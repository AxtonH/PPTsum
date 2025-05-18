# ppt_summary_app_with_chatbot.py
"""
Streamlit PPT Summarizer, Design Assistant & Conversational Chatbot
───────────────────────────────────────────────────────────────────
• Upload a PowerPoint → get deck summary, context, per‑slide insights
• Suggests layouts, charts, visuals, colour schemes
• Built‑in chatbot that answers questions about the deck
"""

# ─── IMPORTS ──────────────────────────────────────────────────────────────────
from io import BytesIO
from typing import List, Dict, Any, Tuple
import os, json, time, requests, re
import streamlit as st
from pptx import Presentation
from dotenv import load_dotenv
from openai import AzureOpenAI

# Handle hydralit_components import gracefully
try:
    import hydralit_components as hc
    HYDRALIT_AVAILABLE = True
except ImportError:
    HYDRALIT_AVAILABLE = False
    st.warning("Advanced navigation features disabled (hydralit_components not available)")

# ─── CONFIGURATION ───────────────────────────────────────────────────────────
load_dotenv()
AZURE_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
CHAT_MODEL = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-4")
DEBUG_MODE = os.getenv("DEBUG_MODE", "False").lower() == "true"

# ─── AZURE HELPERS ───────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_client():
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        st.error("🚨 Missing Azure credentials. Set env vars.")
        st.stop()
    return AzureOpenAI(
        api_key=AZURE_API_KEY,
        azure_endpoint=AZURE_ENDPOINT,
        api_version=API_VERSION,
    )

@st.cache_data(show_spinner=False)
def smoke_test():
    url = f"{AZURE_ENDPOINT.rstrip('/')}/openai/deployments?api-version={API_VERSION}"
    resp = requests.get(url, headers={"api-key": AZURE_API_KEY})
    if resp.status_code != 200:
        st.error(f"🚨 Endpoint check failed [{resp.status_code}]: {resp.text}")
        st.stop()
    return resp.json().get("data", [])

def chat(system: str, user: str, temperature: float = 0.3, max_attempts: int = 2) -> str:
    client = get_client()
    for attempt in range(max_attempts):
        try:
            resp = client.chat.completions.create(
                model=CHAT_MODEL,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                temperature=temperature,
            )
            return resp.choices[0].message.content.strip()
        except Exception:
            if attempt + 1 == max_attempts:
                raise
            time.sleep(1.5)

# ─── SLIDE PROCESSING ────────────────────────────────────────────────────────
def extract_slide_data(ppt_io: BytesIO) -> List[Dict[str, Any]]:
    prs = Presentation(ppt_io)
    data = []
    for i, slide in enumerate(prs.slides, 1):
        text = [sh.text for sh in slide.shapes if hasattr(sh, "text") and sh.text.strip()]
        notes = ""
        if getattr(slide, "notes_slide", None):
            for sh in slide.notes_slide.shapes:
                if hasattr(sh, "text") and sh.text.strip():
                    notes += sh.text + "\n"
        data.append({
            "slide_number": i,
            "slide_text": "\n".join(text),
            "notes_text": notes.strip(),
            "has_notes": bool(notes.strip()),
        })
    return data

# ─── CONTEXT EXTRACTION ──────────────────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=3600)
def identify_context(slides: List[Dict[str, Any]]) -> Dict[str, str]:
    # Extract first few slides plus title/agenda slides
    context_slides = []
    for d in slides[:min(7, len(slides))]:
        if d['slide_text'].strip():
            context_slides.append(d)
    
    # Build context blocks with prioritized content
    context_blocks = []
    for d in context_slides:
        is_title_slide = (d['slide_number'] == 1) or ("title" in d['slide_text'].lower())
        block = f"Slide {d['slide_number']}:"
        if is_title_slide:
            block = "🔹 TITLE SLIDE:\n" + d['slide_text']
        elif "agenda" in d['slide_text'].lower():
            block = "🔸 AGENDA SLIDE:\n" + d['slide_text']
        else:
            block += "\n" + d['slide_text']
        if d["has_notes"]:
            block += f"\nNOTES: {d['notes_text']}"
        context_blocks.append(block)
    
    # Enhanced system prompt
    system_prompt = """Extract presentation metadata as JSON with:
    - topic: Specific subject (e.g., "2024 Marketing Strategy")
    - region: Geographic focus if mentioned
    - purpose: Only from: inform|persuade|propose|educate|report|pitch|review
    
    Rules:
    1. Be SPECIFIC - avoid generic terms
    2. For purpose, ONLY use allowed values
    3. Return ONLY valid JSON like:
    {"topic": "...", "region": "...", "purpose": "..."}"""
    
    # Get raw response with retries
    for attempt in range(3):
        try:
            raw = chat(
                system_prompt,
                "SLIDE CONTENT:\n\n" + "\n".join(context_blocks),
                temperature=0.2
            )
            context = json.loads(raw)
            
            # Validate and clean
            context = {
                "topic": context.get("topic", "Unknown").strip() or "Unknown",
                "region": context.get("region", "Unknown").strip() or "Unknown",
                "purpose": context.get("purpose", "Unknown").strip().lower()
            }
            
            # Validate purpose
            allowed_purposes = {"inform", "persuade", "propose", "educate", "report", "pitch", "review"}
            if context["purpose"] not in allowed_purposes:
                context["purpose"] = "Unknown"
                
            return context
            
        except (json.JSONDecodeError, ValueError):
            if attempt == 2:  # Final attempt
                return {"topic": "Unknown", "region": "Unknown", "purpose": "Unknown"}
            time.sleep(1)

# ─── SUMMARY GENERATION ──────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def summarize_deck(slides: List[Dict[str, Any]]) -> Tuple[str, Dict[str, str], List[str], List[str], List[str], List[str]]:
    ctx = identify_context(slides)
    full = []
    for d in slides:
        segment = f"Slide {d['slide_number']} content:\n{d['slide_text']}"
        if d["has_notes"]:
            segment += f"\n\nNotes:\n{d['notes_text']}"
        full.append(segment)
    
    deck_summary = chat(
        "Summarize this presentation in <100 words for a busy executive.",
        "\n\n".join(full)
    )
    
    slide_sums, design_notes, layouts, charts = [], [], [], []
    for d in slides:
        prompt = f"Slide content: {d['slide_text']}" + (f"\nNotes: {d['notes_text']}" if d["has_notes"] else "")
        resp = chat(
            "Provide:\n1. Key point (≤30 words)\n2. Design tip (≤20 words)",
            prompt
        )
        parts = resp.split("2.")
        slide_sums.append(parts[0].replace("1.", "").strip() if parts else resp)
        design_notes.append(parts[1].strip() if len(parts) > 1 else "—")
        layouts.append(suggest_layout(prompt))
        charts.append(suggest_chart_type(prompt))
    
    return deck_summary, ctx, slide_sums, design_notes, layouts, charts

# ─── DESIGN HELPERS ──────────────────────────────────────────────────────────
def suggest_layout(slide_content: str) -> str:
    return chat(
        "Recommend a slide layout for consulting-style presentations. Be specific about element placement.",
        slide_content,
        temperature=0.4
    )

def suggest_chart_type(slide_content: str) -> str:
    return chat(
        "Suggest the best chart type for this content. If no chart fits, recommend an alternative visual.",
        slide_content,
        temperature=0.4
    )

def generate_image_prompt(slide_summary: str, context: Dict[str, str]) -> str:
    return chat(
        "Create a detailed image prompt for a professional presentation slide.",
        f"Content: {slide_summary}\nContext: {json.dumps(context)}",
        temperature=0.7
    )

# ─── UI COMPONENTS ───────────────────────────────────────────────────────────
def display_context(ctx: Dict[str, str]):
    with st.expander("🔍 Presentation Context", expanded=True):
        cols = st.columns(3)
        
        # Topic
        with cols[0]:
            st.markdown("### 📌 Topic")
            if ctx['topic'] == 'Unknown':
                new_topic = st.text_input("Enter topic", key="topic_input")
                if new_topic:
                    ctx['topic'] = new_topic
            else:
                st.success(ctx['topic'])
        
        # Region
        with cols[1]:
            st.markdown("### 🌍 Region")
            if ctx['region'] == 'Unknown':
                st.info("Not specified")
            else:
                st.success(ctx['region'])
        
        # Purpose
        with cols[2]:
            st.markdown("### 🎯 Purpose")
            purpose_icons = {
                "inform": "📢 Inform",
                "persuade": "🫵 Persuade",
                "propose": "📑 Propose",
                "educate": "🎓 Educate",
                "report": "📈 Report",
                "pitch": "🤝 Pitch",
                "review": "🔍 Review",
                "Unknown": "❓ Unknown"
            }
            st.markdown(purpose_icons.get(ctx['purpose'], "❓") + " " + ctx['purpose'].capitalize())

# ─── MAIN APP ────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(
        "PPT Assistant", 
        layout="wide",
        page_icon="📊",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS
    st.markdown("""
    <style>
        .stApp { background-color: #f8f9fa; }
        .sidebar .sidebar-content { background-color: #ffffff; }
        .stButton>button { width: 100%; }
        .stExpander { background-color: white; border-radius: 8px; }
        .stChatMessage { border-radius: 12px; padding: 1rem; }
    </style>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.title("⚙️ Settings")
        
        # Azure Status
        with st.container(border=True):
            st.subheader("Azure Status")
            if AZURE_API_KEY and AZURE_ENDPOINT:
                try:
                    deployments = smoke_test()
                    st.success("✅ Connected")
                    if DEBUG_MODE:
                        st.markdown(f"**Model:** {CHAT_MODEL}")
                        st.markdown(f"**API Version:** {API_VERSION}")
                except Exception as e:
                    st.error(f"❌ Connection failed: {str(e)}")
            else:
                st.error("❌ Missing credentials")
        
        # File Upload
        st.divider()
        ppt = st.file_uploader(
            "Upload PowerPoint", 
            type=["pptx"],
            help="Upload a .pptx file to analyze"
        )
    
    # Main Content
    st.title("📊 PowerPoint Analysis Suite")
    
    if not ppt:
        st.info("👈 Upload a PowerPoint file to begin")
        return
    
    # Clear session if new file
    if st.session_state.get("_current_file") != ppt.name:
        st.session_state.clear()
        st.session_state._current_file = ppt.name
    
    # Process slides
    slides = extract_slide_data(ppt)
    st.success(f"✔️ Loaded {len(slides)} slides")
    
    # Analysis
    if st.button("🚀 Analyze Presentation", type="primary") or "summaries" in st.session_state:
        if "summaries" not in st.session_state:
            with st.spinner("Analyzing presentation content..."):
                st.session_state.summaries = summarize_deck(slides)
        
        deck_sum, ctx, slide_sums, design_notes, layouts, charts = st.session_state.summaries
        
        # Display Context
        display_context(ctx)
        
        # Deck Summary
        with st.expander("📝 Executive Summary", expanded=True):
            st.write(deck_sum)
        
        # Slide Navigation
        st.divider()
        st.subheader("📑 Slide Analysis")
        
        if HYDRALIT_AVAILABLE:
            menu = [{"id": i, "label": f"Slide {d['slide_number']}", "icon": "bi-easel"} 
                   for i, d in enumerate(slides)]
            sel = hc.nav_bar(menu, sticky_mode="pinned", key="nav")
            idx = sel if isinstance(sel, int) else 0
        else:
            # Fallback to selectbox if hydralit not available
            idx = st.selectbox(
                "Select slide to view:",
                options=range(len(slides)),
                format_func=lambda x: f"Slide {slides[x]['slide_number']}"
            )
        
        current_slide = slides[idx]
        
        # Slide Content
        col1, col2 = st.columns([3, 2])
        
        with col1:
            with st.container(border=True):
                st.markdown("#### 📝 Content Summary")
                st.write(slide_sums[idx])
            
            with st.container(border=True):
                st.markdown("#### 🎨 Design Suggestions")
                st.write(design_notes[idx])
            
            if current_slide["has_notes"]:
                with st.expander("🗒️ Presenter Notes"):
                    st.info(current_slide["notes_text"])
        
        with col2:
            with st.container(border=True):
                st.markdown("#### 🖼️ Recommended Layout")
                st.write(layouts[idx])
            
            with st.container(border=True):
                st.markdown("#### 📊 Data Visualization")
                st.write(charts[idx])
            
            with st.expander("📜 Raw Content"):
                st.text(current_slide["slide_text"])
        
        # Visual Suggestions
        st.divider()
        st.subheader("🖌️ Visual Enhancements")
        
        if st.button("💡 Generate Visual Ideas", key=f"vis_{idx}"):
            with st.spinner("Creating visual concepts..."):
                visual_prompt = generate_image_prompt(
                    f"{slide_sums[idx]}\nDesign tips: {design_notes[idx]}",
                    ctx
                )
                st.session_state[f"visual_{idx}"] = visual_prompt
        
        if f"visual_{idx}" in st.session_state:
            with st.expander("🖼️ Visual Concept"):
                st.markdown("**Image Prompt:**")
                st.code(st.session_state[f"visual_{idx}"], language="text")
        
        # Chatbot
        st.divider()
        st.subheader("💬 Presentation Chatbot")
        
        if "messages" not in st.session_state:
            st.session_state.messages = []
        
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
        
        if prompt := st.chat_input("Ask about the presentation..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    response = chat(
                        "Answer questions about this presentation:",
                        f"Question: {prompt}\n\nPresentation Content:\n{deck_sum}"
                    )
                    st.write(response)
            st.session_state.messages.append({"role": "assistant", "content": response})

if __name__ == "__main__":
    main()
