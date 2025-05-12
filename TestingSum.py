# ppt_summary_app_with_notes.py
"""
Streamlit PPT Summarizer & Image Assistant using Azure OpenAI with environment variables.
Also extracts and incorporates PowerPoint notes.
"""

from io import BytesIO
from typing import List, Dict, Any
import streamlit as st
from pptx import Presentation
from openai import AzureOpenAI
import requests
import os
from dotenv import load_dotenv

# Load environment variables from .env file if it exists
load_dotenv()

# ──────────────────────────────────────────────────
# 🔧 Azure OpenAI client setup with environment variables
# ──────────────────────────────────────────────────
AZURE_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2023-03-15-preview")
CHAT_MODEL = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-4o-mini")
IMAGE_MODEL = os.getenv("AZURE_OPENAI_IMAGE_DEPLOYMENT_NAME", "dall-e-3")

# Add debug option to show available deployments
DEBUG_MODE = os.getenv("DEBUG_MODE", "False").lower() == "true"

# Smoke test: verify key+endpoint
@st.cache_data(show_spinner=False)
def smoke_test():
    # Remove trailing slash if present
    endpoint = AZURE_ENDPOINT.rstrip('/') if AZURE_ENDPOINT else ""
    url = f"{endpoint}/openai/deployments?api-version={API_VERSION}"
    
    if not AZURE_API_KEY or not AZURE_ENDPOINT:
        st.error("🚨 Missing API key or endpoint. Please set environment variables.")
        st.stop()
        
    try:
        resp = requests.get(url, headers={"api-key": AZURE_API_KEY})
        if resp.status_code != 200:
            st.error(f"🚨 Endpoint test failed [{resp.status_code}]: {resp.text}")
            st.stop()
        return resp.json().get("data", [])
    except Exception as e:
        st.error(f"🚨 Connection test failed: {e}")
        st.stop()

# Initialize AzureOpenAI client
def get_client():
    return AzureOpenAI(
        api_key=AZURE_API_KEY,
        azure_endpoint=AZURE_ENDPOINT,
        api_version=API_VERSION,
    )

# ────────────────────────────────────────────────
# 💡 Helper functions
# ────────────────────────────────────────────────
def extract_slide_data(ppt_io: BytesIO) -> List[Dict[str, Any]]:
    """Extract both slide text and notes from PowerPoint"""
    prs = Presentation(ppt_io)
    slides_data = []
    
    for i, slide in enumerate(prs.slides, 1):
        # Extract text from shapes
        slide_text = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text.append(shape.text)
        
        # Extract notes
        notes_text = ""
        if hasattr(slide, "notes_slide") and slide.notes_slide:
            for shape in slide.notes_slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    notes_text += shape.text + "\n"
        
        slides_data.append({
            "slide_number": i,
            "slide_text": "\n".join(slide_text),
            "notes_text": notes_text.strip(),
            "has_notes": bool(notes_text.strip())
        })
    
    return slides_data


def chat(system: str, user: str, temperature: float = 0.3) -> str:
    client = get_client()
    try:
        resp = client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=temperature,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"Chat API error: {str(e)}")
        
        # Check if it's a deployment not found error
        if "DeploymentNotFound" in str(e):
            deployments = list_deployments()
            if deployments:
                deployment_names = [d['id'] for d in deployments]
                st.warning(f"Available chat deployments: {', '.join(deployment_names)}")
                st.info(f"Update your AZURE_OPENAI_DEPLOYMENT_NAME environment variable to use one of these.")
        
        return f"Error: {str(e)}"


def dalle(prompt: str) -> str:
    client = get_client()
    try:
        out = client.images.generate(
            model=IMAGE_MODEL,
            prompt=prompt,
            n=1,
            size="1024x1024",
        )
        return out.data[0].url
    except Exception as e:
        st.error(f"Image generation failed: {e}")
        
        # Check if it's a deployment not found error
        if "DeploymentNotFound" in str(e) or "Resource not found" in str(e):
            deployments = list_deployments()
            if deployments:
                # Filter for likely image models
                image_deployments = [d['id'] for d in deployments 
                                     if 'dall' in d.get('model', '').lower()]
                
                if image_deployments:
                    st.warning(f"Available image deployments: {', '.join(image_deployments)}")
                    st.info(f"Update your AZURE_OPENAI_IMAGE_DEPLOYMENT_NAME environment variable to use one of these.")
                else:
                    st.warning("No DALL-E deployments found. You may need to create one in Azure OpenAI Studio.")
        
        return None


def list_deployments():
    """List all available deployments"""
    endpoint = AZURE_ENDPOINT.rstrip('/') if AZURE_ENDPOINT else ""
    url = f"{endpoint}/openai/deployments?api-version={API_VERSION}"
    
    try:
        resp = requests.get(url, headers={"api-key": AZURE_API_KEY})
        if resp.status_code == 200:
            return resp.json().get("data", [])
        else:
            st.error(f"Failed to list deployments: {resp.status_code} - {resp.text}")
            return []
    except Exception as e:
        st.error(f"Error listing deployments: {str(e)}")
        return []


@st.cache_data(show_spinner=False)
def summarize_deck(slides_data: List[Dict[str, Any]]):
    """Generate summaries with slide content and notes"""
    # Prepare content for overall deck summary
    deck_content = []
    for data in slides_data:
        if data["has_notes"]:
            deck_content.append(f"Slide {data['slide_number']} content:\n{data['slide_text']}\n\nSlide {data['slide_number']} notes:\n{data['notes_text']}")
        else:
            deck_content.append(f"Slide {data['slide_number']} content:\n{data['slide_text']}")
    
    # Generate overall deck summary
    deck_summary = chat(
        "You are an expert presentation analyst. Summarize the entire deck in under 100 words. "
        "Use insights from the presenter notes where available.",
        "\n\n".join(deck_content),
    )
    
    # Generate per-slide summaries
    slide_summaries = []
    for data in slides_data:
        if data["has_notes"]:
            content = f"Slide content: {data['slide_text']}\n\nPresenter notes: {data['notes_text']}"
            system_prompt = "Summarize this slide in ≤40 words. Incorporate insights from the presenter notes."
        else:
            content = data['slide_text']
            system_prompt = "Summarize this slide in ≤40 words."
        
        summary = chat(system_prompt, content, temperature=0.2)
        slide_summaries.append(summary)
    
    return deck_summary, slide_summaries

# ────────────────────────────────────────────────
# 🎛️ Streamlit UI
# ────────────────────────────────────────────────
def main():
    st.set_page_config("PPT Summarizer & Image Assistant", layout="wide")
    st.title("📊 PPT Summarizer & Image Assistant")

    # Show configuration status
    with st.sidebar:
        st.subheader("📌 Azure OpenAI Configuration")
        st.write("API Endpoint:", AZURE_ENDPOINT)
        st.write("API Version:", API_VERSION)
        st.write("Chat Model:", CHAT_MODEL)
        st.write("Image Model:", IMAGE_MODEL)
        
        # Check if environment variables are set
        if not AZURE_API_KEY:
            st.error("❌ AZURE_OPENAI_API_KEY is not set")
        else:
            st.success("✅ API Key is set")
            
        if not AZURE_ENDPOINT:
            st.error("❌ AZURE_OPENAI_ENDPOINT is not set")
        else:
            st.success("✅ Endpoint is set")
            
        # Debug button to show available deployments
        if st.button("🔍 Show Available Deployments"):
            deployments = list_deployments()
            if deployments:
                st.success(f"Found {len(deployments)} deployments:")
                for dep in deployments:
                    st.write(f"- {dep['id']} (Model: {dep.get('model', 'unknown')})")
                
                # Check for DALL-E deployments
                image_deployments = [d['id'] for d in deployments 
                                    if 'dall' in d.get('model', '').lower()]
                if image_deployments:
                    st.info(f"Image deployments: {', '.join(image_deployments)}")
                    if IMAGE_MODEL not in image_deployments:
                        st.warning(f"Your configured image model '{IMAGE_MODEL}' is not in the list!")
                else:
                    st.warning("No DALL-E deployments found!")
            else:
                st.error("No deployments found or couldn't retrieve deployments")

    # Verify key+endpoint if both are available
    if AZURE_API_KEY and AZURE_ENDPOINT:
        try:
            deployments = smoke_test()
            st.success("✅ Connection to Azure OpenAI successful")
            
            # Debug mode: automatically check deployments
            if DEBUG_MODE:
                # Show all deployments
                st.subheader("Available Deployments")
                for dep in deployments:
                    st.write(f"- {dep['id']} ({dep.get('model', 'unknown')})")
                
                # Check if configured models exist
                deployment_ids = [d['id'] for d in deployments]
                if CHAT_MODEL not in deployment_ids:
                    st.warning(f"⚠️ Chat model '{CHAT_MODEL}' not found in deployments!")
                if IMAGE_MODEL not in deployment_ids:
                    st.warning(f"⚠️ Image model '{IMAGE_MODEL}' not found in deployments!")
        except Exception as e:
            st.error(f"🚨 API Connection Error: {e}")
            st.info("Please check your Azure OpenAI credentials and setup.")
            st.stop()
    else:
        st.error("❌ Missing API key or endpoint. Please set environment variables.")
        st.stop()

    # Upload PPT
    file = st.file_uploader("Upload your .pptx file", type=["pptx"])
    if not file:
        st.info("⬆️ Upload a PowerPoint to get started.")
        return

    # Extract slide data (both text and notes)
    slides_data = extract_slide_data(file)
    
    # Show notes statistics
    slides_with_notes = sum(1 for slide in slides_data if slide["has_notes"])
    st.info(f"Found {len(slides_data)} slides, {slides_with_notes} with presenter notes.")

    # Generate summaries
    if st.button("🚀 Generate Summaries") or "summaries" in st.session_state:
        if "summaries" not in st.session_state:
            with st.spinner("Calling Azure OpenAI…"):
                st.session_state.summaries = summarize_deck(slides_data)

        deck, per_slide = st.session_state.summaries
        st.subheader("Overall Deck Summary")
        st.write(deck)
        st.divider()
        st.subheader("Per-Slide Summaries & Illustrations")

        for idx, (data, summ) in enumerate(zip(slides_data, per_slide)):
            with st.expander(f"Slide {data['slide_number']}"):
                if data["has_notes"]:
                    st.markdown("💬 **This slide has presenter notes**")
                
                st.markdown(f"**Summary:** {summ}")
                
                # Show original content in tabs
                tab1, tab2 = st.tabs(["Slide Content", "Presenter Notes"])
                with tab1:
                    st.text(data["slide_text"])
                with tab2:
                    if data["has_notes"]:
                        st.text(data["notes_text"])
                    else:
                        st.info("No presenter notes for this slide.")
                
                # Image generation
                c1, c2 = st.columns([1, 3])
                with c1:
                    if st.button("🎨 Generate Image", key=f"btn{idx}"):
                        prompt = chat(
                            "Craft a vivid, photo-realistic DALL·E-3 prompt (16:9, minimal text). Return only the prompt.",
                            f"Slide summary: {summ}",
                            temperature=0.7,
                        )
                        with st.spinner("Generating image…"):
                            st.session_state[f"img{idx}"] = dalle(prompt)
                            st.session_state[f"prompt{idx}"] = prompt
                with c2:
                    if url := st.session_state.get(f"img{idx}"):
                        st.image(url, use_column_width=True)
                    elif f"prompt{idx}" in st.session_state:
                        st.warning("⚠️ Image generation failed.")
                        st.info("The prompt that would have been used:")
                        st.code(st.session_state[f"prompt{idx}"])

if __name__ == "__main__":
    main()