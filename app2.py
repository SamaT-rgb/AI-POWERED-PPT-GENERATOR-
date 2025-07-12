import os
import time
import re
import requests
from io import BytesIO
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from dotenv import load_dotenv
import google.generativeai as genai
from pexels_api import API as PexelsAPI
from PIL import Image, ImageDraw, ImageFont

# --- Configuration and Setup ---

def configure_app():
    """Set up page configuration and load API keys."""
    st.set_page_config(
        page_title="AI PowerPoint Pro",
        page_icon="✨",
        layout="wide"
    )
    st.title("✨ AI PowerPoint Pro")
    
    load_dotenv()
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    pexels_api_key = os.getenv("PEXELS_API_KEY")
    
    if not gemini_api_key:
        st.error("GEMINI_API_KEY not found. Please set it in your .env file.")
        st.stop()
    if not pexels_api_key:
        st.warning("PEXELS_API_KEY not found. Automatic image search will be disabled.")
    
    genai.configure(api_key=gemini_api_key)
    return pexels_api_key

# --- Image Generation and Fetching with Fallback ---

def create_placeholder_image(text):
    """Creates a grey placeholder image with centered text."""
    width, height = 800, 500
    img = Image.new('RGB', (width, height), color='#DDDDDD')
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 40)
    except IOError:
        font = ImageFont.load_default(size=40)
    
    text_bbox = draw.textbbox((0, 0), text, font=font, align="center")
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    position = ((width - text_width) / 2, (height - text_height) / 2)
    
    draw.text(position, text, fill='#555555', font=font, align="center")
    
    img_buffer = BytesIO()
    img.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    return img_buffer

def fetch_image(api_key, query):
    """Fetches an image from Pexels or returns a placeholder if it fails."""
    if not api_key or not query:
        return create_placeholder_image(f"Image for:\n{query}")

    try:
        api = PexelsAPI(api_key)
        api.search(query, page=1, results_per_page=1)
        if not api.photos:
            st.warning(f"No image found for query: '{query}'. Using placeholder.")
            return create_placeholder_image(f"No image found for:\n'{query}'")
        
        photo_url = api.photos[0].src['large']
        response = requests.get(photo_url)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"Error fetching image for '{query}': {e}. Using placeholder.")
        return create_placeholder_image(f"API Error for:\n'{query}'")

# --- Core AI and PowerPoint Generation Logic ---

def generate_content_from_ai(prompt, max_retries=3, retry_delay=60):
    """Generates content using the Gemini AI model with robust retry logic."""
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        for attempt in range(max_retries):
            try:
                response = model.generate_content(prompt)
                return response.text
            except Exception as e:
                if "429" in str(e):
                    st.warning(f"Rate limit reached. Waiting {retry_delay}s... (Attempt {attempt + 1})")
                    time.sleep(retry_delay)
                else:
                    st.error(f"An unexpected error occurred with the AI model: {e}")
                    return None
        st.error("Failed to generate content after multiple retries.")
        return None
    except Exception as e:
        st.error(f"Failed to initialize the generative model: {e}")
        return None

def parse_ai_response(content):
    """Parses the structured AI response into a list of slide dictionaries."""
    slides_data = []
    slide_blocks = re.findall(
        r'Slide Type:\s*(.*?)\nTitle:\s*(.*?)\nContent:\s*(.*?)\nImage Search Query:\s*(.*?)(?=\n\n---|\Z)',
        content,
        re.DOTALL
    )
    
    for block in slide_blocks:
        slides_data.append({
            "type": block[0].strip(),
            "title": block[1].strip(),
            "content": block[2].strip().replace('- ', '\n- ').lstrip('\n'),
            "query": block[3].strip()
        })
    return slides_data

def add_slide_to_presentation(prs, slide_data, image_stream):
    """Intelligently adds a slide, adapting to the provided template's layouts."""
    title_text = slide_data['title']
    content_text = slide_data['content']

    layout_name_map = {
        "Title Slide": "Title Slide",
        "Content Slide": "Title and Content",
        "Section Header": "Section Header"
    }

    chosen_layout_index = -1
    for i, layout in enumerate(prs.slide_layouts):
        if layout.name == layout_name_map.get(slide_data['type']):
            chosen_layout_index = i
            break
    
    if chosen_layout_index == -1:
        chosen_layout_index = 1 if slide_data['type'] != "Title Slide" else 0

    slide_layout = prs.slide_layouts[chosen_layout_index]
    slide = prs.slides.add_slide(slide_layout)

    if slide.shapes.title:
        slide.shapes.title.text = title_text
    
    content_placeholder = None
    for shape in slide.placeholders:
        if "Content" in shape.name or shape.placeholder_format.idx == 1:
            content_placeholder = shape
            break
    if content_placeholder:
        content_placeholder.text = content_text

    if image_stream:
        pic_placeholder = None
        for shape in slide.placeholders:
            if "Picture" in shape.name or 'PIC' in shape.placeholder_format.type:
                pic_placeholder = shape
                break
        
        if pic_placeholder:
            try:
                pic_placeholder.insert_picture(image_stream)
            except Exception:
                slide.shapes.add_picture(image_stream, Inches(5.5), Inches(1.8), height=Inches(4.5))
        else:
            slide.shapes.add_picture(image_stream, Inches(5.5), Inches(1.8), height=Inches(4.5))


def generate_ppt_from_text(text_input, user_images, auto_image, pexels_key, template_file):
    """Generates a PowerPoint presentation using a template and advanced image handling."""
    prompt = f"""
    Create a professional presentation outline from this text: "{text_input}".
    Generate 4 slides: one "Title Slide" and three "Content Slides".
    For each slide, specify its "Slide Type", a "Title", "Content" (as bullet points), and an "Image Search Query".
    Format the output strictly as follows, using '---' as a separator:

    Slide Type: Title Slide
    Title: [Catchy Title Here]
    Content: [A brief, engaging subtitle or key takeaway]
    Image Search Query: [A query for a powerful opening image, e.g., 'business strategy']

    ---

    Slide Type: Content Slide
    Title: [Informative Title Here]
    Content:
    - First key point
    - Second key point
    - Third key point
    Image Search Query: [A specific query for an image illustrating the content]
    """
    
    with st.spinner("Step 1/3: Generating presentation script with AI..."):
        content = generate_content_from_ai(prompt)
        if not content:
            st.error("Could not generate presentation content.")
            return None
        slides_data = parse_ai_response(content)

    if not slides_data:
        st.error("AI did not return a valid structure. Please check the prompt or try again.")
        return None

    # --- THIS IS THE CORRECTED SECTION ---
    if template_file:
        # Read the uploaded file into a BytesIO buffer to ensure compatibility
        template_buffer = BytesIO(template_file.getvalue())
        prs = Presentation(template_buffer)
    else:
        # If no template, create a blank presentation
        prs = Presentation()
    # --- END OF CORRECTION ---
    
    with st.spinner("Step 2/3: Populating slides and fetching images..."):
        for i, slide_info in enumerate(slides_data):
            image_to_add = None
            if i < len(user_images):
                image_to_add = user_images[i]
            elif auto_image and pexels_key:
                image_to_add = fetch_image(pexels_key, slide_info['query'])
            
            add_slide_to_presentation(prs, slide_info, image_to_add)
            
    with st.spinner("Step 3/3: Finalizing presentation..."):
        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
    
    return ppt_buffer

# --- Streamlit UI ---

def main():
    pexels_key = configure_app()

    st.sidebar.header("🎨 Design Options")
    template_file = st.sidebar.file_uploader(
        "Upload a PowerPoint Template (.potx)",
        type=['potx']
    )
    if template_file:
        st.sidebar.success(f"Using template: {template_file.name}")
    else:
        st.sidebar.info("No template uploaded. Using default blank design.")

    st.sidebar.header("🖼️ Image Options")
    auto_image = st.sidebar.checkbox(
        "Find images automatically", 
        value=True,
        help="Uses Pexels.com. If an image isn't found, a placeholder is created."
    )
    user_images = st.sidebar.file_uploader(
        "Or upload your own images", 
        type=['png', 'jpg', 'jpeg'], 
        accept_multiple_files=True
    )
    
    st.header("✍️ Content Input")
    text_input = st.text_area("Enter the topic or text for your presentation:", height=200, 
        value="The importance of renewable energy. Discuss solar, wind, and hydro power. Cover the environmental benefits, economic advantages, and future challenges of transitioning to a green economy.")
    
    if st.button("🚀 Generate Presentation"):
        if text_input:
            ppt_buffer = generate_ppt_from_text(text_input, user_images, auto_image, pexels_key, template_file)
            if ppt_buffer:
                st.session_state['generated_ppt'] = ppt_buffer
                st.session_state['ppt_name'] = "AI_Generated_Presentation.pptx"
                st.success("Presentation generated successfully!")
        else:
            st.warning("Please enter some text to generate a presentation.")

    if 'generated_ppt' in st.session_state and st.session_state['generated_ppt']:
        st.download_button(
            label="⬇️ Download Presentation",
            data=st.session_state['generated_ppt'],
            file_name=st.session_state['ppt_name'],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            on_click=lambda: st.session_state.update({'generated_ppt': None})
        )

if __name__ == "__main__":
    main()
