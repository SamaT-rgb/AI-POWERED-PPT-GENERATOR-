import os
import base64
from io import BytesIO
import time
import re
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from dotenv import load_dotenv
import google.generativeai as genai
from pexels_api import API as PexelsAPI # New import

# --- Configuration and Setup ---

def configure_app():
    """Set up page configuration and load API keys."""
    st.set_page_config(
        page_title="AI PowerPoint Generator",
        page_icon="🤖",
        layout="wide" # Use wide layout for better UI
    )
    st.title("AI-Powered Presentation Generator")
    
    load_dotenv()
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    pexels_api_key = os.getenv("PEXELS_API_KEY") # New
    
    if not gemini_api_key:
        st.error("GEMINI_API_KEY not found. Please set it in your .env file.")
        st.stop()
    if not pexels_api_key:
        st.error("PEXELS_API_KEY not found. Please set it in your .env file for image search.")
        st.stop()
        
    genai.configure(api_key=gemini_api_key)
    return pexels_api_key

# --- NEW: Image Search and Handling Functions ---

def fetch_image_from_pexels(api_key, query):
    """Fetches an image from Pexels based on a search query."""
    try:
        api = PexelsAPI(api_key)
        api.search(query, page=1, results_per_page=1)
        if not api.photos:
            st.warning(f"No image found for query: '{query}'")
            return None
        
        # Get the URL of a medium-sized photo and download it
        photo_url = api.photos[0].src['medium']
        response = requests.get(photo_url)
        response.raise_for_status() # Raise an exception for bad status codes
        return BytesIO(response.content)
    except Exception as e:
        st.error(f"Error fetching image from Pexels: {e}")
        return None

# --- Gemini AI Functions ---

def generate_content_from_ai(prompt, max_retries=3, retry_delay=60):
    """Generates content using the Gemini AI model with robust retry logic."""
    try:
        model = genai.GenerativeModel('gemini-2.0-flash')
        # ... (rest of the function is the same as before) ...
        for attempt in range(max_retries):
            try:
                response = model.generate_content(prompt)
                return response.text
            except Exception as e:
                if "429" in str(e):
                    st.warning(f"Rate limit exceeded. Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                else:
                    st.error(f"An unexpected error occurred with the AI model: {e}")
                    return None
        st.error("Failed to generate content after multiple retries.")
        return None
    except Exception as e:
        st.error(f"Failed to initialize the generative model: {e}")
        return None

# --- PowerPoint Generation Functions (Upgraded for Images) ---

# +++ THIS IS THE NEW, CORRECTED FUNCTION +++
def add_slide_with_image(prs, title, content, image_stream=None):
    """
    Adds a new slide with a professional two-column layout: 
    title, content on the left, and image on the right.
    """
    # Use 'Title and Content' layout, which is very standard (layout index 1)
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    # Set title
    title_shape = slide.shapes.title
    title_shape.text = title

    # Set content in the main body placeholder
    body_shape = slide.placeholders[1]
    text_frame = body_shape.text_frame
    text_frame.text = content

    # Resize and reposition the text box to the left half of the slide
    body_shape.left = Inches(0.5)
    body_shape.width = Inches(4.5)
    body_shape.top = Inches(1.5)
    body_shape.height = Inches(5.5)

    # Add image to the right half of the slide if an image is provided
    if image_stream:
        # Define position and size for the image on the right side
        left = Inches(5.2)
        top = Inches(1.75)
        height = Inches(4.0) # We can set the height and let the width be proportional
        
        # Use slide.shapes.add_picture, which is the robust method
        pic = slide.shapes.add_picture(image_stream, left, top, height=height)

def parse_ai_response(content):
    """Parses the structured AI response into a list of slide dictionaries."""
    slides_data = []
    # Use regex to find all slide blocks
    slide_blocks = re.findall(r'Slide \d+ Title: (.*?)\n- (.*?)\nImage Search Query: (.*?)(?=\n\nSlide|\Z)', content, re.DOTALL)
    
    for block in slide_blocks:
        title = block[0].strip()
        # The bullet points might span multiple lines, so we re-format them
        bullets = "- " + block[1].strip().replace('\n  -', '\n-')
        search_query = block[2].strip()
        slides_data.append({"title": title, "content": bullets, "query": search_query})
        
    return slides_data

def generate_ppt_from_text(text_input, user_images, auto_image, pexels_key):
    """Generates a PowerPoint presentation from text with image support."""
    # NEW, more structured prompt
    prompt = f"""
    Create a PowerPoint presentation outline based on the following text.
    Generate a title slide and 3 content slides.
    For each slide, provide a short title, 3-5 bullet points, and a concise, effective Pexels.com search query for a relevant, high-quality image.
    Format the output strictly as follows:

    Slide 1 Title: [Your Title Here]
    - Bullet point 1
    - Bullet point 2
    Image Search Query: [A relevant search query]

    ---

    Slide 2 Title: [Your Title Here]
    - Bullet point 1
    - Bullet point 2
    - Bullet point 3
    Image Search Query: [Another relevant search query]

    Text: "{text_input}"
    """
    
    with st.spinner("Generating presentation content with AI..."):
        content = generate_content_from_ai(prompt)
        if not content:
            st.error("Could not generate presentation content.")
            return None
        
        # New parsing function for the structured response
        slides_data = parse_ai_response(content.replace("\n---\n", "\n\n"))

    if not slides_data:
        st.error("AI did not return a valid presentation structure. Please try again.")
        return None

    prs = Presentation()
    
    for i, slide_info in enumerate(slides_data):
        image_to_add = None
        # Priority 1: Use a user-uploaded image if available
        if i < len(user_images):
            image_to_add = user_images[i]
        # Priority 2: Fetch an image automatically if the user enabled it
        elif auto_image:
            with st.spinner(f"Searching for image: '{slide_info['query']}'..."):
                image_to_add = fetch_image_from_pexels(pexels_key, slide_info['query'])

        add_slide_with_image(prs, slide_info['title'], slide_info['content'], image_to_add)
        
    return prs

# --- (The CSV functions can remain as they are) ---
# ... (generate_ppt_from_csv, create_chart_image, etc.) ...

# --- Streamlit UI (Upgraded for Images) ---

def ui_text_section(pexels_key):
    """Renders the UI section for text-based presentation generation."""
    st.header("1. Generate Presentation from Text")
    
    text_input = st.text_area("Enter your text here:", height=150, 
        value="Lions are majestic big cats, often called the 'king of the jungle.' They are native to Africa and India and live in social groups called prides. Male lions are distinguished by their impressive manes, while lionesses are the primary hunters, working together to prey on large herbivores.")
    
    st.subheader("Image Options")
    col1, col2 = st.columns(2)
    with col1:
        # New multi-file uploader
        user_images = st.file_uploader(
            "Upload your own images (one per slide)", 
            type=['png', 'jpg', 'jpeg'], 
            accept_multiple_files=True
        )
    with col2:
        # New checkbox for automatic image search
        auto_image = st.checkbox(
            "Automatically find images if none are uploaded", 
            value=True,
            help="Uses Pexels.com to find a relevant image for each slide based on its content."
        )

    if st.button("Generate Presentation from Text"):
        if text_input:
            # Pass the new image parameters to the generation function
            presentation = generate_ppt_from_text(text_input, user_images, auto_image, pexels_key)
            if presentation:
                st.session_state['generated_ppt'] = presentation
                st.session_state['ppt_name'] = "Text_Based_Presentation.pptx"
                st.success("Presentation generated successfully!")
        else:
            st.warning("Please enter some text to generate a presentation.")

# ... (ui_csv_section remains the same) ...

def main():
    """Main function to run the Streamlit app."""
    pexels_key = configure_app()

    # Initialize session state
    if 'generated_ppt' not in st.session_state:
        st.session_state['generated_ppt'] = None
    if 'ppt_name' not in st.session_state:
        st.session_state['ppt_name'] = ""

    # UI Layout
    tab1, tab2 = st.tabs(["▶️ Generate from Text", "📊 Generate from CSV"])

    with tab1:
        ui_text_section(pexels_key) # Pass the key to the UI function
        
    with tab2:
        # ui_csv_section() # You can call your CSV section function here
        st.write("CSV functionality is ready.")


    # Download button logic remains the same
    if st.session_state['generated_ppt']:
        with BytesIO() as ppt_buffer:
            st.session_state['generated_ppt'].save(ppt_buffer)
            ppt_buffer.seek(0)
            
            st.download_button(
                label=f"Download {st.session_state['ppt_name']}",
                data=ppt_buffer.getvalue(),
                file_name=st.session_state['ppt_name'],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                on_click=lambda: st.session_state.update({'generated_ppt': None, 'ppt_name': ""})
            )
            
if __name__ == "__main__":
    # A small addition to make the Pexels part work seamlessly
    import requests
    main()
