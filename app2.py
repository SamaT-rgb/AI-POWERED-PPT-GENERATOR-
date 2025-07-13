import os
import time
import re
import requests
from io import BytesIO
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt # <--- THIS LINE IS NOW CORRECTED
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from dotenv import load_dotenv
import google.generativeai as genai
from PIL import Image, ImageDraw, ImageFont

# --- Configuration and Setup ---

def configure_app():
    st.set_page_config(page_title="AI PowerPoint Pro", page_icon="✨", layout="wide")
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

# --- Image Generation and Fetching (REBUILT) ---

def create_placeholder_image(text):
    width, height = 1200, 800
    img = Image.new('RGB', (width, height), color='#E0E0E0')
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 60)
    except IOError:
        font = ImageFont.load_default(size=60)
    text_bbox = draw.textbbox((0, 0), text, font=font, align="center")
    position = ((width - (text_bbox[2] - text_bbox[0])) / 2, (height - (text_bbox[3] - text_bbox[1])) / 2)
    draw.text(position, text, fill='#6c757d', font=font, align="center")
    img_buffer = BytesIO()
    img.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    return img_buffer

def fetch_image_from_pexels(api_key, query):
    """Fetches an image directly from the Pexels API using requests."""
    if not api_key or not query:
        return create_placeholder_image(f"Image for:\n{query}")
    headers = {"Authorization": api_key}
    params = {"query": query, "per_page": 1, "page": 1}
    try:
        response = requests.get("https://api.pexels.com/v1/search", headers=headers, params=params)
        response.raise_for_status()
        data = response.json()
        if not data.get("photos"):
            st.warning(f"No image found for query: '{query}'. Using placeholder.")
            return create_placeholder_image(f"No image for:\n'{query}'")
        
        photo_url = data["photos"][0]["src"]["large"]
        image_response = requests.get(photo_url)
        image_response.raise_for_status()
        return BytesIO(image_response.content)
    except requests.exceptions.RequestException as e:
        st.error(f"Pexels API Error for '{query}': {e}. Using placeholder.")
        return create_placeholder_image(f"API Error for:\n'{query}'")

# --- Core AI and PowerPoint Generation (REBUILT) ---

def generate_content_from_ai(prompt):
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"An error occurred with the AI model: {e}")
        return None

def parse_ai_response(content):
    slides_data = []
    slide_blocks = re.findall(
        r'Slide Type:\s*(.*?)\nTitle:\s*(.*?)\nContent:\s*(.*?)\nImage Search Query:\s*(.*?)(?=\n\n---|\Z)',
        content, re.DOTALL
    )
    for block in slide_blocks:
        slides_data.append({
            "type": block[0].strip(),
            "title": block[1].strip(),
            "content": block[2].strip().replace('- ', '\n• ').lstrip('\n'),
            "query": block[3].strip()
        })
    return slides_data

def split_text_for_slides(text, max_chars=600):
    """Splits long text into chunks that will fit on individual slides."""
    if len(text) < max_chars:
        return [text]
    
    paragraphs = text.split('\n')
    chunks = []
    current_chunk = ""
    for p in paragraphs:
        if len(current_chunk) + len(p) < max_chars:
            current_chunk += p + "\n"
        else:
            chunks.append(current_chunk)
            current_chunk = p + "\n"
    chunks.append(current_chunk)
    return chunks

def add_slide_with_layout(prs, title, content_chunk, image_stream=None):
    """A robust function to add a slide with a professional, programmatic layout."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1.0))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(36)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    if image_stream:
        text_left, text_top, text_width, text_height = Inches(0.5), Inches(1.2), Inches(4.5), Inches(5.8)
        img_left, img_top, img_width, img_height = Inches(5.5), Inches(1.5), Inches(4.0), Inches(5.5)

        text_box = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
        text_frame = text_box.text_frame
        text_frame.text = content_chunk
        text_frame.paragraphs[0].font.size = Pt(16)
        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        try:
            image_stream.seek(0) # Ensure buffer is at the beginning
            img = Image.open(image_stream)
            img_w, img_h = img.size
            
            f_w = img_width / Inches(img_w / 914400) # Convert EMU to Inches for calc
            f_h = img_height / Inches(img_h / 914400)
            f = min(f_w, f_h)
            new_w, new_h = Inches(img_w * f / 914400), Inches(img_h * f / 914400)

            final_img_left = img_left + (img_width - new_w) / 2
            final_img_top = img_top + (img_height - new_h) / 2
            
            image_stream.seek(0) # Rewind buffer before adding picture
            slide.shapes.add_picture(image_stream, final_img_left, final_img_top, width=new_w, height=new_h)
        except Exception as e:
            st.error(f"Could not add image to slide: {e}")
            
    else:
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(5.8))
        text_frame = text_box.text_frame
        text_frame.text = content_chunk
        text_frame.paragraphs[0].font.size = Pt(18)
        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

def generate_ppt_from_text(text_input, user_images, auto_image, pexels_key, template_file):
    prompt = f"""
    Create a professional presentation outline from this text: "{text_input}".
    Generate 4 slides: one "Title Slide" and three "Content Slides".
    For each slide, specify its "Slide Type", a "Title", "Content" (as bullet points), and an "Image Search Query".
    Format the output strictly as follows, using '---' as a separator:

    Slide Type: Title Slide
    Title: [Catchy Title Here]
    Content: [A brief, engaging subtitle or key takeaway]
    Image Search Query: [A query for a powerful opening image]

    ---

    Slide Type: Content Slide
    Title: [Informative Title Here]
    Content: - First key point...
    Image Search Query: [A specific query for an image]
    """
    with st.spinner("Step 1/3: Generating presentation script with AI..."):
        content = generate_content_from_ai(prompt)
        if not content: return None
        slides_data = parse_ai_response(content)
        if not slides_data:
            st.error("AI did not return a valid structure. Please try again.")
            return None

    prs = Presentation(pptx=template_file) if template_file else Presentation()
    
    with st.spinner("Step 2/3: Populating slides and fetching images..."):
        user_image_idx = 0
        for slide_info in slides_data:
            image_to_add = None
            if user_image_idx < len(user_images):
                image_to_add = user_images[user_image_idx]
                user_image_idx += 1
            elif auto_image and pexels_key:
                image_to_add = fetch_image_from_pexels(pexels_key, slide_info['query'])
            
            content_chunks = split_text_for_slides(slide_info['content'])
            
            for i, chunk in enumerate(content_chunks):
                slide_title = slide_info['title'] if i == 0 else f"{slide_info['title']} (cont.)"
                image_for_this_slide = image_to_add if i == 0 else None
                add_slide_with_layout(prs, slide_title, chunk, image_for_this_slide)

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
        "Upload a Design Template (.pptx)",
        type=['pptx'],
        help="Upload a .pptx file. The AI will use its slide masters, fonts, and colors."
    )
    if template_file:
        st.sidebar.success(f"Using template: {template_file.name}")
    else:
        st.sidebar.info("No template uploaded. Using default design.")

    st.sidebar.header("🖼️ Image Options")
    auto_image = st.sidebar.checkbox("Find images automatically", value=True)
    user_images = st.sidebar.file_uploader("Or upload your own images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
    
    st.header("✍️ Content Input")
    st.info("The AI will automatically split long content across multiple slides.")
    text_input = st.text_area("Enter the topic or text for your presentation:", height=200, 
        value="Dogs – Man’s Best Friend. Discuss their evolution from wolves, the variety of breeds, and their crucial roles in society as companions, workers, and heroes. Emphasize the deep bond between humans and dogs, the importance of responsible pet ownership, and the need for animal welfare. Also cover their use as service animals, therapy dogs, and in law enforcement.")
    
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
