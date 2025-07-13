import os
import time
import re
import requests
from io import BytesIO
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import PP_PLACEHOLDER
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

# --- Image Generation and Fetching ---
def create_placeholder_image(text):
    width, height = 1200, 800
    img = Image.new('RGB', (width, height), color='#E9ECEF') # Light grey for a softer look
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 60)
    except IOError:
        font = ImageFont.load_default(size=60)
    
    lines = text.split('\n')
    total_text_height = sum(draw.textbbox((0,0), line, font=font)[3] for line in lines)
    current_y = (height - total_text_height) / 2

    for line in lines:
        text_bbox = draw.textbbox((0, 0), line, font=font, align="center")
        position = ((width - (text_bbox[2] - text_bbox[0])) / 2, current_y)
        draw.text(position, line, fill='#6C757D', font=font, align="center")
        current_y += text_bbox[3] + 10 # Add spacing between lines

    img_buffer = BytesIO()
    img.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    return img_buffer

def fetch_image_from_pexels(api_key, query):
    if not api_key or not query:
        return create_placeholder_image(f"Image for:\n{query}")
    headers = {"Authorization": api_key}
    params = {"query": query, "per_page": 1, "page": 1}
    try:
        response = requests.get("https://api.pexels.com/v1/search", headers=headers, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if not data.get("photos"):
            st.warning(f"No image found for query: '{query}'. Using placeholder.")
            return create_placeholder_image(f"No image for:\n'{query}'")
        
        photo_url = data["photos"][0]["src"]["large"]
        image_response = requests.get(photo_url, timeout=10)
        image_response.raise_for_status()
        return BytesIO(image_response.content)
    except requests.exceptions.RequestException as e:
        st.error(f"Pexels API Error for '{query}': {e}. Using placeholder.")
        return create_placeholder_image(f"API Error for:\n'{query}'")

# --- Core AI and PowerPoint Generation ---

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

def split_text_for_slides(text, max_lines=8, max_chars_per_line=90):
    """Splits long text into chunks, respecting line and character limits."""
    chunks = []
    current_chunk_lines = []
    for line in text.split('\n'):
        # Further split long lines
        while len(line) > max_chars_per_line:
            split_pos = line.rfind(' ', 0, max_chars_per_line)
            if split_pos == -1: split_pos = max_chars_per_line
            current_chunk_lines.append(line[:split_pos])
            line = line[split_pos:].lstrip()
            if len(current_chunk_lines) >= max_lines:
                chunks.append("\n".join(current_chunk_lines))
                current_chunk_lines = []
        current_chunk_lines.append(line)
        if len(current_chunk_lines) >= max_lines:
            chunks.append("\n".join(current_chunk_lines))
            current_chunk_lines = []
            
    if current_chunk_lines:
        chunks.append("\n".join(current_chunk_lines))
    return chunks

# --- LAYOUT ENGINE 1: FOR TEMPLATES ---
def add_slide_using_template_layout(prs, slide_info, content_chunk, image_stream=None):
    """Finds the best layout in the template and populates its placeholders."""
    possible_layouts = {
        "Title Slide": ["Title Slide", "Title"],
        "Content Slide": ["Title and Content", "Two Content", "Picture with Caption", "Content with Caption"]
    }.get(slide_info['type'], ["Title and Content"])
    
    chosen_layout = next((l for l_name in possible_layouts for l in prs.slide_layouts if l.name == l_name), prs.slide_layouts[1])

    slide = prs.slides.add_slide(chosen_layout)

    # Populate placeholders intelligently
    title_ph = next((s for s in slide.placeholders if 'Title' in s.name or s.placeholder_format.idx == 0), slide.shapes.title)
    body_ph = next((s for s in slide.placeholders if 'Body' in s.name or 'Content' in s.name or s.placeholder_format.idx in [1, 10, 11]), None)
    pic_ph = next((s for s in slide.placeholders if 'Picture' in s.name or s.placeholder_format.type == PP_PLACEHOLDER.PICTURE), None)
    
    if title_ph: title_ph.text = slide_info['title']
    if body_ph:
        body_ph.text = content_chunk
        body_ph.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        body_ph.text_frame.word_wrap = True
    if image_stream and pic_ph:
        try:
            pic_ph.insert_picture(image_stream)
        except Exception as e:
            st.error(f"Could not insert picture into placeholder: {e}")

# --- LAYOUT ENGINE 2: FOR BLANK PRESENTATIONS ---
def add_slide_with_programmatic_layout(prs, title, content_chunk, image_stream=None):
    """Creates a clean, professional layout from scratch on a blank slide."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1.0))
    title_frame = title_box.text_frame
    title_frame.text = title
    p = title_frame.paragraphs[0]
    p.font.size, p.font.bold, p.alignment = Pt(36), True, PP_ALIGN.LEFT
    
    if image_stream:
        text_container = {"left": Inches(0.5), "top": Inches(1.2), "width": Inches(4.5), "height": Inches(5.8)}
        img_container = {"left": Inches(5.5), "top": Inches(1.5), "width": Inches(4.0), "height": Inches(5.5)}

        text_box = slide.shapes.add_textbox(**text_container)
        text_frame = text_box.text_frame
        text_frame.text = content_chunk
        text_frame.word_wrap, text_frame.auto_size = True, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        text_frame.paragraphs[0].font.size = Pt(16)
        
        try:
            image_stream.seek(0)
            img = Image.open(image_stream)
            img_w, img_h = img.size
            aspect_ratio = float(img_w) / img_h
            container_aspect = float(img_container["width"]) / img_container["height"]
            image_stream.seek(0)

            if aspect_ratio > container_aspect:
                pic = slide.shapes.add_picture(image_stream, img_container["left"], img_container["top"], width=img_container["width"])
            else:
                pic = slide.shapes.add_picture(image_stream, img_container["left"], img_container["top"], height=img_container["height"])

            pic.left = img_container["left"] + (img_container["width"] - pic.width) / 2
            pic.top = img_container["top"] + (img_container["height"] - pic.height) / 2
        except Exception as e:
            st.error(f"Could not add image to slide: {e}")
    else:
        text_container = {"left": Inches(0.5), "top": Inches(1.2), "width": Inches(9), "height": Inches(5.8)}
        text_box = slide.shapes.add_textbox(**text_container)
        text_frame = text_box.text_frame
        text_frame.text = content_chunk
        text_frame.word_wrap, text_frame.auto_size = True, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        text_frame.paragraphs[0].font.size = Pt(18)

# --- MAIN GENERATION FUNCTION ---
def generate_ppt_from_text(text_input, user_images, auto_image, pexels_key, template_file):
    prompt = f"""
    Create a professional presentation outline from this text: "{text_input}".
    Generate 4 slides: one "Title Slide" and three "Content Slides".
    For each slide, specify its "Slide Type", a "Title", "Content", and an "Image Search Query".
    Format the output strictly as follows, using '---' as a separator:

    Slide Type: Title Slide
    Title: [Catchy Title Here]
    Content: [A brief subtitle]
    Image Search Query: [A powerful opening image query]
    ---
    Slide Type: Content Slide
    Title: [Informative Title Here]
    Content: - First key point...
    Image Search Query: [A specific image query]
    """
    with st.spinner("Step 1/3: Generating presentation script..."):
        content = generate_content_from_ai(prompt)
        if not content: return None
        slides_data = parse_ai_response(content)
        if not slides_data:
            st.error("AI did not return a valid structure. Please try again.")
            return None

    prs = Presentation(pptx=template_file) if template_file else Presentation()

    if template_file:
        # Clear existing demo slides from the template
        xml_slides = prs.slides._sldIdLst
        slides_to_remove = list(xml_slides)
        for sld in slides_to_remove:
            xml_slides.remove(sld)
    
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
                slide_info_for_chunk = slide_info.copy()
                if i > 0:
                    slide_info_for_chunk['title'] = f"{slide_info['title']} (cont.)"
                
                image_for_this_slide = image_to_add if i == 0 else None
                
                if template_file:
                    add_slide_using_template_layout(prs, slide_info_for_chunk, chunk, image_for_this_slide)
                else:
                    add_slide_with_programmatic_layout(prs, slide_info_for_chunk['title'], chunk, image_for_this_slide)

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
        help="Upload a .pptx file. The app will use its slide masters, fonts, and colors."
    )
    if template_file:
        st.sidebar.success(f"Using template: {template_file.name}")
    else:
        st.sidebar.info("No template uploaded. Using default programmatic design.")

    st.sidebar.header("🖼️ Image Options")
    auto_image = st.sidebar.checkbox("Find images automatically", value=True)
    user_images = st.sidebar.file_uploader("Or upload your own images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
    
    st.header("✍️ Content Input")
    st.info("The AI will automatically split long content across multiple slides.")
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
            st.warning("Please enter some text.")

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
