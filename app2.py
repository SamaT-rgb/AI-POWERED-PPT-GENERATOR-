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
    if not gemini_api_key: st.error("GEMINI_API_KEY not found."); st.stop()
    if not pexels_api_key: st.warning("PEXELS_API_KEY not found. Image search disabled.")
    genai.configure(api_key=gemini_api_key)
    return pexels_api_key

# --- Image Generation and Fetching ---
def create_placeholder_image(text):
    width, height = 1200, 800; img = Image.new('RGB', (width, height), color='#E9ECEF')
    draw = ImageDraw.Draw(img)
    try: font = ImageFont.truetype("arial.ttf", 60)
    except IOError: font = ImageFont.load_default(size=60)
    lines = text.split('\n'); total_h = sum(draw.textbbox((0,0), l, font=font)[3] for l in lines)
    y = (height - total_h) / 2
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font, align="center")
        pos = ((width - (bbox[2] - bbox[0])) / 2, y)
        draw.text(pos, line, fill='#6C757D', font=font, align="center"); y += bbox[3] + 10
    img_buffer = BytesIO(); img.save(img_buffer, format='PNG'); img_buffer.seek(0)
    return img_buffer

def fetch_image_from_pexels(api_key, query):
    if not api_key or not query: return create_placeholder_image(f"Image for:\n{query}")
    headers, params = {"Authorization": api_key}, {"query": query, "per_page": 1}
    try:
        response = requests.get("https://api.pexels.com/v1/search", headers=headers, params=params, timeout=10)
        response.raise_for_status(); data = response.json()
        if not data.get("photos"):
            st.warning(f"No image for '{query}'. Using placeholder."); return create_placeholder_image(f"No image for:\n'{query}'")
        img_resp = requests.get(data["photos"][0]["src"]["large"], timeout=10)
        img_resp.raise_for_status(); return BytesIO(img_resp.content)
    except requests.exceptions.RequestException as e:
        st.error(f"Pexels API Error: {e}. Using placeholder."); return create_placeholder_image(f"API Error for:\n'{query}'")

# --- Core AI and PowerPoint Generation ---

def generate_content_from_ai(prompt):
    try: model = genai.GenerativeModel('gemini-1.5-flash'); return model.generate_content(prompt).text
    except Exception as e: st.error(f"AI model error: {e}"); return None

def parse_ai_response(content):
    slides_data = []
    slide_blocks = re.findall(r'Slide Type:\s*(.*?)\nTitle:\s*(.*?)\nContent:\s*(.*?)\nImage Search Query:\s*(.*?)(?=\n\n---|\Z)', content, re.DOTALL)
    for type, title, cont, query in slide_blocks:
        slides_data.append({
            "type": type.strip(), "title": title.strip(),
            "content": cont.strip().replace('- ', '\n• ').lstrip('\n'), "query": query.strip()
        })
    return slides_data

def split_text_for_slides(text, max_lines=8, max_chars_per_line=90):
    chunks, current_lines = [], []; text_lines = text.split('\n')
    for line in text_lines:
        while len(line) > max_chars_per_line:
            split_pos = line.rfind(' ', 0, max_chars_per_line); split_pos = split_pos if split_pos != -1 else max_chars_per_line
            current_lines.append(line[:split_pos]); line = line[split_pos:].lstrip()
            if len(current_lines) >= max_lines: chunks.append("\n".join(current_lines)); current_lines = []
        current_lines.append(line)
        if len(current_lines) >= max_lines: chunks.append("\n".join(current_lines)); current_lines = []
    if current_lines: chunks.append("\n".join(current_lines))
    return chunks if chunks else [""]

# --- LAYOUT ENGINE 1: FOR TEMPLATES ---
def add_slide_using_template_layout(prs, slide_info, content_chunk, image_stream=None):
    layout_names = {"Title Slide": ["Title Slide", "Title"], "Content Slide": ["Title and Content", "Picture with Caption", "Two Content"]}.get(slide_info['type'], ["Title and Content"])
    chosen_layout = next((l for name in layout_names for l in prs.slide_layouts if l.name == name), prs.slide_layouts[1])
    slide = prs.slides.add_slide(chosen_layout)
    placeholders = {'title': getattr(slide.shapes, 'title', None), 'body': None, 'pic': None}
    for shape in slide.placeholders:
        ph_type = shape.placeholder_format.type
        if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) and not placeholders['title']: placeholders['title'] = shape
        elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT): placeholders['body'] = shape
        elif ph_type == PP_PLACEHOLDER.PICTURE: placeholders['pic'] = shape
    if placeholders['title']: placeholders['title'].text = slide_info['title']
    if placeholders['body']:
        placeholders['body'].text = content_chunk
        placeholders['body'].text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE; placeholders['body'].text_frame.word_wrap = True
    if image_stream and placeholders['pic']:
        try: placeholders['pic'].insert_picture(image_stream)
        except Exception as e: st.error(f"Could not insert picture into placeholder: {e}")

# --- LAYOUT ENGINE 2: FOR BLANK PRESENTATIONS ---
def add_slide_with_programmatic_layout(prs, title, content_chunk, image_stream=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1.0))
    p = title_box.text_frame.paragraphs[0]; p.text, p.font.size, p.font.bold, p.alignment = title, Pt(36), True, PP_ALIGN.LEFT
    if image_stream:
        text_container = {"left": Inches(0.5), "top": Inches(1.2), "width": Inches(4.5), "height": Inches(5.8)}
        img_container = {"left": Inches(5.5), "top": Inches(1.5), "width": Inches(4.0), "height": Inches(5.5)}
        text_box = slide.shapes.add_textbox(**text_container)
        tf = text_box.text_frame; tf.text, tf.word_wrap, tf.auto_size = content_chunk, True, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.paragraphs[0].font.size = Pt(16)
        try:
            image_stream.seek(0); img = Image.open(image_stream)
            aspect_ratio = float(img.width) / img.height; container_aspect = float(img_container["width"]) / img_container["height"]
            image_stream.seek(0)
            if aspect_ratio > container_aspect: pic = slide.shapes.add_picture(image_stream, img_container["left"], img_container["top"], width=img_container["width"])
            else: pic = slide.shapes.add_picture(image_stream, img_container["left"], img_container["top"], height=img_container["height"])
            pic.left = img_container["left"] + (img_container["width"] - pic.width) / 2; pic.top = img_container["top"] + (img_container["height"] - pic.height) / 2
        except Exception as e: st.error(f"Could not add image: {e}")
    else:
        text_container = {"left": Inches(0.5), "top": Inches(1.2), "width": Inches(9), "height": Inches(5.8)}
        text_box = slide.shapes.add_textbox(**text_container)
        tf = text_box.text_frame; tf.text, tf.word_wrap, tf.auto_size = content_chunk, True, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.paragraphs[0].font.size = Pt(18)

# --- MAIN GENERATION FUNCTION ---
def generate_ppt_from_text(text_input, user_images, auto_image, pexels_key, template_file):
    prompt = f'Create a professional presentation outline from this text: "{text_input}". Generate 4 slides: one "Title Slide" and three "Content Slides". For each slide, specify its "Slide Type", a "Title", "Content", and an "Image Search Query". Format the output strictly as follows, using \'---\' as a separator:\n\nSlide Type: Title Slide\nTitle: [Catchy Title Here]\nContent: [A brief subtitle]\nImage Search Query: [A powerful opening image query]\n---\nSlide Type: Content Slide\nTitle: [Informative Title Here]\nContent: - First key point...\nImage Search Query: [A specific image query]'
    with st.spinner("Step 1/3: Generating presentation script..."):
        content = generate_content_from_ai(prompt)
        if not content: return None
        slides_data = parse_ai_response(content)
        if not slides_data: st.error("AI did not return a valid structure."); return None

    prs = Presentation(template_file) if template_file else Presentation()

    # --- THIS IS THE CORRECTED, STABLE WAY TO DELETE ALL SLIDES ---
    if template_file:
        # Access the low-level XML object that holds the slide list
        xml_slides = prs.slides._sldIdLst
        # Create a list of all slide elements to remove, as we can't modify the list while iterating
        slides_to_remove = list(xml_slides)
        for sld in slides_to_remove:
            xml_slides.remove(sld)

    with st.spinner("Step 2/3: Populating slides..."):
        user_image_idx = 0
        for slide_info in slides_data:
            image_to_add = None
            if user_image_idx < len(user_images):
                image_to_add = user_images[user_image_idx]; user_image_idx += 1
            elif auto_image and pexels_key:
                image_to_add = fetch_image_from_pexels(pexels_key, slide_info['query'])
            content_chunks = split_text_for_slides(slide_info['content'])
            for i, chunk in enumerate(content_chunks):
                current_slide_info = slide_info.copy()
                if i > 0: current_slide_info['title'] = f"{slide_info['title']} (cont.)"
                image_for_this_slide = image_to_add if i == 0 else None
                if template_file:
                    add_slide_using_template_layout(prs, current_slide_info, chunk, image_for_this_slide)
                else:
                    add_slide_with_programmatic_layout(prs, current_slide_info['title'], chunk, image_for_this_slide)

    with st.spinner("Step 3/3: Finalizing presentation..."):
        ppt_buffer = BytesIO(); prs.save(ppt_buffer); ppt_buffer.seek(0)
    return ppt_buffer

# --- Streamlit UI ---
def main():
    pexels_key = configure_app()
    st.sidebar.header("🎨 Design Options")
    template_file = st.sidebar.file_uploader("Upload a Design Template (.pptx)", type=['pptx'], help="Upload a .pptx file. Its slide masters, fonts, and colors will be used.")
    if template_file: st.sidebar.success(f"Using template: {template_file.name}")
    else: st.sidebar.info("No template uploaded. Using default programmatic design.")
    st.sidebar.header("🖼️ Image Options")
    auto_image = st.sidebar.checkbox("Find images automatically", value=True)
    user_images = st.sidebar.file_uploader("Or upload your own images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
    st.header("✍️ Content Input")
    st.info("The AI will automatically split long content across multiple slides.")
    text_input = st.text_area("Enter the topic or text for your presentation:", height=200, value="The importance of renewable energy. Discuss solar, wind, and hydro power. Cover the environmental benefits, economic advantages, and future challenges of transitioning to a green economy.")
    if st.button("🚀 Generate Presentation"):
        if text_input:
            ppt_buffer = generate_ppt_from_text(text_input, user_images, auto_image, pexels_key, template_file)
            if ppt_buffer:
                st.session_state['generated_ppt'] = ppt_buffer; st.session_state['ppt_name'] = "AI_Generated_Presentation.pptx"
                st.success("Presentation generated successfully!")
        else: st.warning("Please enter some text.")
    if 'generated_ppt' in st.session_state and st.session_state['generated_ppt']:
        st.download_button(label="⬇️ Download Presentation", data=st.session_state['generated_ppt'], file_name=st.session_state['ppt_name'], mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", on_click=lambda: st.session_state.update({'generated_ppt': None}))

if __name__ == "__main__":
    main()
