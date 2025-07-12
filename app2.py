import os
import base64
from io import BytesIO
import time
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from dotenv import load_dotenv
import google.generativeai as genai

# --- Configuration and Setup ---

def configure_app():
    """Set up page configuration and load API key."""
    st.set_page_config(
        page_title="AI PowerPoint Generator",
        page_icon="🤖",
        layout="centered"
    )
    st.title("PowerPoint Presentation Generator")
    
    # Load API key from .env file or Streamlit secrets
    load_dotenv()
    api_key = os.getenv("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY")
    
    if not api_key:
        st.error("GEMINI_API_KEY not found. Please set it in your .env file or Streamlit secrets.")
        st.stop()
        
    genai.configure(api_key=api_key)

# --- Gemini AI Functions ---

def generate_content_from_ai(prompt, max_retries=3, retry_delay=60):
    """
    Generates content using the Gemini AI model with robust retry logic.

    Args:
        prompt (str): The prompt to send to the AI model.
        max_retries (int): The maximum number of retries for API calls.
        retry_delay (int): The delay in seconds between retries.

    Returns:
        str: The generated text content or an error message.
    """
    try:
        model = genai.GenerativeModel('gemini-2.0-flash')
        for attempt in range(max_retries):
            try:
                response = model.generate_content(prompt)
                return response.text
            except Exception as e:
                if "429" in str(e): # Handle rate limiting
                    st.warning(f"Rate limit exceeded. Retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                    time.sleep(retry_delay)
                else:
                    st.error(f"An unexpected error occurred with the AI model: {e}")
                    return None
        st.error("Failed to generate content after multiple retries due to rate limiting.")
        return None
    except Exception as e:
        st.error(f"Failed to initialize the generative model: {e}")
        return None


# --- PowerPoint Generation Functions ---

def add_slide_to_ppt(prs, title, content):
    """Adds a new slide with a title and content to the presentation."""
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add and format title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(32)
    title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Add and format content
    content_shape = slide.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.text = content
    for paragraph in text_frame.paragraphs:
        paragraph.font.size = Pt(18)

def create_chart_image(df, chart_type, x_col, y_col):
    """Creates a chart from a DataFrame and returns it as an image buffer."""
    try:
        plt.style.use('seaborn-v0_8-talk')
        fig, ax = plt.subplots(figsize=(8, 5))
        
        if chart_type == "Bar":
            sns.barplot(data=df, x=x_col, y=y_col, ax=ax, palette="viridis")
        elif chart_type == "Line":
            sns.lineplot(data=df, x=x_col, y=y_col, ax=ax, color="dodgerblue", marker='o')

        ax.set_title(f"{chart_type} Chart: {y_col} by {x_col}", fontsize=16)
        ax.set_xlabel(x_col, fontsize=12)
        ax.set_ylabel(y_col, fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format="png", dpi=300)
        plt.close(fig)
        img_buffer.seek(0)
        return img_buffer
    except Exception as e:
        st.error(f"Error creating chart: {e}")
        return None

def generate_ppt_from_text(text_input):
    """Generates a PowerPoint presentation from user-provided text."""
    prompt = f"""
    Create a PowerPoint presentation based on the following text.
    Generate a title slide and at least 3 content slides.
    For each slide, provide a short, clear title and 3-5 bullet points.
    Format the output strictly as follows, using '---' as a slide separator:
    
    Slide 1 Title: [Your Title Here]
    - Bullet point 1
    - Bullet point 2
    - Bullet point 3
    ---
    Slide 2 Title: [Your Title Here]
    - Bullet point 1
    - Bullet point 2
    - Bullet point 3
    
    Text: "{text_input}"
    """
    
    with st.spinner("Generating presentation content with AI..."):
        content = generate_content_from_ai(prompt)
        if not content:
            st.error("Could not generate presentation content.")
            return None

    prs = Presentation()
    slides = content.strip().split("\n---\n")
    
    for i, slide_content in enumerate(slides):
        lines = slide_content.strip().split("\n")
        if not lines:
            continue
            
        title = lines[0].replace(f"Slide {i+1} Title:", "").strip()
        body = "\n".join(lines[1:]).strip()
        add_slide_to_ppt(prs, title, body)
        
    return prs

def generate_ppt_from_csv(df, chart_type, x_col, y_col):
    """Generates a PowerPoint presentation from a CSV file, including a summary and a chart."""
    df_string_for_prompt = df.head().to_string()
    prompt = f"""
    Analyze the following data and provide a summary slide for a PowerPoint presentation.
    The slide should have a title and 3-4 key insights as bullet points.
    Format the output strictly as follows:

    Title: [Your Summary Title]
    - Key insight 1
    - Key insight 2
    - Key insight 3

    Data:
    {df_string_for_prompt}
    """
    with st.spinner("Analyzing data and generating summary..."):
        summary_content = generate_content_from_ai(prompt)
        if not summary_content:
            st.error("Could not generate data summary.")
            return None

    prs = Presentation()
    
    # Add Summary Slide
    lines = summary_content.strip().split("\n")
    title = lines[0].replace("Title:", "").strip()
    body = "\n".join(lines[1:]).strip()
    add_slide_to_ppt(prs, title, body)
    
    # Add Chart Slide
    with st.spinner("Creating data visualization..."):
        chart_img = create_chart_image(df, chart_type, x_col, y_col)
    
    if chart_img:
        chart_slide_layout = prs.slide_layouts[5] # Title only layout
        slide = prs.slides.add_slide(chart_slide_layout)
        slide.shapes.title.text = f"{chart_type} Chart: {y_col} by {x_col}"
        slide.shapes.add_picture(chart_img, Inches(1), Inches(1.5), width=Inches(8))
    else:
        st.warning("Could not create the chart, but the summary slide was generated.")

    return prs

# --- Streamlit UI ---

def ui_text_section():
    """Renders the UI section for text-based presentation generation."""
    st.header("1. Generate Presentation from Text")
    text_input = st.text_area("Enter your text here:", height=200, 
        value="Lions are majestic big cats, often called the 'king of the jungle.' They are native to Africa and India and are the most social of all big cats, living in groups called prides. A pride consists of related females, their offspring, and a few adult males. Male lions are distinguished by their impressive manes, which signal their health and fitness to other lions. Lionesses are the primary hunters, working together to prey on large herbivores like zebras and wildebeest.")
    
    if st.button("Generate from Text"):
        if text_input:
            presentation = generate_ppt_from_text(text_input)
            if presentation:
                # Store presentation in session state
                st.session_state['generated_ppt'] = presentation
                st.session_state['ppt_name'] = "Text_Based_Presentation.pptx"
                st.success("Presentation generated successfully!")
        else:
            st.warning("Please enter some text to generate a presentation.")

def ui_csv_section():
    """Renders the UI section for CSV-based presentation generation."""
    st.header("2. Generate Presentation from CSV")
    csv_file = st.file_uploader("Upload a CSV file", type=["csv"])
    
    if csv_file:
        try:
            df = pd.read_csv(csv_file)
            st.dataframe(df.head())
            
            columns = df.columns.tolist()
            col1, col2, col3 = st.columns(3)
            with col1:
                chart_type = st.selectbox("Chart Type", ["Bar", "Line"], key="chart_type")
            with col2:
                x_col = st.selectbox("X-axis", columns, key="x_col")
            with col3:
                y_col = st.selectbox("Y-axis", columns, key="y_col")
                
            if st.button("Generate from CSV"):
                if x_col and y_col:
                    presentation = generate_ppt_from_csv(df, chart_type, x_col, y_col)
                    if presentation:
                        st.session_state['generated_ppt'] = presentation
                        st.session_state['ppt_name'] = "CSV_Based_Presentation.pptx"
                        st.success("Presentation generated successfully!")
                else:
                    st.warning("Please select columns for the X and Y axes.")
        except Exception as e:
            st.error(f"Error processing CSV file: {e}")

def main():
    """Main function to run the Streamlit app."""
    configure_app()

    # Initialize session state
    if 'generated_ppt' not in st.session_state:
        st.session_state['generated_ppt'] = None
    if 'ppt_name' not in st.session_state:
        st.session_state['ppt_name'] = ""

    # UI Layout
    tab1, tab2 = st.tabs(["From Text", "From CSV"])

    with tab1:
        ui_text_section()
        
    with tab2:
        ui_csv_section()

    # Download button - appears if a presentation is ready in session state
    if st.session_state['generated_ppt']:
        with BytesIO() as ppt_buffer:
            st.session_state['generated_ppt'].save(ppt_buffer)
            ppt_buffer.seek(0)
            
            st.download_button(
                label=f"Download {st.session_state['ppt_name']}",
                data=ppt_buffer.getvalue(),
                file_name=st.session_state['ppt_name'],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                on_click=lambda: st.session_state.update({'generated_ppt': None, 'ppt_name': ""}) # Clear state after click
            )
            
if __name__ == "__main__":
    main()
