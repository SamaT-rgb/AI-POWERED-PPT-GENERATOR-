# ✨ AI PowerPoint Pro

A **Streamlit** web application that automatically generates professional PowerPoint presentations from user-provided text.  
It leverages **Google's Gemini AI** for content creation, **Pexels** for stock photos, and supports **custom design templates** to ensure a high-quality, branded output.

---

## 📋 Features

- **AI-Powered Content:** Automatically generates slide titles, bullet points, and subtitles from a block of text.  
- **Custom Design Templates:** Upload your own `.pptx` file as a design template, preserving branding, fonts, and colors.  
- **Automatic Image Search:** Generates smart search queries for each slide and fetches high-quality stock photos from **Pexels.com**.  
- **Robust Image Fallback:** If an image search fails, a clean placeholder image is generated automatically.  
- **User Image Upload:** Option to override the automatic search by uploading your own images.  
- **Direct Download:** Creates and downloads the final `.pptx` file directly in your browser.

---

## 📄 Example Output

You can **view or download** an example generated presentation here:  
[📂 Roman Empire Presentation (Example)](https://github.com/SamaT-rgb/AI-POWERED-PPT-GENERATOR-/raw/main/Roman_Empire_Presentation%20(1).pptx)

---

## 🚀 Setup and Installation Guide

Follow these steps to run the project locally.

### 1. Prerequisites
- Python **3.8+**
- `pip` (Python package installer)

### 2. Create and Activate Virtual Environment
It’s recommended to use a virtual environment to isolate dependencies.

```bash
# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

### 3. Install Dependencies
bash
Copy
Edit
pip install -r requirements.txt

### 4. Run the Application
bash
Copy
Edit
streamlit run app.py
