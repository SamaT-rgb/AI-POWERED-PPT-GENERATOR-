# ✨ AI PowerPoint Pro

A Streamlit web application that automatically generates professional PowerPoint presentations from user-provided text. It leverages Google's Gemini AI for content creation, Pexels for stock photos, and allows for custom design templates to ensure a high-quality, branded output.

---

## 📋 Features

-   **AI-Powered Content:** Generates slide titles, bullet points, and subtitles automatically from a block of text.
-   **Custom Design Templates:** Upload your own `.pptx` file to be used as a design template, preserving your company's branding, fonts, and colors.
-   **Automatic Image Search:** Intelligently creates search queries for each slide and fetches relevant, high-quality stock photos from Pexels.com.
-   **Robust Image Fallback:** If an image cannot be found or the API fails, a clean placeholder image is generated on the fly, preventing errors and layout issues.
-   **User Image Upload:** Option to override the automatic search and upload your own images for each slide.
-   **Direct Download:** Generates and downloads the final `.pptx` file directly from the browser.

---

## 🚀 Setup and Installation Guide

Follow these steps to get the project running on your local machine. This guide is designed to be a clean start.

### 1. Prerequisites

-   Python 3.8+
-   `pip` (Python package installer)

### 2. Set Up a Virtual Environment

It is highly recommended to use a virtual environment to keep project dependencies isolated.

```bash
# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
