# ✨ AI PowerPoint Pro

A Streamlit web application that automatically generates professional PowerPoint presentations from user-provided text. It leverages Google's Gemini AI for content creation, Pexels for stock photos, and allows for custom design templates to ensure a high-quality, branded output.

![Screenshot of the Streamlit App Interface](https://i.imgur.com/your-screenshot-url.png) 
*(This is a placeholder - you can add a real screenshot of your app running)*

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

3. Install Dependencies
Install all the required Python libraries using the requirements.txt file.
First, create a file named requirements.txt in your project folder and add the following lines:
Generated code
streamlit
python-pptx
python-dotenv
google-generativeai
pexels-api
Pillow
requests
pandas
matplotlib
seaborn
Use code with caution.
Now, run this command in your terminal:
Generated bash
pip install -r requirements.txt
Use code with caution.
Bash
4. Set Up API Keys
This is the most critical step. The application requires API keys from Google and Pexels.
Create a file named .env in the root of your project folder.
Add your keys to this file in the following format:
Generated code
GEMINI_API_KEY="your_google_ai_studio_api_key_here"
PEXELS_API_KEY="your_pexels_api_key_here"
Use code with caution.
Gemini API Key: Get it from Google AI Studio.
Pexels API Key: Get it for free from the Pexels API site.
▶️ How to Run the Application
With your virtual environment active and the .env file configured, run the following command in your terminal:
Generated bash
streamlit run app.py
Use code with caution.
Bash
(Assuming your Python script is named app.py)
Your web browser should automatically open to the application's local URL (e.g., http://localhost:8501).
💡 How to Use the App
Design Template (Highly Recommended):
In PowerPoint, create a presentation with the fonts, colors, and logos you want.
Save this file as a .pptx file.
In the app's sidebar, use the "Upload a Design Template" button to upload this .pptx file.
Image Options:
Keep "Find images automatically" checked to use the Pexels integration.
Alternatively, upload your own images. The first image you upload will be used for the first slide, the second for the second, and so on.
Content Input:
Paste the text you want to turn into a presentation into the text area.
Generate:
Click the "🚀 Generate Presentation" button. The app will show its progress through the three stages.
Download:
Once finished, a "⬇️ Download Presentation" button will appear. Click it to save your file.
