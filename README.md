# Your Text, Your Style – AI Presentation Generator

An intelligent web application that transforms raw text, prose, or markdown into a fully formatted PowerPoint presentation, dynamically adapting to the visual style of your own uploaded template.

---

### 🚀 Live Demo

**https://huggingface.co/spaces/23f3004315/b1**

### ✨ Key Features

* **Multi-Provider AI:** Supports leading LLMs including **Google Gemini**, **OpenAI (GPT-4o)**, and **Anthropic Claude**.
* **True Template Adaptation:** The app intelligently analyzes the slide layouts available in *your* uploaded `.pptx` or `.potx` file.
* **Content-Aware Structuring:** The AI parses your text and intelligently maps it to appropriate titles and bullet points for the detected layouts.
* **Markdown Support:** Paste content directly from markdown documents; the app will clean and process it.
* **AI Guidance:** Steer the AI with optional prompts (e.g., "make this a sales pitch") to control the tone and structure.
* **Secure & Private:** Your API keys are used only for the current session and are never stored or logged.

---

### 🔧 How It Works (Technical Write-up)

This application bridges the gap between raw text and a styled presentation through a two-step, AI-driven process: **Content Structuring** and **Style Application**.

#### 1. How Input Text is Parsed and Mapped to Slides

The core of the content mapping is handled by a Large Language Model (LLM). Instead of simply passing the text to the AI, the application first performs a crucial preliminary step: it analyzes the user-uploaded PowerPoint template to identify all available slide layouts (e.g., 'Title Slide', 'Title and Content', 'Two Column Text').

The names of these layouts are then dynamically inserted into a detailed system prompt that instructs the AI to act as an expert presentation creator. The prompt commands the AI to read the user's raw text, break it down into logical segments, and structure its response into a strict JSON format. For each segment, the AI must choose the most appropriate layout from the provided list and populate its content fields (like `title` and `body`). This turns the unstructured prose into a structured, slide-by-slide plan, with the AI making editorial decisions on how to best summarize and present the information.

#### 2. How the App Applies the Visual Style of the Template

The application achieves style application not by extracting and reapplying individual fonts or colors, but by leveraging the template's built-in **Master Slides and Layouts**. A PowerPoint template (`.pptx` or `.potx`) is a container for pre-designed slide layouts, each with its own placeholder arrangement, typography, color scheme, and branding (like logos).

When the application receives the structured JSON from the AI, it iterates through the plan, slide by slide. For each slide, it reads the layout name chosen by the AI (e.g., 'Title and Content') and finds the corresponding layout object within the user's template file. It then creates a new slide using that exact layout (`prs.slides.add_slide(slide_layout)`). By doing this, the new slide automatically inherits all the stylistic properties defined in the template for that specific layout. This elegant approach ensures that the final output is perfectly consistent with the user's brand and design, without needing to manually manage any style attributes.

---

### 🛠️ Setup and Running Locally

To run this application on your local machine, follow these steps:

1.  **Clone the Repository:**
    ```bash
    git clone [https://github.com/YourUsername/YourRepoName.git](https://github.com/YourUsername/YourRepoName.git)
    cd YourRepoName
    ```

2.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3.  **Install Dependencies:**
    The project's requirements are listed in the `requirements.txt` file.
    ```bash
    pip install -r requirements.txt
    ```

4.  **Run the Streamlit App:**
    ```bash
    streamlit run app.py
    ```
    The application will open in your web browser.

---

### 📜 License

This project is licensed under the MIT License.
