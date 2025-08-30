# Your Text, Your Style – AI Presentation Generator

An intelligent web application that transforms raw text, prose, or markdown into a fully formatted PowerPoint presentation, dynamically adapting to the visual style of your own uploaded template.

---

### 🚀 Live Demo

**https://huggingface.co/spaces/23f3004315/b1**

### ✨ Features

#### Core Functionality
* **Smart Text Processing:** Paste large chunks of text, markdown, or long-form prose.
* **Template Style Preservation:** Maintains your template's layouts, fonts, colors, and visual elements.
* **Multi-LLM Support:** Works with OpenAI, Anthropic Claude, and Google Gemini.
* **Intelligent Slide Mapping:** Automatically determines the optimal number of slides based on content.
* **Image Reuse:** Extracts and reuses images from your uploaded template.

#### Enhanced Features
* **Speaker Notes Generation:** AI-powered speaker notes for better presentation delivery.
* **Real-time Validation:** Instant feedback on requirements and readiness.
* **Slide Preview:** See your presentation structure before downloading.
* **Style Templates:** Quick-select common presentation styles (sales pitch, technical, etc.).
* **Markdown Support:** Full markdown formatting detection and processing.

---

### 🔧 Technical Write-up: AI Presentation Generator

#### How Input Text is Parsed and Mapped to Slides
Our application employs a sophisticated multi-stage approach to transform unstructured text into well-organized presentation slides:

1.  **Content Analysis & Preprocessing** The system first analyzes the input text to detect formatting patterns and structure. When markdown is detected (headers, bold text, bullet points), it's processed to extract hierarchical information while preserving semantic meaning. Long-form prose undergoes natural language processing to identify topic boundaries and key concepts.

2.  **LLM-Powered Intelligent Segmentation** Rather than using fixed slide counts, we leverage large language models (OpenAI, Anthropic Claude, or Google Gemini) to intelligently parse content based on:
    * **Semantic coherence:** Related concepts are grouped together.
    * **Content density:** Appropriate information volume per slide.
    * **Logical flow:** Maintains narrative progression.
    * **Presentation context:** Considers user guidance (e.g., "sales pitch" vs "technical deep-dive").

    The LLM receives the template's available layouts and generates a structured JSON response containing slide-by-slide breakdowns with titles and bullet points optimized for visual presentation.

3.  **Dynamic Slide Allocation** The system determines slide count based on content complexity and specified presentation style. Technical presentations might have more detailed slides, while executive summaries are condensed. This adaptive approach ensures optimal information density and audience engagement.

#### How Visual Style and Assets are Applied from Templates
Our template preservation system ensures generated presentations maintain professional consistency with uploaded templates:

1.  **Template Analysis & Asset Extraction** Upon template upload, the system performs comprehensive analysis:
    * **Layout enumeration:** Identifies all slide layouts with their placeholders.
    * **Image extraction:** Locates and extracts embedded images from slides and master layouts.
    * **Style preservation:** Maintains font families, color schemes, and formatting rules.
    * **Placeholder mapping:** Catalogs available content areas (titles, body text, images).

2.  **Intelligent Layout Matching** The system employs fuzzy matching algorithms to pair generated content with appropriate template layouts:
    * **Semantic matching:** "Title slide" content maps to title layouts.
    * **Content-type alignment:** Bullet-point content matches content/body layouts.
    * **Fallback mechanisms:** Ensures graceful degradation when perfect matches aren't available.

3.  **Asset Integration & Style Application** Template images are strategically reused where contextually appropriate, maintaining visual consistency. The original template's typography, color palette, and spacing are preserved through the `python-pptx` library's style retention capabilities.

4.  **Quality Preservation** Rather than generating new visual elements, the system focuses on intelligent content placement within existing template structures, ensuring professional output that maintains the template's intended aesthetic while accommodating new content effectively.

This approach delivers presentations that look professionally designed while being automatically generated from raw text input.

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

