import streamlit as st
import google.generativeai as genai
import openai
import anthropic
import markdown
from pptx import Presentation
from pptx.util import Inches
import io
import json
import re
from typing import Dict, List, Optional, Any

# --- App Configuration ---
st.set_page_config(
    page_title="Your Text, Your Style – AI Presentation Generator",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
    
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
        margin: 1rem 0;
    }
    
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---

def clean_layout_name(name: str) -> str:
    """Clean and standardize layout names for better matching."""
    return re.sub(r'^Layout \d+:\s*[\'"]?|[\'"]?$', '', name).strip()

def get_ppt_layouts(template_bytes: bytes) -> Dict[str, Any]:
    """
    Analyzes PowerPoint file bytes and returns a dictionary of its layouts.
    """
    if not template_bytes:
        return {}
    
    try:
        template_stream = io.BytesIO(template_bytes)
        prs = Presentation(template_stream)
        
        layouts = {}
        for i, layout in enumerate(prs.slide_layouts):
            layout_name = f"Layout {i}: '{layout.name}'"
            clean_name = clean_layout_name(layout_name)
            
            # Skip blank layouts and add layout info
            if "blank" not in clean_name.lower():
                layouts[layout_name] = {
                    'layout_obj': layout,
                    'clean_name': clean_name,
                    'placeholders': [p.name for p in layout.placeholders if hasattr(p, 'name')]
                }
        
        return layouts
        
    except Exception as e:
        st.error(f"Error reading template file: {str(e)}")
        return {}

def process_markdown_content(content: str) -> str:
    """Convert markdown to plain text while preserving structure."""
    try:
        # Convert markdown to HTML first
        html = markdown.markdown(content)
        # Remove HTML tags for plain text
        clean_text = re.sub(r'<[^>]+>', '', html)
        # Clean up extra whitespace
        clean_text = re.sub(r'\n\s*\n', '\n\n', clean_text)
        return clean_text.strip()
    except:
        # Fallback to original content if markdown processing fails
        return content

# --- LLM API Functions ---

def generate_with_gemini(api_key: str, user_text: str, guidance: str, layout_names: List[str]) -> Optional[List[Dict]]:
    """
    Calls the Google Gemini API to structure the text into presentation slides.
    """
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        
        clean_layouts = [clean_layout_name(name) for name in layout_names]
        
        system_prompt = f"""
        You are an expert presentation creator. Your task is to analyze the provided text and structure it into a series of presentation slides.
        Available slide layouts: {', '.join(clean_layouts)}
        IMPORTANT: Structure your response as a valid JSON object with a "slides" key containing an array of slide objects.
        Each slide object must have:
        1. "layout": A string matching one of the available layouts above.
        2. "content": An object with "title" (string) and "body" (an array of bullet point strings).
        Guidelines: Create 3-8 slides. Use a title slide layout for the first slide. Keep titles concise and break body text into clear bullet points.
        """
        
        user_prompt = f"Text to convert into presentation:\n\n{user_text}\n\n"
        if guidance:
            user_prompt += f"Additional guidance: {guidance}"
        
        response = model.generate_content(
            [system_prompt, user_prompt],
            generation_config=genai.types.GenerationConfig(
                response_mime_type="application/json",
                temperature=0.7
            )
        )
        
        response_data = json.loads(response.text)
        slides = response_data.get('slides', [])
        
        if not slides:
            st.error("AI generated empty slides. Please try again with different text.")
            return None
        return slides
            
    except json.JSONDecodeError as e:
        st.error(f"AI response was not valid JSON. Please try again. Error: {str(e)}")
        st.code(response.text)
        return None
    except Exception as e:
        st.error(f"Gemini API error: {str(e)}")
        return None

def generate_with_anthropic(api_key: str, user_text: str, guidance: str, layout_names: List[str]) -> Optional[List[Dict]]:
    """
    Calls the Anthropic Claude API to structure the text into presentation slides.
    """
    try:
        client = anthropic.Anthropic(api_key=api_key)
        
        clean_layouts = [clean_layout_name(name) for name in layout_names]
        
        system_prompt = f"""
        You are an expert presentation creator. Analyze the provided text and structure it into presentation slides.
        Available layouts: {', '.join(clean_layouts)}
        Return a JSON object with a "slides" key containing an array of slide objects.
        Each slide needs:
        1. "layout": A string matching one of the available layouts.
        2. "content": An object with "title" (string) and "body" (array of strings).
        Create 3-8 slides based on content. Use clear titles and bullet points.
        """
        
        user_prompt = f"Convert this text into a presentation:\n\n{user_text}"
        if guidance:
            user_prompt += f"\n\nGuidance: {guidance}"
        
        response = client.messages.create(
            model="claude-3-5-sonnet-20240620",  # CORRECTED MODEL NAME
            max_tokens=4000,
            temperature=0.7,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )
        
        response_text = response.content[0].text
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            json_str = json_match.group()
        else:
            json_str = response_text
        
        response_data = json.loads(json_str)
        slides = response_data.get('slides', [])
        
        if not slides:
            st.error("AI generated empty slides. Please try again.")
            return None
        return slides
            
    except json.JSONDecodeError as e:
        st.error(f"Claude response was not valid JSON: {str(e)}")
        st.code(response_text)
        return None
    except anthropic.APIError as e:
        st.error(f"Anthropic API error: {e.message}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred with Anthropic API: {str(e)}")
        return None

def generate_with_openai(api_key: str, user_text: str, guidance: str, layout_names: List[str]) -> Optional[List[Dict]]:
    """
    Calls the OpenAI API to structure the text into presentation slides.
    """
    try:
        client = openai.OpenAI(api_key=api_key)
        
        clean_layouts = [clean_layout_name(name) for name in layout_names]
        
        system_prompt = f"""
        You are an expert presentation creator. Analyze the provided text and structure it into presentation slides.
        Available layouts: {', '.join(clean_layouts)}
        Return a JSON object with a "slides" key containing an array of slide objects.
        Each slide needs:
        1. "layout": A string matching one of the available layouts.
        2. "content": An object with "title" (string) and "body" (array of strings).
        Create 3-8 slides based on content. Use clear titles and bullet points.
        """
        
        user_prompt = f"Convert this text into a presentation:\n\n{user_text}"
        if guidance:
            user_prompt += f"\n\nGuidance: {guidance}"
        
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7
        )
        
        response_data = json.loads(response.choices[0].message.content)
        slides = response_data.get('slides', [])
        
        if not slides:
            st.error("AI generated empty slides. Please try again.")
            return None
        return slides
            
    except openai.APIError as e:
        st.error(f"OpenAI API error: {e.message}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred with OpenAI API: {str(e)}")
        return None

# --- Presentation Creation Functions ---

def find_best_layout_match(requested_layout: str, available_layouts: Dict) -> Optional[Any]:
    """Find the best matching layout from available layouts."""
    requested_lower = requested_layout.lower()
    
    # First, try exact match with cleaned names
    for layout_info in available_layouts.values():
        if layout_info['clean_name'].lower() == requested_lower:
            return layout_info['layout_obj']
    
    # Try partial matches based on keywords
    layout_mappings = {
        'title': ['title', 'cover', 'intro'],
        'content': ['content', 'bullet', 'text', 'body'],
        'two': ['two', 'comparison', 'split'],
        'section': ['section', 'divider', 'header']
    }
    
    # Check for keyword matches in user request vs template layout names
    for keyword, synonyms in layout_mappings.items():
        if keyword in requested_lower:
            for layout_info in available_layouts.values():
                layout_name_lower = layout_info['clean_name'].lower()
                for synonym in synonyms:
                    if synonym in layout_name_lower:
                        return layout_info['layout_obj']
    
    # Fallback: if no match, return the first available content-like layout, or any layout
    for layout_info in available_layouts.values():
         if 'content' in layout_info['clean_name'].lower():
              return layout_info['layout_obj']
    if available_layouts:
        return list(available_layouts.values())[0]['layout_obj']
    
    return None

def create_presentation(slides_data: List[Dict], template_bytes: bytes, available_layouts: Dict) -> Optional[io.BytesIO]:
    """
    Creates a PowerPoint presentation from structured slide data.
    """
    try:
        template_stream = io.BytesIO(template_bytes)
        prs = Presentation(template_stream)
        slides_created = 0
        
        for i, slide_info in enumerate(slides_data):
            try:
                layout_name = slide_info.get('layout', '')
                content = slide_info.get('content', {})
                
                slide_layout = find_best_layout_match(layout_name, available_layouts)
                if not slide_layout:
                    st.warning(f"No suitable layout found for slide {i+1} (requested '{layout_name}'). Skipping.")
                    continue
                
                slide = prs.slides.add_slide(slide_layout)
                slides_created += 1
                
                title_set, body_set = False, False
                for shape in slide.placeholders:
                    if not hasattr(shape, 'text_frame'):
                        continue
                    
                    shape_name = getattr(shape, 'name', '').lower()
                    
                    if not title_set and 'title' in shape_name and content.get('title'):
                        shape.text = content['title']
                        title_set = True
                    elif not body_set and any(k in shape_name for k in ['body', 'content', 'text']) and content.get('body'):
                        tf = shape.text_frame
                        tf.clear()
                        body_points = content.get('body', [])
                        if isinstance(body_points, list) and body_points:
                            tf.text = str(body_points[0])
                            for point in body_points[1:]:
                                p = tf.add_paragraph()
                                p.text = str(point)
                                p.level = 0
                        body_set = True
            except Exception as e:
                st.warning(f"Error creating slide {i+1}: {str(e)}")
                continue
        
        if slides_created == 0:
            st.error("No slides were successfully created.")
            return None
        
        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)
        return ppt_stream
        
    except Exception as e:
        st.error(f"Failed to create presentation: {str(e)}")
        return None

# --- Initialize Session State ---
if 'template_bytes' not in st.session_state:
    st.session_state.template_bytes = None
if 'layouts' not in st.session_state:
    st.session_state.layouts = {}
if 'template_name' not in st.session_state:
    st.session_state.template_name = None

# --- UI Layout ---

st.title("✨ Your Text, Your Style – AI Presentation Generator")
st.markdown("Transform any text into a beautifully formatted PowerPoint presentation using your own template and AI.")

# --- Sidebar for Configuration ---
with st.sidebar:
    st.header("⚙️ Configuration")
    llm_provider = st.selectbox(
        "Choose AI Provider", 
        ["OpenAI", "Gemini", "Anthropic"],
        help="Select your preferred AI service"
    )
    api_key = st.text_input(
        f"Enter your {llm_provider} API Key", 
        type="password",
        help=f"Your {llm_provider} API key (never stored)",
        key="api_key_input"
    )
    if api_key:
        st.success("✅ API Key provided")
    st.markdown("---")
    st.header("📋 How It Works")
    st.markdown("""
    1. **Upload Template**: Choose a .pptx or .potx file
    2. **Add Content**: Paste your text content
    3. **Optional Guidance**: Specify tone or style
    4. **Generate**: AI creates structured slides
    5. **Download**: Get your formatted presentation
    """)
    st.markdown("---")
    st.info("🔒 Your API key is only used for this session and never stored.")

# --- Main Content Area ---

st.header("1️⃣ Upload PowerPoint Template")
uploaded_template = st.file_uploader(
    "Choose a PowerPoint template file (.pptx or .potx)",
    type=['pptx', 'potx'],
    help="This template's layouts, fonts, and colors will be applied",
    key='template_uploader'
)

if uploaded_template is not None:
    if st.session_state.template_name != uploaded_template.name or st.session_state.template_bytes is None:
        with st.spinner(f"📋 Analyzing template: {uploaded_template.name}"):
            try:
                st.session_state.template_bytes = uploaded_template.getvalue()
                st.session_state.template_name = uploaded_template.name
                st.session_state.layouts = get_ppt_layouts(st.session_state.template_bytes)
                if st.session_state.layouts:
                    st.success(f"✅ Template analyzed! Found {len(st.session_state.layouts)} usable layouts.")
                else:
                    st.error("❌ Could not find usable layouts in the template.")
            except Exception as e:
                st.error(f"❌ Error processing template: {str(e)}")
                st.session_state.template_bytes = None
                st.session_state.layouts = {}

if st.session_state.layouts:
    with st.expander("📋 View Available Template Layouts", expanded=False):
        st.write("AI will be instructed to choose from these layouts:")
        for i, layout_info in enumerate(st.session_state.layouts.values(), 1):
            st.write(f"**{i}.** {layout_info['clean_name']}")

st.header("2️⃣ Add Your Content")
col1, col2 = st.columns([2, 1])
with col1:
    input_text = st.text_area(
        "Paste your text content here (supports markdown)",
        height=300,
        placeholder="Paste your article, report, notes, or any text content here...",
        help="Supports plain text, markdown, or any written content"
    )
    if input_text:
        char_count = len(input_text)
        word_count = len(input_text.split())
        st.caption(f"📝 {char_count:,} characters, ~{word_count:,} words")
with col2:
    st.subheader("3️⃣ Optional Guidance")
    guidance_text = st.text_input(
        "Presentation style/tone",
        placeholder="e.g., 'executive summary', 'sales pitch'",
        help="Guide the AI on tone, structure, or audience"
    )
    st.write("**Quick Options:**")
    guidance_options = ["Executive summary", "Sales pitch", "Technical presentation", "Training material"]
    for option in guidance_options:
        if st.button(option, key=f"guidance_{option}", use_container_width=True):
            st.session_state.guidance_text = option
            st.rerun()
    guidance_text = st.session_state.get('guidance_text', guidance_text)


st.header("4️⃣ Generate Presentation")
missing_items = []
if not api_key: missing_items.append("API key")
if not st.session_state.template_bytes: missing_items.append("PowerPoint template")
if not input_text.strip(): missing_items.append("text content")
if missing_items:
    st.warning(f"⚠️ Please provide: {', '.join(missing_items)}")

generate_button = st.button(
    "🚀 Generate Presentation",
    type="primary",
    use_container_width=True,
    disabled=bool(missing_items)
)

if generate_button:
    progress_bar = st.progress(0, text="Initializing...")
    try:
        progress_bar.progress(10, text="🤖 AI is analyzing your content...")
        processed_content = process_markdown_content(input_text) if any(c in input_text for c in ['*', '#']) else input_text
        
        layout_names = list(st.session_state.layouts.keys())
        slides = None
        
        if llm_provider == "Gemini":
            slides = generate_with_gemini(api_key, processed_content, guidance_text, layout_names)
        elif llm_provider == "OpenAI":
            slides = generate_with_openai(api_key, processed_content, guidance_text, layout_names)
        elif llm_provider == "Anthropic":
            slides = generate_with_anthropic(api_key, processed_content, guidance_text, layout_names)
        
        if not slides:
            st.error("❌ Failed to generate slide structure. The AI may have returned an empty or invalid response. Please check your text or try again.")
            st.stop()
        
        progress_bar.progress(50, text=f"✅ Generated {len(slides)} slides structure. Preview below.")
        with st.expander("📋 Preview Generated Slides", expanded=True):
            for i, slide in enumerate(slides, 1):
                st.write(f"**Slide {i}: {slide.get('content', {}).get('title', 'Untitled')}** (Layout: *{slide.get('layout')}*)")
                body = slide.get('content', {}).get('body', [])
                if body:
                    for point in body[:2]:
                        st.write(f"• {point[:100]}...") # Truncate long points
                    if len(body) > 2:
                        st.write(f"• ... and {len(body)-2} more points.")
                st.write("---")

        progress_bar.progress(75, text="📊 Applying template styles and creating presentation...")
        presentation_stream = create_presentation(slides, st.session_state.template_bytes, st.session_state.layouts)
        
        if not presentation_stream:
            st.error("❌ Failed to create the final presentation file.")
            st.stop()
        
        progress_bar.progress(100, text="✅ Presentation ready for download!")
        st.success(f"🎉 Successfully created presentation with {len(slides)} slides!")
        
        st.download_button(
            label="📥 Download Presentation (.pptx)",
            data=presentation_stream,
            file_name=f"AI_Generated_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary"
        )
        st.balloons()
        
    except Exception as e:
        st.error(f"❌ An unexpected error occurred: {str(e)}")
    finally:
        progress_bar.empty()

# --- Footer ---
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666; font-size: 0.8em;'>
        Made with ❤️ using Streamlit | Your API keys are never stored or logged
    </div>
    """, 
    unsafe_allow_html=True
)
