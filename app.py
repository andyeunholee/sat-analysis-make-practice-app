import streamlit as st
import google.genai as genai
from google.genai import types
import os
from pathlib import Path
import json
import re
import io
import random
import base64
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend for server
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.ticker as ticker
import numpy as np
from fpdf import FPDF

# Helper: Clean text for export (removes Hangul and Emojis)
def clean_text_for_export(text):
    if not text: return ""
    # Export fonts (standard PDF/Word) handle Latin-1 best. 
    # Encoding to latin-1 with 'ignore' strips Hangul and Emojis.
    return text.encode('latin-1', 'ignore').decode('latin-1')

# Load environment variables
load_dotenv()

# Page Configuration
st.set_page_config(
    page_title="Elite Prep | SAT Analysis & Adaptive Practice",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Brand Colors & Styles
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    
    html, body, [class*="css"]  {
        font-family: 'Inter', sans-serif;
        color: #1f2937;
        background-color: #f3f4f6;
    }
    
    /* Main Container */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 5rem;
        max-width: 1000px;
    }

    /* Header Styling */
    .main-header {
        font-family: 'Inter', sans-serif;
        color: #0C1E41; /* Elite Navy */
        text-align: center;
        font-weight: 800;
        font-size: 2.2rem;
        margin-bottom: 0.5rem;
        letter-spacing: -0.02em;
    }
    .sub-header {
        color: #6b7280;
        text-align: center;
        margin-bottom: 3rem;
        font-size: 1.1rem;
        font-weight: 400;
    }
    
    /* Card Styling */
    .stCard {
        background-color: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 2rem;
        border: 1px solid #e5e7eb;
    }

    /* Button Styling */
    div.stButton > button {
        background-color: #0C1E41;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        border: none;
        transition: all 0.2s ease;
        width: 100%;
        box-shadow: 0 4px 6px rgba(12, 30, 65, 0.2);
    }
    div.stButton > button:hover {
        background-color: #1a3a6e;
        transform: translateY(-1px);
        box-shadow: 0 6px 8px rgba(12, 30, 65, 0.3);
    }

    /* Practice Topic Buttons (Secondary) */
    .topic-btn {
        border: 1px solid #0C1E41 !important;
        background-color: white !important;
        color: #0C1E41 !important;
    }
    
    /* Status Indicators */
    .status-success { color: #166534; font-weight: 500;}
    .status-error { color: #991b1b; font-weight: 500;}

</style>
""", unsafe_allow_html=True)

# Helper: Get configured Gemini client
def _get_client():
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key and 'api_key' in st.session_state and st.session_state.api_key:
        api_key = st.session_state.api_key
    if api_key:
        return genai.Client(api_key=api_key)
    return None

# Helper: Upload file to Gemini (File API)
def upload_to_gemini(path, mime_type=None):
    try:
        client = _get_client()
        if not client:
            return None
        config = types.UploadFileConfig(mime_type=mime_type) if mime_type else None
        file = client.files.upload(path=path, config=config)
        return file
    except Exception as e:
        print(f"Error uploading {path}: {e}")
        return None

# Helper: Load Gemini Model
def get_gemini_response(input_prompt, content_parts, temperature=0.2):
    try:
        client = _get_client()
        if not client:
            st.error("Google API Key not found.")
            return None

        full_payload = [input_prompt] + content_parts

        response = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=full_payload,
            config=types.GenerateContentConfig(
                temperature=temperature,
                max_output_tokens=8192,
            ),
        )

        # Safe access to text
        try:
            return response.text
        except ValueError:
            # Handle cases where response is blocked or empty
            finish_reason = response.candidates[0].finish_reason if response.candidates else "UNKNOWN"
            st.error(f"Generation stopped. Finish Reason: {finish_reason}")
            # If MAX_TOKENS (2), try to return what we have
            if finish_reason == 2 and response.candidates and response.candidates[0].content.parts:
                return response.candidates[0].content.parts[0].text
            return None
    except Exception as e:
        st.error(f"Error calling Gemini API: {str(e)}")
        return None

# Helper: Extract JSON from text
def extract_json_topics(text):
    try:
        # Find JSON block using regex
        match = re.search(r'```json\s*(\{.*?\})\s*```', text, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        else:
            # Try finding just brace content if no markdown tags
            match = re.search(r'\{.*\}', text, re.DOTALL)
            if match:
                return json.loads(match.group(0))
    except Exception as e:
        print(f"JSON Parse Error: {e}")
    return None

# Helper: Force Formatting for Options
def ensure_formatting(text):
    if not text: return ""
    # Remove <br> tags
    text = re.sub(r'<br\s*/?>', '', text)
    # Regex to find A) B) C) D) that are NOT at start of line
    pattern = r"(\s)([A-D]\)) "
    formatted_text = re.sub(pattern, r"\n\n\2 ", text)
    return formatted_text

# Helper: Execute matplotlib code and return PNG image bytes
def execute_figure_code(code_str):
    """Execute matplotlib code string and return PNG image as bytes."""
    try:
        plt.close('all')  # Close any previous figures
        
        # Pre-process the code to fix common issues
        lines = code_str.split('\n')
        cleaned_lines = []
        for line in lines:
            stripped = line.strip()
            # Skip import statements (plt, np, io, math already available)
            if stripped.startswith('import ') or stripped.startswith('from '):
                continue
            # Replace plt.show() with savefig
            if stripped == 'plt.show()':
                continue
            cleaned_lines.append(line)
        
        cleaned_code = '\n'.join(cleaned_lines)
        
        # If no savefig call exists, append one
        if 'savefig' not in cleaned_code:
            cleaned_code += "\nplt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')"
        
        # Create a BytesIO buffer for the image
        buf = io.BytesIO()
        
        # Set up execution namespace with common imports available
        exec_namespace = {
            'plt': plt,
            'np': np,
            'io': io,
            'buf': buf,
            'math': __import__('math'),
            'matplotlib': matplotlib,
            'patches': patches,
            'ticker': ticker,
        }
        
        # Execute the code
        exec(cleaned_code, exec_namespace)
        
        # If buf is still empty, try to save current figure
        if buf.tell() == 0:
            fig = plt.gcf()
            if fig.get_axes():
                fig.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                           facecolor='white', edgecolor='none')
        
        buf.seek(0)
        img_bytes = buf.getvalue()
        plt.close('all')
        
        if len(img_bytes) > 0:
            return img_bytes
        return None
    except Exception as e:
        plt.close('all')
        # Silence warning as requested by user
        # st.warning(f"⚠️ Figure generation error: {e}")
        return None

# Helper: Parse Gemini response into text and figure segments
def parse_response_with_figures(response_text):
    """
    Parse response containing ```python-figure code blocks.
    Returns list of segments: [{"type":"text","content":"..."}, {"type":"figure","image":bytes,"code":"..."}]
    """
    if not response_text:
        return [{"type": "text", "content": ""}]
    
    # Pattern to match ```python-figure ... ``` OR ```python ... ``` blocks containing plt
    pattern = r'```(?:python-figure|python)\s*\n(.*?)```'
    
    segments = []
    last_end = 0
    
    for match in re.finditer(pattern, response_text, re.DOTALL):
        # Add text before this code block
        text_before = response_text[last_end:match.start()].strip()
        if text_before:
            segments.append({"type": "text", "content": text_before})
        
        # Execute the figure code
        code = match.group(1).strip()
        img_bytes = execute_figure_code(code)
        
        if img_bytes:
            segments.append({"type": "figure", "image": img_bytes, "code": code})
        else:
            # If figure generation failed, add a note
            segments.append({"type": "text", "content": "*[Figure could not be generated]*"})
        
        last_end = match.end()
    
    # Add remaining text after last code block
    remaining = response_text[last_end:].strip()
    if remaining:
        segments.append({"type": "text", "content": remaining})
    
    # If no figures found, return entire text as single segment
    if not segments:
        segments.append({"type": "text", "content": response_text})
    
    return segments

# Helper: Convert LaTeX math to readable plain text
def _latex_to_readable(text):
    """Convert LaTeX math expressions to readable Unicode text."""
    if not text:
        return text
    
    def process_math(match):
        math = match.group(1)
        # Handle nested fractions (inner to outer)
        max_iter = 10
        while r'\frac' in math and max_iter > 0:
            math = re.sub(r'\\frac\{([^{}]*)\}\{([^{}]*)\}', r'(\1)/(\2)', math)
            max_iter -= 1
        # Common LaTeX symbols
        replacements = {
            r'\times': '×', r'\div': '÷', r'\pm': '±',
            r'\leq': '≤', r'\geq': '≥', r'\neq': '≠',
            r'\rightarrow': '→', r'\leftarrow': '←',
            r'\cdot': '·', r'\approx': '≈', r'\infty': '∞',
            r'\pi': 'π', r'\theta': 'θ', r'\alpha': 'α', r'\beta': 'β',
            r'\sqrt': '√', r'\le': '≤', r'\ge': '≥',
        }
        for latex_cmd, symbol in replacements.items():
            math = math.replace(latex_cmd, symbol)
        # Subscripts: Try common ones first
        subscripts = {'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉'}
        for k, v in subscripts.items():
            math = math.replace('_{' + k + '}', v).replace('_' + k, v)
        
        # Superscripts: Handle common ones (Latin-1 compatible for PDF)
        superscripts = {'1': '¹', '2': '²', '3': '³'}
        for k, v in superscripts.items():
            math = math.replace('^{' + k + '}', v).replace('^' + k, v)
            
        # Fallback for other superscripts/subscripts
        math = re.sub(r'_\{([^}]*)\}', r'_\1', math)
        math = re.sub(r'\^\{([^}]*)\}', r'^(\1)', math)
        
        # Remove remaining backslashes from unknown commands
        math = re.sub(r'\\([a-zA-Z]+)\s*', r'\1 ', math)
        # Clean extra spaces
        math = re.sub(r'\s+', ' ', math).strip()
        return math
    
    # Process $$...$$ blocks first, then $...$
    text = re.sub(r'\$\$\s*([^$]+?)\s*\$\$', process_math, text)
    text = re.sub(r'\$([^$]+?)\$', process_math, text)
    
    # Also convert bare LaTeX commands outside $ delimiters
    while r'\frac' in text:
        prev = text
        text = re.sub(r'\\frac\{([^{}]*)\}\{([^{}]*)\}', r'(\1)/(\2)', text)
        if text == prev:
            break
    
    bare_replacements = {
        r'\times': '×', r'\cdot': '·', r'\rightarrow': '→',
        r'\approx': '≈', r'\leq': '≤', r'\geq': '≥', r'\neq': '≠',
        r'\pm': '±', r'\div': '÷',
    }
    for latex_cmd, symbol in bare_replacements.items():
        text = text.replace(latex_cmd, symbol)
    
    # Convert bare superscripts for export
    text = text.replace('^{2}', '²').replace('^2', '²')
    text = text.replace('^{3}', '³').replace('^3', '³')
    text = text.replace('^{1}', '¹').replace('^1', '¹')
    
    return text

# Helper: Convert Markdown to Word (.docx) document
def convert_markdown_to_docx(md_text):
    if not md_text: return None
    
    doc = Document()
    
    # Set default font for Normal style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x1f, 0x29, 0x37)
    
    # Configure heading styles
    for level in range(1, 4):
        heading_style = doc.styles[f'Heading {level}']
        heading_style.font.name = 'Arial'
        heading_style.font.color.rgb = RGBColor(0x0C, 0x1E, 0x41)
    
    # Set margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    lines = md_text.split('\n')
    i = 0
    in_code_block = False
    code_block_lines = []
    code_block_lang = ""
    
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        # Handle code block start/end
        if stripped.startswith('```'):
            if not in_code_block:
                # Starting a code block
                code_block_lang = stripped[3:].strip().lower()
                in_code_block = True
                code_block_lines = []
                i += 1
                continue
            else:
                # Ending a code block
                in_code_block = False
                if code_block_lang in ('python-figure', 'python'):
                    # Execute matplotlib code and embed image
                    code_str = '\n'.join(code_block_lines)
                    img_bytes = execute_figure_code(code_str)
                    if img_bytes:
                        img_buf = io.BytesIO(img_bytes)
                        doc.add_picture(img_buf, width=Inches(4.5))
                        # Center the image
                        last_para = doc.paragraphs[-1]
                        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # Regular code block - add as formatted text
                    code_text = '\n'.join(code_block_lines)
                    if code_text.strip():
                        para = doc.add_paragraph()
                        run = para.add_run(code_text)
                        run.font.name = 'Consolas'
                        run.font.size = Pt(9)
                code_block_lang = ""
                code_block_lines = []
                i += 1
                continue
        
        if in_code_block:
            code_block_lines.append(line.rstrip())
            i += 1
            continue
        
        # Strip blockquote markers: "> text" → "text", ">" → ""
        while stripped.startswith('>'):
            stripped = stripped[1:].strip()
        
        # Empty line → skip (no blank rows)
        if not stripped:
            i += 1
            continue
        
        # Convert LaTeX in the line
        stripped = _latex_to_readable(stripped)
        
        # Horizontal rule / Page Break
        if stripped in ('---', '--- PAGE BREAK ---') or stripped.startswith('─'):
            para = doc.add_paragraph()
            run = para.add_run('━' * 60)
            run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
            run.font.size = Pt(8)
            i += 1
            continue
        
        # Headers (### > ## > #)
        if stripped.startswith('### '):
            heading = doc.add_heading(level=3)
            _add_formatted_text(heading, stripped[4:])
            i += 1
            continue
        elif stripped.startswith('## '):
            heading = doc.add_heading(level=2)
            _add_formatted_text(heading, stripped[3:])
            i += 1
            continue
        elif stripped.startswith('# '):
            heading = doc.add_heading(level=1)
            _add_formatted_text(heading, stripped[2:])
            i += 1
            continue
        
        # Bullet points: * or - at start
        if re.match(r'^[\*\-]\s+', stripped):
            content = re.sub(r'^[\*\-]\s+', '', stripped)
            para = doc.add_paragraph(style='List Bullet')
            _add_formatted_text(para, content)
            i += 1
            continue
        
        # Numbered items: "1." or "1)" at start — render as text to preserve original numbering
        num_match = re.match(r'^(\d+)([.)\t])\s*(.*)', stripped)
        if num_match:
            num = num_match.group(1)
            content = num_match.group(3)
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.2)
            run = para.add_run(f"{num}. ")
            run.bold = True
            _add_formatted_text(para, content)
            i += 1
            continue
        
        # Answer option lines: A) B) C) D) — with slight indent
        option_match = re.match(r'^([A-D])\)\s+(.*)', stripped)
        if option_match:
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.3)
            para.paragraph_format.space_before = Pt(2)
            para.paragraph_format.space_after = Pt(2)
            run = para.add_run(f"{option_match.group(1)})  ")
            run.bold = True
            _add_formatted_text(para, option_match.group(2))
            i += 1
            continue
        
        # Regular paragraph
        para = doc.add_paragraph()
        _add_formatted_text(para, stripped)
        i += 1
    
    # Save to bytes buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def _add_formatted_text(paragraph, text):
    """Add text to a paragraph with bold and italic formatting support."""
    if not text:
        return
    
    # Remove Hangul/Emojis for Word consistency
    text = clean_text_for_export(text)
    
    # First convert any remaining LaTeX
    text = _latex_to_readable(text)
    
    # Split by bold markers **...**
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            content = part[2:-2]
            # Handle superscripts inside bold
            _add_superscripts(paragraph, content, True)
        else:
            # Handle italic *...*
            italic_parts = re.split(r'(\*[^*]+\*)', part)
            for ip in italic_parts:
                if not ip:
                    continue
                if ip.startswith('*') and ip.endswith('*') and len(ip) > 2:
                    content = ip[1:-1]
                    _add_superscripts(paragraph, content, False, True)
                else:
                    _add_superscripts(paragraph, ip, False, False)

def _add_superscripts(paragraph, text, bold=False, italic=False):
    """Helper to add text with real superscript formatting for Word."""
    # Replace unicode superscripts with ^ marker for parsing
    text = text.replace('²', '^2').replace('³', '^3').replace('¹', '^1')
    text = text.replace('₀', '_0').replace('₁', '_1').replace('₂', '_2').replace('₃', '_3')
    
    # Split by ^ and _
    parts = re.split(r'(\^[0-9a-zA-Z]+|_[0-9a-zA-Z]+)', text)
    for part in parts:
        if part.startswith('^'):
            run = paragraph.add_run(part[1:])
            run.font.superscript = True
            run.bold = bold
            run.italic = italic
        elif part.startswith('_'):
            run = paragraph.add_run(part[1:])
            run.font.subscript = True
            run.bold = bold
            run.italic = italic
        else:
            run = paragraph.add_run(part)
            run.bold = bold
            run.italic = italic

class ElitePDF(FPDF):
    def header(self):
        # Modern Header with Navy bar
        self.set_fill_color(12, 30, 65) # Elite Navy
        self.rect(0, 0, 210, 15, 'F')
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(255, 255, 255)
        self.set_y(5)
        self.cell(0, 5, "ELITE PREP | SAT ADAPTIVE PRACTICE STATION", align="C", ln=True)
        self.ln(10)

    def footer(self):
        # Subtle Footer
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f"Confidential Elite Prep - Page {self.page_no()}/{{nb}}", align="C")

def convert_markdown_to_pdf(md_text):
    if not md_text: return None
    
    pdf = ElitePDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    
    elite_navy = (12, 30, 65)
    
    # Block grouping to keep questions together
    raw_lines = md_text.split('\n')
    blocks = []
    temp_block = []
    
    for line in raw_lines:
        s = line.strip()
        # Detect new sections or questions to start a new block
        # Match question numbers like 1. or **1.**
        is_q = re.match(r'^(\*\*|)(\d+)[.)]', s)
        if s.startswith('### ') or s.startswith('## ') or s.startswith('# ') or \
           is_q or s.startswith('Question ') or \
           s.startswith('Answer Key') or s == '--- PAGE BREAK ---':
            if temp_block: blocks.append(temp_block)
            temp_block = [line]
        else:
            temp_block.append(line)
    if temp_block: blocks.append(temp_block)
    
    for block in blocks:
        # Estimate height to prevent messy breaks
        est_h = len(block) * 8
        if any('```python-figure' in l for l in block): est_h += 80
        if pdf.get_y() + est_h > (pdf.h - 25): pdf.add_page()
        
        in_code = False
        code_lines = []
        code_lang = ""
        
        for line in block:
            pdf.set_x(18) # Relaxed margins
            s = line.strip()
            
            if s.startswith('```'):
                if not in_code:
                    in_code = True; code_lang = s[3:].strip().lower(); code_lines = []; continue
                else:
                    in_code = False
                    if code_lang in ('python-figure', 'python'):
                        img = execute_figure_code('\n'.join(code_lines))
                        if img: pdf.image(io.BytesIO(img), x=35, w=140); pdf.ln(5)
                    else:
                        pdf.set_font("Courier", size=9); pdf.set_fill_color(245, 245, 245)
                        pdf.multi_cell(0, 5, clean_text_for_export('\n'.join(code_lines)), fill=True)
                        pdf.set_font("Helvetica", size=11)
                    continue
            if in_code: code_lines.append(line); continue
                
            # Process text and clean markdown artifacts
            text = _latex_to_readable(s)
            while text.startswith('>'): text = text[1:].strip()
            
            # Regex for SAT elements
            q_m = re.match(r'^(\*\*|)(\d+)[.)\s\*]+(.*)', text)
            o_m = re.match(r'^(\*\*|)([A-D])\)[.\s\*]*(.*)', text)
            
            if text.startswith('### '):
                pdf.set_text_color(*elite_navy); pdf.set_font("Helvetica", "B", 14)
                pdf.cell(0, 12, clean_text_for_export(text[4:]), ln=True)
                pdf.set_font("Helvetica", size=11); pdf.set_text_color(0)
            elif text.startswith('## ') or text.startswith('# '):
                pdf.set_text_color(*elite_navy); pdf.set_font("Helvetica", "B", 18)
                label = text[text.find(' ')+1:]
                pdf.cell(0, 16, clean_text_for_export(label), ln=True)
                pdf.set_font("Helvetica", size=11); pdf.set_text_color(0)
            elif text in ('---', '--- PAGE BREAK ---'):
                if text == '--- PAGE BREAK ---': pdf.add_page()
                else:
                    pdf.ln(2); pdf.set_draw_color(200, 200, 200)
                    pdf.line(20, pdf.get_y(), 190, pdf.get_y()); pdf.ln(5)
            elif q_m:
                pdf.set_font("Helvetica", "B", 11); pdf.set_text_color(*elite_navy)
                num = q_m.group(2); content = q_m.group(3).strip().replace('**', '')
                pdf.write(7, f"{num}. "); pdf.set_font("Helvetica", size=11); pdf.set_text_color(0)
                pdf.multi_cell(0, 7, clean_text_for_export(content))
            elif o_m:
                pdf.set_x(28); pdf.set_font("Helvetica", "B", 10)
                letter = o_m.group(2); o_text = o_m.group(3).strip().replace('**', '')
                pdf.write(6, f"{letter}) "); pdf.set_font("Helvetica", size=11)
                pdf.multi_cell(0, 6, clean_text_for_export(o_text))
            else:
                cleaned = clean_text_for_export(text).replace('**', '').replace('__', '')
                if cleaned.strip(): pdf.multi_cell(0, 6, cleaned)
                else: pdf.ln(3)

                
    return bytes(pdf.output())

# --- Sidebar ---
with st.sidebar:
    # Logo & Brand (Centered)
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        if Path("Elite-Logo.png").exists():
            st.image("Elite-Logo.png", width='stretch')
    
    st.markdown("<h2 style='text-align: center; color: #0C1E41; margin-top: -10px;'>Elite Prep</h2>", unsafe_allow_html=True)
    
    st.markdown("---")
    
    # API Key (Hidden Label)
    env_key = os.getenv("GOOGLE_API_KEY")
    if env_key:
        st.success("✅ API Key Loaded")
    else:
        api_key_input = st.text_input("Enter Google API Key", type="password")
        if api_key_input:
            st.session_state.api_key = api_key_input
    
    st.markdown("---")
    st.markdown("### 📂 Upload the Student’s SAT Test Result Report")
    uploaded_files = st.file_uploader(
        "Upload Diagnostic PDF(s)", 
        type=['pdf', 'jpg', 'png'], 
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    
st.markdown("<h1 class='main-header'>Elite Prep | SAT Analysis Report & Improvement Plan</h1>", unsafe_allow_html=True)


# Initialization
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'full_report' not in st.session_state:
    st.session_state.full_report = ""
if 'weak_topics' not in st.session_state:
    st.session_state.weak_topics = None
if 'practice_result' not in st.session_state:
    st.session_state.practice_result = ""
if 'all_topics' not in st.session_state:
    st.session_state.all_topics = None
if 'manual_practice_result' not in st.session_state:
    st.session_state.manual_practice_result = ""

# ==========================================
# Student Analysis & Adaptive Practice
# ==========================================
with st.container():
    # 1. Report Generation Phase
    if not st.session_state.analysis_done:
        if uploaded_files:
            st.success(f"📂 {len(uploaded_files)} Report(s) Ready.")
            
            if st.button("🚀 Generate Analysis Report"):
                with st.spinner("Analyzing Student Data... (This takes about 10-20 seconds)"):
                    
                    
                    # Load User Files
                    student_parts = []
                    for up_file in uploaded_files:
                        student_parts.append(
                            types.Part.from_bytes(
                                data=up_file.getvalue(),
                                mime_type=up_file.type
                            )
                        )
                    
                    # --- NEW: Load Knowledge Base ---
                    knowledge_base_text = ""
                    kb_path = Path("knowledge/sat_curriculum.json")
                    if kb_path.exists():
                        try:
                            with open(kb_path, "r", encoding="utf-8") as f:
                                knowledge_base_text = f.read()
                        except Exception as e:
                            st.warning(f"Error loading knowledge base: {e}")
                    
                    all_content = student_parts
                    
                    # Report Prompt
                    prompt = f"""
You are Elite Prep's Senior SAT Consultant. You have been provided with:
1. The student's Diagnostic Test Result images (the actual score sheets to analyze).
2. The official SAT Curriculum Knowledge Base (below).

**KNOWLEDGE BASE:**
```json
{knowledge_base_text}
```

**TASK:** Produce a formal diagnostic report. Output 100% in English. No Korean characters anywhere.

---

**OUTPUT FORMAT — follow this structure exactly, word-for-word section titles:**

Based on the diagnostic test results provided for [Student Name] across [N] practice tests ([list test names]), here is a detailed analysis of the student's performance, key weaknesses, and a targeted study plan.

## Diagnostic Report Overview
Student Name: [Full Name]
Test Trajectory: [Score (Test 1)] → [Score (Test 2)] → ... → [Score (Test N)]
Performance Trend: [2–3 sentence narrative. Mention total score improvement, omitted questions trend, Math plateau, and RW consistency. Be specific with numbers.]

---

## Key Weaknesses & Curricular Solutions

[For each weakness, use the structure below. Number them 1, 2, 3, 4. Always cover exactly 2 Math and 2 Reading & Writing weaknesses. DO NOT compare with other students, focus solely on the requested student. Use the KNOWLEDGE BASE to map the weakness to the correct Curricular Category and Chapters.]

### [N]. [Subject]: [Weakness Category Name] ([Severity: e.g., "Critical Weakness" / "Foundational Gap" / "Inconsistency"])
**Specific Skills Lacking:** [2–3 sentence explanation. Reference specific test scores and percentages from the student's actual results. Describe exactly what conceptual knowledge is missing based on the Knowledge Base 'common_mistakes' and 'key_concepts'.]
**Curriculum Chapters:**
- Chapter [X]: [Chapter Name]
- Chapter [X]: [Chapter Name]
- (add more as needed)

---

## Tutoring Recommendations & Logistics
**Target Score:** [Score - use the Scoring Guidelines from the Knowledge Base based on their starting score]
**Goal Breakdown:** [1–2 sentences: Math target, RW target, and why.]
**Total Tutoring Hours:** [N] Hours [use the Scoring Guidelines]
**Split:** [Math Hours] Hours Math / [RW Hours] Hours Reading & Writing. ([1 sentence justification.])
**Frequency:** [Frequency from Scoring Guidelines]

### Week-by-Week Study Plan
**Week 1 ([Theme]):**
- Math: [Chapter(s)]
- RW: [Chapter(s) + brief focus note]

**Week 2 ([Theme]):**
- Math: [Chapter(s)]
- RW: [Chapter(s) + brief focus note]

[Continue for the recommended number of weeks (typically 6-8). Second to last week = timed drills. Last week = full practice test review.]

---

**CRITICAL OUTPUT INSTRUCTION:**
After the full report, append a RAW JSON block listing the weak lesson subjects for the practice question generator. Format exactly:
```json
{{
    "English": ["Chapter X: Name", "Chapter X: Name"],
    "Math": ["Chapter X: Name", "Chapter X: Name"]
}}
```
"""
                    
                    response_text = get_gemini_response(prompt, all_content)
                    
                    if response_text:
                        st.session_state.full_report = response_text
                        st.session_state.analysis_done = True
                        
                        # Parse JSON
                        topics = extract_json_topics(response_text)
                        if topics:
                            st.session_state.weak_topics = topics
                        else:
                            st.warning("Could not auto-detect weak topics for the practice menu. (JSON Parsing Failed)")
                            # Fallback default
                            st.session_state.weak_topics = {"English": ["General Practice"], "Math": ["General Practice"]}
                        
                        st.rerun() # Reload to show report
                    else:
                        st.error("Analysis Failed. Please try again.")
        else:
            st.info("👋 Please upload the student's SAT test result in the sidebar to start the analysis.")

    # 2. Display Report
    if st.session_state.analysis_done:
        st.markdown("### 📊 Analysis Report")
        
        # Remove the JSON block from display if possible for cleanliness (Optional, keeping it simple for now)
        clean_report = re.sub(r'```json\s*\{.*?\}\s*```', '', st.session_state.full_report, flags=re.DOTALL)
        
        with st.expander("📄 View Full Analysis Report", expanded=True):
            st.markdown(clean_report)
