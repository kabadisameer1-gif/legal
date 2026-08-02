# app.py - Legal Guide AI
# Run with: streamlit run app.py

import os
import re
import io
import streamlit as st
from dotenv import load_dotenv
import google.generativeai as genai
import pdfplumber
from docx import Document as DocxDocument

# ---------- PAGE CONFIGURATION & DESIGN SYSTEM ----------
st.set_page_config(
    page_title="Legal Guide AI • Concierge",
    page_icon="👑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------- THEME SYSTEM & DESIGN PALETTES ----------
def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    if len(hex_str) != 6:
        return (255, 215, 0)
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

def get_theme_css(theme_name: str, custom_hex: str = "#00E5FF") -> str:
    themes = {
        "🏛️ Classic Legal (Default)": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #334155 0%, #1E293B 35%, #0F172A 70%, #070A12 100%)",
            "primary": "#CBD5E1",
            "secondary": "#94A3B8",
            "dark": "#334155",
            "light": "#F8FAFC",
            "glow": "rgba(203, 213, 225, 0.4)",
            "border": "rgba(148, 163, 184, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(203, 213, 225, 0.2) 0%, rgba(148, 163, 184, 0.1) 100%)"
        },
        "📜 Classic Parchment & Gold": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #2D261E 0%, #1C1712 35%, #120E0B 70%, #090705 100%)",
            "primary": "#E6C687",
            "secondary": "#C5A059",
            "dark": "#5C4033",
            "light": "#FDFBF7",
            "glow": "rgba(230, 198, 135, 0.45)",
            "border": "rgba(197, 160, 89, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(230, 198, 135, 0.2) 0%, rgba(197, 160, 89, 0.1) 100%)"
        },
        "👑 Royal Gold": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #1E1B4B 0%, #0F172A 35%, #070A14 70%, #020308 100%)",
            "primary": "#FFD700",
            "secondary": "#D4AF37",
            "dark": "#78350F",
            "light": "#FFE5A3",
            "glow": "rgba(255, 215, 0, 0.45)",
            "border": "rgba(255, 215, 0, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(255, 215, 0, 0.2) 0%, rgba(184, 134, 11, 0.1) 100%)"
        },
        "💎 Imperial Emerald": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #064E3B 0%, #022C22 35%, #061A14 70%, #020D0A 100%)",
            "primary": "#10B981",
            "secondary": "#059669",
            "dark": "#064E3B",
            "light": "#A7F3D0",
            "glow": "rgba(16, 185, 129, 0.45)",
            "border": "rgba(16, 185, 129, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(16, 185, 129, 0.2) 0%, rgba(5, 150, 105, 0.1) 100%)"
        },
        "🔮 Cyber Sapphire": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #1E3A8A 0%, #0F172A 35%, #070F26 70%, #020617 100%)",
            "primary": "#06B6D4",
            "secondary": "#0284C7",
            "dark": "#1E3A8A",
            "light": "#BAE6FD",
            "glow": "rgba(6, 182, 212, 0.45)",
            "border": "rgba(6, 182, 212, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(6, 182, 212, 0.2) 0%, rgba(2, 132, 199, 0.1) 100%)"
        },
        "💜 Amethyst Velvet": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #4C1D95 0%, #240046 35%, #10002B 70%, #050014 100%)",
            "primary": "#C084FC",
            "secondary": "#9333EA",
            "dark": "#581C87",
            "light": "#F5D0FE",
            "glow": "rgba(192, 132, 252, 0.45)",
            "border": "rgba(192, 132, 252, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(192, 132, 252, 0.2) 0%, rgba(147, 51, 234, 0.1) 100%)"
        },
        "🩸 Ruby Crimson": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #881337 0%, #4C0519 35%, #1F020A 70%, #0B0104 100%)",
            "primary": "#F43F5E",
            "secondary": "#E11D48",
            "dark": "#881337",
            "light": "#FECDD3",
            "glow": "rgba(244, 63, 94, 0.45)",
            "border": "rgba(244, 63, 94, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(244, 63, 94, 0.2) 0%, rgba(225, 29, 72, 0.1) 100%)"
        },
        "⚡ Cyberpunk Amber": {
            "bg_radial": "radial-gradient(ellipse at 50% -10%, #451A03 0%, #1C1917 35%, #0C0A09 70%, #030202 100%)",
            "primary": "#F59E0B",
            "secondary": "#D97706",
            "dark": "#78350F",
            "light": "#FDE68A",
            "glow": "rgba(245, 158, 11, 0.45)",
            "border": "rgba(245, 158, 11, 0.35)",
            "badge_bg": "linear-gradient(135deg, rgba(245, 158, 11, 0.2) 0%, rgba(217, 119, 6, 0.1) 100%)"
        }
    }

    if theme_name in themes:
        cfg = themes[theme_name]
    else:
        r, g, b = hex_to_rgb(custom_hex)
        lr, lg, lb = min(255, r + 70), min(255, g + 70), min(255, b + 70)
        dr, dg, db = max(0, r - 60), max(0, g - 60), max(0, b - 60)
        
        cfg = {
            "bg_radial": f"radial-gradient(ellipse at 50% -10%, rgb({dr},{dg},{db}) 0%, #0F172A 35%, #070A14 70%, #020308 100%)",
            "primary": custom_hex,
            "secondary": f"rgb({dr},{dg},{db})",
            "dark": f"rgb({max(0, dr-20)},{max(0, dg-20)},{max(0, db-20)})",
            "light": f"rgb({lr},{lg},{lb})",
            "glow": f"rgba({r}, {g}, {b}, 0.45)",
            "border": f"rgba({r}, {g}, {b}, 0.35)",
            "badge_bg": f"linear-gradient(135deg, rgba({r}, {g}, {b}, 0.2) 0%, rgba({dr}, {dg}, {db}, 0.1) 100%)"
        }

    return f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel+Decorative:wght@700;900&family=Cinzel:wght@400;600;700;800;900&family=Cormorant+Garamond:ital,wght@0,500;0,700;1,400;1,600&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&family=Outfit:wght@300;400;500;600;700&display=swap');

    :root {{
        --bg-radial: {cfg['bg_radial']};
        --accent-primary: {cfg['primary']};
        --accent-secondary: {cfg['secondary']};
        --accent-dark: {cfg['dark']};
        --accent-light: {cfg['light']};
        --accent-glow: {cfg['glow']};
        --accent-border: {cfg['border']};
        --badge-bg: {cfg['badge_bg']};
    }}

    /* Global Background & Classic Body Text */
    .stApp, .stMarkdown, p, span, label, li {{
        background: var(--bg-radial) !important;
        font-family: 'Cormorant Garamond', 'Georgia', 'Times New Roman', serif !important;
        font-size: 1.1rem;
        color: #F8FAFC !important;
    }}
    
    /* Main Container Padding */
    .main .block-container {{
        padding-top: 1.8rem;
        padding-bottom: 5rem;
        max-width: 1260px;
    }}
    
    /* Scrollbar */
    ::-webkit-scrollbar {{
        width: 10px;
        height: 10px;
    }}
    ::-webkit-scrollbar-track {{
        background: #020308;
    }}
    ::-webkit-scrollbar-thumb {{
        background: linear-gradient(180deg, var(--accent-primary) 0%, var(--accent-secondary) 50%, var(--accent-dark) 100%);
        border-radius: 5px;
        border: 2px solid #020308;
    }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background: rgba(4, 6, 14, 0.94) !important;
        backdrop-filter: blur(30px);
        -webkit-backdrop-filter: blur(30px);
        border-right: 1px solid var(--accent-border) !important;
        box-shadow: 10px 0 45px rgba(0, 0, 0, 0.8);
    }}
    
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {{
        font-family: 'Cinzel Decorative', 'Cinzel', serif !important;
        color: var(--accent-light) !important;
        letter-spacing: 2px !important;
        text-transform: uppercase !important;
        font-weight: 800 !important;
    }}

    /* Radio Buttons in Sidebar */
    section[data-testid="stSidebar"] .stRadio > div {{
        gap: 10px;
    }}
    
    section[data-testid="stSidebar"] .stRadio label {{
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8) 0%, rgba(30, 27, 75, 0.4) 100%) !important;
        border: 1px solid var(--accent-border) !important;
        border-radius: 12px !important;
        padding: 0.75rem 1.1rem !important;
        margin-bottom: 0.25rem !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer;
        font-weight: 500 !important;
    }}
    
    section[data-testid="stSidebar"] .stRadio label:hover {{
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.9) 0%, rgba(49, 46, 129, 0.6) 100%) !important;
        border-color: var(--accent-primary) !important;
        transform: translateX(6px);
        box-shadow: 0 4px 20px var(--accent-glow);
    }}

    /* Header Banner */
    .app-header {{
        background: linear-gradient(135deg, rgba(30, 27, 75, 0.85) 0%, rgba(15, 23, 42, 0.95) 50%, rgba(7, 10, 20, 0.95) 100%);
        border: 1px solid var(--accent-border);
        border-radius: 24px;
        padding: 3.2rem 2rem 2.8rem 2rem;
        text-align: center;
        box-shadow: 0 20px 60px rgba(0, 0, 0, 0.8), inset 0 1px 0 rgba(255, 230, 160, 0.35);
        margin-bottom: 2.2rem;
        position: relative;
        overflow: hidden;
    }}
    
    .app-header::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 4px;
        background: linear-gradient(90deg, transparent, var(--accent-secondary), var(--accent-primary), #FFF8DC, var(--accent-primary), var(--accent-secondary), transparent);
    }}
    
    .app-header::after {{
        content: '';
        position: absolute;
        bottom: 0; left: 15%; right: 15%; height: 1px;
        background: linear-gradient(90deg, transparent, var(--accent-border), transparent);
    }}

    .app-insignia {{
        font-size: 3.2rem;
        line-height: 1;
        margin-bottom: 0.8rem;
        filter: drop-shadow(0 0 20px var(--accent-glow));
    }}

    .app-badge {{
        display: inline-block;
        background: var(--badge-bg);
        border: 1px solid var(--accent-border);
        color: var(--accent-light);
        font-family: 'Cinzel', 'Georgia', serif;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 3px;
        text-transform: uppercase;
        padding: 0.4rem 1.4rem;
        border-radius: 50px;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 15px var(--accent-glow);
    }}

    .app-title {{
        font-family: 'Cinzel Decorative', 'Cinzel', serif;
        font-size: 3.2rem;
        font-weight: 900;
        background: linear-gradient(135deg, #FFFFFF 0%, #FFF8DC 20%, var(--accent-primary) 50%, var(--accent-secondary) 75%, var(--accent-dark) 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.8rem;
        letter-spacing: 3px;
        line-height: 1.2;
        filter: drop-shadow(0 4px 15px rgba(0, 0, 0, 0.9));
    }}
    
    .app-subtitle {{
        font-family: 'Cormorant Garamond', serif;
        font-size: 1.35rem;
        font-style: italic;
        color: #E2E8F0;
        font-weight: 500;
        letter-spacing: 1px;
        max-width: 850px;
        margin: 0 auto;
    }}

    /* Key Metrics Live Dashboard Bar */
    .metrics-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
        gap: 1.3rem;
        margin-bottom: 2.8rem;
    }}

    .app-metric-card {{
        background: linear-gradient(145deg, rgba(20, 26, 48, 0.75) 0%, rgba(10, 14, 28, 0.85) 100%);
        backdrop-filter: blur(20px);
        border: 1px solid var(--accent-border);
        border-radius: 16px;
        padding: 1.3rem 1.5rem;
        display: flex;
        align-items: center;
        gap: 1.1rem;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.08);
        transition: all 0.35s ease;
    }}

    .app-metric-card:hover {{
        border-color: var(--accent-primary);
        transform: translateY(-4px);
        box-shadow: 0 15px 35px var(--accent-glow);
    }}

    .app-metric-icon {{
        font-size: 2.2rem;
        line-height: 1;
        background: var(--badge-bg);
        border-radius: 12px;
        padding: 0.7rem;
        border: 1px solid var(--accent-border);
    }}

    .app-metric-val {{
        font-family: 'Cinzel', serif;
        font-size: 1.3rem;
        font-weight: 800;
        color: var(--accent-light);
        line-height: 1.2;
    }}

    .app-metric-lbl {{
        font-size: 0.82rem;
        color: #94A3B8;
        letter-spacing: 0.5px;
    }}

    /* Section & Card Headings (Cinzel Classic Roman Inscription Style) */
    h1, h2, h3, h4, h5, h6,
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown h5, .stMarkdown h6,
    .app-card h1, .app-card h2, .app-card h3, .app-card h4, .chip-title {{
        font-family: 'Cinzel Decorative', 'Cinzel', serif !important;
        color: var(--accent-light) !important;
        letter-spacing: 1.8px !important;
        font-weight: 700 !important;
    }}

    .stMarkdown h3 {{
        font-size: 1.55rem !important;
        letter-spacing: 1.6px !important;
    }}
    
    /* Custom Glass Cards */
    .app-card {{
        background: linear-gradient(145deg, rgba(20, 26, 48, 0.8) 0%, rgba(10, 14, 28, 0.9) 100%);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid var(--accent-border);
        border-radius: 18px;
        padding: 2.2rem;
        margin-bottom: 2rem;
        box-shadow: 0 12px 35px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 230, 160, 0.15);
        transition: all 0.35s ease;
        position: relative;
    }}
    
    .app-card::before {{
        content: '';
        position: absolute;
        top: 0; left: 20px; right: 20px; height: 2px;
        background: linear-gradient(90deg, transparent, var(--accent-border), transparent);
    }}
    
    .app-card:hover {{
        border-color: var(--accent-primary);
        box-shadow: 0 16px 45px var(--accent-glow);
    }}

    /* Custom 24k Gold Foil Shimmer Buttons */
    div.stButton > button {{
        background: linear-gradient(135deg, var(--accent-light) 0%, var(--accent-primary) 35%, var(--accent-secondary) 70%, var(--accent-dark) 100%) !important;
        color: #0A0E1A !important;
        font-family: 'Cinzel', 'Cormorant Garamond', serif !important;
        font-weight: 800 !important;
        font-size: 0.98rem !important;
        border-radius: 12px !important;
        border: 1px solid #FFF8DC !important;
        padding: 0.75rem 2rem !important;
        box-shadow: 0 5px 25px var(--accent-glow) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }}
    
    div.stButton > button:hover {{
        background: linear-gradient(135deg, #FFFFFF 0%, var(--accent-light) 30%, var(--accent-primary) 70%, var(--accent-secondary) 100%) !important;
        box-shadow: 0 8px 35px var(--accent-glow) !important;
        transform: translateY(-3px) !important;
        color: #000000 !important;
    }}

    div.stButton > button:active {{
        transform: translateY(1px) !important;
    }}

    /* Text Inputs, Textareas, Selectboxes */
    .stTextInput input, .stTextArea textarea, div[data-baseweb="select"], .stSelectbox {{
        background-color: rgba(6, 9, 17, 0.9) !important;
        border: 1px solid var(--accent-border) !important;
        color: #F8FAFC !important;
        border-radius: 12px !important;
        font-family: 'Cormorant Garamond', 'Georgia', serif !important;
        font-size: 1.1rem !important;
    }}
    
    .stTextInput input:focus, .stTextArea textarea:focus {{
        border-color: var(--accent-primary) !important;
        box-shadow: 0 0 20px var(--accent-glow) !important;
    }}

    /* Chat Messages Styling */
    .stChatMessage {{
        background: linear-gradient(145deg, rgba(20, 26, 48, 0.75) 0%, rgba(10, 14, 28, 0.85) 100%) !important;
        border: 1px solid var(--accent-border) !important;
        border-radius: 18px !important;
        margin-bottom: 1.3rem !important;
        padding: 1.3rem !important;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.4);
    }}
    
    .stChatMessage[data-testid="stChatMessageUser"] {{
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.8) 0%, rgba(49, 46, 129, 0.4) 100%) !important;
        border-color: var(--accent-border) !important;
    }}

    /* File Uploader Custom Styling */
    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.7) !important;
        border: 2px dashed var(--accent-border) !important;
        border-radius: 16px !important;
        padding: 1.8rem !important;
        transition: all 0.3s ease;
    }}
    
    div[data-testid="stFileUploader"]:hover {{
        border-color: var(--accent-primary) !important;
        background: rgba(30, 41, 59, 0.85) !important;
        box-shadow: 0 0 25px var(--accent-glow);
    }}

    /* Disclaimer Box */
    .app-disclaimer-box {{
        background: linear-gradient(135deg, rgba(20, 26, 48, 0.95) 0%, rgba(4, 6, 14, 0.98) 100%);
        border: 1px solid var(--accent-border);
        border-left: 6px solid var(--accent-primary);
        border-radius: 14px;
        padding: 1.6rem 2rem;
        color: #CBD5E1;
        font-size: 0.9rem;
        margin-top: 4rem;
        box-shadow: 0 10px 35px rgba(0, 0, 0, 0.6);
        position: relative;
    }}

    /* Quick Prompt Chips Title */
    .chip-title {{
        font-size: 0.85rem;
        color: var(--accent-light);
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        margin-bottom: 0.6rem;
    }}
</style>
"""

# Default session state initialization for Theme
if "theme_name" not in st.session_state:
    st.session_state.theme_name = "🏛️ Classic Legal (Default)"
if "custom_hex" not in st.session_state:
    st.session_state.custom_hex = "#00E5FF"

# Inject Dynamic Theme CSS
st.markdown(get_theme_css(st.session_state.theme_name, st.session_state.custom_hex), unsafe_allow_html=True)

# ---------- LOAD ENVIRONMENT & API KEY ----------
load_dotenv()

SYSTEM_PROMPT = """You are Legal Guide AI, a legal concierge assistant that helps ordinary
people in India understand their legal rights and procedures in simple, plain English.
Rules:
1. Explain in simple, everyday language (avoid complex legal jargon).
2. Structure your answers logically with headers: Relevant Rights, Guidance, Next Steps.
3. Always include helpline numbers for urgent situations (e.g. Cyber Crime: 1930, Women Helpline: 181, Police: 112).
4. Always end with a clear disclaimer that this is informational and not formal legal advice.
"""

DISCLAIMER = "\n\n---\n**Disclaimer:** *This is general legal information generated by AI, not formal legal advice. Please consult a licensed advocate for specific legal cases.*"

# Helper function to query Gemini API with fallback models
def ask_gemini(prompt: str, api_key: str, model_name: str = "gemini-3.6-flash") -> str:
    if not api_key or not api_key.strip():
        return "❌ **API Key Missing**: Please enter your Google Gemini API key in the sidebar or update `.env`."
    
    key = api_key.strip()
    genai.configure(api_key=key)
    
    models_to_try = [model_name, "gemini-3.6-flash", "gemini-flash-latest", "gemini-3.5-flash", "gemini-2.0-flash"]
    models_to_try = list(dict.fromkeys(models_to_try))
    
    last_err = ""
    for m in models_to_try:
        try:
            model = genai.GenerativeModel(m)
            response = model.generate_content(prompt)
            if response and response.text:
                return response.text
        except Exception as e:
            last_err = str(e)
            if "404" in last_err or "not found" in last_err or "no longer available" in last_err:
                continue
            elif "429" in last_err or "Quota" in last_err or "ResourceExhausted" in last_err:
                return f"❌ **API Quota Exceeded (429)** for model `{m}`.\n\n*Solution*: Get a free Gemini API key from [Google AI Studio](https://aistudio.google.com/app/apikey) and paste it in the sidebar."
            else:
                return f"❌ **Error from Gemini API ({m})**: {last_err}"
                
    return f"❌ **Model Error**: Could not connect to Gemini models. Details: {last_err}"


# ---------- SIDEBAR & NAVIGATION ----------
with st.sidebar:
    st.markdown("""
        <div style="text-align: center; padding: 1.2rem 0;">
            <div style="font-size: 3.2rem; line-height: 1; filter: drop-shadow(0 0 15px rgba(255, 215, 0, 0.6));">👑</div>
            <h2 style="margin: 0.6rem 0 0 0; font-size: 1.6rem; color: #FFE5A3;">LEGAL AI</h2>
            <div style="color: #FFD700; font-size: 0.8rem; font-weight: 700; letter-spacing: 2.5px; text-transform: uppercase;">Sovereign Concierge</div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.header("🔑 Key Access & Engine")
    
    env_api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or ""
    
    api_key_input = st.text_input(
        "Google Gemini API Key:",
        value=env_api_key,
        type="password",
        help="Get a free key from Google AI Studio (starts with AIzaSy...)"
    )
    
    if api_key_input:
        if api_key_input.startswith("AIzaSy"):
            st.success("✅ Active Key (`AIzaSy...`)")
        else:
            st.info("ℹ️ Custom Key Detected")
    else:
        st.warning("⚠️ Key Required")
        
    model_choice = st.selectbox(
        "AI Model Engine:",
        ["gemini-3.6-flash", "gemini-flash-latest", "gemini-3.5-flash", "gemini-2.0-flash"],
        index=0
    )
    
    # ---- Theme selection ----
    theme_options = [
        "🏛️ Classic Legal (Default)",
        "📜 Classic Parchment & Gold",
        "👑 Royal Gold",
        "💎 Imperial Emerald",
        "🔮 Cyber Sapphire",
        "💜 Amethyst Velvet",
        "🩸 Ruby Crimson",
        "⚡ Cyberpunk Amber",
        "Custom"
    ]
    selected_theme = st.selectbox("Select Theme:", theme_options, index=0, key="theme_select")
    if selected_theme == "Custom":
        custom_color = st.color_picker("Pick a custom accent color", "#00E5FF")
        st.session_state.custom_hex = custom_color
    else:
        st.session_state.custom_hex = "#00E5FF"
    st.session_state.theme_name = selected_theme
    
    st.markdown("---")
    st.header("🏛️ Sovereign Suite")
    app_mode = st.radio(
        "Select Feature:",
        [
            "💬 AI Legal Chatbot",
            "⚖️ Rights Finder",
            "📜 Law Simplifier",
            "🛡️ Scam & Fraud Checker",
            "📑 Document Analyzer",
            "🚨 Emergency Legal Help",
            "📋 Legal Procedure Guide"
        ]
    )
    
    st.markdown("---")
    st.info(
        "🔑 **Free Key Setup:**\n"
        "1. Visit [Google AI Studio](https://aistudio.google.com/app/apikey)\n"
        "2. Click **Create API Key**\n"
        "3. Paste key above or save in `.env`."
    )

# ---------- HEADER BANNER ----------
st.markdown("""
<div class="app-header">
    <div class="app-insignia">👑</div>
    <div class="app-badge">✦ SOVEREIGN LEGAL ATELIER ✦</div>
    <div class="app-title">LEGAL GUIDE AI</div>
    <div class="app-subtitle">Sovereign Legal Intelligence • Statutory Rights Protection • Plain English Legal Simplification</div>
</div>
""", unsafe_allow_html=True)

# ---------- LIVE METRICS DASHBOARD ----------
st.markdown("""
<div class="metrics-grid">
    <div class="app-metric-card">
        <div class="app-metric-icon">🔒</div>
        <div>
            <div class="app-metric-val">256-BIT</div>
            <div class="app-metric-lbl">Private Sovereign Encryption</div>
        </div>
    </div>
    <div class="app-metric-card">
        <div class="app-metric-icon">🏛️</div>
        <div>
            <div class="app-metric-val">INDIAN LAW</div>
            <div class="app-metric-lbl">Constitution & Rights Grounded</div>
        </div>
    </div>
    <div class="app-metric-card">
        <div class="app-metric-icon">⚡</div>
        <div>
            <div class="app-metric-val">SUB-SECOND</div>
            <div class="app-metric-lbl">Generative Response</div>
        </div>
    </div>
    <div class="app-metric-card">
        <div class="app-metric-icon">💎</div>
        <div>
            <div class="app-metric-val">FREE ACCESS</div>
            <div class="app-metric-lbl">Public Legal Empowerment</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# ===========================================================================
# 1. 💬 AI LEGAL CHATBOT
# ===========================================================================
if app_mode == "💬 AI Legal Chatbot":
    st.markdown("""
    <div class="app-card">
        <h3 style="margin-top:0; color:#FFE5A3;">💬 AI Legal Assistant Concierge</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Consult the Sovereign Legal Concierge on Indian law, tenant protection, employee rights, or civil procedures.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="chip-title">💡 Popular Queries:</div>', unsafe_allow_html=True)
    chip_col1, chip_col2, chip_col3, chip_col4 = st.columns(4)
    
    prompt_to_set = None
    with chip_col1:
        if st.button("🏠 Landlord Eviction Notice", key="chip_eviction"):
            prompt_to_set = "What are my rights if my landlord asks me to vacate without 30 days written notice in India?"
    with chip_col2:
        if st.button("💼 Salary Withheld", key="chip_salary"):
            prompt_to_set = "My company is delaying my full & final settlement after resignation. What legal remedies do I have?"
    with chip_col3:
        if st.button("🛒 Defective Laptop Refund", key="chip_consumer"):
            prompt_to_set = "How can I claim a refund for a defective laptop from an online seller under Consumer Protection Act?"
    with chip_col4:
        if st.button("💳 Bank Cyber Fraud", key="chip_cyber"):
            prompt_to_set = "Someone stole Rs 25,000 from my account via a fake UPI link. What are the immediate legal steps?"

    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Welcome to **Legal Guide AI Sovereign Atelier**. How may I assist you with your statutory rights or administrative legal procedures today?"}
        ]
        
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_input = st.chat_input("e.g. What are my rights if my landlord asks me to vacate without 30 days notice?")
    
    active_prompt = prompt_to_set or user_input
    
    if active_prompt:
        st.session_state.messages.append({"role": "user", "content": active_prompt})
        with st.chat_message("user"):
            st.markdown(active_prompt)
            
        with st.chat_message("assistant"):
            with st.spinner("Analyzing Indian Legal Statutes & Sovereign Rights Database..."):
                prompt = f"{SYSTEM_PROMPT}\n\nUser Question: {active_prompt}"
                answer = ask_gemini(prompt, api_key_input, model_choice)
                if not answer.endswith("advice."):
                    answer += DISCLAIMER
                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})


# ===========================================================================
# 2. ⚖️ RIGHTS FINDER
# ===========================================================================
elif app_mode == "⚖️ Rights Finder":
    st.markdown("""
    <div class="app-card">
        <h3 style="margin-top:0; color:#FFE5A3;">⚖️ Legal Rights Finder</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Describe your situation to identify protecting acts, notice periods, and statutory legal remedies under Indian law.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="chip-title">💡 Sample Cases:</div>', unsafe_allow_html=True)
    sc1, sc2, sc3 = st.columns(3)
    preset_situation = ""
    with sc1:
        if st.button("🏢 Fired Without Notice", key="rights_fired"):
            preset_situation = "My company terminated my employment immediately without giving 1 month notice or pay in lieu of notice."
    with sc2:
        if st.button("💰 Security Deposit Refund", key="rights_deposit"):
            preset_situation = "My landlord refuses to return my Rs 60,000 security deposit after I vacated the flat in clean condition."
    with sc3:
        if st.button("📞 Unsolicited Harassment Calls", key="rights_calls"):
            preset_situation = "Recovery agents are calling my contacts and harassing me for a loan I never took."

    situation = st.text_area(
        "Describe your situation in detail:",
        value=preset_situation,
        placeholder="e.g. My employer is withholding my salary for 2 months after I submitted my resignation...",
        height=150
    )
    
    if st.button("Analyze & Find My Rights", type="primary"):
        if not situation.strip():
            st.warning("⚠️ Please describe your situation above.")
        else:
            with st.spinner("Evaluating legal statutes & rights..."):
                prompt = f"""{SYSTEM_PROMPT}

A user described the following situation:
"{situation}"

Please analyze and provide:
1. **Applicable Legal Rights**: What rights protect the user under Indian Law?
2. **Key Timelines & Notice Periods**: Mandatory notice periods or legal deadlines.
3. **Actionable Next Steps**: Concrete steps to resolve this issue.
4. **Relevant Authorities**: Government bodies, ombudsmen, or forums to contact.
"""
                result = ask_gemini(prompt, api_key_input, model_choice)
                st.markdown("""
                <div class="app-card">
                    <h3 style="color:#FFE5A3;">🏛️ Legal Analysis & Statutory Rights</h3>
                </div>
                """, unsafe_allow_html=True)
                st.markdown(result + DISCLAIMER)


# ===========================================================================
# 3. 📜 LAW SIMPLIFIER
# ===========================================================================
elif app_mode == "📜 Law Simplifier":
    st.markdown("""
    <div class="app-card">
        <h3 style="margin-top:0; color:#FFE5A3;">📜 Legal Text Simplifier</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Translate complex legal jargon, contract clauses, or court notices into clear, plain English.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="chip-title">💡 Try Sample Legal Clauses:</div>', unsafe_allow_html=True)
    lc1, lc2 = st.columns(2)
    sample_clause = ""
    with lc1:
        if st.button("📜 Termination Clause Example", key="clause_term"):
            sample_clause = "Notwithstanding anything contained herein to the contrary, the Lessor reserves the absolute right to terminate this lease covenant unilaterally with 7 days notice without forfeiture penalty."
    with lc2:
        if st.button("📜 Indemnity & Liability Clause Example", key="clause_indem"):
            sample_clause = "The Employee agrees to indemnify, defend, and hold harmless the Company against all third-party claims, liabilities, and legal costs arising from direct or indirect performance of duties."

    legal_text = st.text_area(
        "Paste Legal Text / Clause Here:",
        value=sample_clause,
        placeholder="e.g. Notwithstanding anything contained herein to the contrary, the Lessor reserves the absolute right...",
        height=160
    )
    
    if st.button("Translate To Plain English", type="primary"):
        if not legal_text.strip():
            st.warning("⚠️ Please paste legal text to simplify.")
        else:
            with st.spinner("Translating legal terminology into plain English..."):
                prompt = f"""Translate the following complex legal text into simple, easy-to-understand plain English. 
Highlight any potential risks or hidden obligations.

Legal Text:
\"\"\"{legal_text}\"\"\"
"""
                simplified = ask_gemini(prompt, api_key_input, model_choice)
                st.success("✨ Simplified Plain English Translation:")
                st.markdown(simplified)


# ===========================================================================
# 4. 🛡️ SCAM & FRAUD CHECKER
# ===========================================================================
elif app_mode == "🛡️ Scam & Fraud Checker":
    st.markdown("""
    <div class="app-card">
        <h3 style="margin-top:0; color:#FFE5A3;">🛡️ Scam & Cyber Fraud Detector</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Inspect suspicious SMS messages, emails, lottery claims, or job offers for fraud red flags.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="chip-title">💡 Try Sample Fraud Messages:</div>', unsafe_allow_html=True)
    mc1, mc2 = st.columns(2)
    sample_scam = ""
    with mc1:
        if st.button("🚨 Bank Account Blocked SMS", key="scam_bank"):
            sample_scam = "URGENT: Your SBI bank account will be suspended today due to missing PAN. Click http://bit.ly/sbi-verify-now to complete KYC or call 9876543210 immediately."
    with mc2:
        if st.button("💼 Work-from-home Telegram Scam", key="scam_job"):
            sample_scam = "Earn Rs 5,000 per day reviewing Youtube videos! No experience needed. Contact HR manager on Telegram @JobOffer2026 immediately to start."

    scam_msg = st.text_area(
        "Paste Message / Email / Offer:",
        value=sample_scam,
        placeholder="e.g. Urgent! Your bank account has been blocked. Click here to verify OTP immediately...",
        height=150
    )
    
    if st.button("Inspect For Scam Risks", type="primary"):
        if not scam_msg.strip():
            st.warning("⚠️ Please paste a message to inspect.")
        else:
            flags = []
            if re.search(r"http[s]?://[^\s]+", scam_msg):
                flags.append("🔗 External Link Detected")
            if re.search(r"(otp|urgent|verify now|account blocked|lottery|prize|click here|claim reward|job offer|kyc)", scam_msg, re.IGNORECASE):
                flags.append("🚨 Common Phishing / Urgency Keywords Detected")
            if re.search(r"(\b\d{10}\b|whatsapp|telegram)", scam_msg, re.IGNORECASE):
                flags.append("📱 Off-platform Contact Request (WhatsApp/Telegram)")
                
            st.markdown("### 🤖 Security Scan Results:")
            if flags:
                for f in flags:
                    st.error(f)
            else:
                st.success("✅ No common automated scam keywords found.")
                
            with st.spinner("Running AI Phishing & Fraud Analysis..."):
                prompt = f"""Analyze the following message for potential scam, phishing, or financial fraud risks under Indian Cyber Law.
Explain the red flags, common scam tactics used, and what immediate protective actions the user should take.

Message to analyze:
\"\"\"{scam_msg}\"\"\"
"""
                analysis = ask_gemini(prompt, api_key_input, model_choice)
                st.markdown("### 🔍 AI Fraud Analysis & Protective Recommendations")
                st.markdown(analysis)


# ===========================================================================
# 5. 📑 DOCUMENT ANALYZER
# ===========================================================================
elif app_mode == "📑 Document Analyzer":
    st.markdown("""
    <div class="app-card">
        <h3 style="margin-top:0; color:#FFE5A3;">📑 Document & Contract Intelligence</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Upload rental agreements, contracts, or employment offer letters (.pdf or .docx) for automatic risk analysis.</p>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload Legal Document (.pdf or .docx):", type=["pdf", "docx"])
    
    if uploaded_file is not None:
        file_text = ""
        try:
            if uploaded_file.name.endswith(".pdf"):
                with pdfplumber.open(uploaded_file) as pdf:
                    for page in pdf.pages:
                        extracted = page.extract_text()
                        if extracted:
                            file_text += extracted + "\n"
            elif uploaded_file.name.endswith(".docx"):
                doc = DocxDocument(uploaded_file)
                file_text = "\n".join([p.text for p in doc.paragraphs])
        except Exception as e:
            st.error(f"❌ Error parsing file: {e}")
            
        if file_text.strip():
            st.success(f"📄 Successfully Parsed **{uploaded_file.name}** ({len(file_text)} characters)")
            with st.expander("🔍 View Extracted Document Preview"):
                st.text(file_text[:2500] + ("..." if len(file_text) > 2500 else ""))
                
            if st.button("Run AI Contract Review", type="primary"):
                with st.spinner("Extracting key clauses & evaluating liabilities..."):
                    prompt = f"""Analyze this legal document carefully. Provide a structured review covering:
1. **Document Summary**: Core purpose of this document.
2. **Important Clauses & Obligations**: Key commitments required from signee.
3. **Red Flags & Risks**: Unfair penalties, hidden charges, or liabilities.
4. **Questions to Ask Before Signing**: 3-5 critical questions to clarify.

Document Text:
\"\"\"{file_text[:7500]}\"\"\"
"""
                    doc_analysis = ask_gemini(prompt, api_key_input, model_choice)
                    st.markdown("## 🔍 Executive Contract Review")
                    st.markdown(doc_analysis + DISCLAIMER)
        else:
            st.error("❌ No text could be extracted from this file.")


# ===========================================================================
# 6. 🚨 EMERGENCY LEGAL HELP
# ===========================================================================
elif app_mode == "🚨 Emergency Legal Help":
    st.markdown("""
    <div class="app-card" style="border-color: rgba(239, 68, 68, 0.5);">
        <h3 style="margin-top:0; color:#FCA5A5;">🚨 Emergency Assistance & National Helplines</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Immediate action protocols and official national emergency contact numbers in India.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div style="background: rgba(239, 68, 68, 0.18); border: 1px solid rgba(239, 68, 68, 0.5); border-radius: 14px; padding: 1.4rem; text-align: center;">
            <div style="font-size: 2rem;">🚨</div>
            <h4 style="color: #FCA5A5; margin: 0.4rem 0; font-family:'Cinzel',serif;">Cyber Crime Portal</h4>
            <div style="font-size: 1.5rem; font-weight: 800; color: #FFF;">Call 1930</div>
            <div style="font-size: 0.82rem; color: #94A3B8; margin-top: 0.3rem;"><a href="https://cybercrime.gov.in" target="_blank" style="color:#FFE5A3;">cybercrime.gov.in</a></div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div style="background: rgba(255, 215, 0, 0.15); border: 1px solid rgba(255, 215, 0, 0.5); border-radius: 14px; padding: 1.4rem; text-align: center;">
            <div style="font-size: 2rem;">👩</div>
            <h4 style="color: #FFE5A3; margin: 0.4rem 0; font-family:'Cinzel',serif;">Women Helpline</h4>
            <div style="font-size: 1.5rem; font-weight: 800; color: #FFF;">Call 181 / 1091</div>
            <div style="font-size: 0.82rem; color: #94A3B8; margin-top: 0.3rem;">24/7 Protection & Safety</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div style="background: rgba(59, 130, 246, 0.18); border: 1px solid rgba(59, 130, 246, 0.5); border-radius: 14px; padding: 1.4rem; text-align: center;">
            <div style="font-size: 2rem;">🚓</div>
            <h4 style="color: #93C5FD; margin: 0.4rem 0; font-family:'Cinzel',serif;">National Police</h4>
            <div style="font-size: 1.5rem; font-weight: 800; color: #FFF;">Call 112 / 100</div>
            <div style="font-size: 0.82rem; color: #94A3B8; margin-top: 0.3rem;">Immediate Law Enforcement</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("---")
    issue_type = st.selectbox(
        "Select Urgent Emergency Scenario:",
        ["Cyber Fraud / Financial Theft", "Domestic Violence / Abuse", "Harassment / Stalking", "Unlawful Police Arrest / Custody"]
    )
    
    if st.button("Generate Emergency Protocol Guide", type="primary"):
        with st.spinner("Generating emergency protocol..."):
            prompt = f"""{SYSTEM_PROMPT}

The user is experiencing an urgent emergency regarding: "{issue_type}".
Provide immediate, clear, 1-2-3 step instructions on:
1. Immediate safety measures & evidence preservation
2. Law enforcement reporting procedure
3. Relevant helpline numbers & support organisations in India.
"""
            emergency_guide = ask_gemini(prompt, api_key_input, model_choice)
            st.markdown(emergency_guide + DISCLAIMER)


# ===========================================================================
# 7. 📋 LEGAL PROCEDURE GUIDE
# ===========================================================================
elif app_mode == "📋 Legal Procedure Guide":
    st.markdown("""
    <div class="app-card">
        <h3 style="margin-top:0; color:#FFE5A3;">📋 Step-by-Step Legal Procedures</h3>
        <p style="color: #CBD5E1; margin-bottom: 0;">Comprehensive walkthroughs for administrative, civil, and police procedures in India.</p>
    </div>
    """, unsafe_allow_html=True)
    
    procedure_topic = st.selectbox(
        "Select Official Legal Procedure:",
        [
            "How to File an FIR at a Police Station",
            "How to File a Consumer Forum Complaint for Defective Product/Service",
            "How to File an RTI (Right to Information) Application",
            "Tenant Rights & Legal Eviction Notice Rules",
            "Employee Rights for Salary Delay or Illegal Termination"
        ]
    )
    
    if st.button("Generate Step-by-Step Walkthrough", type="primary"):
        with st.spinner("Retrieving procedure manual..."):
            prompt = f"""{SYSTEM_PROMPT}

Provide a comprehensive, clear, step-by-step guide for: "{procedure_topic}" in India.
Include:
1. **Required Documents**: List of prerequisite papers.
2. **Step-by-Step Process**: Exact procedure & submission details.
3. **Timeline & Standard Costs**: Official timelines & statutory fees.
4. **Escalation Path**: What to do if ignored or rejected.
"""
            proc_guide = ask_gemini(prompt, api_key_input, model_choice)
            st.markdown(proc_guide + DISCLAIMER)


# ---------- FOOTER DISCLAIMER ----------
st.markdown("""
<div class="app-disclaimer-box">
    <div style="font-family: 'Cinzel Decorative', 'Cinzel', serif; color: #FFE5A3; font-weight: 800; margin-bottom: 0.5rem; font-size: 1.05rem; letter-spacing: 1px;">👑 SOVEREIGN LEGAL DISCLAIMER</div>
    Legal Guide AI is an artificial intelligence platform designed strictly for informational and educational purposes. It does not constitute formal legal representation or binding legal advice. Please consult a licensed advocate or attorney for your specific legal case.
</div>
""", unsafe_allow_html=True)