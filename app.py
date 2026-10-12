"""
DataPilot — Advanced Data Analytics & Storytelling Chat App (Gen-Z Aesthetic Edition)
Free stack: Python, Streamlit, Pandas, Matplotlib, python-pptx.
Works 100% offline without any paid API keys.
Deterministic intent parsing + Pandas calculations + Automated PowerPoint slide generation.
"""
from io import BytesIO
from pathlib import Path
import re
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Configure page with custom icon & wide layout
st.set_page_config(
    page_title="DataPilot // Next-Gen Data Intelligence",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------- Dark Mode & Neon Matplotlib Theme -------------------------
def apply_matplotlib_theme():
    plt.style.use('dark_background')
    plt.rcParams.update({
        'figure.facecolor': '#0d1117',
        'axes.facecolor': '#0d1117',
        'savefig.facecolor': '#0d1117',
        'grid.color': '#21262d',
        'grid.linestyle': '--',
        'grid.alpha': 0.6,
        'text.color': '#f0f6fc',
        'axes.labelcolor': '#c9d1d9',
        'xtick.color': '#8b949e',
        'ytick.color': '#8b949e',
        'axes.edgecolor': '#30363d',
        'axes.linewidth': 1.2,
        'font.sans-serif': ['DejaVu Sans', 'Arial', 'Helvetica'],
    })

apply_matplotlib_theme()
NEON_COLORS = ['#00F0FF', '#FF007F', '#A855F7', '#00FFA3', '#FFE600', '#FF6B00', '#3B82F6', '#EC4899']

# ------------------------- Gen-Z / Cyber-Chic Glassmorphic CSS -------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&display=swap');

:root {
  --dp-bg: #07090e;
  --dp-panel: rgba(17, 21, 33, 0.72);
  --dp-panel-solid: #0f131f;
  --dp-purple: #8B5CF6;
  --dp-pink: #EC4899;
  --dp-cyan: #06B6D4;
  --dp-lime: #10B981;
  --dp-yellow: #F59E0B;
  --dp-text: #F8FAFC;
  --dp-muted: #94A3B8;
  --dp-border: rgba(255, 255, 255, 0.08);
  --dp-glow-purple: 0 0 25px rgba(139, 92, 246, 0.35);
  --dp-glow-pink: 0 0 25px rgba(236, 72, 153, 0.35);
  --dp-glow-cyan: 0 0 25px rgba(6, 182, 212, 0.35);
  --dp-grad-primary: linear-gradient(135deg, #FF007A 0%, #7928CA 50%, #00DFD8 100%);
  --dp-grad-vibe: linear-gradient(135deg, #8B5CF6 0%, #EC4899 100%);
  --dp-grad-cyber: linear-gradient(135deg, #06B6D4 0%, #3B82F6 100%);
}

/* Base Body Styling */
html, body, [class*="css"] {
  font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
}

.stApp {
  background-color: var(--dp-bg);
  background-image: 
    radial-gradient(at 0% 0%, rgba(139, 92, 246, 0.12) 0px, transparent 50%),
    radial-gradient(at 100% 0%, rgba(6, 182, 212, 0.10) 0px, transparent 50%),
    radial-gradient(at 50% 100%, rgba(236, 72, 153, 0.08) 0px, transparent 50%);
  background-attachment: fixed;
  color: var(--dp-text);
}

.block-container {
  padding-top: 0.5rem;
  padding-bottom: 3.5rem;
  max-width: 1650px;
}

/* Header & Sidebar */
header[data-testid="stHeader"] {
  background: rgba(7, 9, 14, 0.85);
  backdrop-filter: blur(16px);
  border-bottom: 1px solid var(--dp-border);
}

section[data-testid="stSidebar"] {
  background: rgba(11, 14, 22, 0.85) !important;
  backdrop-filter: blur(20px);
  border-right: 1px solid var(--dp-border);
}

section[data-testid="stSidebar"] * {
  color: var(--dp-text);
}

/* Sidebar Branding Banner */
.dp-sidebar-hero {
  padding: 1.1rem 0.9rem;
  margin-bottom: 1.2rem;
  background: linear-gradient(145deg, rgba(139, 92, 246, 0.15), rgba(236, 72, 153, 0.08));
  border: 1px solid rgba(139, 92, 246, 0.3);
  border-radius: 16px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
}
.dp-sidebar-badge {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: rgba(139, 92, 246, 0.25);
  border: 1px solid rgba(139, 92, 246, 0.5);
  padding: 3px 9px;
  border-radius: 999px;
  font-size: 0.68rem;
  font-weight: 800;
  letter-spacing: 0.1em;
  color: #E9D5FF;
  text-transform: uppercase;
}
.dp-sidebar-badge span.pulse-dot {
  width: 7px;
  height: 7px;
  background: #10B981;
  border-radius: 50%;
  display: inline-block;
  box-shadow: 0 0 8px #10B981;
  animation: pulse 1.5s infinite;
}
@keyframes pulse {
  0% { transform: scale(0.95); opacity: 0.8; }
  50% { transform: scale(1.3); opacity: 1; }
  100% { transform: scale(0.95); opacity: 0.8; }
}

.dp-sidebar-title {
  font-family: 'Outfit', sans-serif;
  font-size: 1.35rem;
  font-weight: 800;
  margin: 0.5rem 0 0.2rem;
  background: var(--dp-grad-primary);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  letter-spacing: -0.02em;
}
.dp-sidebar-sub {
  font-size: 0.78rem;
  color: var(--dp-muted);
  line-height: 1.3;
}

/* File Uploader Custom Gen-Z Styling */
[data-testid="stFileUploader"] {
  background: rgba(18, 22, 34, 0.6);
  border: 2px dashed rgba(139, 92, 246, 0.4);
  border-radius: 16px;
  padding: 12px;
  transition: all 0.3s ease;
}
[data-testid="stFileUploader"]:hover {
  border-color: #EC4899;
  background: rgba(236, 72, 153, 0.05);
  box-shadow: 0 0 20px rgba(236, 72, 153, 0.2);
}

/* Metrics Cards - Glassmorphism on Steroids */
[data-testid="stMetric"] {
  background: linear-gradient(135deg, rgba(255, 255, 255, 0.04) 0%, rgba(255, 255, 255, 0.01) 100%) !important;
  backdrop-filter: blur(16px);
  border: 1px solid var(--dp-border) !important;
  border-top: 2px solid var(--dp-purple) !important;
  padding: 18px 20px !important;
  border-radius: 18px !important;
  box-shadow: 0 8px 28px rgba(0, 0, 0, 0.35);
  transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
  position: relative;
  overflow: hidden;
}
[data-testid="stMetric"]:hover {
  transform: translateY(-4px);
  border-color: rgba(236, 72, 153, 0.5) !important;
  box-shadow: 0 14px 32px rgba(139, 92, 246, 0.25);
}
[data-testid="stMetricLabel"] {
  color: #94A3B8 !important;
  font-size: 0.78rem !important;
  font-weight: 700 !important;
  letter-spacing: 0.12em !important;
  text-transform: uppercase !important;
}
[data-testid="stMetricValue"] {
  color: #FFFFFF !important;
  font-family: 'Outfit', sans-serif !important;
  font-size: 2.1rem !important;
  font-weight: 800 !important;
  text-shadow: 0 0 15px rgba(255, 255, 255, 0.3);
}

/* Tabs: Floating Capsule Bar */
[data-testid="stTabs"] {
  background: rgba(15, 18, 28, 0.7);
  padding: 6px;
  border-radius: 18px;
  border: 1px solid var(--dp-border);
  backdrop-filter: blur(12px);
  margin-top: 1rem;
  margin-bottom: 1.5rem;
}
[data-testid="stTabs"] button[role="tab"] {
  color: #94A3B8 !important;
  border-radius: 12px !important;
  padding: 0.7rem 1.25rem !important;
  font-family: 'Outfit', sans-serif !important;
  font-weight: 700 !important;
  font-size: 0.95rem !important;
  border: none !important;
  background: transparent !important;
  transition: all 0.25s ease !important;
}
[data-testid="stTabs"] button[role="tab"]:hover {
  color: #FFFFFF !important;
  background: rgba(255, 255, 255, 0.05) !important;
}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
  color: #FFFFFF !important;
  background: var(--dp-grad-vibe) !important;
  box-shadow: 0 4px 18px rgba(139, 92, 246, 0.45) !important;
}

/* Buttons: High-Energy Gradients */
.stButton > button, .stDownloadButton > button {
  background: var(--dp-grad-vibe) !important;
  color: #FFFFFF !important;
  border: none !important;
  border-radius: 12px !important;
  padding: 0.65rem 1.4rem !important;
  font-family: 'Outfit', sans-serif !important;
  font-weight: 700 !important;
  letter-spacing: 0.03em !important;
  box-shadow: 0 4px 18px rgba(139, 92, 246, 0.35) !important;
  transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
  transform: translateY(-2px) scale(1.02) !important;
  box-shadow: 0 8px 26px rgba(236, 72, 153, 0.55) !important;
  background: linear-gradient(135deg, #9333EA 0%, #F43F5E 100%) !important;
}
.stButton > button:active, .stDownloadButton > button:active {
  transform: scale(0.98) !important;
}

/* DataFrames & Tables */
[data-testid="stDataFrame"], [data-testid="stTable"] {
  border: 1px solid var(--dp-border) !important;
  border-radius: 16px !important;
  overflow: hidden !important;
  background: rgba(14, 18, 28, 0.6) !important;
  box-shadow: 0 6px 20px rgba(0, 0, 0, 0.35) !important;
}

/* Chat Messages */
[data-testid="stChatMessage"] {
  background: rgba(18, 22, 34, 0.65) !important;
  border: 1px solid var(--dp-border) !important;
  border-radius: 18px !important;
  padding: 1rem 1.25rem !important;
  margin-bottom: 0.85rem !important;
  backdrop-filter: blur(10px) !important;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25) !important;
}
[data-testid="stChatMessage"][data-role="user"] {
  border-left: 3px solid var(--dp-pink) !important;
}
[data-testid="stChatMessage"][data-role="assistant"] {
  border-left: 3px solid var(--dp-cyan) !important;
}

/* Inputs & Select Boxes */
input, textarea, [data-baseweb="select"] > div {
  background-color: rgba(16, 20, 31, 0.85) !important;
  border: 1px solid rgba(255, 255, 255, 0.12) !important;
  border-radius: 12px !important;
  color: #FFFFFF !important;
}
input:focus, textarea:focus, [data-baseweb="select"] > div:focus-within {
  border-color: var(--dp-purple) !important;
  box-shadow: 0 0 15px rgba(139, 92, 246, 0.4) !important;
}

/* Glass Panels & Modern Cards */
.dp-glass-card {
  background: linear-gradient(135deg, rgba(255, 255, 255, 0.04) 0%, rgba(255, 255, 255, 0.01) 100%);
  border: 1px solid var(--dp-border);
  border-radius: 18px;
  padding: 1.4rem;
  backdrop-filter: blur(16px);
  margin-bottom: 1.2rem;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3);
}

.dp-tag-pill {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  background: rgba(255, 255, 255, 0.06);
  border: 1px solid rgba(255, 255, 255, 0.1);
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 0.72rem;
  font-weight: 600;
  color: var(--dp-text);
}

/* Workspace Headline */
.dp-workspace-heading {
  padding: 1rem 0 0.4rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 12px;
}
.dp-kicker {
  background: var(--dp-grad-primary);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  letter-spacing: 0.22em;
  font-size: 0.75rem;
  font-weight: 800;
  text-transform: uppercase;
}
.dp-workspace-heading h2 {
  font-family: 'Outfit', sans-serif;
  color: #FFFFFF;
  font-size: clamp(1.6rem, 2.5vw, 2.3rem);
  font-weight: 800;
  letter-spacing: -0.02em;
  line-height: 1.1;
  margin: 0.3rem 0;
}
.dp-workspace-heading p {
  color: var(--dp-muted);
  font-size: 0.95rem;
  margin: 0;
}

/* Custom Scrollbar */
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: #07090e; }
::-webkit-scrollbar-thumb {
  background: linear-gradient(180deg, #8B5CF6, #EC4899);
  border-radius: 4px;
}
::-webkit-scrollbar-thumb:hover { background: #EC4899; }
</style>
""", unsafe_allow_html=True)

# ------------------------- Mind-Blowing Gen-Z Interactive Hero HTML -------------------------
HERO_HTML = r'''<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>DataPilot // Next-Gen Data Intelligence</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800;900&family=Plus+Jakarta+Sans:wght@400;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body {
      height: 100%;
      background: #07090e;
      overflow: hidden;
      font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
      color: #F8FAFC;
      -webkit-font-smoothing: antialiased;
    }

    /* Ambient Canvas & Mesh Container */
    .hero-container {
      position: relative;
      width: 100%;
      height: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 24px 20px;
      overflow: hidden;
    }

    canvas#particle-canvas {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      z-index: 0;
      pointer-events: none;
    }

    /* Glowing Radial Mesh Auroras */
    .aurora-blob {
      position: absolute;
      border-radius: 50%;
      filter: blur(90px);
      opacity: 0.35;
      pointer-events: none;
      z-index: 1;
      animation: floatAurora 14s ease-in-out infinite alternate;
    }
    .blob-1 {
      top: -10%; left: 15%;
      width: 500px; height: 500px;
      background: radial-gradient(circle, #8B5CF6 0%, transparent 70%);
    }
    .blob-2 {
      bottom: -15%; right: 10%;
      width: 550px; height: 550px;
      background: radial-gradient(circle, #EC4899 0%, transparent 70%);
      animation-delay: -5s;
    }
    .blob-3 {
      top: 30%; right: 25%;
      width: 450px; height: 450px;
      background: radial-gradient(circle, #06B6D4 0%, transparent 70%);
      animation-delay: -9s;
    }
    @keyframes floatAurora {
      0% { transform: translate(0, 0) scale(1); }
      50% { transform: translate(40px, -30px) scale(1.1); }
      100% { transform: translate(-30px, 40px) scale(0.95); }
    }

    /* Content Wrapper */
    .hero-content {
      position: relative;
      z-index: 2;
      max-width: 1200px;
      width: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      text-align: center;
    }

    /* Eyebrow Pill */
    .eyebrow-pill {
      display: inline-flex;
      align-items: center;
      gap: 9px;
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid rgba(255, 255, 255, 0.12);
      backdrop-filter: blur(16px);
      padding: 6px 16px;
      border-radius: 999px;
      font-size: 0.78rem;
      font-weight: 700;
      letter-spacing: 0.14em;
      text-transform: uppercase;
      color: #E2E8F0;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      margin-bottom: 1.1rem;
      animation: fadeInDown 0.8s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    .eyebrow-pill .dot {
      width: 8px; height: 8px;
      background: #10B981;
      border-radius: 50%;
      box-shadow: 0 0 10px #10B981;
      animation: pulse 1.8s infinite;
    }

    /* Mega Gen-Z Headline */
    .mega-headline {
      font-family: 'Outfit', sans-serif;
      font-weight: 900;
      font-size: clamp(2.4rem, 5.4vw, 4.8rem);
      line-height: 1.02;
      letter-spacing: -0.035em;
      text-transform: uppercase;
      margin-bottom: 1.1rem;
      animation: fadeInUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) 0.15s both;
    }
    .text-grad-1 {
      background: linear-gradient(135deg, #FFFFFF 20%, #CBD5E1 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      display: block;
    }
    .text-grad-2 {
      background: linear-gradient(135deg, #FF007A 0%, #7928CA 45%, #00DFD8 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      display: block;
      filter: drop-shadow(0 0 35px rgba(236, 72, 153, 0.45));
    }
    .text-grad-3 {
      background: linear-gradient(135deg, #A855F7 0%, #EC4899 50%, #F59E0B 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      display: inline-block;
    }

    /* Subtext */
    .hero-subtext {
      max-width: 760px;
      font-size: clamp(0.95rem, 1.4vw, 1.2rem);
      color: #94A3B8;
      line-height: 1.5;
      font-weight: 400;
      margin-bottom: 2rem;
      animation: fadeInUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) 0.3s both;
    }

    /* 4 Gen-Z Floating Feature Cards */
    .feature-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 16px;
      width: 100%;
      max-width: 1100px;
      margin-bottom: 1.8rem;
      animation: fadeInUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) 0.45s both;
    }
    .feature-card {
      background: rgba(18, 22, 34, 0.65);
      backdrop-filter: blur(18px);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 18px;
      padding: 16px 14px;
      text-align: left;
      position: relative;
      overflow: hidden;
      transition: all 0.35s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .feature-card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 2px;
      background: linear-gradient(90deg, transparent, var(--card-color, #8B5CF6), transparent);
      opacity: 0.8;
    }
    .feature-card:hover {
      transform: translateY(-5px);
      border-color: rgba(255, 255, 255, 0.22);
      box-shadow: 0 14px 30px rgba(0, 0, 0, 0.5), 0 0 20px var(--card-glow, rgba(139, 92, 246, 0.25));
    }
    .card-icon {
      font-size: 1.5rem;
      margin-bottom: 8px;
      display: inline-block;
    }
    .card-title {
      font-family: 'Outfit', sans-serif;
      font-weight: 700;
      font-size: 0.96rem;
      color: #FFFFFF;
      margin-bottom: 4px;
      letter-spacing: -0.01em;
    }
    .card-desc {
      font-size: 0.78rem;
      color: #94A3B8;
      line-height: 1.35;
    }

    /* Live Status Chips Bar */
    .status-bar {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 18px;
      flex-wrap: wrap;
      animation: fadeInUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) 0.6s both;
    }
    .status-item {
      display: flex;
      align-items: center;
      gap: 6px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 0.74rem;
      font-weight: 600;
      color: #CBD5E1;
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(255, 255, 255, 0.06);
      padding: 5px 12px;
      border-radius: 999px;
    }

    @keyframes fadeInUp {
      from { opacity: 0; transform: translateY(22px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @keyframes fadeInDown {
      from { opacity: 0; transform: translateY(-15px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @keyframes pulse {
      0%, 100% { transform: scale(1); opacity: 0.8; }
      50% { transform: scale(1.35); opacity: 1; }
    }

    @media (max-width: 900px) {
      .feature-grid { grid-template-columns: repeat(2, 1fr); gap: 12px; }
      .hero-container { padding: 18px 12px; }
    }
    @media (max-width: 550px) {
      .feature-grid { grid-template-columns: 1fr; }
      .mega-headline { font-size: 2.2rem; }
      .status-bar { gap: 8px; }
    }
  </style>
</head>
<body>
  <div class="hero-container">
    <div class="aurora-blob blob-1"></div>
    <div class="aurora-blob blob-2"></div>
    <div class="aurora-blob blob-3"></div>
    <canvas id="particle-canvas"></canvas>

    <main class="hero-content">
      <div class="eyebrow-pill">
        <span class="dot"></span>
        <span>NEXT-GEN AI DATA INTELLIGENCE // ZERO API COST // 100% PRIVATE</span>
      </div>

      <h1 class="mega-headline">
        <span class="text-grad-1">DATA CRUNCHING.</span>
        <span class="text-grad-2">ZERO FLUFF.</span>
        <span class="text-grad-3">HITS DIFFERENT 🔥</span>
      </h1>

      <p class="hero-subtext">
        Chat with your datasets in plain English, sniff out messy data, generate aesthetic dark-mode charts, and spin up executive PowerPoint decks in 3 seconds flat.
      </p>

      <div class="feature-grid">
        <div class="feature-card" style="--card-color: #06B6D4; --card-glow: rgba(6, 182, 212, 0.3);">
          <div class="card-icon">💬</div>
          <div class="card-title">Chat with Data</div>
          <div class="card-desc">Deterministic pandas math. Instant answers with zero hallucinations.</div>
        </div>

        <div class="feature-card" style="--card-color: #EC4899; --card-glow: rgba(236, 72, 153, 0.3);">
          <div class="card-icon">📊</div>
          <div class="card-title">Aesthetic Charts</div>
          <div class="card-desc">Bar, line, scatter & box plots styled to pure cyber perfection.</div>
        </div>

        <div class="feature-card" style="--card-color: #8B5CF6; --card-glow: rgba(139, 92, 246, 0.3);">
          <div class="card-icon">📑</div>
          <div class="card-title">1-Click Slide Deck</div>
          <div class="card-desc">Auto-generate 6-slide PowerPoint presentations ready for executives.</div>
        </div>

        <div class="feature-card" style="--card-color: #10B981; --card-glow: rgba(16, 185, 129, 0.3);">
          <div class="card-icon">🛡️</div>
          <div class="card-title">Dirty-Data Radar</div>
          <div class="card-desc">Identify missing values, ghost duplicate records & dirty schemas.</div>
        </div>
      </div>

      <div class="status-bar">
        <div class="status-item"><span style="color:#10B981">●</span> ENGINE: PANDAS 2.x</div>
        <div class="status-item"><span style="color:#06B6D4">⚡</span> SPEED: &lt;10ms LOCAL</div>
        <div class="status-item"><span style="color:#EC4899">🔒</span> PRIVACY: 100% OFFLINE</div>
        <div class="status-item"><span style="color:#F59E0B">💸</span> COST: $0.00 FOREVER</div>
      </div>
    </main>
  </div>

  <script>
    // Lightweight Interactive Constellation Particles
    (() => {
      const canvas = document.getElementById('particle-canvas');
      const ctx = canvas.getContext('2d');
      let w = canvas.width = window.innerWidth;
      let h = canvas.height = window.innerHeight;

      const particles = [];
      const count = Math.min(Math.floor((w * h) / 18000), 75);
      const colors = ['#8B5CF6', '#EC4899', '#06B6D4', '#10B981'];

      for (let i = 0; i < count; i++) {
        particles.push({
          x: Math.random() * w,
          y: Math.random() * h,
          vx: (Math.random() - 0.5) * 0.7,
          vy: (Math.random() - 0.5) * 0.7,
          radius: Math.random() * 2 + 1,
          color: colors[Math.floor(Math.random() * colors.length)]
        });
      }

      const animate = () => {
        ctx.clearRect(0, 0, w, h);
        for (let i = 0; i < particles.length; i++) {
          const p = particles[i];
          p.x += p.vx;
          p.y += p.vy;
          if (p.x < 0 || p.x > w) p.vx *= -1;
          if (p.y < 0 || p.y > h) p.vy *= -1;

          ctx.beginPath();
          ctx.arc(p.x, p.y, p.radius, 0, Math.PI * 2);
          ctx.fillStyle = p.color;
          ctx.globalAlpha = 0.5;
          ctx.fill();

          for (let j = i + 1; j < particles.length; j++) {
            const p2 = particles[j];
            const dist = Math.hypot(p.x - p2.x, p.y - p2.y);
            if (dist < 110) {
              ctx.beginPath();
              ctx.moveTo(p.x, p.y);
              ctx.lineTo(p2.x, p2.y);
              ctx.strokeStyle = '#ffffff';
              ctx.globalAlpha = (1 - dist / 110) * 0.12;
              ctx.lineWidth = 0.8;
              ctx.stroke();
            }
          }
        }
        ctx.globalAlpha = 1.0;
        requestAnimationFrame(animate);
      };

      animate();

      window.addEventListener('resize', () => {
        w = canvas.width = window.innerWidth;
        h = canvas.height = window.innerHeight;
      });
    })();
  </script>
</body>
</html>
'''

# Render Hero Banner with sleek height
components.html(HERO_HTML, height=620, scrolling=False)

# ------------------------- Workspace Heading -------------------------
st.markdown("""
<div class="dp-workspace-heading">
  <div>
    <div class="dp-kicker">⚡ LIVE DATA WORKSPACE</div>
    <h2>Turn raw tables into high-key insights.</h2>
    <p>Upload a file, ask questions, inspect data hygiene, and export executive presentations with zero friction.</p>
  </div>
  <div style="display:flex; gap:8px; align-items:center;">
    <span class="dp-tag-pill">🔥 100% Free Engine</span>
    <span class="dp-tag-pill">🛡️ Zero Cloud Tracking</span>
    <span class="dp-tag-pill">🚀 Ready to Crunch</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ------------------------- Load data -------------------------
@st.cache_data(show_spinner=False)
def read_csv_bytes(raw: bytes, filename: str):
    errors = []
    for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin1"):
        try:
            return pd.read_csv(BytesIO(raw), encoding=encoding, sep=None, engine="python")
        except Exception as exc:
            errors.append(f"{encoding}: {exc}")
    raise ValueError("Could not read this CSV. " + " | ".join(errors[-2:]))

@st.cache_data(show_spinner=False)
def read_excel_bytes(raw: bytes):
    return pd.read_excel(BytesIO(raw))

# Sidebar Gen-Z Branding Box
st.sidebar.markdown("""
<div class="dp-sidebar-hero">
  <div class="dp-sidebar-badge">
    <span class="pulse-dot"></span>
    ONLINE // LOCAL
  </div>
  <div class="dp-sidebar-title">DATAPILOT // PRO</div>
  <div class="dp-sidebar-sub">Next-gen analytical engine powered by deterministic pandas calculations.</div>
</div>
""", unsafe_allow_html=True)

uploaded = st.sidebar.file_uploader("📂 Drop dataset here (CSV, XLSX, XLS)", type=["csv", "xlsx", "xls"])
if uploaded:
    try:
        raw = uploaded.getvalue()
        if uploaded.name.lower().endswith((".xlsx", ".xls")):
            df = read_excel_bytes(raw)
        else:
            df = read_csv_bytes(raw, uploaded.name)
        st.sidebar.success(f"✨ Loaded {len(df):,} rows × {len(df.columns)} columns")
    except Exception as exc:
        st.error(f"Could not load the file: {exc}")
        st.stop()
else:
    df = pd.DataFrame({
        "Product": ["Laptop", "Phone", "Laptop", "Tablet", "Phone", "Tablet", "Laptop", "Phone"],
        "Sales": [50000, 20000, 45000, 15000, 25000, 12000, 52000, 23000],
        "Quantity": [5, 10, 4, 3, 12, 2, 6, 8],
        "Region": ["Mumbai", "Pune", "Mumbai", "Nashik", "Pune", "Mumbai", "Nashik", "Mumbai"]
    })
    st.sidebar.info("💡 Sample Sales data is active. Drop your own CSV/XLSX to crunch live numbers.")

# Inspect column types
all_cols = list(df.columns)
numeric_cols = df.select_dtypes(include="number").columns.tolist()
categorical_cols = df.select_dtypes(include=["object", "category", "bool"]).columns.tolist()
date_cols = []
for col in df.columns:
    if col not in numeric_cols:
        parsed = pd.to_datetime(df[col], errors="coerce")
        if len(df) and parsed.notna().mean() >= 0.8:
            date_cols.append(col)

# ------------------------- Analysis helpers -------------------------
def choose_column_from_question(question, candidates):
    q = question.lower()
    for col in sorted(candidates, key=lambda x: len(str(x)), reverse=True):
        name = str(col).lower().strip()
        if name and name in q:
            return col
    synonyms = {
        "sales": ["sales", "revenue", "turnover", "amount", "profit", "price", "cost"],
        "product": ["product", "item", "category"],
        "region": ["region", "market", "state", "city", "country", "department"],
        "date": ["date", "time", "month", "year"]
    }
    for col in candidates:
        low = str(col).lower()
        for key, words in synonyms.items():
            if any(w in low for w in words) and any(w in q for w in words):
                return col
    return None

def resolve_aggregation(question):
    q = question.lower()
    if any(w in q for w in ["average", "avg", "mean", "on average"]):
        return "mean"
    if any(w in q for w in ["median"]):
        return "median"
    if any(w in q for w in ["count", "number of", "how many"]):
        return "count"
    if any(w in q for w in ["minimum", "lowest", "smallest", "min "]):
        return "min"
    if any(w in q for w in ["maximum", "highest", "largest", "max "]):
        return "max"
    if any(w in q for w in ["profit", "total", "sum", "revenue", "sales"]):
        return "sum"
    return "mean"

def human_agg(agg):
    return {"mean": "average", "median": "median", "count": "count", "min": "minimum", "max": "maximum", "sum": "total"}[agg]

def make_grouped_result(data, group_col, value_col, agg):
    if agg == "count":
        result = data.groupby(group_col, dropna=False)[value_col].count()
    else:
        result = data.groupby(group_col, dropna=False)[value_col].agg(agg)
    result = result.sort_values(ascending=False)
    return result

def detect_chart_type(q):
    if any(w in q for w in ["histogram", "distribution", "frequency"]):
        return "histogram"
    if "scatter" in q or "relationship" in q or "correlation" in q:
        return "scatter"
    if "line chart" in q or "trend" in q or "over time" in q:
        return "line"
    if "pie" in q:
        return "pie"
    if "box plot" in q or "boxplot" in q:
        return "box"
    return "bar"

def answer_question(data, question):
    q = question.lower().strip()
    if not q:
        return "Type a question about your dataset."

    if any(w in q for w in ["missing", "null", "empty", "data quality"]):
        miss = data.isna().sum()
        miss = miss[miss > 0].sort_values(ascending=False)
        if miss.empty:
            return "✨ **Clean bill of health!** Zero missing values were found."
        st.dataframe(miss.rename("Missing values").reset_index(names="Column"), use_container_width=True)
        return f"⚠️ Found missing values in **{len(miss)} column(s)**, affecting **{int(miss.sum())} cells**."

    if "duplicate" in q:
        dups = int(data.duplicated().sum())
        if dups == 0:
            return "✨ **No duplicate rows detected!** All records are unique."
        return f"⚠️ Duplicate rows detected: **{dups:,}**"

    if any(w in q for w in ["summary", "overview", "describe dataset", "dataset info", "profile"]):
        return (f"📊 **Dataset Profile:**\n\n"
                f"- **Rows:** `{len(data):,}` | **Columns:** `{len(data.columns)}`\n"
                f"- **Missing cells:** `{int(data.isna().sum().sum()):,}`\n"
                f"- **Duplicate rows:** `{int(data.duplicated().sum()):,}`\n"
                f"- **Numeric features:** `{', '.join(map(str, numeric_cols)) or 'None'}`\n"
                f"- **Categorical features:** `{', '.join(map(str, categorical_cols)) or 'None'}`")

    # Chart requests
    chart_request = any(w in q for w in ["chart", "plot", "graph", "visualize", "visualise", "bar", "histogram", "scatter", "pie", "box plot", "boxplot", "trend"])
    group_col = choose_column_from_question(question, categorical_cols + date_cols)
    value_col = choose_column_from_question(question, numeric_cols)

    if chart_request:
        chart_type = detect_chart_type(q)
        if chart_type == "histogram":
            col = value_col or (numeric_cols[0] if numeric_cols else None)
            if col is None:
                return "⚠️ A histogram needs a numeric column. None detected in this dataset."
            fig, ax = plt.subplots(figsize=(9, 4.5))
            data[col].dropna().plot(kind="hist", bins=20, ax=ax, color='#00F0FF', edgecolor="#07090e", alpha=0.9)
            ax.set_title(f"Distribution of {col}", fontsize=13, fontweight='bold', pad=12)
            ax.set_xlabel(str(col), fontweight='bold')
            ax.set_ylabel("Frequency", fontweight='bold')
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
            return f"✨ Rendered high-vibe histogram for **{col}**."

        if chart_type == "scatter":
            if len(numeric_cols) < 2:
                return "⚠️ Scatter plots require at least 2 numerical columns."
            x_col = value_col or numeric_cols[0]
            y_candidates = [c for c in numeric_cols if c != x_col]
            y_col = choose_column_from_question(question, y_candidates) or y_candidates[0]
            fig, ax = plt.subplots(figsize=(9, 4.5))
            ax.scatter(data[x_col], data[y_col], color='#EC4899', edgecolors='#FFFFFF', s=60, alpha=0.8)
            ax.set_xlabel(str(x_col), fontweight='bold')
            ax.set_ylabel(str(y_col), fontweight='bold')
            ax.set_title(f"{y_col} vs {x_col}", fontsize=13, fontweight='bold', pad=12)
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
            return f"✨ Rendered scatter plot of **{y_col} vs {x_col}**."

        if group_col is None and categorical_cols:
            group_col = categorical_cols[0]
        if value_col is None and numeric_cols:
            value_col = numeric_cols[0]
        if group_col is None or value_col is None:
            return "⚠️ For a grouped chart, select both a category and numeric column in the Explore tab."

        agg = resolve_aggregation(q)
        grouped = make_grouped_result(data, group_col, value_col, agg)
        if chart_type == "pie":
            grouped = grouped.head(8)
            fig, ax = plt.subplots(figsize=(7, 6))
            grouped.plot(kind="pie", ax=ax, autopct="%1.1f%%", colors=NEON_COLORS[:len(grouped)], textprops={'color': '#FFFFFF', 'fontweight': 'bold'})
            ax.set_ylabel("")
            ax.set_title(f"{human_agg(agg).title()} {value_col} by {group_col}", fontsize=13, fontweight='bold', pad=12)
        else:
            kind = "line" if chart_type == "line" else "bar"
            fig, ax = plt.subplots(figsize=(10, 4.8))
            if kind == "line":
                grouped.plot(kind=kind, ax=ax, marker="o", color='#00F0FF', linewidth=2.5, markersize=7)
            else:
                grouped.plot(kind=kind, ax=ax, color='#8B5CF6', edgecolor='#07090e')
            ax.set_title(f"{human_agg(agg).title()} {value_col} by {group_col}", fontsize=13, fontweight='bold', pad=12)
            ax.set_xlabel(str(group_col), fontweight='bold')
            ax.set_ylabel(f"{human_agg(agg).title()} {value_col}", fontweight='bold')
            ax.tick_params(axis="x", rotation=35)
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
        st.dataframe(grouped.rename(f"{human_agg(agg).title()} {value_col}").reset_index(), use_container_width=True)
        return f"✨ Rendered a **{chart_type} chart** displaying {human_agg(agg)} **{value_col}** by **{group_col}**."

    # Grouped analytics
    group_terms = ["by product", "per product", "for each product", "by category", "per category",
                   "by region", "per region", "by city", "by country", "by state", "by department",
                   "group by", "for each", "per "]
    asks_grouping = any(term in q for term in group_terms)
    if asks_grouping or ("average" in q or "avg" in q or "mean" in q) or any(w in q for w in ["highest", "lowest", "maximum", "minimum", "top"]):
        if value_col is None and numeric_cols:
            value_col = numeric_cols[0]
        if group_col is None and categorical_cols and asks_grouping:
            group_col = categorical_cols[0]

        if group_col is not None and value_col is not None and asks_grouping:
            agg = resolve_aggregation(q)
            grouped = make_grouped_result(data, group_col, value_col, agg)
            st.dataframe(grouped.rename(f"{human_agg(agg).title()} {value_col}").reset_index(), use_container_width=True)
            if "highest" in q or "top" in q:
                return f"🔥 **Top {human_agg(agg)} {value_col}:** **{grouped.index[0]}** (`{grouped.iloc[0]:,.2f}`)."
            if "lowest" in q:
                return f"📉 **Lowest {human_agg(agg)} {value_col}:** **{grouped.index[-1]}** (`{grouped.iloc[-1]:,.2f}`)."
            return f"Calculated {human_agg(agg)} **{value_col}** for each **{group_col}**."

    if value_col is not None and any(w in q for w in ["average", "avg", "mean", "median", "maximum", "minimum", "highest", "lowest", "sum", "total"]):
        series = data[value_col].dropna()
        if series.empty:
            return f"Column **{value_col}** contains no numeric values."
        if "median" in q:
            value, label = series.median(), "Median"
        elif any(w in q for w in ["maximum", "highest", "max"]):
            value, label = series.max(), "Maximum"
        elif any(w in q for w in ["minimum", "lowest", "min"]):
            value, label = series.min(), "Minimum"
        elif any(w in q for w in ["sum", "total"]):
            value, label = series.sum(), "Total"
        else:
            value, label = series.mean(), "Average"
        return f"🎯 {label} **{value_col}**: **`{value:,.2f}`**"

    return ("🤔 I couldn't map that directly. Try asking something like:\n"
            "- `Show average Sales by Product as a bar chart`\n"
            "- `What is the highest Sales?`\n"
            "- `Compare total Sales by Region`\n"
            "- `Create a histogram of Quantity`\n"
            "- `Find missing values`\n"
            "- `Show dataset summary`")

# ------------------------- Storytelling + PowerPoint -------------------------
def make_story(data):
    rows, cols = data.shape
    missing_total = int(data.isna().sum().sum())
    duplicate_total = int(data.duplicated().sum())
    numeric = data.select_dtypes(include="number").columns.tolist()
    categorical = data.select_dtypes(include=["object", "category", "bool"]).columns.tolist()

    def find_col(candidates):
        lowered = {str(c).strip().lower(): c for c in data.columns}
        for candidate in candidates:
            if candidate in lowered:
                return lowered[candidate]
        for col in data.columns:
            name = str(col).strip().lower()
            if any(candidate in name for candidate in candidates):
                return col
        return None

    measure = find_col(["sales", "revenue", "amount", "profit", "quantity", "price", "total"])
    if measure not in numeric:
        measure = numeric[0] if numeric else None
    group = find_col(["product", "category", "region", "market", "department", "segment", "city", "country", "state"])
    if group not in categorical:
        group = categorical[0] if categorical else None

    findings = []
    chart_info = None
    if measure is not None:
        values = pd.to_numeric(data[measure], errors="coerce").dropna()
        if not values.empty:
            findings.append(f"Average {measure}: {values.mean():,.2f}; median: {values.median():,.2f}.")
            findings.append(f"Lowest {measure}: {values.min():,.2f}; highest: {values.max():,.2f}.")
            if values.nunique() > 1:
                findings.append(f"The spread range of {measure} is {values.max() - values.min():,.2f}.")

            if group is not None:
                try:
                    grouped = data.groupby(group, dropna=False)[measure].mean().dropna().sort_values(ascending=False)
                    if not grouped.empty:
                        findings.append(f"Highest average {measure} by {group}: {grouped.index[0]} ({grouped.iloc[0]:,.2f}).")
                        findings.append(f"Lowest average {measure} by {group}: {grouped.index[-1]} ({grouped.iloc[-1]:,.2f}).")
                        chart_info = (group, measure, grouped.head(10))
                except (TypeError, ValueError, KeyError):
                    pass

    if not findings:
        findings.append("No usable numerical columns detected. Review the column types and category counts.")

    recommendations = []
    if missing_total:
        recommendations.append("Investigate missing values before making executive decisions; decide whether to impute or drop.")
    else:
        recommendations.append("Zero missing cells detected; confirm the dataset represents the complete business window.")
    if duplicate_total:
        recommendations.append("Audit duplicate records to differentiate repeated events from pipeline ingestion bugs.")
    else:
        recommendations.append("Zero exact duplicate records detected across all features.")
    if group is not None and measure is not None:
        recommendations.append(f"Deep-dive into the performance gap between the top and bottom performers by {group}.")
    if numeric:
        recommendations.append("Run outlier analysis and confirm measurement units before presenting final takeaways.")
    recommendations.append("Treat these insights as exploratory correlations; cross-validate with real-world business context.")

    return {
        "rows": rows, "cols": cols, "missing": missing_total, "duplicates": duplicate_total,
        "numeric_cols": numeric, "categorical_cols": categorical,
        "headline": f"Dataset encompasses {rows:,} records across {cols:,} analytical features.",
        "quality": f"{missing_total:,} missing cells and {duplicate_total:,} duplicate records identified.",
        "findings": findings, "recommendations": recommendations,
        "chart_info": chart_info, "measure": measure, "group": group,
    }

def add_text_slide(prs, title, bullets, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.35), Inches(12), Inches(0.7))
    title_box.text_frame.text = title
    title_p = title_box.text_frame.paragraphs[0]
    title_p.font.size = Pt(28)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(15, 23, 42)
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.12), Inches(11.8), Inches(0.45))
        sub.text_frame.text = subtitle
        sub.text_frame.paragraphs[0].font.size = Pt(13)
        sub.text_frame.paragraphs[0].font.color.rgb = RGBColor(100, 116, 139)
    top = 1.85 if subtitle else 1.5
    for i, bullet in enumerate(bullets[:5]):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(top + i * 0.9),
            Inches(11.6), Inches(0.68)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(241, 245, 249)
        shape.line.color.rgb = RGBColor(203, 213, 225)
        frame = shape.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.margin_left = Inches(0.18)
        frame.margin_right = Inches(0.12)
        frame.margin_top = Inches(0.08)
        paragraph = frame.paragraphs[0]
        paragraph.text = str(bullet)
        paragraph.font.size = Pt(15)
        paragraph.font.color.rgb = RGBColor(30, 41, 59)
        paragraph.alignment = PP_ALIGN.LEFT
    return slide

def chart_png(chart_info):
    group, measure, grouped = chart_info
    fig, ax = plt.subplots(figsize=(9, 4.5))
    grouped.sort_values().plot(kind="barh", ax=ax, color='#8B5CF6', edgecolor='#0d1117')
    ax.set_title(f"Average {measure} by {group} (top groups)", fontsize=13, fontweight='bold', pad=12)
    ax.set_xlabel(f"Average {measure}", fontweight='bold')
    ax.set_ylabel(str(group), fontweight='bold')
    fig.tight_layout()
    out = BytesIO()
    fig.savefig(out, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    out.seek(0)
    return out

def create_presentation(story, dataset_name):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    add_text_slide(
        prs, "Data Story: Key Findings",
        [f"Dataset: {dataset_name}", story["headline"], f"Data quality check: {story['quality']}"],
        "Automatically generated by DataPilot AI Engine"
    )
    add_text_slide(
        prs, "1. Dataset Foundation",
        [f"{story['rows']:,} rows and {story['cols']:,} columns loaded.",
         f"Numeric features detected: {len(story['numeric_cols'])}.",
         f"Categorical features detected: {len(story['categorical_cols'])}.",
         f"Primary metric selected: {story['measure'] or 'None detected'}."],
        "High-level structural breakdown of the dataset"
    )
    add_text_slide(prs, "2. Key Metrics & Patterns", story["findings"][:5], "Descriptive statistics across core dimensions")
    if story["chart_info"]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.35), Inches(12), Inches(0.7))
        title_box.text_frame.text = "3. Segment Comparison"
        p = title_box.text_frame.paragraphs[0]
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)
        slide.shapes.add_picture(chart_png(story["chart_info"]), Inches(0.9), Inches(1.3), width=Inches(11.5))
    add_text_slide(
        prs, "4. Data Hygiene Audit",
        [f"Missing cells: {story['missing']:,}.",
         f"Exact duplicate rows: {story['duplicates']:,}.",
         "Missing or duplicated records can heavily distort statistical aggregation.",
         "Validate column definitions, units, and ranges before finalizing business conclusions."],
        "Confidence check on dataset reliability"
    )
    add_text_slide(prs, "5. Recommended Action Items", story["recommendations"][:5], "Actionable next steps based on statistical findings")
    add_text_slide(
        prs, "Conclusion",
        ["Use these findings to guide targeted business experiments.",
         "Cross-reference anomalies with domain subject-matter experts.",
         "Correlations highlight opportunities but do not demonstrate causality."],
        "DataPilot Pro // Gen-Z Analytics Engine"
    )
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ------------------------- Multi-Tab Modern Workspace -------------------------
tab_overview, tab_chat, tab_explore, tab_quality, tab_story, tab_export = st.tabs([
    "⚡ Overview",
    "💬 Chat with Data",
    "📊 Explore & Chart",
    "🛡️ Data Quality",
    "📑 Story & PowerPoint",
    "🚀 Export Data"
])

# TAB 1: OVERVIEW
with tab_overview:
    st.markdown("### 📊 Dataset High-Level Diagnostics")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows Loaded", f"{len(df):,}")
    c2.metric("Columns Total", f"{len(df.columns):,}")
    c3.metric("Missing Cells", f"{int(df.isna().sum().sum()):,}")
    c4.metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 🔍 First 20 Records Preview")
    st.dataframe(df.head(20), use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns(2)
    with left:
        st.markdown("#### 🧬 Schema & Types Breakdown")
        type_df = pd.DataFrame({
            "Column": df.columns.astype(str),
            "Type": [str(t) for t in df.dtypes],
            "Missing": [int(df[c].isna().sum()) for c in df.columns],
            "Unique Values": [int(df[c].nunique(dropna=True)) for c in df.columns],
        })
        st.dataframe(type_df, use_container_width=True, hide_index=True)
    with right:
        st.markdown("#### 📈 Numerical Statistics Profile")
        if numeric_cols:
            st.dataframe(df[numeric_cols].describe().T, use_container_width=True)
        else:
            st.info("No numerical columns detected in this dataset.")

# TAB 2: CHAT WITH DATA
with tab_chat:
    st.markdown("### 💬 Conversational Analytics")
    st.caption("Ask anything in natural language. Powered by deterministic pandas operations with zero cloud tokens.")

    # Funky Quick Chips
    st.markdown("""
    <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:1rem;">
      <span class="dp-tag-pill">💡 Try asking:</span>
      <span class="dp-tag-pill">"Average Sales by Product"</span>
      <span class="dp-tag-pill">"Bar chart of Sales by Region"</span>
      <span class="dp-tag-pill">"What is the highest Sales?"</span>
      <span class="dp-tag-pill">"Check data quality"</span>
    </div>
    """, unsafe_allow_html=True)

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = [{
            "role": "assistant",
            "content": "👋 Hey! I'm your DataPilot co-pilot. Drop a question like **'Show average Sales by Product as a bar chart'** or **'What is the highest Sales?'** to get started."
        }]

    for item in st.session_state.chat_history:
        with st.chat_message(item["role"]):
            st.markdown(item["content"])

    question = st.chat_input("Ask a question about your dataset...")
    if question:
        st.session_state.chat_history.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        with st.chat_message("assistant"):
            with st.spinner("Crunching numbers..."):
                try:
                    answer = answer_question(df, question)
                except Exception as exc:
                    answer = f"⚠️ Analysis failed: {exc}"
            st.markdown(answer)
        st.session_state.chat_history.append({"role": "assistant", "content": answer})

# TAB 3: EXPLORE & CHART
with tab_explore:
    st.markdown("### 🎨 Manual Chart Studio")
    st.caption("Customize your visual story with full control over dimensions, measures, and aggregations.")

    if not numeric_cols:
        st.warning("⚠️ No numerical columns detected for charting.")
    else:
        with st.container():
            x1, x2, x3 = st.columns(3)
            with x1:
                chart_kind = st.selectbox("Chart Style", ["Bar", "Line", "Histogram", "Scatter", "Box"])
            with x2:
                measure = st.selectbox("Numeric Column (Metric)", numeric_cols)
            with x3:
                category_options = categorical_cols + date_cols
                category = st.selectbox("Group / Category Column", ["(none)"] + category_options)

            aggregation = st.selectbox("Aggregation Math", ["Average (mean)", "Total (sum)", "Count", "Median", "Maximum", "Minimum"])
            agg_map = {"Average (mean)": "mean", "Total (sum)": "sum", "Count": "count", "Median": "median", "Maximum": "max", "Minimum": "min"}

            if st.button("🚀 Render Visual", type="primary"):
                try:
                    fig, ax = plt.subplots(figsize=(10, 4.8))
                    if chart_kind == "Histogram":
                        df[measure].dropna().plot(kind="hist", bins=20, ax=ax, color='#00F0FF', edgecolor='#07090e')
                        ax.set_xlabel(str(measure), fontweight='bold')
                        ax.set_ylabel("Frequency", fontweight='bold')
                    elif chart_kind == "Scatter":
                        if len(numeric_cols) < 2:
                            st.error("Scatter plot requires at least two numerical columns.")
                            st.stop()
                        y_candidates = [c for c in numeric_cols if c != measure]
                        y_measure = y_candidates[0]
                        ax.scatter(df[measure], df[y_measure], color='#EC4899', edgecolors='#FFFFFF', s=60, alpha=0.8)
                        ax.set_xlabel(str(measure), fontweight='bold')
                        ax.set_ylabel(str(y_measure), fontweight='bold')
                    elif chart_kind == "Box":
                        if category != "(none)":
                            df.boxplot(column=measure, by=category, ax=ax, patch_artist=True,
                                       boxprops=dict(facecolor="#8B5CF6", color="#FFFFFF"),
                                       medianprops=dict(color="#00FFA3", linewidth=2))
                        else:
                            df[measure].plot(kind="box", ax=ax, patch_artist=True,
                                             boxprops=dict(facecolor="#8B5CF6", color="#FFFFFF"),
                                             medianprops=dict(color="#00FFA3", linewidth=2))
                    elif category == "(none)":
                        st.error("Please pick a group/category column for bar or line charts.")
                        st.stop()
                    else:
                        grouped = df.groupby(category, dropna=False)[measure].agg(agg_map[aggregation]).sort_values(ascending=False)
                        if chart_kind == "Line":
                            grouped.plot(kind="line", ax=ax, marker="o", color='#00F0FF', linewidth=2.5, markersize=7)
                        else:
                            grouped.plot(kind=chart_kind.lower(), ax=ax, color='#8B5CF6', edgecolor='#07090e')
                        ax.set_xlabel(str(category), fontweight='bold')
                        ax.set_ylabel(f"{aggregation} {measure}", fontweight='bold')
                        ax.tick_params(axis="x", rotation=35)
                        st.dataframe(grouped.rename("Result").reset_index(), use_container_width=True)
                    ax.set_title(f"{chart_kind} Analysis", fontsize=13, fontweight='bold', pad=12)
                    fig.tight_layout()
                    st.pyplot(fig)
                    plt.close(fig)
                except Exception as exc:
                    st.error(f"Could not render visual: {exc}")

# TAB 4: DATA QUALITY
with tab_quality:
    st.markdown("### 🛡️ Data Hygiene & Integrity Audit")
    st.caption("Catch anomalies, missing fields, and duplicates before they corrupt downstream modeling.")

    q1, q2 = st.columns([3, 1])
    with q1:
        missing_df = pd.DataFrame({
            "Column": df.columns.astype(str),
            "Missing Count": [int(df[c].isna().sum()) for c in df.columns],
            "Missing %": [round(float(df[c].isna().mean() * 100), 2) for c in df.columns],
            "Unique Values": [int(df[c].nunique(dropna=True)) for c in df.columns],
            "Data Type": [str(df[c].dtype) for c in df.columns],
        })
        st.dataframe(missing_df, use_container_width=True, hide_index=True)
    with q2:
        st.metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")
        st.metric("Total Missing Cells", f"{int(df.isna().sum().sum()):,}")
        has_missing = df.isna().any(axis=1).sum() > 0
        if has_missing:
            missing_rows = df[df.isna().any(axis=1)]
            st.download_button(
                "⬇️ Download Dirty Rows",
                missing_rows.to_csv(index=False),
                "dirty_rows.csv",
                "text/csv"
            )
        else:
            st.success("✨ Zero dirty rows found!")

# TAB 5: STORY & POWERPOINT
with tab_story:
    st.markdown("### 📑 Automated Executive Story & Slide Deck")
    st.caption("Instantly synthesize takeaways, segment breakdowns, and compile a PowerPoint deck (.pptx).")

    story = make_story(df)
    sm1, sm2, sm3, sm4 = st.columns(4)
    sm1.metric("Dataset Rows", f"{story['rows']:,}")
    sm2.metric("Features", f"{story['cols']:,}")
    sm3.metric("Missing Cells", f"{story['missing']:,}")
    sm4.metric("Duplicate Rows", f"{story['duplicates']:,}")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"#### 💡 Core Takeaway: {story['headline']}")

    c_findings, c_recom = st.columns(2)
    with c_findings:
        st.markdown("##### 📌 Key Data Patterns")
        for finding in story["findings"]:
            st.write(f"• {finding}")
    with c_recom:
        st.markdown("##### 🎯 Recommended Next Steps")
        for recommendation in story["recommendations"]:
            st.write(f"• {recommendation}")

    if story["chart_info"]:
        group_col, measure_col, grouped_vals = story["chart_info"]
        st.markdown(f"##### 📊 Segment Comparison: {measure_col} by {group_col}")
        fig, ax = plt.subplots(figsize=(10, 4.2))
        grouped_vals.sort_values().plot(kind="barh", ax=ax, color='#00F0FF', edgecolor='#07090e')
        ax.set_xlabel(f"Average {measure_col}", fontweight='bold')
        ax.set_ylabel(str(group_col), fontweight='bold')
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("✨ Compile Executive PowerPoint (.pptx)", type="primary", key="create_story_pptx"):
        try:
            file_name = uploaded.name if uploaded is not None else "Sample sales dataset"
            pptx_bytes = create_presentation(story, file_name).getvalue()
            st.session_state["story_pptx"] = pptx_bytes
            st.success("🎉 PowerPoint generated successfully! Click below to download.")
        except Exception as exc:
            st.error(f"Could not build slide deck: {exc}")

    if "story_pptx" in st.session_state:
        st.download_button(
            "⬇️ Download PowerPoint Deck (.pptx)",
            data=st.session_state["story_pptx"],
            file_name="DataPilot_Executive_Deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="download_story_pptx"
        )

# TAB 6: EXPORT
with tab_export:
    st.markdown("### 🚀 Export Engine")
    st.caption("Grab cleaned datasets, column profiles, or diagnostic reports in one click.")

    report = [
        "================================================",
        "          DATAPILOT ANALYTICS REPORT            ",
        "================================================",
        f"Rows: {len(df)}",
        f"Columns: {len(df.columns)}",
        f"Missing cells: {int(df.isna().sum().sum())}",
        f"Duplicate rows: {int(df.duplicated().sum())}",
        "\n--- COLUMN TYPES ---",
        df.dtypes.to_string(),
        "\n--- NUMERICAL PROFILE ---",
        df.describe(include="all").to_string(),
    ]
    report_text = "\n".join(report)

    e1, e2, e3 = st.columns(3)
    with e1:
        st.download_button("📄 Download Text Report (.txt)", report_text, "datapilot_report.txt", "text/plain", use_container_width=True)
    with e2:
        st.download_button("📊 Export Cleaned Data (.csv)", df.to_csv(index=False), "cleaned_dataset.csv", "text/csv", use_container_width=True)
    with e3:
        prof_df = pd.DataFrame({
            "column": df.columns.astype(str),
            "dtype": [str(x) for x in df.dtypes],
            "missing": [int(df[c].isna().sum()) for c in df.columns],
            "unique": [int(df[c].nunique(dropna=True)) for c in df.columns]
        })
        st.download_button("🧬 Export Column Profile (.csv)", prof_df.to_csv(index=False), "column_profile.csv", "text/csv", use_container_width=True)

# Sidebar Footer
st.sidebar.markdown("---")
st.sidebar.caption("⚡ **DataPilot Gen-Z Edition** • Zero API tokens • 100% Client-side Pandas calculations • Your data stays local.")
