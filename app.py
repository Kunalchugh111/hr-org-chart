"""
Nodely — Streamlit launcher.

Routes between the marketing landing page (default) and the org-chart builder
based on a query param. Clicking any "Try the app" button on the marketing
page sets ?view=app and reloads, which makes us serve nodely.html instead.
"""
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Nodely — Org charts that breathe",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Hide all of Streamlit's default chrome
st.markdown(
    """
    <style>
      #MainMenu, footer, header, .stDeployButton,
      [data-testid="stSidebar"], [data-testid="stToolbar"],
      [data-testid="stDecoration"], [data-testid="stHeader"]
        { display: none !important; }
      .stApp { background: #ffffff !important; }
      .block-container {
        padding: 0 !important;
        max-width: 100% !important;
        margin: 0 !important;
      }
      iframe { border: none !important; border-radius: 0 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

base = Path(__file__).parent
view = st.query_params.get("view", "marketing")

if view == "app":
    target = base / "nodely.html"
    if not target.exists():
        st.error(f"Couldn't find nodely.html at {target.resolve()}")
        st.stop()
    components.html(target.read_text(encoding="utf-8"), height=900, scrolling=False)
else:
    target = base / "marketing.html"
    if not target.exists():
        st.error(f"Couldn't find marketing.html at {target.resolve()}")
        st.stop()
    # Marketing is a long scrollable page — give it plenty of vertical room
    components.html(target.read_text(encoding="utf-8"), height=4800, scrolling=True)
