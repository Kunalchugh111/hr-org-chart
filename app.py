"""
Nodely — Streamlit launcher.

Routes between the marketing landing page (default) and the org-chart builder
based on a query param. Clicking any "Try the app" button on the marketing
page sets ?view=app via a programmatic <a target="_top"> click, which makes
us serve nodely.html on the next request.
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

# Hide Streamlit's chrome, then make the components iframe fill the viewport.
# The iframe height we pass to components.html() is a *floor* — this CSS
# stretches the iframe to 100vh so the marketing page looks edge-to-edge
# even though it's technically rendered inside a child document.
st.markdown(
    """
    <style>
      #MainMenu, footer, header, .stDeployButton,
      [data-testid="stSidebar"], [data-testid="stToolbar"],
      [data-testid="stDecoration"], [data-testid="stHeader"]
        { display: none !important; }
      .stApp { background: #ffffff !important; overflow: hidden !important; }
      .block-container {
        padding: 0 !important;
        max-width: 100% !important;
        margin: 0 !important;
      }
      [data-testid="stAppViewContainer"], .main, .main > .block-container {
        padding: 0 !important;
        margin: 0 !important;
        overflow: hidden !important;
      }
      /* Stretch the components.html iframe to fill the viewport.
         The marketing page is taller than the viewport, so it scrolls
         INSIDE the iframe — which is exactly what makes the nav anchor
         links (Features / How it works / Pricing) work correctly. */
      iframe {
        border: none !important;
        border-radius: 0 !important;
        height: 100vh !important;
        width: 100% !important;
        display: block !important;
      }
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
    # Chart app: viewport-height floor, CSS above stretches to 100vh.
    components.html(
        target.read_text(encoding="utf-8"),
        height=900,
        scrolling=False,
    )
else:
    target = base / "marketing.html"
    if not target.exists():
        st.error(f"Couldn't find marketing.html at {target.resolve()}")
        st.stop()
    # Marketing iframe: viewport-height floor + CSS 100vh stretch.
    # scrolling=True so the marketing page (which is taller than viewport)
    # scrolls INSIDE the iframe.  This is what makes the nav anchor links
    # + smooth-scroll-to-top work without cross-origin parent access.
    components.html(
        target.read_text(encoding="utf-8"),
        height=900,
        scrolling=True,
    )
