import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import json

# =========================================================
# 1) Page Configuration
# =========================================================
st.set_page_config(
    page_title="OrgDesign Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# 2) Streamlit Styling — Refined Dark Sidebar + Clean Main
# =========================================================
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
  }

  #MainMenu, footer, header, .stDeployButton { display: none !important; }

  .stApp { background: #0f1117 !important; }

  .block-container {
    padding-top: 0.5rem !important;
    padding-bottom: 2rem !important;
    max-width: 99% !important;
  }

  /* Sidebar dark theme */
  [data-testid="stSidebar"] {
    background: #161b27 !important;
    border-right: 1px solid #2a3147 !important;
  }
  [data-testid="stSidebar"] * {
    color: #e2e8f0 !important;
  }
  [data-testid="stSidebar"] .stMarkdown h3 {
    color: #ffffff !important;
    font-size: 1.1rem !important;
  }
  [data-testid="stSidebar"] .stMarkdown h5 {
    color: #94a3b8 !important;
    font-size: 0.72rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
  }

  /* File uploader */
  [data-testid="stFileUploader"] {
    background: #1e2536 !important;
    border: 1.5px dashed #3b4563 !important;
    border-radius: 12px !important;
  }

  /* Selectbox */
  [data-testid="stSelectbox"] > div > div {
    background: #1e2536 !important;
    border: 1px solid #2a3147 !important;
    border-radius: 10px !important;
    color: #e2e8f0 !important;
    font-weight: 600 !important;
  }

  /* Toggle */
  [data-testid="stToggle"] label { color: #e2e8f0 !important; }

  /* Slider */
  [data-testid="stSlider"] label { color: #94a3b8 !important; }

  /* Caption */
  .stApp .stCaption { color: #64748b !important; }

  /* Buttons */
  .stButton > button {
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-family: 'DM Sans', sans-serif !important;
    transition: all 0.2s !important;
    width: 100% !important;
  }
  .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #6366f1, #8b5cf6) !important;
    color: #ffffff !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(99,102,241,0.35) !important;
  }
  .stButton > button[kind="secondary"] {
    background: #1e2536 !important;
    color: #e2e8f0 !important;
    border: 1px solid #2a3147 !important;
  }

  /* Divider */
  hr { border-color: #2a3147 !important; }

  /* Info box */
  [data-testid="stInfo"] {
    background: #1e2536 !important;
    border: 1px solid #3b4563 !important;
    border-radius: 12px !important;
    color: #94a3b8 !important;
  }

  iframe { border: none !important; border-radius: 16px !important; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 3) Initialize State
# =========================================================
if "draft_moves" not in st.session_state:
    st.session_state.draft_moves = {}

# =========================================================
# 4) Sidebar
# =========================================================
with st.sidebar:
    st.markdown("### 🏢 OrgDesign Pro")
    st.markdown("---")

    st.markdown("##### 📂 Data Source")
    uploaded_file = st.file_uploader(
        "Upload HR Roster (CSV / XLSX)",
        type=["csv", "xlsx", "xls"],
        help="Required columns: Employee Code, Employee Name, L1 Manager Code"
    )

    with st.expander("📋 Required Columns"):
        st.markdown("""
| Column | Required |
|--------|----------|
| Employee Code | ✅ |
| Employee Name | ✅ |
| L1 Manager Code | ✅ |
| Designation | Optional |
| Grade | Optional |
| Sub Function | Optional |
        """)

    df = None
    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine="openpyxl")

            df.columns = df.columns.str.strip()

            if "Employee Code" in df.columns:
                df["Clean_Emp_Code"] = (
                    df["Employee Code"].astype(str)
                    .str.replace(".0", "", regex=False)
                    .str.strip()
                )
            else:
                st.error("❌ Missing required column: 'Employee Code'")

            if "L1 Manager Code" in df.columns:
                df["L1 Manager Code"] = (
                    df["L1 Manager Code"].astype(str)
                    .str.replace(".0", "", regex=False)
                    .str.strip()
                )

            total = len(df)
            managers = df["L1 Manager Code"].nunique() if "L1 Manager Code" in df.columns else 0
            st.success(f"✅ {total} employees loaded")
            st.caption(f"{managers} unique managers found")

        except Exception as e:
            st.error(f"Error reading file: {e}")

    st.markdown("---")
    st.markdown("##### ⚙️ Display Settings")

    if df is not None and "Sub Function" in df.columns:
        subs = ["All"] + sorted(df["Sub Function"].dropna().unique())
        selected_sub = st.selectbox("Filter by Sub Function", subs, index=0)
    else:
        selected_sub = "All"

    color_scheme = st.selectbox(
        "Card Accent Color",
        ["Indigo", "Emerald", "Rose", "Amber", "Sky"],
        index=0
    )

    chart_height = st.slider("Chart Height (px)", 600, 4000, 1800, 100)

    st.markdown("---")
    st.markdown("##### 💡 Tips")
    st.caption("🔍 Use the search bar to find any employee")
    st.caption("🖱️ Click cards to expand/collapse branches")
    st.caption("📌 Drag to pan, scroll to zoom")
    st.caption("🖼️ Save Image captures full chart as PNG")

# =========================================================
# 5) Main Logic
# =========================================================
if df is None:
    # Landing screen
    st.markdown("""
    <div style="
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        min-height: 75vh; text-align: center; padding: 2rem;
        font-family: 'DM Sans', sans-serif;
    ">
        <div style="font-size: 4rem; margin-bottom: 1rem;">🏢</div>
        <h1 style="
            font-size: 2.4rem; font-weight: 800;
            background: linear-gradient(135deg, #6366f1, #a78bfa);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            margin-bottom: 0.75rem;
        ">OrgDesign Pro</h1>
        <p style="color: #64748b; font-size: 1.05rem; max-width: 480px; line-height: 1.6; margin-bottom: 2rem;">
            Upload your HR roster CSV or Excel file in the sidebar to generate a fully interactive
            org chart with search, collapse, zoom and PNG export.
        </p>
        <div style="
            background: #1e2536; border: 1px solid #2a3147; border-radius: 16px;
            padding: 1.5rem 2rem; max-width: 420px; text-align: left;
        ">
            <p style="color: #94a3b8; font-size: 0.8rem; text-transform: uppercase;
                      letter-spacing: 0.08em; font-weight: 700; margin-bottom: 1rem;">
                Required columns
            </p>
            <div style="display: flex; flex-direction: column; gap: 8px;">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="background: #6366f1; color: #fff; border-radius: 6px;
                                 padding: 2px 8px; font-size: 0.72rem; font-weight: 700;">REQUIRED</span>
                    <span style="color: #e2e8f0; font-family: 'DM Mono', monospace; font-size: 0.85rem;">Employee Code</span>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="background: #6366f1; color: #fff; border-radius: 6px;
                                 padding: 2px 8px; font-size: 0.72rem; font-weight: 700;">REQUIRED</span>
                    <span style="color: #e2e8f0; font-family: 'DM Mono', monospace; font-size: 0.85rem;">Employee Name</span>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="background: #6366f1; color: #fff; border-radius: 6px;
                                 padding: 2px 8px; font-size: 0.72rem; font-weight: 700;">REQUIRED</span>
                    <span style="color: #e2e8f0; font-family: 'DM Mono', monospace; font-size: 0.85rem;">L1 Manager Code</span>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="background: #2a3147; color: #94a3b8; border-radius: 6px;
                                 padding: 2px 8px; font-size: 0.72rem; font-weight: 700;">OPTIONAL</span>
                    <span style="color: #64748b; font-family: 'DM Mono', monospace; font-size: 0.85rem;">Designation, Grade, Sub Function</span>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

required_cols = ["Employee Code", "Employee Name"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required column(s): {', '.join(missing)}")
    st.stop()

base_df = df.copy()

if "Clean_Emp_Code" in base_df.columns and "L1 Manager Code" in base_df.columns:
    for e_code, m_code in st.session_state.draft_moves.items():
        base_df.loc[base_df["Clean_Emp_Code"] == e_code, "L1 Manager Code"] = m_code

if selected_sub != "All" and "Sub Function" in base_df.columns:
    base_df = base_df[base_df["Sub Function"] == selected_sub]

valid_ids = set(base_df["Clean_Emp_Code"].dropna().astype(str).tolist())

view_data = []
for _, row in base_df.iterrows():
    eid = str(row.get("Clean_Emp_Code", "")).strip()
    mid = str(row.get("L1 Manager Code", "")).strip()
    if mid not in valid_ids:
        mid = ""
    view_data.append({
        "id": eid,
        "manager": mid,
        "name": str(row.get("Employee Name", "Unknown")),
        "title": str(row.get("Designation", "N/A")),
        "grade": str(row.get("Grade", "N/A")),
        "sub": str(row.get("Sub Function", "N/A"))
    })

# Color accent map
accent_colors = {
    "Indigo": "#6366f1",
    "Emerald": "#10b981",
    "Rose": "#f43f5e",
    "Amber": "#f59e0b",
    "Sky": "#0ea5e9",
}
accent = accent_colors.get(color_scheme, "#6366f1")

# =========================================================
# 6) HTML Component — Full rewrite
# =========================================================
html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>OrgDesign Pro</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<style>
  :root {{
    --bg:           #0f1117;
    --surface:      #161b27;
    --surface2:     #1e2536;
    --border:       #2a3147;
    --border-light: #3b4563;
    --text:         #f0f4ff;
    --text-muted:   #8896b3;
    --accent:       {accent};
    --accent-glow:  {accent}44;
    --connector:    #3b4563;
    --success:      #10b981;
    --warning:      #f59e0b;
    --danger:       #f43f5e;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  html, body {{
    height: 100%;
    background: var(--bg);
    font-family: 'DM Sans', sans-serif;
    color: var(--text);
    overflow: hidden;
  }}

  body {{
    display: flex;
    flex-direction: column;
    height: 100vh;
  }}

  /* ── Toolbar ─────────────────────────────── */
  .toolbar {{
    flex-shrink: 0;
    height: 58px;
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    padding: 0 16px;
    gap: 8px;
    z-index: 200;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
  }}

  .brand {{
    font-weight: 800;
    font-size: 1rem;
    letter-spacing: -0.02em;
    color: var(--text);
    display: flex;
    align-items: center;
    gap: 8px;
    margin-right: 4px;
  }}
  .brand-dot {{
    width: 8px; height: 8px;
    background: var(--accent);
    border-radius: 50%;
    box-shadow: 0 0 10px var(--accent);
  }}

  .divider {{ width: 1px; height: 28px; background: var(--border); margin: 0 4px; }}
  .spacer {{ flex: 1; }}

  /* Search */
  .search-wrap {{
    position: relative;
    display: flex;
    align-items: center;
    flex: 1;
    max-width: 280px;
  }}
  .search-icon {{
    position: absolute;
    left: 11px;
    font-size: 0.85rem;
    pointer-events: none;
    opacity: 0.5;
  }}
  #search-input {{
    width: 100%;
    background: var(--surface2);
    border: 1.5px solid var(--border);
    border-radius: 10px;
    padding: 7px 12px 7px 32px;
    font-size: 0.84rem;
    font-weight: 500;
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    outline: none;
    transition: border-color 0.2s, box-shadow 0.2s;
  }}
  #search-input:focus {{
    border-color: var(--accent);
    box-shadow: 0 0 0 3px var(--accent-glow);
  }}
  #search-input::placeholder {{ color: var(--text-muted); }}

  /* Search results dropdown */
  #search-results {{
    position: absolute;
    top: calc(100% + 6px);
    left: 0; right: 0;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    box-shadow: 0 12px 40px rgba(0,0,0,0.5);
    max-height: 280px;
    overflow-y: auto;
    z-index: 999;
    display: none;
  }}
  #search-results.visible {{ display: block; }}
  .sr-item {{
    padding: 10px 14px;
    cursor: pointer;
    border-bottom: 1px solid var(--border);
    transition: background 0.12s;
  }}
  .sr-item:last-child {{ border-bottom: none; }}
  .sr-item:hover {{ background: var(--surface2); }}
  .sr-name {{ font-weight: 700; font-size: 0.85rem; }}
  .sr-sub  {{ font-size: 0.75rem; color: var(--text-muted); margin-top: 2px; }}

  /* Zoom controls */
  .zoom-strip {{
    display: flex; align-items: center; gap: 3px;
    background: var(--surface2); border-radius: 10px;
    padding: 3px; border: 1.5px solid var(--border);
  }}

  /* Buttons */
  .btn {{
    background: var(--surface2);
    border: 1.5px solid var(--border);
    padding: 7px 13px;
    border-radius: 9px;
    cursor: pointer;
    font-weight: 700;
    font-size: 0.82rem;
    color: var(--text);
    transition: all 0.15s;
    display: inline-flex; align-items: center; gap: 6px;
    white-space: nowrap;
    user-select: none;
    font-family: 'DM Sans', sans-serif;
    line-height: 1;
  }}
  .btn:hover {{
    border-color: var(--accent);
    color: var(--accent);
    box-shadow: 0 0 0 2px var(--accent-glow);
  }}
  .btn.accent {{
    background: var(--accent);
    color: #fff; border: none;
    box-shadow: 0 4px 14px var(--accent-glow);
  }}
  .btn.accent:hover {{
    opacity: 0.9;
    box-shadow: 0 6px 20px var(--accent-glow);
    color: #fff;
  }}
  .btn.ghost {{ background: transparent; border: none; padding: 6px 10px; border-radius: 8px; }}
  .btn.ghost:hover {{ background: var(--surface2); color: var(--text); border: none; box-shadow: none; }}

  .zoom-label {{
    font-size: 0.76rem; color: var(--text); min-width: 46px;
    text-align: center; font-weight: 700;
    font-variant-numeric: tabular-nums;
    font-family: 'DM Mono', monospace;
  }}

  /* Stats bar */
  .stats-bar {{
    flex-shrink: 0;
    height: 38px;
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    display: flex; align-items: center;
    padding: 0 20px; gap: 20px;
    font-size: 0.75rem;
  }}
  .stat {{ display: flex; align-items: center; gap: 6px; }}
  .stat-val {{ font-weight: 700; color: var(--text); font-variant-numeric: tabular-nums; }}
  .stat-lbl {{ color: var(--text-muted); }}
  .stat-dot {{ width: 6px; height: 6px; border-radius: 50%; background: var(--accent); }}

  /* Canvas */
  .canvas-wrapper {{
    flex: 1;
    overflow: auto;
    -webkit-overflow-scrolling: touch;
    background:
      radial-gradient(circle at 20% 20%, rgba(99,102,241,0.04) 0%, transparent 50%),
      radial-gradient(circle at 80% 80%, rgba(99,102,241,0.03) 0%, transparent 50%),
      var(--bg);
    background-size: 100% 100%, 100% 100%, 20px 20px;
    position: relative;
    cursor: grab;
  }}
  .canvas-wrapper:active {{ cursor: grabbing; }}

  .canvas-wrapper::before {{
    content: '';
    position: fixed;
    inset: 0;
    background-image:
      linear-gradient(var(--border) 1px, transparent 1px),
      linear-gradient(90deg, var(--border) 1px, transparent 1px);
    background-size: 40px 40px;
    opacity: 0.2;
    pointer-events: none;
    z-index: 0;
  }}

  .canvas-content {{
    display: inline-block;
    padding: 60px 80px 120px 80px;
    transform-origin: top left;
    min-width: 100%;
    position: relative;
    z-index: 1;
  }}

  /* Tree structure */
  .tree {{ display: inline-block; }}
  .tree ul {{
    padding-top: 24px;
    position: relative;
    list-style: none;
    display: flex;
    justify-content: center;
    gap: 0;
  }}

  .tree li {{
    display: table-cell;
    vertical-align: top;
    text-align: center;
    position: relative;
    padding: 24px 8px 0 8px;
  }}

  /* Connector lines — fix orphan lines */
  .tree li::before, .tree li::after {{
    content: '';
    position: absolute;
    top: 0; right: 50%;
    border-top: 2px solid var(--connector);
    width: 50%; height: 24px;
  }}
  .tree li::after {{
    right: auto; left: 50%;
    border-left: 2px solid var(--connector);
  }}
  .tree li:only-child::before,
  .tree li:only-child::after {{ display: none; }}
  .tree li:first-child::before,
  .tree li:last-child::after  {{ display: none; }}
  .tree li:first-child::after {{ border-radius: 6px 0 0 0; }}
  .tree li:last-child::before {{ border-radius: 0 6px 0 0; }}

  .tree ul ul::before {{
    content: '';
    position: absolute;
    top: 0; left: 50%;
    border-left: 2px solid var(--connector);
    height: 24px;
  }}

  /* Node card */
  .node-card {{
    display: inline-block;
    width: 220px;
    background: var(--surface);
    border: 1.5px solid var(--border);
    border-radius: 14px;
    cursor: pointer;
    text-align: left;
    border-top: 3px solid var(--accent);
    transition: transform 0.15s, box-shadow 0.15s, border-color 0.15s;
    box-shadow: 0 2px 12px rgba(0,0,0,0.25);
    position: relative;
    font-family: 'DM Sans', sans-serif;
  }}

  .node-card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 8px 30px rgba(0,0,0,0.4), 0 0 0 2px var(--accent-glow);
    border-color: var(--accent);
    z-index: 10;
  }}

  .node-card.highlighted {{
    border-color: var(--warning) !important;
    border-top-color: var(--warning) !important;
    box-shadow: 0 0 0 3px rgba(245,158,11,0.3), 0 8px 30px rgba(0,0,0,0.4) !important;
  }}

  .node-card.collapsed-node {{
    border-top-color: var(--text-muted);
    opacity: 0.75;
  }}

  .card-header {{
    padding: 8px 11px 7px;
    background: rgba(255,255,255,0.03);
    border-bottom: 1px solid var(--border);
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-radius: 12px 12px 0 0;
    gap: 6px;
  }}

  .sub-tag {{
    font-size: 0.58rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    background: rgba(99,102,241,0.15);
    color: var(--accent);
    padding: 2px 7px;
    border-radius: 999px;
    max-width: 110px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    border: 1px solid rgba(99,102,241,0.2);
  }}

  .grade-tag {{
    font-size: 0.65rem;
    font-weight: 700;
    color: var(--text-muted);
    white-space: nowrap;
    font-family: 'DM Mono', monospace;
  }}

  .card-body {{ padding: 11px 12px 9px; }}

  .emp-name {{
    font-size: 0.88rem;
    font-weight: 800;
    color: var(--text);
    margin-bottom: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    letter-spacing: -0.01em;
  }}

  .emp-title {{
    font-size: 0.75rem;
    font-weight: 500;
    color: var(--text-muted);
    line-height: 1.4;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }}

  .card-footer {{
    padding: 7px 12px;
    border-top: 1px solid var(--border);
    font-size: 0.68rem;
    font-weight: 600;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-radius: 0 0 12px 12px;
    background: rgba(255,255,255,0.02);
    font-family: 'DM Mono', monospace;
  }}

  .emp-id {{ color: var(--text-muted); }}

  .reports-badge {{
    background: rgba(255,255,255,0.06);
    padding: 2px 7px;
    border-radius: 999px;
    color: var(--text-muted);
    font-size: 0.65rem;
    font-weight: 700;
  }}
  .reports-badge.has-reports {{
    background: rgba(99,102,241,0.15);
    color: var(--accent);
    border: 1px solid rgba(99,102,241,0.2);
  }}

  .collapse-btn {{
    position: absolute;
    bottom: -12px;
    left: 50%;
    transform: translateX(-50%);
    width: 22px; height: 22px;
    background: var(--surface);
    border: 1.5px solid var(--border);
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    cursor: pointer;
    font-size: 0.65rem;
    color: var(--text-muted);
    transition: all 0.15s;
    z-index: 5;
  }}
  .collapse-btn:hover {{
    background: var(--accent);
    border-color: var(--accent);
    color: #fff;
    box-shadow: 0 0 10px var(--accent-glow);
  }}

  /* Collapsed subtree */
  li.collapsed > ul {{ display: none; }}

  /* ── Export overlay ───────────────────────── */
  .export-overlay {{
    position: fixed; inset: 0; z-index: 9999;
    background: rgba(0,0,0,0.85);
    backdrop-filter: blur(8px);
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; gap: 16px; color: var(--text);
  }}
  .export-spinner {{
    width: 48px; height: 48px;
    border: 3px solid rgba(255,255,255,0.1);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin 0.7s linear infinite;
  }}
  @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
  .export-title {{ font-size: 1.1rem; font-weight: 700; }}
  .export-sub {{ font-size: 0.8rem; color: var(--text-muted); font-weight: 500; }}

  /* ── Export-specific overrides ─────────────
     These ensure text is readable in png output */
  .export-stage .node-card {{
    background: #ffffff !important;
    border-color: #dde3ec !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
  }}
  .export-stage .card-header {{
    background: #f8fafc !important;
    border-bottom-color: #dde3ec !important;
  }}
  .export-stage .emp-name  {{ color: #0b1220 !important; }}
  .export-stage .emp-title {{ color: #374151 !important; }}
  .export-stage .emp-id, .export-stage .grade-tag {{ color: #6b7280 !important; }}
  .export-stage .sub-tag   {{ color: #4f46e5 !important; background: #eef2ff !important; border-color: #c7d2fe !important; }}
  .export-stage .reports-badge {{ color: #6b7280 !important; background: #f3f4f6 !important; }}
  .export-stage .reports-badge.has-reports {{ color: #4f46e5 !important; background: #eef2ff !important; }}
  .export-stage .card-footer {{ background: #f8fafc !important; border-top-color: #dde3ec !important; }}
  .export-stage .collapse-btn {{ display: none !important; }}
  /* Force text visibility — critical fix */
  .export-stage * {{
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }}
</style>
</head>

<body>

<!-- ═══════════════════════════════════════════
     TOOLBAR
═══════════════════════════════════════════ -->
<div class="toolbar">
  <div class="brand">
    <div class="brand-dot"></div>
    OrgDesign Pro
  </div>

  <div class="divider"></div>

  <div class="search-wrap">
    <span class="search-icon">🔍</span>
    <input id="search-input" type="text" placeholder="Search by name or ID…" autocomplete="off"/>
    <div id="search-results"></div>
  </div>

  <div class="divider"></div>

  <div class="zoom-strip">
    <button class="btn ghost" onclick="zoomBy(-0.1)" title="Zoom out">−</button>
    <span class="zoom-label" id="zoom-level">100%</span>
    <button class="btn ghost" onclick="zoomBy(0.1)" title="Zoom in">+</button>
    <button class="btn ghost" onclick="fitToScreen(true)" title="Fit to screen">⊡</button>
  </div>

  <button class="btn" onclick="centerView()" title="Center chart">🧭</button>
  <button class="btn" onclick="expandAll()" title="Expand all nodes">⊞ Expand</button>
  <button class="btn" onclick="collapseAll()" title="Collapse all">⊟ Collapse</button>

  <div class="divider"></div>

  <button class="btn" onclick="downloadCSV()">💾 CSV</button>
  <button class="btn accent" onclick="exportImage()">🖼️ Save PNG</button>
</div>

<!-- Stats bar -->
<div class="stats-bar" id="stats-bar">
  <div class="stat"><div class="stat-dot"></div><span class="stat-val" id="stat-total">—</span><span class="stat-lbl">employees</span></div>
  <div class="stat"><span class="stat-val" id="stat-roots">—</span><span class="stat-lbl">root nodes</span></div>
  <div class="stat"><span class="stat-val" id="stat-depth">—</span><span class="stat-lbl">max depth</span></div>
  <div class="stat"><span class="stat-val" id="stat-visible">—</span><span class="stat-lbl">currently shown</span></div>
</div>

<!-- Canvas -->
<div class="canvas-wrapper" id="canvas-wrapper">
  <div class="canvas-content" id="canvas-content">
    <div class="tree" id="org-tree"></div>
  </div>
</div>


<script>
const viewData = {json.dumps(view_data)};

let state = {{
  zoom: 1,
  exporting: false,
  highlighted: null
}};

let autoFitTimer = null;
let isPanning = false, panStartX, panStartY, panScrollLeft, panScrollTop;

const treeEl        = document.getElementById('org-tree');
const canvasContent = document.getElementById('canvas-content');
const wrapper       = document.getElementById('canvas-wrapper');
const zoomDisplay   = document.getElementById('zoom-level');
const searchInput   = document.getElementById('search-input');
const searchResults = document.getElementById('search-results');

// ── Utility ──────────────────────────────────
function esc(str) {{
  return String(str ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}}

function childrenOf(id) {{ return viewData.filter(x => x.manager === id); }}

function treeDepth(nodeId, memo={{}}) {{
  if (memo[nodeId] !== undefined) return memo[nodeId];
  const kids = childrenOf(nodeId);
  if (!kids.length) {{ memo[nodeId] = 0; return 0; }}
  memo[nodeId] = 1 + Math.max(...kids.map(k => treeDepth(k.id, memo)));
  return memo[nodeId];
}}

// ── Stats ─────────────────────────────────────
function updateStats(roots) {{
  document.getElementById('stat-total').textContent   = viewData.length;
  document.getElementById('stat-roots').textContent   = roots.length;
  const depths = roots.map(r => treeDepth(r.id));
  document.getElementById('stat-depth').textContent   = depths.length ? Math.max(...depths) : 0;
  document.getElementById('stat-visible').textContent = document.querySelectorAll('.node-card').length;
}}

// ── Render ────────────────────────────────────
function render() {{
  treeEl.innerHTML = '';
  const roots = viewData.filter(d => !d.manager || d.manager === '');
  if (!roots.length) {{
    treeEl.innerHTML = '<p style="padding:48px;color:#8896b3;font-size:0.9rem;font-weight:600;background:#161b27;border:1px solid #2a3147;border-radius:14px;">No root nodes found — ensure L1 Manager Code column is present and at least one employee has no manager.</p>';
    return;
  }}
  const ul = document.createElement('ul');
  roots.forEach(r => ul.appendChild(createNodeLI(r, 0)));
  treeEl.appendChild(ul);

  clearTimeout(autoFitTimer);
  autoFitTimer = setTimeout(() => fitToScreen(true), 160);
  setTimeout(() => updateStats(roots), 200);
}}

function createNodeLI(node, depth) {{
  const li    = document.createElement('li');
  li.dataset.id = node.id;

  const card  = document.createElement('div');
  card.className = 'node-card';
  if (node.id === state.highlighted) card.classList.add('highlighted');

  const reports = childrenOf(node.id).length;

  card.innerHTML =
    '<div class="card-header">' +
      '<span class="sub-tag" title="' + esc(node.sub) + '">' + esc(node.sub) + '</span>' +
      '<span class="grade-tag">GR&nbsp;' + esc(node.grade) + '</span>' +
    '</div>' +
    '<div class="card-body">' +
      '<div class="emp-name" title="' + esc(node.name) + '">' + esc(node.name) + '</div>' +
      '<div class="emp-title" title="' + esc(node.title) + '">' + esc(node.title) + '</div>' +
    '</div>' +
    '<div class="card-footer">' +
      '<span class="emp-id">' + esc(String(node.id).slice(0,12)) + '</span>' +
      '<span class="reports-badge' + (reports > 0 ? ' has-reports' : '') + '">' +
        reports + (reports === 1 ? ' report' : ' reports') +
      '</span>' +
    '</div>';

  if (reports > 0) {{
    const colBtn = document.createElement('div');
    colBtn.className = 'collapse-btn';
    colBtn.title = 'Collapse/expand';
    colBtn.innerHTML = '▾';
    colBtn.addEventListener('click', function(e) {{
      e.stopPropagation();
      toggleCollapse(li, colBtn);
    }});
    card.appendChild(colBtn);
  }}

  li.appendChild(card);

  const children = childrenOf(node.id);
  if (children.length) {{
    const ul = document.createElement('ul');
    children.forEach(c => ul.appendChild(createNodeLI(c, depth + 1)));
    li.appendChild(ul);
  }}
  return li;
}}

function toggleCollapse(li, btn) {{
  li.classList.toggle('collapsed');
  const isCollapsed = li.classList.contains('collapsed');
  btn.innerHTML = isCollapsed ? '▸' : '▾';
  btn.style.color = isCollapsed ? 'var(--warning)' : '';
  const card = li.querySelector('.node-card');
  if (card) card.classList.toggle('collapsed-node', isCollapsed);
  setTimeout(() => updateStats(viewData.filter(d => !d.manager)), 50);
}}

function expandAll() {{
  document.querySelectorAll('li.collapsed').forEach(li => {{
    li.classList.remove('collapsed');
    const card = li.querySelector('.node-card');
    if (card) card.classList.remove('collapsed-node');
    const btn = li.querySelector('.collapse-btn');
    if (btn) {{ btn.innerHTML = '▾'; btn.style.color = ''; }}
  }});
  setTimeout(() => updateStats(viewData.filter(d => !d.manager)), 50);
}}

function collapseAll() {{
  // Collapse all non-root nodes
  document.querySelectorAll('li').forEach(li => {{
    if (!li.closest('ul')?.closest('li')) return; // skip root
    const ul = li.querySelector(':scope > ul');
    if (ul) {{
      li.classList.add('collapsed');
      const card = li.querySelector('.node-card');
      if (card) card.classList.add('collapsed-node');
      const btn = li.querySelector('.collapse-btn');
      if (btn) {{ btn.innerHTML = '▸'; btn.style.color = 'var(--warning)'; }}
    }}
  }});
  setTimeout(() => updateStats(viewData.filter(d => !d.manager)), 50);
}}

// ── Zoom & Pan ────────────────────────────────
function applyZoom(z) {{
  state.zoom = Math.max(0.1, Math.min(2.5, z));
  canvasContent.style.transform = 'scale(' + state.zoom + ')';
  zoomDisplay.textContent = Math.round(state.zoom * 100) + '%';
}}

function zoomBy(delta) {{ applyZoom(state.zoom + delta); }}

function fitToScreen(alsoCenter = false) {{
  if (state.exporting) return;
  requestAnimationFrame(() => {{
    const tW = treeEl.scrollWidth, tH = treeEl.scrollHeight;
    const aW = wrapper.clientWidth - 100, aH = wrapper.clientHeight - 100;
    if (tW < 10 || tH < 10) return;
    const target = Math.min(aW / tW, aH / tH, 1.0);
    applyZoom(Math.max(target, 0.2));
    if (alsoCenter) setTimeout(centerView, 60);
  }});
}}

function centerView() {{
  const scaledW = treeEl.scrollWidth * state.zoom;
  wrapper.scrollLeft = Math.max(0, (scaledW - wrapper.clientWidth) / 2);
  wrapper.scrollTop  = 0;
}}

// Drag to pan
wrapper.addEventListener('mousedown', e => {{
  if (e.target.closest('.node-card') || e.target.closest('.collapse-btn')) return;
  isPanning = true;
  panStartX = e.clientX; panStartY = e.clientY;
  panScrollLeft = wrapper.scrollLeft; panScrollTop = wrapper.scrollTop;
  wrapper.style.cursor = 'grabbing';
}});
window.addEventListener('mousemove', e => {{
  if (!isPanning) return;
  wrapper.scrollLeft = panScrollLeft - (e.clientX - panStartX);
  wrapper.scrollTop  = panScrollTop  - (e.clientY - panStartY);
}});
window.addEventListener('mouseup', () => {{
  isPanning = false;
  wrapper.style.cursor = '';
}});

// Scroll-to-zoom
wrapper.addEventListener('wheel', e => {{
  if (e.ctrlKey || e.metaKey) {{
    e.preventDefault();
    zoomBy(e.deltaY < 0 ? 0.08 : -0.08);
  }}
}}, {{ passive: false }});

// ── Search ────────────────────────────────────
searchInput.addEventListener('input', function() {{
  const q = this.value.trim().toLowerCase();
  if (!q) {{ searchResults.classList.remove('visible'); return; }}

  const matches = viewData
    .filter(d => d.name.toLowerCase().includes(q) || d.id.toLowerCase().includes(q))
    .slice(0, 8);

  if (!matches.length) {{
    searchResults.innerHTML = '<div class="sr-item" style="color:var(--text-muted);">No results</div>';
    searchResults.classList.add('visible');
    return;
  }}

  searchResults.innerHTML = matches.map(d =>
    '<div class="sr-item" onclick="highlightNode(\'' + esc(d.id) + '\')">' +
      '<div class="sr-name">' + esc(d.name) + '</div>' +
      '<div class="sr-sub">' + esc(d.title) + ' &bull; ' + esc(d.id) + '</div>' +
    '</div>'
  ).join('');
  searchResults.classList.add('visible');
}});

document.addEventListener('click', e => {{
  if (!e.target.closest('.search-wrap')) {{
    searchResults.classList.remove('visible');
  }}
}});

function highlightNode(id) {{
  // Remove old highlight
  document.querySelectorAll('.node-card.highlighted').forEach(c => c.classList.remove('highlighted'));
  state.highlighted = id;

  // Expand path to node
  expandAll();

  // Find and highlight
  const li = document.querySelector('li[data-id="' + CSS.escape(id) + '"]');
  if (li) {{
    const card = li.querySelector('.node-card');
    if (card) {{
      card.classList.add('highlighted');
      setTimeout(() => {{
        const rect   = card.getBoundingClientRect();
        const wRect  = wrapper.getBoundingClientRect();
        const scrollX = wrapper.scrollLeft + (rect.left - wRect.left) - (wRect.width  / 2) + (rect.width  / 2);
        const scrollY = wrapper.scrollTop  + (rect.top  - wRect.top ) - (wRect.height / 2) + (rect.height / 2);
        wrapper.scrollTo({{ left: scrollX, top: scrollY, behavior: 'smooth' }});
        card.style.transition = 'box-shadow 0.2s';
        card.style.boxShadow  = '0 0 0 4px rgba(245,158,11,0.6), 0 12px 40px rgba(0,0,0,0.5)';
        setTimeout(() => {{ card.style.boxShadow = ''; }}, 2000);
      }}, 80);
    }}
  }}

  searchInput.value = '';
  searchResults.classList.remove('visible');
}}

// ── CSV Export ────────────────────────────────
function csvEscape(v) {{ return '"' + String(v ?? '').replace(/"/g, '""') + '"'; }}
function triggerDownload(blob, name) {{
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name; a.click();
  URL.revokeObjectURL(url);
}}
function downloadCSV() {{
  const headers = ['Employee Code','L1 Manager Code','Employee Name','Designation','Grade','Sub Function'];
  let csv = headers.join(',') + '\\n';
  viewData.forEach(r => {{
    csv += [r.id, r.manager, r.name, r.title, r.grade, r.sub].map(csvEscape).join(',') + '\\n';
  }});
  triggerDownload(new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }}), 'org_chart_export.csv');
}}

// ── PNG Export (FIXED) ────────────────────────
// KEY FIX: Instead of positioning off-screen, we render the stage
// in a hidden but IN-FLOW container so html2canvas captures text correctly.
// We also strip overflow:hidden, set explicit white background on all cards,
// and use a light theme so text is always visible.
async function exportImage() {{
  if (state.exporting) return;
  state.exporting = true;
  clearTimeout(autoFitTimer);

  // Show overlay
  const overlay = document.createElement('div');
  overlay.className = 'export-overlay';
  overlay.innerHTML =
    '<div class="export-spinner"></div>' +
    '<div class="export-title">Rendering org chart…</div>' +
    '<div class="export-sub">Large charts may take a moment</div>';
  document.body.appendChild(overlay);

  // Reset zoom for clean capture
  const savedZoom = state.zoom;
  applyZoom(1);
  await new Promise(r => setTimeout(r, 100));

  // Build a light-themed export stage
  const stage = document.createElement('div');
  stage.className = 'export-stage tree';
  stage.style.cssText = [
    'position:absolute',
    'top:0', 'left:0',
    'visibility:hidden',
    'pointer-events:none',
    'background:#f0f4f8',
    'padding:48px',
    'display:inline-block',
    'white-space:nowrap',
  ].join(';');

  // Clone the tree content
  const cloned = treeEl.cloneNode(true);

  // CRITICAL FIX: Remove all overflow:hidden so text renders in html2canvas
  cloned.querySelectorAll('*').forEach(el => {{
    el.style.overflow = 'visible';
    el.style.webkitLineClamp = 'unset';
    // Remove collapsed state for export
    el.classList.remove('collapsed');
  }});
  // Re-show all collapsed subtrees
  cloned.querySelectorAll('ul').forEach(ul => ul.style.display = '');
  // Remove collapse buttons
  cloned.querySelectorAll('.collapse-btn').forEach(b => b.remove());

  stage.appendChild(cloned);
  document.body.appendChild(stage);

  await new Promise(r => setTimeout(r, 120));
  if (document.fonts?.ready) await document.fonts.ready;
  await new Promise(r => setTimeout(r, 80));

  try {{
    const canvas = await html2canvas(stage, {{
      backgroundColor: '#f0f4f8',
      scale: 2,
      useCORS: true,
      logging: false,
      allowTaint: true,
      // CRITICAL: width/height must match stage content, not viewport
      width:  stage.scrollWidth,
      height: stage.scrollHeight,
      windowWidth:  stage.scrollWidth  + 200,
      windowHeight: stage.scrollHeight + 200,
      x: 0,
      y: 0,
    }});

    canvas.toBlob(blob => {{
      const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
      triggerDownload(blob, 'orgchart_' + stamp + '.png');
    }}, 'image/png');

  }} catch(err) {{
    console.error('Export error:', err);
    alert('Export failed: ' + err.message);
  }} finally {{
    stage.remove();
    overlay.remove();
    applyZoom(savedZoom);
    state.exporting = false;
  }}
}}

// ── Init ─────────────────────────────────────
render();
</script>

</body>
</html>"""

    components.html(html_template, height=chart_height, scrolling=True)
