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
# 2) Streamlit Styling
# =========================================================
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
  html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
  }
  #MainMenu, footer, header, .stDeployButton { display: none !important; }
  .stApp { background-color: #f0f4f8 !important; }
  .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; max-width: 99% !important; }
  [data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e5e7eb !important; }
  .stButton > button { border-radius: 10px !important; font-weight: 700 !important; transition: all 0.2s !important; width: 100% !important; }
  .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #4f46e5, #7c3aed) !important;
    color: #ffffff !important; border: none !important;
    box-shadow: 0 6px 14px rgba(79,70,229,0.25) !important;
  }
  .stButton > button[kind="secondary"] { background: #ffffff !important; color: #111827 !important; border: 1px solid #e5e7eb !important; }
  [data-testid="stSelectbox"] > div > div {
    background: #ffffff !important; border: 1px solid #e5e7eb !important;
    border-radius: 10px !important; color: #111827 !important; font-weight: 600 !important;
  }
  iframe { border: none !important; }
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
    uploaded_file = st.file_uploader("Upload HR Roster", type=["csv", "xlsx", "xls"], label_visibility="collapsed")

    df = None
    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = df.columns.str.strip()
            if "Employee Code" in df.columns:
                df["Clean_Emp_Code"] = df["Employee Code"].astype(str).str.replace(".0", "", regex=False).str.strip()
            else:
                st.error("Missing required column: 'Employee Code'")
            if "L1 Manager Code" in df.columns:
                df["L1 Manager Code"] = df["L1 Manager Code"].astype(str).str.replace(".0", "", regex=False).str.strip()
        except Exception as e:
            st.error(f"Error reading file: {e}")

    st.markdown("---")
    st.markdown("##### ⚙️ Controls")
    enable_interactive = st.toggle("Interactive Draft Mode", value=False)

    if df is not None and "Sub Function" in df.columns:
        subs = ["All"] + sorted(df["Sub Function"].dropna().unique())
        selected_sub = st.selectbox("Sub Function", subs, index=0)
    else:
        selected_sub = "All"

    chart_height = st.slider("Chart Height (px)", 600, 4000, 1400, 100)
    st.caption("Chart auto-fits on load. Use ⊡ Fit anytime to re-fit.")

# =========================================================
# 5) Main Logic
# =========================================================
if df is None:
    st.info("👈 Upload an HR CSV/XLSX file in the sidebar to generate the org chart.")
else:
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

    emp_cols = ["Clean_Emp_Code", "Employee Name"]
    if "Designation" in df.columns:
        emp_cols.append("Designation")
    all_emp_list = df[emp_cols].dropna().copy()
    if "Designation" not in all_emp_list.columns:
        all_emp_list["Designation"] = "N/A"
    all_emp_list.columns = ["id", "name", "title"]
    all_emp_list["id"] = all_emp_list["id"].astype(str).str.replace(".0", "", regex=False).str.strip()
    all_emps = all_emp_list.to_dict("records")

    # =========================================================
    # 6) HTML Component
    # =========================================================
    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>OrgDesign Pro</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<style>
  :root {{
    --bg: #f0f4f8;
    --card: #ffffff;
    --border: #dde3ec;
    --text-main: #0b1220;
    --text-sub:  #1e293b;
    --accent: #4f46e5;
    --accent-hover: #4338ca;
    --success: #059669;
    --danger: #dc2626;
    --connector: #94a3b8;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  /* FIX 1: body scrolls — NOT overflow hidden */
  html, body {{
    height: 100%; background: var(--bg);
    font-family: 'Inter', sans-serif; color: var(--text-main);
    overflow: auto;
  }}
  body {{ display: flex; flex-direction: column; min-height: 100vh; }}

  /* Toolbar */
  .toolbar {{
    position: sticky; top: 0; height: 54px;
    background: #ffffff; border-bottom: 1px solid var(--border);
    display: flex; align-items: center; padding: 0 16px; gap: 8px;
    z-index: 100; flex-shrink: 0;
    box-shadow: 0 2px 10px rgba(0,0,0,0.07);
  }}
  .brand {{ font-weight: 900; font-size: 1.05rem; letter-spacing: -0.02em; margin-right: auto; color: var(--text-main); }}
  .tool-btn {{
    background: #f8fafc; border: 1.5px solid var(--border);
    padding: 6px 13px; border-radius: 9px; cursor: pointer;
    font-weight: 700; font-size: 0.82rem; color: var(--text-main);
    transition: 0.15s; display: inline-flex; align-items: center; gap: 6px;
    white-space: nowrap; user-select: none; font-family: 'Inter', sans-serif;
  }}
  .tool-btn:hover {{ background: #eef2f7; border-color: #b0bec9; }}
  .tool-btn:active {{ transform: scale(0.97); }}
  .tool-btn.primary {{
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    color: #fff; border: none; box-shadow: 0 4px 14px rgba(79,70,229,0.30);
  }}
  .tool-btn.primary:hover {{ filter: brightness(1.1); }}
  .tool-btn.success {{
    background: linear-gradient(135deg, #059669, #0d9488);
    color: #fff; border: none; box-shadow: 0 4px 14px rgba(5,150,105,0.28);
  }}
  .tool-btn.success:hover {{ filter: brightness(1.1); }}
  .zoom-controls {{
    display: flex; align-items: center; gap: 3px;
    background: #f1f5f9; border-radius: 10px; padding: 3px;
    border: 1.5px solid var(--border);
  }}
  .zoom-controls .tool-btn {{ border: none; background: transparent; padding: 5px 10px; border-radius: 7px; }}
  .zoom-controls .tool-btn:hover {{ background: #e2e8f0; }}
  .zoom-label {{ font-size: 0.78rem; color: var(--text-main); min-width: 44px; text-align: center; font-weight: 800; }}

  /* FIX 2: Canvas — overflow auto both axes */
  .canvas-wrapper {{
    flex: 1; overflow: auto; -webkit-overflow-scrolling: touch;
    background-image: radial-gradient(#d1d9e6 1.2px, transparent 1.2px);
    background-size: 22px 22px; position: relative; cursor: grab;
    min-height: 0;
  }}
  .canvas-wrapper.dragging {{ cursor: grabbing; }}
  .canvas-wrapper::-webkit-scrollbar {{ width: 8px; height: 8px; }}
  .canvas-wrapper::-webkit-scrollbar-thumb {{ background: #b0bec9; border-radius: 8px; }}
  .canvas-wrapper::-webkit-scrollbar-track {{ background: #eef2f7; }}

  /* FIX 3: inline-block + transform-origin top left so zoom anchors correctly */
  .canvas-content {{
    display: inline-block;
    padding: 48px 64px 96px 64px;
    transform-origin: top left;
    min-width: 100%;
  }}

  /* FIX 4: Tree — no conflicting width:max-content on ul */
  .tree {{ display: inline-block; }}
  .tree ul {{
    padding-top: 20px; position: relative;
    list-style: none; display: flex; justify-content: center;
  }}
  .tree li {{
    display: table-cell; vertical-align: top;
    text-align: center; position: relative; padding: 20px 7px 0 7px;
  }}

  /* Connectors */
  .tree li::before, .tree li::after {{
    content: ''; position: absolute; top: 0; right: 50%;
    border-top: 2px solid var(--connector); width: 50%; height: 20px;
  }}
  .tree li::after {{ right: auto; left: 50%; border-left: 2px solid var(--connector); }}
  /* FIX 8: Only-child — hide ALL connectors including dangling stem */
  .tree li:only-child::before,
  .tree li:only-child::after {{ display: none !important; }}
  .tree li:only-child {{ padding-top: 0; }}
  .tree li:only-child > ul::before {{ display: none !important; }}
  .tree li:first-child::before, .tree li:last-child::after {{ border: none !important; }}
  .tree li:last-child::before {{ border-right: 2px solid var(--connector); border-radius: 0 5px 0 0; }}
  .tree li:first-child::after {{ border-radius: 5px 0 0 0; }}
  .tree ul ul::before {{
    content: ''; position: absolute; top: 0; left: 50%;
    border-left: 2px solid var(--connector); width: 0; height: 20px;
  }}

  /* Cards */
  .node-card {{
    display: inline-block; width: 200px;
    background: var(--card); border: 1.5px solid var(--border);
    border-radius: 14px; box-shadow: 0 2px 8px rgba(11,18,32,0.07);
    transition: box-shadow 0.18s, transform 0.18s; cursor: pointer;
    text-align: left; border-top: 4px solid var(--success); color: var(--text-main);
  }}
  .node-card:hover {{ transform: translateY(-3px); box-shadow: 0 12px 28px rgba(11,18,32,0.13); }}
  .node-card.selected {{ border: 2px solid var(--accent); box-shadow: 0 0 0 4px rgba(79,70,229,0.18); }}
  .node-card.is-root {{ border-top-color: var(--accent); }}
  .node-card.selecting-target {{ cursor: crosshair; }}

  .card-h {{
    padding: 9px 11px; background: #f8fafc;
    border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: center;
    border-radius: 13px 13px 0 0; gap: 8px;
  }}
  .sub-tag {{
    font-size: 0.58rem; font-weight: 800; text-transform: uppercase;
    background: #e2e8f0; color: #0b1220; padding: 2px 8px; border-radius: 999px;
    max-width: 120px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }}
  .grade-tag {{ font-size: 0.68rem; color: #0b1220; font-weight: 800; white-space: nowrap; }}
  .card-b {{ padding: 11px 12px 10px; }}
  .emp-name {{
    font-size: 0.88rem; font-weight: 900; color: #0b1220; margin-bottom: 3px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }}
  .emp-title {{
    font-size: 0.76rem; font-weight: 700; color: #1e293b; line-height: 1.35;
    display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; overflow: hidden;
  }}
  .card-f {{
    padding: 7px 12px; border-top: 1px solid #eef1f6;
    font-size: 0.68rem; color: #0b1220; font-weight: 700;
    display: flex; justify-content: space-between;
    border-radius: 0 0 14px 14px; background: #f8fafc;
  }}

  /* Modal */
  .modal-bg {{
    position: fixed; inset: 0; background: rgba(11,18,32,0.45);
    backdrop-filter: blur(3px); opacity: 0; pointer-events: none;
    transition: opacity 0.22s; z-index: 200;
  }}
  .modal-bg.open {{ opacity: 1; pointer-events: auto; }}
  .action-sheet {{
    position: fixed; bottom: 0; left: 0; right: 0;
    max-width: 540px; margin: 0 auto; background: #ffffff;
    border-top: 1px solid var(--border); border-radius: 22px 22px 0 0;
    padding: 22px 24px; box-shadow: 0 -12px 40px rgba(0,0,0,0.14);
    transform: translateY(100%);
    transition: transform 0.28s cubic-bezier(0.16, 1, 0.3, 1); z-index: 201;
  }}
  .modal-bg.open .action-sheet {{ transform: translateY(0); }}
  .sheet-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }}
  .sheet-title {{ font-weight: 900; font-size: 1.05rem; color: var(--text-main); }}
  .close-btn {{ background: none; border: none; font-size: 1.6rem; color: #334155; cursor: pointer; line-height: 1; }}
  .close-btn:hover {{ color: var(--text-main); }}
  .emp-summary {{
    display: flex; gap: 12px; align-items: center; padding: 12px;
    background: #f8fafc; border-radius: 14px; margin-bottom: 14px;
    border: 1.5px solid #e5e7eb;
  }}
  .avatar {{
    width: 42px; height: 42px; flex-shrink: 0;
    background: linear-gradient(135deg, #e0e7ff, #ede9fe);
    border-radius: 50%; display: flex; align-items: center; justify-content: center;
    font-weight: 900; color: #3730a3; font-size: 1.1rem;
  }}
  .emp-sum-name {{ font-weight: 900; color: #0b1220; font-size: 0.95rem; }}
  .emp-sum-title {{ font-size: 0.82rem; color: #1e293b; font-weight: 700; margin-top: 2px; }}
  .action-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }}
  .act-btn {{
    padding: 12px; border-radius: 12px; font-weight: 900; font-size: 0.86rem;
    cursor: pointer; transition: 0.18s; display: flex; align-items: center;
    justify-content: center; gap: 7px; font-family: 'Inter', sans-serif;
  }}
  .act-btn.move {{ background: #eff6ff; border: 1.5px solid #bfdbfe; color: #1d4ed8; }}
  .act-btn.move:hover {{ background: #dbeafe; }}
  .act-btn.del {{ background: #fef2f2; border: 1.5px solid #fecaca; color: #b91c1c; }}
  .act-btn.del:hover {{ background: #fee2e2; }}
  .search-box {{ margin-top: 12px; display: none; }}
  .search-box.active {{ display: block; }}
  .s-input {{
    width: 100%; padding: 10px 13px; border: 1.5px solid var(--border);
    border-radius: 11px; font-size: 0.9rem; outline: none; font-weight: 700;
    color: var(--text-main); background: #fff; font-family: 'Inter', sans-serif;
  }}
  .s-input:focus {{ border-color: var(--accent); box-shadow: 0 0 0 3px rgba(79,70,229,0.11); }}
  .s-results {{
    max-height: 168px; overflow-y: auto; margin-top: 8px;
    border: 1.5px solid var(--border); border-radius: 11px; background: #fff;
  }}
  .s-item {{
    padding: 9px 13px; cursor: pointer; font-size: 0.86rem;
    border-bottom: 1px solid #f1f5f9; color: var(--text-main); font-weight: 700;
  }}
  .s-item:hover {{ background: #f8fafc; }}
  .s-item:last-child {{ border: none; }}
  .mode-banner {{
    background: #fffbeb; color: #78350f; padding: 10px 13px;
    border-radius: 11px; font-size: 0.86rem; font-weight: 900;
    margin-bottom: 12px; text-align: center; display: none; border: 1.5px solid #fcd34d;
  }}
  .mode-banner.show {{ display: block; }}

  /* Export loading overlay */
  .export-loading {{
    position: fixed; inset: 0; z-index: 9999;
    background: rgba(11,18,32,0.72); backdrop-filter: blur(4px);
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; gap: 18px; color: #fff;
    font-size: 1.1rem; font-weight: 800; font-family: 'Inter', sans-serif;
  }}
  .export-spinner {{
    width: 52px; height: 52px; border: 5px solid rgba(255,255,255,0.2);
    border-top-color: #fff; border-radius: 50%;
    animation: spin 0.75s linear infinite;
  }}
  @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
</style>
</head>
<body>

<div class="toolbar">
  <div class="brand">🏢 OrgDesign Pro</div>
  <div class="zoom-controls">
    <button class="tool-btn" onclick="zoomBy(-0.1)">−</button>
    <span class="zoom-label" id="zoom-level">100%</span>
    <button class="tool-btn" onclick="zoomBy(0.1)">+</button>
    <button class="tool-btn" onclick="fitToScreen()">⊡ Fit</button>
  </div>
  <button class="tool-btn" onclick="centerView()">🧭 Center</button>
  <button class="tool-btn success" onclick="exportImage()">🖼️ Save Image</button>
  <button class="tool-btn primary" onclick="downloadCSV()">💾 CSV</button>
</div>

<div class="canvas-wrapper" id="canvas-wrapper">
  <div class="canvas-content" id="canvas-content">
    <div class="tree" id="org-tree"></div>
  </div>
</div>

<div class="modal-bg" id="modal-bg" onclick="closeModal()">
  <div class="action-sheet" onclick="event.stopPropagation()">
    <div class="mode-banner" id="mode-banner">👆 Click any manager card to assign them as new manager</div>
    <div class="sheet-header">
      <span class="sheet-title">Manage Employee</span>
      <button class="close-btn" onclick="closeModal()">×</button>
    </div>
    <div class="emp-summary">
      <div class="avatar" id="sel-avatar">A</div>
      <div>
        <div class="emp-sum-name" id="sel-name">Name</div>
        <div class="emp-sum-title" id="sel-title">Title</div>
      </div>
    </div>
    <div class="action-grid">
      <button class="act-btn move" id="btn-move" onclick="startMoveMode()">🔄 Reassign Manager</button>
      <button class="act-btn del" onclick="removeEmployee()">🗑️ Remove Node</button>
    </div>
    <div class="search-box" id="search-box">
      <input type="text" class="s-input" id="search-input" placeholder="Search manager by name…"/>
      <div class="s-results" id="search-results"></div>
    </div>
  </div>
</div>

<script>
const viewData = {json.dumps(view_data)};
const allEmps  = {json.dumps(all_emps)};

let state = {{ selected: null, mode: 'normal', zoom: 1 }};

const treeEl        = document.getElementById('org-tree');
const canvasContent = document.getElementById('canvas-content');
const wrapper       = document.getElementById('canvas-wrapper');
const modalBg       = document.getElementById('modal-bg');
const searchBox     = document.getElementById('search-box');
const searchInput   = document.getElementById('search-input');
const resultsDiv    = document.getElementById('search-results');
const modeBanner    = document.getElementById('mode-banner');
const zoomDisplay   = document.getElementById('zoom-level');

function esc(str) {{
  return String(str ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}}

// ── Render ────────────────────────────────────────────────────────────────────
function render() {{
  treeEl.innerHTML = '';
  const roots = viewData.filter(d => !d.manager || d.manager === '');
  if (!roots.length) {{
    treeEl.innerHTML = '<p style="padding:40px;color:#475569;font-size:0.9rem;">No root nodes found — check L1 Manager Code column.</p>';
    return;
  }}
  const ul = document.createElement('ul');
  roots.forEach(r => ul.appendChild(createNode(r)));
  treeEl.appendChild(ul);
  // FIX 5+7: always auto-fit after every render (covers filter changes too)
  setTimeout(fitToScreen, 100);
}}

function createNode(node) {{
  const li   = document.createElement('li');
  const card = document.createElement('div');
  card.className = 'node-card' +
    (!node.manager ? ' is-root' : '') +
    (state.selected === node.id ? ' selected' : '') +
    (state.mode === 'selecting_manager' ? ' selecting-target' : '');

  const reports = viewData.filter(x => x.manager === node.id).length;
  card.innerHTML =
    '<div class="card-h">' +
      '<span class="sub-tag" title="' + esc(node.sub) + '">' + esc(node.sub) + '</span>' +
      '<span class="grade-tag">GR ' + esc(node.grade) + '</span>' +
    '</div>' +
    '<div class="card-b">' +
      '<div class="emp-name" title="' + esc(node.name) + '">' + esc(node.name) + '</div>' +
      '<div class="emp-title" title="' + esc(node.title) + '">' + esc(node.title) + '</div>' +
    '</div>' +
    '<div class="card-f">' +
      '<span>' + esc(String(node.id).slice(0,10)) + '</span>' +
      '<span>' + reports + ' direct' + (reports !== 1 ? 's' : '') + '</span>' +
    '</div>';

  card.onclick = e => {{ e.stopPropagation(); handleNodeClick(node); }};
  li.appendChild(card);

  const children = viewData.filter(c => c.manager === node.id);
  if (children.length) {{
    const ul = document.createElement('ul');
    children.forEach(c => ul.appendChild(createNode(c)));
    li.appendChild(ul);
  }}
  return li;
}}

// ── Interaction ───────────────────────────────────────────────────────────────
function handleNodeClick(node) {{
  if (state.mode === 'selecting_manager' && state.selected !== node.id) {{
    executeMove(state.selected, node.id);
  }} else if (state.mode === 'normal') {{
    state.selected = node.id;
    openModal(node);
    render();
  }}
}}

function openModal(node) {{
  document.getElementById('sel-name').textContent  = node.name;
  document.getElementById('sel-title').textContent = node.title;
  document.getElementById('sel-avatar').textContent = String(node.name || 'A').charAt(0).toUpperCase();
  modalBg.classList.add('open');
}}

function closeModal() {{
  modalBg.classList.remove('open');
  state.selected = null;
  state.mode = 'normal';
  searchBox.classList.remove('active');
  modeBanner.classList.remove('show');
  document.getElementById('btn-move').innerHTML = '🔄 Reassign Manager';
  searchInput.value = ''; resultsDiv.innerHTML = '';
  render();
}}

function startMoveMode() {{
  state.mode = 'selecting_manager';
  modeBanner.classList.add('show');
  searchBox.classList.add('active');
  document.getElementById('btn-move').innerHTML = '⏳ Selecting…';
  searchInput.focus(); render();
}}

function executeMove(empId, newMgrId) {{
  if (!empId || !newMgrId || empId === newMgrId) return;
  const emp = viewData.find(x => x.id === empId);
  if (!emp) return;
  emp.manager = newMgrId;
  closeModal();
  showToast('✅ Manager updated');
}}

function removeEmployee() {{
  if (!confirm('Remove this employee? Their direct reports will move up.')) return;
  const idx = viewData.findIndex(x => x.id === state.selected);
  if (idx > -1) {{
    const node = viewData[idx];
    viewData.filter(c => c.manager === node.id).forEach(c => c.manager = node.manager);
    viewData.splice(idx, 1);
  }}
  closeModal();
}}

// ── Search ────────────────────────────────────────────────────────────────────
searchInput.addEventListener('input', e => {{
  const q = e.target.value.toLowerCase().trim();
  resultsDiv.innerHTML = '';
  if (!q) return;
  allEmps
    .filter(emp => String(emp.name||'').toLowerCase().includes(q) && String(emp.id) !== String(state.selected))
    .slice(0, 14)
    .forEach(emp => {{
      const div = document.createElement('div');
      div.className = 's-item';
      div.innerHTML = '<b>' + esc(emp.name) + '</b> <span style="color:#475569;font-weight:600">— ' + esc(emp.title) + '</span>';
      div.onclick = () => executeMove(state.selected, emp.id);
      resultsDiv.appendChild(div);
    }});
}});

// ── Zoom ──────────────────────────────────────────────────────────────────────
function applyZoom(z) {{
  state.zoom = Math.max(0.1, Math.min(2.5, z));
  // FIX 3: transform-origin top left — no content jump on zoom
  canvasContent.style.transformOrigin = 'top left';
  canvasContent.style.transform = 'scale(' + state.zoom + ')';
  zoomDisplay.textContent = Math.round(state.zoom * 100) + '%';
}}

function zoomBy(delta) {{ applyZoom(state.zoom + delta); }}

// FIX 7: Auto-fit — scale tree to fill wrapper, then center
function fitToScreen() {{
  requestAnimationFrame(() => {{
    const treeW  = treeEl.scrollWidth;
    const treeH  = treeEl.scrollHeight;
    const availW = wrapper.clientWidth  - 80;
    const availH = wrapper.clientHeight - 80;
    if (treeW < 10 || treeH < 10) return;
    const scaleW = availW / treeW;
    const scaleH = availH / treeH;
    // Fit whole chart; don't upscale tiny charts beyond 1.0
    applyZoom(Math.min(scaleW, scaleH, 1.0));
    setTimeout(() => {{
      // After zoom applied, center horizontally
      const scaledW = treeEl.scrollWidth * state.zoom;
      wrapper.scrollLeft = Math.max(0, (scaledW - wrapper.clientWidth) / 2);
      wrapper.scrollTop  = 0;
    }}, 40);
  }});
}}

// FIX 5: centerView has no gate flag — always works
function centerView() {{
  const maxLeft = wrapper.scrollWidth  - wrapper.clientWidth;
  wrapper.scrollLeft = Math.max(0, Math.floor(maxLeft / 2));
  wrapper.scrollTop  = 0;
}}

// ── Drag to pan (mouse) ───────────────────────────────────────────────────────
let isPanning = false, panX = 0, panY = 0, panSL = 0, panST = 0;
wrapper.addEventListener('mousedown', e => {{
  if (e.target.closest('.node-card,.toolbar,.action-sheet')) return;
  isPanning = true; wrapper.classList.add('dragging');
  panX = e.pageX; panY = e.pageY;
  panSL = wrapper.scrollLeft; panST = wrapper.scrollTop;
}});
window.addEventListener('mouseup', () => {{ isPanning = false; wrapper.classList.remove('dragging'); }});
wrapper.addEventListener('mousemove', e => {{
  if (!isPanning) return; e.preventDefault();
  wrapper.scrollLeft = panSL - (e.pageX - panX);
  wrapper.scrollTop  = panST - (e.pageY - panY);
}});

// Touch pan
let touch0 = null;
wrapper.addEventListener('touchstart', e => {{
  if (e.touches.length === 1) touch0 = {{ x: e.touches[0].pageX, y: e.touches[0].pageY, sl: wrapper.scrollLeft, st: wrapper.scrollTop }};
}}, {{passive:true}});
wrapper.addEventListener('touchmove', e => {{
  if (!touch0 || e.touches.length !== 1) return;
  wrapper.scrollLeft = touch0.sl - (e.touches[0].pageX - touch0.x);
  wrapper.scrollTop  = touch0.st - (e.touches[0].pageY - touch0.y);
}}, {{passive:true}});
wrapper.addEventListener('touchend', () => {{ touch0 = null; }});

// ── CSV Download ──────────────────────────────────────────────────────────────
function csvEscape(v) {{ return '"' + String(v??'').replace(/"/g,'""') + '"'; }}

function downloadCSV() {{
  const headers = ['Employee Code','L1 Manager Code','Employee Name','Designation','Grade','Sub Function'];
  let csv = headers.join(',') + '\\n';
  viewData.forEach(r => {{
    csv += [r.id, r.manager, r.name, r.title, r.grade, r.sub].map(csvEscape).join(',') + '\\n';
  }});
  triggerDownload(new Blob([csv], {{type:'text/csv;charset=utf-8;'}}), 'org_draft_updated.csv');
}}

function triggerDownload(blob, filename) {{
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}}

// ── Image Export ──────────────────────────────────────────────────────────────
async function exportImage() {{
  const overlay = document.createElement('div');
  overlay.className = 'export-loading';
  overlay.innerHTML =
    '<div class="export-spinner"></div>' +
    '<div>Rendering org chart…</div>' +
    '<div style="font-size:0.78rem;opacity:0.65;font-weight:600">Large charts may take a moment</div>';
  document.body.appendChild(overlay);

  const savedZoom = state.zoom;
  // Reset zoom to 1 for full-resolution capture
  applyZoom(1);
  await new Promise(r => setTimeout(r, 150));

  try {{
    const canvas = await html2canvas(canvasContent, {{
      backgroundColor: '#f0f4f8',
      scale: 2,           // retina quality
      useCORS: true,
      logging: false,
      width:  canvasContent.scrollWidth,
      height: canvasContent.scrollHeight,
      windowWidth:  canvasContent.scrollWidth  + 200,
      windowHeight: canvasContent.scrollHeight + 200,
    }});

    // Compose final image with branded header + footer
    const HEADER = 64;
    const FOOTER = 44;
    const PAD    = 32;
    const final  = document.createElement('canvas');
    final.width  = canvas.width  + PAD * 2;
    final.height = canvas.height + HEADER + FOOTER + PAD * 2;

    const ctx = final.getContext('2d');

    // Background
    ctx.fillStyle = '#f0f4f8';
    ctx.fillRect(0, 0, final.width, final.height);

    // White header band
    const hBand = HEADER + PAD;
    ctx.fillStyle = '#ffffff';
    ctx.shadowColor = 'rgba(0,0,0,0.06)';
    ctx.shadowBlur  = 12;
    ctx.fillRect(0, 0, final.width, hBand);
    ctx.shadowBlur = 0;

    // Accent bar at very top
    const grad = ctx.createLinearGradient(0, 0, final.width, 0);
    grad.addColorStop(0,   '#4f46e5');
    grad.addColorStop(0.5, '#7c3aed');
    grad.addColorStop(1,   '#0d9488');
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, final.width, 5);

    // Title text
    ctx.fillStyle = '#0b1220';
    ctx.font = 'bold 28px Inter, Arial, sans-serif';
    ctx.fillText('🏢  OrgDesign Pro', PAD, PAD + 30);

    // Subtitle / date
    ctx.fillStyle = '#475569';
    ctx.font = '500 15px Inter, Arial, sans-serif';
    const dateStr = 'Org Chart Export · ' + new Date().toLocaleDateString('en-GB', {{day:'2-digit',month:'short',year:'numeric'}});
    ctx.fillText(dateStr, PAD, PAD + 52);

    // Employee count badge
    const countText = viewData.length + ' employees';
    ctx.font = 'bold 13px Inter, Arial, sans-serif';
    const badgeW = ctx.measureText(countText).width + 20;
    const badgeX = final.width - PAD - badgeW;
    const badgeY = PAD + 14;
    ctx.fillStyle = '#4f46e5';
    ctx.beginPath();
    ctx.roundRect(badgeX, badgeY, badgeW, 26, 13);
    ctx.fill();
    ctx.fillStyle = '#ffffff';
    ctx.fillText(countText, badgeX + 10, badgeY + 17);

    // Org chart content
    ctx.drawImage(canvas, PAD, hBand + 8);

    // Footer
    const footerY = hBand + 8 + canvas.height + 16;
    ctx.fillStyle = '#94a3b8';
    ctx.font = '500 12px Inter, Arial, sans-serif';
    ctx.fillText('Generated by OrgDesign Pro', PAD, footerY);

    // Bottom accent line
    ctx.fillStyle = grad;
    ctx.fillRect(0, final.height - 4, final.width, 4);

    final.toBlob(blob => {{
      const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
      triggerDownload(blob, 'orgchart_' + stamp + '.png');
      showToast('🖼️ Image saved!');
    }}, 'image/png');

  }} catch(err) {{
    showToast('❌ Export failed — ' + err.message);
    console.error(err);
  }} finally {{
    overlay.remove();
    applyZoom(savedZoom);
  }}
}}

// ── Toast ─────────────────────────────────────────────────────────────────────
function showToast(msg) {{
  const t = document.createElement('div');
  t.style.cssText = 'position:fixed;bottom:22px;right:22px;background:#0b1220;color:#fff;' +
    'padding:11px 18px;border-radius:12px;font-size:0.9rem;font-weight:800;z-index:9999;' +
    'transform:translateY(18px);opacity:0;transition:0.22s;font-family:Inter,sans-serif;' +
    'box-shadow:0 8px 24px rgba(0,0,0,0.22);';
  t.textContent = msg;
  document.body.appendChild(t);
  requestAnimationFrame(() => {{ t.style.transform = 'translateY(0)'; t.style.opacity = '1'; }});
  setTimeout(() => {{
    t.style.opacity = '0'; t.style.transform = 'translateY(18px)';
    setTimeout(() => t.remove(), 250);
  }}, 2400);
}}

render();
</script>
</body>
</html>"""

    components.html(html_template, height=chart_height, scrolling=True)
