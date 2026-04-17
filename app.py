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
  /* (keep your styling as-is) */
  html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif !important;
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

    chart_height = st.slider("Chart Height (px)", 600, 4000, 1600, 100)
    st.caption("Chart auto-fits on load. Use ⊡ Fit anytime to re-fit. Drag to pan.")

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
    # 6) HTML Component  (ONLY FIXES FOR "NO TEXT" BELOW)
    # =========================================================
    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>OrgDesign Pro</title>

<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.jstyle>
  :root {{
    --bg: #f0f4f8;
    --card: #ffffff;
    --border: #dde3ec;
    --text-main: #0b1220;
    --text-sub:  #0f172a;
    --accent: #4f46e5;
    --accent-hover: #4338ca;
    --success: #059669;
    --danger: #dc2626;
    --connector: #94a3b8;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  html, body {{
    height: 100%;
    background: var(--bg);
    font-family: Arial, Helvetica, sans-serif; /* NORMAL FONT (no Google font dependency) */
    color: var(--text-main);
    overflow: hidden;
  }}

  body {{
    display: flex;
    flex-direction: column;
    height: 100vh;
  }}

  .toolbar {{
    flex-shrink: 0;
    height: 56px;
    background: #ffffff;
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    padding: 0 16px;
    gap: 8px;
    z-index: 100;
    box-shadow: 0 2px 10px rgba(0,0,0,0.07);
  }}

  .brand {{
    font-weight: 900;
    font-size: 1.05rem;
    letter-spacing: -0.02em;
    margin-right: auto;
    color: var(--text-main);
  }}

  .tool-btn {{
    background: #f8fafc;
    border: 1.5px solid var(--border);
    padding: 7px 13px;
    border-radius: 10px;
    cursor: pointer;
    font-weight: 800;
    font-size: 0.84rem;
    color: var(--text-main);
    transition: 0.15s;
    display: inline-flex;
    align-items: center;
    gap: 7px;
    white-space: nowrap;
    user-select: none;
    font-family: Arial, Helvetica, sans-serif;
  }}

  .tool-btn.primary {{
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    color: #fff; border: none;
    box-shadow: 0 4px 14px rgba(79,70,229,0.30);
  }}

  .tool-btn.success {{
    background: linear-gradient(135deg, #059669, #0d9488);
    color: #fff; border: none;
    box-shadow: 0 4px 14px rgba(5,150,105,0.28);
  }}

  .zoom-controls {{
    display: flex;
    align-items: center;
    gap: 3px;
    background: #f1f5f9;
    border-radius: 10px;
    padding: 3px;
    border: 1.5px solid var(--border);
  }}

  .zoom-controls .tool-btn {{
    border: none;
    background: transparent;
    padding: 6px 10px;
    border-radius: 8px;
  }}

  .zoom-label {{
    font-size: 0.78rem;
    color: var(--text-main);
    min-width: 48px;
    text-align: center;
    font-weight: 900;
  }}

  .canvas-wrapper {{
    flex: 1;
    overflow: auto;
    -webkit-overflow-scrolling: touch;
    background-image: radial-gradient(#d1d9e6 1.2px, transparent 1.2px);
    background-size: 22px 22px;
    position: relative;
    cursor: grab;
  }}

  .canvas-content {{
    display: inline-block;
    padding: 48px 64px 96px 64px;
    transform-origin: top left;
    min-width: 100%;
  }}

  .tree {{ display: inline-block; }}
  .tree ul {{
    padding-top: 20px;
    position: relative;
    list-style: none;
    display: flex;
    justify-content: center;
  }}
  .tree li {{
    display: table-cell;
    vertical-align: top;
    text-align: center;
    position: relative;
    padding: 20px 7px 0 7px;
  }}

  .tree li::before, .tree li::after {{
    content: '';
    position: absolute;
    top: 0;
    right: 50%;
    border-top: 2px solid var(--connector);
    width: 50%;
    height: 20px;
  }}
  .tree li::after {{
    right: auto;
    left: 50%;
    border-left: 2px solid var(--connector);
  }}

  .node-card {{
    display: inline-block;
    width: 210px;
    background: var(--card);
    border: 1.5px solid var(--border);
    border-radius: 14px;
    box-shadow: 0 2px 8px rgba(11,18,32,0.07);
    cursor: pointer;
    text-align: left;
    border-top: 4px solid var(--success);
    color: var(--text-main);
  }}

  /* IMPORTANT: do NOT use -webkit-text-fill-color here; it can break canvas capture */
  .node-card * {{
    color: var(--text-main);
    opacity: 1;
  }}

  .card-h {{
    padding: 9px 11px;
    background: #f8fafc;
    border-bottom: 1px solid var(--border);
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-radius: 13px 13px 0 0;
    gap: 8px;
  }}

  .sub-tag {{
    font-size: 0.6rem;
    font-weight: 900;
    text-transform: uppercase;
    background: #e2e8f0;
    padding: 2px 8px;
    border-radius: 999px;
    max-width: 120px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }}

  .grade-tag {{
    font-size: 0.7rem;
    font-weight: 900;
    white-space: nowrap;
  }}

  .card-b {{ padding: 11px 12px 10px; }}

  .emp-name {{
    font-size: 0.9rem;
    font-weight: 900;
    margin-bottom: 3px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }}

  .emp-title {{
    font-size: 0.78rem;
    font-weight: 800;
    line-height: 1.35;
    max-height: calc(1.35em * 2);
    overflow: hidden;
  }}

  .card-f {{
    padding: 7px 12px;
    border-top: 1px solid #eef1f6;
    font-size: 0.7rem;
    font-weight: 800;
    display: flex;
    justify-content: space-between;
    border-radius: 0 0 14px 14px;
    background: #f8fafc;
  }}

  /* Export loading overlay */
  .export-loading {{
    position: fixed; inset: 0; z-index: 9999;
    background: rgba(11,18,32,0.72);
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; gap: 18px; color: #fff;
    font-size: 1.1rem; font-weight: 900;
  }}
  .export-spinner {{
    width: 52px; height: 52px; border: 5px solid rgba(255,255,255,0.2);
    border-top-color: #fff; border-radius: 50%;
    animation: spin 0.75s linear infinite;
  }}
  @keyframes spin {{ to {{ transform: rotate(360deg); }} }}

  /* ✅ Export-only font boost so text is not lost on huge charts */
  .export-mode .emp-name {{ font-size: 16px !important; }}
  .export-mode .emp-title {{ font-size: 13px !important; }}
  .export-mode .card-f    {{ font-size: 12px !important; }}
</style>
</head>

<body>
<div class="toolbar">
  <div class="brand">🏢 OrgDesign Pro</div>
  <div class="zoom-controls">
    <button class="tool-btn" onclick="zoomBy(-0.1)">−</button>
    <span class="zoom-label" id="zoom-level">100%</span>
    <button class="tool-btn" onclick="zoomBy(0.1)">+</button>
    <button class="tool-btn" onclick="fitToScreen(true)">⊡ Fit</button>
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

<script>
const viewData = {json.dumps(view_data)};
const allEmps  = {json.dumps(all_emps)};

/* ✅ FIX 1: add exporting lock + debounce timer */
let state = {{ selected: null, mode: 'normal', zoom: 1, exporting: false }};
let autoFitTimer = null;

const treeEl        = document.getElementById('org-tree');
const canvasContent = document.getElementById('canvas-content');
const wrapper       = document.getElementById('canvas-wrapper');
const zoomDisplay   = document.getElementById('zoom-level');

/* ✅ FIX 2: correct esc() (yours was wrong) */
function esc(str) {{
  return String(str ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}}

function render() {{
  treeEl.innerHTML = '';
  const roots = viewData.filter(d => !d.manager || d.manager === '');
  if (!roots.length) {{
    treeEl.innerHTML = '<p style="padding:40px;color:#0f172a;font-size:0.95rem;font-weight:800;">No root nodes found — check L1 Manager Code column.</p>';
    return;
  }}
  const ul = document.createElement('ul');
  roots.forEach(r => ul.appendChild(createNode(r)));
  treeEl.appendChild(ul);

  /* ✅ FIX 3: debounce auto-fit, and never during export */
  clearTimeout(autoFitTimer);
  autoFitTimer = setTimeout(() => fitToScreen(true), 140);
}}

function createNode(node) {{
  const li   = document.createElement('li');
  const card = document.createElement('div');
  card.className = 'node-card';

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

  li.appendChild(card);

  const children = viewData.filter(c => c.manager === node.id);
  if (children.length) {{
    const ul = document.createElement('ul');
    children.forEach(c => ul.appendChild(createNode(c)));
    li.appendChild(ul);
  }}
  return li;
}}

function applyZoom(z) {{
  state.zoom = Math.max(0.1, Math.min(2.5, z));
  canvasContent.style.transformOrigin = 'top left';
  canvasContent.style.transform = 'scale(' + state.zoom + ')';
  zoomDisplay.textContent = Math.round(state.zoom * 100) + '%';
}}

function zoomBy(delta) {{ applyZoom(state.zoom + delta); }}

/* ✅ FIX 4: clamp minimum zoom so text never becomes microscopic */
function fitToScreen(alsoCenter=false) {{
  if (state.exporting) return;   // ✅ do NOT change zoom while exporting

  requestAnimationFrame(() => {{
    const treeW  = treeEl.scrollWidth;
    const treeH  = treeEl.scrollHeight;
    const availW = wrapper.clientWidth  - 80;
    const availH = wrapper.clientHeight - 80;
    if (treeW < 10 || treeH < 10) return;

    const scaleW = availW / treeW;
    const scaleH = availH / treeH;

    const target = Math.min(scaleW, scaleH, 1.0);
    applyZoom(Math.max(target, 0.35));  // ✅ MIN 35% keeps text visible (use 0.25 if org is massive)

    if (alsoCenter) setTimeout(centerView, 60);
  }});
}}

function centerView() {{
  const scaledW = treeEl.scrollWidth * state.zoom;
  const maxLeft = Math.max(0, scaledW - wrapper.clientWidth);
  wrapper.scrollLeft = Math.floor(maxLeft / 2);
  wrapper.scrollTop  = 0;
}}

function csvEscape(v) {{ return '"' + String(v??'').replace(/"/g,'""') + '"'; }}
function triggerDownload(blob, filename) {{
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}}
function downloadCSV() {{
  const headers = ['Employee Code','L1 Manager Code','Employee Name','Designation','Grade','Sub Function'];
  let csv = headers.join(',') + '\\n';
  viewData.forEach(r => {{
    csv += [r.id, r.manager, r.name, r.title, r.grade, r.sub].map(csvEscape).join(',') + '\\n';
  }});
  triggerDownload(new Blob([csv], {{type:'text/csv;charset=utf-8;'}}), 'org_draft_updated.csv');
}}

/* ✅ FIX 5: export lock + export-mode + higher scale */
async function exportImage() {{
  state.exporting = true;
  clearTimeout(autoFitTimer);

  const overlay = document.createElement('div');
  overlay.className = 'export-loading';
  overlay.innerHTML =
    '<div class="export-spinner"></div>' +
    '<div>Rendering org chart…</div>' +
    '<div style="font-size:0.78rem;opacity:0.7;font-weight:700;">Large charts may take a moment</div>';
  document.body.appendChild(overlay);

  const savedZoom = state.zoom;
  applyZoom(1);

  // Wait for fonts/layout
  try {{
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
  }} catch(e) {{}}
  await new Promise(r => setTimeout(r, 160));

  // Create export stage
  const stage = document.createElement('div');
  stage.className = 'export-mode';       // ✅ boosts font sizes during export
  stage.style.position = 'fixed';
  stage.style.left = '-10000px';
  stage.style.top = '0';
  stage.style.background = '#f0f4f8';
  stage.style.padding = '40px';
  stage.style.width = (treeEl.scrollWidth + 200) + 'px';

  const clonedTree = treeEl.cloneNode(true);
  stage.appendChild(clonedTree);
  document.body.appendChild(stage);

  try {{
    const canvas = await html2canvas(stage, {{
      backgroundColor: '#f0f4f8',
      scale: 3,     // ✅ higher DPI so text renders (try 4 if still faint)
      useCORS: true,
      logging: false
    }});

    canvas.toBlob(blob => {{
      const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
      triggerDownload(blob, 'orgchart_' + stamp + '.png');
    }}, 'image/png');

  }} catch(err) {{
    console.error(err);
    alert('Export failed: ' + err.message);
  }} finally {{
    stage.remove();
    overlay.remove();
    applyZoom(savedZoom);
    state.exporting = false;
  }}
}}

render();
</script>

</body>
</html>
"""

    components.html(html_template, height=chart_height, scrolling=True)
``
