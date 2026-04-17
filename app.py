import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import json

# ── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OrgDesign Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Syne:wght@700;800&display=swap');

  html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
  }
  #MainMenu, footer, header, .stDeployButton { visibility: hidden; display: none !important; }

  .stApp {
    background: #f0f2f5 !important;
    min-height: 100vh;
  }

  .block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 3rem !important;
    max-width: 99% !important;
  }

  h1 {
    font-family: 'Syne', sans-serif !important;
    font-weight: 800 !important;
    color: #0a0a0f !important;
    letter-spacing: -0.04em !important;
    font-size: 2.2rem !important;
  }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #0f0f14 !important;
    border-right: 1px solid #1e1e2a !important;
  }
  [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
  [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4, [data-testid="stSidebar"] h5 {
    color: #f8fafc !important;
    font-family: 'Syne', sans-serif !important;
  }
  [data-testid="stSidebar"] .stMarkdown p {
    color: #94a3b8 !important;
  }

  /* Selectbox in sidebar */
  [data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: #1a1a24 !important;
    border: 1px solid #2d2d3d !important;
    border-radius: 8px !important;
    color: #e2e8f0 !important;
  }

  /* File uploader */
  [data-testid="stFileUploadDropzone"] {
    background: #1a1a24 !important;
    border: 2px dashed #2d2d3d !important;
    border-radius: 12px !important;
  }
  [data-testid="stFileUploadDropzone"]:hover {
    border-color: #6366f1 !important;
  }

  /* Buttons */
  .stButton > button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
  }
  .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
    border: none !important;
    color: #fff !important;
    box-shadow: 0 4px 12px rgba(99,102,241,0.3) !important;
  }
  .stButton > button[kind="secondary"] {
    background: #1a1a24 !important;
    border: 1px solid #2d2d3d !important;
    color: #e2e8f0 !important;
  }

  /* Toggle */
  [data-testid="stToggle"] > label > div > div[aria-checked="true"] {
    background-color: #6366f1 !important;
    border-color: #6366f1 !important;
  }

  /* Info badge */
  .stat-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #1a1a24;
    border: 1px solid #2d2d3d;
    border-radius: 20px;
    padding: 5px 12px;
    font-size: 0.8rem;
    color: #94a3b8;
    margin-bottom: 6px;
  }
  .stat-pill b { color: #f8fafc; }

  .divider {
    border: none;
    border-top: 1px solid #1e1e2a;
    margin: 1rem 0;
  }
</style>
""", unsafe_allow_html=True)

# ── Session State ─────────────────────────────────────────────────────────────
if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏢 OrgDesign Pro")
    st.markdown("<p style='color:#64748b;font-size:0.83rem;margin-top:-0.4rem;margin-bottom:1.2rem;'>Interactive Org Architecture</p>", unsafe_allow_html=True)
    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    st.markdown("##### 📂 Upload HR Data")
    uploaded_file = st.file_uploader("", type=['csv', 'xlsx'], label_visibility="collapsed")

    df = None
    if uploaded_file:
        st.session_state.file_uploaded = True
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        if 'Employee Code' in df.columns:
            df['Clean_Emp_Code'] = df['Employee Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
            if 'L1 Manager Code' in df.columns:
                df['L1 Manager Code'] = df['L1 Manager Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
    else:
        st.session_state.file_uploaded = False

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)
    st.markdown("##### ⚙️ Controls")

    selected_sub = "All"
    include_retainers = True
    include_inactive = False
    enable_draft = False

    if st.session_state.file_uploaded and df is not None:
        if 'Sub Function' in df.columns:
            sub_functions = sorted(df['Sub Function'].dropna().unique())
            selected_sub = st.selectbox("Sub Function", ["All"] + list(sub_functions))

        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)

        st.markdown("<hr class='divider'>", unsafe_allow_html=True)
        st.markdown("##### 🖊️ Draft Mode")
        enable_draft = st.toggle("Enable Draft Editing", value=False,
                                  help="Right-click any card to move, replace, or delete.")
        if enable_draft:
            st.markdown("""
            <div style='background:#1a2740;border:1px solid #1e3a5f;border-radius:10px;padding:12px;margin-top:8px;'>
              <p style='color:#60a5fa;font-size:0.83rem;margin:0;font-weight:600;'>✦ Draft Mode Active</p>
              <ul style='color:#7eb3f8;font-size:0.78rem;margin:6px 0 0 0;padding-left:1.2rem;'>
                <li>Right-click any card for options</li>
                <li>Change Manager → select new parent</li>
                <li>Replace Person → swap employee</li>
                <li>Delete → remove from chart</li>
              </ul>
            </div>
            """, unsafe_allow_html=True)

        # Stats
        st.markdown("<hr class='divider'>", unsafe_allow_html=True)
        st.markdown("##### 📊 Summary")
        total = len(df)
        st.markdown(f"<div class='stat-pill'>👥 Total <b>{total}</b></div>", unsafe_allow_html=True)
        if 'Sub Function' in df.columns:
            subs = df['Sub Function'].nunique()
            st.markdown(f"<div class='stat-pill'>🗂 Sub Functions <b>{subs}</b></div>", unsafe_allow_html=True)
        if 'Grade' in df.columns:
            grades = df['Grade'].nunique()
            st.markdown(f"<div class='stat-pill'>🏅 Grades <b>{grades}</b></div>", unsafe_allow_html=True)

# ── Main ──────────────────────────────────────────────────────────────────────
if not st.session_state.file_uploaded or df is None:
    # Landing page
    st.markdown("""
    <div style='display:flex;flex-direction:column;align-items:center;justify-content:center;padding:5rem 2rem;text-align:center;'>
      <div style='font-size:4rem;margin-bottom:1rem;'>🏢</div>
      <h1 style='font-family:Syne,sans-serif;font-size:2.8rem;font-weight:800;letter-spacing:-0.04em;color:#0a0a0f;margin-bottom:0.5rem;'>OrgDesign Pro</h1>
      <p style='color:#64748b;font-size:1.1rem;max-width:520px;margin-bottom:3rem;line-height:1.6;'>
        Transform your HR data into a living, interactive org chart. Upload a CSV or Excel file with employee and manager data to get started.
      </p>
      <div style='background:#fff;border:2px dashed #cbd5e1;border-radius:18px;padding:2rem 3rem;color:#94a3b8;font-size:0.95rem;'>
        ⬅️ Upload your file in the sidebar to begin
      </div>
      <div style='margin-top:2.5rem;display:flex;gap:1.5rem;flex-wrap:wrap;justify-content:center;'>
        <div style='background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:1rem 1.5rem;font-size:0.85rem;color:#475569;min-width:160px;'>
          <div style='font-size:1.5rem;margin-bottom:4px;'>📊</div><b>Required columns</b><br>Employee Code, Employee Name,<br>L1 Manager Code, Designation
        </div>
        <div style='background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:1rem 1.5rem;font-size:0.85rem;color:#475569;min-width:160px;'>
          <div style='font-size:1.5rem;margin-bottom:4px;'>✨</div><b>Optional columns</b><br>Grade, Sub Function,<br>Employment Type, Status
        </div>
        <div style='background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:1rem 1.5rem;font-size:0.85rem;color:#475569;min-width:160px;'>
          <div style='font-size:1.5rem;margin-bottom:4px;'>🖊️</div><b>Draft Mode</b><br>Right-click to move, replace,<br>or delete anyone
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

else:
    # ── Filter Data ───────────────────────────────────────────────────────────
    fdf = df.copy()

    if not include_retainers and 'Employment Type' in fdf.columns:
        fdf = fdf[~fdf['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
    if not include_inactive and 'Status' in fdf.columns:
        fdf = fdf[~fdf['Status'].astype(str).str.contains('Inactive', case=False, na=False)]
    if selected_sub != "All" and 'Sub Function' in fdf.columns:
        fdf = fdf[fdf['Sub Function'] == selected_sub]

    valid_ids = set(str(v).strip() for v in fdf['Clean_Emp_Code' if 'Clean_Emp_Code' in fdf.columns else 'Employee Code'].tolist())

    # ── Build org_data JSON ───────────────────────────────────────────────────
    id_col = 'Clean_Emp_Code' if 'Clean_Emp_Code' in fdf.columns else 'Employee Code'
    org_data = []
    for _, row in fdf.iterrows():
        emp_id = str(row.get(id_col, '')).strip()
        mgr_id = str(row.get('L1 Manager Code', '')).strip()
        if mgr_id not in valid_ids:
            mgr_id = ""
        org_data.append({
            "id": emp_id,
            "manager": mgr_id,
            "name": str(row.get('Employee Name', '')).strip(),
            "title": str(row.get('Designation', '')).strip(),
            "grade": str(row.get('Grade', 'N/A')).strip(),
            "sub": str(row.get('Sub Function', '')).strip(),
        })

    data_json = json.dumps(org_data)
    title_text = selected_sub if selected_sub != "All" else "Full Organization"

    col1, col2 = st.columns([6, 1])
    with col1:
        st.title(f"Org Architecture — {title_text}")
    with col2:
        st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)
        node_count = len(org_data)
        st.markdown(f"<div style='background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:8px 14px;text-align:center;font-size:0.85rem;color:#475569;'><b style='color:#0a0a0f;font-size:1.1rem;'>{node_count}</b><br>nodes</div>", unsafe_allow_html=True)

    # ── HTML: Standard View ───────────────────────────────────────────────────
    if not enable_draft:
        js_rows = []
        for row in org_data:
            sub_short = row['sub'][:14] if row['sub'] else "—"
            box_html = (
                f"<div class='ocard'>"
                f"<div class='ocard-top'>"
                f"<span class='obadge'>{sub_short}</span>"
                f"<span class='ograde'>GR {row['grade']}</span>"
                f"</div>"
                f"<div class='ocard-body'>"
                f"<div class='oname'>{row['name']}</div>"
                f"<div class='otitle'>{row['title']}</div>"
                f"</div>"
                f"<div class='ocard-foot'>ID: {row['id']}</div>"
                f"</div>"
            )
            clean = box_html.replace('\n','').replace("'","\\'")
            mgr = f"'{row['manager']}'" if row['manager'] else "''"
            js_rows.append(f"[{{'v':'{row['id']}','f':\"{clean}\"}},{mgr},'']")

        rows_str = ",\n".join(js_rows)

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script src="https://www.gstatic.com/charts/loader.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Syne:wght@700;800&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: #f0f2f5; font-family: 'DM Sans', sans-serif; padding: 16px; overflow: hidden; height: 100vh; display: flex; flex-direction: column; }}

  .toolbar {{ display: flex; align-items: center; gap: 10px; margin-bottom: 14px; flex-shrink: 0; }}
  .tbtn {{ background: #fff; border: 1px solid #e2e8f0; color: #374151; padding: 8px 18px; border-radius: 8px; font-weight: 600; font-size: 13px; cursor: pointer; transition: all .18s; font-family:'DM Sans',sans-serif; }}
  .tbtn:hover {{ background: #f8fafc; transform: translateY(-1px); box-shadow: 0 3px 8px rgba(0,0,0,.06); }}
  .tbtn-primary {{ background: linear-gradient(135deg,#6366f1,#4f46e5); color: #fff; border-color: transparent; box-shadow: 0 4px 12px rgba(99,102,241,.25); }}

  #scroll-outer {{
    flex: 1;
    overflow: auto;
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 32px;
    min-height: 0;
  }}
  #chart_div {{ display: inline-block; min-width: 100%; }}

  /* Google Charts connector lines */
  .google-visualization-orgchart-lineleft,
  .google-visualization-orgchart-lineright,
  .google-visualization-orgchart-linebottom,
  .google-visualization-orgchart-linetop {{
    border-color: #cbd5e1 !important;
    border-width: 2px !important;
  }}

  .myNode {{ border: none !important; background: transparent !important; padding: 0 !important; box-shadow: none !important; margin: 10px 16px !important; }}

  /* Cards */
  @keyframes fadeUp {{ from{{opacity:0;transform:translateY(10px)}} to{{opacity:1;transform:translateY(0)}} }}
  .ocard {{
    width: 220px;
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 6px rgba(0,0,0,.04);
    animation: fadeUp .3s ease-out;
    transition: box-shadow .2s, transform .2s, border-color .2s;
    border-top: 3px solid #6366f1;
    cursor: pointer;
  }}
  .ocard:hover {{ box-shadow: 0 10px 24px rgba(0,0,0,.09); transform: translateY(-3px); border-color: #818cf8; }}
  .selectedNode .ocard {{ box-shadow: 0 0 0 3px rgba(99,102,241,.25); border-color: #6366f1; }}
  .ocard-top {{ background: #fafbff; border-bottom: 1px solid #f1f5f9; padding: 9px 12px; display: flex; justify-content: space-between; align-items: center; }}
  .obadge {{ background: #eef2ff; color: #4f46e5; font-size: 9.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; padding: 3px 7px; border-radius: 4px; }}
  .ograde {{ color: #94a3b8; font-size: 10px; font-weight: 600; }}
  .ocard-body {{ padding: 11px 12px; }}
  .oname {{ font-size: 13.5px; font-weight: 700; color: #0f172a; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .otitle {{ font-size: 11px; color: #64748b; line-height: 1.4; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .ocard-foot {{ background: #fafbff; border-top: 1px solid #f1f5f9; padding: 6px 12px; font-size: 10.5px; color: #94a3b8; }}

  ::-webkit-scrollbar {{ width: 8px; height: 8px; }}
  ::-webkit-scrollbar-track {{ background: #f1f5f9; border-radius: 8px; }}
  ::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 8px; }}
</style>
</head>
<body>
  <div class="toolbar">
    <button class="tbtn tbtn-primary" id="btn-png">⬇ Download PNG</button>
  </div>
  <div id="scroll-outer">
    <div id="chart_div"></div>
  </div>
<script>
google.charts.load('current', {{packages:['orgchart']}});
google.charts.setOnLoadCallback(function() {{
  var data = new google.visualization.DataTable();
  data.addColumn('string','Name');
  data.addColumn('string','Manager');
  data.addColumn('string','ToolTip');
  data.addRows([{rows_str}]);
  var chart = new google.visualization.OrgChart(document.getElementById('chart_div'));
  chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'medium', nodeClass:'myNode', selectedNodeClass:'selectedNode'}});
}});

document.getElementById('btn-png').addEventListener('click', function() {{
  html2canvas(document.getElementById('chart_div'), {{backgroundColor:'#fff', scale:2}}).then(function(canvas){{
    canvas.toBlob(function(blob){{saveAs(blob,'org_chart.png');}});
  }});
}});
</script>
</body>
</html>"""

        components.html(html, height=820, scrolling=False)

    # ── HTML: Draft / Edit Mode ───────────────────────────────────────────────
    else:
        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Syne:wght@700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'DM Sans', sans-serif;
    background: #f0f2f5;
    padding: 16px;
    overflow: hidden;
    height: 100vh;
    display: flex;
    flex-direction: column;
  }}

  /* ── Toolbar ── */
  .toolbar {{
    display: flex; gap: 10px; margin-bottom: 14px; flex-shrink: 0; align-items: center; flex-wrap: wrap;
  }}
  .tbtn {{
    background: #fff; border: 1px solid #e2e8f0; color: #374151;
    padding: 8px 16px; border-radius: 8px; font-weight: 600; font-size: 13px;
    cursor: pointer; transition: all .18s; font-family:'DM Sans',sans-serif;
  }}
  .tbtn:hover {{ background: #f8fafc; transform: translateY(-1px); box-shadow: 0 3px 8px rgba(0,0,0,.06); }}
  .tbtn-primary {{ background: linear-gradient(135deg,#6366f1,#4f46e5); color:#fff; border-color:transparent; box-shadow: 0 4px 12px rgba(99,102,241,.25); }}
  .tbtn-danger  {{ background: #fff1f2; border-color: #fecdd3; color: #e11d48; }}
  .hint {{ margin-left: auto; font-size: 12px; color: #94a3b8; }}

  /* ── Scroll canvas ── */
  #canvas-wrap {{
    flex: 1;
    overflow: auto;
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 40px 40px 80px;
    min-height: 0;
    position: relative;
  }}
  #tree-root {{ display: inline-block; min-width: max-content; }}

  /* ── Tree layout ── */
  .tree-node {{ display: inline-flex; flex-direction: column; align-items: center; }}
  .tree-children {{
    display: flex; flex-direction: row; justify-content: center;
    gap: 0; margin-top: 0; position: relative;
  }}
  .child-col {{
    display: flex; flex-direction: column; align-items: center;
    position: relative; padding: 0 16px;
  }}

  /* Connector lines via pseudo */
  .tree-node > .tree-children::before {{
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 0;
    border-top: 2px solid #cbd5e1;
    margin: 0 50%;
    width: 0;
  }}
  .child-col::before {{
    content: ''; position: absolute;
    top: -30px; left: 50%; width: 2px; height: 30px;
    background: #cbd5e1;
    transform: translateX(-50%);
  }}
  .connector-down {{
    width: 2px; height: 30px; background: #cbd5e1; flex-shrink: 0;
  }}
  .connector-h {{
    position: absolute; top: -30px; left: 0; right: 0;
    height: 0; border-top: 2px solid #cbd5e1;
  }}

  /* ── Cards ── */
  @keyframes fadeUp {{ from{{opacity:0;transform:translateY(8px)}} to{{opacity:1;transform:translateY(0)}} }}
  .card {{
    width: 210px;
    background: #fff;
    border: 1.5px solid #e2e8f0;
    border-top: 3px solid #6366f1;
    border-radius: 11px;
    overflow: hidden;
    box-shadow: 0 2px 6px rgba(0,0,0,.04);
    animation: fadeUp .25s ease-out;
    transition: box-shadow .2s, transform .2s;
    cursor: context-menu;
    user-select: none;
  }}
  .card:hover {{ box-shadow: 0 8px 20px rgba(0,0,0,.08); transform: translateY(-2px); }}
  .card.is-root {{ border-top-color: #f59e0b; }}
  .card.is-modified {{ border-top-color: #10b981; }}
  .card.highlight {{ box-shadow: 0 0 0 3px rgba(99,102,241,.3) !important; border-color:#6366f1 !important; }}

  .card-top {{ background: #fafbff; border-bottom: 1px solid #f1f5f9; padding: 8px 10px; display: flex; justify-content: space-between; align-items: center; }}
  .badge {{ background: #eef2ff; color: #4338ca; font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing:.05em; padding: 3px 6px; border-radius: 4px; }}
  .grade {{ color: #94a3b8; font-size: 10px; font-weight: 600; }}
  .card-body {{ padding: 10px 12px; }}
  .cname {{ font-size: 13px; font-weight: 700; color: #0f172a; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .ctitle {{ font-size: 10.5px; color: #64748b; line-height: 1.4; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-top: 2px; }}
  .card-foot {{ background: #fafbff; border-top: 1px solid #f1f5f9; padding: 5px 12px; font-size: 10px; color: #94a3b8; display: flex; justify-content: space-between; }}

  /* ── Context Menu ── */
  #ctx-menu {{
    display: none; position: fixed; z-index: 9999;
    background: #fff; border: 1px solid #e2e8f0;
    border-radius: 12px; box-shadow: 0 8px 30px rgba(0,0,0,.13);
    padding: 6px; min-width: 200px;
    animation: fadeUp .12s ease-out;
  }}
  .ctx-header {{
    padding: 8px 12px; font-size: 12px; font-weight: 700; color: #0f172a;
    border-bottom: 1px solid #f1f5f9; margin-bottom: 4px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 190px;
  }}
  .ctx-item {{
    display: flex; align-items: center; gap: 8px;
    padding: 8px 12px; border-radius: 7px;
    font-size: 13px; color: #374151; cursor: pointer;
    transition: background .12s;
  }}
  .ctx-item:hover {{ background: #f1f5f9; }}
  .ctx-item.danger {{ color: #e11d48; }}
  .ctx-item.danger:hover {{ background: #fff1f2; }}
  .ctx-sep {{ height: 1px; background: #f1f5f9; margin: 4px 0; }}

  /* ── Modal ── */
  .modal-overlay {{
    display: none; position: fixed; inset: 0;
    background: rgba(0,0,0,.35); z-index: 10000;
    align-items: center; justify-content: center;
  }}
  .modal-overlay.open {{ display: flex; }}
  .modal {{
    background: #fff; border-radius: 16px; padding: 24px;
    width: 380px; max-width: 95vw; max-height: 80vh;
    overflow-y: auto;
    box-shadow: 0 20px 60px rgba(0,0,0,.18);
    animation: fadeUp .2s ease-out;
  }}
  .modal h3 {{ font-family:'Syne',sans-serif; font-size:1.1rem; font-weight:800; color:#0f172a; margin-bottom:12px; }}
  .modal p {{ font-size:13px; color:#64748b; margin-bottom:14px; }}
  .modal input {{
    width: 100%; padding: 9px 12px; border: 1.5px solid #e2e8f0;
    border-radius: 8px; font-family: 'DM Sans',sans-serif; font-size:13px;
    color: #0f172a; margin-bottom: 10px; outline: none;
    transition: border-color .18s;
  }}
  .modal input:focus {{ border-color: #6366f1; }}
  .modal-list {{ list-style: none; }}
  .modal-list li {{
    padding: 9px 12px; border-radius: 8px; cursor: pointer;
    font-size: 13px; transition: background .12s; display: flex; flex-direction: column; gap: 2px;
  }}
  .modal-list li:hover {{ background: #f1f5f9; }}
  .modal-list li .ml-name {{ font-weight: 600; color: #0f172a; }}
  .modal-list li .ml-sub {{ color: #94a3b8; font-size: 11px; }}
  .modal-actions {{ display: flex; gap: 8px; margin-top: 12px; justify-content: flex-end; }}
  .mbtn {{ padding: 8px 18px; border-radius: 8px; font-weight:600; font-size:13px; cursor:pointer; font-family:'DM Sans',sans-serif; }}
  .mbtn-cancel {{ background:#fff; border:1px solid #e2e8f0; color:#374151; }}
  .mbtn-ok {{ background: linear-gradient(135deg,#6366f1,#4f46e5); color:#fff; border:none; box-shadow: 0 4px 10px rgba(99,102,241,.25); }}

  /* ── Toast ── */
  #toast {{
    position: fixed; bottom: 20px; right: 20px;
    background: #1e293b; color: #f8fafc; padding: 11px 18px;
    border-radius: 10px; font-size: 13px; font-weight: 500;
    box-shadow: 0 6px 20px rgba(0,0,0,.2);
    transform: translateY(60px); opacity: 0;
    transition: all .25s; z-index: 99999;
  }}
  #toast.show {{ transform: translateY(0); opacity: 1; }}

  ::-webkit-scrollbar {{ width: 8px; height: 8px; }}
  ::-webkit-scrollbar-track {{ background: #f1f5f9; border-radius: 8px; }}
  ::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 8px; }}
</style>
</head>
<body>

<!-- Toolbar -->
<div class="toolbar">
  <button class="tbtn tbtn-primary" id="btn-save">💾 Download CSV</button>
  <button class="tbtn" id="btn-reset">↺ Reset All</button>
  <span class="hint">Right-click any card for options</span>
</div>

<!-- Org tree canvas -->
<div id="canvas-wrap">
  <div id="tree-root"></div>
</div>

<!-- Context Menu -->
<div id="ctx-menu">
  <div class="ctx-header" id="ctx-title">Person Name</div>
  <div class="ctx-item" id="ctx-change-mgr">🔀 Change Manager</div>
  <div class="ctx-item" id="ctx-replace">🔄 Replace Person</div>
  <div class="ctx-item" id="ctx-edit">✏️ Edit Details</div>
  <div class="ctx-sep"></div>
  <div class="ctx-item danger" id="ctx-delete">🗑 Delete from Chart</div>
</div>

<!-- Modal: Change Manager -->
<div class="modal-overlay" id="modal-mgr">
  <div class="modal">
    <h3>Change Manager</h3>
    <p id="modal-mgr-desc">Select a new manager for this person.</p>
    <input type="text" id="mgr-search" placeholder="Search by name or ID…">
    <ul class="modal-list" id="mgr-list"></ul>
    <div class="modal-actions">
      <button class="mbtn mbtn-cancel" id="mgr-cancel">Cancel</button>
    </div>
  </div>
</div>

<!-- Modal: Replace Person -->
<div class="modal-overlay" id="modal-replace">
  <div class="modal">
    <h3>Replace Person</h3>
    <p id="modal-replace-desc">Select who will replace this person (they will swap Employee Codes / managers).</p>
    <input type="text" id="replace-search" placeholder="Search by name or ID…">
    <ul class="modal-list" id="replace-list"></ul>
    <div class="modal-actions">
      <button class="mbtn mbtn-cancel" id="replace-cancel">Cancel</button>
    </div>
  </div>
</div>

<!-- Modal: Edit Details -->
<div class="modal-overlay" id="modal-edit">
  <div class="modal">
    <h3>Edit Details</h3>
    <p>Update name, title, grade, or sub-function.</p>
    <input type="text" id="edit-name" placeholder="Name">
    <input type="text" id="edit-title" placeholder="Designation / Title">
    <input type="text" id="edit-grade" placeholder="Grade">
    <input type="text" id="edit-sub" placeholder="Sub Function">
    <div class="modal-actions">
      <button class="mbtn mbtn-cancel" id="edit-cancel">Cancel</button>
      <button class="mbtn mbtn-ok" id="edit-ok">Save</button>
    </div>
  </div>
</div>

<!-- Toast -->
<div id="toast"></div>

<script>
// ─── Data ────────────────────────────────────────────────────────────────────
let orgData = JSON.parse(JSON.stringify({data_json}));
const originalData = JSON.parse(JSON.stringify({data_json}));
let modifiedIds = new Set();

// ─── Render ──────────────────────────────────────────────────────────────────
function render() {{
  const root = document.getElementById('tree-root');
  root.innerHTML = '';
  const roots = orgData.filter(n => !n.manager || !orgData.find(m => m.id === n.manager));
  if (roots.length === 0) {{ root.innerHTML = '<p style="color:#94a3b8;padding:20px;">No data to display.</p>'; return; }}
  roots.forEach(r => root.appendChild(buildNode(r)));
}}

function getChildren(id) {{
  return orgData.filter(n => n.manager === id);
}}

function buildNode(node) {{
  const wrap = document.createElement('div');
  wrap.className = 'tree-node';

  // Card
  const card = document.createElement('div');
  card.className = 'card' + (modifiedIds.has(node.id) ? ' is-modified' : '') + (!node.manager || !orgData.find(m => m.id === node.manager) ? ' is-root' : '');
  card.dataset.id = node.id;
  card.innerHTML = `
    <div class="card-top">
      <span class="badge">${{node.sub ? node.sub.substring(0,14) : '—'}}</span>
      <span class="grade">GR ${{node.grade || '—'}}</span>
    </div>
    <div class="card-body">
      <div class="cname">${{node.name}}</div>
      <div class="ctitle">${{node.title}}</div>
    </div>
    <div class="card-foot">
      <span>ID: ${{node.id.substring(0,8)}}</span>
      <span>↳ ${{getChildren(node.id).length}}</span>
    </div>
  `;
  card.addEventListener('contextmenu', (e) => {{ e.preventDefault(); openCtx(e, node.id); }});

  wrap.appendChild(card);

  // Children
  const children = getChildren(node.id);
  if (children.length > 0) {{
    const downLine = document.createElement('div');
    downLine.className = 'connector-down';
    wrap.appendChild(downLine);

    const childRow = document.createElement('div');
    childRow.className = 'tree-children';

    // Horizontal bar
    const hBar = document.createElement('div');
    hBar.style.cssText = `position:absolute;top:0;left:calc(100%/${children.length}/2);right:calc(100%/${children.length}/2);height:0;border-top:2px solid #cbd5e1;`;
    if (children.length > 1) childRow.appendChild(hBar);

    children.forEach(child => {{
      const col = document.createElement('div');
      col.className = 'child-col';
      col.appendChild(buildNode(child));
      childRow.appendChild(col);
    }});

    wrap.appendChild(childRow);
  }}

  return wrap;
}}

// ─── Context Menu ────────────────────────────────────────────────────────────
let ctxId = null;
const ctxMenu = document.getElementById('ctx-menu');

function openCtx(e, id) {{
  ctxId = id;
  const node = orgData.find(n => n.id === id);
  document.getElementById('ctx-title').textContent = node.name;

  ctxMenu.style.display = 'block';
  const x = Math.min(e.clientX, window.innerWidth - 220);
  const y = Math.min(e.clientY, window.innerHeight - 220);
  ctxMenu.style.left = x + 'px';
  ctxMenu.style.top = y + 'px';
}}

document.addEventListener('click', () => {{ ctxMenu.style.display = 'none'; }});
document.addEventListener('keydown', (e) => {{ if(e.key==='Escape') {{ ctxMenu.style.display='none'; closeAllModals(); }} }});

document.getElementById('ctx-change-mgr').addEventListener('click', () => {{
  ctxMenu.style.display = 'none';
  if (!ctxId) return;
  const node = orgData.find(n => n.id === ctxId);
  document.getElementById('modal-mgr-desc').textContent = `Select a new manager for "${{node.name}}"`;
  populateList('mgr-list', ctxId, 'manager');
  document.getElementById('modal-mgr').classList.add('open');
  document.getElementById('mgr-search').value = '';
  document.getElementById('mgr-search').focus();
}});

document.getElementById('ctx-replace').addEventListener('click', () => {{
  ctxMenu.style.display = 'none';
  if (!ctxId) return;
  const node = orgData.find(n => n.id === ctxId);
  document.getElementById('modal-replace-desc').textContent = `Who replaces "${{node.name}}"? The replacement will take over their manager and direct reports.`;
  populateList('replace-list', ctxId, 'replace');
  document.getElementById('modal-replace').classList.add('open');
  document.getElementById('replace-search').value = '';
  document.getElementById('replace-search').focus();
}});

document.getElementById('ctx-edit').addEventListener('click', () => {{
  ctxMenu.style.display = 'none';
  if (!ctxId) return;
  const node = orgData.find(n => n.id === ctxId);
  document.getElementById('edit-name').value = node.name;
  document.getElementById('edit-title').value = node.title;
  document.getElementById('edit-grade').value = node.grade;
  document.getElementById('edit-sub').value = node.sub;
  document.getElementById('modal-edit').classList.add('open');
  document.getElementById('edit-name').focus();
}});

document.getElementById('ctx-delete').addEventListener('click', () => {{
  ctxMenu.style.display = 'none';
  if (!ctxId) return;
  const node = orgData.find(n => n.id === ctxId);
  if (!confirm(`Delete "${{node.name}}" from the chart? Their direct reports will be re-parented to their manager.`)) return;

  // Re-parent children
  const mgr = node.manager;
  orgData.forEach(n => {{ if (n.manager === ctxId) {{ n.manager = mgr; modifiedIds.add(n.id); }} }});
  orgData = orgData.filter(n => n.id !== ctxId);
  modifiedIds.delete(ctxId);
  render();
  toast(`🗑 Deleted ${{node.name}}`);
}});

// ─── Populate List ────────────────────────────────────────────────────────────
function populateList(listId, excludeId, mode) {{
  const list = document.getElementById(listId);
  const searchInput = document.getElementById(listId === 'mgr-list' ? 'mgr-search' : 'replace-search');

  function draw(filter) {{
    list.innerHTML = '';
    let items = orgData.filter(n => n.id !== excludeId && !isDescendant(excludeId, n.id));
    if (filter) items = items.filter(n => n.name.toLowerCase().includes(filter) || n.id.toLowerCase().includes(filter));
    if (items.length === 0) {{ list.innerHTML = '<li style="padding:10px 12px;color:#94a3b8;font-size:13px;">No results found</li>'; return; }}
    items.slice(0, 50).forEach(n => {{
      const li = document.createElement('li');
      li.innerHTML = `<span class="ml-name">${{n.name}}</span><span class="ml-sub">${{n.title}} · ${{n.sub}}</span>`;
      li.addEventListener('click', () => {{
        if (mode === 'manager') changeManager(excludeId, n.id);
        else if (mode === 'replace') replacePerson(excludeId, n.id);
        closeAllModals();
      }});
      list.appendChild(li);
    }});
  }}

  draw('');
  searchInput.addEventListener('input', () => draw(searchInput.value.toLowerCase().trim()));
}}

// ─── Actions ──────────────────────────────────────────────────────────────────
function changeManager(empId, newMgrId) {{
  if (empId === newMgrId) return;
  if (isDescendant(empId, newMgrId)) {{ toast('⚠️ Cannot move a manager under their own report.'); return; }}
  const node = orgData.find(n => n.id === empId);
  node.manager = newMgrId;
  modifiedIds.add(empId);
  render();
  toast(`✅ Manager updated`);
}}

function replacePerson(oldId, newId) {{
  const old = orgData.find(n => n.id === oldId);
  const rep = orgData.find(n => n.id === newId);

  // Swap managers
  const oldMgr = old.manager;
  old.manager = rep.manager;
  rep.manager = oldMgr;

  // Re-point children of old to new
  orgData.forEach(n => {{
    if (n.manager === oldId && n.id !== newId) {{ n.manager = newId; modifiedIds.add(n.id); }}
  }});
  // Re-point children of new to old
  orgData.forEach(n => {{
    if (n.manager === newId && n.id !== oldId) {{ n.manager = oldId; modifiedIds.add(n.id); }}
  }});

  modifiedIds.add(oldId);
  modifiedIds.add(newId);
  render();
  toast(`🔄 ${{old.name}} ↔ ${{rep.name}} swapped`);
}}

document.getElementById('edit-ok').addEventListener('click', () => {{
  if (!ctxId) return;
  const node = orgData.find(n => n.id === ctxId);
  node.name = document.getElementById('edit-name').value.trim() || node.name;
  node.title = document.getElementById('edit-title').value.trim() || node.title;
  node.grade = document.getElementById('edit-grade').value.trim() || node.grade;
  node.sub = document.getElementById('edit-sub').value.trim() || node.sub;
  modifiedIds.add(ctxId);
  closeAllModals();
  render();
  toast('✏️ Details updated');
}});

// ─── Helpers ──────────────────────────────────────────────────────────────────
function isDescendant(ancestorId, targetId) {{
  const kids = orgData.filter(n => n.manager === ancestorId);
  for (let k of kids) {{
    if (k.id === targetId || isDescendant(k.id, targetId)) return true;
  }}
  return false;
}}

function closeAllModals() {{
  document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('open'));
}}
document.getElementById('mgr-cancel').addEventListener('click', closeAllModals);
document.getElementById('replace-cancel').addEventListener('click', closeAllModals);
document.getElementById('edit-cancel').addEventListener('click', closeAllModals);
document.querySelectorAll('.modal-overlay').forEach(m => {{
  m.addEventListener('click', (e) => {{ if(e.target === m) closeAllModals(); }});
}});

function toast(msg) {{
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 3000);
}}

// ─── Download ─────────────────────────────────────────────────────────────────
document.getElementById('btn-save').addEventListener('click', () => {{
  const headers = ['Employee Code','L1 Manager Code','Employee Name','Designation','Grade','Sub Function'];
  const rows = orgData.map(n => [n.id, n.manager, n.name, n.title, n.grade, n.sub]
    .map(v => '"' + String(v ?? '').replace(/"/g,'""') + '"').join(','));
  const csv = [headers.join(','), ...rows].join('\\n');
  saveAs(new Blob([csv], {{type:'text/csv;charset=utf-8;'}}), 'org_chart_draft.csv');
  toast('💾 CSV downloaded');
}});

document.getElementById('btn-reset').addEventListener('click', () => {{
  if (!confirm('Reset all changes?')) return;
  orgData = JSON.parse(JSON.stringify(originalData));
  modifiedIds.clear();
  render();
  toast('↺ Reset complete');
}});

// ─── Init ─────────────────────────────────────────────────────────────────────
render();
</script>
</body>
</html>"""

        components.html(html, height=840, scrolling=False)
