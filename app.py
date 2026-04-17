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
# 2) Streamlit Styling (Sidebar / Buttons)
# =========================================================
st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

      html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
      }

      #MainMenu, footer, header, .stDeployButton { display: none !important; }

      .stApp { background-color: #f8fafc !important; }
      .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 99% !important;
      }

      [data-testid="stSidebar"] {
        background: #ffffff !important;
        border-right: 1px solid #e5e7eb !important;
      }

      .stButton > button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        transition: all 0.2s !important;
        width: 100% !important;
      }

      .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #4f46e5, #7c3aed) !important;
        color: #ffffff !important;
        border: none !important;
        box-shadow: 0 6px 14px rgba(79, 70, 229, 0.25) !important;
      }

      .stButton > button[kind="secondary"] {
        background: #ffffff !important;
        color: #111827 !important;   /* higher contrast */
        border: 1px solid #e5e7eb !important;
      }

      [data-testid="stSelectbox"] > div > div {
        background: #ffffff !important;
        border: 1px solid #e5e7eb !important;
        border-radius: 10px !important;
        color: #111827 !important;   /* higher contrast */
        font-weight: 600 !important;
      }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# 3) Initialize State
# =========================================================
if "draft_moves" not in st.session_state:
    st.session_state.draft_moves = {}

# =========================================================
# 4) Sidebar (Upload + Filter)
# =========================================================
with st.sidebar:
    st.markdown("### 🏢 OrgDesign Pro")

    uploaded_file = st.file_uploader(
        "Upload HR Roster",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed"
    )

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
                    df["Employee Code"]
                    .astype(str)
                    .str.replace(".0", "", regex=False)
                    .str.strip()
                )
            else:
                st.error("Missing required column: 'Employee Code'")

            if "L1 Manager Code" in df.columns:
                df["L1 Manager Code"] = (
                    df["L1 Manager Code"]
                    .astype(str)
                    .str.replace(".0", "", regex=False)
                    .str.strip()
                )
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

    st.caption("Tip: Drag inside the org area to pan horizontally/vertically. Use Center to jump to middle.")

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

    # Apply stored draft moves if any
    if "Clean_Emp_Code" in base_df.columns and "L1 Manager Code" in base_df.columns:
        for e_code, m_code in st.session_state.draft_moves.items():
            base_df.loc[base_df["Clean_Emp_Code"] == e_code, "L1 Manager Code"] = m_code

    # Filter sub-function
    if selected_sub != "All" and "Sub Function" in base_df.columns:
        base_df = base_df[base_df["Sub Function"] == selected_sub]

    valid_ids = set(base_df["Clean_Emp_Code"].dropna().astype(str).tolist())

    # Build view data for chart
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

    # Build list for search
    emp_cols = ["Clean_Emp_Code", "Employee Name"]
    if "Designation" in df.columns:
        emp_cols.append("Designation")

    all_emp_list = df[emp_cols].dropna().copy()
    if "Designation" not in all_emp_list.columns:
        all_emp_list["Designation"] = "N/A"

    all_emp_list.columns = ["id", "name", "title"]
    all_emp_list["id"] = (
        all_emp_list["id"].astype(str).str.replace(".0", "", regex=False).str.strip()
    )
    all_emps = all_emp_list.to_dict("records")

    # =========================================================
    # 6) HTML Component (High-contrast + horizontal pan + center)
    # =========================================================
    html_template = f"""
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>OrgDesign Pro</title>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap">

  <style>
    :root {{
      --bg: #f8fafc;
      --card: #ffffff;
      --border: #e5e7eb;

      /* FIX: REAL contrast (no pale greys) */
      --text-main: #0b1220;
      --text-sub:  #111827;

      --accent: #4f46e5;
      --accent-hover: #4338ca;
      --danger: #dc2626;
      --success: #059669;
    }}

    * {{ box-sizing: border-box; }}
    html, body {{ height: 100%; width: 100%; }}
    body {{
      margin: 0; padding: 0;
      background: var(--bg);
      font-family: 'Inter', sans-serif;
      color: var(--text-main);
      overflow: hidden;
      height: 100vh;
      display: flex;
      flex-direction: column;
    }}

    /* Toolbar */
    .toolbar {{
      height: 56px;
      background: #ffffff;
      border-bottom: 1px solid var(--border);
      display: flex;
      align-items: center;
      padding: 0 16px;
      gap: 10px;
      flex-shrink: 0;
    }}
    .brand {{
      font-weight: 900;
      font-size: 1.05rem;
      letter-spacing: -0.02em;
      margin-right: auto;
      color: var(--text-main);
    }}

    .tool-btn {{
      background: #ffffff;
      border: 1px solid var(--border);
      padding: 7px 12px;
      border-radius: 10px;
      cursor: pointer;
      font-weight: 800;
      font-size: 0.85rem;
      color: var(--text-main);
      transition: 0.2s;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      user-select: none;
    }}
    .tool-btn:hover {{
      background: #f1f5f9;
      border-color: #cbd5e1;
    }}
    .tool-btn.primary {{
      background: var(--accent);
      color: #fff;
      border: none;
      box-shadow: 0 8px 18px rgba(79,70,229,0.22);
    }}
    .tool-btn.primary:hover {{ background: var(--accent-hover); }}

    .zoom-controls {{
      display: flex;
      align-items: center;
      gap: 6px;
      background: #f1f5f9;
      border-radius: 12px;
      padding: 4px;
      border: 1px solid #e2e8f0;
    }}
    .zoom-label {{
      font-size: 0.8rem;
      color: var(--text-main);
      min-width: 48px;
      text-align: center;
      font-weight: 900;
    }}

    /* Canvas Wrapper: REAL horizontal scroll + panning */
    .canvas-wrapper {{
      flex: 1;
      overflow: auto;                  /* FIX: allows horizontal + vertical scroll */
      -webkit-overflow-scrolling: touch;
      background-image: radial-gradient(#e2e8f0 1px, transparent 1px);
      background-size: 24px 24px;
      position: relative;
      cursor: grab;
    }}
    .canvas-wrapper.dragging {{ cursor: grabbing; }}

    /* Canvas content must size to chart width, not viewport */
    .canvas-content {{
      padding: 40px 40px 80px 40px;
      transform-origin: top left;      /* FIX: zoom doesn't push content off-screen */
      transition: transform 0.18s ease;
      width: max-content;             /* FIX: enables true horizontal scroll width */
      min-width: 100%;
    }}

    /* Scrollbars visible */
    .canvas-wrapper::-webkit-scrollbar {{ width: 10px; height: 10px; }}
    .canvas-wrapper::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 10px; }}
    .canvas-wrapper::-webkit-scrollbar-track {{ background: #f1f5f9; }}

    /* Tree Layout */
    .tree {{
      display: inline-block;          /* key: shrink-wrap tree size */
    }}
    .tree ul {{
      padding-top: 18px;
      position: relative;
      list-style: none;
      margin: 0;
      padding-left: 0;

      display: flex;
      justify-content: center;
      width: max-content;             /* key: chart grows naturally */
      margin-left: auto;
      margin-right: auto;
    }}
    .tree li {{
      display: table-cell;
      vertical-align: top;
      text-align: center;
      position: relative;
      padding: 18px 6px 0 6px;        /* tighter spacing reduces "span is big" feel */
    }}

    /* connectors */
    .tree li::before, .tree li::after {{
      content: '';
      position: absolute;
      top: 0;
      right: 50%;
      border-top: 2px solid #cbd5e1;
      width: 50%;
      height: 18px;
    }}
    .tree li::after {{
      right: auto;
      left: 50%;
      border-left: 2px solid #cbd5e1;
    }}
    .tree li:only-child::after, .tree li:only-child::before {{ display: none; }}
    .tree li:only-child {{ padding-top: 0; }}
    .tree li:first-child::before, .tree li:last-child::after {{ border: 0 none; }}
    .tree li:last-child::before {{
      border-right: 2px solid #cbd5e1;
      border-radius: 0 5px 0 0;
    }}
    .tree li:first-child::after {{ border-radius: 5px 0 0 0; }}
    .tree ul ul::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 50%;
      border-left: 2px solid #cbd5e1;
      width: 0;
      height: 18px;
    }}

    /* Cards */
    .node-card {{
      display: inline-block;
      width: 220px;                   /* slightly smaller to reduce width */
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 14px;
      box-shadow: 0 2px 7px rgba(0,0,0,0.06);
      transition: 0.18s;
      cursor: pointer;
      text-align: left;
      border-top: 4px solid var(--success);
      color: var(--text-main);
    }}
    .node-card:hover {{
      transform: translateY(-3px);
      box-shadow: 0 14px 26px rgba(0,0,0,0.12);
      border-color: #cbd5e1;
      z-index: 10;
    }}
    .node-card.selected {{
      border: 2px solid var(--accent);
      box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.18);
    }}
    .node-card.is-root {{ border-top-color: var(--accent); }}

    .card-h {{
      padding: 10px 12px;
      background: #f8fafc;
      border-bottom: 1px solid var(--border);
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-radius: 14px 14px 0 0;
      gap: 10px;
    }}

    /* FIX: not light grey */
    .sub-tag {{
      font-size: 0.62rem;
      font-weight: 900;
      text-transform: uppercase;
      background: #e2e8f0;
      color: #0b1220;
      padding: 3px 10px;
      border-radius: 999px;
      max-width: 140px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}
    .grade-tag {{
      font-size: 0.72rem;
      color: #0b1220;                /* FIX: dark */
      font-weight: 900;
      white-space: nowrap;
    }}

    .card-b {{ padding: 12px; }}

    .emp-name, .emp-title {{
      overflow: hidden;
      text-overflow: ellipsis;
      display: -webkit-box;
      -webkit-box-orient: vertical;
      color: var(--text-main);
    }}
    .emp-name {{
      font-size: 0.92rem;
      font-weight: 900;
      margin-bottom: 4px;
      -webkit-line-clamp: 1;
    }}
    .emp-title {{
      font-size: 0.82rem;
      font-weight: 700;
      color: #111827;               /* FIX: dark, readable */
      line-height: 1.3;
      -webkit-line-clamp: 2;
    }}

    .card-f {{
      padding: 9px 12px;
      border-top: 1px solid #eef2f7;
      font-size: 0.72rem;
      color: #0b1220;               /* FIX: dark */
      font-weight: 800;
      display: flex;
      justify-content: space-between;
      border-radius: 0 0 14px 14px;
    }}

    /* Modal */
    .modal-bg {{
      position: fixed;
      inset: 0;
      background: rgba(15, 23, 42, 0.45);
      backdrop-filter: blur(2px);
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.25s;
      z-index: 50;
    }}
    .modal-bg.open {{ opacity: 1; pointer-events: auto; }}

    .action-sheet {{
      position: fixed;
      bottom: 0; left: 0; right: 0;
      max-width: 560px;
      margin: 0 auto;
      background: #ffffff;
      border-top: 1px solid var(--border);
      border-radius: 20px 20px 0 0;
      padding: 22px;
      box-shadow: 0 -10px 40px rgba(0,0,0,0.12);
      transform: translateY(100%);
      transition: transform 0.28s cubic-bezier(0.16, 1, 0.3, 1);
      z-index: 51;
    }}
    .modal-bg.open .action-sheet {{ transform: translateY(0); }}

    .sheet-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 14px;
    }}
    .sheet-title {{
      font-weight: 900;
      font-size: 1.05rem;
      color: var(--text-main);
    }}
    .close-btn {{
      background: none;
      border: none;
      font-size: 1.6rem;
      color: #111827;
      cursor: pointer;
      padding: 4px;
      font-weight: 900;
    }}

    .emp-summary {{
      display: flex;
      gap: 12px;
      align-items: center;
      padding: 12px;
      background: #f8fafc;
      border-radius: 14px;
      margin-bottom: 14px;
      border: 1px solid #e5e7eb;
    }}
    .avatar {{
      width: 40px;
      height: 40px;
      background: #e2e8f0;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 900;
      color: #0b1220;
    }}

    .action-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }}
    .act-btn {{
      padding: 12px;
      border-radius: 14px;
      font-weight: 900;
      font-size: 0.86rem;
      cursor: pointer;
      border: 1px solid var(--border);
      transition: 0.18s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      color: var(--text-main);
    }}
    .act-btn:hover {{ background: #f8fafc; }}
    .act-btn.move {{
      background: #eff6ff;
      border-color: #bfdbfe;
      color: #1d4ed8;
    }}
    .act-btn.del {{
      background: #fef2f2;
      border-color: #fecaca;
      color: #b91c1c;
    }}

    .search-box {{ margin-top: 12px; display: none; }}
    .search-box.active {{ display: block; }}
    .s-input {{
      width: 100%;
      padding: 10px 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      font-size: 0.92rem;
      outline: none;
      font-weight: 700;
      color: var(--text-main);
    }}
    .s-input:focus {{
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.12);
    }}
    .s-results {{
      max-height: 170px;
      overflow-y: auto;
      margin-top: 8px;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: #fff;
    }}
    .s-item {{
      padding: 10px 12px;
      cursor: pointer;
      font-size: 0.88rem;
      border-bottom: 1px solid #f1f5f9;
      color: var(--text-main);
      font-weight: 700;
    }}
    .s-item:hover {{ background: #f8fafc; }}
    .s-item:last-child {{ border: none; }}

    .mode-banner {{
      background: #fffbeb;
      color: #92400e;
      padding: 10px 12px;
      border-radius: 12px;
      font-size: 0.86rem;
      font-weight: 900;
      margin-bottom: 12px;
      text-align: center;
      display: none;
      border: 1px solid #fcd34d;
    }}
    .mode-banner.show {{ display: block; }}
  </style>
</head>

<body>
  <div class="toolbar">
    <div class="brand">🏢 OrgDesign Pro</div>

    <div class="zoom-controls">
      <button class="tool-btn" onclick="zoom(-0.1)">−</button>
      <span class="zoom-label" id="zoom-level">100%</span>
      <button class="tool-btn" onclick="zoom(0.1)">+</button>
      <button class="tool-btn" onclick="resetZoom()">Fit</button>
    </div>

    <button class="tool-btn" onclick="centerView()">🧭 Center</button>
    <button class="tool-btn primary" id="dl-btn">💾 Download CSV</button>
  </div>

  <div class="canvas-wrapper" id="canvas-wrapper" title="Drag to pan • Scroll to zoom (trackpad) • Shift+wheel for horizontal on mouse">
    <div class="canvas-content" id="canvas-content">
      <div class="tree" id="org-tree"></div>
    </div>
  </div>

  <div class="modal-bg" id="modal-bg" onclick="closeModal()">
    <div class="action-sheet" onclick="event.stopPropagation()">
      <div class="mode-banner" id="mode-banner">👆 Click any manager in the tree to assign them</div>

      <div class="sheet-header">
        <span class="sheet-title">Manage Employee</span>
        <button class="close-btn" onclick="closeModal()">×</button>
      </div>

      <div class="emp-summary">
        <div class="avatar" id="sel-avatar">A</div>
        <div>
          <div style="font-weight:900; color:#0b1220" id="sel-name">Name</div>
          <div style="font-size:0.9rem; color:#111827; font-weight:800" id="sel-title">Title</div>
        </div>
      </div>

      <div class="action-grid">
        <button class="act-btn move" id="btn-move" onclick="startMoveMode()">🔄 Reassign Manager</button>
        <button class="act-btn del" onclick="removeEmployee()">🗑️ Remove Node</button>
      </div>

      <div class="search-box" id="search-box">
        <input type="text" class="s-input" id="search-input" placeholder="Search manager by name..." />
        <div class="s-results" id="search-results"></div>
      </div>
    </div>
  </div>

  <script>
    const viewData = {json.dumps(view_data)};
    const allEmps = {json.dumps(all_emps)};

    let state = {{
      selected: null,
      mode: 'normal',
      zoom: 1
    }};

    const treeEl = document.getElementById('org-tree');
    const canvasContent = document.getElementById('canvas-content');
    const wrapper = document.getElementById('canvas-wrapper');

    const modalBg = document.getElementById('modal-bg');
    const searchBox = document.getElementById('search-box');
    const searchInput = document.getElementById('search-input');
    const resultsDiv = document.getElementById('search-results');
    const modeBanner = document.getElementById('mode-banner');
    const zoomDisplay = document.getElementById('zoom-level');

    function escapeHtml(str) {{
      return String(str ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
    }}

    function safeSliceId(id) {{
      return String(id ?? '').slice(0, 8);
    }}

    function render() {{
      treeEl.innerHTML = '';
      const roots = viewData.filter(d => !d.manager || d.manager === '');
      const ul = document.createElement('ul');
      roots.forEach(r => ul.appendChild(createNode(r)));
      treeEl.appendChild(ul);

      // After rendering, center view (helps Streamlit Cloud default scrollLeft=0 issue)
      // Only do it once on initial render.
      if (!window.__didCenter) {{
        window.__didCenter = true;
        setTimeout(centerView, 120);
      }}
    }}

    function createNode(node) {{
      const li = document.createElement('li');
      const card = document.createElement('div');
      card.className = 'node-card';

      if (!node.manager) card.classList.add('is-root');
      if (state.selected === node.id) card.classList.add('selected');
      if (state.mode === 'selecting_manager') card.style.cursor = 'crosshair';

      const reports = viewData.filter(x => x.manager === node.id).length;

      const nodeName = escapeHtml(node.name);
      const nodeTitle = escapeHtml(node.title);
      const nodeSub = escapeHtml(node.sub);
      const nodeGrade = escapeHtml(node.grade);

      card.innerHTML = `
        <div class="card-h">
          <span class="sub-tag" title="${{nodeSub}}">${{nodeSub}}</span>
          <span class="grade-tag">GR: ${{nodeGrade}}</span>
        </div>
        <div class="card-b">
          <div class="emp-name" title="${{nodeName}}">${{nodeName}}</div>
          <div class="emp-title" title="${{nodeTitle}}">${{nodeTitle}}</div>
        </div>
        <div class="card-f">
          <span>ID: ${{safeSliceId(node.id)}}</span>
          <span>${{reports}} Directs</span>
        </div>
      `;

      card.onclick = (e) => {{
        e.stopPropagation();
        handleNodeClick(node);
      }};

      li.appendChild(card);

      const children = viewData.filter(c => c.manager === node.id);
      if (children.length > 0) {{
        const ul = document.createElement('ul');
        children.forEach(c => ul.appendChild(createNode(c)));
        li.appendChild(ul);
      }}
      return li;
    }}

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
      document.getElementById('sel-name').textContent = node.name;
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
      searchInput.value = '';
      resultsDiv.innerHTML = '';
      render();
    }}

    function startMoveMode() {{
      state.mode = 'selecting_manager';
      modeBanner.classList.add('show');
      searchBox.classList.add('active');
      document.getElementById('btn-move').innerHTML = '⏳ Select new manager...';
      searchInput.focus();
      render();
    }}

    function executeMove(empId, newMgrId) {{
      if (!empId || !newMgrId || empId === newMgrId) return;
      const emp = viewData.find(x => x.id === empId);
      if (!emp) return;

      emp.manager = newMgrId;
      closeModal();
      render();
      showToast('✅ Manager updated');
    }}

    function removeEmployee() {{
      if(!confirm('Remove this employee from the draft? Their direct reports will move up.')) return;

      const idx = viewData.findIndex(x => x.id === state.selected);
      if (idx > -1) {{
        const node = viewData[idx];
        const children = viewData.filter(c => c.manager === node.id);
        children.forEach(c => c.manager = node.manager);
        viewData.splice(idx, 1);
      }}
      closeModal();
      render();
    }}

    // Search Logic
    searchInput.addEventListener('input', (e) => {{
      const q = e.target.value.toLowerCase().trim();
      resultsDiv.innerHTML = '';
      if(!q) return;

      allEmps
        .filter(emp =>
          String(emp.name || '').toLowerCase().includes(q) &&
          String(emp.id || '') !== String(state.selected || '')
        )
        .slice(0, 12)
        .forEach(emp => {{
          const div = document.createElement('div');
          div.className = 's-item';
          div.innerHTML = `<b>${{escapeHtml(emp.name)}}</b> <span style="color:#111827; font-weight:800">— ${{escapeHtml(emp.title)}}</span>`;
          div.onclick = () => executeMove(state.selected, emp.id);
          resultsDiv.appendChild(div);
        }});
    }});

    // Zoom
    function zoom(delta) {{
      state.zoom = Math.max(0.3, Math.min(1.8, state.zoom + delta));
      canvasContent.style.transform = `scale(${{state.zoom}})`;
      zoomDisplay.textContent = `${{Math.round(state.zoom * 100)}}%`;
    }}

    function resetZoom() {{
      state.zoom = 1;
      canvasContent.style.transform = 'scale(1)';
      zoomDisplay.textContent = '100%';
      setTimeout(centerView, 60);
    }}

    // Center the view (critical for wide org charts on Streamlit Cloud)
    function centerView() {{
      const maxLeft = wrapper.scrollWidth - wrapper.clientWidth;
      const maxTop  = wrapper.scrollHeight - wrapper.clientHeight;
      wrapper.scrollLeft = Math.max(0, Math.floor(maxLeft / 2));
      wrapper.scrollTop  = Math.max(0, Math.floor(maxTop * 0.05));
    }}

    // Drag-to-pan (mouse)
    let isDown = false;
    let startX, startY, scrollLeft, scrollTop;

    wrapper.addEventListener('mousedown', (e) => {{
      // Don't start drag when clicking on cards/buttons/modal etc.
      if (e.target.closest('.node-card') || e.target.closest('.toolbar') || e.target.closest('.action-sheet')) return;
      isDown = true;
      wrapper.classList.add('dragging');
      startX = e.pageX - wrapper.offsetLeft;
      startY = e.pageY - wrapper.offsetTop;
      scrollLeft = wrapper.scrollLeft;
      scrollTop = wrapper.scrollTop;
    }});

    window.addEventListener('mouseup', () => {{
      isDown = false;
      wrapper.classList.remove('dragging');
    }});

    wrapper.addEventListener('mousemove', (e) => {{
      if (!isDown) return;
      e.preventDefault();
      const x = e.pageX - wrapper.offsetLeft;
      const y = e.pageY - wrapper.offsetTop;
      const walkX = (x - startX);
      const walkY = (y - startY);
      wrapper.scrollLeft = scrollLeft - walkX;
      wrapper.scrollTop = scrollTop - walkY;
    }});

    // CSV Download
    function csvEscape(val) {{
      const s = String(val ?? '');
      return '"' + s.replaceAll('"', '""') + '"';
    }}

    document.getElementById('dl-btn').onclick = () => {{
      const headers = ['Employee Code', 'L1 Manager Code', 'Employee Name', 'Designation', 'Grade', 'Sub Function'];
      const rows = viewData.map(r => [r.id, r.manager, r.name, r.title, r.grade, r.sub]);

      let csv = headers.join(',') + '\\n';
      rows.forEach(r => {{
        csv += r.map(csvEscape).join(',') + '\\n';
      }});

      const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'org_draft_updated.csv';
      a.click();
      URL.revokeObjectURL(url);
    }};

    function showToast(msg) {{
      const t = document.createElement('div');
      t.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#0b1220;color:#fff;padding:10px 16px;border-radius:12px;font-size:0.92rem;font-weight:900;z-index:999;transform:translateY(20px);opacity:0;transition:0.25s;';
      t.textContent = msg;
      document.body.appendChild(t);
      requestAnimationFrame(() => {{ t.style.transform = 'translateY(0)'; t.style.opacity = '1'; }});
      setTimeout(() => {{
        t.style.opacity = '0';
        t.style.transform = 'translateY(20px)';
        setTimeout(() => t.remove(), 250);
      }}, 2000);
    }}

    render();
  </script>
</body>
</html>
"""

    # Streamlit Cloud: you NEED a large height + scrolling enabled
    # Increase height if your org is huge.
    components.html(html_template, height=1600, scrolling=True)
