import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import json

# 1. Page Configuration
st.set_page_config(
    page_title="OrgDesign Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Premium Light SaaS Styling
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important; }
        #MainMenu, footer, header, .stDeployButton { visibility: hidden !important; display: none !important; }
        .stApp { background-color: #f8fafc !important; }
        .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; max-width: 99% !important; }
        
        [data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e5e7eb !important; }
        .stButton > button { border-radius: 10px !important; font-weight: 600 !important; transition: all 0.2s !important; width: 100% !important; }
        .stButton > button[kind="primary"] { background: linear-gradient(135deg, #6366f1, #8b5cf6) !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(99, 102, 241, 0.25) !important; }
        .stButton > button[kind="primary"]:hover { transform: translateY(-1px) !important; }
        .stButton > button[kind="secondary"] { background: #ffffff !important; color: #334155 !important; border: 1px solid #e5e7eb !important; }
        [data-testid="stSelectbox"] > div > div { background: #ffffff !important; border: 1px solid #e5e7eb !important; border-radius: 10px !important; color: #334155 !important; }
    </style>
""", unsafe_allow_html=True)

# 3. Initialize State
if 'draft_moves' not in st.session_state:
    st.session_state.draft_moves = {}

# 4. Sidebar
with st.sidebar:
    st.markdown("### 🏢 OrgDesign Pro")
    uploaded_file = st.file_uploader("Upload HR Roster", type=['csv', 'xlsx'], label_visibility="collapsed")
    
    df = None
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
            else: df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
            
            if 'Employee Code' in df.columns:
                df['Clean_Emp_Code'] = df['Employee Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
            if 'L1 Manager Code' in df.columns:
                df['L1 Manager Code'] = df['L1 Manager Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
        except Exception as e:
            st.error(f"Error reading file: {e}")

    st.markdown("---")
    st.markdown("##### ⚙️ Controls")
    enable_interactive = st.toggle("Interactive Draft Mode", value=False)
    
    if df is not None and 'Sub Function' in df.columns:
        subs = ["All"] + sorted(df['Sub Function'].dropna().unique())
        selected_sub = st.selectbox("Sub Function", subs, index=0)
    else:
        selected_sub = "All"

# 5. Main Logic
if df is None:
    st.info("👈 Upload an HR CSV/XLSX file in the sidebar to generate the org chart.")
else:
    # Prepare Data
    base_df = df.copy()
    
    # Apply previous draft moves if any
    for e_code, m_code in st.session_state.draft_moves.items():
        base_df.loc[base_df['Clean_Emp_Code'] == e_code, 'L1 Manager Code'] = m_code

    if selected_sub != "All" and 'Sub Function' in base_df.columns:
        base_df = base_df[base_df['Sub Function'] == selected_sub]
        
    valid_ids = set(base_df['Clean_Emp_Code'].tolist())
    
    # Build View Data (Current Tree)
    view_data = []
    for _, row in base_df.iterrows():
        eid = row.get('Clean_Emp_Code', '')
        mid = row.get('L1 Manager Code', '')
        if mid not in valid_ids: mid = ""
        
        view_data.append({
            "id": eid,
            "manager": mid,
            "name": str(row.get('Employee Name', 'Unknown')),
            "title": str(row.get('Designation', 'N/A')),
            "grade": str(row.get('Grade', 'N/A')),
            "sub": str(row.get('Sub Function', 'N/A'))
        })

    # Build All Employees List (for search/dropdown)
    all_emp_list = df[['Clean_Emp_Code', 'Employee Name', 'Designation']].dropna().copy()
    all_emp_list.columns = ['id', 'name', 'title']
    all_emp_list['id'] = all_emp_list['id'].astype(str).str.replace('.0', '', regex=False).str.strip()
    all_emps = all_emp_list.to_dict('records')

    # 6. Render HTML Component
    html_template = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interactive Org Chart</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg: #f8fafc;
            --card: #ffffff;
            --border: #e2e8f0;
            --text-main: #0f172a;
            --text-sub: #475569;
            --accent: #6366f1;
            --accent-hover: #4f46e5;
            --danger: #ef4444;
            --success: #10b981;
        }}
        * {{ box-sizing: border-box; }}
        body {{ margin: 0; padding: 0; background: var(--bg); font-family: 'Inter', sans-serif; color: var(--text-main); overflow: hidden; height: 100vh; display: flex; flex-direction: column; }}
        
        /* Header */
        .toolbar {{ height: 56px; background: #ffffff; border-bottom: 1px solid var(--border); display: flex; align-items: center; padding: 0 20px; gap: 12px; z-index: 10; flex-shrink: 0; }}
        .brand {{ font-weight: 800; font-size: 1.1rem; letter-spacing: -0.02em; margin-right: auto; }}
        .tool-btn {{ background: #ffffff; border: 1px solid var(--border); padding: 6px 12px; border-radius: 8px; cursor: pointer; font-weight: 500; font-size: 0.85rem; color: var(--text-main); transition: 0.2s; display: flex; align-items: center; gap: 6px; }}
        .tool-btn:hover {{ background: #f1f5f9; border-color: #cbd5e1; }}
        .tool-btn.primary {{ background: var(--accent); color: #fff; border: none; }}
        .tool-btn.primary:hover {{ background: var(--accent-hover); }}
        .zoom-controls {{ display: flex; align-items: center; gap: 4px; background: #f1f5f9; border-radius: 8px; padding: 2px; }}
        .zoom-controls .tool-btn {{ border: none; background: transparent; padding: 6px 10px; }}
        .zoom-controls .tool-btn:hover {{ background: #e2e8f0; }}

        /* Canvas Area */
        .canvas-wrapper {{ flex: 1; overflow: auto; background-image: radial-gradient(#e2e8f0 1px, transparent 1px); background-size: 24px 24px; position: relative; }}
        .canvas-content {{ min-width: 100%; min-height: 100%; padding: 60px; display: flex; justify-content: center; transform-origin: top center; transition: transform 0.2s ease; }}
        
        /* Org Chart Tree (CSS Only) */
        .tree ul {{ padding-top: 20px; position: relative; transition: all 0.5s; display: flex; justify-content: center; list-style: none; margin: 0; padding-left: 0; }}
        .tree li {{ display: table-cell; vertical-align: top; text-align: center; position: relative; padding: 20px 8px 0 8px; transition: all 0.5s; }}
        .tree li::before, .tree li::after {{ content: ''; position: absolute; top: 0; right: 50%; border-top: 2px solid #cbd5e1; width: 50%; height: 20px; }}
        .tree li::after {{ right: auto; left: 50%; border-left: 2px solid #cbd5e1; }}
        .tree li:only-child::after, .tree li:only-child::before {{ display: none; }}
        .tree li:only-child {{ padding-top: 0; }}
        .tree li:first-child::before, .tree li:last-child::after {{ border: 0 none; }}
        .tree li:last-child::before {{ border-right: 2px solid #cbd5e1; border-radius: 0 5px 0 0; }}
        .tree li:first-child::after {{ border-radius: 5px 0 0 0; }}
        .tree ul ul::before {{ content: ''; position: absolute; top: 0; left: 50%; border-left: 2px solid #cbd5e1; width: 0; height: 20px; }}
        
        /* Cards */
        .node-card {{
            display: inline-block; width: 220px; padding: 0; background: var(--card); 
            border: 1px solid var(--border); border-radius: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.05);
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1); cursor: pointer; position: relative;
            border-top: 3px solid #10b981; text-align: left;
        }}
        .node-card:hover {{ transform: translateY(-4px); box-shadow: 0 12px 20px rgba(0,0,0,0.1); border-color: #cbd5e1; z-index: 10; }}
        .node-card.selected {{ border: 2px solid var(--accent); box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.15); }}
        .node-card.is-target {{ border: 2px solid var(--success); background: #f0fdf4; transform: scale(1.03); }}
        .node-card.is-root {{ border-top-color: var(--accent); }}
        
        .card-h {{ padding: 12px 14px; background: #f8fafc; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; border-radius: 12px 12px 0 0; }}
        .sub-tag {{ font-size: 0.6rem; font-weight: 700; text-transform: uppercase; background: #f1f5f9; color: #475569; padding: 2px 6px; border-radius: 4px; }}
        .grade-tag {{ font-size: 0.7rem; color: #94a3b8; font-weight: 600; }}
        
        .card-b {{ padding: 14px; }}
        .emp-name {{ font-size: 0.9rem; font-weight: 700; color: var(--text-main); margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .emp-title {{ font-size: 0.75rem; color: var(--text-sub); line-height: 1.4; }}
        
        .card-f {{ padding: 8px 14px; background: #ffffff; border-top: 1px solid #f1f5f9; font-size: 0.65rem; color: #94a3b8; display: flex; justify-content: space-between; border-radius: 0 0 12px 12px; }}

        /* Action Modal */
        .modal-bg {{ position: fixed; inset: 0; background: rgba(15, 23, 42, 0.4); backdrop-filter: blur(2px); opacity: 0; pointer-events: none; transition: opacity 0.3s; z-index: 50; }}
        .modal-bg.open {{ opacity: 1; pointer-events: auto; }}
        .action-sheet {{
            position: fixed; bottom: 0; left: 0; right: 0; max-width: 500px; margin: 0 auto;
            background: #ffffff; border-top: 1px solid var(--border); border-radius: 20px 20px 0 0;
            padding: 24px; box-shadow: 0 -10px 40px rgba(0,0,0,0.1); transform: translateY(100%); transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1); z-index: 51;
        }}
        .modal-bg.open .action-sheet {{ transform: translateY(0); }}
        
        .sheet-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }}
        .sheet-title {{ font-weight: 700; font-size: 1.1rem; }}
        .close-btn {{ background: none; border: none; font-size: 1.5rem; color: #64748b; cursor: pointer; padding: 4px; }}
        .close-btn:hover {{ color: #0f172a; }}
        
        .emp-summary {{ display: flex; gap: 12px; align-items: center; padding: 12px; background: #f8fafc; border-radius: 10px; margin-bottom: 16px; }}
        .avatar {{ width: 36px; height: 36px; background: #e2e8f0; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; color: #475569; }}
        
        .action-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }}
        .act-btn {{ padding: 12px; border-radius: 10px; font-weight: 600; font-size: 0.85rem; cursor: pointer; border: 1px solid var(--border); transition: 0.2s; display: flex; align-items: center; justify-content: center; gap: 8px; }}
        .act-btn:hover {{ background: #f8fafc; }}
        .act-btn.move {{ background: #eff6ff; border-color: #bfdbfe; color: #2563eb; }}
        .act-btn.move:hover {{ background: #dbeafe; }}
        .act-btn.del {{ background: #fef2f2; border-color: #fecaca; color: #dc2626; }}
        .act-btn.del:hover {{ background: #fee2e2; }}
        
        .search-box {{ margin-top: 12px; display: none; }}
        .search-box.active {{ display: block; }}
        .s-input {{ width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.85rem; outline: none; }}
        .s-input:focus {{ border-color: var(--accent); box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.1); }}
        .s-results {{ max-height: 150px; overflow-y: auto; margin-top: 8px; border: 1px solid var(--border); border-radius: 8px; background: #fff; }}
        .s-item {{ padding: 8px 12px; cursor: pointer; font-size: 0.8rem; border-bottom: 1px solid #f1f5f9; }}
        .s-item:hover {{ background: #f8fafc; }}
        .s-item:last-child {{ border: none; }}
        
        .mode-banner {{ background: #fefce8; color: #ca8a04; padding: 8px 12px; border-radius: 8px; font-size: 0.8rem; font-weight: 500; margin-bottom: 12px; text-align: center; display: none; border: 1px solid #fef08a; }}
        .mode-banner.show {{ display: block; }}
    </style>
    </head>
    <body>
        <div class="toolbar">
            <div class="brand">🏢 OrgDesign Pro</div>
            <div class="zoom-controls">
                <button class="tool-btn" onclick="zoom(-0.1)">−</button>
                <span style="font-size:0.7rem; color:#64748b; min-width:30px; text-align:center;" id="zoom-level">100%</span>
                <button class="tool-btn" onclick="zoom(0.1)">+</button>
                <button class="tool-btn" onclick="resetZoom()">Fit</button>
            </div>
            <button class="tool-btn primary" id="dl-btn">💾 Download CSV</button>
        </div>

        <div class="canvas-wrapper" id="canvas-wrapper">
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
                        <div style="font-weight:700; color:#0f172a" id="sel-name">Name</div>
                        <div style="font-size:0.8rem; color:#475569" id="sel-title">Title</div>
                    </div>
                </div>
                <div class="action-grid">
                    <button class="act-btn move" id="btn-move" onclick="startMoveMode()">🔄 Reassign Manager</button>
                    <button class="act-btn del" onclick="removeEmployee()">🗑️ Remove Node</button>
                </div>
                <div class="search-box" id="search-box">
                    <input type="text" class="s-input" id="search-input" placeholder="Search manager by name...">
                    <div class="s-results" id="search-results"></div>
                </div>
            </div>
        </div>

    <script>
        const viewData = {json.dumps(view_data)};
        const allEmps = {json.dumps(all_emps)};
        
        let state = {{
            selected: null,
            mode: 'normal', // 'normal', 'selecting_manager'
            zoom: 1
        }};

        const treeEl = document.getElementById('org-tree');
        const canvasContent = document.getElementById('canvas-content');
        const modalBg = document.getElementById('modal-bg');
        const searchBox = document.getElementById('search-box');
        const searchInput = document.getElementById('search-input');
        const resultsDiv = document.getElementById('search-results');
        const modeBanner = document.getElementById('mode-banner');
        const zoomDisplay = document.getElementById('zoom-level');

        function render() {{
            treeEl.innerHTML = '';
            const roots = viewData.filter(d => !d.manager || d.manager === '');
            const ul = document.createElement('ul');
            roots.forEach(r => ul.appendChild(createNode(r)));
            treeEl.appendChild(ul);
        }}

        function createNode(node) {{
            const li = document.createElement('li');
            const card = document.createElement('div');
            card.className = 'node-card';
            if (!node.manager) card.classList.add('is-root');
            if (state.selected === node.id) card.classList.add('selected');
            if (state.mode === 'selecting_manager') card.style.cursor = 'crosshair';
            
            const reports = viewData.filter(x => x.manager === node.id).length;
            card.innerHTML = `
                <div class="card-h">
                    <span class="sub-tag">${{node.sub}}</span>
                    <span class="grade-tag">GR: ${{node.grade}}</span>
                </div>
                <div class="card-b">
                    <div class="emp-name">${{node.name}}</div>
                    <div class="emp-title">${{node.title}}</div>
                </div>
                <div class="card-f">
                    <span>ID: ${{node.id.substring(0,6)}}</span>
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
                // Assign manager
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
            document.getElementById('sel-avatar').textContent = node.name.charAt(0).toUpperCase();
            modalBg.classList.add('open');
        }}

        function closeModal() {{
            modalBg.classList.remove('open');
            state.selected = null;
            state.mode = 'normal';
            searchBox.classList.remove('active');
            modeBanner.classList.remove('show');
            document.getElementById('btn-move').innerHTML = '🔄 Reassign Manager';
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
            if (empId === newMgrId) return;
            const emp = viewData.find(x => x.id === empId);
            if (!emp) return;
            
            emp.manager = newMgrId;
            closeModal();
            render();
            showToast('✅ Manager updated successfully');
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
            const q = e.target.value.toLowerCase();
            resultsDiv.innerHTML = '';
            if(!q) return;
            
            allEmps.filter(emp => 
                emp.name.toLowerCase().includes(q) && 
                emp.id !== state.selected
            ).slice(0, 10).forEach(emp => {{
                const div = document.createElement('div');
                div.className = 's-item';
                div.innerHTML = `<b>${{emp.name}}</b> <span style="color:#64748b">${{emp.title}}</span>`;
                div.onclick = () => {{
                    executeMove(state.selected, emp.id);
                }};
                resultsDiv.appendChild(div);
            }});
        }});

        // Zoom
        function zoom(delta) {{
            state.zoom = Math.max(0.3, Math.min(1.5, state.zoom + delta));
            canvasContent.style.transform = `scale(${{state.zoom}})`;
            zoomDisplay.textContent = `${{Math.round(state.zoom * 100)}}%`;
        }}
        function resetZoom() {{
            state.zoom = 1;
            canvasContent.style.transform = `scale(1)`;
            zoomDisplay.textContent = '100%';
        }}

        // Download
        document.getElementById('dl-btn').onclick = () => {{
            const headers = ['Employee Code', 'L1 Manager Code', 'Employee Name', 'Designation', 'Grade', 'Sub Function'];
            const rows = viewData.map(r => [r.id, r.manager, r.name, r.title, r.grade, r.sub]);
            let csv = headers.join(',') + '\\n';
            rows.forEach(r => {{
                csv += r.map(v => `"$ {{String(v || '').replace(/"/g, '""')}}"`).join(',') + '\\n';
            }});
            
            const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = 'org_draft_updated.csv'; a.click();
            URL.revokeObjectURL(url);
        }};

        function showToast(msg) {{
            const t = document.createElement('div');
            t.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#0f172a;color:#fff;padding:10px 20px;border-radius:8px;font-size:0.85rem;font-weight:500;z-index:999;transform:translateY(20px);opacity:0;transition:0.3s;';
            t.textContent = msg;
            document.body.appendChild(t);
            requestAnimationFrame(() => {{ t.style.transform = 'translateY(0)'; t.style.opacity = '1'; }});
            setTimeout(() => {{ t.style.opacity = '0'; t.style.transform = 'translateY(20px)'; setTimeout(()=>t.remove(), 300); }}, 2500);
        }}

        render();
    </script>
    </body>
    </html>
    """
    
    components.html(html_template, height=850)
