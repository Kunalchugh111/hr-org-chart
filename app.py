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

        /* Base Reset */
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important; }
        
        /* Hide Branding */
        #MainMenu, footer, header, .stDeployButton { visibility: hidden !important; display: none !important; }

        /* App Background */
        .stApp {
            background-color: #fafafa !important;
            background-image: 
                radial-gradient(circle at 10% 10%, rgba(99, 102, 241, 0.04) 0%, transparent 40%),
                radial-gradient(circle at 90% 90%, rgba(16, 185, 129, 0.04) 0%, transparent 40%);
        }
        
        .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 99% !important; }
        h1 { font-weight: 800 !important; color: #0f172a !important; font-size: 2.5rem !important; letter-spacing: -0.02em !important; }

        /* Sidebar Styling */
        [data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e5e7eb !important; }
        [data-testid="stSidebar"] h3 { color: #111827 !important; border-bottom: 1px solid #f3f4f6 !important; padding-bottom: 0.5rem !important; }
        
        /* Inputs & Buttons */
        .stButton > button {
            border-radius: 10px !important; font-weight: 600 !important; transition: all 0.2s !important;
        }
        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #8b5cf6, #6366f1) !important; color: white !important; border: none !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.25) !important;
        }
        .stButton > button[kind="primary"]:hover { transform: translateY(-1px) !important; box-shadow: 0 6px 16px rgba(99, 102, 241, 0.35) !important; }
        
        .stButton > button[kind="secondary"] {
            background: #ffffff !important; color: #111827 !important; border: 1px solid #e5e7eb !important;
        }
        .stButton > button[kind="secondary"]:hover { background: #f9fafb !important; border-color: #d1d5db !important; }
        
        [data-testid="stSelectbox"] > div > div { background: #ffffff !important; border: 1px solid #e5e7eb !important; border-radius: 10px !important; }
    </style>
""", unsafe_allow_html=True)

# 3. Initialize Session State
if 'draft_moves' not in st.session_state:
    st.session_state.draft_moves = {}

# 4. Sidebar Logic
with st.sidebar:
    st.markdown("### 🏢 OrgDesign Pro")
    
    st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
    st.markdown("##### 📊 Upload HR Data")
    uploaded_file = st.file_uploader("", type=['csv', 'xlsx'], label_visibility="collapsed")
    
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
        else: df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()
        
        # Data Cleaning
        if 'Employee Code' in df.columns:
            df['Clean_Emp_Code'] = df['Employee Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
        if 'L1 Manager Code' in df.columns:
            df['L1 Manager Code'] = df['L1 Manager Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
    else:
        df = None

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### ⚙️ Settings")
    selected_sub = "All"
    enable_draft = st.toggle("Interactive Draft Mode", value=False)
    
    if df is not None:
        st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
        
        if 'Sub Function' in df.columns:
            sub_functions = sorted(df['Sub Function'].dropna().unique())
            selected_sub = st.selectbox("Sub Function View", ["All"] + sub_functions, index=0)
        
        st.markdown("</div>", unsafe_allow_html=True)

# 5. Main Logic
if df is None:
    st.info("👈 Upload an HR file to start.")
else:
    # Filter Data
    base_df = df.copy()
    
    # Apply Draft Moves
    if 'draft_moves' in st.session_state and st.session_state.draft_moves:
        for e_code, m_code in st.session_state.draft_moves.items():
            base_df.loc[base_df['Clean_Emp_Code'] == e_code, 'L1 Manager Code'] = m_code

    # Filter View
    if selected_sub != "All" and 'Sub Function' in base_df.columns:
        base_df = base_df[base_df['Sub Function'] == selected_sub]
        
    valid_ids = set(base_df['Clean_Emp_Code'].tolist())
    
    # Prepare JSON Payload
    # We only pass the current view to the chart, but we pass ALL employees for the "Reassign Manager" dropdown logic
    all_employees = []
    if 'Employee Name' in df.columns:
        for _, row in df.iterrows():
            all_employees.append({
                "id": str(row.get('Clean_Emp_Code', '')),
                "name": str(row.get('Employee Name', '')),
                "title": str(row.get('Designation', ''))
            })

    view_data = []
    for _, row in base_df.iterrows():
        emp_id = row.get('Clean_Emp_Code', '')
        mgr_id = row.get('L1 Manager Code', '')
        
        # Ensure manager is valid within this view, otherwise treat as root
        if mgr_id not in valid_ids:
            mgr_id = ""
            
        view_data.append({
            "id": emp_id,
            "manager": mgr_id,
            "name": str(row.get('Employee Name', '')),
            "title": str(row.get('Designation', '')),
            "grade": str(row.get('Grade', '')),
            "sub": str(row.get('Sub Function', ''))
        })
    
    # 6. Render HTML Component
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Org Chart</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        :root {{
            --bg: #fafafa;
            --card-bg: #ffffff;
            --primary: #6366f1;
            --primary-light: #e0e7ff;
            --text-main: #111827;
            --text-sub: #4b5563;
            --border: #e5e7eb;
            --shadow: 0 2px 8px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02);
            --shadow-hover: 0 12px 24px rgba(0,0,0,0.08), 0 4px 8px rgba(99, 102, 241, 0.12);
            --radius: 12px;
        }}
        
        body {{ 
            background: var(--bg); 
            font-family: 'Inter', sans-serif; 
            margin: 0; 
            padding: 0; 
            color: var(--text-main);
            overflow: hidden; /* Prevent body scroll, use container scroll */
        }}

        /* Layout */
        .header-bar {{
            position: fixed; top: 0; left: 0; right: 0; height: 60px; 
            background: rgba(255,255,255,0.9); backdrop-filter: blur(8px);
            border-bottom: 1px solid var(--border); z-index: 100;
            display: flex; align-items: center; padding: 0 24px; justify-content: space-between;
        }}
        .header-title {{ font-size: 1.1rem; font-weight: 700; display: flex; align-items: center; gap: 10px; }}
        .badge-draft {{ background: #d1fae5; color: #065f46; font-size: 0.7rem; padding: 2px 8px; border-radius: 10px; font-weight: 700; text-transform: uppercase; }}
        
        .main-scroll {{
            height: calc(100vh - 60px); margin-top: 60px; 
            overflow: auto; 
            background-image: 
                radial-gradient(circle at 20px 20px, #f3f4f6 1px, transparent 1px);
            background-size: 40px 40px;
        }}
        
        /* Tree Container */
        .tree-center {{ min-height: 100%; display: flex; justify-content: center; padding-top: 60px; padding-bottom: 100px; }}
        
        /* Recursive Tree Structure */
        ul {{ padding-top: 20px; position: relative; transition: all 0.5s; display: flex; justify-content: center; list-style: none; margin: 0; padding-left: 0; }}
        li {{ display: table-cell; vertical-align: top; text-align: center; position: relative; padding: 20px 5px 0 5px; }}
        
        /* Connectors */
        li::before, li::after {{
            content: ''; position: absolute; top: 0; right: 50%; border-top: 2px solid #cbd5e1;
            width: 50%; height: 20px;
        }}
        li::after {{ right: auto; left: 50%; border-left: 2px solid #cbd5e1; }}
        li:only-child::after, li:only-child::before {{ display: none; }}
        li:only-child {{ padding-top: 0; }}
        li:first-child::before, li:last-child::after {{ border: 0 none; }}
        li:last-child::before {{ border-right: 2px solid #cbd5e1; border-radius: 0 5px 0 0; }}
        li:first-child::after {{ border-radius: 5px 0 0 0; }}
        
        ul ul::before {{
            content: ''; position: absolute; top: 0; left: 50%;
            border-left: 2px solid #cbd5e1; width: 0; height: 20px;
        }}
        
        /* Card Design */
        .card {{
            display: inline-block; width: 240px; padding: 0;
            background: var(--card-bg); border: 1px solid var(--border); border-radius: var(--radius);
            box-shadow: var(--shadow); transition: all 0.2s ease; position: relative;
            cursor: pointer; z-index: 10;
        }}
        .card:hover {{ box-shadow: var(--shadow-hover); transform: translateY(-3px); z-index: 20; }}
        
        .card.selected {{
            border: 2px solid var(--primary); box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.2) !important;
        }}
        
        .card-header {{
            padding: 12px 16px; background: #f8fafc; border-bottom: 1px solid var(--border);
            display: flex; justify-content: space-between; align-items: center;
            border-radius: var(--radius) var(--radius) 0 0;
        }}
        .card-header.is-root {{ background: linear-gradient(135deg, #f5f3ff 0%, #ffffff 100%); }}
        
        .sub-badge {{
            font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
            background: #f1f5f9; color: #475569; padding: 2px 6px; border-radius: 4px;
        }}
        .grade-tag {{ font-size: 0.7rem; color: #94a3b8; font-weight: 600; }}
        
        .card-body {{ padding: 16px; text-align: left; }}
        .emp-name {{ font-size: 0.95rem; font-weight: 700; color: var(--text-main); margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .emp-title {{ font-size: 0.8rem; color: var(--text-sub); line-height: 1.4; }}
        
        .card-footer {{
            padding: 8px 16px; background: #ffffff; border-top: 1px solid #f1f5f9;
            font-size: 0.7rem; color: #94a3b8; display: flex; justify-content: space-between;
            border-radius: 0 0 var(--radius) var(--radius);
        }}
        
        /* Modal / Action Panel */
        .modal-overlay {{
            position: fixed; bottom: -100%; left: 0; right: 0; height: 100vh;
            background: rgba(0,0,0,0.2); backdrop-filter: blur(2px); z-index: 500;
            transition: background 0.3s; opacity: 0;
        }}
        .modal-overlay.active {{ bottom: 0; opacity: 1; }}
        
        .action-panel {{
            position: fixed; bottom: -400px; left: 0; right: 0; max-width: 600px;
            margin: 0 auto; background: #ffffff; border-top: 1px solid #e2e8f0;
            border-radius: 20px 20px 0 0; box-shadow: 0 -10px 40px rgba(0,0,0,0.1);
            padding: 24px; z-index: 501; transition: bottom 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }}
        .modal-overlay.active .action-panel {{ bottom: 0; }}
        
        .panel-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }}
        .panel-title {{ font-weight: 700; font-size: 1.2rem; color: #1e293b; }}
        .close-btn {{ background: none; border: none; font-size: 1.5rem; color: #64748b; cursor: pointer; padding: 4px 8px; border-radius: 6px; }}
        .close-btn:hover {{ background: #f1f5f9; color: #0f172a; }}
        
        .selected-info {{
            display: flex; align-items: center; gap: 12px; padding: 12px; background: #f8fafc;
            border-radius: 12px; border: 1px solid #e2e8f0; margin-bottom: 20px;
        }}
        .avatar {{
            width: 40px; height: 40px; border-radius: 50%; background: #e2e8f0;
            display: flex; align-items: center; justify-content: center; font-weight: 700; color: #475569;
        }}
        
        .action-btn {{
            width: 100%; padding: 12px; margin-bottom: 10px; border-radius: 10px;
            font-weight: 600; font-size: 0.9rem; cursor: pointer; transition: 0.2s;
            display: flex; align-items: center; gap: 10px;
        }}
        .btn-move {{ background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; }}
        .btn-move:hover {{ background: #dbeafe; border-color: #93c5fd; }}
        .btn-delete {{ background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }}
        .btn-delete:hover {{ background: #fee2e2; border-color: #fca5a5; }}
        
        /* Search for Manager */
        .search-container {{ display: none; flex-direction: column; gap: 8px; margin-top: 10px; }}
        .search-container.active {{ display: flex; }}
        .search-input {{
            padding: 10px; border: 1px solid #cbd5e1; border-radius: 8px; font-family: inherit;
        }}
        .search-results {{
            max-height: 150px; overflow-y: auto; background: #ffffff; border: 1px solid #e2e8f0;
            border-radius: 8px;
        }}
        .result-item {{
            padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #f1f5f9;
            font-size: 0.85rem; display: flex; justify-content: space-between;
        }}
        .result-item:hover {{ background: #f8fafc; }}
        .result-item:last-child {{ border-bottom: none; }}
        
    </style>
    </head>
    <body>
        <div class="header-bar">
            <div class="header-title">
                OrgDesign Pro <span class="badge-draft">Interactive</span>
            </div>
            <button id="download-btn" style="background:#0f172a; color:white; border:none; padding:8px 16px; border-radius:8px; cursor:pointer; font-weight:600;">
                💾 Download CSV
            </button>
        </div>

        <div class="main-scroll">
            <div class="tree-center" id="org-tree">
                <!-- Tree renders here -->
            </div>
        </div>

        <div class="modal-overlay" id="overlay" onclick="closePanel()">
            <div class="action-panel" id="panel">
                <div class="panel-header">
                    <span class="panel-title">Manage Employee</span>
                    <button class="close-btn" onclick="closePanel()">×</button>
                </div>
                <div class="selected-info" id="sel-info">
                    <div class="avatar" id="sel-avatar">A</div>
                    <div>
                        <div style="font-weight:700" id="sel-name">Name</div>
                        <div style="font-size:0.8rem; color:#64748b" id="sel-title">Title</div>
                    </div>
                </div>
                
                <button class="action-btn btn-move" id="btn-move-action" onclick="startMoveMode()">
                    🔄 Reassign Manager
                </button>
                <button class="action-btn btn-delete" onclick="deleteEmployee()">
                    🗑️ Remove Employee
                </button>

                <div class="search-container" id="search-container">
                    <input type="text" class="search-input" id="manager-search" placeholder="Search for a new manager...">
                    <div class="search-results" id="search-results"></div>
                </div>
            </div>
        </div>

    <script>
        // Data Setup
        const viewData = {json.dumps(view_data)};
        const allEmps = {json.dumps(all_employees)};
        let selectedId = null;
        let movingId = null;

        const treeEl = document.getElementById('org-tree');
        const overlay = document.getElementById('overlay');
        const panel = document.getElementById('panel');
        const searchContainer = document.getElementById('search-container');
        const searchInput = document.getElementById('manager-search');
        const resultsList = document.getElementById('search-results');

        // 1. Recursive Tree Render
        function renderTree() {{
            treeEl.innerHTML = '';
            const roots = viewData.filter(d => !d.manager || d.manager === '');
            const ul = document.createElement('ul');
            
            roots.forEach(r => {{
                ul.appendChild(createNode(r));
            }});
            treeEl.appendChild(ul);
        }}

        function createNode(node) {{
            const li = document.createElement('li');
            
            // Card
            const card = document.createElement('div');
            card.className = 'card';
            if(selectedId === node.id) card.classList.add('selected');
            
            card.innerHTML = `
                <div class="card-header ${{!node.manager ? 'is-root' : ''}}">
                    <span class="sub-badge">${{node.sub || 'N/A'}}</span>
                    <span class="grade-tag">GR: ${{node.grade}}</span>
                </div>
                <div class="card-body">
                    <div class="emp-name" title="${{node.name}}">${{node.name}}</div>
                    <div class="emp-title">${{node.title}}</div>
                </div>
                <div class="card-footer">
                    <span>ID: ${{node.id}}</span>
                    <span>
                        ${{viewData.filter(x => x.manager === node.id).length}} Reports
                    </span>
                </div>
            `;
            
            card.onclick = (e) => {{
                e.stopPropagation();
                handleCardClick(node);
            }};
            
            li.appendChild(card);
            
            // Children
            const children = viewData.filter(c => c.manager === node.id);
            if (children.length > 0) {{
                const ul = document.createElement('ul');
                children.forEach(child => {{
                    ul.appendChild(createNode(child));
                }});
                li.appendChild(ul);
            }}
            
            return li;
        }}

        // 2. Interaction Handlers
        function handleCardClick(node) {{
            if (movingId && movingId !== node.id) {{
                // If we are in move mode, and clicked a different person -> Set them as new manager
                executeMove(movingId, node.id);
            }} else {{
                // Normal selection
                selectedId = node.id;
                movingId = null;
                searchContainer.classList.remove('active');
                openPanel(node);
                renderTree();
            }}
        }}

        function openPanel(node) {{
            document.getElementById('sel-name').textContent = node.name;
            document.getElementById('sel-title').textContent = node.title;
            document.getElementById('sel-avatar').textContent = node.name.charAt(0).toUpperCase();
            overlay.classList.add('active');
        }}

        function closePanel() {{
            overlay.classList.remove('active');
            selectedId = null;
            movingId = null;
            searchContainer.classList.remove('active');
            renderTree();
        }}

        // 3. Move Logic
        function startMoveMode() {{
            movingId = selectedId;
            searchContainer.classList.add('active');
            document.getElementById('btn-move-action').innerHTML = '👇 Click on the NEW manager in the tree or search here.';
            searchInput.focus();
        }}

        // Search Logic
        searchInput.addEventListener('input', (e) => {{
            const val = e.target.value.toLowerCase();
            resultsList.innerHTML = '';
            
            const matches = allEmps.filter(emp => 
                emp.name.toLowerCase().includes(val) && 
                emp.id !== movingId // Can't be own manager
            );
            
            matches.forEach(emp => {{
                const div = document.createElement('div');
                div.className = 'result-item';
                div.innerHTML = `<span>${{emp.name}}</span> <span style="color:#94a3b8">${{emp.title.substring(0,20)}}</span>`;
                div.onclick = () => {{
                    executeMove(movingId, emp.id);
                }};
                resultsList.appendChild(div);
            }});
        }});

        function executeMove(employeeId, newManagerId) {{
            if(employeeId === newManagerId) return; // Safety check
            
            const emp = viewData.find(x => x.id === employeeId);
            emp.manager = newManagerId;
            
            // Visual Feedback
            movingId = null;
            selectedId = null;
            closePanel();
            renderTree();
            
            // Optional: Scroll to the new position? 
            // For now just re-render.
        }}

        function deleteEmployee() {{
            if(confirm('Remove this employee from the view? (This is a draft change)')) {{
                // To delete, we set a flag or remove from viewData. 
                // We will remove from viewData and update tree.
                const idx = viewData.findIndex(x => x.id === selectedId);
                if(idx > -1) {{
                    // Reassign their children to their manager or leave as root
                    const children = viewData.filter(c => c.manager === selectedId);
                    const currentManager = viewData.find(x => x.id === selectedId).manager;
                    children.forEach(c => c.manager = currentManager || '');
                    
                    viewData.splice(idx, 1);
                }}
                closePanel();
                renderTree();
            }}
        }}

        // Download CSV
        document.getElementById('download-btn').onclick = () => {{
            const csvRows = ['Employee Code,L1 Manager Code,Employee Name,Designation,Grade,Sub Function'];
            viewData.forEach(row => {{
                csvRows.push(`"${{row.id}}","${{row.manager}}","${{row.name.replace(/"/g, '""')}}","${{row.title.replace(/"/g, '""')}}","${{row.grade}}","${{row.sub.replace(/"/g, '""')}}"`);
            }});
            const blob = new Blob([csvRows.join('\\n')], {{ type: 'text/csv' }});
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'org_design_draft.csv';
            a.click();
        }};

        // Initial Render
        renderTree();
        
    </script>
    </body>
    </html>
    """
    
    components.html(html_content, height=800)
