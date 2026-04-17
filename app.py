import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import json
import io

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

        /* Base Reset & Typography */
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important; }
        
        /* Hide Streamlit Branding */
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header { visibility: hidden; }
        .stDeployButton { display: none !important; }

        /* App Background - Premium Light */
        .stApp {
            background-color: #fafafa !important;
            background-image: 
                radial-gradient(circle at 0% 0%, rgba(99, 102, 241, 0.05) 0%, transparent 40%),
                radial-gradient(circle at 100% 100%, rgba(16, 185, 129, 0.05) 0%, transparent 40%);
            min-height: 100vh;
        }
        
        /* Main Container */
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
            max-width: 99% !important;
            position: relative;
            z-index: 1;
        }

        /* Header Styling */
        h1 {
            font-weight: 800 !important;
            color: #0f172a !important;
            letter-spacing: -0.03em !important;
            font-size: 2.6rem !important;
            margin-bottom: 0.5rem !important;
        }
        
        h2, h3 {
            color: #1e293b !important;
            font-weight: 700 !important;
        }

        /* Premium Light Sidebar */
        [data-testid="stSidebar"] {
            background: #ffffff !important;
            border-right: 1px solid #e5e7eb !important;
            box-shadow: 4px 0 24px rgba(0, 0, 0, 0.02) !important;
        }
        [data-testid="stSidebar"] * { color: #374151 !important; }
        [data-testid="stSidebar"] h3 {
            color: #111827 !important;
            font-size: 1.05rem !important;
            margin-bottom: 0.75rem !important;
            padding-bottom: 0.5rem !important;
            border-bottom: 2px solid #f3f4f6 !important;
        }

        /* Form Elements */
        [data-testid="stSidebar"] div[data-baseweb="select"] > div,
        .stSelectbox > div > div {
            background-color: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 10px !important;
            color: #374151 !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important;
        }
        [data-testid="stSidebar"] div[data-baseweb="select"]:hover > div,
        .stSelectbox > div > div:hover {
            border-color: #8b5cf6 !important;
        }
        
        /* File Uploader */
        [data-testid="stFileUploadDropzone"] {
            background: #ffffff !important;
            border: 2px dashed #d1d5db !important;
            border-radius: 14px !important;
            transition: all 0.2s;
            min-height: 120px !important;
        }
        [data-testid="stFileUploadDropzone"]:hover {
            border-color: #8b5cf6 !important;
            background: #fdfbff !important;
        }

        /* Buttons */
        .stButton > button {
            border-radius: 10px !important;
            font-weight: 600 !important;
            font-size: 0.9rem !important;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
            width: 100% !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important;
        }
        
        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%) !important;
            border: none !important;
            color: #ffffff !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.2) !important;
        }
        .stButton > button[kind="primary"]:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 6px 16px rgba(99, 102, 241, 0.3) !important;
        }
        
        .stButton > button[kind="secondary"] {
            background: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            color: #111827 !important;
        }
        .stButton > button[kind="secondary"]:hover {
            background: #f9fafb !important;
            border-color: #d1d5db !important;
        }

        /* Toggles */
        [data-testid="stToggle"] > label > div > div {
            background-color: #f3f4f6 !important;
            border-color: #d1d5db !important;
        }
        [data-testid="stToggle"] > label > div > div[aria-checked="true"] {
            background-color: #8b5cf6 !important;
            border-color: #8b5cf6 !important;
        }

        /* Alert */
        .stAlert {
            background: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 12px !important;
            color: #374151 !important;
        }
        .stAlert > div { background-color: transparent !important; }
    </style>
""", unsafe_allow_html=True)

# 3. Initialize Session State
if 'draft_moves' not in st.session_state:
    st.session_state.draft_moves = {}

if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False

# 4. Sidebar Controls
with st.sidebar:
    st.markdown("### 🏢 OrgDesign Pro")
    st.markdown("<p style='color: #6b7280; font-size: 0.85rem; margin-top: -0.5rem; margin-bottom: 1.5rem;'>Interactive Org Architecture</p>", unsafe_allow_html=True)
    
    st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
    st.markdown("##### 📊 Upload HR Data")
    uploaded_file = st.file_uploader("", type=['csv', 'xlsx'], label_visibility="collapsed")
    
    if uploaded_file is not None:
        st.session_state.file_uploaded = True
        if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
        else: df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()
        
        if 'Employee Code' in df.columns:
            df['Clean_Emp_Code'] = df['Employee Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
            if 'L1 Manager Code' in df.columns:
                df['L1 Manager Code'] = df['L1 Manager Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
            name_to_code = dict(zip(df['Employee Name'], df['Clean_Emp_Code']))
    else:
        st.session_state.file_uploaded = False
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("### ⚙️ Controls")
    selected_sub = "All"
    include_retainers = True
    include_inactive = False
    
    if st.session_state.file_uploaded:
        st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
        st.markdown("##### 🔍 Filters")
        
        if 'Sub Function' in df.columns:
            sub_functions = sorted(df['Sub Function'].dropna().unique())
            selected_sub = st.selectbox("Sub Function View", ["All"] + sub_functions, index=0)
        
        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div style='background: linear-gradient(135deg, #f5f3ff 0%, #ffffff 100%); border: 1px solid #ddd6fe; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
        st.markdown("##### 🧪 Draft Mode")
        enable_draft = st.toggle("Enable Interactive Draft", value=False, help="Allows drag & drop re-parenting and inline editing.")
        
        if enable_draft:
            st.markdown("<div style='background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 10px; padding: 0.75rem; margin-top: 0.5rem;'>", unsafe_allow_html=True)
            st.markdown("""
                <p style='color: #059669; font-size: 0.85rem; margin: 0; font-weight: 600; text-align: center;'>✨ Interactive Workspace</p>
                <ul style='color: #059669; font-size: 0.8rem; margin: 0.5rem 0 0 1rem; padding-left: 0.5rem;'>
                    <li>Drag cards onto managers to move them.</li>
                    <li>Double-click a card to edit details.</li>
                    <li>Click "Download Drafted CSV" to save changes.</li>
                </ul>
            """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)

# 5. Main Dashboard Area
if not st.session_state.file_uploaded:
    st.markdown("<div style='text-align: center; padding: 4rem 2rem; max-width: 900px; margin: 0 auto;'>", unsafe_allow_html=True)
    st.title("🏢 OrgDesign Pro")
    st.markdown("<p style='color: #4b5563; font-size: 1.2rem; margin-top: 0.5rem; margin-bottom: 3rem;'>Transform HR data into interactive organizational charts.</p>", unsafe_allow_html=True)
    st.markdown("<div style='background: #ffffff; border: 1px dashed #d1d5db; border-radius: 16px; padding: 2rem;'><p style='color: #9ca3af;'>Upload a file to begin</p></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.title(f"Org Architecture: {selected_sub if selected_sub != 'All' else 'Full Organization'}")
    
    base_filtered_df = df.copy()
    
    # Filters
    if not include_retainers and 'Employment Type' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[~base_filtered_df['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
    if not include_inactive and 'Status' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[~base_filtered_df['Status'].astype(str).str.contains('Inactive', case=False, na=False)]

    if selected_sub != "All" and 'Sub Function' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[base_filtered_df['Sub Function'] == selected_sub]
        
    valid_ids = [str(vid).strip() for vid in base_filtered_df['Employee Code'].tolist()]
    
    # Prepare JSON Data for Frontend
    org_data = []
    for index, row in base_filtered_df.iterrows():
        emp_id = str(row.get('Employee Code', '')).strip()
        mgr_id = str(row.get('L1 Manager Code', '')).strip()
        if mgr_id not in valid_ids: mgr_id = ""
        
        org_data.append({
            "id": emp_id,
            "manager": mgr_id,
            "name": str(row.get('Employee Name', '')).strip(),
            "title": str(row.get('Designation', '')).strip(),
            "grade": str(row.get('Grade', '')).strip(),
            "sub": str(row.get('Sub Function', '')).strip()
        })
    
    data_json = json.dumps(org_data)

    # 6. HTML Templates
    if enable_draft:
        # --- INTERACTIVE DRAFT MODE ---
        html_template = f"""
        <html>
          <head>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
            <style>
              :root {{
                --bg: #fafafa;
                --card-bg: #ffffff;
                --border: #e5e7eb;
                --accent: #8b5cf6;
                --text-main: #111827;
                --text-sub: #4b5563;
                --shadow: 0 2px 8px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02);
                --shadow-hover: 0 12px 24px rgba(0,0,0,0.08), 0 4px 8px rgba(139, 92, 246, 0.12);
              }}
              
              body {{ background: var(--bg); font-family: 'Inter', sans-serif; color: var(--text-main); margin: 0; padding: 20px; }}
              
              .toolbar {{ display: flex; gap: 12px; margin-bottom: 24px; flex-wrap: wrap; align-items: center; }}
              .btn {{ background: #ffffff; border: 1px solid #d1d5db; padding: 10px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px; transition: all 0.2s; color: #374151; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }}
              .btn:hover {{ background: #f9fafb; transform: translateY(-1px); box-shadow: 0 4px 6px rgba(0,0,0,0.06); }}
              .btn-primary {{ background: linear-gradient(135deg, #8b5cf6, #6366f1); color: white; border: none; box-shadow: 0 4px 12px rgba(99, 102, 241, 0.2); }}
              .btn-primary:hover {{ box-shadow: 0 6px 16px rgba(99, 102, 241, 0.3); }}
              
              #chart-container {{ overflow: auto; padding: 40px; background: #ffffff; border: 1px solid #e5e7eb; border-radius: 16px; box-shadow: inset 0 1px 3px rgba(0,0,0,0.02); min-height: 500px; position: relative; }}
              
              .node {{ display: inline-block; vertical-align: top; text-align: center; position: relative; margin: 10px; }}
              .children {{ display: flex; flex-direction: row; justify-content: center; position: relative; margin-top: 30px; }}
              .children::before {{ content: ''; position: absolute; top: -30px; left: 50%; width: 0; height: 30px; border-left: 2px solid #d1d5db; }}
              .child-wrapper {{ position: relative; margin: 0 15px; display: inline-block; vertical-align: top; }}
              .child-wrapper::before {{ content: ''; position: absolute; top: -30px; left: 0; right: 0; height: 0; border-top: 2px solid #d1d5db; }}
              .child-wrapper:first-child::before {{ left: 50%; }}
              .child-wrapper:last-child::before {{ right: 50%; }}
              .child-wrapper:only-child::before {{ display: none; }}
              
              .card {{
                background: var(--card-bg);
                border: 1px solid var(--border);
                border-radius: 12px;
                width: 240px;
                padding: 0;
                box-shadow: var(--shadow);
                transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
                cursor: grab;
                position: relative;
                overflow: visible;
                border-top: 3px solid #10b981;
              }}
              .card:hover {{ box-shadow: var(--shadow-hover); transform: translateY(-3px); z-index: 10; }}
              .card.dragging {{ opacity: 0.4; transform: scale(0.95); }}
              .card.drop-target {{ border-color: var(--accent); box-shadow: 0 0 0 4px rgba(139, 92, 246, 0.15); transform: scale(1.05); }}
              .card.is-root {{ border-top-color: #8b5cf6; }}
              .card.is-moved {{ border-top-color: #3b82f6; }}
              .card.is-moved::after {{ content: '✎'; position: absolute; top: -8px; right: -8px; background: #3b82f6; color: white; border-radius: 50%; width: 20px; height: 20px; font-size: 12px; display: flex; align-items: center; justify-content: center; border: 2px solid white; }}
              
              .card-header {{ padding: 12px 14px; background: #f8f9fb; border-bottom: 1px solid #f3f4f6; display: flex; justify-content: space-between; align-items: center; }}
              .badge {{ background: #f3f4f6; color: #4b5563; font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 4px; text-transform: uppercase; }}
              .grade {{ color: #9ca3af; font-size: 11px; font-weight: 600; }}
              
              .card-body {{ padding: 14px; text-align: left; }}
              .name-input {{ width: 100%; border: 1px solid transparent; background: transparent; font-weight: 700; font-size: 15px; color: #111827; margin-bottom: 4px; font-family: inherit; }}
              .name-input:focus, .title-input:focus {{ border-color: #d1d5db; background: #ffffff; outline: none; padding: 2px 4px; border-radius: 4px; }}
              .title-input {{ width: 100%; border: 1px solid transparent; background: transparent; font-size: 12px; color: #4b5563; font-family: inherit; }}
              
              .card-footer {{ background: #f9fafb; padding: 8px 14px; border-top: 1px solid #f3f4f6; display: flex; justify-content: space-between; font-size: 11px; color: #6b7280; }}
              
              .toast {{ position: fixed; bottom: 20px; right: 20px; background: #1f2937; color: white; padding: 12px 20px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.2); transform: translateY(100px); opacity: 0; transition: all 0.3s; z-index: 100; }}
              .toast.show {{ transform: translateY(0); opacity: 1; }}
            </style>
          </head>
          <body>
            <div class="toolbar">
              <button class="btn btn-primary" id="download-draft-btn">💾 Download Drafted CSV</button>
              <button class="btn" id="reset-btn">↺ Reset Changes</button>
              <div style="flex-grow: 1; text-align: right; font-size: 0.9rem; color: #6b7280;">
                🧪 Drag & Drop to move • Double-click to edit
              </div>
            </div>
            <div id="chart-container"></div>
            <div id="toast" class="toast">✅ Changes saved!</div>

            <script>
              let orgData = {data_json};
              let movedIds = new Set();
              
              const container = document.getElementById('chart-container');
              const toast = document.getElementById('toast');

              function render() {{
                container.innerHTML = '';
                const roots = orgData.filter(n => !n.manager || !orgData.find(m => m.id === n.manager));
                roots.forEach(root => {{
                  container.appendChild(createNode(root));
                }});
              }}

              function createNode(node) {{
                const wrapper = document.createElement('div');
                wrapper.className = 'node';
                
                const card = document.createElement('div');
                card.className = 'card';
                if (!node.manager) card.classList.add('is-root');
                if (movedIds.has(node.id)) card.classList.add('is-moved');
                
                card.setAttribute('draggable', 'true');
                card.dataset.id = node.id;

                // Drag Events
                card.addEventListener('dragstart', (e) => {{
                  e.dataTransfer.setData('text/plain', node.id);
                  setTimeout(() => card.classList.add('dragging'), 0);
                }});
                card.addEventListener('dragend', () => card.classList.remove('dragging'));

                // Drop Events
                card.addEventListener('dragover', (e) => {{
                  e.preventDefault();
                  if (e.currentTarget.dataset.id !== node.id) {{
                    card.classList.add('drop-target');
                  }}
                }});
                card.addEventListener('dragleave', () => card.classList.remove('drop-target'));
                card.addEventListener('drop', (e) => {{
                  e.preventDefault();
                  card.classList.remove('drop-target');
                  const draggedId = e.dataTransfer.getData('text/plain');
                  if (draggedId && draggedId !== node.id) {{
                    handleDrop(draggedId, node.id);
                  }}
                }});

                // Double Click Edit
                card.addEventListener('dblclick', (e) => {{
                  if (e.target.tagName === 'INPUT') return;
                  enableEditing(card, node);
                }});

                // Inner HTML
                card.innerHTML = `
                  <div class="card-header">
                    <span class="badge">${{node.sub.substring(0, 12)}}</span>
                    <span class="grade">GR: ${{node.grade}}</span>
                  </div>
                  <div class="card-body">
                    <div class="name-input" readonly>${{node.name}}</div>
                    <div class="title-input" readonly>${{node.title}}</div>
                  </div>
                  <div class="card-footer">
                    <span>Directs: ${{getDirectReports(node.id).length}}</span>
                    <span>ID: ${{node.id.substring(0, 6)}}</span>
                  </div>
                `;
                wrapper.appendChild(card);

                const children = orgData.filter(n => n.manager === node.id);
                if (children.length > 0) {{
                  const childrenWrapper = document.createElement('div');
                  childrenWrapper.className = 'children';
                  children.forEach(child => {{
                    childrenWrapper.appendChild(createNode(child));
                  }});
                  wrapper.appendChild(childrenWrapper);
                }}

                return wrapper;
              }}

              function enableEditing(card, node) {{
                const nameInput = card.querySelector('.name-input');
                const titleInput = card.querySelector('.title-input');
                
                nameInput.removeAttribute('readonly');
                nameInput.style.border = '1px solid #d1d5db';
                nameInput.style.background = '#fff';
                nameInput.focus();

                const finishEdit = () => {{
                  nameInput.setAttribute('readonly', true);
                  titleInput.setAttribute('readonly', true);
                  nameInput.style.border = 'none';
                  titleInput.style.border = 'none';
                  titleInput.style.background = 'transparent';
                  nameInput.style.background = 'transparent';
                  node.name = nameInput.value;
                  node.title = titleInput.value;
                  render(); // Re-render to update stats etc
                }};

                nameInput.addEventListener('blur', finishEdit, {{ once: true }});
                nameInput.addEventListener('keydown', (e) => {{
                  if (e.key === 'Enter') nameInput.blur();
                }});
              }}

              function handleDrop(draggedId, targetId) {{
                // Cycle detection: target cannot be a descendant of dragged
                if (isDescendant(draggedId, targetId)) {{
                  showToast('⚠️ Invalid move: Cannot move manager to their own reportee.');
                  return;
                }}
                
                const draggedNode = orgData.find(n => n.id === draggedId);
                draggedNode.manager = targetId;
                movedIds.add(draggedId);
                render();
                showToast('🔄 Employee moved successfully');
              }}

              function isDescendant(ancestorId, targetId) {{
                const children = orgData.filter(n => n.manager === ancestorId);
                for (let child of children) {{
                  if (child.id === targetId) return true;
                  if (isDescendant(child.id, targetId)) return true;
                }}
                return false;
              }}

              function getDirectReports(managerId) {{
                return orgData.filter(n => n.manager === managerId);
              }}

              function showToast(msg) {{
                toast.textContent = msg;
                toast.classList.add('show');
                setTimeout(() => toast.classList.remove('show'), 3000);
              }}

              // Button Actions
              document.getElementById('reset-btn').addEventListener('click', () => {{
                movedIds.clear();
                orgData = {data_json}; // Reset to initial
                render();
                showToast('↺ Draft reset');
              }});

              document.getElementById('download-draft-btn').addEventListener('click', () => {{
                const csvContent = convertToCSV(orgData);
                const blob = new Blob([csvContent], {{ type: 'text/csv;charset=utf-8;' }});
                saveAs(blob, 'org_chart_draft.csv');
              }});

              function convertToCSV(data) {{
                const headers = ['Employee Code', 'L1 Manager Code', 'Employee Name', 'Designation', 'Grade', 'Sub Function'];
                const rows = data.map(n => [n.id, n.manager, n.name, n.title, n.grade, n.sub]);
                return [headers.join(','), ...rows.map(r => r.map(val => `"$ {{String(val).replace(/"/g, '""') }}"`).join(','))].join('\\n');
              }}

              render();
            </script>
          </body>
        </html>
        """
    else:
        # --- STANDARD VIEW (Google Charts) ---
        # Build standard data string for Google Charts
        js_rows = []
        for row in org_data:
            # Google Charts format: [{{v: id, f: html}}, manager_id, '']
            box_html = f"<div class='beautiful-card'><div class='card-header'><span class='badge'>{row['sub'][:12]}</span><span class='grade'>GR: {row['grade']}</span></div><div class='card-body'><div class='card-name'>{row['name']}</div><div class='card-title'>{row['title']}</div></div><div class='card-footer'><div class='stat'><span>ID</span><b style='font-size:11px'>{row['id']}</b></div></div></div>"
            box_html_clean = box_html.replace('\n', '').replace('\r', '').replace("'", "\\'")
            mgr_str = f"'{row['manager']}'" if row['manager'] else "''"
            js_rows.append(f"[{{'v': '{row['id']}', 'f': \"{box_html_clean}\"}}, {mgr_str}, '']")
        
        rows_str = ",\n".join(js_rows)

        html_template = f"""
        <html>
          <head>
            <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
            <style>
              @keyframes slideUp {{ from {{ opacity: 0; transform: translateY(15px); }} to {{ opacity: 1; transform: translateY(0); }} }}
              body {{ margin: 0; background: #fafafa; font-family: 'Inter', sans-serif; }}
              
              .btn-container {{ display: flex; gap: 12px; margin-bottom: 24px; flex-wrap: wrap; padding: 0 4px; }}
              .btn {{ background: #ffffff; color: #374151; border: 1px solid #d1d5db; padding: 12px 24px; border-radius: 10px; cursor: pointer; font-weight: 600; font-size: 14px; transition: all 0.2s; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }}
              .btn:hover {{ background: #f9fafb; transform: translateY(-1px); box-shadow: 0 4px 6px rgba(0,0,0,0.06); }}
              
              #scroll_wrapper {{ overflow: auto; background: #ffffff; border: 1px solid #e5e7eb; border-radius: 16px; padding: 30px; box-shadow: inset 0 1px 3px rgba(0,0,0,0.02); }}
              #chart_div {{ min-width: 100%; display: flex; justify-content: center; }}

              .google-visualization-orgchart-lineleft, .google-visualization-orgchart-lineright, .google-visualization-orgchart-linebottom {{ border-color: #d1d5db !important; border-width: 2px !important; }}
              
              .myNode {{ border: none !important; background: none !important; padding: 0 !important; box-shadow: none !important; margin: 16px; cursor: pointer; }}
              .selectedNode .beautiful-card {{ border-color: #8b5cf6 !important; box-shadow: 0 0 0 3px rgba(139, 92, 246, 0.2) !important; }}
              
              .beautiful-card {{ border-radius: 14px; background: #ffffff; border: 1px solid #e5e7eb; width: 240px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.04); animation: slideUp 0.4s ease-out; transition: all 0.25s; border-top: 4px solid #10b981; }}
              .beautiful-card:hover {{ transform: translateY(-4px); box-shadow: 0 12px 24px rgba(0,0,0,0.08); border-color: #c7d2fe; z-index: 10; }}
              
              .card-header {{ padding: 14px 16px; background: #f8f9fb; border-bottom: 1px solid #f3f4f6; display: flex; justify-content: space-between; }}
              .badge {{ background: #f3f4f6; color: #4b5563; padding: 4px 8px; border-radius: 4px; font-size: 10px; font-weight: 700; text-transform: uppercase; }}
              .grade {{ color: #9ca3af; font-size: 11px; font-weight: 600; }}
              .card-body {{ padding: 16px; text-align: left; }}
              .card-name {{ font-size: 15px; font-weight: 700; color: #111827; margin-bottom: 4px; }}
              .card-title {{ font-size: 12px; color: #4b5563; line-height: 1.4; }}
              .card-footer {{ background: #f9fafb; padding: 8px 16px; border-top: 1px solid #f3f4f6; text-align: left; font-size: 11px; color: #6b7280; }}
              
              /* Scrollbar */
              ::-webkit-scrollbar {{ width: 10px; height: 10px; }}
              ::-webkit-scrollbar-track {{ background: #f3f4f6; border-radius: 10px; }}
              ::-webkit-scrollbar-thumb {{ background: #d1d5db; border-radius: 10px; border: 2px solid #f3f4f6; }}
            </style>
          </head>
          <body>
            <div class="btn-container">
              <button class="btn" id="download-btn">⬇️ Download View</button>
              <button class="btn" id="zip-btn">📦 Export All (Not avail)</button>
            </div>
            <div id="scroll_wrapper">
              <div id="chart_div"></div>
            </div>
            <script>
              google.charts.load('current', {{packages:["orgchart"]}});
              google.charts.setOnLoadCallback(drawChart);
              
              function drawChart() {{
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Name');
                data.addColumn('string', 'Manager');
                data.addColumn('string', 'ToolTip');
                data.addRows([{rows_str}]);
                
                var chart = new google.visualization.OrgChart(document.getElementById('chart_div'));
                chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'large', nodeClass:'myNode', selectedNodeClass:'selectedNode'}});
              }}
              
              document.getElementById('download-btn').addEventListener('click', () => {{
                html2canvas(document.getElementById('chart_div'), {{ backgroundColor: "#ffffff", scale: 2 }}).then(canvas => {{
                  let link = document.createElement('a');
                  link.download = 'org_chart_view.png';
                  link.href = canvas.toDataURL();
                  link.click();
                }});
              }});
            </script>
          </body>
        </html>
        """
    
    components.html(html_template, height=1000, scrolling=True)
