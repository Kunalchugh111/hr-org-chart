import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. Page Configuration
st.set_page_config(
    page_title="OrgDesign Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Premium SaaS Styling
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

        /* App Background - Premium Dark Gradient */
        .stApp {
            background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%) !important;
            min-height: 100vh;
        }
        
        /* Ambient Glow Effects */
        .stApp::before {
            content: '';
            position: fixed;
            top: -20%;
            right: -10%;
            width: 600px;
            height: 600px;
            background: radial-gradient(circle, rgba(139, 92, 246, 0.15) 0%, transparent 70%);
            border-radius: 50%;
            pointer-events: none;
            z-index: 0;
        }
        .stApp::after {
            content: '';
            position: fixed;
            bottom: -10%;
            left: -10%;
            width: 500px;
            height: 500px;
            background: radial-gradient(circle, rgba(59, 130, 246, 0.12) 0%, transparent 70%);
            border-radius: 50%;
            pointer-events: none;
            z-index: 0;
        }

        /* Main Container */
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
            max-width: 98% !important;
            position: relative;
            z-index: 1;
        }

        /* Header Styling */
        h1 {
            font-weight: 800 !important;
            color: #ffffff !important;
            letter-spacing: -0.04em !important;
            font-size: 3rem !important;
            margin-bottom: 0.5rem !important;
            background: linear-gradient(135deg, #ffffff 0%, #c7d2fe 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        h2, h3 {
            color: #e2e8f0 !important;
            font-weight: 700 !important;
            letter-spacing: -0.02em !important;
        }

        /* Premium Sidebar */
        [data-testid="stSidebar"] {
            background: rgba(15, 23, 42, 0.8) !important;
            backdrop-filter: blur(20px) !important;
            border-right: 1px solid rgba(139, 92, 246, 0.2) !important;
        }
        [data-testid="stSidebar"]::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            bottom: 0;
            width: 1px;
            background: linear-gradient(180deg, rgba(139, 92, 246, 0.4) 0%, transparent 50%, rgba(59, 130, 246, 0.4) 100%);
        }
        [data-testid="stSidebar"] * { color: #cbd5e1 !important; }
        [data-testid="stSidebar"] h3 {
            color: #a5b4fc !important;
            font-size: 1.1rem !important;
            margin-bottom: 1rem !important;
            padding-bottom: 0.5rem !important;
            border-bottom: 1px solid rgba(139, 92, 246, 0.2) !important;
        }

        /* Form Elements in Sidebar */
        [data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background-color: rgba(30, 41, 59, 0.8) !important;
            border: 1px solid rgba(139, 92, 246, 0.3) !important;
            border-radius: 8px !important;
            color: #e2e8f0 !important;
        }
        [data-testid="stSidebar"] div[data-baseweb="select"]:hover > div {
            border-color: rgba(139, 92, 246, 0.6) !important;
        }
        
        /* File Uploader */
        [data-testid="stFileUploadDropzone"] {
            background: rgba(30, 41, 59, 0.4) !important;
            backdrop-filter: blur(10px) !important;
            border: 2px dashed rgba(139, 92, 246, 0.4) !important;
            border-radius: 12px !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            min-height: 120px !important;
        }
        [data-testid="stFileUploadDropzone"]:hover {
            border-color: #8b5cf6 !important;
            background: rgba(139, 92, 246, 0.1) !important;
            transform: translateY(-2px);
        }
        [data-testid="stFileUploadDropzone"] p { color: #cbd5e1 !important; }

        /* Toggles */
        [data-testid="stToggle"] > label > div > div {
            background-color: rgba(139, 92, 246, 0.2) !important;
            border-color: #8b5cf6 !important;
        }
        [data-testid="stToggle"] > label > div > div[aria-checked="true"] {
            background-color: #8b5cf6 !important;
            border-color: #8b5cf6 !important;
        }

        /* Buttons */
        .stButton > button {
            background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
            padding: 0.75rem 1.5rem !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
            box-shadow: 0 4px 15px rgba(139, 92, 246, 0.3) !important;
            width: 100% !important;
        }
        .stButton > button:hover {
            transform: translateY(-2px) scale(1.02) !important;
            box-shadow: 0 8px 25px rgba(139, 92, 246, 0.4) !important;
        }
        .stButton > button:active {
            transform: translateY(0) scale(0.98) !important;
        }

        /* Select Box in Main Area */
        .stSelectbox > div > div {
            background-color: rgba(30, 41, 59, 0.8) !important;
            border: 1px solid rgba(139, 92, 246, 0.3) !important;
            border-radius: 10px !important;
            color: #e2e8f0 !important;
        }

        /* Info/Success/Error Boxes */
        .stAlert {
            background: rgba(30, 41, 59, 0.6) !important;
            border: 1px solid rgba(139, 92, 246, 0.3) !important;
            border-radius: 12px !important;
            backdrop-filter: blur(10px) !important;
            color: #e2e8f0 !important;
        }
        .stAlert div { color: #e2e8f0 !important; }
        .stAlert > div { background-color: transparent !important; }

        /* Metrics/Stats Cards */
        [data-testid="stMetric"] {
            background: rgba(30, 41, 59, 0.6) !important;
            border: 1px solid rgba(139, 92, 246, 0.2) !important;
            border-radius: 12px !important;
            padding: 1rem !important;
            backdrop-filter: blur(10px) !important;
        }

        /* Divider */
        .stDivider {
            border-color: rgba(139, 92, 246, 0.2) !important;
        }

        /* Tooltip */
        [data-testid="stTooltipHoverContent"] {
            background-color: #1e293b !important;
            border: 1px solid rgba(139, 92, 246, 0.3) !important;
        }
    </style>
""", unsafe_allow_html=True)

# 3. Initialize Session State
if 'draft_moves' not in st.session_state:
    st.session_state.draft_moves = {}

if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False

# 4. Sidebar Controls
with st.sidebar:
    st.markdown("### ✨ OrgDesign Pro")
    st.markdown("<p style='color: #94a3b8; font-size: 0.85rem; margin-top: -0.5rem; margin-bottom: 1.5rem;'>Premium Organizational Architecture Platform</p>", unsafe_allow_html=True)
    
    st.markdown("<div style='background: rgba(30, 41, 59, 0.6); border: 1px solid rgba(139, 92, 246, 0.2); border-radius: 12px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
    st.markdown("##### 📊 Upload HR Data")
    st.markdown("<p style='color: #64748b; font-size: 0.8rem; margin-top: -0.5rem;'>Drag & drop or browse your roster file</p>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("", type=['csv', 'xlsx'], label_visibility="collapsed")
    
    if uploaded_file is not None:
        st.session_state.file_uploaded = True
        if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
        else: df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()
        
        if 'Employee Code' in df.columns:
            df['Clean_Emp_Code'] = df['Employee Code'].astype(str).str.replace('.0', '', regex=False).str.strip()
            name_to_code = dict(zip(df['Employee Name'], df['Clean_Emp_Code']))
    else:
        st.session_state.file_uploaded = False
        
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("### ⚙️ Configuration")
    
    selected_sub = "All"
    include_retainers = True
    include_inactive = False
    
    if st.session_state.file_uploaded:
        st.markdown("<div style='background: rgba(30, 41, 59, 0.6); border: 1px solid rgba(139, 92, 246, 0.2); border-radius: 12px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
        st.markdown("##### 🔍 Data Filters")
        
        if 'Sub Function' in df.columns:
            sub_functions = sorted(df['Sub Function'].dropna().unique())
            selected_sub = st.selectbox("Sub Function View", ["All"] + sub_functions, index=0)
        
        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div style='background: rgba(30, 41, 59, 0.6); border: 1px solid rgba(139, 92, 246, 0.2); border-radius: 12px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
        st.markdown("##### 🧪 Scenario Modeling")
        enable_draft = st.toggle("Enable Draft Mode")
        
        if enable_draft:
            emp_names = sorted(df['Employee Name'].dropna().astype(str).tolist())
            move_emp = st.selectbox("Employee to Move", [""] + emp_names)
            new_mgr = st.selectbox("New Manager", [""] + emp_names)
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Apply", use_container_width=True):
                    if move_emp and new_mgr:
                        e_code = name_to_code.get(move_emp)
                        m_code = name_to_code.get(new_mgr)
                        if e_code and m_code:
                            st.session_state.draft_moves[e_code] = m_code
                            st.rerun()
            with col2:
                if st.button("↺ Reset", use_container_width=True, type="secondary"):
                    st.session_state.draft_moves = {}
                    st.rerun()
                        
            if st.session_state.draft_moves:
                st.markdown(f"<div style='background: rgba(34, 197, 94, 0.1); border: 1px solid rgba(34, 197, 94, 0.3); border-radius: 8px; padding: 0.5rem; text-align: center; margin-top: 0.5rem;'><p style='color: #4ade80; font-size: 0.8rem; margin: 0;'>✨ {len(st.session_state.draft_moves)} moves active</p></div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)

# 5. Main Dashboard Area
if not st.session_state.file_uploaded:
    st.markdown("<div style='text-align: center; padding: 4rem 2rem;'>", unsafe_allow_html=True)
    st.title("🏢 OrgDesign Pro")
    st.markdown("<p style='color: #94a3b8; font-size: 1.25rem; margin-top: 0.5rem; margin-bottom: 2rem;'>Transform your HR data into stunning organizational charts</p>", unsafe_allow_html=True)
    
    st.markdown("""
        <div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem; max-width: 800px; margin: 0 auto; text-align: left;'>
            <div style='background: rgba(30, 41, 59, 0.6); border: 1px solid rgba(139, 92, 246, 0.2); border-radius: 12px; padding: 1.5rem; backdrop-filter: blur(10px);'>
                <div style='font-size: 2rem; margin-bottom: 0.5rem;'>📊</div>
                <h3 style='color: #e2e8f0; margin: 0 0 0.5rem 0; font-size: 1.1rem;'>Smart Parsing</h3>
                <p style='color: #64748b; font-size: 0.85rem; margin: 0;'>Automatically detects and maps reporting hierarchies</p>
            </div>
            <div style='background: rgba(30, 41, 59, 0.6); border: 1px solid rgba(139, 92, 246, 0.2); border-radius: 12px; padding: 1.5rem; backdrop-filter: blur(10px);'>
                <div style='font-size: 2rem; margin-bottom: 0.5rem;'>🎨</div>
                <h3 style='color: #e2e8f0; margin: 0 0 0.5rem 0; font-size: 1.1rem;'>Premium Design</h3>
                <p style='color: #64748b; font-size: 0.85rem; margin: 0;'>Beautiful, modern cards with rich details</p>
            </div>
            <div style='background: rgba(30, 41, 59, 0.6); border: 1px solid rgba(139, 92, 246, 0.2); border-radius: 12px; padding: 1.5rem; backdrop-filter: blur(10px);'>
                <div style='font-size: 2rem; margin-bottom: 0.5rem;'>🧪</div>
                <h3 style='color: #e2e8f0; margin: 0 0 0.5rem 0; font-size: 1.1rem;'>Scenario Mode</h3>
                <p style='color: #64748b; font-size: 0.85rem; margin: 0;'>Model org changes before implementing them</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div style='margin-top: 3rem; padding: 1.5rem; background: rgba(30, 41, 59, 0.4); border: 1px dashed rgba(139, 92, 246, 0.3); border-radius: 12px; max-width: 600px; margin-left: auto; margin-right: auto;'>
            <p style='color: #94a3b8; margin: 0;'>
                <strong style='color: #a5b4fc;'>👈 Get Started:</strong> Upload your HR roster (CSV/XLSX) from the sidebar to begin mapping your organizational structure.
            </p>
        </div>
    """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.title(f"Org Architecture: {selected_sub if selected_sub != 'All' else 'Full Organization'}")
    
    st.markdown("""
        <div style='display: flex; align-items: center; gap: 0.5rem; margin-top: -10px; margin-bottom: 1.5rem;'>
            <span style='color: #64748b; font-size: 0.9rem;'>💡 Tip: Double-click any manager's card to collapse or expand their reporting lines.</span>
        </div>
    """, unsafe_allow_html=True)
    
    base_filtered_df = df.copy()
    
    # APPLY DRAFT MOVES TO THE DATAFRAME FIRST
    if st.session_state.draft_moves:
        for e_code, m_code in st.session_state.draft_moves.items():
            base_filtered_df.loc[base_filtered_df['Clean_Emp_Code'] == e_code, 'L1 Manager Code'] = m_code

    if not include_retainers and 'Employment Type' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[~base_filtered_df['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
    if not include_inactive and 'Status' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[~base_filtered_df['Status'].astype(str).str.contains('Inactive', case=False, na=False)]

    # Build data for all sub-functions (for ZIP download)
    all_subs_js_dict_items = []
    
    if 'Sub Function' in base_filtered_df.columns:
        unique_subs = base_filtered_df['Sub Function'].dropna().unique()
        for sub in unique_subs:
            sub_df = base_filtered_df[base_filtered_df['Sub Function'] == sub]
            valid_ids_sub = [str(vid).replace('.0', '').strip() for vid in sub_df['Employee Code'].tolist()]
            js_rows_sub = []
            
            for index, row in sub_df.iterrows():
                emp_id = str(row.get('Employee Code', '')).replace('.0', '').strip()
                manager_id = str(row.get('L1 Manager Code', '')).replace('.0', '').strip()
                name = str(row.get('Employee Name', '')).strip()
                grade = str(row.get('Grade', '')).strip()
                designation = str(row.get('Designation', '')).strip()
                onroll = str(row.get('Onroll Reportees', '0')).strip()
                sub_func = str(row.get('Sub Function', '')).strip()
                
                if manager_id not in valid_ids_sub: manager_id = ""
                    
                box_html = f"<div class='beautiful-card'><div class='card-header'><span class='badge'>{sub_func[:15]}</span><span class='grade'>GR: {grade}</span></div><div class='card-body'><div class='card-name'>{name}</div><div class='card-title'>{designation}</div></div><div class='card-footer'><div class='stat'><span>On-Roll</span><b>{onroll}</b></div><div class='stat'><span>Appr HC</span><b style='color:#ea580c;'>-</b></div><div class='stat'><span>Off-Roll</span><b style='color:#65a30d;'>-</b></div></div></div>"
                box_html_clean = box_html.replace('\n', '').replace('\r', '').replace("'", "\\'")
                mgr_str = f"'{manager_id}'" if manager_id else "''"
                js_rows_sub.append(f"[{{'v': '{emp_id}', 'f': \"{box_html_clean}\"}}, {mgr_str}, '']")
                
            sub_rows_str = "[\n" + ",\n".join(js_rows_sub) + "\n]"
            all_subs_js_dict_items.append(f'"{sub}": {sub_rows_str}')

    all_data_dict_str = "{\n" + ",\n".join(all_subs_js_dict_items) + "\n}"

    current_view_df = base_filtered_df.copy()
    if selected_sub != "All" and 'Sub Function' in current_view_df.columns:
        current_view_df = current_view_df[current_view_df['Sub Function'] == selected_sub]
        
    valid_ids_current = [str(vid).replace('.0', '').strip() for vid in current_view_df['Employee Code'].tolist()]
    
    js_rows_current = []
    for index, row in current_view_df.iterrows():
        emp_id = str(row.get('Employee Code', '')).replace('.0', '').strip()
        manager_id = str(row.get('L1 Manager Code', '')).replace('.0', '').strip()
        name = str(row.get('Employee Name', '')).strip()
        grade = str(row.get('Grade', '')).strip()
        designation = str(row.get('Designation', '')).strip()
        onroll = str(row.get('Onroll Reportees', '0')).strip()
        sub_func = str(row.get('Sub Function', '')).strip()
        
        if manager_id not in valid_ids_current: manager_id = ""
            
        # Draft Mode Visual Cue
        is_moved = emp_id in st.session_state.draft_moves
        top_border = "border-top: 4px solid #8b5cf6;" if is_moved else "border-top: 4px solid #10b981;"
        bg_tint = "background-color: rgba(139, 92, 246, 0.05);" if is_moved else ""
        
        box_html = f"<div class='beautiful-card' style='{top_border} {bg_tint}'><div class='card-header'><span class='badge'>{sub_func[:15]}</span><span class='grade'>GR: {grade}</span></div><div class='card-body'><div class='card-name'>{name}</div><div class='card-title'>{designation}</div></div><div class='card-footer'><div class='stat'><span>On-Roll</span><b>{onroll}</b></div><div class='stat'><span>Appr HC</span><b style='color:#ea580c;'>-</b></div><div class='stat'><span>Off-Roll</span><b style='color:#65a30d;'>-</b></div></div></div>"
        box_html_clean = box_html.replace('\n', '').replace('\r', '').replace("'", "\\'")
        mgr_str = f"'{manager_id}'" if manager_id else "''"
        js_rows_current.append(f"[{{'v': '{emp_id}', 'f': \"{box_html_clean}\"}}, {mgr_str}, '']")

    all_rows_formatted = ",\n".join(js_rows_current)
    safe_filename = selected_sub.replace(" ", "_") + "_Org_Chart" if selected_sub != "All" else "Full_Organization_Chart"

    # 6. HTML Template with Google Charts
    html_template = f"""
    <html>
      <head>
        <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        
        <script type="text/javascript">
          const allDataMap = {all_data_dict_str};
          var chart;
          
          google.charts.load('current', {{packages:["orgchart"]}});
          google.charts.setOnLoadCallback(drawChart);
          
          function drawChart() {{
            var data = new google.visualization.DataTable();
            data.addColumn('string', 'Name');
            data.addColumn('string', 'Manager');
            data.addColumn('string', 'ToolTip');
            data.addRows([{all_rows_formatted}]);
            
            chart = new google.visualization.OrgChart(document.getElementById('chart_div'));
            chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'large', nodeClass:'myNode', selectedNodeClass:'selectedNode'}});
          }}
          
          function downloadImage() {{
              const chartContainer = document.getElementById("chart_div");
              html2canvas(chartContainer, {{ 
                  backgroundColor: "#0f172a",
                  scale: 2
              }}).then(canvas => {{
                  let link = document.createElement('a');
                  link.download = '{safe_filename}.png';
                  link.href = canvas.toDataURL("image/png");
                  link.click();
              }});
          }}
          
          async function downloadAllZip() {{
              const btn = document.getElementById('zip-btn');
              const originalText = btn.innerHTML;
              btn.innerHTML = '⏳ Generating ZIP...'; btn.style.pointerEvents = 'none'; btn.style.opacity = '0.7';
              
              var zip = new JSZip();
              var subFunctions = Object.keys(allDataMap);
              
              for (let i = 0; i < subFunctions.length; i++) {{
                  let sub = subFunctions[i];
                  let rows = allDataMap[sub];
                  
                  var data = new google.visualization.DataTable();
                  data.addColumn('string', 'Name'); data.addColumn('string', 'Manager'); data.addColumn('string', 'ToolTip');
                  data.addRows(rows);
                  chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'large', nodeClass:'myNode'}});
                  
                  await new Promise(r => setTimeout(r, 800));
                  
                  const chartContainer = document.getElementById("chart_div");
                  const canvas = await html2canvas(chartContainer, {{ backgroundColor: "#0f172a", scale: 2 }});
                  
                  const imgData = canvas.toDataURL("image/png").split(',')[1];
                  zip.file(sub.replace(/ /g, "_") + "_Org_Chart.png", imgData, {{base64: true}});
              }}
              
              drawChart();
              
              btn.innerHTML = '📦 Wrapping ZIP...';
              zip.generateAsync({{type:"blob"}}).then(function(content) {{
                  saveAs(content, "All_HR_Org_Charts.zip");
                  btn.innerHTML = originalText; btn.style.pointerEvents = 'auto'; btn.style.opacity = '1';
              }});
          }}
       </script>
       <style>
         @keyframes slideUpFade {{ 0% {{ opacity: 0; transform: translateY(20px); }} 100% {{ opacity: 1; transform: translateY(0); }} }}
         @keyframes pulseGlow {{ 0%, 100% {{ box-shadow: 0 0 20px rgba(139, 92, 246, 0.3); }} 50% {{ box-shadow: 0 0 30px rgba(139, 92, 246, 0.5); }} }}
         
         body {{ 
             margin: 0; padding: 0; 
             background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%);
             font-family: 'Inter', sans-serif;
             min-height: 100vh;
             color: #e2e8f0;
         }}
         
         /* Org Chart Lines */
         .google-visualization-orgchart-lineleft, 
         .google-visualization-orgchart-lineright, 
         .google-visualization-orgchart-linebottom {{ 
             border-color: rgba(148, 163, 184, 0.4) !important; 
             border-width: 2px !important; 
         }}
         
         /* Node Styling */
         .myNode {{ 
             border: none !important; 
             background: none !important; 
             padding: 0 !important; 
             box-shadow: none !important; 
             margin: 16px; 
             cursor: pointer; 
         }}
         
         .selectedNode .beautiful-card {{
             border-color: #8b5cf6 !important;
             box-shadow: 0 0 20px rgba(139, 92, 246, 0.4) !important;
         }}
         
         .beautiful-card {{ 
             border-radius: 14px; 
             background: rgba(30, 41, 59, 0.8); 
             border: 1px solid rgba(139, 92, 246, 0.2); 
             position: relative; 
             width: 280px; 
             overflow: hidden; 
             box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3); 
             animation: slideUpFade 0.5s forwards; 
             transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); 
             backdrop-filter: blur(10px);
         }}
         .beautiful-card:hover {{ 
             transform: translateY(-6px) scale(1.02); 
             box-shadow: 0 12px 32px rgba(139, 92, 246, 0.25); 
             border-color: rgba(139, 92, 246, 0.5); 
             z-index: 10; 
         }}
         
         .card-header {{ 
             padding: 14px 18px; 
             display: flex; 
             justify-content: space-between; 
             align-items: center; 
             background: linear-gradient(180deg, rgba(139, 92, 246, 0.1) 0%, rgba(30, 41, 59, 0.4) 100%); 
             border-bottom: 1px solid rgba(148, 163, 184, 0.2); 
         }}
         .badge {{ 
             background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%); 
             color: #ffffff; 
             padding: 5px 12px; 
             border-radius: 20px; 
             font-size: 11px; 
             font-weight: 700; 
             text-transform: uppercase; 
             letter-spacing: 0.5px; 
             box-shadow: 0 2px 8px rgba(139, 92, 246, 0.3);
         }}
         .grade {{ 
             color: #94a3b8; 
             font-size: 12px; 
             font-weight: 600; 
             background: rgba(148, 163, 184, 0.1);
             padding: 4px 10px;
             border-radius: 6px;
         }}
         
         .card-body {{ 
             padding: 18px; 
             text-align: left; 
         }}
         .card-name {{ 
             font-size: 16px; 
             font-weight: 700; 
             color: #ffffff; 
             margin-bottom: 6px; 
         }}
         .card-title {{ 
             font-size: 13px; 
             color: #94a3b8; 
             font-weight: 500; 
             line-height: 1.5; 
         }}
         
         .card-footer {{ 
             background: rgba(15, 23, 42, 0.5); 
             border-top: 1px solid rgba(148, 163, 184, 0.15); 
             padding: 12px 18px; 
             display: flex; 
             justify-content: space-between; 
         }}
         .stat {{ 
             display: flex; 
             flex-direction: column; 
             align-items: center; 
             gap: 4px; 
         }}
         .stat span {{ 
             font-size: 10px; 
             color: #64748b; 
             text-transform: uppercase; 
             font-weight: 700; 
             letter-spacing: 0.5px; 
         }}
         .stat b {{ 
             font-size: 14px; 
             color: #e2e8f0; 
         }}
         
         /* Buttons */
         .btn-container {{ 
             display: flex; 
             gap: 16px; 
             margin-bottom: 28px; 
             flex-wrap: wrap; 
             padding: 0 4px;
         }}
         .download-btn {{ 
             background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%); 
             color: #ffffff; 
             border: none; 
             padding: 14px 28px; 
             border-radius: 10px; 
             cursor: pointer; 
             font-weight: 600; 
             font-size: 15px; 
             transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); 
             display: inline-flex; 
             align-items: center; 
             box-shadow: 0 4px 15px rgba(139, 92, 246, 0.4); 
             font-family: 'Inter', sans-serif;
         }}
         .download-btn:hover {{ 
             transform: translateY(-2px); 
             box-shadow: 0 8px 25px rgba(139, 92, 246, 0.5); 
         }}
         .download-btn:active {{
             transform: translateY(0) scale(0.98);
         }}
         
         .zip-btn {{ 
             background: linear-gradient(135deg, #10b981 0%, #059669 100%); 
             box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4); 
         }}
         .zip-btn:hover {{ 
             box-shadow: 0 8px 25px rgba(16, 185, 129, 0.5); 
         }}
         
         /* Chart Container */
         #scroll_wrapper {{ 
             overflow-x: auto; 
             width: 100%; 
             padding: 40px 20px; 
             background: rgba(15, 23, 42, 0.4); 
             border-radius: 20px; 
             border: 1px solid rgba(139, 92, 246, 0.2); 
             backdrop-filter: blur(10px);
         }}
         #chart_div {{ 
             display: inline-block; 
             min-width: 100%; 
             padding: 20px; 
             background-color: transparent; 
         }}
         
         /* Scrollbar */
         ::-webkit-scrollbar {{
             width: 10px;
             height: 10px;
         }}
         ::-webkit-scrollbar-track {{
             background: rgba(30, 41, 59, 0.5);
             border-radius: 10px;
         }}
         ::-webkit-scrollbar-thumb {{
             background: linear-gradient(180deg, #8b5cf6, #6366f1);
             border-radius: 10px;
         }}
         ::-webkit-scrollbar-thumb:hover {{
             background: linear-gradient(180deg, #7c3aed, #4f46e5);
         }}
       </style>
      </head>
      <body>
        <div class="btn-container">
            <button class="download-btn" onclick="downloadImage()">⬇️ Download Current View</button>
            <button class="download-btn zip-btn" id="zip-btn" onclick="downloadAllZip()">📦 Download All Charts ZIP</button>
        </div>
        <div id="scroll_wrapper">
            <div id="chart_div"></div>
        </div>
      </body>
    </html>
    """
    
    components.html(html_template, height=950, scrolling=True)
