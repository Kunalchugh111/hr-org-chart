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
                radial-gradient(circle at 10% 20%, rgba(139, 92, 246, 0.05) 0%, transparent 40%),
                radial-gradient(circle at 90% 80%, rgba(59, 130, 246, 0.05) 0%, transparent 40%),
                radial-gradient(circle at 50% 50%, rgba(16, 185, 129, 0.03) 0%, transparent 50%);
            min-height: 100vh;
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
            color: #0f172a !important;
            letter-spacing: -0.03em !important;
            font-size: 2.8rem !important;
            margin-bottom: 0.5rem !important;
        }
        
        h2, h3 {
            color: #1e293b !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em !important;
        }

        /* Premium Light Sidebar */
        [data-testid="stSidebar"] {
            background: #ffffff !important;
            border-right: 1px solid #e5e7eb !important;
            box-shadow: 4px 0 20px rgba(0, 0, 0, 0.02) !important;
        }
        [data-testid="stSidebar"] * { color: #374151 !important; }
        [data-testid="stSidebar"] h3 {
            color: #111827 !important;
            font-size: 1.1rem !important;
            margin-bottom: 1rem !important;
            padding-bottom: 0.5rem !important;
            border-bottom: 2px solid #f3f4f6 !important;
        }

        /* Form Elements in Sidebar */
        [data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background-color: #f9fafb !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 10px !important;
            color: #374151 !important;
        }
        [data-testid="stSidebar"] div[data-baseweb="select"]:hover > div {
            border-color: #8b5cf6 !important;
            background-color: #ffffff !important;
        }
        
        /* File Uploader */
        [data-testid="stFileUploadDropzone"] {
            background: linear-gradient(135deg, #fafafa 0%, #ffffff 100%) !important;
            border: 2px dashed #d1d5db !important;
            border-radius: 14px !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            min-height: 140px !important;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.02);
        }
        [data-testid="stFileUploadDropzone"]:hover {
            border-color: #8b5cf6 !important;
            background: linear-gradient(135deg, #faf5ff 0%, #ffffff 100%) !important;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(139, 92, 246, 0.08);
        }
        [data-testid="stFileUploadDropzone"] p { color: #4b5563 !important; }

        /* Toggles */
        [data-testid="stToggle"] > label > div > div {
            background-color: #f3f4f6 !important;
            border-color: #d1d5db !important;
        }
        [data-testid="stToggle"] > label > div > div[aria-checked="true"] {
            background-color: #8b5cf6 !important;
            border-color: #8b5cf6 !important;
            box-shadow: 0 2px 8px rgba(139, 92, 246, 0.25);
        }

        /* Buttons */
        .stButton > button {
            background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 12px !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
            padding: 0.8rem 1.5rem !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
            box-shadow: 0 4px 12px rgba(139, 92, 246, 0.25), 0 1px 2px rgba(0,0,0,0.05) !important;
            width: 100% !important;
        }
        .stButton > button:hover {
            transform: translateY(-2px) scale(1.01) !important;
            box-shadow: 0 6px 20px rgba(139, 92, 246, 0.35), 0 2px 4px rgba(0,0,0,0.08) !important;
            background: linear-gradient(135deg, #9061f9 0%, #8b5cf6 100%) !important;
        }
        .stButton > button:active {
            transform: translateY(0) scale(0.99) !important;
        }
        .stButton > button[type="secondary"] {
            background: #f9fafb !important;
            color: #4b5563 !important;
            border: 1px solid #e5e7eb !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important;
        }
        .stButton > button[type="secondary"]:hover {
            background: #f3f4f6 !important;
            color: #111827 !important;
            border-color: #d1d5db !important;
        }

        /* Select Box in Main Area */
        .stSelectbox > div > div {
            background-color: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 10px !important;
            color: #374151 !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03) !important;
        }
        .stSelectbox > div > div:hover {
            border-color: #8b5cf6 !important;
        }

        /* Info/Success/Error Boxes */
        .stAlert {
            background: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 12px !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important;
            color: #374151 !important;
        }
        .stAlert div { color: #374151 !important; }
        .stAlert > div { background-color: transparent !important; }

        /* Metrics/Stats Cards */
        [data-testid="stMetric"] {
            background: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            border-radius: 12px !important;
            padding: 1rem !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important;
        }

        /* Divider */
        .stDivider {
            border-color: #f3f4f6 !important;
        }

        /* Tooltip */
        [data-testid="stTooltipHoverContent"] {
            background-color: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            color: #111827 !important;
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
    st.markdown("### 🏢 OrgDesign Pro")
    st.markdown("<p style='color: #6b7280; font-size: 0.85rem; margin-top: -0.5rem; margin-bottom: 1.5rem;'>Enterprise Organizational Architecture Platform</p>", unsafe_allow_html=True)
    
    st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
    st.markdown("##### 📊 Upload HR Data")
    st.markdown("<p style='color: #6b7280; font-size: 0.8rem; margin-top: -0.5rem;'>CSV or Excel roster files supported</p>", unsafe_allow_html=True)
    
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
        st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
        st.markdown("##### 🔍 Data Filters")
        
        if 'Sub Function' in df.columns:
            sub_functions = sorted(df['Sub Function'].dropna().unique())
            selected_sub = st.selectbox("Sub Function View", ["All"] + sub_functions, index=0)
        
        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div style='background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 1rem; margin-bottom: 1rem;'>", unsafe_allow_html=True)
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
                st.markdown(f"<div style='background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 10px; padding: 0.6rem; text-align: center; margin-top: 0.75rem;'><p style='color: #059669; font-size: 0.85rem; margin: 0; font-weight: 600;'>✨ {len(st.session_state.draft_moves)} active moves</p></div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)

# 5. Main Dashboard Area
if not st.session_state.file_uploaded:
    st.markdown("<div style='text-align: center; padding: 4rem 2rem; max-width: 900px; margin: 0 auto;'>", unsafe_allow_html=True)
    st.title("🏢 OrgDesign Pro")
    st.markdown("<p style='color: #4b5563; font-size: 1.2rem; margin-top: 0.5rem; margin-bottom: 3rem;'>Transform your HR data into stunning, interactive organizational charts</p>", unsafe_allow_html=True)
    
    st.markdown("""
        <div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem; max-width: 800px; margin: 0 auto; text-align: left;'>
            <div style='background: #ffffff; border: 1px solid #e5e7eb; border-radius: 16px; padding: 1.5rem; box-shadow: 0 4px 12px rgba(0,0,0,0.03); transition: all 0.3s;'>
                <div style='font-size: 2.2rem; margin-bottom: 0.75rem;'>📊</div>
                <h3 style='color: #111827; margin: 0 0 0.5rem 0; font-size: 1.15rem;'>Smart Parsing</h3>
                <p style='color: #6b7280; font-size: 0.88rem; margin: 0; line-height: 1.5;'>Automatically detects reporting hierarchies and structures complex org trees.</p>
            </div>
            <div style='background: #ffffff; border: 1px solid #e5e7eb; border-radius: 16px; padding: 1.5rem; box-shadow: 0 4px 12px rgba(0,0,0,0.03); transition: all 0.3s;'>
                <div style='font-size: 2.2rem; margin-bottom: 0.75rem;'>🎨</div>
                <h3 style='color: #111827; margin: 0 0 0.5rem 0; font-size: 1.15rem;'>Premium Design</h3>
                <p style='color: #6b7280; font-size: 0.88rem; margin: 0; line-height: 1.5;'>Beautiful, modern cards with rich details, shadows, and interactive states.</p>
            </div>
            <div style='background: #ffffff; border: 1px solid #e5e7eb; border-radius: 16px; padding: 1.5rem; box-shadow: 0 4px 12px rgba(0,0,0,0.03); transition: all 0.3s;'>
                <div style='font-size: 2.2rem; margin-bottom: 0.75rem;'>🧪</div>
                <h3 style='color: #111827; margin: 0 0 0.5rem 0; font-size: 1.15rem;'>Scenario Mode</h3>
                <p style='color: #6b7280; font-size: 0.88rem; margin: 0; line-height: 1.5;'>Model org changes safely before implementing them in production.</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div style='margin-top: 3rem; padding: 1.5rem; background: #ffffff; border: 1px dashed #d1d5db; border-radius: 16px; max-width: 600px; margin-left: auto; margin-right: auto; box-shadow: 0 2px 8px rgba(0,0,0,0.02);'>
            <p style='color: #374151; margin: 0; font-size: 0.95rem;'>
                <strong style='color: #7c3aed;'>👈 Get Started:</strong> Upload your HR roster (CSV/XLSX) from the sidebar to begin mapping your organizational structure.
            </p>
        </div>
    """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.title(f"Org Architecture: {selected_sub if selected_sub != 'All' else 'Full Organization'}")
    
    st.markdown("""
        <div style='display: flex; align-items: center; gap: 0.5rem; margin-top: -10px; margin-bottom: 1.5rem;'>
            <span style='background: #fef3c7; border: 1px solid #fde68a; padding: 4px 10px; border-radius: 8px; font-size: 0.85rem; color: #b45309; font-weight: 500;'>💡 Tip: Double-click any manager's card to collapse or expand their reporting lines.</span>
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
        
        box_html = f"<div class='beautiful-card' style='{top_border}'><div class='card-header'><span class='badge'>{sub_func[:15]}</span><span class='grade'>GR: {grade}</span></div><div class='card-body'><div class='card-name'>{name}</div><div class='card-title'>{designation}</div></div><div class='card-footer'><div class='stat'><span>On-Roll</span><b>{onroll}</b></div><div class='stat'><span>Appr HC</span><b style='color:#ea580c;'>-</b></div><div class='stat'><span>Off-Roll</span><b style='color:#65a30d;'>-</b></div></div></div>"
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
                  backgroundColor: "#fafafa",
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
                  const canvas = await html2canvas(chartContainer, {{ backgroundColor: "#fafafa", scale: 2 }});
                  
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
         @keyframes slideUpFade {{ 0% {{ opacity: 0; transform: translateY(15px); }} 100% {{ opacity: 1; transform: translateY(0); }} }}
         @keyframes softPulse {{ 0%, 100% {{ transform: translateY(0); }} 50% {{ transform: translateY(-2px); }} }}
         
         body {{ 
             margin: 0; padding: 0; 
             background: linear-gradient(180deg, #fafafa 0%, #ffffff 100%);
             font-family: 'Inter', sans-serif;
             min-height: 100vh;
             color: #111827;
         }}
         
         /* Org Chart Lines - Premium Light */
         .google-visualization-orgchart-lineleft, 
         .google-visualization-orgchart-lineright, 
         .google-visualization-orgchart-linebottom {{ 
             border-color: #d1d5db !important; 
             border-width: 2px !important; 
         }}
         .google-visualization-orgchart-node-medium {{
             padding: 0 !important;
         }}
         
         /* Node Styling */
         .myNode {{ 
             border: none !important; 
             background: none !important; 
             padding: 0 !important; 
             box-shadow: none !important; 
             margin: 20px; 
             cursor: pointer; 
         }}
         
         .selectedNode .beautiful-card {{
             border-color: #8b5cf6 !important;
             box-shadow: 0 8px 24px rgba(139, 92, 246, 0.2), 0 0 0 3px rgba(139, 92, 246, 0.15) !important;
             transform: translateY(-2px);
         }}
         
         .beautiful-card {{ 
             border-radius: 16px; 
             background: #ffffff; 
             border: 1px solid #e5e7eb; 
             position: relative; 
             width: 280px; 
             overflow: hidden; 
             box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04), 0 1px 2px rgba(0,0,0,0.02); 
             animation: slideUpFade 0.4s ease-out forwards; 
             transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); 
         }}
         .beautiful-card:hover {{ 
             transform: translateY(-4px); 
             box-shadow: 0 12px 24px rgba(0, 0, 0, 0.08), 0 4px 8px rgba(139, 92, 246, 0.12); 
             border-color: #c7d2fe; 
             z-index: 10; 
         }}
         
         .card-header {{ 
             padding: 16px 18px; 
             display: flex; 
             justify-content: space-between; 
             align-items: center; 
             background: linear-gradient(180deg, #f8f9fb 0%, #ffffff 100%); 
             border-bottom: 1px solid #f3f4f6; 
         }}
         .badge {{ 
             background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%); 
             color: #ffffff; 
             padding: 6px 12px; 
             border-radius: 8px; 
             font-size: 11px; 
             font-weight: 700; 
             text-transform: uppercase; 
             letter-spacing: 0.6px; 
             box-shadow: 0 2px 6px rgba(139, 92, 246, 0.25);
         }}
         .grade {{ 
             color: #6b7280; 
             font-size: 12px; 
             font-weight: 600; 
             background: #f3f4f6;
             padding: 5px 10px;
             border-radius: 6px;
             border: 1px solid #e5e7eb;
         }}
         
         .card-body {{ 
             padding: 20px 18px 14px; 
             text-align: left; 
         }}
         .card-name {{ 
             font-size: 16px; 
             font-weight: 700; 
             color: #0f172a; 
             margin-bottom: 6px; 
             line-height: 1.3;
         }}
         .card-title {{ 
             font-size: 13px; 
             color: #4b5563; 
             font-weight: 500; 
             line-height: 1.5; 
         }}
         
         .card-footer {{ 
             background: #f9fafb; 
             border-top: 1px solid #f3f4f6; 
             padding: 14px 18px; 
             display: flex; 
             justify-content: space-between; 
         }}
         .stat {{ 
             display: flex; 
             flex-direction: column; 
             align-items: center; 
             gap: 5px; 
         }}
         .stat span {{ 
             font-size: 10px; 
             color: #6b7280; 
             text-transform: uppercase; 
             font-weight: 700; 
             letter-spacing: 0.5px; 
         }}
         .stat b {{ 
             font-size: 14px; 
             color: #0f172a; 
             font-weight: 600;
         }}
         
         /* Buttons */
         .btn-container {{ 
             display: flex; 
             gap: 16px; 
             margin-bottom: 32px; 
             flex-wrap: wrap; 
             padding: 0 8px;
         }}
         .download-btn {{ 
             background: linear-gradient(135deg, #ffffff 0%, #f9fafb 100%); 
             color: #374151; 
             border: 1px solid #d1d5db; 
             padding: 14px 26px; 
             border-radius: 12px; 
             cursor: pointer; 
             font-weight: 600; 
             font-size: 15px; 
             transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); 
             display: inline-flex; 
             align-items: center; 
             gap: 8px;
             box-shadow: 0 1px 3px rgba(0,0,0,0.04); 
             font-family: 'Inter', sans-serif;
         }}
         .download-btn:hover {{ 
             background: linear-gradient(135deg, #f9fafb 0%, #ffffff 100%);
             border-color: #c0c4c8;
             box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
             transform: translateY(-2px);
         }}
         .download-btn:active {{
             transform: translateY(0);
         }}
         
         .zip-btn {{ 
             background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%); 
             color: #ffffff;
             border: none;
             box-shadow: 0 4px 12px rgba(139, 92, 246, 0.25), 0 1px 2px rgba(0,0,0,0.05);
         }}
         .zip-btn:hover {{ 
             box-shadow: 0 6px 18px rgba(139, 92, 246, 0.35);
             transform: translateY(-2px);
         }}
         
         /* Chart Container */
         #scroll_wrapper {{ 
             overflow-x: auto; 
             width: 100%; 
             padding: 40px 24px 48px; 
             background: #ffffff; 
             border-radius: 20px; 
             border: 1px solid #e5e7eb; 
             box-shadow: inset 0 1px 3px rgba(0,0,0,0.02), 0 2px 8px rgba(0,0,0,0.03);
         }}
         #chart_div {{ 
             display: inline-block; 
             min-width: 100%; 
             padding: 12px; 
             background-color: transparent; 
         }}
         
         /* Scrollbar */
         ::-webkit-scrollbar {{
             width: 12px;
             height: 12px;
         }}
         ::-webkit-scrollbar-track {{
             background: #f3f4f6;
             border-radius: 12px;
         }}
         ::-webkit-scrollbar-thumb {{
             background: #d1d5db;
             border-radius: 12px;
             border: 3px solid #f3f4f6;
         }}
         ::-webkit-scrollbar-thumb:hover {{
             background: #9ca3af;
         }}
       </style>
      </head>
      <body>
        <div class="btn-container">
            <button class="download-btn" onclick="downloadImage()">⬇️ Download View</button>
            <button class="download-btn zip-btn" id="zip-btn" onclick="downloadAllZip()">📦 Export All Charts</button>
        </div>
        <div id="scroll_wrapper">
            <div id="chart_div"></div>
        </div>
      </body>
    </html>
    """
    
    components.html(html_template, height=950, scrolling=True)
