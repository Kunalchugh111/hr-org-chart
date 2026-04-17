import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. Page Configuration
st.set_page_config(page_title="HR Org Design", layout="wide", initial_sidebar_state="expanded")

# --- SPECTACULAR UI/UX OVERHAUL ---
st.markdown("""
    <style>
        /* Import Premium Web Font */
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&display=swap');

        /* Hide Streamlit Chrome to look like a real app */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Global Font Override */
        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', -apple-system, sans-serif !important;
        }

        /* Main App Background - Soft, Colorful Mesh Gradient */
        .stApp {
            background-color: #f8fafc;
            background-image:
                radial-gradient(at 10% 0%, hsla(142, 69%, 92%, 1) 0px, transparent 50%),
                radial-gradient(at 90% 10%, hsla(28, 100%, 90%, 1) 0px, transparent 50%),
                radial-gradient(at 50% 100%, hsla(210, 100%, 95%, 1) 0px, transparent 50%);
            background-attachment: fixed;
        }

        /* Clean, Bold Typography for the Main Dashboard */
        h1 {
            font-weight: 800 !important;
            color: #0f172a !important;
            letter-spacing: -1.5px;
            font-size: 2.8rem !important;
            margin-bottom: 1rem !important;
        }

        /* Sleek, Dark "Pro" Sidebar */
        [data-testid="stSidebar"] {
            background-color: #0f172a !important;
            border-right: 1px solid #1e293b;
        }
        
        /* Force Sidebar Text to White */
        [data-testid="stSidebar"] * {
            color: #f8fafc !important;
        }
        
        /* Sidebar Dropdown Styling */
        [data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background-color: #1e293b !important;
            border: 1px solid #334155 !important;
            border-radius: 8px !important;
            color: white !important;
        }
        
        /* File Uploader - Frosted Glass Look */
        [data-testid="stFileUploadDropzone"] {
            background: rgba(255, 255, 255, 0.4) !important;
            backdrop-filter: blur(10px) !important;
            border: 1px solid rgba(255, 255, 255, 0.8) !important;
            border-radius: 16px !important;
            box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.05) !important;
            transition: all 0.3s ease;
        }
        [data-testid="stFileUploadDropzone"]:hover {
            border-color: #22c55e !important;
            background: rgba(255, 255, 255, 0.6) !important;
            transform: translateY(-2px);
        }

        /* Main Canvas Adjustments */
        .block-container { 
            padding-top: 3rem; 
            padding-bottom: 2rem; 
            max-width: 95%;
        }
    </style>
""", unsafe_allow_html=True)

# 2. Sidebar Controls
with st.sidebar:
    st.markdown("### ⚡ Architecture Generator")
    st.markdown("Upload your roster to map the organizational structure.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload HR Data", type=['csv', 'xlsx'])
    
    selected_sub = "All"
    include_retainers = True
    include_inactive = False
    
    if uploaded_file is not None:
        st.markdown("<br>### 🎛️ Data Filters", unsafe_allow_html=True)
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()
            
        if 'Sub Function' in df.columns:
            sub_functions = df['Sub Function'].dropna().unique()
            selected_sub = st.selectbox("Sub Function View", ["All"] + list(sub_functions))
        else:
            sub_functions = []
        
        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)

# 3. Main Dashboard Area & Data Processing
if uploaded_file is None:
    st.title("Organizational Architecture")
    st.info("👈 Please upload your HR data in the sidebar to begin building the map.")
else:
    st.title(f"Organizational Architecture: {selected_sub if selected_sub != 'All' else 'Full Organization'}")
    
    base_filtered_df = df.copy()
    if not include_retainers and 'Employment Type' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[~base_filtered_df['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
    if not include_inactive and 'Status' in base_filtered_df.columns:
        base_filtered_df = base_filtered_df[~base_filtered_df['Status'].astype(str).str.contains('Inactive', case=False, na=False)]

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
                
                if manager_id not in valid_ids_sub:
                    manager_id = ""
                    
                box_html = f"<div class='beautiful-card'><div class='card-header'><span class='badge'>{sub_func[:15]}</span><span class='grade'>GR: {grade}</span></div><div class='card-body'><div class='card-name'>{name}</div><div class='card-title'>{designation}</div></div><div class='card-footer'><div class='stat'><span>On-Roll</span><b>{onroll}</b></div><div class='stat'><span>Appr HC</span><b style='color:#ea580c;'>-</b></div><div class='stat'><span>Off-Roll</span><b style='color:#65a30d;'>-</b></div></div></div>"
                box_html_clean = box_html.replace('\n', '').replace('\r', '')
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
        
        if manager_id not in valid_ids_current:
            manager_id = ""
            
        box_html = f"<div class='beautiful-card'><div class='card-header'><span class='badge'>{sub_func[:15]}</span><span class='grade'>GR: {grade}</span></div><div class='card-body'><div class='card-name'>{name}</div><div class='card-title'>{designation}</div></div><div class='card-footer'><div class='stat'><span>On-Roll</span><b>{onroll}</b></div><div class='stat'><span>Appr HC</span><b style='color:#ea580c;'>-</b></div><div class='stat'><span>Off-Roll</span><b style='color:#65a30d;'>-</b></div></div></div>"
        box_html_clean = box_html.replace('\n', '').replace('\r', '')
        mgr_str = f"'{manager_id}'" if manager_id else "''"
        js_rows_current.append(f"[{{'v': '{emp_id}', 'f': \"{box_html_clean}\"}}, {mgr_str}, '']")

    all_rows_formatted = ",\n".join(js_rows_current)
    safe_filename = selected_sub.replace(" ", "_") + "_Org_Chart" if selected_sub != "All" else "Full_Organization_Chart"

    # 5. HTML Template with JSZip integration (UNCHANGED CORE LOGIC)
    html_template = f"""
    <html>
      <head>
        <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
        
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
            chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'large', nodeClass:'myNode'}});
          }}
          
          function downloadImage() {{
              const chartContainer = document.getElementById("chart_div");
              html2canvas(chartContainer, {{ 
                  backgroundColor: "#ffffff",
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
              btn.innerHTML = '⏳ Generating ZIP...';
              btn.style.pointerEvents = 'none';
              btn.style.opacity = '0.7';
              
              var zip = new JSZip();
              var subFunctions = Object.keys(allDataMap);
              
              for (let i = 0; i < subFunctions.length; i++) {{
                  let sub = subFunctions[i];
                  let rows = allDataMap[sub];
                  
                  var data = new google.visualization.DataTable();
                  data.addColumn('string', 'Name');
                  data.addColumn('string', 'Manager');
                  data.addColumn('string', 'ToolTip');
                  data.addRows(rows);
                  chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'large', nodeClass:'myNode'}});
                  
                  await new Promise(r => setTimeout(r, 800));
                  
                  const chartContainer = document.getElementById("chart_div");
                  const canvas = await html2canvas(chartContainer, {{ 
                      backgroundColor: "#ffffff", 
                      scale: 2 
                  }});
                  
                  const imgData = canvas.toDataURL("image/png").split(',')[1];
                  zip.file(sub.replace(/ /g, "_") + "_Org_Chart.png", imgData, {{base64: true}});
              }}
              
              drawChart();
              
              btn.innerHTML = '📦 Wrapping ZIP...';
              zip.generateAsync({{type:"blob"}}).then(function(content) {{
                  saveAs(content, "All_HR_Org_Charts.zip");
                  btn.innerHTML = 'Create All & Download ZIP';
                  btn.style.pointerEvents = 'auto';
                  btn.style.opacity = '1';
              }});
          }}
       </script>
       <style>
         @keyframes slideUpFade {{ 0% {{ opacity: 0; transform: translateY(20px); }} 100% {{ opacity: 1; transform: translateY(0); }} }}
         /* Internal HTML Font matching Streamlit */
         body {{ margin: 0; padding: 0; background-color: transparent; font-family: 'Plus Jakarta Sans', -apple-system, sans-serif; }}
         .google-visualization-orgchart-lineleft, .google-visualization-orgchart-lineright, .google-visualization-orgchart-linebottom {{ border-color: #cbd5e1 !important; border-width: 2px !important; }}
         .myNode {{ border: none !important; background: none !important; padding: 0 !important; box-shadow: none !important; margin: 12px; }}
         
         .beautiful-card {{ border-radius: 12px; background: #ffffff; border: 1px solid #dcfce7; border-top: 4px solid #22c55e; position: relative; width: 250px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(34, 197, 94, 0.05); animation: slideUpFade 0.5s forwards; transition: all 0.3s ease; }}
         .beautiful-card:hover {{ transform: translateY(-4px); box-shadow: 0 12px 20px -5px rgba(34, 197, 94, 0.15); border-color: #86efac; z-index: 10; }}
         .card-header {{ padding: 12px 16px; display: flex; justify-content: space-between; align-items: center; background: linear-gradient(180deg, #f0fdf4 0%, #ffffff 100%); border-bottom: 1px solid #f1f5f9; }}
         .badge {{ background: #dcfce7; color: #166534; padding: 4px 10px; border-radius: 20px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }}
         .grade {{ color: #64748b; font-size: 11px; font-weight: 600; }}
         .card-body {{ padding: 16px; text-align: left; }}
         .card-name {{ font-size: 15px; font-weight: 700; color: #0f172a; margin-bottom: 4px; }}
         .card-title {{ font-size: 12px; color: #475569; font-weight: 400; line-height: 1.4; }}
         .card-footer {{ background: #f8fafc; border-top: 1px solid #f1f5f9; padding: 10px 16px; display: flex; justify-content: space-between; }}
         .stat {{ display: flex; flex-direction: column; align-items: center; gap: 2px; }}
         .stat span {{ font-size: 9px; color: #64748b; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; }}
         .stat b {{ font-size: 13px; color: #0f172a; }}
         
         .btn-container {{ display: flex; gap: 15px; margin-bottom: 24px; flex-wrap: wrap; }}
         .download-btn {{ background: #0f172a; color: #ffffff; border: none; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 14px; transition: all 0.2s; display: inline-flex; align-items: center; box-shadow: 0 4px 6px -1px rgba(15, 23, 42, 0.2); }}
         .download-btn:hover {{ background: #1e293b; transform: translateY(-1px); }}
         .zip-btn {{ background: #16a34a; box-shadow: 0 4px 6px -1px rgba(22, 163, 74, 0.2); }}
         .zip-btn:hover {{ background: #15803d; }}
         
         #scroll_wrapper {{ overflow-x: auto; width: 100%; padding: 40px 20px; background-color: rgba(255,255,255,0.7); border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: inset 0 2px 4px rgba(0,0,0,0.02); }}
         #chart_div {{ display: inline-block; min-width: 100%; padding: 20px; background-color: transparent; }}
       </style>
      </head>
      <body>
        <div class="btn-container">
            <button class="download-btn" onclick="downloadImage()">Download Current View</button>
            <button class="download-btn zip-btn" id="zip-btn" onclick="downloadAllZip()">Create All & Download ZIP</button>
        </div>
        <div id="scroll_wrapper">
            <div id="chart_div"></div>
        </div>
      </body>
    </html>
    """
    
    components.html(html_template, height=900, scrolling=True)
