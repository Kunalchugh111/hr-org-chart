import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. Page Configuration (Light Theme)
st.set_page_config(page_title="HR Org Design", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 0rem; }
        h1 { font-family: -apple-system, BlinkMacSystemFont, sans-serif; font-weight: 800; color: #0f172a; }
        .stApp { background-color: #ffffff; color: #0f172a; }
    </style>
""", unsafe_allow_html=True)

# 2. Sidebar Controls
with st.sidebar:
    st.markdown("### 🌿 Architecture Generator")
    st.markdown("Upload your roster to generate the dynamic structure.")
    st.markdown("---")
    
    uploaded_file = st.file_uploader("Upload HR Data", type=['csv', 'xlsx'])
    
    selected_sub = "All"
    include_retainers = True
    include_inactive = False
    
    if uploaded_file is not None:
        st.markdown("### 🎛️ Filters")
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip()
            
        if 'Sub Function' in df.columns:
            sub_functions = df['Sub Function'].dropna().unique()
            selected_sub = st.selectbox("Sub Function", ["All"] + list(sub_functions))
        
        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)

# 3. Main Dashboard Area
if uploaded_file is None:
    st.title("Organizational Architecture")
    st.info("👈 Please upload your HR data in the sidebar to begin building the map.")
else:
    st.title(f"Organizational Architecture: {selected_sub if selected_sub != 'All' else 'Full Organization'}")
    
    filtered_df = df.copy()
    if selected_sub != "All" and 'Sub Function' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Sub Function'] == selected_sub]
        
    if not include_retainers and 'Employment Type' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
        
    if not include_inactive and 'Status' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Status'].astype(str).str.contains('Inactive', case=False, na=False)]

    valid_ids = [str(vid).replace('.0', '').strip() for vid in filtered_df['Employee Code'].tolist()]
    
    js_rows = []
    for index, row in filtered_df.iterrows():
        emp_id = str(row.get('Employee Code', '')).replace('.0', '').strip()
        manager_id = str(row.get('L1 Manager Code', '')).replace('.0', '').strip()
        
        name = str(row.get('Employee Name', '')).strip()
        grade = str(row.get('Grade', '')).strip()
        designation = str(row.get('Designation', '')).strip()
        onroll = str(row.get('Onroll Reportees', '0')).strip()
        sub_func = str(row.get('Sub Function', '')).strip()
        
        if manager_id not in valid_ids:
            manager_id = ""
            
        # BEAUTIFUL UI: Premium light green cards with pure white background
        box_html = f"""
        <div class='beautiful-card'>
            <div class='card-header'>
                <span class='badge'>{sub_func[:15]}</span>
                <span class='grade'>GR: {grade}</span>
            </div>
            <div class='card-body'>
                <div class='card-name'>{name}</div>
                <div class='card-title'>{designation}</div>
            </div>
            <div class='card-footer'>
                <div class='stat'><span>On-Roll</span><b>{onroll}</b></div>
                <div class='stat'><span>Appr HC</span><b style='color:#ea580c;'>-</b></div>
                <div class='stat'><span>Off-Roll</span><b style='color:#65a30d;'>-</b></div>
            </div>
        </div>
        """
        
        # Stripping newlines to prevent crashes
        box_html_clean = box_html.replace('\n', '').replace('\r', '')

        mgr_str = f"'{manager_id}'" if manager_id else "''"
        row_string = f"[{{'v': '{emp_id}', 'f': \"{box_html_clean}\"}}, {mgr_str}, '']"
        js_rows.append(row_string)

    all_rows_formatted = ",\n".join(js_rows)

    # 5. HTML Template - Clipped Download FIXED
    html_template = f"""
    <html>
      <head>
        <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <script type="text/javascript">
          google.charts.load('current', {{packages:["orgchart"]}});
          google.charts.setOnLoadCallback(drawChart);
          
          function drawChart() {{
            var data = new google.visualization.DataTable();
            data.addColumn('string', 'Name');
            data.addColumn('string', 'Manager');
            data.addColumn('string', 'ToolTip');
            data.addRows([{all_rows_formatted}]);
            var chart = new google.visualization.OrgChart(document.getElementById('chart_div'));
            chart.draw(data, {{allowHtml:true, allowCollapse:true, size:'large', nodeClass:'myNode'}});
          }}
          
          function downloadImage() {{
              // THE FIX: We target the inner 'chart_div' directly so it captures the FULL expanding tree, not the cut-off scrollbox.
              const chartContainer = document.getElementById("chart_div");
              
              html2canvas(chartContainer, {{ 
                  backgroundColor: "#ffffff", // Pure white background
                  scale: 2 // High-res export
              }}).then(canvas => {{
                  let link = document.createElement('a');
                  link.download = 'Beautiful_Org_Chart.png';
                  link.href = canvas.toDataURL("image/png");
                  link.click();
              }});
          }}
       </script>
       <style>
         @keyframes slideUpFade {{
            0% {{ opacity: 0; transform: translateY(20px); }}
            100% {{ opacity: 1; transform: translateY(0); }}
         }}
         
         body {{ margin: 0; padding: 0; background-color: #ffffff; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
         
         /* Connecting Lines Styling - Soft Grey */
         .google-visualization-orgchart-lineleft, .google-visualization-orgchart-lineright, .google-visualization-orgchart-linebottom {{
             border-color: #cbd5e1 !important;
             border-width: 2px !important;
         }}
         
         .myNode {{ border: none !important; background: none !important; padding: 0 !important; box-shadow: none !important; margin: 12px; }}
         
         /* Beautiful Light Green Card Design */
         .beautiful-card {{
            border-radius: 12px;
            background: #ffffff;
            border: 1px solid #dcfce7;
            border-top: 4px solid #22c55e; /* Vibrant Coromandel Green Accent Line */
            position: relative;
            width: 250px;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(34, 197, 94, 0.05), 0 2px 4px -1px rgba(34, 197, 94, 0.03);
            animation: slideUpFade 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards;
            transition: all 0.3s ease;
         }}
         
         .beautiful-card:hover {{
            transform: translateY(-4px);
            box-shadow: 0 12px 20px -5px rgba(34, 197, 94, 0.15), 0 8px 10px -6px rgba(34, 197, 94, 0.1);
            border-color: #86efac;
            z-index: 10;
         }}
         
         .card-header {{
            padding: 12px 16px;
            display: flex; justify-content: space-between; align-items: center;
            background: linear-gradient(180deg, #f0fdf4 0%, #ffffff 100%);
            border-bottom: 1px solid #f1f5f9;
         }}
         
         .badge {{
            background: #dcfce7;
            color: #166534;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
         }}
         
         .grade {{ color: #64748b; font-size: 11px; font-weight: 600; }}
         
         .card-body {{ padding: 16px; text-align: left; }}
         .card-name {{ font-size: 15px; font-weight: 700; color: #0f172a; margin-bottom: 4px; }}
         .card-title {{ font-size: 12px; color: #475569; font-weight: 400; line-height: 1.4; }}
         
         .card-footer {{
            background: #f8fafc;
            border-top: 1px solid #f1f5f9;
            padding: 10px 16px;
            display: flex; justify-content: space-between;
         }}
         .stat {{ display: flex; flex-direction: column; align-items: center; gap: 2px; }}
         .stat span {{ font-size: 9px; color: #64748b; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; }}
         .stat b {{ font-size: 13px; color: #0f172a; }}
         
         /* Sleek Green Download Button */
         .download-btn {{ 
            background: #16a34a;
            color: #ffffff; border: none; padding: 12px 24px; border-radius: 8px; 
            cursor: pointer; font-weight: 600; font-size: 14px; margin-bottom: 24px; 
            transition: all 0.2s; display: inline-flex; align-items: center; gap: 8px;
            box-shadow: 0 4px 6px -1px rgba(22, 163, 74, 0.2);
         }}
         .download-btn:hover {{ background: #15803d; transform: translateY(-1px); }}
         
         #scroll_wrapper {{
            overflow-x: auto; width: 100%; padding: 40px 20px; 
            background-color: #ffffff; border-radius: 12px;
            border: 1px solid #e2e8f0;
         }}
         
         #chart_div {{
             display: inline-block; min-width: 100%; padding: 20px; background-color: #ffffff;
         }}
       </style>
      </head>
      <body>
        <button class="download-btn" onclick="downloadImage()">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            Download High-Res Architecture
        </button>
        <div id="scroll_wrapper">
            <div id="chart_div"></div>
        </div>
      </body>
    </html>
    """
    
    components.html(html_template, height=900, scrolling=True)
