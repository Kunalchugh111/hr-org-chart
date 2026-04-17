import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. Page Configuration (Dark Theme App)
st.set_page_config(page_title="Connected Intelligence | Org Design", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 0rem; }
        h1 { font-family: -apple-system, BlinkMacSystemFont, sans-serif; font-weight: 800; color: #f8fafc; }
        .stApp { background-color: #0f172a; color: white; }
    </style>
""", unsafe_allow_html=True)

# 2. Sidebar Controls
with st.sidebar:
    st.markdown("### 🧬 Architecture Generator")
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
            
        # SPECTACULAR UI: Dark mode glass cards with neon gradient accents
        box_html = f"""
        <div class='glass-card'>
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
                <div class='stat'><span>Appr HC</span><b style='color:#fb923c;'>-</b></div>
                <div class='stat'><span>Off-Roll</span><b style='color:#a3e635;'>-</b></div>
            </div>
        </div>
        """
        
        # THE CRITICAL BUG FIX: Stripping newlines so JavaScript doesn't crash
        box_html_clean = box_html.replace('\n', '').replace('\r', '')

        mgr_str = f"'{manager_id}'" if manager_id else "''"
        row_string = f"[{{'v': '{emp_id}', 'f': \"{box_html_clean}\"}}, {mgr_str}, '']"
        js_rows.append(row_string)

    all_rows_formatted = ",\n".join(js_rows)

    # 5. HTML Template with High-End CSS Animations and Download Fix
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
              const chartContainer = document.getElementById("scroll_wrapper");
              
              // Download Fix: Forced solid background so it doesn't render black
              html2canvas(chartContainer, {{ 
                  backgroundColor: "#0b1120", 
                  scale: 2,
                  width: chartContainer.scrollWidth,
                  height: chartContainer.scrollHeight
              }}).then(canvas => {{
                  let link = document.createElement('a');
                  link.download = 'Spectacular_Org_Chart.png';
                  link.href = canvas.toDataURL("image/png");
                  link.click();
              }});
          }}
       </script>
       <style>
         @keyframes slideUpFade {{
            0% {{ opacity: 0; transform: translateY(30px); }}
            100% {{ opacity: 1; transform: translateY(0); }}
         }}
         
         body {{ margin: 0; padding: 0; background-color: #0b1120; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
         
         /* Connecting Lines Styling */
         .google-visualization-orgchart-lineleft, .google-visualization-orgchart-lineright, .google-visualization-orgchart-linebottom {{
             border-color: #334155 !important;
             border-width: 2px !important;
         }}
         
         .myNode {{ border: none !important; background: none !important; padding: 0 !important; box-shadow: none !important; margin: 12px; }}
         
         /* Spectacular Card Design */
         .glass-card {{
            border-radius: 16px;
            background: linear-gradient(145deg, #1e293b, #0f172a);
            border: 1px solid #334155;
            position: relative;
            width: 260px;
            overflow: hidden;
            box-shadow: 0 10px 15px -3px rgba(0,0,0,0.5);
            animation: slideUpFade 0.6s cubic-bezier(0.16, 1, 0.3, 1) forwards;
            transition: all 0.3s ease;
         }}
         
         /* Glowing Neon Top/Bottom Borders */
         .glass-card::before {{
            content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, #f97316, #fb923c); /* Orange */
         }}
         .glass-card::after {{
            content: ''; position: absolute; bottom: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, #65a30d, #a3e635); /* Green */
         }}
         
         .glass-card:hover {{
            transform: translateY(-5px) scale(1.02);
            box-shadow: 0 20px 25px -5px rgba(0,0,0,0.6), 0 0 15px rgba(249, 115, 22, 0.3);
            border-color: #475569;
            z-index: 10;
         }}
         
         .card-header {{
            padding: 12px 16px;
            display: flex; justify-content: space-between; align-items: center;
            border-bottom: 1px solid rgba(255,255,255,0.05);
         }}
         
         .badge {{
            background: rgba(249, 115, 22, 0.15);
            color: #fdba74;
            padding: 4px 8px;
            border-radius: 20px;
            font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;
         }}
         .grade {{ color: #94a3b8; font-size: 11px; font-weight: 600; }}
         
         .card-body {{ padding: 20px 16px; text-align: center; }}
         .card-name {{ font-size: 16px; font-weight: 700; color: #f8fafc; margin-bottom: 6px; letter-spacing: 0.2px; }}
         .card-title {{ font-size: 12px; color: #cbd5e1; font-weight: 400; line-height: 1.4; }}
         
         .card-footer {{
            background: rgba(0,0,0,0.2);
            padding: 12px 16px;
            display: flex; justify-content: space-between;
         }}
         .stat {{ display: flex; flex-direction: column; align-items: center; gap: 4px; }}
         .stat span {{ font-size: 9px; color: #64748b; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; }}
         .stat b {{ font-size: 14px; color: #f8fafc; }}
         
         /* Sleek Action Button */
         .download-btn {{ 
            background: linear-gradient(135deg, #f97316 0%, #ea580c 100%);
            color: #ffffff; border: none; padding: 12px 28px; border-radius: 30px; 
            cursor: pointer; font-weight: 700; font-size: 14px; margin-bottom: 24px; 
            transition: all 0.3s; display: inline-flex; align-items: center; gap: 8px;
            box-shadow: 0 4px 10px rgba(234, 88, 12, 0.3);
         }}
         .download-btn:hover {{ transform: translateY(-2px); box-shadow: 0 6px 15px rgba(234, 88, 12, 0.5); }}
         
         #scroll_wrapper {{
            overflow-x: auto; width: 100%; padding: 40px 20px; 
            background-color: #0b1120; border-radius: 16px;
            border: 1px solid #1e293b;
         }}
       </style>
      </head>
      <body>
        <button class="download-btn" onclick="downloadImage()">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            Download High-Res Architecture
        </button>
        <div id="scroll_wrapper">
            <div id="chart_div" style="display: inline-block; min-width: 100%;"></div>
        </div>
      </body>
    </html>
    """
    
    components.html(html_template, height=900, scrolling=True)
