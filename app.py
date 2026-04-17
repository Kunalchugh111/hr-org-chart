import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. Page Configuration
st.set_page_config(page_title="Connected Intelligence | Org Design", layout="wide", initial_sidebar_state="expanded")

# --- UI TWEAKS FOR STREAMLIT ---
st.markdown("""
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 0rem; }
        h1 { font-family: -apple-system, BlinkMacSystemFont, sans-serif; font-weight: 700; color: #1e293b; }
        .stButton>button { border-radius: 8px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

# 2. Sidebar for Controls (SaaS App feel)
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/8/82/Coromandel_International_Logo.svg/512px-Coromandel_International_Logo.svg.png", width=150)
    st.markdown("### Org Chart Generator")
    st.markdown("Upload your roster to generate dynamic reporting structures.")
    st.markdown("---")
    
    uploaded_file = st.file_uploader("Upload HR Data", type=['csv', 'xlsx'])
    
    selected_sub = "All"
    include_retainers = True
    include_inactive = False
    
    if uploaded_file is not None:
        st.markdown("### Filters")
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        df.columns = df.columns.str.strip() # Clean invisible spaces from headers
            
        if 'Sub Function' in df.columns:
            sub_functions = df['Sub Function'].dropna().unique()
            selected_sub = st.selectbox("Sub Function", ["All"] + list(sub_functions))
        
        include_retainers = st.toggle("Include Retainers", value=True)
        include_inactive = st.toggle("Include Inactive", value=False)

# 3. Main Dashboard Area
if uploaded_file is None:
    st.title("Organizational Architecture")
    st.info("👈 Please upload your HR data in the sidebar to begin building the chart.")
else:
    st.title(f"Organizational Architecture: {selected_sub if selected_sub != 'All' else 'Full Organization'}")
    
    # Apply Filters
    filtered_df = df.copy()
    if selected_sub != "All" and 'Sub Function' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Sub Function'] == selected_sub]
        
    if not include_retainers and 'Employment Type' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
        
    if not include_inactive and 'Status' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Status'].astype(str).str.contains('Inactive', case=False, na=False)]

    valid_ids = [str(vid).replace('.0', '').strip() for vid in filtered_df['Employee Code'].tolist()]
    
    # 4. Generate the Chart Nodes with Premium CSS
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
            
        # The new highly stylized HTML Card
        box_html = f"""
        <div class='premium-card'>
            <div class='card-header'>
                <span>{sub_func[:15]}</span>
                <span>GR: {grade}</span>
            </div>
            <div class='card-body'>
                <div class='card-name'>{name}</div>
                <div class='card-title'>{designation}</div>
            </div>
            <div class='card-footer'>
                <span>On-roll: <b>{onroll}</b></span>
                <span>HC: <b>-</b></span>
                <span>Off: <b>-</b></span>
            </div>
        </div>
        """

        mgr_str = f"'{manager_id}'" if manager_id else "''"
        row_string = f"[{{'v': '{emp_id}', 'f': \"{box_html}\"}}, {mgr_str}, '']"
        js_rows.append(row_string)

    all_rows_formatted = ",\n".join(js_rows)

    # 5. HTML Template with spectacular styling
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
              const chartContainer = document.getElementById("chart_div");
              html2canvas(chartContainer, {{ backgroundColor: "#f8fafc", scale: 2 }}).then(canvas => {{
                  let link = document.createElement('a');
                  link.download = 'Spectacular_Org_Chart.png';
                  link.href = canvas.toDataURL("image/png");
                  link.click();
              }});
          }}
       </script>
       <style>
         body {{ margin: 0; padding: 0; background-color: #ffffff; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
         
         /* Hide default Google OrgChart borders */
         .myNode {{ border: none !important; background: none !important; padding: 0 !important; box-shadow: none !important; margin: 10px; }}
         
         /* Premium Card CSS */
         .premium-card {{
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            background: #ffffff;
            width: 250px;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
         }}
         .premium-card:hover {{
            transform: translateY(-3px);
            box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
            border-color: #cbd5e1;
         }}
         .card-header {{
            background-color: #f8fafc;
            border-bottom: 1px solid #e2e8f0;
            padding: 10px 14px;
            display: flex;
            justify-content: space-between;
            font-size: 11px;
            color: #64748b;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
         }}
         .card-body {{
            padding: 18px 14px;
            text-align: left;
         }}
         .card-name {{
            font-size: 15px;
            font-weight: 700;
            color: #0f172a;
            margin-bottom: 4px;
         }}
         .card-title {{
            font-size: 13px;
            color: #475569;
            line-height: 1.4;
         }}
         .card-footer {{
            background-color: #f8fafc;
            border-top: 1px solid #e2e8f0;
            padding: 10px 14px;
            font-size: 11px;
            color: #475569;
            display: flex;
            justify-content: space-between;
         }}
         
         /* Modern Download Button */
         .download-btn {{ 
            background-color: #0f172a; 
            color: #ffffff; 
            border: none; 
            padding: 12px 24px; 
            border-radius: 30px; 
            cursor: pointer; 
            font-weight: 600;
            font-size: 14px;
            margin-bottom: 20px; 
            transition: background-color 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 8px;
         }}
         .download-btn:hover {{ background-color: #334155; }}
         
         /* Scroll Container */
         #scroll_wrapper {{
            overflow-x: auto; 
            width: 100%; 
            padding: 40px 20px; 
            background-color: #f8fafc; 
            border-radius: 16px;
            border: 1px solid #e2e8f0;
         }}
       </style>
      </head>
      <body>
        <button class="download-btn" onclick="downloadImage()">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            Download High-Res Map
        </button>
        <div id="scroll_wrapper">
            <div id="chart_div" style="display: inline-block; min-width: 100%;"></div>
        </div>
      </body>
    </html>
    """
    
    components.html(html_template, height=850, scrolling=True)
