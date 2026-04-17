import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. Page Configuration (Clean, Premium Aesthetic)
st.set_page_config(page_title="HR Org Chart Generator", layout="wide")
st.title("Automated Org Chart Generator")
st.markdown("Upload your employee roster to instantly generate and download dynamic reporting structures.")

# 2. File Uploader
uploaded_file = st.file_uploader("Upload HR Data (CSV or Excel)", type=['csv', 'xlsx'])

if uploaded_file is not None:
    # Read the data
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
    
    st.markdown("### Filter Options")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if 'Sub Function' in df.columns:
            sub_functions = df['Sub Function'].dropna().unique()
            selected_sub = st.selectbox("Select Sub Function:", ["All"] + list(sub_functions))
        else:
            selected_sub = "All"
            st.warning("Column 'Sub Function' not found.")

    with col2:
        # Toggle for Retainers (Assuming a column named 'Employment Type' exists)
        include_retainers = st.toggle("Include Retainers", value=True)
        
    with col3:
        # Toggle for Inactive (Assuming a column named 'Status' exists)
        include_inactive = st.toggle("Include Inactive Employees", value=False)

 # 3. Apply Filters
    filtered_df = df.copy()
    if selected_sub != "All":
        filtered_df = filtered_df[filtered_df['Sub Function'] == selected_sub]
        
    if not include_retainers and 'Employment Type' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Employment Type'].astype(str).str.contains('Retainer', case=False, na=False)]
        
    if not include_inactive and 'Status' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Status'].astype(str).str.contains('Inactive', case=False, na=False)]

    # FIX: Strip out the hidden '.0' from the Pandas decimal conversion
    valid_ids = [str(vid).replace('.0', '').strip() for vid in filtered_df['Employee Code'].tolist()]
    
    # 4. Generate the Chart
    js_rows = []
    for index, row in filtered_df.iterrows():
        # FIX: Clean the IDs here as well so they match perfectly
        emp_id = str(row.get('Employee Code', '')).replace('.0', '').strip()
        manager_id = str(row.get('L1 Manager Code', '')).replace('.0', '').strip()
        
        name = str(row.get('Employee Name', '')).strip()
        grade = str(row.get('Grade', '')).strip()
        designation = str(row.get('Designation', '')).strip()
        onroll = str(row.get('Onroll Reportees', '0')).strip()
        sub_func = str(row.get('Sub Function', '')).strip()
        
        if manager_id not in valid_ids:
            manager_id = ""
            
        box_html = f"""<div class='customBox'><table style='width:100%; border-collapse: collapse;'><tr style='background-color:#c5e1a5;'><td style='border:1px solid #a5d6a7; padding:4px;'>{sub_func[:12]}</td><td style='border:1px solid #a5d6a7; padding:4px;'>Gr: {grade}</td></tr><tr style='background-color:#ffffff;'><td colspan='2' style='border:1px solid #a5d6a7; padding:6px;'><b>{name}</b><br>{designation}</td></tr><tr style='background-color:#f1f8e9; font-size:11px;'><td colspan='2' style='border:1px solid #a5d6a7; padding:4px;'>On-roll: {onroll}</td></tr></table></div>"""

        mgr_str = f"'{manager_id}'" if manager_id else "''"
        row_string = f"[{{'v': '{emp_id}', 'f': \"{box_html}\"}}, {mgr_str}, '']"
        js_rows.append(row_string)

    all_rows_formatted = ",\n".join(js_rows)

    # 5. The HTML Template with the Image Downloader built-in
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
          
     // Function to download the chart as an image (FIXED for large scrolling charts)
          function downloadImage() {
              const chartContainer = document.getElementById("chart_div");
              
              html2canvas(chartContainer, {
                  width: chartContainer.scrollWidth,
                  height: chartContainer.scrollHeight,
                  scrollX: 0,
                  scrollY: 0,
                  backgroundColor: "#ffffff" // Changed to white so the background isn't transparent
              }).then(canvas => {
                  let link = document.createElement('a');
                  link.download = 'Org_Chart_Full.png';
                  link.href = canvas.toDataURL("image/png");
                  link.click();
              });
          }
          }}
       </script>
       <style>
         .myNode {{ border: none !important; background: none !important; padding: 0 !important; box-shadow: none !important; margin: 5px; }}
         .customBox {{ border: 2px solid #5a8231; border-radius: 5px; font-family: Arial; text-align: center; width: 220px; overflow: hidden; box-shadow: 2px 2px 5px rgba(0,0,0,0.2);}}
         .btn {{ background-color: #000; color: #fff; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-family: Arial; margin-bottom: 15px; }}
         .btn:hover {{ background-color: #333; }}
       </style>
      </head>
      <body>
        <button class="btn" onclick="downloadImage()">Download Chart as Image</button>
        <div id="chart_div" style="overflow-x: auto; width: 100%; padding: 20px;"></div>
      </body>
    </html>
    """
    
    st.markdown("---")
    components.html(html_template, height=800, scrolling=True)
