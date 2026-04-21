import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="OrgDesign Pro",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
  #MainMenu, footer, header, .stDeployButton,
  [data-testid="stSidebar"], [data-testid="stToolbar"],
  [data-testid="stDecoration"], [data-testid="stHeader"] { display: none !important; }
  .stApp { background: #ffffff !important; }
  .block-container { padding: 0 !important; max-width: 100% !important; margin: 0 !important; }
  iframe { border-radius: 0 !important; border: none !important; }
</style>
""", unsafe_allow_html=True)

APP_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>OrgDesign Pro</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Plus+Jakarta+Sans:ital,wght@0,400;0,500;0,600;0,700;0,800;1,400&display=swap" rel="stylesheet"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<style>
:root {
  --bg:#ffffff; --bg2:#f8fafc; --bg3:#f1f5f9; --bg4:#e8edf5;
  --border:#e2e8f0; --border2:#cbd5e1;
  --text:#0f172a; --text2:#475569; --text3:#94a3b8;
  --accent:#4f46e5; --accent2:#6366f1; --accent-light:#eef2ff; --accent-mid:#c7d2fe;
  --success:#059669; --success-light:#d1fae5;
  --warning:#d97706; --warning-light:#fef3c7;
  --danger:#dc2626;
  --shadow-xs:0 1px 2px rgba(0,0,0,0.05);
  --shadow-sm:0 1px 4px rgba(0,0,0,0.07),0 1px 2px rgba(0,0,0,0.04);
  --shadow-md:0 4px 16px rgba(0,0,0,0.08),0 2px 4px rgba(0,0,0,0.04);
  --shadow-lg:0 12px 40px rgba(0,0,0,0.1),0 4px 12px rgba(0,0,0,0.05);
  --r:10px; --r-lg:14px; --r-xl:18px;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;background:var(--bg);font-family:'Plus Jakarta Sans',sans-serif;color:var(--text);overflow:hidden;font-size:14px}
body{display:flex;flex-direction:column}

/* NAV */
.topnav{flex-shrink:0;height:54px;background:var(--bg);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 24px;gap:12px;z-index:100;box-shadow:var(--shadow-xs)}
.brand{font-family:'Syne',sans-serif;font-weight:800;font-size:1.05rem;color:var(--text);display:flex;align-items:center;gap:9px;letter-spacing:-0.02em;flex-shrink:0}
.brand-icon{width:30px;height:30px;background:linear-gradient(135deg,var(--accent),var(--accent2));border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:0.9rem;box-shadow:0 3px 10px rgba(79,70,229,0.3)}
.nav-sep{width:1px;height:26px;background:var(--border);flex-shrink:0}
.step-trail{display:flex;align-items:center;gap:2px;flex:1;justify-content:center}
.step-item{display:flex;align-items:center;gap:6px;font-size:0.76rem;font-weight:600;color:var(--text3);transition:color 0.2s;white-space:nowrap;padding:4px 6px;border-radius:6px}
.step-item.active{color:var(--accent);background:var(--accent-light)}
.step-item.done{color:var(--success)}
.step-dot{width:22px;height:22px;border-radius:50%;background:var(--bg3);border:2px solid var(--border2);display:flex;align-items:center;justify-content:center;font-size:0.62rem;font-weight:800;color:var(--text3);transition:all 0.2s;flex-shrink:0}
.step-item.active .step-dot{background:var(--accent);border-color:var(--accent);color:#fff}
.step-item.done .step-dot{background:var(--success);border-color:var(--success);color:#fff;font-size:0.7rem}
.step-arrow{color:var(--border2);font-size:0.8rem;margin:0 1px}

/* MAIN / SCREENS */
.main{flex:1;overflow:hidden;position:relative}
.screen{position:absolute;inset:0;overflow-y:auto;display:flex;flex-direction:column;padding:32px 36px;background:var(--bg);opacity:0;pointer-events:none;transform:translateX(18px);transition:opacity 0.22s ease,transform 0.22s ease}
.screen.active{opacity:1;pointer-events:auto;transform:translateX(0)}
#screen-chart{padding:0;overflow:hidden}

/* UPLOAD */
.upload-center{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100%;gap:24px;padding:24px}
.upload-hero{text-align:center}
.upload-hero h1{font-family:'Syne',sans-serif;font-weight:800;font-size:2rem;color:var(--text);letter-spacing:-0.03em;margin-bottom:8px}
.upload-hero p{color:var(--text3);font-size:0.9rem;max-width:400px;line-height:1.6}
.upload-zone{width:520px;max-width:100%;border:2px dashed var(--border2);border-radius:var(--r-xl);padding:48px 32px;text-align:center;cursor:pointer;transition:all 0.2s;background:var(--bg);position:relative}
.upload-zone:hover,.upload-zone.drag-over{border-color:var(--accent);background:var(--accent-light)}
.upload-zone input[type="file"]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.upload-emoji{font-size:2.8rem;margin-bottom:14px;display:block}
.upload-zone h3{font-weight:800;font-size:1.15rem;color:var(--text);margin-bottom:6px}
.upload-zone p{font-size:0.84rem;color:var(--text3);line-height:1.5}
.upload-zone p span{color:var(--accent);font-weight:700}
.info-cards{display:flex;gap:14px;width:520px;max-width:100%}
.info-card{flex:1;background:var(--bg2);border:1px solid var(--border);border-radius:var(--r-lg);padding:16px}
.info-card-title{font-size:0.72rem;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:10px}
.info-card-row{font-size:0.8rem;color:var(--text2);font-weight:500;padding:3px 0;display:flex;align-items:center;gap:6px}
.info-card-row::before{content:'';width:5px;height:5px;background:var(--border2);border-radius:50%;flex-shrink:0}

/* SECTION */
.section-header{margin-bottom:24px}
.section-title{font-family:'Syne',sans-serif;font-weight:700;font-size:1.45rem;color:var(--text);letter-spacing:-0.02em;margin-bottom:4px}
.section-sub{font-size:0.84rem;color:var(--text2)}

/* COL MAP */
.detected-chips{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:24px}
.col-chip{display:inline-flex;align-items:center;gap:7px;padding:5px 11px;background:var(--bg2);border:1.5px solid var(--border);border-radius:999px;font-size:0.76rem;font-weight:600;color:var(--text2)}
.col-chip .chip-sample{color:var(--text3);font-size:0.7rem;font-style:italic}
.map-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;max-width:740px;margin-bottom:28px}
.map-card{background:var(--bg);border:1.5px solid var(--border);border-radius:var(--r-lg);padding:16px;transition:border-color 0.2s,box-shadow 0.2s}
.map-card:focus-within{border-color:var(--accent);box-shadow:0 0 0 3px rgba(79,70,229,0.08)}
.map-card-label{font-size:0.7rem;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:8px;display:flex;align-items:center;gap:7px}
.badge-req{background:#fee2e2;color:var(--danger);padding:1px 7px;border-radius:999px;font-size:0.6rem;font-weight:700}
.badge-opt{background:var(--bg3);color:var(--text3);padding:1px 7px;border-radius:999px;font-size:0.6rem;font-weight:700}
.map-select{width:100%;background:var(--bg3);border:1.5px solid var(--border);border-radius:8px;padding:8px 10px;font-size:0.84rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 10px center;transition:border-color 0.15s}
.map-select:focus{border-color:var(--accent);background-color:var(--bg)}
.map-hint{font-size:0.72rem;color:var(--text3);margin-top:6px}
.data-preview-table{width:100%;max-width:740px;border-collapse:collapse;margin-bottom:28px;font-size:0.78rem}
.data-preview-table th{background:var(--bg3);padding:7px 12px;text-align:left;font-weight:700;color:var(--text2);border:1px solid var(--border);font-size:0.72rem;text-transform:uppercase;letter-spacing:0.04em}
.data-preview-table td{padding:6px 12px;border:1px solid var(--border);color:var(--text2);max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.data-preview-table tr:nth-child(even) td{background:var(--bg2)}

/* CARD DESIGNER */
.card-design-layout{display:grid;grid-template-columns:260px 1fr;gap:24px;flex:1;min-height:0}
.fields-panel{background:var(--bg2);border:1.5px solid var(--border);border-radius:var(--r-lg);padding:18px;overflow-y:auto}
.fields-panel-title{font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:12px}
.fields-section{margin-bottom:16px}
.fields-section-label{font-size:0.66rem;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:8px;padding-bottom:4px;border-bottom:1px solid var(--border)}
.field-chip{display:inline-flex;align-items:center;gap:6px;padding:6px 11px;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;font-size:0.78rem;font-weight:600;color:var(--text);cursor:grab;margin:3px;transition:all 0.15s;user-select:none}
.field-chip:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.field-chip.placed{background:var(--accent-light);border-color:var(--accent-mid);color:var(--accent)}
.field-chip.dragging{opacity:0.4;transform:scale(0.93)}
.drag-icon{color:var(--text3);font-size:0.7rem;cursor:grab}
.card-preview-area{display:flex;flex-direction:column;align-items:flex-start;gap:14px}
.preview-label{font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3)}
.preview-card{width:300px;background:var(--bg);border:2px solid var(--border);border-top:4px solid var(--accent);border-radius:var(--r-lg);box-shadow:var(--shadow-md)}
.preview-card-header{padding:9px 12px;background:var(--bg2);border-bottom:1px solid var(--border);border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;gap:8px}
.preview-card-body{padding:12px 14px}
.preview-card-footer{padding:8px 12px;border-top:1px solid var(--border);border-radius:0 0 12px 12px;background:var(--bg2);display:flex;justify-content:space-between;align-items:center;gap:8px}
.card-zone{flex:1;min-height:30px;min-width:60px;border:2px dashed var(--border2);border-radius:7px;padding:4px 8px;font-size:0.7rem;color:var(--text3);display:flex;align-items:center;justify-content:center;transition:all 0.15s;position:relative;cursor:default}
.card-zone .zone-ph{opacity:0.6;font-style:italic}
.card-zone.drop-target{border-color:var(--accent);background:var(--accent-light)}
.card-zone.filled{border-style:solid;border-color:var(--accent-mid);background:var(--accent-light);flex-direction:column;gap:2px;align-items:flex-start;justify-content:center}
.zone-field{font-weight:700;font-size:0.7rem;color:var(--accent)}
.zone-val{font-size:0.68rem;color:var(--text2);font-style:italic}
.zone-remove{position:absolute;top:3px;right:4px;font-size:0.6rem;cursor:pointer;opacity:0.5;line-height:1}
.zone-remove:hover{opacity:1}
.card-zone-subtitle{width:100%}
.preview-hint{font-size:0.76rem;color:var(--text3);max-width:300px;line-height:1.5}
.preview-name-fixed{display:flex;align-items:center;gap:6px;background:var(--bg3);border:1px dashed var(--border2);border-radius:6px;padding:6px 10px}
.preview-name-fixed span{font-size:0.8rem;color:var(--text2);font-weight:600}
.preview-name-fixed .lock{font-size:0.75rem;opacity:0.5}

/* PHOTO */
.ncard-photo{width:38px;height:38px;border-radius:50%;object-fit:cover;border:2px solid rgba(255,255,255,0.8);display:block;flex-shrink:0}
.ncard-photo-fallback{width:38px;height:38px;border-radius:50%;font-size:0.72rem;font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0}

/* FILTER SETUP */
.filter-setup{max-width:640px}
.filter-chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:24px}
.filter-chip{padding:7px 15px;background:var(--bg3);border:1.5px solid var(--border);border-radius:999px;font-size:0.82rem;font-weight:600;color:var(--text2);cursor:pointer;transition:all 0.15s;user-select:none}
.filter-chip:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.filter-chip.selected{background:var(--accent);border-color:var(--accent);color:#fff}
.filter-counter{font-size:0.72rem;color:var(--text3);font-weight:600;margin-bottom:12px}
.filter-preview-box{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:16px;margin-top:8px}
.fpr-row{display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid var(--border);font-size:0.8rem}
.fpr-row:last-child{border-bottom:none}
.fpr-col{font-weight:700;color:var(--text);min-width:130px}
.fpr-vals{display:flex;flex-wrap:wrap;gap:5px;flex:1}
.fv-pill{padding:2px 9px;background:var(--bg);border:1px solid var(--border);border-radius:999px;font-size:0.72rem;font-weight:500;color:var(--text2)}

/* BUTTONS */
.btn{padding:9px 20px;border-radius:var(--r);font-size:0.84rem;font-weight:700;cursor:pointer;border:none;transition:all 0.15s;display:inline-flex;align-items:center;gap:7px;font-family:'Plus Jakarta Sans',sans-serif;line-height:1;white-space:nowrap}
.btn-primary{background:var(--accent);color:#fff;box-shadow:0 4px 14px rgba(79,70,229,0.3)}
.btn-primary:hover{background:#4338ca;transform:translateY(-1px);box-shadow:0 6px 20px rgba(79,70,229,0.4)}
.btn-ghost{background:transparent;color:var(--text2);border:1.5px solid var(--border)}
.btn-ghost:hover{background:var(--bg3);color:var(--text);border-color:var(--border2)}
.btn-sm{padding:6px 13px;font-size:0.78rem;border-radius:8px}
.btn-row{display:flex;gap:10px;margin-top:28px}
.btn-export-all{background:linear-gradient(135deg,#7c3aed,#0284c7)!important;color:#fff!important;border:none!important;box-shadow:0 4px 14px rgba(124,58,237,0.35)!important}
.btn-export-all:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(124,58,237,0.45)!important}

/* CHART SCREEN */
.chart-toolbar{flex-shrink:0;height:52px;background:var(--bg);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 12px;gap:6px;box-shadow:var(--shadow-xs);position:relative;z-index:20;overflow-x:auto}
.chart-toolbar::-webkit-scrollbar{height:0}
.stats-bar{flex-shrink:0;height:34px;background:var(--bg2);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 18px;gap:18px;font-size:0.73rem}
.stat-item{display:flex;align-items:center;gap:6px;color:var(--text3);font-weight:600}
.stat-item strong{color:var(--text);font-weight:800}
.stat-dot{width:6px;height:6px;border-radius:50%;background:var(--accent)}
.filter-bar{flex-shrink:0;background:var(--bg);border-bottom:1px solid var(--border);padding:7px 18px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;min-height:44px}
.filter-dropdown-wrap{display:flex;align-items:center;gap:6px;font-size:0.79rem}
.filter-dropdown-label{font-weight:700;color:var(--text2)}
.filter-dropdown{background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:5px 28px 5px 10px;font-size:0.79rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center;transition:border-color 0.15s}
.filter-dropdown:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(79,70,229,0.1)}

/* Photo button in toolbar */
.photo-btn{display:flex;align-items:center;gap:5px;padding:5px 10px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0}
.photo-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.photo-btn.loaded{border-color:#059669;color:#059669;background:#d1fae5}
.photo-count{background:#059669;color:#fff;border-radius:999px;padding:1px 6px;font-size:0.65rem;font-weight:800}

/* Depth wrap */
.depth-wrap{display:flex;align-items:center;gap:5px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:3px 6px 3px 9px;flex-shrink:0}
.depth-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);white-space:nowrap}
.depth-select{background:transparent;border:none;border-radius:6px;padding:3px 20px 3px 4px;font-size:0.78rem;font-weight:700;color:var(--accent);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%234f46e5'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 3px center}

/* CHART CANVAS */
.chart-canvas-wrap{flex:1;overflow:auto;background:var(--bg3);cursor:grab;position:relative}
.chart-canvas-wrap:active{cursor:grabbing}
.chart-canvas-wrap::before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(148,163,184,0.12) 1px,transparent 1px),linear-gradient(90deg,rgba(148,163,184,0.12) 1px,transparent 1px);background-size:32px 32px;pointer-events:none;z-index:0}
.chart-canvas-content{display:inline-block;padding:56px 80px 120px 80px;transform-origin:top left;position:relative;z-index:1}
.org-tree{display:inline-block}
.org-tree ul{padding-top:24px;position:relative;list-style:none;display:flex;justify-content:center}
.org-tree li{display:table-cell;vertical-align:top;text-align:center;position:relative;padding:24px 7px 0 7px}
.org-tree li::before,.org-tree li::after{content:'';position:absolute;top:0;right:50%;border-top:2px solid #cbd5e1;width:50%;height:24px}
.org-tree li::after{right:auto;left:50%;border-left:2px solid #cbd5e1}
.org-tree li:only-child::before,.org-tree li:only-child::after{display:none}
.org-tree li:first-child::before,.org-tree li:last-child::after{display:none}
.org-tree li:first-child::after{border-radius:6px 0 0 0}
.org-tree li:last-child::before{border-radius:0 6px 0 0}
.org-tree ul ul::before{content:'';position:absolute;top:0;left:50%;border-left:2px solid #cbd5e1;height:24px}

/* NODE CARD */
.node-card{display:inline-block;width:220px;background:var(--bg);border:1.5px solid var(--border);border-top:3px solid var(--accent);border-radius:var(--r-lg);cursor:pointer;text-align:left;transition:transform 0.15s,box-shadow 0.15s,border-color 0.15s;box-shadow:var(--shadow-sm);position:relative;font-family:'Plus Jakarta Sans',sans-serif}
.node-card:hover{transform:translateY(-3px);box-shadow:0 8px 28px rgba(0,0,0,0.12),0 0 0 2px rgba(79,70,229,0.12);border-color:var(--accent);z-index:10}
.node-card.highlighted{border-color:var(--warning)!important;border-top-color:var(--warning)!important;box-shadow:0 0 0 3px rgba(217,119,6,0.2),0 8px 24px rgba(0,0,0,0.1)!important}
.node-card.collapsed-node{opacity:0.65}
.ncard-header{padding:7px 11px 6px;background:var(--bg2);border-bottom:1px solid var(--border);border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;gap:6px}
.ncard-badge{font-size:0.59rem;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;background:var(--accent-light);color:var(--accent);padding:2px 8px;border-radius:999px;max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;border:1px solid var(--accent-mid)}
.ncard-badge2{font-size:0.65rem;font-weight:700;color:var(--text3);white-space:nowrap}
.ncard-body{padding:10px 12px 8px}
.ncard-body-inner{display:flex;align-items:center;gap:9px}
.ncard-text-wrap{flex:1;min-width:0}
.ncard-name{font-size:0.86rem;font-weight:800;color:var(--text);margin-bottom:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:-0.01em}
.ncard-sub{font-size:0.73rem;color:var(--text2);line-height:1.4;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.ncard-footer{padding:6px 12px;border-top:1px solid var(--border);border-radius:0 0 12px 12px;background:var(--bg2);display:flex;justify-content:space-between;align-items:center;font-size:0.67rem;font-weight:600}
.ncard-fl{color:var(--text3)}
.ncard-fr{background:var(--bg3);padding:2px 8px;border-radius:999px;color:var(--text3);font-size:0.64rem}
.ncard-fr.has-r{background:var(--accent-light);color:var(--accent);border:1px solid var(--accent-mid)}
.collapse-btn{position:absolute;bottom:-11px;left:50%;transform:translateX(-50%);width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:0.58rem;color:var(--text3);transition:all 0.15s;z-index:5;box-shadow:var(--shadow-xs)}
.collapse-btn:hover{background:var(--accent);border-color:var(--accent);color:#fff}
li.collapsed>ul{display:none}

/* SEARCH */
.search-wrap{position:relative;flex:1;max-width:240px}
.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:0.8rem;pointer-events:none;opacity:0.45}
#chart-search{width:100%;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:6px 10px 6px 29px;font-size:0.8rem;font-weight:500;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;transition:border-color 0.15s}
#chart-search:focus{border-color:var(--accent);background:var(--bg)}
#chart-search::placeholder{color:var(--text3)}
#chart-search-results{position:absolute;top:calc(100% + 5px);left:0;right:0;background:var(--bg);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--shadow-lg);max-height:260px;overflow-y:auto;z-index:999;display:none}
#chart-search-results.visible{display:block}
.sr-item{padding:9px 13px;cursor:pointer;border-bottom:1px solid var(--border);transition:background 0.1s}
.sr-item:last-child{border-bottom:none}
.sr-item:hover{background:var(--bg3)}
.sr-name{font-weight:700;font-size:0.82rem}
.sr-sub{font-size:0.72rem;color:var(--text3);margin-top:2px}

/* ZOOM */
.zoom-strip{display:flex;align-items:center;gap:1px;background:var(--bg2);border-radius:8px;padding:2px;border:1.5px solid var(--border)}
.btn-zoom{background:transparent;border:none;border-radius:6px;width:26px;height:26px;cursor:pointer;font-size:0.85rem;font-weight:700;color:var(--text2);font-family:'Plus Jakarta Sans',sans-serif;display:flex;align-items:center;justify-content:center;transition:background 0.12s}
.btn-zoom:hover{background:var(--bg3);color:var(--text)}
.zoom-label{font-size:0.72rem;font-weight:800;color:var(--text);min-width:42px;text-align:center;font-variant-numeric:tabular-nums}

/* EXPORT OVERLAY */
.export-overlay{position:fixed;inset:0;z-index:9999;background:rgba(255,255,255,0.88);backdrop-filter:blur(6px);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px}
.export-spinner{width:44px;height:44px;border:3px solid var(--border2);border-top-color:var(--accent);border-radius:50%;animation:spin 0.7s linear infinite}
.export-steps{display:flex;flex-direction:column;align-items:center;gap:5px;margin-top:6px}
.export-step{font-size:0.76rem;color:var(--text3);font-weight:500;display:flex;align-items:center;gap:8px;min-width:260px}
.export-step.active{color:var(--accent);font-weight:700}
.export-step.done{color:var(--success)}
@keyframes spin{to{transform:rotate(360deg)}}
.no-data{padding:40px;color:var(--text3);font-size:0.9rem;font-weight:600;background:var(--bg);border:1.5px solid var(--border);border-radius:var(--r-lg);max-width:440px}

/* VACANT */
.node-card.vacant{border-top-color:#dc2626!important;background:#fff5f5!important}
.node-card.vacant .ncard-header{background:#fee2e2!important}
.node-card.vacant .ncard-name{color:#dc2626!important}
.node-card.vacant .ncard-footer{background:#fee2e2!important}
.vacant-badge{display:inline-flex;align-items:center;gap:4px;font-size:0.59rem;font-weight:800;text-transform:uppercase;letter-spacing:0.05em;background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:999px;border:1px solid #fca5a5}

/* EDIT BTN */
.ncard-edit-btn{position:absolute;top:6px;right:6px;width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:6px;font-size:0.65rem;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 0.15s,background 0.15s,border-color 0.15s;z-index:8}
.node-card:hover .ncard-edit-btn{opacity:1}
.ncard-edit-btn:hover{background:var(--accent);border-color:var(--accent);color:#fff}

/* COLOR PALETTE */
.color-palette{display:flex;flex-wrap:wrap;gap:7px;margin-top:6px}
.color-swatch{width:24px;height:24px;border-radius:6px;cursor:pointer;border:2.5px solid transparent;transition:transform 0.1s,border-color 0.1s;flex-shrink:0}
.color-swatch:hover{transform:scale(1.15)}
.color-swatch.selected{border-color:var(--text);box-shadow:0 0 0 2px #fff inset}

/* VACANT SELECTOR */
.vacant-setup{margin-top:14px}
.vacant-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:6px}
.vacant-select{flex:1;min-width:100px;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:5px 8px;font-size:0.78rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;appearance:none;cursor:pointer;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 7px center}
.vacant-select:focus{border-color:var(--accent)}

/* MODAL */
.modal-overlay{position:fixed;inset:0;z-index:8000;background:rgba(15,23,42,0.45);backdrop-filter:blur(4px);display:flex;align-items:center;justify-content:center;padding:20px}
.modal-overlay.hidden{display:none}
.modal-box{background:var(--bg);border:1px solid var(--border);border-radius:var(--r-xl);box-shadow:0 24px 80px rgba(0,0,0,0.18);width:440px;max-width:100%;display:flex;flex-direction:column;max-height:80vh;overflow:hidden}
.modal-header{padding:18px 20px 14px;border-bottom:1px solid var(--border);display:flex;align-items:flex-start;justify-content:space-between}
.modal-title{font-family:'Syne',sans-serif;font-weight:700;font-size:1rem;color:var(--text)}
.modal-sub{font-size:0.76rem;color:var(--text3);margin-top:3px}
.modal-close{background:none;border:none;font-size:1.1rem;cursor:pointer;color:var(--text3);line-height:1;padding:2px 6px;border-radius:6px}
.modal-close:hover{background:var(--bg3);color:var(--text)}
.modal-body{padding:16px 20px;flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:12px}
.modal-search{width:100%;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:8px 12px;font-size:0.84rem;font-weight:500;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none}
.modal-search:focus{border-color:var(--accent);background:var(--bg)}
.modal-list{display:flex;flex-direction:column;gap:1px;max-height:280px;overflow-y:auto}
.modal-emp-row{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:8px;cursor:pointer;transition:background 0.1s;border:2px solid transparent}
.modal-emp-row:hover{background:var(--bg2)}
.modal-emp-row.selected{background:var(--accent-light);border-color:var(--accent-mid)}
.modal-emp-avatar{width:32px;height:32px;border-radius:10px;background:var(--accent-light);color:var(--accent);font-size:0.75rem;font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.modal-emp-avatar.vacant-av{background:#fee2e2;color:#dc2626}
.modal-emp-name{font-weight:700;font-size:0.82rem;color:var(--text)}
.modal-emp-sub{font-size:0.71rem;color:var(--text3)}
.modal-emp-row.make-root .modal-emp-name{color:var(--warning)}
.modal-footer{padding:14px 20px;border-top:1px solid var(--border);display:flex;gap:10px;justify-content:flex-end}
.modal-note{font-size:0.73rem;color:var(--text3);flex:1;display:flex;align-items:center}
.tb-sep{width:1px;height:22px;background:var(--border);flex-shrink:0}

/* Photo folder drop zone in chart toolbar */
.photo-folder-input{display:none}
</style>
</head>
<body>

<nav class="topnav">
  <div class="brand"><div class="brand-icon">🏢</div>OrgDesign Pro</div>
  <div class="nav-sep"></div>
  <div class="step-trail">
    <div class="step-item active" id="nav-step-upload"><div class="step-dot">1</div><span>Upload</span></div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-map"><div class="step-dot">2</div><span>Map Columns</span></div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-card"><div class="step-dot">3</div><span>Design Card</span></div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-filter"><div class="step-dot">4</div><span>Set Filters</span></div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-chart"><div class="step-dot">5</div><span>Org Chart</span></div>
  </div>
</nav>

<main class="main">

  <!-- Screen 1: Upload -->
  <div class="screen active" id="screen-upload">
    <div class="upload-center">
      <div class="upload-hero">
        <h1>Build your Org Chart</h1>
        <p>Upload your HR roster and we'll guide you through designing a beautiful, interactive org chart in minutes.</p>
      </div>
      <div class="upload-zone" id="upload-dropzone">
        <input type="file" id="file-input" accept=".csv,.xlsx,.xls"/>
        <span class="upload-emoji">📊</span>
        <h3>Drop your file here</h3>
        <p>Supports CSV and Excel (.xlsx, .xls)<br>or <span>click to browse</span></p>
      </div>
      <div class="info-cards">
        <div class="info-card">
          <div class="info-card-title">✅ Required Columns</div>
          <div class="info-card-row">Employee Code / ID</div>
          <div class="info-card-row">Employee Name</div>
          <div class="info-card-row">Manager Code / ID</div>
        </div>
        <div class="info-card">
          <div class="info-card-title">📸 Photo Tip</div>
          <div class="info-card-row">Name photos by Employee ID</div>
          <div class="info-card-row">e.g. EMP001.jpg, E001.png</div>
          <div class="info-card-row">Load folder from Chart toolbar</div>
        </div>
      </div>
    </div>
  </div>

  <!-- Screen 2: Column Map -->
  <div class="screen" id="screen-map">
    <div class="section-header">
      <div class="section-title">Map Your Columns</div>
      <div class="section-sub">We detected <span id="col-count">0</span> columns. Auto-mapped where possible.</div>
    </div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Detected Columns</div>
    <div class="detected-chips" id="detected-columns"></div>
    <div class="map-grid">
      <div class="map-card">
        <div class="map-card-label">👤 Employee ID <span class="badge-req">Required</span></div>
        <select class="map-select" id="map-empId"></select>
        <div class="map-hint">Unique identifier — also used to match photos</div>
      </div>
      <div class="map-card">
        <div class="map-card-label">🏷️ Employee Name <span class="badge-req">Required</span></div>
        <select class="map-select" id="map-empName"></select>
        <div class="map-hint">Full name shown on the card</div>
      </div>
      <div class="map-card">
        <div class="map-card-label">🔗 Manager ID <span class="badge-opt">Optional</span></div>
        <select class="map-select" id="map-managerId"></select>
        <div class="map-hint">Links employee to their manager</div>
      </div>
    </div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Data Preview (first 3 rows)</div>
    <div id="data-preview-wrap" style="margin-bottom:24px;overflow-x:auto"></div>
    <div class="btn-row" style="margin-top:0">
      <button class="btn btn-ghost" onclick="goTo('upload')">← Back</button>
      <button class="btn btn-primary" onclick="confirmColumnMap()">Continue to Card Design →</button>
    </div>
  </div>

  <!-- Screen 3: Card Designer -->
  <div class="screen" id="screen-card">
    <div class="section-header" style="margin-bottom:18px">
      <div class="section-title">Design Your Card</div>
      <div class="section-sub">Drag fields into card zones and pick an accent color. Load employee photos from the Chart toolbar (Step 5).</div>
    </div>
    <div class="card-design-layout">
      <div class="fields-panel">
        <div class="fields-panel-title">Available Fields</div>
        <div id="card-fields-panel"></div>
        <div class="fields-section" style="margin-top:18px">
          <div class="fields-section-label" style="margin-bottom:8px">🎨 Card Accent Color</div>
          <div class="color-palette" id="color-palette"></div>
        </div>
        <div class="fields-section vacant-setup">
          <div class="fields-section-label" style="margin-bottom:8px">🔴 Mark Vacant Positions</div>
          <div style="font-size:0.74rem;color:var(--text3);margin-bottom:7px;line-height:1.5">Select a column and value that identifies vacant roles.</div>
          <div class="vacant-row">
            <select class="vacant-select" id="vacant-col" onchange="onVacantColChange()"><option value="">Column…</option></select>
            <select class="vacant-select" id="vacant-val" style="display:none"><option value="">Value…</option></select>
          </div>
        </div>
      </div>
      <div class="card-preview-area">
        <div class="preview-label">Live Card Preview</div>
        <div id="card-preview"></div>
        <div class="preview-hint">Drag a field chip onto any zone. Drop another on it to swap. Click ✕ to clear.</div>
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" onclick="goTo('map')">← Back</button>
      <button class="btn btn-primary" onclick="confirmCardDesign()">Continue to Filters →</button>
    </div>
  </div>

  <!-- Screen 4: Filter Setup -->
  <div class="screen" id="screen-filter">
    <div class="section-header">
      <div class="section-title">Set Up Filters</div>
      <div class="section-sub">Choose up to 3 columns to use as filters. The last filter drives "Export All" — it will export one chart per value of that column.</div>
    </div>
    <div class="filter-setup">
      <div class="filter-counter" id="filter-counter">0 of 3 filters selected</div>
      <div class="filter-chips" id="filter-chip-picker"></div>
      <div id="filter-preview-area"></div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" onclick="goTo('card')">← Back</button>
      <button class="btn btn-primary" onclick="launchChart()">🚀 Launch Org Chart</button>
    </div>
  </div>

  <!-- Screen 5: Chart -->
  <div class="screen" id="screen-chart">
    <div class="chart-toolbar">
      <button class="btn btn-ghost btn-sm" onclick="goTo('filter')">← Setup</button>
      <div class="tb-sep"></div>
      <div class="search-wrap">
        <span class="search-icon">🔍</span>
        <input id="chart-search" type="text" placeholder="Search name or ID…" autocomplete="off"/>
        <div id="chart-search-results"></div>
      </div>
      <div class="tb-sep"></div>
      <div class="zoom-strip">
        <button class="btn-zoom" onclick="zoomBy(-0.1)">−</button>
        <span class="zoom-label" id="zoom-level">100%</span>
        <button class="btn-zoom" onclick="zoomBy(0.1)">+</button>
        <button class="btn-zoom" onclick="fitToScreen(true)" title="Fit">⊡</button>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="centerView()">🧭</button>
      <button class="btn btn-ghost btn-sm" onclick="expandAll()">⊞</button>
      <button class="btn btn-ghost btn-sm" onclick="collapseAll()">⊟</button>
      <div class="tb-sep"></div>

      <!-- Bottom-up level depth -->
      <div class="depth-wrap" title="Show N levels counting from the bottom (leaves up)">
        <span class="depth-label">Levels ↑</span>
        <select class="depth-select" id="depth-select" onchange="setMaxDepth(parseInt(this.value))">
          <option value="0">All</option>
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
        </select>
      </div>
      <div class="tb-sep"></div>

      <!-- Photo folder picker -->
      <input type="file" id="photo-folder-input" class="photo-folder-input" accept="image/*" multiple/>
      <div class="photo-btn" id="photo-btn" onclick="openPhotoFolder()" title="Select a folder of employee photos (named by Employee ID)">
        📸 <span id="photo-btn-label">Load Photos</span>
        <span class="photo-count" id="photo-count" style="display:none">0</span>
      </div>
      <div class="tb-sep"></div>

      <div style="flex:1"></div>
      <button class="btn btn-ghost btn-sm" onclick="downloadCSV()">💾 CSV</button>
      <button class="btn btn-ghost btn-sm" onclick="exportPNG()">🖼️ PNG</button>
      <button class="btn btn-ghost btn-sm" onclick="exportPPTX()">📊 PPTX</button>
      <button class="btn btn-sm btn-export-all" onclick="exportAll()">📦 Export All</button>
    </div>

    <div class="stats-bar">
      <div class="stat-item"><div class="stat-dot"></div><strong id="stat-total">—</strong>&nbsp;employees</div>
      <div class="stat-item"><strong id="stat-roots">—</strong>&nbsp;root nodes</div>
      <div class="stat-item"><strong id="stat-vis">—</strong>&nbsp;visible</div>
      <div class="stat-item" id="stat-photos" style="display:none;color:var(--success)">📸 <strong id="stat-photos-val">0</strong> photos</div>
      <div class="stat-item" id="stat-depth-info" style="display:none;color:var(--accent)">↑ <strong id="stat-depth-val">—</strong> levels from bottom</div>
      <div class="stat-item" id="stat-filtered" style="display:none;color:var(--warning)">⚠️ Filtered</div>
    </div>

    <div class="filter-bar" id="filter-bar" style="display:none"></div>

    <div class="chart-canvas-wrap" id="chart-canvas-wrap">
      <div class="chart-canvas-content" id="chart-canvas-content">
        <div class="org-tree" id="org-tree"></div>
      </div>
    </div>
  </div>
</main>

<!-- REASSIGN MODAL -->
<div class="modal-overlay hidden" id="reassign-modal">
  <div class="modal-box">
    <div class="modal-header">
      <div>
        <div class="modal-title">Reassign Manager</div>
        <div class="modal-sub" id="reassign-subject">Moving <strong>—</strong></div>
      </div>
      <button class="modal-close" onclick="closeReassignModal()">✕</button>
    </div>
    <div class="modal-body">
      <input class="modal-search" id="reassign-search" type="text" placeholder="Search employee name or ID…" autocomplete="off" oninput="filterReassignList()"/>
      <div class="modal-list" id="reassign-list"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-sm" onclick="removeCurrentNode()" style="background:#fee2e2;border:1.5px solid #fca5a5;color:#dc2626;margin-right:auto">🗑 Remove</button>
      <span class="modal-note" id="reassign-note">Select a new manager above</span>
      <button class="btn btn-ghost btn-sm" onclick="closeReassignModal()">Cancel</button>
      <button class="btn btn-primary btn-sm" id="reassign-confirm-btn" onclick="confirmReassign()" disabled>Reassign</button>
    </div>
  </div>
</div>

<script>
// ════════════════════════════════════════════════
// STATE
// ════════════════════════════════════════════════
const S = {
  rawRows:[],columns:[],colSamples:{},
  colMap:{empId:'',empName:'',managerId:''},
  cardSlots:{badge1:'',badge2:'',subtitle:'',footer1:'',footer2:''},
  cardAccent:'#4f46e5',
  vacantCol:'',vacantVal:'',
  filterCols:[],activeFilters:{},
  managerOverrides:{},removedIds:new Set(),
  viewData:[],childMap:{},descCount:{},
  nodeHeight:{},       // max height from node to deepest leaf (0 = leaf)
  zoom:1,highlighted:null,
  draggingField:null,
  reassignTarget:null,reassignPick:null,
  maxDepth:0,          // 0=all; N = show N levels counting from the BOTTOM (leaves=1)
  photoMap:{},         // empId (lowercase) → object URL
  photoObjUrls:[],     // track for revocation
};

// ════════════════════════════════════════════════
// NAVIGATION
// ════════════════════════════════════════════════
function goTo(step){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
  document.getElementById('screen-'+step)?.classList.add('active');
  const order=['upload','map','card','filter','chart'];
  const cur=order.indexOf(step);
  order.forEach((s,i)=>{
    const el=document.getElementById('nav-step-'+s); if(!el) return;
    el.className='step-item'+(i<cur?' done':i===cur?' active':'');
    const dot=el.querySelector('.step-dot');
    if(dot) dot.textContent=i<cur?'✓':String(i+1);
  });
  if(step==='chart'){setTimeout(()=>initPan(),80);setTimeout(()=>initSearch(),80);}
}

// ════════════════════════════════════════════════
// FILE HANDLING
// ════════════════════════════════════════════════
function handleFile(file){
  const ext=file.name.split('.').pop().toLowerCase();
  if(ext==='csv'){
    Papa.parse(file,{header:true,skipEmptyLines:true,
      complete:r=>initData(r.data),error:e=>alert('CSV error: '+e.message)});
  } else if(['xlsx','xls'].includes(ext)){
    const reader=new FileReader();
    reader.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      initData(XLSX.utils.sheet_to_json(ws,{defval:''}));
    };
    reader.readAsArrayBuffer(file);
  } else { alert('Please upload a CSV or Excel (.xlsx/.xls) file.'); }
}

function initData(rows){
  S.rawRows=rows;
  S.columns=rows.length?Object.keys(rows[0]):[];
  S.colSamples={};
  S.columns.forEach(col=>{
    S.colSamples[col]=[...new Set(rows.slice(0,25).map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='undefined'&&v!=='null'))].slice(0,3);
  });
  S.colMap=autoDetect(S.columns);
  buildMapScreen();
  goTo('map');
}

function autoDetect(cols){
  const lc=cols.map(c=>c.toLowerCase());
  const find=pats=>{for(const p of pats){const i=lc.findIndex(c=>c.includes(p));if(i>=0)return cols[i];}return '';};
  return {
    empId:   find(['employee code','emp code','emp id','employee id','empcode','empid','staff id','staff code','person id','worker id']),
    empName: find(['employee name','emp name','full name','person name','staff name','worker name','name']),
    managerId:find(['l1 manager code','l1 manager','manager code','manager id','reports to','supervisor','mgr code','mgrid']),
  };
}

// ════════════════════════════════════════════════
// PHOTO FOLDER LOADING
// ════════════════════════════════════════════════
// Primary: modern File System Access API (showDirectoryPicker)
// Fallback: <input multiple accept="image/*"> (the user picks the folder contents)
async function openPhotoFolder(){
  if('showDirectoryPicker' in window){
    try{
      const dirHandle=await window.showDirectoryPicker({mode:'read'});
      await loadFromDirectoryHandle(dirHandle);
    }catch(e){
      if(e.name!=='AbortError') document.getElementById('photo-folder-input').click();
    }
  } else {
    document.getElementById('photo-folder-input').click();
  }
}

async function loadFromDirectoryHandle(dirHandle){
  // Revoke old URLs
  S.photoObjUrls.forEach(u=>URL.revokeObjectURL(u));
  S.photoObjUrls=[];
  const newMap={};
  const IMG_EXTS=new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);
  for await(const[name,handle] of dirHandle.entries()){
    if(handle.kind==='file'){
      const ext=name.split('.').pop().toLowerCase();
      if(IMG_EXTS.has(ext)){
        const file=await handle.getFile();
        const key=name.replace(/\.[^.]+$/,'').toLowerCase().trim();
        const url=URL.createObjectURL(file);
        newMap[key]=url;
        S.photoObjUrls.push(url);
      }
    }
  }
  S.photoMap=newMap;
  updatePhotoUI();
  if(S.viewData.length) renderChart();
}

function loadFromFileInput(files){
  // Fallback: user picked files (may be from a folder via "Upload folder" dialog)
  S.photoObjUrls.forEach(u=>URL.revokeObjectURL(u));
  S.photoObjUrls=[];
  const newMap={};
  const IMG_EXTS=new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);
  Array.from(files).forEach(file=>{
    const name=file.name;
    const ext=name.split('.').pop().toLowerCase();
    if(IMG_EXTS.has(ext)){
      const key=name.replace(/\.[^.]+$/,'').toLowerCase().trim();
      const url=URL.createObjectURL(file);
      newMap[key]=url;
      S.photoObjUrls.push(url);
    }
  });
  S.photoMap=newMap;
  updatePhotoUI();
  if(S.viewData.length) renderChart();
}

function updatePhotoUI(){
  const count=Object.keys(S.photoMap).length;
  const btn=document.getElementById('photo-btn');
  const label=document.getElementById('photo-btn-label');
  const badge=document.getElementById('photo-count');
  const stat=document.getElementById('stat-photos');
  const statVal=document.getElementById('stat-photos-val');
  if(count>0){
    btn.classList.add('loaded');
    label.textContent='Photos';
    badge.textContent=count;
    badge.style.display='';
    if(stat){stat.style.display='flex';statVal.textContent=count;}
  } else {
    btn.classList.remove('loaded');
    label.textContent='Load Photos';
    badge.style.display='none';
    if(stat) stat.style.display='none';
  }
}

function getPhotoUrl(node){
  if(!Object.keys(S.photoMap).length) return '';
  const empId=node.id.toLowerCase().trim();
  if(S.photoMap[empId]) return S.photoMap[empId];
  // Fallback: name-based matching (spaces→underscore, lowercased)
  const nameKey=node.name.toLowerCase().trim().replace(/\s+/g,'_');
  if(S.photoMap[nameKey]) return S.photoMap[nameKey];
  const nameKey2=node.name.toLowerCase().trim().replace(/\s+/g,'');
  if(S.photoMap[nameKey2]) return S.photoMap[nameKey2];
  return '';
}

// ════════════════════════════════════════════════
// SCREEN 2
// ════════════════════════════════════════════════
function buildMapScreen(){
  document.getElementById('col-count').textContent=S.columns.length;
  const chips=document.getElementById('detected-columns');
  chips.innerHTML=S.columns.map(c=>{
    const s=S.colSamples[c].join(', ');
    return`<div class="col-chip">📋 ${esc(c)}${s?`<span class="chip-sample">${esc(s)}</span>`:''}</div>`;
  }).join('');
  const blank='<option value="">— select —</option>';
  const opts=blank+S.columns.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
  ['empId','empName','managerId'].forEach(k=>{
    const sel=document.getElementById('map-'+k);
    if(!sel)return;sel.innerHTML=opts;sel.value=S.colMap[k]||'';
  });
  const wrap=document.getElementById('data-preview-wrap');
  const preview=S.rawRows.slice(0,3);
  if(!preview.length){wrap.innerHTML='';return;}
  let html='<table class="data-preview-table"><thead><tr>'+S.columns.map(c=>`<th>${esc(c)}</th>`).join('')+'</tr></thead><tbody>';
  preview.forEach(row=>{html+='<tr>'+S.columns.map(c=>`<td title="${esc(String(row[c]||''))}">${esc(String(row[c]||'').substring(0,22))}</td>`).join('')+'</tr>';});
  html+='</tbody></table>';wrap.innerHTML=html;
}

function confirmColumnMap(){
  S.colMap.empId=document.getElementById('map-empId').value;
  S.colMap.empName=document.getElementById('map-empName').value;
  S.colMap.managerId=document.getElementById('map-managerId').value;
  if(!S.colMap.empId||!S.colMap.empName){alert('Please map Employee ID and Employee Name.');return;}
  buildCardScreen();goTo('card');
}

// ════════════════════════════════════════════════
// SCREEN 3 — CARD DESIGNER
// ════════════════════════════════════════════════
const AUTO_FIELDS=[
  {id:'__auto_reports__',icon:'📊',label:'Direct Reports',desc:'Count of direct reports'},
  {id:'__auto_teamsize__',icon:'👥',label:'Total Team Size',desc:'All descendants count'},
];

function buildCardScreen(){
  const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId].filter(Boolean));
  const available=S.columns.filter(c=>!core.has(c));
  const panel=document.getElementById('card-fields-panel');
  panel.innerHTML=
    '<div class="fields-section"><div class="fields-section-label">Column Fields</div>'+
    (available.length?available.map(f=>`<div class="field-chip" draggable="true" data-field="${esc(f)}" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)"><span class="drag-icon">⠿</span>${esc(f)}</div>`).join('')
      :'<div style="font-size:0.78rem;color:var(--text3);font-style:italic">No extra columns</div>')+
    '</div><div class="fields-section"><div class="fields-section-label">Auto-Calculated</div>'+
    AUTO_FIELDS.map(f=>`<div class="field-chip" draggable="true" data-field="${f.id}" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)" title="${f.desc}"><span class="drag-icon">⠿</span><span>${f.icon}</span>${f.label}</div>`).join('')+
    '</div>';

  if(!S.cardSlots.footer2) S.cardSlots.footer2='__auto_reports__';

  const COLORS=['#4f46e5','#7c3aed','#db2777','#dc2626','#d97706','#059669','#0891b2','#0284c7','#374151','#0f172a'];
  document.getElementById('color-palette').innerHTML=COLORS.map(c=>
    `<div class="color-swatch${S.cardAccent===c?' selected':''}" style="background:${c}" title="${c}" onclick="setCardAccent('${c}')"></div>`
  ).join('');

  const core2=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId].filter(Boolean));
  const colSel=document.getElementById('vacant-col');
  if(colSel){
    colSel.innerHTML='<option value="">Column…</option>'+S.columns.filter(c=>!core2.has(c)).map(c=>`<option value="${esc(c)}"${S.vacantCol===c?' selected':''}>${esc(c)}</option>`).join('');
    if(S.vacantCol) populateVacantValues(S.vacantCol);
  }
  renderCardPreview();syncChipStates();
}

function onDragStart(e){S.draggingField=e.currentTarget.dataset.field;e.currentTarget.classList.add('dragging');e.dataTransfer.effectAllowed='move';}
function onDragEnd(e){e.currentTarget.classList.remove('dragging');}
function onZoneDragOver(e){e.preventDefault();e.currentTarget.classList.add('drop-target');}
function onZoneDragLeave(e){e.currentTarget.classList.remove('drop-target');}
function onZoneDrop(e,zone){
  e.preventDefault();e.currentTarget.classList.remove('drop-target');
  if(!S.draggingField)return;
  Object.keys(S.cardSlots).forEach(z=>{if(S.cardSlots[z]===S.draggingField)S.cardSlots[z]='';});
  S.cardSlots[zone]=S.draggingField;S.draggingField=null;
  renderCardPreview();syncChipStates();
}
function clearZone(zone){S.cardSlots[zone]='';renderCardPreview();syncChipStates();}
function syncChipStates(){
  const placed=new Set(Object.values(S.cardSlots).filter(Boolean));
  document.querySelectorAll('.field-chip').forEach(c=>c.classList.toggle('placed',placed.has(c.dataset.field)));
}
function fieldLabel(id){if(!id)return '';const af=AUTO_FIELDS.find(f=>f.id===id);if(af)return af.icon+' '+af.label;return id;}
function fieldSampleVal(id){
  if(!id)return '';if(id==='__auto_reports__')return '12';if(id==='__auto_teamsize__')return '48';
  const row=S.rawRows.find(r=>r[id])||S.rawRows[0]||{};return String(row[id]||'Sample').substring(0,22);
}
function zoneHtml(zoneId,placeholder,extraClass=''){
  const v=S.cardSlots[zoneId];
  const dA=`ondragover="onZoneDragOver(event)" ondragleave="onZoneDragLeave(event)" ondrop="onZoneDrop(event,'${zoneId}')"`;
  if(v)return`<div class="card-zone filled ${extraClass}" ${dA}><span class="zone-field">${esc(fieldLabel(v))}</span><span class="zone-val">${esc(fieldSampleVal(v))}</span><span class="zone-remove" onclick="clearZone('${zoneId}')">✕</span></div>`;
  return`<div class="card-zone ${extraClass}" ${dA}><span class="zone-ph">${placeholder}</span></div>`;
}

function renderCardPreview(){
  const sampleRow=S.rawRows.find(r=>r[S.colMap.empName])||S.rawRows[0]||{};
  const sampleName=String(sampleRow[S.colMap.empName]||'Employee Name').substring(0,26);
  const ac=S.cardAccent;
  document.getElementById('card-preview').innerHTML=`
    <div class="preview-card" style="border-top-color:${ac}">
      <div class="preview-card-header">${zoneHtml('badge1','+ Badge Left')}${zoneHtml('badge2','+ Badge Right')}</div>
      <div class="preview-card-body">
        <div style="display:flex;align-items:center;gap:9px;margin-bottom:7px">
          <div style="width:38px;height:38px;border-radius:50%;background:${ac}18;color:${ac};font-size:0.72rem;font-weight:800;display:flex;align-items:center;justify-content:center;border:2px solid ${ac}44;flex-shrink:0">AB</div>
          <div class="preview-name-fixed" style="flex:1;margin:0"><span class="lock">🔒</span><span>${esc(sampleName)}</span></div>
        </div>
        ${zoneHtml('subtitle','+ Subtitle / Designation','card-zone-subtitle')}
      </div>
      <div class="preview-card-footer">${zoneHtml('footer1','+ Footer Left')}${zoneHtml('footer2','+ Footer Right')}</div>
    </div>
    <div style="margin-top:10px;font-size:0.72rem;color:var(--text3)">Accent: <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:${ac};vertical-align:middle;margin-left:4px"></span> <strong style="color:${ac}">${ac}</strong></div>`;
}

function setCardAccent(color){S.cardAccent=color;document.querySelectorAll('.color-swatch').forEach(s=>s.classList.toggle('selected',s.style.background===color));renderCardPreview();}
function onVacantColChange(){
  const col=document.getElementById('vacant-col').value;S.vacantCol=col;S.vacantVal='';
  const v=document.getElementById('vacant-val');
  if(col){v.style.display='';populateVacantValues(col);}else{v.style.display='none';}
}
function populateVacantValues(col){
  const vals=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();
  const v=document.getElementById('vacant-val');if(!v)return;
  v.innerHTML='<option value="">Value…</option>'+vals.map(x=>`<option value="${esc(x)}"${S.vacantVal===x?' selected':''}>${esc(x)}</option>`).join('');
  v.style.display='';v.onchange=()=>{S.vacantVal=v.value;};
}
function isVacant(node){return S.vacantCol&&S.vacantVal&&node[S.vacantCol]===S.vacantVal;}
function confirmCardDesign(){const v=document.getElementById('vacant-val');if(v)S.vacantVal=v.value;buildFilterScreen();goTo('filter');}

// ════════════════════════════════════════════════
// SCREEN 4 — FILTER SETUP
// ════════════════════════════════════════════════
function buildFilterScreen(){
  const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId].filter(Boolean));
  const filterable=S.columns.filter(c=>!core.has(c));
  document.getElementById('filter-chip-picker').innerHTML=filterable.map(col=>
    `<div class="filter-chip ${S.filterCols.includes(col)?'selected':''}" data-col="${esc(col)}" onclick="toggleFilterCol('${esc(col)}')">${esc(col)}</div>`
  ).join('');
  renderFilterPreview();
}
function toggleFilterCol(col){
  if(S.filterCols.includes(col)){S.filterCols=S.filterCols.filter(c=>c!==col);}
  else if(S.filterCols.length<3){S.filterCols.push(col);}
  else{S.filterCols.shift();S.filterCols.push(col);}
  document.querySelectorAll('.filter-chip').forEach(c=>c.classList.toggle('selected',S.filterCols.includes(c.dataset.col)));
  renderFilterPreview();
}
function renderFilterPreview(){
  document.getElementById('filter-counter').textContent=`${S.filterCols.length} of 3 filters selected`;
  const area=document.getElementById('filter-preview-area');
  if(!S.filterCols.length){area.innerHTML=`<div style="font-size:0.82rem;color:var(--text3);padding:12px 0">No filters — full chart will display. "Export All" will export one chart per value of the last filter you add.</div>`;return;}
  const lastCol=S.filterCols[S.filterCols.length-1];
  area.innerHTML=`<div class="filter-preview-box">
    ${S.filterCols.map((col,i)=>{
      const isLast=i===S.filterCols.length-1;
      const vals=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort().slice(0,10);
      return`<div class="fpr-row">
        <span class="fpr-col">${esc(col)}${isLast?` <span style="background:var(--accent);color:#fff;border-radius:999px;padding:1px 7px;font-size:0.58rem;font-weight:700;margin-left:4px">Export All splits here</span>`:''}</span>
        <div class="fpr-vals">${vals.map(v=>`<span class="fv-pill">${esc(v)}</span>`).join('')}${vals.length>=10?'<span style="font-size:0.7rem;color:var(--text3)">+ more</span>':''}</div>
      </div>`;
    }).join('')}
  </div>`;
}
function launchChart(){S.activeFilters={};S.maxDepth=0;buildViewData();buildFilterBar();renderChart();goTo('chart');}

// ════════════════════════════════════════════════
// DATA PREP
// ════════════════════════════════════════════════
function buildViewData(){
  const {empId,empName,managerId}=S.colMap;
  let nodes=S.rawRows.map(row=>{
    const id=String(row[empId]||'').replace(/\.0$/,'').trim();
    const mgr=managerId?String(row[managerId]||'').replace(/\.0$/,'').trim():'';
    const node={id,name:String(row[empName]||'Unknown'),manager:mgr};
    S.columns.forEach(col=>{node[col]=String(row[col]||'');});
    return node;
  }).filter(n=>n.id&&!S.removedIds.has(n.id));

  const validIds=new Set(nodes.map(n=>n.id));
  nodes.forEach(n=>{if(S.managerOverrides.hasOwnProperty(n.id))n.manager=S.managerOverrides[n.id];});
  nodes.forEach(n=>{if(!validIds.has(n.manager)||n.manager===n.id)n.manager='';});

  const hasFilter=Object.values(S.activeFilters).some(v=>v);
  if(hasFilter){
    const matching=new Set(nodes.filter(n=>Object.entries(S.activeFilters).every(([c,v])=>!v||n[c]===v)).map(n=>n.id));
    const byId=Object.fromEntries(nodes.map(n=>[n.id,n]));
    const keep=new Set(matching);
    matching.forEach(id=>{
      let cur=byId[id];const visited=new Set();
      while(cur&&cur.manager&&byId[cur.manager]&&!visited.has(cur.id)){visited.add(cur.id);keep.add(cur.manager);cur=byId[cur.manager];}
    });
    nodes=nodes.filter(n=>keep.has(n.id));
  }

  S.viewData=nodes;
  S.childMap={};
  nodes.forEach(n=>{if(!S.childMap[n.manager])S.childMap[n.manager]=[];S.childMap[n.manager].push(n);});

  // Descendant count (for auto-fields)
  S.descCount={};
  function calcDesc(id){
    if(S.descCount[id]!==undefined)return S.descCount[id];
    const kids=S.childMap[id]||[];
    S.descCount[id]=kids.reduce((s,k)=>s+1+calcDesc(k.id),0);
    return S.descCount[id];
  }
  nodes.filter(n=>!n.manager).forEach(r=>calcDesc(r.id));

  // Node height from bottom (0=leaf, 1=parent of leaf, etc.)
  S.nodeHeight={};
  function calcHeight(id){
    if(S.nodeHeight[id]!==undefined)return S.nodeHeight[id];
    const kids=S.childMap[id]||[];
    S.nodeHeight[id]=kids.length?1+Math.max(...kids.map(k=>calcHeight(k.id))):0;
    return S.nodeHeight[id];
  }
  nodes.filter(n=>!n.manager).forEach(r=>calcHeight(r.id));
  // Also compute for nodes that might be orphan-ish
  nodes.forEach(n=>{if(S.nodeHeight[n.id]===undefined)calcHeight(n.id);});
}

function childrenOf(id){return S.childMap[id]||[];}
function countDescendants(id){return S.descCount[id]||0;}
function nodeHeight(id){return S.nodeHeight[id]||0;}

// ════════════════════════════════════════════════
// FILTER BAR
// ════════════════════════════════════════════════
function buildFilterBar(){
  const bar=document.getElementById('filter-bar');
  if(!S.filterCols.length){bar.style.display='none';return;}
  bar.style.display='flex';
  const allVals={};
  S.filterCols.forEach(col=>{
    allVals[col]=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();
  });
  bar.innerHTML='<span style="font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);flex-shrink:0">Filters</span>'+
    S.filterCols.map(col=>{
      const cur=S.activeFilters[col]||'';
      const opts=`<option value="">All ${esc(col)}</option>`+allVals[col].map(v=>`<option value="${esc(v)}"${cur===v?' selected':''}>${esc(v)}</option>`).join('');
      return`<div class="filter-dropdown-wrap"><span class="filter-dropdown-label">${esc(col)}</span><select class="filter-dropdown" onchange="applyFilter('${esc(col)}',this.value)">${opts}</select></div>`;
    }).join('')+
    (Object.values(S.activeFilters).some(v=>v)?`<button class="btn btn-ghost btn-sm" onclick="clearAllFilters()" style="margin-left:auto">✕ Clear All</button>`:'');
}
function applyFilter(col,val){
  if(val)S.activeFilters[col]=val;else delete S.activeFilters[col];
  requestAnimationFrame(()=>setTimeout(()=>{buildViewData();renderChart();buildFilterBar();},0));
}
function clearAllFilters(){S.activeFilters={};requestAnimationFrame(()=>setTimeout(()=>{buildViewData();renderChart();buildFilterBar();},0));}

// ════════════════════════════════════════════════
// BOTTOM-UP LEVEL DEPTH FILTER
// ════════════════════════════════════════════════
function setMaxDepth(n){
  S.maxDepth=n;
  const ds=document.getElementById('depth-select');if(ds)ds.value=n;
  renderChart();
}

// ════════════════════════════════════════════════
// ORG CHART RENDER
// ════════════════════════════════════════════════
function getSlotVal(node,slot){
  const f=S.cardSlots[slot];if(!f)return '';
  if(f==='__auto_reports__')return childrenOf(node.id).length+' reports';
  if(f==='__auto_teamsize__')return countDescendants(node.id)+' people';
  return String(node[f]||'').substring(0,28);
}

function renderChart(){
  const tree=document.getElementById('org-tree');
  tree.innerHTML='';

  const ds=document.getElementById('depth-select');if(ds)ds.value=S.maxDepth;
  const depthInfo=document.getElementById('stat-depth-info');
  const depthVal=document.getElementById('stat-depth-val');
  if(depthInfo&&depthVal){
    depthInfo.style.display=S.maxDepth>0?'flex':'none';
    if(S.maxDepth>0)depthVal.textContent=S.maxDepth;
  }

  let roots;
  if(S.maxDepth>0){
    // Bottom-up: visible = nodes whose nodeHeight < maxDepth
    // Virtual root = visible node whose manager is not visible (or has no manager)
    const visible=new Set(S.viewData.filter(n=>(S.nodeHeight[n.id]||0)<S.maxDepth).map(n=>n.id));
    roots=S.viewData.filter(n=>visible.has(n.id)&&(!n.manager||!visible.has(n.manager)));
  } else {
    roots=S.viewData.filter(n=>!n.manager);
  }

  if(!roots.length){
    tree.innerHTML=`<div class="no-data">No nodes found for this depth setting.<br>Try increasing the Levels ↑ selector or selecting "All".</div>`;
    updateStats(roots);return;
  }

  const ul=document.createElement('ul');
  roots.forEach(r=>ul.appendChild(mkNodeLI(r)));
  tree.appendChild(ul);
  updateStats(roots);
  clearTimeout(window._fit);
  window._fit=setTimeout(()=>fitToScreen(true),180);
}

function mkNodeLI(node){
  const li=document.createElement('li');
  li.dataset.id=node.id;

  const vacant=isVacant(node);
  const ac=S.cardAccent;
  const card=document.createElement('div');
  card.className='node-card'+(node.id===S.highlighted?' highlighted':'')+(vacant?' vacant':'');
  if(!vacant)card.style.borderTopColor=ac;

  const acLight=ac+'18', acMid=ac+'55';
  const reports=childrenOf(node.id).length;
  const badge1=getSlotVal(node,'badge1'),badge2=getSlotVal(node,'badge2');
  const subtitle=getSlotVal(node,'subtitle');
  const footer1=getSlotVal(node,'footer1')||node.id.substring(0,14);
  const footer2=getSlotVal(node,'footer2');
  const hasR=reports>0;
  const initials=node.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();

  // Photo: look up in photoMap
  const photoUrl=getPhotoUrl(node);
  let photoHtml='';
  if(photoUrl){
    photoHtml=`<img class="ncard-photo" src="${esc(photoUrl)}" crossorigin="anonymous" onerror="this.onerror=null;this.style.display='none';this.nextElementSibling.style.display='flex'">` +
      `<div class="ncard-photo-fallback" style="display:none;background:${acLight};color:${ac};border:2px solid ${acMid}">${esc(initials)}</div>`;
  } else if(Object.keys(S.photoMap).length>0){
    // Photos are loaded but no match → show fallback avatar
    photoHtml=`<div class="ncard-photo-fallback" style="display:flex;background:${acLight};color:${ac};border:2px solid ${acMid}">${esc(initials)}</div>`;
  }
  // If no photos loaded at all, no photo area shown

  card.innerHTML=
    '<div class="ncard-header">'+
      (vacant?'<span class="vacant-badge">🔴 Vacant</span>':
        badge1?`<span class="ncard-badge" title="${esc(badge1)}" style="background:${acLight};color:${ac};border-color:${acMid}">${esc(badge1)}</span>`:'<span></span>')+
      (badge2&&!vacant?`<span class="ncard-badge2">${esc(badge2)}</span>`:'')+
    '</div>'+
    '<div class="ncard-body"><div class="ncard-body-inner">'+
      (photoHtml?`<div style="flex-shrink:0">${photoHtml}</div>`:'')+
      '<div class="ncard-text-wrap">'+
        `<div class="ncard-name" title="${esc(node.name)}">${esc(node.name)}</div>`+
        (subtitle?`<div class="ncard-sub" title="${esc(subtitle)}">${esc(subtitle)}</div>`:'')+
      '</div>'+
    '</div></div>'+
    '<div class="ncard-footer">'+
      `<span class="ncard-fl">${esc(footer1)}</span>`+
      (footer2?`<span class="ncard-fr has-r" style="background:${acLight};color:${ac}">${esc(footer2)}</span>`:
        `<span class="ncard-fr${hasR?' has-r':''}" ${hasR?`style="background:${acLight};color:${ac}"`:''}>` +
        `${reports} ${reports===1?'report':'reports'}</span>`)+
    '</div>'+
    `<div class="ncard-edit-btn" title="Reassign manager" onclick="openReassignModal(event,'${esc(node.id)}')">✎</div>`;

  if(hasR){
    const cb=document.createElement('div');
    cb.className='collapse-btn';cb.innerHTML='▾';cb.title='Collapse / expand';
    cb.addEventListener('click',e=>{e.stopPropagation();toggleCollapse(li,cb);});
    card.appendChild(cb);
  }
  li.appendChild(card);

  // Children: respect bottom-up depth limit
  // In bottom-up mode, only render children that are still within the visible set
  let kids=childrenOf(node.id);
  if(S.maxDepth>0){
    kids=kids.filter(k=>(S.nodeHeight[k.id]||0)<S.maxDepth);
  }
  if(kids.length){
    const ul=document.createElement('ul');
    kids.forEach(k=>ul.appendChild(mkNodeLI(k)));
    li.appendChild(ul);
  }
  return li;
}

function toggleCollapse(li,btn){li.classList.toggle('collapsed');const c=li.classList.contains('collapsed');btn.innerHTML=c?'▸':'▾';btn.style.color=c?'var(--warning)':'';li.querySelector('.node-card')?.classList.toggle('collapsed-node',c);setTimeout(()=>updateStats(),60);}
function expandAll(){document.querySelectorAll('li.collapsed').forEach(li=>{li.classList.remove('collapsed');li.querySelector('.node-card')?.classList.remove('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▾';b.style.color='';}});setTimeout(()=>updateStats(),60);}
function collapseAll(){document.querySelectorAll('li').forEach(li=>{if(!li.parentElement?.parentElement?.closest('li'))return;if(li.querySelector(':scope > ul')){li.classList.add('collapsed');li.querySelector('.node-card')?.classList.add('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▸';b.style.color='var(--warning)';}}}); setTimeout(()=>updateStats(),60);}
function updateStats(roots){
  if(!roots)roots=S.maxDepth>0?S.viewData.filter(n=>(S.nodeHeight[n.id]||0)===S.maxDepth-1||(S.nodeHeight[n.id]||0)<S.maxDepth&&(!n.manager||!((S.nodeHeight[n.manager]||0)<S.maxDepth))):S.viewData.filter(n=>!n.manager);
  document.getElementById('stat-total').textContent=S.viewData.length;
  document.getElementById('stat-roots').textContent=roots.length;
  document.getElementById('stat-vis').textContent=document.querySelectorAll('.node-card').length;
  document.getElementById('stat-filtered').style.display=Object.values(S.activeFilters).some(v=>v)?'flex':'none';
}

// ════════════════════════════════════════════════
// ZOOM & PAN
// ════════════════════════════════════════════════
function cwrap(){return document.getElementById('chart-canvas-wrap');}
function ccontent(){return document.getElementById('chart-canvas-content');}
function applyZoom(z){S.zoom=Math.max(0.1,Math.min(3,z));ccontent().style.transform='scale('+S.zoom+')';document.getElementById('zoom-level').textContent=Math.round(S.zoom*100)+'%';}
function zoomBy(d){applyZoom(S.zoom+d);}
function fitToScreen(andCenter){
  requestAnimationFrame(()=>{
    const tree=document.getElementById('org-tree');const wrap=cwrap();if(!tree||!wrap)return;
    const tw=tree.scrollWidth,th=tree.scrollHeight,aw=wrap.clientWidth-100,ah=wrap.clientHeight-100;
    if(tw<10||th<10)return;applyZoom(Math.max(0.12,Math.min(1,aw/tw,ah/th)));
    if(andCenter)setTimeout(centerView,70);
  });
}
function centerView(){const wrap=cwrap();const tree=document.getElementById('org-tree');if(!wrap||!tree)return;const sw=tree.scrollWidth*S.zoom;wrap.scrollLeft=Math.max(0,(sw-wrap.clientWidth)/2);wrap.scrollTop=0;}
let _panning=false,_px,_py,_psl,_pst;
function initPan(){
  const wrap=cwrap();if(!wrap)return;
  wrap.onmousedown=e=>{if(e.target.closest('.node-card,.collapse-btn'))return;_panning=true;_px=e.clientX;_py=e.clientY;_psl=wrap.scrollLeft;_pst=wrap.scrollTop;wrap.style.cursor='grabbing';};
  window.onmousemove=e=>{if(!_panning)return;cwrap().scrollLeft=_psl-(e.clientX-_px);cwrap().scrollTop=_pst-(e.clientY-_py);};
  window.onmouseup=()=>{_panning=false;if(cwrap())cwrap().style.cursor='';};
  wrap.addEventListener('wheel',e=>{if(e.ctrlKey||e.metaKey){e.preventDefault();zoomBy(e.deltaY<0?0.08:-0.08);}},{passive:false});
}

// ════════════════════════════════════════════════
// SEARCH
// ════════════════════════════════════════════════
function initSearch(){
  const input=document.getElementById('chart-search');const box=document.getElementById('chart-search-results');if(!input)return;
  input.addEventListener('input',function(){
    const q=this.value.trim().toLowerCase();if(!q){box.classList.remove('visible');return;}
    const hits=S.viewData.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)).slice(0,8);
    box.innerHTML=hits.length?hits.map(n=>`<div class="sr-item" onclick="highlightNode('${esc(n.id)}')"><div class="sr-name">${esc(n.name)}</div><div class="sr-sub">${esc(getSlotVal(n,'subtitle')||n.id)}</div></div>`).join(''):'<div class="sr-item" style="color:var(--text3);font-size:0.8rem">No results</div>';
    box.classList.add('visible');
  });
  document.addEventListener('click',e=>{if(!e.target.closest('.search-wrap'))box.classList.remove('visible');});
}
function highlightNode(id){
  document.querySelectorAll('.node-card.highlighted').forEach(c=>c.classList.remove('highlighted'));
  S.highlighted=id;expandAll();
  const li=document.querySelector(`li[data-id="${CSS.escape(id)}"]`);
  if(li){
    const card=li.querySelector('.node-card');if(card){
      card.classList.add('highlighted');
      setTimeout(()=>{const r=card.getBoundingClientRect();const w=cwrap();const wr=w.getBoundingClientRect();w.scrollTo({left:w.scrollLeft+(r.left-wr.left)-wr.width/2+r.width/2,top:w.scrollTop+(r.top-wr.top)-wr.height/2+r.height/2,behavior:'smooth'});},80);
    }
  }
  document.getElementById('chart-search').value='';document.getElementById('chart-search-results').classList.remove('visible');
}

// ════════════════════════════════════════════════
// SHARED EXPORT HELPERS
// ════════════════════════════════════════════════
function inlineStyles(root){
  const PROPS=['color','backgroundColor','borderTopColor','borderBottomColor','borderLeftColor','borderRightColor','borderTopWidth','borderTopStyle','borderRadius','fontFamily','fontSize','fontWeight','fontStyle','lineHeight','padding','paddingTop','paddingBottom','paddingLeft','paddingRight','margin','display','flexDirection','justifyContent','alignItems','gap','whiteSpace','overflow','textOverflow','opacity','boxShadow','backgroundImage'];
  root.querySelectorAll('*').forEach(el=>{
    const cs=window.getComputedStyle(el);PROPS.forEach(p=>{el.style[p]=cs[p];});
    el.style.webkitLineClamp='unset';el.style.overflow='visible';el.classList.remove('collapsed');
  });
}

async function buildRenderStage(){
  const stage=document.createElement('div');
  stage.style.cssText='position:fixed;top:0;left:-99999px;background:#ffffff;padding:56px;display:inline-block;white-space:nowrap;z-index:-999;pointer-events:none';
  const cloned=document.getElementById('org-tree').cloneNode(true);
  cloned.querySelectorAll('ul').forEach(ul=>{ul.style.display='';});
  cloned.querySelectorAll('li').forEach(li=>{li.classList.remove('collapsed');});
  cloned.querySelectorAll('.collapse-btn,.ncard-edit-btn').forEach(b=>b.remove());
  stage.appendChild(cloned);document.body.appendChild(stage);
  await new Promise(r=>setTimeout(r,220));
  if(document.fonts?.ready)await document.fonts.ready;
  await new Promise(r=>setTimeout(r,80));
  inlineStyles(stage);
  await new Promise(r=>setTimeout(r,60));
  return stage;
}

async function renderToCanvas(stage){
  return html2canvas(stage,{backgroundColor:'#ffffff',scale:2,useCORS:true,logging:false,allowTaint:true,foreignObjectRendering:false,width:stage.scrollWidth,height:stage.scrollHeight,scrollX:0,scrollY:0});
}

function triggerDownload(blob,fname){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=fname;a.click();URL.revokeObjectURL(url);}

function csvEsc(v){return '"'+String(v??'').replace(/"/g,'""')+'"';}
function buildCSVContent(){
  const cols=[S.colMap.empId,S.colMap.empName,S.colMap.managerId,...S.columns.filter(c=>c!==S.colMap.empId&&c!==S.colMap.empName&&c!==S.colMap.managerId)].filter(Boolean);
  let csv=cols.map(csvEsc).join(',')+'\n';
  S.viewData.forEach(n=>{csv+=cols.map(c=>csvEsc(n[c]||'')).join(',')+'\n';});
  return csv;
}
function downloadCSV(){
  triggerDownload(new Blob([buildCSVContent()],{type:'text/csv;charset=utf-8;'}),'orgchart_export.csv');
}

// ════════════════════════════════════════════════
// PNG EXPORT
// ════════════════════════════════════════════════
async function exportPNG(){
  const overlay=makeOverlay('Rendering org chart…','Large charts may take a moment');
  document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,120));
  let stage;
  try{
    stage=await buildRenderStage();
    const canvas=await renderToCanvas(stage);
    const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');
    const fp=Object.values(S.activeFilters).filter(Boolean).map(v=>v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');
    await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,`orgchart_${fp?fp+'_':''}${stamp}.png`);res();}, 'image/png'));
  }catch(e){alert('PNG export failed: '+e.message);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);}
}

function makeOverlay(title,sub){
  const o=document.createElement('div');o.className='export-overlay';
  o.innerHTML=`<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px" id="_ov_title">${title}</div><div style="font-size:0.75rem;color:#94a3b8;margin-top:4px" id="_ov_sub">${sub}</div>`;
  return o;
}

// ════════════════════════════════════════════════
// PPTX HELPERS
// ════════════════════════════════════════════════
function xe(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function pptxRect(id,x,y,cx,cy,fill){return`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="r${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;}
function pptxTxt(id,x,y,cx,cy,text,sz,bold,color,algn='ctr'){return`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="${algn}"/><a:r><a:rPr lang="en-US" sz="${sz}" b="${bold?1:0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${xe(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;}
function pptxSlideWrap(bg,content,rels){
  const SW=12192000,SH=6858000;
  return[`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/><a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>${content}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,rels];
}

async function buildPPTXBlob(imgB64,cW,cH,titleSuffix=''){
  const ac=S.cardAccent.replace('#','');
  const stamp=new Date().toLocaleDateString('en-IN',{day:'numeric',month:'long',year:'numeric'});
  const activeF=Object.entries(S.activeFilters).filter(([,v])=>v);
  const filterLine=activeF.map(([k,v])=>`${k}: ${v}`).join('  |  ')||(titleSuffix||'All Employees');
  const roots=S.viewData.filter(n=>!n.manager).length;
  const vacants=S.vacantCol&&S.vacantVal?S.viewData.filter(n=>n[S.vacantCol]===S.vacantVal).length:0;
  const SW=12192000,SH=6858000,HBAR=457200,PAD=457200;
  const aspect=cW/cH;const avW=SW-PAD*2,avH=SH-HBAR-PAD*2;
  let iW,iH;if(aspect>avW/avH){iW=avW;iH=Math.round(avW/aspect);}else{iH=avH;iW=Math.round(avH*aspect);}
  const iX=Math.round((SW-iW)/2),iY=HBAR+Math.round((SH-HBAR-iH)/2);

  const [s1xml,s1rels]=pptxSlideWrap('F8FAFC',
    pptxRect(2,'0','0',SW,HBAR,ac)+
    pptxTxt(3,PAD,Math.round(SH*0.28),SW-PAD*2,1200000,'Org Chart',5400,true,'0F172A')+
    pptxTxt(4,PAD,Math.round(SH*0.28)+1300000,SW-PAD*2,500000,`Generated: ${stamp}`,1800,false,'64748B')+
    pptxTxt(5,PAD,Math.round(SH*0.28)+1900000,SW-PAD*2,450000,filterLine,1600,true,ac),
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`);

  const picXml=`<p:pic><p:nvPicPr><p:cNvPr id="10" name="OrgChart"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${iX}" y="${iY}"/><a:ext cx="${iW}" cy="${iH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const [s2xml,]=pptxSlideWrap('F1F5F9',
    pptxRect(2,'0','0',SW,HBAR,ac)+
    pptxTxt(3,PAD,0,SW*0.4,HBAR,'Org Chart',1600,true,'FFFFFF','l')+
    pptxTxt(4,SW-2400000,0,2200000,HBAR,`${S.viewData.length} employees`,1400,false,'FFFFFF','r')+picXml,
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`);

  const statItems=[{label:'Total Employees',val:S.viewData.length},{label:'Root Nodes',val:roots},{label:'Vacant',val:vacants},{label:'Depth',val:S.maxDepth>0?`${S.maxDepth} lvls`:'All'}];
  const boxW=Math.round(SW*0.18),boxH=Math.round(SH*0.22),totalBW=boxW*4+Math.round(SW*0.025)*3,bStartX=Math.round((SW-totalBW)/2),bY=Math.round(SH*0.32);
  let sc=pptxRect(2,'0','0',SW,HBAR,ac)+pptxTxt(3,PAD,0,SW*0.5,HBAR,'Summary',1600,true,'FFFFFF','l');
  statItems.forEach((st,i)=>{const bx=bStartX+i*(boxW+Math.round(SW*0.025));sc+=pptxRect(10+i*2,bx,bY,boxW,boxH,'F8FAFC');sc+=pptxTxt(11+i*2,bx,bY+Math.round(boxH*0.1),boxW,Math.round(boxH*0.55),String(st.val),3600,true,ac);sc+=pptxTxt(20+i,bx,bY+Math.round(boxH*0.65),boxW,Math.round(boxH*0.3),st.label,1200,false,'64748B');});
  if(filterLine!=='All Employees')sc+=pptxTxt(30,PAD,Math.round(SH*0.72),SW-PAD*2,380000,'Filters: '+filterLine,1200,false,'94A3B8');
  const [s3xml,]=pptxSlideWrap('FFFFFF',sc,`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`);

  const BOILERPLATE={
    '[Content_Types].xml':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
    '_rels/.rels':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    'ppt/presentation.xml':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst><p:sldSz cx="${SW}" cy="${SH}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>`,
    'ppt/theme/theme1.xml':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="OrgTheme"><a:themeElements><a:clrScheme name="OrgScheme"><a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1><a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1><a:dk2><a:srgbClr val="0F172A"/></a:dk2><a:lt2><a:srgbClr val="F1F5F9"/></a:lt2><a:accent1><a:srgbClr val="${ac}"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="EF4444"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="06B6D4"/></a:accent6><a:hlink><a:srgbClr val="${ac}"/></a:hlink><a:folHlink><a:srgbClr val="64748B"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
    'ppt/slideMasters/slideMaster1.xml':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lstStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></a:lstStyle></p:titleStyle><p:bodyStyle><a:lstStyle/></p:bodyStyle><p:otherStyle><a:lstStyle/></p:otherStyle></p:txStyles></p:sldMaster>`,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
    'ppt/slideLayouts/slideLayout1.xml':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels':`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`
  };

  const mkRels=r=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${r}</Relationships>`;
  const zip=new JSZip();
  Object.entries(BOILERPLATE).forEach(([k,v])=>zip.file(k,v));
  zip.file('ppt/slides/slide1.xml',s1xml);
  zip.file('ppt/slides/_rels/slide1.xml.rels',mkRels(s1rels));
  zip.file('ppt/slides/slide2.xml',s2xml);
  zip.file('ppt/slides/_rels/slide2.xml.rels',mkRels(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`));
  zip.file('ppt/slides/slide3.xml',s3xml);
  zip.file('ppt/slides/_rels/slide3.xml.rels',mkRels(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`));
  zip.file('ppt/media/image1.png',imgB64,{base64:true});
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'DEFLATE'});
}

async function exportPPTX(){
  if(typeof JSZip==='undefined'){alert('ZIP library failed to load.');return;}
  const overlay=makeOverlay('Building PowerPoint…','Rendering then packaging slides');
  document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,120));
  let stage;
  try{
    stage=await buildRenderStage();
    const canvas=await renderToCanvas(stage);
    const blob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height);
    const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');
    const fp=Object.values(S.activeFilters).filter(Boolean).map(v=>v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');
    triggerDownload(blob,`orgchart_${fp?fp+'_':''}${dp}.pptx`);
  }catch(e){alert('PPTX failed: '+e.message);console.error(e);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);}
}

// ════════════════════════════════════════════════
// EXPORT ALL — iterates last filter column values
// ════════════════════════════════════════════════
async function exportAll(){
  if(typeof JSZip==='undefined'){alert('ZIP library failed to load.');return;}

  const lastFilterCol=S.filterCols[S.filterCols.length-1]||null;

  // ── No last filter: export current view as CSV+PNG+PPTX ──
  if(!lastFilterCol){
    await _exportAllSingleView();
    return;
  }

  // ── Collect values for last filter, respecting any already-active parent filters ──
  const parentFilters=Object.entries(S.activeFilters).filter(([k])=>k!==lastFilterCol);
  const relevantRows=S.rawRows.filter(row=>parentFilters.every(([col,val])=>!val||String(row[col]||'').trim()===val));
  const lastVals=[...new Set(relevantRows.map(r=>String(r[lastFilterCol]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();

  if(!lastVals.length){await _exportAllSingleView();return;}

  const overlay=document.createElement('div');
  overlay.className='export-overlay';
  overlay.innerHTML=`<div class="export-spinner"></div>
    <div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px" id="_ea_title">Exporting ${lastVals.length} charts…</div>
    <div class="export-steps">
      <div id="_ea_step" class="export-step active" style="min-width:320px">Preparing…</div>
      <div id="_ea_prog" style="font-size:0.7rem;color:var(--text3);margin-top:4px">0 / ${lastVals.length}</div>
    </div>`;
  document.body.appendChild(overlay);

  const savedZoom=S.zoom; applyZoom(1); await new Promise(r=>setTimeout(r,120));
  const savedFilters={...S.activeFilters};

  const outerZip=new JSZip();
  let successCount=0;

  try{
    for(let i=0;i<lastVals.length;i++){
      const val=lastVals[i];
      const safeName=val.replace(/[^a-zA-Z0-9]/g,'_');

      // Update overlay
      const stepEl=document.getElementById('_ea_step');
      const progEl=document.getElementById('_ea_prog');
      if(stepEl) stepEl.textContent=`📊 ${val}`;
      if(progEl) progEl.textContent=`${i+1} / ${lastVals.length}`;

      // Apply filter
      S.activeFilters[lastFilterCol]=val;
      buildViewData();
      renderChart();
      await new Promise(r=>setTimeout(r,300));

      // CSV
      const csvContent=buildCSVContent();
      outerZip.file(`${safeName}/${safeName}.csv`,csvContent);

      // Render chart to canvas
      const stage=await buildRenderStage();
      let canvas;
      try{canvas=await renderToCanvas(stage);}finally{stage.remove();}

      // PNG
      const pngBlob=await new Promise(res=>canvas.toBlob(res,'image/png'));
      if(pngBlob) outerZip.file(`${safeName}/${safeName}.png`,pngBlob);

      // PPTX
      try{
        const pptxBlob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height,val);
        outerZip.file(`${safeName}/${safeName}.pptx`,pptxBlob);
      }catch(pe){console.warn('PPTX failed for '+val,pe);}

      successCount++;
    }

    // Done — generate the outer ZIP
    const titleEl=document.getElementById('_ea_title');
    if(titleEl) titleEl.textContent=`Packaging ${successCount} exports…`;
    const masterBlob=await outerZip.generateAsync({type:'blob',compression:'DEFLATE'});
    const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');
    const colSafe=lastFilterCol.replace(/[^a-zA-Z0-9]/g,'_');
    triggerDownload(masterBlob,`orgchart_all_${colSafe}_${dp}.zip`);

    if(titleEl) titleEl.textContent=`✅ ${successCount} charts exported!`;
    await new Promise(r=>setTimeout(r,1200));

  }catch(e){
    alert('Export All failed: '+e.message);console.error(e);
  }finally{
    // Restore previous state
    S.activeFilters=savedFilters;
    buildViewData();
    buildFilterBar();
    renderChart();
    overlay.remove();
    applyZoom(savedZoom);
  }
}

// Fallback: no last filter — export current view
async function _exportAllSingleView(){
  const overlay=document.createElement('div');
  overlay.className='export-overlay';
  overlay.innerHTML=`<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px">Exporting current view…</div>
    <div class="export-steps">
      <div class="export-step active" id="_sv_s1">💾 CSV</div>
      <div class="export-step" id="_sv_s2">🖼️ PNG</div>
      <div class="export-step" id="_sv_s3">📊 PPTX</div>
    </div>`;
  document.body.appendChild(overlay);
  const setStep=n=>[1,2,3].forEach(i=>{const el=document.getElementById(`_sv_s${i}`);if(el)el.className='export-step'+(i<n?' done':i===n?' active':'');});
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,120));
  let stage;
  try{
    setStep(1);downloadCSV();await new Promise(r=>setTimeout(r,400));
    setStep(2);
    stage=await buildRenderStage();
    const canvas=await renderToCanvas(stage);
    const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');
    await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,`orgchart_${stamp}.png`);res();},'image/png'));
    await new Promise(r=>setTimeout(r,300));
    setStep(3);
    const blob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height);
    triggerDownload(blob,`orgchart_${stamp}.pptx`);
    [1,2,3].forEach(i=>{const el=document.getElementById(`_sv_s${i}`);if(el)el.className='export-step done';});
    await new Promise(r=>setTimeout(r,900));
  }catch(e){alert('Export failed: '+e.message);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);}
}

// ════════════════════════════════════════════════
// REASSIGN
// ════════════════════════════════════════════════
function openReassignModal(e,nodeId){
  e.stopPropagation();
  const node=S.viewData.find(n=>n.id===nodeId);if(!node)return;
  S.reassignTarget=node;S.reassignPick=null;
  document.getElementById('reassign-subject').innerHTML=`Moving <strong>${esc(node.name)}</strong>`;
  document.getElementById('reassign-note').textContent='Select a new manager above';
  document.getElementById('reassign-confirm-btn').disabled=true;
  const desc=getAllDescendants(nodeId);desc.add(nodeId);
  renderReassignList(S.viewData.filter(n=>!desc.has(n.id)),node.manager);
  document.getElementById('reassign-search').value='';
  document.getElementById('reassign-modal').classList.remove('hidden');
  document.getElementById('reassign-search').focus();
}
function getAllDescendants(id){const r=new Set();const q=[...(S.childMap[id]||[])];while(q.length){const n=q.shift();r.add(n.id);(S.childMap[n.id]||[]).forEach(k=>q.push(k));}return r;}
function renderReassignList(candidates,curMgrId){
  const list=document.getElementById('reassign-list');
  let html=`<div class="modal-emp-row make-root${S.reassignPick==='__ROOT__'?' selected':''}" onclick="pickReassignTarget('__ROOT__',this)"><div class="modal-emp-avatar" style="background:#fef9c3;color:#d97706;font-size:0.9rem">🏠</div><div><div class="modal-emp-name">Make Root (no manager)</div><div class="modal-emp-sub">Move to top level</div></div></div>`;
  html+=candidates.map(n=>{
    const isCur=n.id===curMgrId;const isSel=S.reassignPick&&S.reassignPick.id===n.id;
    const init=n.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();
    return`<div class="modal-emp-row${isSel?' selected':''}" onclick="pickReassignTarget('${esc(n.id)}',this)"><div class="modal-emp-avatar${isVacant(n)?' vacant-av':''}">${init}</div><div style="flex:1;min-width:0"><div class="modal-emp-name">${esc(n.name)}${isCur?' <span style="color:var(--text3);font-weight:500;font-size:0.68rem">(current)</span>':''}</div><div class="modal-emp-sub">${esc(getSlotVal(n,'subtitle')||n.id)}</div></div></div>`;
  }).join('');
  list.innerHTML=html;
}
function filterReassignList(){
  const q=document.getElementById('reassign-search').value.trim().toLowerCase();
  const node=S.reassignTarget;if(!node)return;
  const desc=getAllDescendants(node.id);desc.add(node.id);
  const all=S.viewData.filter(n=>!desc.has(n.id));
  renderReassignList(q?all.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)):all,node.manager);
}
function pickReassignTarget(id,rowEl){
  document.querySelectorAll('#reassign-list .modal-emp-row').forEach(r=>r.classList.remove('selected'));
  rowEl.classList.add('selected');
  if(id==='__ROOT__'){S.reassignPick='__ROOT__';document.getElementById('reassign-note').textContent='→ Will become a root node';}
  else{S.reassignPick=S.viewData.find(n=>n.id===id)||null;document.getElementById('reassign-note').textContent=S.reassignPick?`→ New manager: ${S.reassignPick.name}`:'';}
  document.getElementById('reassign-confirm-btn').disabled=false;
}
function confirmReassign(){
  if(!S.reassignTarget)return;
  const newMgr=S.reassignPick==='__ROOT__'?'':(S.reassignPick?S.reassignPick.id:null);
  if(newMgr===null)return;
  S.managerOverrides[S.reassignTarget.id]=newMgr;
  closeReassignModal();buildViewData();renderChart();
}
function closeReassignModal(){document.getElementById('reassign-modal').classList.add('hidden');S.reassignTarget=null;S.reassignPick=null;}
function removeCurrentNode(){
  if(!S.reassignTarget)return;
  const{id,name}=S.reassignTarget;
  if(!confirm(`Remove "${name}" from the chart?`))return;
  S.removedIds.add(id);delete S.managerOverrides[id];
  closeReassignModal();buildViewData();renderChart();
}

// ════════════════════════════════════════════════
// UTILITY
// ════════════════════════════════════════════════
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');}

// ════════════════════════════════════════════════
// INIT
// ════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded',()=>{
  const zone=document.getElementById('upload-dropzone');
  const input=document.getElementById('file-input');
  zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('drag-over');});
  zone.addEventListener('dragleave',()=>zone.classList.remove('drag-over'));
  zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('drag-over');const f=e.dataTransfer.files[0];if(f)handleFile(f);});
  input.addEventListener('change',e=>{if(e.target.files[0])handleFile(e.target.files[0]);});

  // Photo folder fallback input (multiple image files)
  const photoInput=document.getElementById('photo-folder-input');
  photoInput.addEventListener('change',e=>{if(e.target.files.length)loadFromFileInput(e.target.files);});
});
</script>
</body>
</html>"""

components.html(APP_HTML, height=870, scrolling=False)
