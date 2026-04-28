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

APP_HTML = r'''<!DOCTYPE html>
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
.main{flex:1;overflow:hidden;position:relative}
.screen{position:absolute;inset:0;overflow-y:auto;display:flex;flex-direction:column;padding:32px 36px;background:var(--bg);opacity:0;pointer-events:none;transform:translateX(18px);transition:opacity 0.22s ease,transform 0.22s ease}
.screen.active{opacity:1;pointer-events:auto;transform:translateX(0)}
#screen-chart{padding:0;overflow:hidden}
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
.section-header{margin-bottom:24px}
.section-title{font-family:'Syne',sans-serif;font-weight:700;font-size:1.45rem;color:var(--text);letter-spacing:-0.02em;margin-bottom:4px}
.section-sub{font-size:0.84rem;color:var(--text2)}
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
/* CARD DESIGN */
.card-design-layout{display:grid;grid-template-columns:280px 1fr;gap:24px;flex:1;min-height:0}
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
.preview-card{width:320px;background:var(--bg);border:2px solid var(--border);border-top:4px solid var(--accent);border-radius:var(--r-lg);box-shadow:var(--shadow-md)}
/* 3-slot header/footer rows */
.preview-card-header{padding:7px 10px;background:var(--bg2);border-bottom:1px solid var(--border);border-radius:12px 12px 0 0;display:flex;align-items:center;gap:5px}
.preview-card-body{padding:12px 14px}
.preview-card-footer{padding:7px 10px;border-top:1px solid var(--border);border-radius:0 0 12px 12px;background:var(--bg2);display:flex;align-items:center;gap:5px}
.card-zone{flex:1;min-height:26px;min-width:50px;border:2px dashed var(--border2);border-radius:7px;padding:3px 6px;font-size:0.68rem;color:var(--text3);display:flex;align-items:center;justify-content:center;transition:all 0.15s;position:relative;cursor:default}
.card-zone .zone-ph{opacity:0.6;font-style:italic}
.card-zone.drop-target{border-color:var(--accent);background:var(--accent-light)}
.card-zone.filled{border-style:solid;border-color:var(--accent-mid);background:var(--accent-light);flex-direction:column;gap:1px;align-items:flex-start;justify-content:center}
.zone-field{font-weight:700;font-size:0.67rem;color:var(--accent)}
.zone-val{font-size:0.65rem;color:var(--text2);font-style:italic}
.zone-remove{position:absolute;top:2px;right:3px;font-size:0.58rem;cursor:pointer;opacity:0.5;line-height:1}
.zone-remove:hover{opacity:1}
.preview-hint{font-size:0.76rem;color:var(--text3);max-width:320px;line-height:1.5}
.preview-name-fixed{display:flex;align-items:center;gap:6px;background:var(--bg3);border:1px dashed var(--border2);border-radius:6px;padding:6px 10px}
.preview-name-fixed span{font-size:0.8rem;color:var(--text2);font-weight:600}
.preview-name-fixed .lock{font-size:0.75rem;opacity:0.5}
.ncard-photo{object-fit:cover;object-position:center top;display:block;flex-shrink:0}
.ncard-photo-fallback{font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0;letter-spacing:-0.02em}
/* Employment type color setup */
.emp-type-setup{margin-top:14px}
.emp-type-row{display:flex;align-items:center;gap:8px;margin-bottom:7px}
.emp-type-label{font-size:0.78rem;font-weight:700;color:var(--text2);min-width:80px}
.emp-type-color-input{width:34px;height:28px;border-radius:7px;border:2px solid var(--border2);cursor:pointer;padding:0;background:none}
.emp-type-value-select{flex:1;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:4px 8px;font-size:0.76rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;cursor:pointer}
/* Filter screen */
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
.btn{padding:9px 20px;border-radius:var(--r);font-size:0.84rem;font-weight:700;cursor:pointer;border:none;transition:all 0.15s;display:inline-flex;align-items:center;gap:7px;font-family:'Plus Jakarta Sans',sans-serif;line-height:1;white-space:nowrap}
.btn-primary{background:var(--accent);color:#fff;box-shadow:0 4px 14px rgba(79,70,229,0.3)}
.btn-primary:hover{background:#4338ca;transform:translateY(-1px);box-shadow:0 6px 20px rgba(79,70,229,0.4)}
.btn-ghost{background:transparent;color:var(--text2);border:1.5px solid var(--border)}
.btn-ghost:hover{background:var(--bg3);color:var(--text);border-color:var(--border2)}
.btn-sm{padding:6px 13px;font-size:0.78rem;border-radius:8px}
.btn-row{display:flex;gap:10px;margin-top:28px}
.btn-export-all{background:linear-gradient(135deg,#7c3aed,#0284c7)!important;color:#fff!important;border:none!important;box-shadow:0 4px 14px rgba(124,58,237,0.35)!important}
.btn-export-all:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(124,58,237,0.45)!important}
/* CHART */
.chart-toolbar{flex-shrink:0;height:52px;background:var(--bg);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 12px;gap:6px;box-shadow:var(--shadow-xs);position:relative;z-index:20;overflow:visible}
.stats-bar{flex-shrink:0;height:34px;background:var(--bg2);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 18px;gap:18px;font-size:0.73rem}
.stat-item{display:flex;align-items:center;gap:6px;color:var(--text3);font-weight:600}
.stat-item strong{color:var(--text);font-weight:800}
.stat-dot{width:6px;height:6px;border-radius:50%;background:var(--accent)}
.filter-bar{flex-shrink:0;background:var(--bg);border-bottom:1px solid var(--border);padding:7px 18px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;min-height:44px}
.filter-dropdown-wrap{display:flex;align-items:center;gap:6px;font-size:0.79rem}
.filter-dropdown-label{font-weight:700;color:var(--text2)}
.filter-dropdown{background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:5px 28px 5px 10px;font-size:0.79rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center;transition:border-color 0.15s}
.filter-dropdown:focus{border-color:var(--accent)}
.photo-btn{display:flex;align-items:center;gap:5px;padding:5px 10px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0}
.photo-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.photo-btn.loaded{border-color:#059669;color:#059669;background:#d1fae5}
.photo-count{background:#059669;color:#fff;border-radius:999px;padding:1px 6px;font-size:0.65rem;font-weight:800}
.depth-wrap{display:flex;align-items:center;gap:5px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:3px 6px 3px 9px;flex-shrink:0}
.depth-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);white-space:nowrap}
.depth-select{background:transparent;border:none;border-radius:6px;padding:3px 20px 3px 4px;font-size:0.78rem;font-weight:700;color:var(--accent);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%234f46e5'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 3px center}
/* Manager-only mode toggle */
.mgr-mode-btn{display:flex;align-items:center;gap:6px;padding:5px 11px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0;user-select:none}
.mgr-mode-btn:hover{border-color:#7c3aed;color:#7c3aed;background:#f5f3ff}
.mgr-mode-btn.active{background:#f5f3ff;border-color:#7c3aed;color:#7c3aed;box-shadow:0 0 0 2px #ddd6fe}
.mgr-mode-dot{width:8px;height:8px;border-radius:50%;background:var(--border2);transition:background 0.15s}
.mgr-mode-btn.active .mgr-mode-dot{background:#7c3aed}
/* Summary fields selector */
.summary-fields-wrap{display:flex;align-items:center;gap:5px;background:#fdf4ff;border:1.5px solid #e9d5ff;border-radius:8px;padding:3px 6px 3px 9px;flex-shrink:0}
.summary-fields-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:#7c3aed;white-space:nowrap}
.summary-field-select{background:transparent;border:none;padding:3px 18px 3px 4px;font-size:0.75rem;font-weight:700;color:#7c3aed;font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%237c3aed'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 2px center;max-width:110px}
/* Compact summary list card (manager-only mode) */
.summary-list-card{display:inline-block;min-width:200px;max-width:280px;background:var(--bg);border:1.5px solid var(--border);border-top:3px solid #7c3aed;border-radius:var(--r-lg);box-shadow:var(--shadow-sm);font-family:'Plus Jakarta Sans',sans-serif;cursor:default}
.summary-list-header{padding:7px 12px;background:#f5f3ff;border-bottom:1px solid #e9d5ff;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center}
.summary-list-title{font-size:0.68rem;font-weight:800;color:#7c3aed;text-transform:uppercase;letter-spacing:0.05em}
.summary-list-count{font-size:0.68rem;font-weight:800;background:#7c3aed;color:#fff;border-radius:999px;padding:1px 8px}
/* NOTE: max-height/overflow-y set inline per instance so export captures all rows */
.summary-list-body{padding:4px 0}
.summary-person-row{display:flex;align-items:center;gap:8px;padding:6px 12px;border-bottom:1px solid var(--border);cursor:pointer;transition:background 0.1s}
.summary-person-row:last-child{border-bottom:none}
.summary-person-row:hover{background:var(--bg2)}
.summary-person-avatar{width:26px;height:26px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:0.6rem;font-weight:800;flex-shrink:0}
.summary-person-info{flex:1;min-width:0}
.summary-person-name{font-size:0.75rem;font-weight:700;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.summary-person-field2{font-size:0.68rem;color:var(--text3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
/* Node card */
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
.org-tree li.collapsed > ul{display:none!important}
.node-card{display:inline-block;width:270px;background:var(--bg);border:1.5px solid var(--border);border-top:3px solid var(--accent);border-radius:var(--r-lg);cursor:pointer;text-align:left;transition:transform 0.15s,box-shadow 0.15s,border-color 0.15s;box-shadow:var(--shadow-sm);position:relative;font-family:'Plus Jakarta Sans',sans-serif}
.node-card:hover{transform:translateY(-3px);box-shadow:0 8px 28px rgba(0,0,0,0.12),0 0 0 2px rgba(79,70,229,0.12);border-color:var(--accent);z-index:10}
.node-card.highlighted{border-color:var(--warning)!important;border-top-color:var(--warning)!important;box-shadow:0 0 0 3px rgba(217,119,6,0.2),0 8px 24px rgba(0,0,0,0.1)!important}
.node-card.collapsed-node{opacity:0.65}
/* 3-slot ncard header/footer */
.ncard-header{padding:6px 10px;background:var(--bg2);border-bottom:1px solid var(--border);border-radius:12px 12px 0 0;display:flex;align-items:center;gap:4px}
.ncard-footer{padding:6px 10px;border-top:1px solid var(--border);border-radius:0 0 12px 12px;background:var(--bg2);display:flex;align-items:center;gap:4px}
.ncard-slot{font-size:0.64rem;font-weight:700;color:var(--text3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;text-align:center}
.ncard-slot.has-val{color:var(--accent)}
.ncard-body{padding:14px 14px 12px}
.ncard-body-inner{display:flex;gap:10px}
.ncard-body-b1{border-top:1px dashed var(--border2);margin-top:6px;padding-top:5px;text-align:center;font-size:0.72rem;color:var(--accent);font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ncard-text-wrap{width:100%;text-align:center}
.ncard-name{font-size:0.88rem;font-weight:800;color:var(--text);margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:-0.01em}
.ncard-sub{font-size:0.74rem;color:var(--text2);line-height:1.4;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.ncard-photo{object-fit:cover;object-position:center top;display:block;flex-shrink:0}
.ncard-photo-fallback{font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0;letter-spacing:-0.02em}
.collapse-btn{position:absolute;bottom:-11px;left:50%;transform:translateX(-50%);width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:0.58rem;color:var(--text3);transition:all 0.15s;z-index:5;box-shadow:var(--shadow-xs)}
.collapse-btn:hover{background:var(--accent);border-color:var(--accent);color:#fff}
.search-wrap{position:relative;flex:1;max-width:240px}
.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:0.8rem;pointer-events:none;opacity:0.45}
#chart-search{width:100%;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:6px 10px 6px 29px;font-size:0.8rem;font-weight:500;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;transition:border-color 0.15s}
#chart-search:focus{border-color:var(--accent);background:var(--bg)}
#chart-search::placeholder{color:var(--text3)}
#chart-search-results{position:fixed;background:var(--bg);border:1.5px solid var(--border);border-radius:var(--r);box-shadow:var(--shadow-lg);max-height:300px;overflow-y:auto;z-index:99999;display:none;min-width:260px}
#chart-search-results.visible{display:block}
.sr-item{padding:10px 14px;cursor:pointer;border-bottom:1px solid var(--border);transition:background 0.1s}
.sr-item:last-child{border-bottom:none}
.sr-item:hover{background:var(--bg3)}
.sr-name{font-weight:700;font-size:0.83rem;color:var(--text)}
.sr-sub{font-size:0.72rem;color:var(--text3);margin-top:2px}
.zoom-strip{display:flex;align-items:center;gap:1px;background:var(--bg2);border-radius:8px;padding:2px;border:1.5px solid var(--border)}
.btn-zoom{background:transparent;border:none;border-radius:6px;width:26px;height:26px;cursor:pointer;font-size:0.85rem;font-weight:700;color:var(--text2);font-family:'Plus Jakarta Sans',sans-serif;display:flex;align-items:center;justify-content:center;transition:background 0.12s}
.btn-zoom:hover{background:var(--bg3);color:var(--text)}
.zoom-label{font-size:0.72rem;font-weight:800;color:var(--text);min-width:42px;text-align:center;font-variant-numeric:tabular-nums}
.export-overlay{position:fixed;inset:0;z-index:9999;background:rgba(255,255,255,0.94);backdrop-filter:blur(8px);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px}
.export-spinner{width:44px;height:44px;border:3px solid var(--border2);border-top-color:var(--accent);border-radius:50%;animation:spin 0.7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.no-data{padding:40px;color:var(--text3);font-size:0.9rem;font-weight:600;background:var(--bg);border:1.5px solid var(--border);border-radius:var(--r-lg);max-width:440px}
.color-palette{display:flex;flex-wrap:wrap;gap:7px;margin-top:6px}
.color-swatch{width:24px;height:24px;border-radius:6px;cursor:pointer;border:2.5px solid transparent;transition:transform 0.1s,border-color 0.1s;flex-shrink:0}
.color-swatch:hover{transform:scale(1.15)}
.color-swatch.selected{border-color:var(--text);box-shadow:0 0 0 2px #fff inset}
.shape-btn{padding:5px 11px;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:600;color:var(--text2);cursor:pointer;transition:all 0.15s;user-select:none}
.shape-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.shape-btn.selected{background:var(--accent);border-color:var(--accent);color:#fff}
.vacant-setup{margin-top:14px}
.vacant-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:6px}
.vacant-select{flex:1;min-width:100px;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:5px 8px;font-size:0.78rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;appearance:none;cursor:pointer;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 7px center}
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
.modal-emp-name{font-weight:700;font-size:0.82rem;color:var(--text)}
.modal-emp-sub{font-size:0.71rem;color:var(--text3)}
.modal-footer{padding:14px 20px;border-top:1px solid var(--border);display:flex;gap:10px;justify-content:flex-end}
.modal-note{font-size:0.73rem;color:var(--text3);flex:1;display:flex;align-items:center}
.tb-sep{width:1px;height:22px;background:var(--border);flex-shrink:0}
.photo-folder-input{display:none}
/* Body slot b1 zone in card design preview */
.preview-b1-zone{margin:4px 0 0 0;padding-top:6px;border-top:1px dashed var(--border2)}
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
  <!-- UPLOAD -->
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

  <!-- MAP -->
  <div class="screen" id="screen-map">
    <div class="section-header">
      <div class="section-title">Map Your Columns</div>
      <div class="section-sub">We detected <span id="col-count">0</span> columns. Auto-mapped where possible.</div>
    </div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Detected Columns</div>
    <div class="detected-chips" id="detected-columns"></div>
    <div class="map-grid">
      <div class="map-card"><div class="map-card-label">👤 Employee ID <span class="badge-req">Required</span></div><select class="map-select" id="map-empId"></select><div class="map-hint">Unique identifier — also used to match photos</div></div>
      <div class="map-card"><div class="map-card-label">🏷️ Employee Name <span class="badge-req">Required</span></div><select class="map-select" id="map-empName"></select><div class="map-hint">Full name shown on the card</div></div>
      <div class="map-card"><div class="map-card-label">🔗 Manager ID <span class="badge-opt">Optional</span></div><select class="map-select" id="map-managerId"></select><div class="map-hint">Links employee to their manager</div></div>
    </div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Data Preview (first 3 rows)</div>
    <div id="data-preview-wrap" style="margin-bottom:24px;overflow-x:auto"></div>
    <div class="btn-row" style="margin-top:0">
      <button class="btn btn-ghost" onclick="goTo('upload')">← Back</button>
      <button class="btn btn-primary" onclick="confirmColumnMap()">Continue to Card Design →</button>
    </div>
  </div>

  <!-- CARD DESIGN -->
  <div class="screen" id="screen-card">
    <div class="section-header" style="margin-bottom:18px">
      <div class="section-title">Design Your Card</div>
      <div class="section-sub">Drag fields into 3 header, 1 body, and 3 footer slots. Configure employment type colors.</div>
    </div>
    <div class="card-design-layout">
      <div class="fields-panel">
        <div class="fields-panel-title">Available Fields</div>
        <div id="card-fields-panel"></div>

        <div class="fields-section" style="margin-top:16px">
          <div class="fields-section-label" style="margin-bottom:8px">🎨 Card Accent Color</div>
          <div class="color-palette" id="color-palette"></div>
        </div>

        <div class="fields-section" style="margin-top:14px">
          <div class="fields-section-label" style="margin-bottom:8px">🖼️ Photo Size &amp; Shape</div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:7px">Size (px)</div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
            <input type="range" id="photo-size-slider" min="40" max="160" step="10" value="80"
              style="flex:1;accent-color:var(--accent);cursor:pointer"
              oninput="S.photoSize=parseInt(this.value);document.getElementById('photo-size-val').textContent=this.value+'px';renderCardPreview();renderChart()"/>
            <span id="photo-size-val" style="font-size:0.78rem;font-weight:700;color:var(--accent);min-width:36px">80px</span>
          </div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:7px">Shape</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px">
            <div class="shape-btn selected" data-shape="circle" onclick="setPhotoShape('circle')">⬤ Circle</div>
            <div class="shape-btn" data-shape="rounded" onclick="setPhotoShape('rounded')">▪ Rounded</div>
            <div class="shape-btn" data-shape="square" onclick="setPhotoShape('square')">■ Square</div>
          </div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:7px">Placement</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <div class="shape-btn selected" data-placement="top" onclick="setPhotoPlacement('top')">⬆ Top</div>
            <div class="shape-btn" data-placement="left" onclick="setPhotoPlacement('left')">◀ Left</div>
            <div class="shape-btn" data-placement="right" onclick="setPhotoPlacement('right')">▶ Right</div>
            <div class="shape-btn" data-placement="none" onclick="setPhotoPlacement('none')">✕ None</div>
          </div>
        </div>

        <!-- Employment Type Color Mapping -->
        <div class="fields-section emp-type-setup">
          <div class="fields-section-label" style="margin-bottom:10px">🎨 Employment Type Colors</div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:10px;line-height:1.5">Map a column + values to control card border color. Overrides global accent.</div>
          <div style="margin-bottom:8px">
            <select class="vacant-select" id="emp-type-col" onchange="onEmpTypeColChange()" style="width:100%;margin-bottom:8px"><option value="">Select column…</option></select>
          </div>
          <div id="emp-type-rows" style="display:none">
            <div class="emp-type-row">
              <div class="emp-type-label">Active</div>
              <select class="emp-type-value-select" id="emp-val-active"><option value="">Value…</option></select>
              <input type="color" class="emp-type-color-input" id="emp-color-active" value="#059669" title="Color for Active"/>
            </div>
            <div class="emp-type-row">
              <div class="emp-type-label">Vacant</div>
              <select class="emp-type-value-select" id="emp-val-vacant"><option value="">Value…</option></select>
              <input type="color" class="emp-type-color-input" id="emp-color-vacant" value="#dc2626" title="Color for Vacant"/>
            </div>
            <div class="emp-type-row">
              <div class="emp-type-label">Resigned</div>
              <select class="emp-type-value-select" id="emp-val-resigned"><option value="">Value…</option></select>
              <input type="color" class="emp-type-color-input" id="emp-color-resigned" value="#d97706" title="Color for Resigned"/>
            </div>
          </div>
        </div>
      </div>

      <div class="card-preview-area">
        <div class="preview-label">Live Card Preview — 3 Header + 1 Body + 3 Footer Slots</div>
        <div id="card-preview"></div>
        <div class="preview-hint">Drag field chips onto header, body, or footer zones. Name is always fixed in the card body. Body slot appears beneath the name.</div>
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" onclick="goTo('map')">← Back</button>
      <button class="btn btn-primary" onclick="confirmCardDesign()">Continue to Filters →</button>
    </div>
  </div>

  <!-- FILTER -->
  <div class="screen" id="screen-filter">
    <div class="section-header">
      <div class="section-title">Set Up Filters</div>
      <div class="section-sub">Choose up to 3 columns to use as filters. The last filter drives "Export All".</div>
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

  <!-- CHART -->
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
      <div class="depth-wrap" title="Skip top N levels">
        <span class="depth-label">Skip Top</span>
        <select class="depth-select" id="depth-select" onchange="setSkipDepth(parseInt(this.value))">
          <option value="0">None</option>
          <option value="1">L1</option><option value="2">L2</option><option value="3">L3</option>
          <option value="4">L4</option><option value="5">L5</option><option value="6">L6</option>
        </select>
      </div>
      <div class="tb-sep"></div>
      <!-- Manager-Only Mode Toggle -->
      <div class="mgr-mode-btn" id="mgr-mode-btn" onclick="toggleManagerMode()" title="Manager View: managers shown as full cards; leaf employees (no reports) shown as compact list">
        <div class="mgr-mode-dot"></div>
        👔 Manager View
      </div>
      <!-- Summary fields selectors (visible when mgr mode on) -->
      <div class="summary-fields-wrap" id="summary-fields-wrap" style="display:none">
        <span class="summary-fields-label">Show</span>
        <select class="summary-field-select" id="summary-field1" onchange="S.summaryField1=this.value;if(S.managerMode)renderChart()" title="First field in summary list">
          <option value="">Field 1…</option>
        </select>
        <span style="font-size:0.7rem;color:#7c3aed;font-weight:700">+</span>
        <select class="summary-field-select" id="summary-field2" onchange="S.summaryField2=this.value;if(S.managerMode)renderChart()" title="Second field in summary list">
          <option value="">Field 2…</option>
        </select>
      </div>
      <div class="tb-sep"></div>
      <input type="file" id="photo-folder-input" class="photo-folder-input" accept="image/*" multiple/>
      <div class="photo-btn" id="photo-btn" onclick="openPhotoFolder()">
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
      <div class="stat-item"><strong id="stat-roots">—</strong>&nbsp;shown roots</div>
      <div class="stat-item"><strong id="stat-vis">—</strong>&nbsp;visible cards</div>
      <div class="stat-item" id="stat-photos" style="display:none;color:var(--success)">📸 <strong id="stat-photos-val">0</strong> photos</div>
      <div class="stat-item" id="stat-mgr-mode" style="display:none;color:#7c3aed">👔 <strong id="stat-mgr-val">—</strong></div>
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

<!-- Reassign modal -->
<div class="modal-overlay hidden" id="reassign-modal">
  <div class="modal-box">
    <div class="modal-header">
      <div><div class="modal-title">Reassign Manager</div><div class="modal-sub" id="reassign-subject">Moving <strong>—</strong></div></div>
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
const S={
  rawRows:[],columns:[],colSamples:{},
  colMap:{empId:'',empName:'',managerId:''},
  cardSlots:{h1:'',h2:'',h3:'',b1:'',f1:'',f2:'',f3:''},
  cardAccent:'#4f46e5',
  empTypeCol:'',
  empTypeMap:{},
  empTypeLabels:{active:'',vacant:'',resigned:''},
  empTypeColors:{active:'#059669',vacant:'#dc2626',resigned:'#d97706'},
  filterCols:[],activeFilters:{},
  managerOverrides:{},removedIds:new Set(),
  viewData:[],childMap:{},descCount:{},nodeHeight:{},nodeDepth:{},
  zoom:1,highlighted:null,draggingField:null,
  reassignTarget:null,reassignPick:null,
  skipDepth:0,
  photoMap:{},photoObjUrls:[],
  photoSize:80,photoShape:'circle',photoPlacement:'top',
  managerMode:false,
  summaryField1:'',summaryField2:'',
};

function goTo(step){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));
  document.getElementById('screen-'+step)?.classList.add('active');
  const order=['upload','map','card','filter','chart'];
  const cur=order.indexOf(step);
  order.forEach((s,i)=>{
    const el=document.getElementById('nav-step-'+s);if(!el)return;
    el.className='step-item'+(i<cur?' done':i===cur?' active':'');
    const dot=el.querySelector('.step-dot');
    if(dot)dot.textContent=i<cur?'✓':String(i+1);
  });
  if(step==='chart'){setTimeout(()=>initPan(),80);setTimeout(()=>initSearch(),80);setTimeout(()=>populateSummaryFields(),120);}
}

/* ── FILE HANDLING ── */
function handleFile(file){
  const ext=file.name.split('.').pop().toLowerCase();
  if(ext==='csv'){Papa.parse(file,{header:true,skipEmptyLines:true,complete:r=>initData(r.data),error:e=>alert('CSV error: '+e.message)});}
  else if(['xlsx','xls'].includes(ext)){const reader=new FileReader();reader.onload=e=>{const wb=XLSX.read(e.target.result,{type:'array'});initData(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''}));};reader.readAsArrayBuffer(file);}
  else{alert('Please upload a CSV or Excel file.');}
}
function initData(rows){
  S.rawRows=rows;S.columns=rows.length?Object.keys(rows[0]):[];
  S.colSamples={};
  S.columns.forEach(col=>{S.colSamples[col]=[...new Set(rows.slice(0,25).map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='undefined'&&v!=='null'))].slice(0,3);});
  S.colMap=autoDetect(S.columns);buildMapScreen();goTo('map');
}
function autoDetect(cols){
  const lc=cols.map(c=>c.toLowerCase());
  const find=pats=>{for(const p of pats){const i=lc.findIndex(c=>c.includes(p));if(i>=0)return cols[i];}return '';};
  return{empId:find(['employee code','emp code','emp id','employee id','empcode','empid','staff id']),empName:find(['employee name','emp name','full name','person name','staff name','name']),managerId:find(['l1 manager code','l1 manager','manager code','manager id','reports to','supervisor','mgr code','mgrid'])};
}

/* ── PHOTOS ── */
async function openPhotoFolder(){
  if('showDirectoryPicker' in window){try{const d=await window.showDirectoryPicker({mode:'read'});await loadFromDirectoryHandle(d);}catch(e){if(e.name!=='AbortError')document.getElementById('photo-folder-input').click();}}
  else{document.getElementById('photo-folder-input').click();}
}
async function loadFromDirectoryHandle(dirHandle){
  S.photoObjUrls.forEach(u=>URL.revokeObjectURL(u));S.photoObjUrls=[];const newMap={};const IMG=new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);
  for await(const[name,handle] of dirHandle.entries()){if(handle.kind==='file'){const ext=name.split('.').pop().toLowerCase();if(IMG.has(ext)){const f=await handle.getFile();const k=name.replace(/\.[^.]+$/,'').toLowerCase().trim();const u=URL.createObjectURL(f);newMap[k]=u;S.photoObjUrls.push(u);}}}
  S.photoMap=newMap;updatePhotoUI();if(S.viewData.length)renderChart();
}
function loadFromFileInput(files){
  S.photoObjUrls.forEach(u=>URL.revokeObjectURL(u));S.photoObjUrls=[];const newMap={};const IMG=new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);
  Array.from(files).forEach(file=>{const ext=file.name.split('.').pop().toLowerCase();if(IMG.has(ext)){const k=file.name.replace(/\.[^.]+$/,'').toLowerCase().trim();const u=URL.createObjectURL(file);newMap[k]=u;S.photoObjUrls.push(u);}});
  S.photoMap=newMap;updatePhotoUI();if(S.viewData.length)renderChart();
}
function updatePhotoUI(){
  const count=Object.keys(S.photoMap).length;
  document.getElementById('photo-btn').classList.toggle('loaded',count>0);
  document.getElementById('photo-btn-label').textContent=count>0?'Photos':'Load Photos';
  const badge=document.getElementById('photo-count');badge.textContent=count;badge.style.display=count>0?'':'none';
  const stat=document.getElementById('stat-photos');if(stat){stat.style.display=count>0?'flex':'none';document.getElementById('stat-photos-val').textContent=count;}
}
function getPhotoUrl(node){
  if(!Object.keys(S.photoMap).length)return '';
  const id=node.id.toLowerCase().trim();if(S.photoMap[id])return S.photoMap[id];
  const nk=node.name.toLowerCase().trim().replace(/\s+/g,'_');if(S.photoMap[nk])return S.photoMap[nk];
  const nk2=node.name.toLowerCase().trim().replace(/\s+/g,'');if(S.photoMap[nk2])return S.photoMap[nk2];
  return '';
}

/* ── MAP SCREEN ── */
function buildMapScreen(){
  document.getElementById('col-count').textContent=S.columns.length;
  document.getElementById('detected-columns').innerHTML=S.columns.map(c=>`<div class="col-chip">📋 ${esc(c)}${S.colSamples[c].length?`<span class="chip-sample">${esc(S.colSamples[c].join(', '))}</span>`:''}</div>`).join('');
  const blank='<option value="">— select —</option>';
  const opts=blank+S.columns.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
  ['empId','empName','managerId'].forEach(k=>{const sel=document.getElementById('map-'+k);if(!sel)return;sel.innerHTML=opts;sel.value=S.colMap[k]||'';});
  const wrap=document.getElementById('data-preview-wrap');const preview=S.rawRows.slice(0,3);
  if(!preview.length){wrap.innerHTML='';return;}
  let html='<table class="data-preview-table"><thead><tr>'+S.columns.map(c=>`<th>${esc(c)}</th>`).join('')+'</tr></thead><tbody>';
  preview.forEach(row=>{html+='<tr>'+S.columns.map(c=>`<td title="${esc(String(row[c]||''))}">${esc(String(row[c]||'').substring(0,22))}</td>`).join('')+'</tr>';});
  wrap.innerHTML=html+'</tbody></table>';
}
function confirmColumnMap(){
  S.colMap.empId=document.getElementById('map-empId').value;
  S.colMap.empName=document.getElementById('map-empName').value;
  S.colMap.managerId=document.getElementById('map-managerId').value;
  if(!S.colMap.empId||!S.colMap.empName){alert('Please map Employee ID and Employee Name.');return;}
  buildCardScreen();goTo('card');
}

/* ── CARD DESIGN SCREEN ── */
const AUTO_FIELDS=[
  {id:'__auto_reports__',icon:'📊',label:'Direct Reports',desc:'Count of direct reports'},
  {id:'__auto_teamsize__',icon:'👥',label:'Total Team Size',desc:'All descendants count'},
];

function buildCardScreen(){
  const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId].filter(Boolean));
  const available=S.columns.filter(c=>!core.has(c));
  document.getElementById('card-fields-panel').innerHTML=
    '<div class="fields-section"><div class="fields-section-label">Column Fields</div>'+
    (available.length?available.map(f=>`<div class="field-chip" draggable="true" data-field="${esc(f)}" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)"><span class="drag-icon">⠿</span>${esc(f)}</div>`).join(''):'<div style="font-size:0.78rem;color:var(--text3);font-style:italic">No extra columns</div>')+
    '</div><div class="fields-section"><div class="fields-section-label">Auto-Calculated</div>'+
    AUTO_FIELDS.map(f=>`<div class="field-chip" draggable="true" data-field="${f.id}" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)" title="${f.desc}"><span class="drag-icon">⠿</span><span>${f.icon}</span>${f.label}</div>`).join('')+'</div>';

  if(!S.cardSlots.f3)S.cardSlots.f3='__auto_reports__';

  const COLORS=['#4f46e5','#7c3aed','#db2777','#dc2626','#d97706','#059669','#0891b2','#0284c7','#374151','#0f172a'];
  document.getElementById('color-palette').innerHTML=COLORS.map(c=>`<div class="color-swatch${S.cardAccent===c?' selected':''}" style="background:${c}" onclick="setCardAccent('${c}')"></div>`).join('');

  const empColSel=document.getElementById('emp-type-col');
  if(empColSel){
    empColSel.innerHTML='<option value="">Select column…</option>'+S.columns.filter(c=>!core.has(c)).map(c=>`<option value="${esc(c)}"${S.empTypeCol===c?' selected':''}>${esc(c)}</option>`).join('');
    if(S.empTypeCol)populateEmpTypeValues(S.empTypeCol);
  }
  renderCardPreview();syncChipStates();
}

function onEmpTypeColChange(){
  S.empTypeCol=document.getElementById('emp-type-col').value;
  if(S.empTypeCol){populateEmpTypeValues(S.empTypeCol);}
  else{document.getElementById('emp-type-rows').style.display='none';}
  if(S.viewData.length)renderChart();
}
function populateEmpTypeValues(col){
  const vals=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();
  const rows=document.getElementById('emp-type-rows');rows.style.display='';
  ['active','vacant','resigned'].forEach(key=>{
    const sel=document.getElementById('emp-val-'+key);
    sel.innerHTML='<option value="">Value…</option>'+vals.map(v=>`<option value="${esc(v)}"${S.empTypeLabels[key]===v?' selected':''}>${esc(v)}</option>`).join('');
    sel.onchange=()=>{S.empTypeLabels[key]=sel.value;buildEmpTypeMap();if(S.viewData.length)renderChart();};
    const colorInput=document.getElementById('emp-color-'+key);
    colorInput.value=S.empTypeColors[key];
    colorInput.oninput=()=>{S.empTypeColors[key]=colorInput.value;buildEmpTypeMap();if(S.viewData.length)renderChart();};
  });
  buildEmpTypeMap();
}
function buildEmpTypeMap(){
  S.empTypeMap={};
  ['active','vacant','resigned'].forEach(key=>{
    const v=S.empTypeLabels[key];if(v)S.empTypeMap[v]=S.empTypeColors[key];
  });
}
function getNodeBorderColor(node){
  if(S.empTypeCol&&S.empTypeMap){
    const val=String(node[S.empTypeCol]||'').trim();
    if(S.empTypeMap[val])return S.empTypeMap[val];
  }
  return S.cardAccent;
}

function onDragStart(e){S.draggingField=e.currentTarget.dataset.field;e.currentTarget.classList.add('dragging');e.dataTransfer.effectAllowed='move';}
function onDragEnd(e){e.currentTarget.classList.remove('dragging');}
function onZoneDragOver(e){e.preventDefault();e.currentTarget.classList.add('drop-target');}
function onZoneDragLeave(e){e.currentTarget.classList.remove('drop-target');}
function onZoneDrop(e,zone){e.preventDefault();e.currentTarget.classList.remove('drop-target');if(!S.draggingField)return;Object.keys(S.cardSlots).forEach(z=>{if(S.cardSlots[z]===S.draggingField)S.cardSlots[z]='';});S.cardSlots[zone]=S.draggingField;S.draggingField=null;renderCardPreview();syncChipStates();}
function clearZone(zone){S.cardSlots[zone]='';renderCardPreview();syncChipStates();}
function syncChipStates(){const placed=new Set(Object.values(S.cardSlots).filter(Boolean));document.querySelectorAll('.field-chip').forEach(c=>c.classList.toggle('placed',placed.has(c.dataset.field)));}
function fieldLabel(id){if(!id)return '';const af=AUTO_FIELDS.find(f=>f.id===id);if(af)return af.icon+' '+af.label;return id;}
function fieldSampleVal(id){if(!id)return '';if(id==='__auto_reports__')return '12';if(id==='__auto_teamsize__')return '48';const row=S.rawRows.find(r=>r[id])||S.rawRows[0]||{};return String(row[id]||'Sample').substring(0,18);}
function zoneHtml(zoneId,placeholder){
  const dA=`ondragover="onZoneDragOver(event)" ondragleave="onZoneDragLeave(event)" ondrop="onZoneDrop(event,'${zoneId}')"`;
  const v=S.cardSlots[zoneId];
  if(v)return`<div class="card-zone filled" ${dA}><span class="zone-field">${esc(fieldLabel(v))}</span><span class="zone-val">${esc(fieldSampleVal(v))}</span><span class="zone-remove" onclick="clearZone('${zoneId}')">✕</span></div>`;
  return`<div class="card-zone" ${dA}><span class="zone-ph">${placeholder}</span></div>`;
}

function renderCardPreview(){
  const sampleRow=S.rawRows.find(r=>r[S.colMap.empName])||S.rawRows[0]||{};
  const sampleName=String(sampleRow[S.colMap.empName]||'Employee Name').substring(0,26);
  const ac=S.cardAccent;
  const ps=S.photoSize,pr=getPhotoRadius();
  const photoDiv=`<div style="width:${ps}px;height:${ps}px;border-radius:${pr};background:linear-gradient(150deg,${ac}18,${ac}30);color:${ac};font-size:${Math.round(ps*0.28)}px;font-weight:800;display:flex;align-items:center;justify-content:center;border:3px solid ${ac}55;flex-shrink:0;box-shadow:0 6px 20px ${ac}44">AB</div>`;
  const b1ZoneHtml=`<div class="preview-b1-zone">${zoneHtml('b1','Body slot')}</div>`;
  const nameBlock=`<div style="width:100%;text-align:center"><div style="font-size:0.88rem;font-weight:800;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-bottom:4px">🔒 ${esc(sampleName)}</div>${b1ZoneHtml}</div>`;
  let bodyHtml;
  const pl=S.photoPlacement;
  if(pl==='none'){bodyHtml=`<div style="display:flex;flex-direction:column;gap:6px">${nameBlock}</div>`;}
  else if(pl==='top'){bodyHtml=`<div style="display:flex;flex-direction:column;align-items:center;gap:10px">${photoDiv}${nameBlock}</div>`;}
  else if(pl==='left'){bodyHtml=`<div style="display:flex;flex-direction:row;align-items:flex-start;gap:10px">${photoDiv}<div style="flex:1;min-width:0">${nameBlock}</div></div>`;}
  else{bodyHtml=`<div style="display:flex;flex-direction:row-reverse;align-items:flex-start;gap:10px">${photoDiv}<div style="flex:1;min-width:0">${nameBlock}</div></div>`;}

  document.getElementById('card-preview').innerHTML=
    `<div class="preview-card" style="border-top-color:${ac}">
      <div class="preview-card-header">${zoneHtml('h1','H1')}${zoneHtml('h2','H2')}${zoneHtml('h3','H3')}</div>
      <div class="preview-card-body">${bodyHtml}</div>
      <div class="preview-card-footer">${zoneHtml('f1','F1')}${zoneHtml('f2','F2')}${zoneHtml('f3','F3')}</div>
    </div>
    <div style="margin-top:10px;font-size:0.72rem;color:var(--text3)">Accent: <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:${ac};vertical-align:middle;margin-left:4px"></span> <strong style="color:${ac}">${ac}</strong></div>`;
}
function setCardAccent(color){S.cardAccent=color;document.querySelectorAll('.color-swatch').forEach(s=>s.classList.toggle('selected',s.style.background===color));renderCardPreview();}
function setPhotoShape(shape){S.photoShape=shape;document.querySelectorAll('.shape-btn').forEach(b=>b.classList.toggle('selected',b.dataset.shape===shape));renderCardPreview();if(S.viewData.length)renderChart();}
function setPhotoPlacement(p){S.photoPlacement=p;document.querySelectorAll('[data-placement]').forEach(b=>b.classList.toggle('selected',b.dataset.placement===p));renderCardPreview();if(S.viewData.length)renderChart();}
function getPhotoRadius(){if(S.photoShape==='circle')return'50%';if(S.photoShape==='rounded')return'12px';return'4px';}
function confirmCardDesign(){buildEmpTypeMap();buildFilterScreen();goTo('filter');}

/* ── FILTER SCREEN ── */
function buildFilterScreen(){
  const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId].filter(Boolean));
  const filterable=S.columns.filter(c=>!core.has(c));
  document.getElementById('filter-chip-picker').innerHTML=filterable.map(col=>`<div class="filter-chip ${S.filterCols.includes(col)?'selected':''}" data-col="${esc(col)}" onclick="toggleFilterCol('${esc(col)}')">${esc(col)}</div>`).join('');
  renderFilterPreview();
}
function toggleFilterCol(col){
  if(S.filterCols.includes(col))S.filterCols=S.filterCols.filter(c=>c!==col);
  else if(S.filterCols.length<3)S.filterCols.push(col);
  else{S.filterCols.shift();S.filterCols.push(col);}
  document.querySelectorAll('.filter-chip').forEach(c=>c.classList.toggle('selected',S.filterCols.includes(c.dataset.col)));
  renderFilterPreview();
}
function renderFilterPreview(){
  document.getElementById('filter-counter').textContent=`${S.filterCols.length} of 3 filters selected`;
  const area=document.getElementById('filter-preview-area');
  if(!S.filterCols.length){area.innerHTML=`<div style="font-size:0.82rem;color:var(--text3);padding:12px 0">No filters — full chart will display.</div>`;return;}
  area.innerHTML=`<div class="filter-preview-box">${S.filterCols.map((col,i)=>{const isLast=i===S.filterCols.length-1;const vals=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort().slice(0,10);return`<div class="fpr-row"><span class="fpr-col">${esc(col)}${isLast?` <span style="background:var(--accent);color:#fff;border-radius:999px;padding:1px 7px;font-size:0.58rem;font-weight:700;margin-left:4px">Export All</span>`:''}</span><div class="fpr-vals">${vals.map(v=>`<span class="fv-pill">${esc(v)}</span>`).join('')}${vals.length>=10?'<span style="font-size:0.7rem;color:var(--text3)">+ more</span>':''}</div></div>`;}).join('')}</div>`;
}
function launchChart(){S.activeFilters={};S.skipDepth=0;buildViewData();buildFilterBar();renderChart();goTo('chart');}

/* ── SUMMARY FIELDS ── */
function populateSummaryFields(){
  const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId].filter(Boolean));
  const opts='<option value="">—</option>'+'<option value="__name__">👤 Name</option>'+S.columns.filter(c=>!core.has(c)).map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
  const s1=document.getElementById('summary-field1');const s2=document.getElementById('summary-field2');
  if(s1){s1.innerHTML=opts;if(S.summaryField1)s1.value=S.summaryField1;}
  if(s2){s2.innerHTML=opts;if(S.summaryField2)s2.value=S.summaryField2;}
  document.getElementById('depth-select').value=S.skipDepth;
}

/* ── MANAGER MODE ── */
function toggleManagerMode(){
  S.managerMode=!S.managerMode;
  const btn=document.getElementById('mgr-mode-btn');
  btn.classList.toggle('active',S.managerMode);
  document.getElementById('summary-fields-wrap').style.display=S.managerMode?'flex':'none';
  const stat=document.getElementById('stat-mgr-mode');if(stat)stat.style.display=S.managerMode?'flex':'none';
  renderChart();
}

/* Is a node a manager (has ≥1 direct report in current viewData)? */
function isManager(nodeId){return (S.childMap[nodeId]||[]).length>0;}

/* ── VIEW DATA ── */
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
    matching.forEach(id=>{let cur=byId[id];const vis=new Set();while(cur&&cur.manager&&byId[cur.manager]&&!vis.has(cur.id)){vis.add(cur.id);keep.add(cur.manager);cur=byId[cur.manager];}});
    nodes=nodes.filter(n=>keep.has(n.id));
  }
  S.viewData=nodes;S.childMap={};
  nodes.forEach(n=>{if(!S.childMap[n.manager])S.childMap[n.manager]=[];S.childMap[n.manager].push(n);});
  S.descCount={};
  function calcD(id){if(S.descCount[id]!==undefined)return S.descCount[id];const kids=S.childMap[id]||[];S.descCount[id]=kids.reduce((s,k)=>s+1+calcD(k.id),0);return S.descCount[id];}
  nodes.filter(n=>!n.manager).forEach(r=>calcD(r.id));
  S.nodeHeight={};
  function calcH(id){if(S.nodeHeight[id]!==undefined)return S.nodeHeight[id];const kids=S.childMap[id]||[];S.nodeHeight[id]=kids.length?1+Math.max(...kids.map(k=>calcH(k.id))):0;return S.nodeHeight[id];}
  nodes.filter(n=>!n.manager).forEach(r=>calcH(r.id));
  nodes.forEach(n=>{if(S.nodeHeight[n.id]===undefined)calcH(n.id);});
  S.nodeDepth={};
  function calcDepth(id,d){S.nodeDepth[id]=d;(S.childMap[id]||[]).forEach(k=>calcDepth(k.id,d+1));}
  nodes.filter(n=>!n.manager).forEach(r=>calcDepth(r.id,0));
  nodes.forEach(n=>{if(S.nodeDepth[n.id]===undefined)S.nodeDepth[n.id]=0;});
}
function childrenOf(id){return S.childMap[id]||[];}
function countDescendants(id){return S.descCount[id]||0;}

/* ── FILTER BAR ── */
function buildFilterBar(){
  const bar=document.getElementById('filter-bar');
  if(!S.filterCols.length){bar.style.display='none';return;}
  bar.style.display='flex';
  const allVals={};
  S.filterCols.forEach(col=>{allVals[col]=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();});
  bar.innerHTML='<span style="font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);flex-shrink:0">Filters</span>'+
    S.filterCols.map(col=>`<div class="filter-dropdown-wrap"><span class="filter-dropdown-label">${esc(col)}</span><select class="filter-dropdown" onchange="applyFilter('${esc(col)}',this.value)"><option value="">All ${esc(col)}</option>${allVals[col].map(v=>`<option value="${esc(v)}"${(S.activeFilters[col]===v)?' selected':''}>${esc(v)}</option>`).join('')}</select></div>`).join('')+
    (Object.values(S.activeFilters).some(v=>v)?`<button class="btn btn-ghost btn-sm" onclick="clearAllFilters()" style="margin-left:auto">✕ Clear All</button>`:'');
}
function applyFilter(col,val){if(val)S.activeFilters[col]=val;else delete S.activeFilters[col];requestAnimationFrame(()=>setTimeout(()=>{buildViewData();renderChart();buildFilterBar();},0));}
function clearAllFilters(){S.activeFilters={};requestAnimationFrame(()=>setTimeout(()=>{buildViewData();renderChart();buildFilterBar();},0));}
function setSkipDepth(n){S.skipDepth=n;const ds=document.getElementById('depth-select');if(ds)ds.value=n;renderChart();}

/* ── SLOT VALUE HELPERS ── */
function getSlotVal(node,slot){
  const f=S.cardSlots[slot];if(!f)return '';
  if(f==='__auto_reports__')return childrenOf(node.id).length+' reports';
  if(f==='__auto_teamsize__')return countDescendants(node.id)+' people';
  return String(node[f]||'').substring(0,28);
}

/* ── RENDER CHART ── */
function renderChart(){
  const tree=document.getElementById('org-tree');tree.innerHTML='';
  const ds=document.getElementById('depth-select');if(ds)ds.value=S.skipDepth;
  let roots;
  if(S.skipDepth>0){roots=S.viewData.filter(n=>(S.nodeDepth[n.id]||0)===S.skipDepth);}
  else{roots=S.viewData.filter(n=>!n.manager);}
  if(!roots.length){tree.innerHTML=`<div class="no-data">No nodes found. Try a lower Skip Top value.</div>`;updateStats(roots);return;}
  const ul=document.createElement('ul');
  roots.forEach(r=>ul.appendChild(mkNodeLI(r,0)));
  tree.appendChild(ul);updateStats(roots);
  clearTimeout(window._fit);window._fit=setTimeout(()=>fitToScreen(true),180);
}

/* ── NODE LI ── */
function mkNodeLI(node,depth){
  depth=depth||0;
  const li=document.createElement('li');li.dataset.id=node.id;
  const borderColor=getNodeBorderColor(node);
  const ac=borderColor;
  const acLight=ac+'18',acMid=ac+'55';
  const kids=childrenOf(node.id);
  const card=document.createElement('div');
  card.className='node-card'+(node.id===S.highlighted?' highlighted':'');
  card.style.borderTopColor=ac;

  const h1=getSlotVal(node,'h1'),h2=getSlotVal(node,'h2'),h3=getSlotVal(node,'h3');
  const f1=getSlotVal(node,'f1'),f2=getSlotVal(node,'f2'),f3=getSlotVal(node,'f3')||node.id.substring(0,14);
  const b1=getSlotVal(node,'b1');
  const subtitle=h2;
  const ps=S.photoSize,pr=getPhotoRadius(),pfs=Math.round(ps*0.28)+'px';
  const photoInlineSize=`width:${ps}px;height:${ps}px;border-radius:${pr};`;
  const initials=node.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();
  const photoUrl=getPhotoUrl(node);
  let photoHtml='';
  if(photoUrl){
    photoHtml=`<img class="ncard-photo" src="${esc(photoUrl)}" crossorigin="anonymous" style="${photoInlineSize}border:3px solid ${acMid};box-shadow:0 8px 24px ${ac}66" onerror="this.onerror=null;this.style.display='none';this.nextElementSibling.style.display='flex'"><div class="ncard-photo-fallback" style="display:none;${photoInlineSize}font-size:${pfs};background:linear-gradient(150deg,${acLight},${ac}28);color:${ac};border:3px solid ${acMid};">${esc(initials)}</div>`;
  } else if(Object.keys(S.photoMap).length>0){
    photoHtml=`<div class="ncard-photo-fallback" style="display:flex;${photoInlineSize}font-size:${pfs};background:linear-gradient(150deg,${acLight},${ac}28);color:${ac};border:3px solid ${acMid};">${esc(initials)}</div>`;
  }

  const b1row=b1?`<div class="ncard-body-b1" title="${esc(b1)}">${esc(b1)}</div>`:'';
  const textBlock=`<div class="ncard-text-wrap"><div class="ncard-name" title="${esc(node.name)}">${esc(node.name)}</div>${subtitle?`<div class="ncard-sub" title="${esc(subtitle)}">${esc(subtitle)}</div>`:''}${b1row}</div>`;
  let bodyHtml;
  const pl=S.photoPlacement;
  if(!photoHtml||pl==='none'){bodyHtml=`<div class="ncard-body-inner" style="flex-direction:column">${textBlock}</div>`;}
  else if(pl==='top'){bodyHtml=`<div class="ncard-body-inner" style="flex-direction:column;align-items:center"><div style="flex-shrink:0">${photoHtml}</div>${textBlock}</div>`;}
  else if(pl==='left'){bodyHtml=`<div class="ncard-body-inner" style="flex-direction:row;align-items:flex-start"><div style="flex-shrink:0">${photoHtml}</div><div style="flex:1;min-width:0">${textBlock}</div></div>`;}
  else{bodyHtml=`<div class="ncard-body-inner" style="flex-direction:row-reverse;align-items:flex-start"><div style="flex-shrink:0">${photoHtml}</div><div style="flex:1;min-width:0">${textBlock}</div></div>`;}

  card.innerHTML=
    `<div class="ncard-header" style="background:${acLight};border-bottom-color:${ac}33">
      <span class="ncard-slot${h1?' has-val':''}" title="${esc(h1)}">${esc(h1)||'—'}</span>
      <span class="ncard-slot${h2?' has-val':''}" title="${esc(h2)}">${esc(h2)||'—'}</span>
      <span class="ncard-slot${h3?' has-val':''}" title="${esc(h3)}">${esc(h3)||'—'}</span>
    </div>
    <div class="ncard-body">${bodyHtml}</div>
    <div class="ncard-footer" style="background:${acLight};border-top-color:${ac}33">
      <span class="ncard-slot${f1?' has-val':''}" title="${esc(f1)}">${esc(f1)||'—'}</span>
      <span class="ncard-slot${f2?' has-val':''}" title="${esc(f2)}">${esc(f2)||'—'}</span>
      <span class="ncard-slot${f3?' has-val':''}" title="${esc(f3)}">${esc(f3)||node.id.substring(0,14)}</span>
    </div>
    <div class="ncard-export-btn" title="Export this person's team" onclick="exportSubtree(event,'${esc(node.id)}')" style="position:absolute;top:6px;right:30px;width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:6px;font-size:0.6rem;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 0.15s;z-index:8">📸</div>
    <div class="ncard-edit-btn" title="Reassign manager" onclick="openReassignModal(event,'${esc(node.id)}')" style="position:absolute;top:6px;right:6px;width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:6px;font-size:0.65rem;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 0.15s;z-index:8">✎</div>`;
  card.querySelectorAll('.ncard-edit-btn,.ncard-export-btn').forEach(b=>{card.addEventListener('mouseenter',()=>b.style.opacity='1');card.addEventListener('mouseleave',()=>b.style.opacity='0');});

  if(kids.length){
    const cb=document.createElement('div');cb.className='collapse-btn';cb.innerHTML='▾';cb.title='Collapse / expand';
    cb.addEventListener('click',e=>{e.stopPropagation();toggleCollapse(li,cb);});
    card.appendChild(cb);
  }
  li.appendChild(card);

  if(kids.length){
    /*
     * MANAGER VIEW SUMMARISATION RULES:
     *
     * The node card itself is ALWAYS a full card — never summarised.
     * Among a node's children:
     *   - Manager children (≥1 direct report) → ALWAYS rendered as full cards in the tree
     *   - Leaf children (0 direct reports)     → shown in a compact static summary list
     *
     * Example A → B(mgr), C(mgr), D(mgr), E(leaf):
     *   B, C, D → full cards in tree  ✓
     *   E       → appears in compact list attached below A  ✓
     *
     * Example B → X(mgr), Y(leaf), Z(leaf):
     *   X       → full card in tree  ✓
     *   Y, Z    → appear in compact list attached below B  ✓
     *
     * The compact list is static — no click / no drill-down.
     */
    if(S.managerMode){
      const managerKids=kids.filter(k=>isManager(k.id));  // have reportees → full cards
      const leafKids=kids.filter(k=>!isManager(k.id));    // no reportees  → compact list

      const ul=document.createElement('ul');

      // Manager children always get full tree cards
      managerKids.forEach(k=>ul.appendChild(mkNodeLI(k,depth+1)));

      // Leaf children are shown in a compact static list (no click)
      if(leafKids.length>0){
        ul.appendChild(mkLeafSummaryLI(leafKids,ac));
      }

      li.appendChild(ul);
    } else {
      const ul=document.createElement('ul');
      kids.forEach(k=>ul.appendChild(mkNodeLI(k,depth+1)));
      li.appendChild(ul);
    }
  }
  return li;
}

/* ── LEAF SUMMARY LIST NODE ──
 *
 * Renders a compact static card listing leaf employees (no direct reports).
 * Used only in Manager View. Static — no click, no drill-down.
 *
 * CRITICAL EXPORT FIX:
 * Built entirely with createElement + direct style assignment (zero innerHTML).
 * html2canvas reliably renders DOM nodes built this way; innerHTML-based nodes
 * in off-screen stages often paint blank even with correct styles.
 */
function mkLeafSummaryLI(leafNodes,ac){
  const li=document.createElement('li');
  const f1=S.summaryField1,f2=S.summaryField2;
  const count=leafNodes.length;
  const FONT="'Plus Jakarta Sans',sans-serif";

  /* ── card wrapper ── */
  const card=document.createElement('div');
  card.className='summary-list-card'; // triggers inlineStyles skip — preserves all inline styles
  Object.assign(card.style,{
    display:'inline-block',minWidth:'200px',maxWidth:'280px',
    background:'#ffffff',border:'1.5px solid #e2e8f0',
    borderTop:'3px solid #7c3aed',borderRadius:'14px',
    boxShadow:'0 1px 4px rgba(0,0,0,0.07)',
    fontFamily:FONT,overflow:'visible',verticalAlign:'top',
  });

  /* ── header ── */
  const header=document.createElement('div');
  Object.assign(header.style,{
    padding:'7px 12px',background:'#f5f3ff',
    borderBottom:'1px solid #e9d5ff',
    borderRadius:'12px 12px 0 0',
    display:'flex',justifyContent:'space-between',alignItems:'center',
    fontFamily:FONT,
  });
  const titleSpan=document.createElement('span');
  Object.assign(titleSpan.style,{fontSize:'0.68rem',fontWeight:'800',color:'#7c3aed',textTransform:'uppercase',letterSpacing:'0.05em',fontFamily:FONT});
  titleSpan.textContent='ICs ('+count+')';
  const countBadge=document.createElement('span');
  Object.assign(countBadge.style,{fontSize:'0.68rem',fontWeight:'800',background:'#7c3aed',color:'#ffffff',borderRadius:'999px',padding:'1px 8px',fontFamily:FONT});
  countBadge.textContent=String(count);
  header.appendChild(titleSpan);header.appendChild(countBadge);
  card.appendChild(header);

  /* ── body ── */
  const body=document.createElement('div');
  Object.assign(body.style,{padding:'4px 0',overflow:'visible',fontFamily:FONT});

  leafNodes.forEach((n,idx)=>{
    const initials=n.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();
    const borderC=getNodeBorderColor(n);
    const photoUrl=getPhotoUrl(n);

    /* row */
    const row=document.createElement('div');
    Object.assign(row.style,{
      display:'flex',alignItems:'center',gap:'8px',
      padding:'6px 12px',fontFamily:FONT,
      borderBottom: idx<leafNodes.length-1 ? '1px solid #e2e8f0' : 'none',
    });
    row.title=n.name;

    /* avatar */
    const fallback=document.createElement('div');
    Object.assign(fallback.style,{
      display:'flex',alignItems:'center',justifyContent:'center',
      width:'26px',height:'26px',minWidth:'26px',
      borderRadius:'8px',fontSize:'0.6rem',fontWeight:'800',
      fontFamily:FONT,background:borderC+'18',color:borderC,
      border:'2px solid '+borderC+'44',flexShrink:'0',
    });
    fallback.textContent=initials;

    if(photoUrl){
      const img=document.createElement('img');
      Object.assign(img.style,{width:'26px',height:'26px',minWidth:'26px',borderRadius:'8px',objectFit:'cover',objectPosition:'center top',border:'2px solid '+borderC+'55',flexShrink:'0',display:'block'});
      img.src=photoUrl;img.crossOrigin='anonymous';
      img.onerror=function(){this.style.display='none';fallback.style.display='flex';};
      fallback.style.display='none';
      row.appendChild(img);row.appendChild(fallback);
    } else {
      row.appendChild(fallback);
    }

    /* text block */
    const nameVal=n.name.substring(0,24);
    const f1IsName=(f1==='__name__');
    const primaryVal=f1?(f1IsName?nameVal:(String(n[f1]||'').trim()||nameVal).substring(0,24)):nameVal;
    const showNameSub=f1&&!f1IsName&&primaryVal!==nameVal;
    const val2=f2?(f2==='__name__'?n.name.substring(0,22):String(n[f2]||'').substring(0,22)):'';

    const textWrap=document.createElement('div');
    Object.assign(textWrap.style,{flex:'1',minWidth:'0',overflow:'hidden',fontFamily:FONT});

    const nameLine=document.createElement('div');
    Object.assign(nameLine.style,{fontSize:'0.75rem',fontWeight:'700',color:'#0f172a',whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis',fontFamily:FONT,maxWidth:'210px'});
    nameLine.textContent=primaryVal;
    textWrap.appendChild(nameLine);

    if(showNameSub){
      const subLine=document.createElement('div');
      Object.assign(subLine.style,{fontSize:'0.65rem',color:'#475569',fontWeight:'600',fontFamily:FONT,whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis',maxWidth:'210px'});
      subLine.textContent=nameVal;
      textWrap.appendChild(subLine);
    }
    if(val2){
      const val2Line=document.createElement('div');
      Object.assign(val2Line.style,{fontSize:'0.68rem',color:'#94a3b8',fontFamily:FONT,whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis',maxWidth:'210px'});
      val2Line.textContent=val2;
      textWrap.appendChild(val2Line);
    }

    row.appendChild(textWrap);
    body.appendChild(row);
  });

  card.appendChild(body);
  li.appendChild(card);
  return li;
}



/* ── COLLAPSE / EXPAND ── */
function toggleCollapse(li,btn){
  li.classList.toggle('collapsed');const c=li.classList.contains('collapsed');
  const childUl=li.querySelector(':scope > ul');if(childUl)childUl.style.display=c?'none':'';
  btn.innerHTML=c?'▸':'▾';btn.style.color=c?'var(--warning)':'';
  li.querySelector('.node-card')?.classList.toggle('collapsed-node',c);
  setTimeout(()=>updateStats(),60);
}
function expandAll(){document.querySelectorAll('li.collapsed').forEach(li=>{li.classList.remove('collapsed');const u=li.querySelector(':scope > ul');if(u)u.style.display='';li.querySelector('.node-card')?.classList.remove('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▾';b.style.color='';}});setTimeout(()=>updateStats(),60);}
function collapseAll(){document.querySelectorAll('li').forEach(li=>{if(!li.parentElement?.parentElement?.closest('li'))return;if(li.querySelector(':scope > ul')){li.classList.add('collapsed');const u=li.querySelector(':scope > ul');if(u)u.style.display='none';li.querySelector('.node-card')?.classList.add('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▸';b.style.color='var(--warning)';};}});setTimeout(()=>updateStats(),60);}

function updateStats(roots){
  if(!roots)roots=S.skipDepth>0?S.viewData.filter(n=>(S.nodeDepth[n.id]||0)===S.skipDepth):S.viewData.filter(n=>!n.manager);
  document.getElementById('stat-total').textContent=S.viewData.length;
  document.getElementById('stat-roots').textContent=roots.length;
  document.getElementById('stat-vis').textContent=document.querySelectorAll('.node-card:not(.summary-list-card)').length;
  document.getElementById('stat-filtered').style.display=Object.values(S.activeFilters).some(v=>v)?'flex':'none';
  const mgrStat=document.getElementById('stat-mgr-mode');const mgrVal=document.getElementById('stat-mgr-val');
  if(mgrStat){mgrStat.style.display=S.managerMode?'flex':'none';if(S.managerMode&&mgrVal){const leafCount=S.viewData.filter(n=>!isManager(n.id)).length;mgrVal.textContent=leafCount+' ICs in lists';}}
}

/* ── ZOOM / PAN ── */
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
  wrap.onmousedown=e=>{if(e.target.closest('.node-card,.summary-list-card,.collapse-btn'))return;_panning=true;_px=e.clientX;_py=e.clientY;_psl=wrap.scrollLeft;_pst=wrap.scrollTop;wrap.style.cursor='grabbing';};
  window.onmousemove=e=>{if(!_panning)return;cwrap().scrollLeft=_psl-(e.clientX-_px);cwrap().scrollTop=_pst-(e.clientY-_py);};
  window.onmouseup=()=>{_panning=false;if(cwrap())cwrap().style.cursor='';};
  wrap.addEventListener('wheel',e=>{if(e.ctrlKey||e.metaKey){e.preventDefault();zoomBy(e.deltaY<0?0.08:-0.08);}},{passive:false});
}

/* ── SEARCH ── */
function initSearch(){
  const input=document.getElementById('chart-search');const box=document.getElementById('chart-search-results');if(!input)return;
  function positionBox(){const r=input.getBoundingClientRect();box.style.top=(r.bottom+4)+'px';box.style.left=r.left+'px';box.style.width=Math.max(270,r.width)+'px';}
  input.addEventListener('input',function(){
    const q=this.value.trim().toLowerCase();if(!q){box.classList.remove('visible');return;}
    const hits=S.viewData.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)).slice(0,10);
    box.innerHTML=hits.length?hits.map(n=>`<div class="sr-item" onclick="highlightNode('${esc(n.id)}')"><div class="sr-name">${esc(n.name)}</div><div class="sr-sub">${esc(n.id)}</div></div>`).join(''):'<div class="sr-item" style="color:var(--text3);font-size:0.8rem;padding:12px 13px">No results</div>';
    positionBox();box.classList.add('visible');
  });
  input.addEventListener('focus',()=>{if(input.value.trim())positionBox();});
  document.addEventListener('click',e=>{if(!e.target.closest('.search-wrap'))box.classList.remove('visible');});
  window.addEventListener('resize',()=>{if(box.classList.contains('visible'))positionBox();});
}
function highlightNode(id){
  document.querySelectorAll('.node-card.highlighted').forEach(c=>c.classList.remove('highlighted'));
  S.highlighted=id;expandAll();
  const li=document.querySelector(`li[data-id="${CSS.escape(id)}"]`);
  if(li){const card=li.querySelector('.node-card');if(card){card.classList.add('highlighted');setTimeout(()=>{const r=card.getBoundingClientRect();const w=cwrap();const wr=w.getBoundingClientRect();w.scrollTo({left:w.scrollLeft+(r.left-wr.left)-wr.width/2+r.width/2,top:w.scrollTop+(r.top-wr.top)-wr.height/2+r.height/2,behavior:'smooth'});},80);}}
  document.getElementById('chart-search').value='';document.getElementById('chart-search-results').classList.remove('visible');
}

/* ── EXPORT HELPERS ── */
function inlineStyles(root){
  const PROPS=['color','backgroundColor','borderTopColor','borderBottomColor','borderLeftColor','borderRightColor','borderTopWidth','borderTopStyle','borderRadius','fontFamily','fontSize','fontWeight','fontStyle','lineHeight','padding','paddingTop','paddingBottom','paddingLeft','paddingRight','margin','display','flexDirection','justifyContent','alignItems','gap','whiteSpace','overflow','textOverflow','opacity','boxShadow','borderWidth','borderStyle','borderColor'];
  root.querySelectorAll('*').forEach(el=>{
    // Skip summary list elements — they are 100% inline-styled already.
    // Overwriting their computed styles would destroy the resolved colours/layout.
    if(el.closest('.summary-list-card')){return;}

    const cs=window.getComputedStyle(el);
    PROPS.forEach(p=>{try{const v=cs[p];if(v)el.style[p]=v;}catch(e){}});
    // Open up hidden/auto overflow so html2canvas captures everything.
    // Exceptions: node-card, ncard-name, ncard-sub need overflow:hidden for text clipping.
    const ov=el.style.overflow;
    const ovY=cs.overflowY;
    const isTextClipper=el.classList.contains('node-card')||el.classList.contains('ncard-name')||el.classList.contains('ncard-sub');
    if(!isTextClipper&&(ov==='hidden'||ovY==='auto'||ovY==='scroll')){
      el.style.overflow='visible';
      el.style.overflowY='visible';
      el.style.overflowX='visible';
    }
    el.classList.remove('collapsed');
  });
}

async function buildRenderStage(rootNodeId){
  const PAD=20;
  const stage=document.createElement('div');
  // Must be on-screen (so browser paints it) but behind the export overlay (z:9999).
  // No opacity/visibility hiding — those prevent html2canvas from capturing content.
  // The overlay is always appended BEFORE this is called, so it covers the stage visually.
  stage.style.cssText=`position:fixed;top:0;left:0;z-index:9998;background:#f8fafc;padding:${PAD}px;display:inline-block;white-space:nowrap;pointer-events:none;font-family:'Plus Jakarta Sans',sans-serif`;
  let sourceTree;
  if(rootNodeId){
    sourceTree=document.createElement('div');sourceTree.className='org-tree';
    const ul=document.createElement('ul');const node=S.viewData.find(n=>n.id===rootNodeId);
    if(node)ul.appendChild(mkNodeLI(node,0));sourceTree.appendChild(ul);
  } else {
    sourceTree=document.createElement('div');sourceTree.className='org-tree';
    const exportRoots=S.skipDepth>0?S.viewData.filter(n=>(S.nodeDepth[n.id]||0)===S.skipDepth):S.viewData.filter(n=>!n.manager);
    const exportUl=document.createElement('ul');
    exportRoots.forEach(r=>exportUl.appendChild(mkNodeLI(r,0)));sourceTree.appendChild(exportUl);
  }
  sourceTree.querySelectorAll('li').forEach(li=>li.classList.remove('collapsed'));
  sourceTree.querySelectorAll('ul').forEach(ul=>{ul.style.display='';});
  sourceTree.querySelectorAll('.collapse-btn,.ncard-edit-btn,.ncard-export-btn').forEach(b=>b.remove());
  stage.appendChild(sourceTree);
  document.body.appendChild(stage);
  // Give browser time to fully lay out and paint
  await new Promise(r=>setTimeout(r,400));
  if(document.fonts?.ready)await document.fonts.ready;
  await new Promise(r=>setTimeout(r,150));
  inlineStyles(stage);
  await new Promise(r=>setTimeout(r,100));
  return stage;
}
async function renderToCanvas(stage){
  const W=Math.ceil(stage.scrollWidth),H=Math.ceil(stage.scrollHeight);
  return html2canvas(stage,{
    backgroundColor:'#f8fafc',scale:3,useCORS:true,logging:false,
    allowTaint:true,foreignObjectRendering:false,
    width:W,height:H,scrollX:0,scrollY:0,
    windowWidth:Math.max(W+200,window.innerWidth),
    windowHeight:Math.max(H+200,window.innerHeight),
    x:0,y:0,
  });
}
function triggerDownload(blob,fname){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=fname;a.click();URL.revokeObjectURL(url);}
function csvEsc(v){return'"'+String(v??'').replace(/"/g,'""')+'"';}
function buildCSVContent(){
  const cols=[S.colMap.empId,S.colMap.empName,S.colMap.managerId,...S.columns.filter(c=>c!==S.colMap.empId&&c!==S.colMap.empName&&c!==S.colMap.managerId)].filter(Boolean);
  return cols.map(csvEsc).join(',')+'\n'+S.viewData.map(n=>cols.map(c=>csvEsc(n[c]||'')).join(',')).join('\n');
}
function downloadCSV(){triggerDownload(new Blob([buildCSVContent()],{type:'text/csv;charset=utf-8;'}),'orgchart_export.csv');}
function makeOverlay(title,sub){
  const o=document.createElement('div');o.className='export-overlay';
  o.innerHTML=`<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px">${title}</div><div style="font-size:0.75rem;color:#94a3b8;margin-top:4px">${sub}</div>`;
  return o;
}

async function exportPNG(){
  const overlay=makeOverlay('Rendering org chart…','Capturing at 3× resolution');document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));let stage;
  try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');const fp=Object.values(S.activeFilters).filter(Boolean).map(v=>v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');const mode=S.managerMode?'_mgr_view':'';await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,`orgchart_${fp?fp+'_':''}${mode}${stamp}.png`);res();},'image/png'));}
  catch(e){alert('PNG export failed: '+e.message);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);}
}

async function exportSubtree(e,nodeId){
  e.stopPropagation();const node=S.viewData.find(n=>n.id===nodeId);if(!node)return;
  const includeIds=new Set([nodeId]);function collectDesc(id){(S.childMap[id]||[]).forEach(k=>{includeIds.add(k.id);collectDesc(k.id);});}collectDesc(nodeId);
  const overlay=makeOverlay(`Exporting ${node.name}'s team (${includeIds.size})…`,'');document.body.appendChild(overlay);
  const savedViewData=S.viewData,savedChildMap=S.childMap,savedDescCount=S.descCount,savedNodeHeight=S.nodeHeight,savedNodeDepth=S.nodeDepth;
  const hadOverride=S.managerOverrides.hasOwnProperty(nodeId);const prevOverride=S.managerOverrides[nodeId];
  S.viewData=savedViewData.filter(n=>includeIds.has(n.id));S.managerOverrides[nodeId]='';
  S.childMap={};S.viewData.forEach(n=>{const mgr=(n.id===nodeId)?'':n.manager;if(!S.childMap[mgr])S.childMap[mgr]=[];S.childMap[mgr].push(n);});
  S.descCount={};S.nodeHeight={};S.nodeDepth={};
  function cD(id){const k=S.childMap[id]||[];S.descCount[id]=k.reduce((s,c)=>s+1+cD(c.id),0);return S.descCount[id];}
  function cH(id){const k=S.childMap[id]||[];S.nodeHeight[id]=k.length?1+Math.max(...k.map(c=>cH(c.id))):0;return S.nodeHeight[id];}
  function cDep(id,d){S.nodeDepth[id]=d;(S.childMap[id]||[]).forEach(k=>cDep(k.id,d+1));}
  cD(nodeId);cH(nodeId);cDep(nodeId,0);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,120));let stage;
  try{stage=await buildRenderStage(nodeId);const canvas=await renderToCanvas(stage);const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');const safeName=node.name.replace(/[^a-zA-Z0-9]/g,'_');await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,`team_${safeName}_${stamp}.png`);res();},'image/png'));}
  catch(ex){alert('Subtree export failed: '+ex.message);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);if(hadOverride)S.managerOverrides[nodeId]=prevOverride;else delete S.managerOverrides[nodeId];S.viewData=savedViewData;S.childMap=savedChildMap;S.descCount=savedDescCount;S.nodeHeight=savedNodeHeight;S.nodeDepth=savedNodeDepth;renderChart();}
}

/* ── PPTX ── */
function xe(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
const SW=12192000,SH=6858000;
function pptxRect(id,x,y,cx,cy,fill){return`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="r${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;}
function pptxTxt(id,x,y,cx,cy,text,sz,bold,color,algn){algn=algn||'ctr';return`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="${algn}"/><a:r><a:rPr lang="en-US" sz="${sz}" b="${bold?1:0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${xe(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;}
function pptxImg(id,x,y,cx,cy,rId){return`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="img${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;}
function pptxSlide(bg,content,rels){return[`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/><a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>${content}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,rels];}

async function buildPPTXBlob(imgB64,cW,cH,titleSuffix){
  titleSuffix=titleSuffix||'';const ac=S.cardAccent.replace('#','');
  const stamp=new Date().toLocaleDateString('en-IN',{day:'numeric',month:'long',year:'numeric'});
  const activeF=Object.entries(S.activeFilters).filter(([,v])=>v);
  const filterLine=activeF.map(([k,v])=>`${k}: ${v}`).join('  |  ')||(titleSuffix||'All Employees');
  const roots=S.viewData.filter(n=>!n.manager).length;
  const mgrCount=S.viewData.filter(n=>isManager(n.id)).length;
  const modeNote=S.managerMode?` | Manager View (ICs in lists)`:'';
  const summaryNote=S.managerMode&&(S.summaryField1||S.summaryField2)?` | IC fields: ${[S.summaryField1,S.summaryField2].filter(Boolean).map(f=>f==='__name__'?'Name':f).join(' + ')}`:'';
  const [s1xml,s1rels]=pptxSlide('F1F5F9',
    pptxRect(2,0,0,SW,Math.round(SH*0.52),ac)+pptxRect(3,0,Math.round(SH*0.52),SW,Math.round(SH*0.48),'FFFFFF')+
    pptxTxt(4,Math.round(SW*0.08),Math.round(SH*0.12),Math.round(SW*0.84),Math.round(SH*0.22),'Org Chart',7600,true,'FFFFFF','l')+
    pptxTxt(5,Math.round(SW*0.08),Math.round(SH*0.35),Math.round(SW*0.84),420000,filterLine+modeNote,2200,true,'FFFFFF','l')+
    pptxTxt(6,Math.round(SW*0.08),Math.round(SH*0.44),Math.round(SW*0.84),340000,'Generated: '+stamp+summaryNote,1500,false,'C7D2FE','l')+
    pptxTxt(7,Math.round(SW*0.08),Math.round(SH*0.59),Math.round(SW*0.38),400000,String(S.viewData.length),5200,true,ac,'l')+
    pptxTxt(8,Math.round(SW*0.08),Math.round(SH*0.74),Math.round(SW*0.38),310000,'Total Employees',1600,false,'64748B','l')+
    pptxTxt(9,Math.round(SW*0.55),Math.round(SH*0.59),Math.round(SW*0.35),400000,String(S.viewData.filter(n=>isManager(n.id)).length),4000,true,'64748B','l')+
    pptxTxt(10,Math.round(SW*0.55),Math.round(SH*0.74),Math.round(SW*0.35),310000,'Managers',1600,false,'64748B','l'),
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`);
  const imgAspect=cW/cH,slideAspect=SW/SH;
  let iW,iH,iX,iY;
  if(imgAspect>=slideAspect){iW=SW;iH=Math.round(SW/imgAspect);iX=0;iY=Math.round((SH-iH)/2);}
  else{iH=SH;iW=Math.round(SH*imgAspect);iX=Math.round((SW-iW)/2);iY=0;}
  const capY=SH-Math.round(SH*0.065);
  const [s2xml,]=pptxSlide('FFFFFF',
    pptxImg(20,iX,iY,iW,iH,'rId2')+pptxTxt(21,Math.round(SW*0.04),capY,Math.round(SW*0.92),Math.round(SH*0.055),filterLine+modeNote+' · '+stamp+' · '+S.viewData.length+' employees',1000,false,'64748B','r'),
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`);
  const statItems=[{label:'Total Employees',val:S.viewData.length,color:ac},{label:'Managers',val:mgrCount,color:'7c3aed'},{label:'Roots',val:roots,color:'0891b2'},{label:'Mode',val:S.managerMode?'Manager View':'Full Tree',color:'059669'}];
  const boxW=Math.round(SW*0.19),boxH=Math.round(SH*0.3),gap=Math.round((SW-boxW*4)*0.2);
  const totalBW=boxW*4+gap*3,bStartX=Math.round((SW-totalBW)/2),bY=Math.round(SH*0.3);
  let sc=pptxRect(2,0,0,SW,Math.round(SH*0.2),ac)+pptxTxt(3,Math.round(SW*0.04),0,Math.round(SW*0.6),Math.round(SH*0.2),'Summary Dashboard',1800,true,'FFFFFF','l')+pptxTxt(4,Math.round(SW*0.65),0,Math.round(SW*0.3),Math.round(SH*0.2),stamp,1200,false,'C7D2FE','r')+pptxTxt(5,Math.round(SW*0.04),Math.round(SH*0.22),Math.round(SW*0.92),Math.round(SH*0.06),filterLine+modeNote,1600,false,'64748B','l');
  statItems.forEach((st,i)=>{const bx=bStartX+i*(boxW+gap);sc+=pptxRect(10+i*2,bx,bY,boxW,boxH,'F8FAFC')+pptxRect(11+i*2,bx,bY,boxW,Math.round(boxH*0.05),st.color)+pptxTxt(20+i*2,bx,bY+Math.round(boxH*0.1),boxW,Math.round(boxH*0.52),String(st.val),4800,true,st.color)+pptxTxt(21+i*2,bx,bY+Math.round(boxH*0.7),boxW,Math.round(boxH*0.28),st.label,1300,false,'64748B');});
  const [s3xml,]=pptxSlide('FFFFFF',sc,`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`);
  const BP={
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
  Object.entries(BP).forEach(([k,v])=>zip.file(k,v));
  zip.file('ppt/slides/slide1.xml',s1xml);zip.file('ppt/slides/_rels/slide1.xml.rels',mkRels(s1rels));
  zip.file('ppt/slides/slide2.xml',s2xml);zip.file('ppt/slides/_rels/slide2.xml.rels',mkRels(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`));
  zip.file('ppt/slides/slide3.xml',s3xml);zip.file('ppt/slides/_rels/slide3.xml.rels',mkRels(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`));
  zip.file('ppt/media/image1.png',imgB64,{base64:true});
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'DEFLATE'});
}

async function exportPPTX(){
  if(typeof JSZip==='undefined'){alert('ZIP library failed to load.');return;}
  const overlay=makeOverlay('Building PowerPoint…','Rendering then packaging');document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));let stage;
  try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const blob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height);const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');const fp=Object.values(S.activeFilters).filter(Boolean).map(v=>v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');const mode=S.managerMode?'_mgr':'';triggerDownload(blob,`orgchart_${fp?fp+'_':''}${mode}${dp}.pptx`);}
  catch(e){alert('PPTX failed: '+e.message);console.error(e);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);}
}

async function exportAll(){
  if(typeof JSZip==='undefined'){alert('ZIP library failed to load.');return;}
  const lastFilterCol=S.filterCols[S.filterCols.length-1]||null;
  if(!lastFilterCol){await _exportAllSingleView();return;}
  const parentFilters=Object.entries(S.activeFilters).filter(([k])=>k!==lastFilterCol);
  const relevantRows=S.rawRows.filter(row=>parentFilters.every(([col,val])=>!val||String(row[col]||'').trim()===val));
  const lastVals=[...new Set(relevantRows.map(r=>String(r[lastFilterCol]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();
  if(!lastVals.length){await _exportAllSingleView();return;}
  const overlay=document.createElement('div');overlay.className='export-overlay';
  overlay.innerHTML=`<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px" id="_ea_title">Exporting ${lastVals.length} charts…</div><div id="_ea_step" style="font-size:0.8rem;color:#64748b;margin-top:4px">Preparing…</div><div id="_ea_prog" style="font-size:0.7rem;color:var(--text3);margin-top:2px">0 / ${lastVals.length}</div>`;
  document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));
  const savedFilters={...S.activeFilters};const outerZip=new JSZip();let successCount=0;
  try{
    for(let i=0;i<lastVals.length;i++){
      const val=lastVals[i];const safeName=val.replace(/[^a-zA-Z0-9]/g,'_');
      document.getElementById('_ea_step').textContent=`📊 ${val}`;
      document.getElementById('_ea_prog').textContent=`${i+1} / ${lastVals.length}`;
      S.activeFilters[lastFilterCol]=val;buildViewData();renderChart();await new Promise(r=>setTimeout(r,320));
      outerZip.file(`${safeName}/${safeName}.csv`,buildCSVContent());
      let stage2;
      try{stage2=await buildRenderStage();const canvas2=await renderToCanvas(stage2);outerZip.file(`${safeName}/${safeName}.png`,canvas2.toDataURL('image/png').split(',')[1],{base64:true});const pptxBlob=await buildPPTXBlob(canvas2.toDataURL('image/png').split(',')[1],canvas2.width,canvas2.height,val);outerZip.file(`${safeName}/${safeName}.pptx`,pptxBlob);successCount++;}
      finally{if(stage2)stage2.remove();}
    }
  }finally{S.activeFilters=savedFilters;buildViewData();renderChart();buildFilterBar();overlay.remove();applyZoom(savedZoom);}
  if(successCount>0){const zipBlob=await outerZip.generateAsync({type:'blob',compression:'DEFLATE'});const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');triggerDownload(zipBlob,`orgcharts_all_${dp}.zip`);}
}
async function _exportAllSingleView(){
  const overlay=makeOverlay('Exporting current view…','PNG + PPTX + CSV');document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));let stage;
  try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');const zip=new JSZip();zip.file('orgchart.csv',buildCSVContent());zip.file('orgchart.png',canvas.toDataURL('image/png').split(',')[1],{base64:true});const pptxBlob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height);zip.file('orgchart.pptx',pptxBlob);const zipBlob=await zip.generateAsync({type:'blob',compression:'DEFLATE'});triggerDownload(zipBlob,`orgchart_${dp}.zip`);}
  catch(e){alert('Export failed: '+e.message);}
  finally{if(stage)stage.remove();overlay.remove();applyZoom(savedZoom);}
}

/* ── REASSIGN MODAL ── */
let _reassignAllNodes=[];
function openReassignModal(e,nodeId){
  e.stopPropagation();S.reassignTarget=nodeId;S.reassignPick=null;
  const node=S.viewData.find(n=>n.id===nodeId);
  document.getElementById('reassign-subject').innerHTML=`Moving <strong>${esc(node?.name||nodeId)}</strong>`;
  document.getElementById('reassign-search').value='';
  document.getElementById('reassign-confirm-btn').disabled=true;
  document.getElementById('reassign-note').textContent='Select a new manager above';
  _reassignAllNodes=[{id:'__root__',name:'Make Root (no manager)',manager:''},...S.viewData.filter(n=>n.id!==nodeId)];
  renderReassignList(_reassignAllNodes);
  document.getElementById('reassign-modal').classList.remove('hidden');
}
function closeReassignModal(){document.getElementById('reassign-modal').classList.add('hidden');S.reassignTarget=null;S.reassignPick=null;}
function filterReassignList(){const q=document.getElementById('reassign-search').value.trim().toLowerCase();renderReassignList(q?_reassignAllNodes.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)):_reassignAllNodes);}
function renderReassignList(nodes){
  document.getElementById('reassign-list').innerHTML=nodes.slice(0,60).map(n=>{
    const isRoot=n.id==='__root__';const initials=n.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();
    return`<div class="modal-emp-row${S.reassignPick===n.id?' selected':''}" onclick="pickReassign('${esc(n.id)}','${esc(n.name)}')">
      <div class="modal-emp-avatar">${isRoot?'🔼':esc(initials)}</div>
      <div><div class="modal-emp-name">${esc(n.name)}</div><div class="modal-emp-sub">${isRoot?'Will appear as root node':esc(n.id)}</div></div>
    </div>`;
  }).join('');
}
function pickReassign(id,name){
  S.reassignPick=id;document.getElementById('reassign-confirm-btn').disabled=false;document.getElementById('reassign-note').textContent=`→ ${name}`;
  const q=document.getElementById('reassign-search').value.trim().toLowerCase();
  renderReassignList(q?_reassignAllNodes.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)):_reassignAllNodes);
}
function confirmReassign(){if(!S.reassignTarget||!S.reassignPick)return;S.managerOverrides[S.reassignTarget]=S.reassignPick==='__root__'?'':S.reassignPick;closeReassignModal();buildViewData();renderChart();}
function removeCurrentNode(){if(!S.reassignTarget)return;S.removedIds.add(S.reassignTarget);closeReassignModal();buildViewData();renderChart();}

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

/* ── EVENT BINDINGS ── */
document.getElementById('file-input').addEventListener('change',e=>{if(e.target.files[0])handleFile(e.target.files[0]);});
document.getElementById('photo-folder-input').addEventListener('change',e=>{if(e.target.files.length)loadFromFileInput(e.target.files);});
const dz=document.getElementById('upload-dropzone');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag-over');});
dz.addEventListener('dragleave',()=>dz.classList.remove('drag-over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag-over');const f=e.dataTransfer.files[0];if(f)handleFile(f);});
document.getElementById('reassign-modal').addEventListener('click',e=>{if(e.target===e.currentTarget)closeReassignModal();});
</script>
</body>
</html>'''

components.html(APP_HTML, height=900, scrolling=False)
