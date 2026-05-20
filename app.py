import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="OrgDesign Pro",page_icon="🏢",layout="wide",initial_sidebar_state="collapsed")
st.markdown("""<style>
#MainMenu,footer,header,.stDeployButton,[data-testid="stSidebar"],[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stHeader"]{display:none!important}
.stApp{background:#ffffff!important}
.block-container{padding:0!important;max-width:100%!important;margin:0!important}
iframe{border-radius:0!important;border:none!important}
</style>""",unsafe_allow_html=True)

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
:root{--bg:#ffffff;--bg2:#f8fafc;--bg3:#f1f5f9;--bg4:#e8edf5;--border:#e2e8f0;--border2:#cbd5e1;--text:#0f172a;--text2:#475569;--text3:#94a3b8;--accent:#4f46e5;--accent2:#6366f1;--accent-light:#eef2ff;--accent-mid:#c7d2fe;--success:#059669;--success-light:#d1fae5;--warning:#d97706;--warning-light:#fef3c7;--danger:#dc2626;--shadow-xs:0 1px 2px rgba(0,0,0,0.05);--shadow-sm:0 1px 4px rgba(0,0,0,0.07),0 1px 2px rgba(0,0,0,0.04);--shadow-md:0 4px 16px rgba(0,0,0,0.08),0 2px 4px rgba(0,0,0,0.04);--shadow-lg:0 12px 40px rgba(0,0,0,0.1),0 4px 12px rgba(0,0,0,0.05);--r:10px;--r-lg:14px;--r-xl:18px}
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
/* MAP GRID: 2x2 for 4 fields */
.map-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;max-width:740px;margin-bottom:28px}
.map-card{background:var(--bg);border:1.5px solid var(--border);border-radius:var(--r-lg);padding:16px;transition:border-color 0.2s,box-shadow 0.2s}
.map-card:focus-within{border-color:var(--accent);box-shadow:0 0 0 3px rgba(79,70,229,0.08)}
.map-card-label{font-size:0.7rem;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);margin-bottom:8px;display:flex;align-items:center;gap:7px}
.badge-req{background:#fee2e2;color:var(--danger);padding:1px 7px;border-radius:999px;font-size:0.6rem;font-weight:700}
.badge-opt{background:var(--bg3);color:var(--text3);padding:1px 7px;border-radius:999px;font-size:0.6rem;font-weight:700}
.badge-fro{background:#f5f3ff;color:#7c3aed;padding:1px 7px;border-radius:999px;font-size:0.6rem;font-weight:700}
.map-select{width:100%;background:var(--bg3);border:1.5px solid var(--border);border-radius:8px;padding:8px 10px;font-size:0.84rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;cursor:pointer;appearance:none;background-repeat:no-repeat;background-position:right 10px center;transition:border-color 0.15s}
.map-select:focus{border-color:var(--accent);background-color:var(--bg)}
.map-hint{font-size:0.72rem;color:var(--text3);margin-top:6px}
.data-preview-table{width:100%;max-width:740px;border-collapse:collapse;margin-bottom:28px;font-size:0.78rem}
.data-preview-table th{background:var(--bg3);padding:7px 12px;text-align:left;font-weight:700;color:var(--text2);border:1px solid var(--border);font-size:0.72rem;text-transform:uppercase;letter-spacing:0.04em}
.data-preview-table td{padding:6px 12px;border:1px solid var(--border);color:var(--text2);max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.data-preview-table tr:nth-child(even) td{background:var(--bg2)}
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
.preview-card{width:320px;background:var(--bg);border:2px solid var(--accent);border-radius:var(--r-lg);box-shadow:var(--shadow-md)}
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
.ncard-photo{object-fit:cover;object-position:center top;display:block;flex-shrink:0}
.ncard-photo-fallback{font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0;letter-spacing:-0.02em}
.emp-type-setup{margin-top:14px}
.emp-type-row{display:flex;align-items:center;gap:8px;margin-bottom:7px}
.emp-type-label{font-size:0.78rem;font-weight:700;color:var(--text2);min-width:80px}
.emp-type-color-input{width:34px;height:28px;border-radius:7px;border:2px solid var(--border2);cursor:pointer;padding:0;background:none}
.emp-type-value-select{flex:1;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:4px 8px;font-size:0.76rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;cursor:pointer}
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
.btn-export-all{background:linear-gradient(135deg,#7c3aed,#0284c7)!important;color:#fff!important;border:none!important;box-shadow:0 4px 14px rgba(124,58,237,0.35)!important}
.btn-export-all:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(124,58,237,0.45)!important}
.chart-toolbar{flex-shrink:0;min-height:52px;height:auto;background:var(--bg);border-bottom:1px solid var(--border);display:flex;align-items:center;flex-wrap:wrap;padding:6px 12px;gap:6px;box-shadow:var(--shadow-xs);position:relative;z-index:20;overflow:visible}
.stats-bar{flex-shrink:0;height:34px;background:var(--bg2);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 18px;gap:18px;font-size:0.73rem}
.stat-item{display:flex;align-items:center;gap:6px;color:var(--text3);font-weight:600}
.stat-item strong{color:var(--text);font-weight:800}
.stat-dot{width:6px;height:6px;border-radius:50%;background:var(--accent)}
.filter-bar{flex-shrink:0;background:var(--bg);border-bottom:1px solid var(--border);padding:7px 18px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;min-height:44px}
.filter-dropdown-wrap{display:flex;align-items:center;gap:6px;font-size:0.79rem}
.filter-dropdown-label{font-weight:700;color:var(--text2)}
.filter-dropdown{background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:5px 28px 5px 10px;font-size:0.79rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-repeat:no-repeat;background-position:right 8px center;transition:border-color 0.15s}
.filter-dropdown:focus{border-color:var(--accent)}
.photo-btn{display:flex;align-items:center;gap:5px;padding:5px 10px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0}
.photo-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.photo-btn.loaded{border-color:#059669;color:#059669;background:#d1fae5}
.photo-count{background:#059669;color:#fff;border-radius:999px;padding:1px 6px;font-size:0.65rem;font-weight:800}
.depth-wrap{display:flex;align-items:center;gap:5px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:3px 6px 3px 9px;flex-shrink:0}
.depth-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);white-space:nowrap}
.depth-select{background:transparent;border:none;border-radius:6px;padding:3px 20px 3px 4px;font-size:0.78rem;font-weight:700;color:var(--accent);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-repeat:no-repeat;background-position:right 3px center}
.mgr-mode-btn{display:flex;align-items:center;gap:6px;padding:5px 11px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0;user-select:none}
.mgr-mode-btn:hover{border-color:#7c3aed;color:#7c3aed;background:#f5f3ff}
.mgr-mode-btn.active{background:#f5f3ff;border-color:#7c3aed;color:#7c3aed;box-shadow:0 0 0 2px #ddd6fe}
.mgr-mode-dot{width:8px;height:8px;border-radius:50%;background:var(--border2);transition:background 0.15s}
.mgr-mode-btn.active .mgr-mode-dot{background:#7c3aed}
.summary-fields-wrap{display:flex;align-items:center;gap:5px;background:#fdf4ff;border:1.5px solid #e9d5ff;border-radius:8px;padding:3px 6px 3px 9px;flex-shrink:0}
.summary-fields-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:#7c3aed;white-space:nowrap}
.summary-field-select{background:transparent;border:none;padding:3px 18px 3px 4px;font-size:0.75rem;font-weight:700;color:#7c3aed;font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-repeat:no-repeat;background-position:right 2px center;max-width:110px}
.summary-list-card{display:inline-block;width:240px;background:#ffffff;border:1.5px solid #e2e8f0;border-top:3px solid #7c3aed;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,0.07);font-family:'Plus Jakarta Sans',sans-serif;overflow:hidden;vertical-align:top;text-align:left;position:relative;z-index:3}
.chart-canvas-wrap{flex:1;overflow:auto;background:var(--bg3);cursor:grab;position:relative}
.chart-canvas-wrap:active{cursor:grabbing}
.chart-canvas-content{display:inline-block;padding:56px 80px 120px 80px;transform-origin:top left;position:relative;z-index:1}
.org-tree{display:inline-block}
.org-tree ul{padding-top:24px;position:relative;list-style:none;display:flex;justify-content:center}
.org-tree li{display:table-cell;vertical-align:top;text-align:center;position:relative;padding:24px 7px 0 7px}
.org-tree li::before,.org-tree li::after{content:'';position:absolute;top:0;right:50%;border-top:2px solid #94a3b8;width:50%;height:24px}
.org-tree li::after{right:auto;left:50%;border-left:2px solid #94a3b8}
.org-tree li:only-child::before,.org-tree li:only-child::after{display:none}
.org-tree li:first-child::before,.org-tree li:last-child::after{display:none}
.org-tree li:first-child::after{border-radius:6px 0 0 0}
.org-tree li:last-child::before{border-radius:0 6px 0 0}
.org-tree ul ul::before{content:'';position:absolute;top:0;left:50%;border-left:2px solid #94a3b8;height:24px}
.org-tree li.collapsed > ul{display:none!important}
.row-size-wrap{display:flex;align-items:center;gap:5px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:3px 6px 3px 9px;flex-shrink:0}
.row-size-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);white-space:nowrap}
.row-size-select{background:transparent;border:none;border-radius:6px;padding:3px 18px 3px 4px;font-size:0.78rem;font-weight:700;color:var(--accent);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none}
/* ── Children row chunking for max-6-per-row ── */
.children-rows-wrap{display:flex;flex-direction:column;align-items:center;gap:14px;position:relative;padding-top:24px}
.children-rows-wrap::before{content:'';position:absolute;top:0;left:50%;border-left:2px solid #94a3b8;height:24px;margin-left:-1px;z-index:0}
.children-row-ul{padding-top:24px;position:relative;list-style:none;display:flex;justify-content:center;flex-wrap:nowrap}
.children-row-ul::before{display:none}
.children-row-ul.row-cont{padding-top:30px;border-top:1.5px dashed #94a3b8;margin-top:6px;position:relative}
.children-row-ul.row-cont::after{content:'\2937 continued';position:absolute;top:-9px;left:50%;transform:translateX(-50%);background:var(--bg);padding:0 9px;font-size:0.6rem;font-weight:800;color:var(--text3);text-transform:uppercase;letter-spacing:0.07em;border-radius:6px}
/* Drag-to-reassign visual states */
.node-card.drop-halo{box-shadow:0 0 0 4px rgba(34,197,94,0.45),0 0 30px rgba(34,197,94,0.35)!important;border-color:#16a34a!important}
.node-card.drop-halo-bad{box-shadow:0 0 0 4px rgba(239,68,68,0.4),0 0 30px rgba(239,68,68,0.25)!important;border-color:#dc2626!important;cursor:not-allowed}
.node-card.node-dragging{opacity:0.45;cursor:grabbing!important}
#root-drop-zone{display:none;align-items:center;gap:6px;padding:5px 11px;background:#fef3c7;border:1.5px dashed #d97706;border-radius:8px;font-size:0.74rem;font-weight:700;color:#92400e;flex-shrink:0;transition:all 0.15s}
#root-drop-zone.dragging-active{display:inline-flex}
#root-drop-zone.over{background:#d97706;color:#fff;transform:scale(1.05)}
.toast{position:fixed;bottom:30px;left:50%;transform:translateX(-50%) translateY(8px);background:#0f172a;color:#fff;padding:10px 18px;border-radius:10px;font-size:0.82rem;font-weight:600;z-index:10000;box-shadow:0 12px 40px rgba(0,0,0,0.25);opacity:0;transition:opacity 0.18s,transform 0.18s;pointer-events:none;display:flex;align-items:center;gap:10px}
.toast.visible{opacity:1;transform:translateX(-50%) translateY(0)}
.toast .toast-action{background:rgba(255,255,255,0.18);border:none;color:#fff;font-size:0.74rem;font-weight:800;padding:4px 10px;border-radius:6px;cursor:pointer;font-family:inherit;pointer-events:auto}
.toast .toast-action:hover{background:rgba(255,255,255,0.28)}
/* ── Node card: full 4-side border color ── */
.node-card{display:inline-block;width:270px;background:var(--bg);border:2px solid var(--accent);border-radius:var(--r-lg);cursor:pointer;text-align:left;transition:transform 0.15s,box-shadow 0.15s,border-color 0.15s;box-shadow:var(--shadow-sm);position:relative;z-index:3;font-family:'Plus Jakarta Sans',sans-serif}
.node-card:hover{transform:translateY(-3px);box-shadow:0 8px 28px rgba(0,0,0,0.12),0 0 0 2px rgba(79,70,229,0.12);z-index:10}
.node-card.highlighted{box-shadow:0 0 0 3px rgba(217,119,6,0.2),0 8px 24px rgba(0,0,0,0.1)!important}
.node-card.collapsed-node{opacity:0.65}
.ncard-header{padding:6px 10px;background:var(--bg2);border-bottom:1px solid var(--border);border-radius:12px 12px 0 0;display:flex;align-items:center;gap:4px}
.ncard-footer{padding:6px 10px;border-top:1px solid var(--border);border-radius:0 0 12px 12px;background:var(--bg2);display:flex;align-items:center;gap:4px}
.ncard-slot{font-size:0.64rem;font-weight:700;color:var(--text3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;text-align:center}
.ncard-slot.has-val{color:var(--card-accent,var(--accent))}
.ncard-body{padding:14px 14px 12px}
.ncard-body-inner{display:flex;gap:10px}
.ncard-body-b1{border-top:1px dashed var(--border2);margin-top:6px;padding-top:5px;text-align:center;font-size:0.72rem;color:var(--accent);font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ncard-text-wrap{width:100%;text-align:center}
.ncard-name{font-size:0.88rem;font-weight:800;color:var(--text);margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:-0.01em}
.ncard-sub{font-size:0.74rem;color:var(--text2);line-height:1.4;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.collapse-btn{position:absolute;bottom:-11px;left:50%;transform:translateX(-50%);width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:0.58rem;color:var(--text3);transition:all 0.15s;z-index:5;box-shadow:var(--shadow-xs)}
.collapse-btn:hover{background:var(--accent);border-color:var(--accent);color:#fff}
.search-wrap{position:relative;flex:0 0 auto;width:240px;height:34px}
.search-icon{position:absolute;left:13px;top:50%;transform:translateY(-50%);font-size:0.95rem;pointer-events:none;color:#fff;z-index:2;filter:drop-shadow(0 1px 1px rgba(0,0,0,0.25))}
#chart-search{width:100%;height:100%;background:linear-gradient(135deg,#6d28d9 0%,#7c3aed 60%,#a855f7 100%);border:none;border-radius:999px;padding:0 14px 0 36px;font-size:0.8rem;font-weight:600;color:#fff;font-family:'Plus Jakarta Sans',sans-serif;outline:none;box-shadow:0 0 0 2px rgba(124,58,237,0.16),0 4px 12px rgba(124,58,237,0.3),inset 0 1px 0 rgba(255,255,255,0.22);transition:box-shadow 0.18s ease}
#chart-search:focus{box-shadow:0 0 0 3px rgba(192,38,211,0.28),0 6px 18px rgba(192,38,211,0.42),inset 0 1px 0 rgba(255,255,255,0.32)}
#chart-search::placeholder{color:rgba(255,255,255,0.85);font-weight:700;letter-spacing:0.05em;text-transform:uppercase;font-size:0.7rem}
#chart-search-results{position:fixed;background:var(--bg);border:1.5px solid var(--border);border-radius:var(--r);box-shadow:var(--shadow-lg);max-height:320px;overflow-y:auto;z-index:99999;display:none;min-width:280px}
#chart-search-results.visible{display:block}
.sr-item{display:flex;align-items:center;gap:10px;padding:9px 14px;cursor:pointer;border-bottom:1px solid var(--border);transition:background 0.1s ease}
.sr-item:last-child{border-bottom:none}
.sr-item:hover{background:rgba(124,58,237,0.08)}
.sr-item.sr-selected{background:rgba(124,58,237,0.14)}
.sr-info{flex:1;min-width:0}
.sr-name{font-weight:700;font-size:0.84rem;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sr-sub{font-size:0.7rem;color:var(--text3);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sr-chev{flex-shrink:0;color:#7c3aed;font-size:1.3rem;font-weight:300;line-height:1;opacity:0.5}
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
.vacant-select{flex:1;min-width:100px;background:var(--bg);border:1.5px solid var(--border);border-radius:8px;padding:5px 8px;font-size:0.78rem;font-weight:600;color:var(--text);font-family:'Plus Jakarta Sans',sans-serif;outline:none;appearance:none;cursor:pointer;background-repeat:no-repeat;background-position:right 7px center}
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
.preview-b1-zone{margin:4px 0 0 0;padding-top:6px;border-top:1px dashed var(--border2)}
.export-stage-root .org-tree li,.export-stage-root .org-tree ul,.export-stage-root .node-card,.export-stage-root .summary-list-card{overflow:visible!important}
.bg-control-wrap{display:flex;align-items:center;gap:5px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;padding:3px 6px;flex-shrink:0}
.bg-control-label{font-size:0.65rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3);white-space:nowrap}
.bg-color-input{width:24px;height:22px;border-radius:5px;border:1.5px solid var(--border2);cursor:pointer;padding:0;background:none;flex-shrink:0}
.bg-color-input:disabled{opacity:0.4;cursor:not-allowed}
.bg-transparent-btn{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;background:var(--bg);border:1.5px solid var(--border);border-radius:6px;font-size:0.7rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;user-select:none;line-height:1.2;font-family:'Plus Jakarta Sans',sans-serif}
.bg-transparent-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.bg-transparent-btn.active{background:#fef3c7;border-color:#d97706;color:#92400e;box-shadow:0 0 0 2px #fde68a}
.chart-canvas-wrap.transparent-preview{background-color:#ffffff!important;background-image:linear-gradient(45deg,#e2e8f0 25%,transparent 25%),linear-gradient(-45deg,#e2e8f0 25%,transparent 25%),linear-gradient(45deg,transparent 75%,#e2e8f0 75%),linear-gradient(-45deg,transparent 75%,#e2e8f0 75%)!important;background-size:20px 20px!important;background-position:0 0,0 10px,10px -10px,-10px 0!important}
.photo-match-select{background:transparent;border:none;padding:3px 18px 3px 4px;font-size:0.75rem;font-weight:700;color:var(--accent);font-family:'Plus Jakarta Sans',sans-serif;cursor:pointer;outline:none;appearance:none;background-repeat:no-repeat;background-position:right 2px center;max-width:130px}
/* ── Person View Modal ── */
.pv-modal-box{background:var(--bg);border:1px solid var(--border);border-radius:var(--r-xl);box-shadow:0 24px 80px rgba(0,0,0,0.22);width:94vw;max-width:1400px;height:90vh;display:flex;flex-direction:column;overflow:hidden}
.pv-modal-header{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;flex-wrap:wrap;gap:8px}
.pv-chart-area{flex:1;overflow:auto;background:var(--bg3);position:relative;cursor:grab}
.pv-chart-area:active{cursor:grabbing}
.pv-tree-content{display:inline-block;padding:48px 60px 80px 60px;position:relative;transform-origin:top left}
/* ── Depth buttons in PV ── */
.pv-depth-btn{padding:4px 11px;border:1.5px solid var(--border2);border-radius:999px;font-size:0.72rem;font-weight:700;cursor:pointer;background:var(--bg);color:var(--text3);transition:all 0.15s;font-family:'Plus Jakarta Sans',sans-serif}
.pv-depth-btn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent-light)}
.pv-depth-btn.selected{background:var(--accent);border-color:var(--accent);color:#fff}
/* ── FRO legend indicator ── */
.fro-legend{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;background:#f5f3ff;border:1px solid #ddd6fe;border-radius:6px;font-size:0.7rem;font-weight:700;color:#7c3aed}
.fro-legend-line{width:20px;height:2px;background:repeating-linear-gradient(90deg,#7c3aed 0,#7c3aed 5px,transparent 5px,transparent 9px);display:inline-block}
/* ── Data Quality banner + modal ── */
.dq-toolbar-btn{display:inline-flex;align-items:center;gap:6px;padding:5px 11px;background:#fef3c7;border:1.5px solid #f59e0b;border-radius:8px;font-size:0.74rem;font-weight:800;color:#92400e;cursor:pointer;font-family:inherit}
.dq-toolbar-btn:hover{background:#fde68a}
.dq-toolbar-btn.clean{background:#d1fae5;border-color:#059669;color:#065f46}
.dq-section{border:1px solid var(--border);border-radius:10px;padding:12px 14px;margin-bottom:10px;background:var(--bg2)}
.dq-section.empty{opacity:0.55}
.dq-section h4{font-size:0.86rem;font-weight:800;color:var(--text);display:flex;align-items:center;gap:8px;margin-bottom:6px}
.dq-count{background:var(--danger);color:#fff;border-radius:999px;padding:1px 8px;font-size:0.65rem;font-weight:800}
.dq-section.empty .dq-count{background:#16a34a}
.dq-list{font-size:0.78rem;color:var(--text2);line-height:1.6;max-height:140px;overflow-y:auto}
.dq-list .dq-row{padding:3px 0;border-bottom:1px dashed var(--border);display:flex;justify-content:space-between;gap:10px}
.dq-list .dq-row:last-child{border-bottom:none}
.dq-list .dq-row .dq-fix{font-size:0.7rem;font-weight:700;color:var(--accent);cursor:pointer}
.dq-list .dq-row .dq-fix:hover{text-decoration:underline}
.dq-bulk{display:flex;gap:8px;margin-top:10px}
/* ── Live map-select conflict ── */
.map-select.conflict{border-color:var(--danger)!important;background:#fef2f2!important}
.map-hint.warn{color:#b45309;font-weight:600}
.map-hint.err{color:var(--danger);font-weight:600}
.map-hint.ok{color:#059669;font-weight:600}
/* ── Grid Mode v2: tree stays visible, drag to translate, SVG connectors ── */
#org-grid{display:none}
.chart-canvas-content.grid-mode #org-tree li::before,
.chart-canvas-content.grid-mode #org-tree li::after,
.chart-canvas-content.grid-mode #org-tree ul ul::before,
.chart-canvas-content.grid-mode .children-rows-wrap::before,
.chart-canvas-content.grid-mode .children-row-ul.row-cont{border:none!important;background:transparent!important}
.chart-canvas-content.grid-mode #org-tree li::before,
.chart-canvas-content.grid-mode #org-tree li::after,
.chart-canvas-content.grid-mode #org-tree ul ul::before{display:none!important}
.chart-canvas-content.grid-mode #fro-svg{display:none}
.chart-canvas-content.grid-mode .node-card.grid-translated{z-index:6;box-shadow:0 6px 20px rgba(15,23,42,0.18)!important}
.grid-overlay{position:absolute;top:0;left:0;width:100%;height:100%;background-image:linear-gradient(to right,rgba(148,163,184,0.22) 1px,transparent 1px),linear-gradient(to bottom,rgba(148,163,184,0.22) 1px,transparent 1px);background-size:20px 20px;background-position:0 0;pointer-events:none;z-index:0;display:none}
.chart-canvas-content.grid-mode .grid-overlay.visible{display:block}
.grid-svg{position:absolute;top:0;left:0;pointer-events:none;z-index:1;display:none}
.chart-canvas-content.grid-mode .grid-svg{display:block}
/* Multi-select selection indicator */
.node-card.selected{box-shadow:0 0 0 3px rgba(34,197,94,0.6),0 8px 24px rgba(34,197,94,0.2)!important;outline:none!important}
.align-toolbar{position:fixed;top:64px;left:50%;transform:translateX(-50%);background:#0f172a;color:#fff;padding:6px;border-radius:12px;box-shadow:0 16px 48px rgba(0,0,0,0.35);z-index:8500;display:none;gap:3px;align-items:center;font-family:'Plus Jakarta Sans',sans-serif}
.align-toolbar.visible{display:flex}
.align-toolbar .at-label{padding:0 10px;font-size:0.72rem;opacity:0.85;font-weight:700}
.align-toolbar .at-sep{width:1px;height:18px;background:rgba(255,255,255,0.18);margin:0 2px}
.align-toolbar .at-btn{background:rgba(255,255,255,0.06);border:none;color:#fff;width:30px;height:30px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:700;font-family:inherit;transition:background 0.15s;display:flex;align-items:center;justify-content:center;padding:0}
.align-toolbar .at-btn:hover{background:rgba(255,255,255,0.18)}
.align-toolbar .at-btn.danger:hover{background:rgba(239,68,68,0.4)}
/* Person-View grid mirror — translate-based on the EXISTING PV tree */
#pv-org-grid{display:none}
.pv-tree-content.grid-mode #pv-org-tree li::before,
.pv-tree-content.grid-mode #pv-org-tree li::after,
.pv-tree-content.grid-mode #pv-org-tree ul ul::before,
.pv-tree-content.grid-mode #pv-org-tree .children-rows-wrap::before{display:none!important}
.pv-tree-content.grid-mode #pv-fro-svg{display:none}
.pv-tree-content.grid-mode #pv-grid-svg{display:block}
.pv-tree-content.grid-mode .node-card.grid-translated{z-index:6;box-shadow:0 6px 20px rgba(15,23,42,0.18)!important}
.grid-mode-btn{display:flex;align-items:center;gap:6px;padding:5px 11px;background:var(--bg2);border:1.5px solid var(--border);border-radius:8px;font-size:0.74rem;font-weight:700;color:var(--text2);cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0;font-family:inherit}
.grid-mode-btn:hover{border-color:#0891b2;color:#0891b2;background:#ecfeff}
.grid-mode-btn.active{background:#ecfeff;border-color:#0891b2;color:#0891b2;box-shadow:0 0 0 2px #cffafe}
.grid-mode-dot{width:8px;height:8px;border-radius:50%;background:var(--border2)}
.grid-mode-btn.active .grid-mode-dot{background:#0891b2}
/* ── Print stylesheet (Phase 6) — A3 landscape, chart only ── */
@media print{
  @page{size:A3 landscape;margin:8mm}
  html,body{background:#fff!important;overflow:visible!important;height:auto!important}
  .topnav,.chart-toolbar,.stats-bar,.filter-bar,.modal-overlay,.collapse-btn,.ncard-edit-btn,.ncard-export-btn,.toast,#root-drop-zone,#fro-legend,.dq-toolbar-btn,#chart-search-results{display:none!important}
  .main,.screen,#screen-chart,.chart-canvas-wrap,.chart-canvas-content{position:static!important;overflow:visible!important;background:#fff!important;padding:0!important;width:auto!important;height:auto!important;transform:none!important}
  .org-tree{transform:none!important}
  .node-card{break-inside:avoid;page-break-inside:avoid;box-shadow:none!important}
  .summary-list-card{break-inside:avoid;page-break-inside:avoid;box-shadow:none!important}
  /* Hide grid overlay (Phase 9) when printing */
  .grid-overlay{display:none!important}
}
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
  <div class="screen active" id="screen-upload">
    <div class="upload-center">
      <div class="upload-hero"><h1>Build your Org Chart</h1><p>Upload your HR roster and we'll guide you through designing a beautiful, interactive org chart in minutes.</p></div>
      <div class="upload-zone" id="upload-dropzone" onclick="document.getElementById('file-input').click()">
        <input type="file" id="file-input" accept=".csv,.xlsx,.xls" onchange="if(this.files[0])handleFile(this.files[0])"/>
        <span class="upload-emoji">📊</span><h3>Drop your file here</h3><p>Supports CSV and Excel (.xlsx, .xls)<br>or <span>click to browse</span></p>
      </div>
      <div id="lib-status" style="font-size:0.78rem;color:var(--text2);font-weight:600;display:flex;gap:14px;flex-wrap:wrap;justify-content:center;margin-top:-8px"></div>
      <div style="display:flex;gap:10px"><button class="btn btn-ghost btn-sm" onclick="loadDemoData()">📁 Load demo data</button></div>
      <div class="info-cards">
        <div class="info-card"><div class="info-card-title">Required Columns</div><div class="info-card-row">Employee Code / ID</div><div class="info-card-row">Employee Name</div><div class="info-card-row">Manager Code / ID</div></div>
        <div class="info-card"><div class="info-card-title">FRO &amp; Photo Tips</div><div class="info-card-row">Map FRO column for dotted functional lines</div><div class="info-card-row">Name photos by any column value</div><div class="info-card-row">Pick match column in Chart toolbar</div></div>
      </div>
      <div style="font-size:0.74rem;color:var(--text3);margin-top:6px">Trouble re-uploading? <a href="#" onclick="event.preventDefault();clearPersisted();alert('Saved session cleared. Reload the page or upload again.');" style="color:var(--accent);text-decoration:underline;font-weight:700">Clear saved session</a></div>
    </div>
  </div>
  <div class="screen" id="screen-map">
    <div class="section-header"><div class="section-title">Map Your Columns</div><div class="section-sub">We detected <span id="col-count">0</span> columns. Auto-mapped where possible.</div></div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Detected Columns</div>
    <div class="detected-chips" id="detected-columns"></div>
    <div class="map-grid">
      <div class="map-card"><div class="map-card-label">Employee ID <span class="badge-req">Required</span></div><select class="map-select" id="map-empId" onchange="onMapChange()"></select><div class="map-hint" id="hint-empId">Unique identifier — also used to match photos</div></div>
      <div class="map-card"><div class="map-card-label">Employee Name <span class="badge-req">Required</span></div><select class="map-select" id="map-empName" onchange="onMapChange()"></select><div class="map-hint" id="hint-empName">Full name shown on the card</div></div>
      <div class="map-card"><div class="map-card-label">Manager ID <span class="badge-opt">Optional</span></div><select class="map-select" id="map-managerId" onchange="onMapChange()"></select><div class="map-hint" id="hint-managerId">Links employee to their direct line manager</div></div>
      <div class="map-card" style="border-color:#ddd6fe"><div class="map-card-label">FRO / Functional Manager ID <span class="badge-fro">Optional · Dotted Line</span></div><select class="map-select" id="map-froId" style="border-color:#ddd6fe" onchange="onMapChange()"></select><div class="map-hint" id="hint-froId">Functional reporting officer — shown as a purple dotted line on the chart</div></div>
    </div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Data Preview (first 3 rows)</div>
    <div id="data-preview-wrap" style="margin-bottom:24px;overflow-x:auto"></div>
    <div style="display:flex;gap:12px;margin-top:0"><button class="btn btn-ghost" onclick="goTo('upload')">Back</button><button class="btn btn-primary" onclick="confirmColumnMap()">Continue to Card Design</button></div>
  </div>
  <div class="screen" id="screen-card">
    <div class="section-header" style="margin-bottom:18px"><div class="section-title">Design Your Card</div><div class="section-sub">Drag fields into slots. Configure employment type colors.</div></div>
    <div class="card-design-layout">
      <div class="fields-panel">
        <div class="fields-panel-title">Available Fields</div><div id="card-fields-panel"></div>
        <div class="fields-section" style="margin-top:16px"><div class="fields-section-label" style="margin-bottom:8px">Card Accent Color</div><div class="color-palette" id="color-palette"></div></div>
        <div class="fields-section" style="margin-top:14px">
          <div class="fields-section-label" style="margin-bottom:8px">Photo Size &amp; Shape</div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:7px">Size (px)</div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px"><input type="range" id="photo-size-slider" min="40" max="160" step="10" value="80" style="flex:1;accent-color:var(--accent);cursor:pointer" oninput="S.photoSize=parseInt(this.value);document.getElementById('photo-size-val').textContent=this.value+'px';renderCardPreview();renderChart();persistState();" aria-label="Photo size in pixels"/><span id="photo-size-val" style="font-size:0.78rem;font-weight:700;color:var(--accent);min-width:36px">80px</span></div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:7px">Shape</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px"><div class="shape-btn selected" data-shape="circle" onclick="setPhotoShape('circle')">Circle</div><div class="shape-btn" data-shape="rounded" onclick="setPhotoShape('rounded')">Rounded</div><div class="shape-btn" data-shape="square" onclick="setPhotoShape('square')">Square</div></div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:7px">Placement</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap"><div class="shape-btn selected" data-placement="top" onclick="setPhotoPlacement('top')">Top</div><div class="shape-btn" data-placement="left" onclick="setPhotoPlacement('left')">Left</div><div class="shape-btn" data-placement="right" onclick="setPhotoPlacement('right')">Right</div><div class="shape-btn" data-placement="none" onclick="setPhotoPlacement('none')">None</div></div>
        </div>
        <div class="fields-section emp-type-setup">
          <div class="fields-section-label" style="margin-bottom:10px">Employment Type Colors</div>
          <div style="font-size:0.73rem;color:var(--text3);margin-bottom:10px;line-height:1.5">Map a column + values to control card border color.</div>
          <div style="margin-bottom:8px"><select class="vacant-select" id="emp-type-col" onchange="onEmpTypeColChange()" style="width:100%;margin-bottom:8px"><option value="">Select column...</option></select></div>
          <div id="emp-type-rows" style="display:none">
            <div class="emp-type-row"><div class="emp-type-label">Active</div><select class="emp-type-value-select" id="emp-val-active"><option value="">Value...</option></select><input type="color" class="emp-type-color-input" id="emp-color-active" value="#059669"/></div>
            <div class="emp-type-row"><div class="emp-type-label">Vacant</div><select class="emp-type-value-select" id="emp-val-vacant"><option value="">Value...</option></select><input type="color" class="emp-type-color-input" id="emp-color-vacant" value="#dc2626"/></div>
            <div class="emp-type-row"><div class="emp-type-label">Resigned</div><select class="emp-type-value-select" id="emp-val-resigned"><option value="">Value...</option></select><input type="color" class="emp-type-color-input" id="emp-color-resigned" value="#d97706"/></div>
          </div>
        </div>
      </div>
      <div class="card-preview-area"><div class="preview-label">Live Card Preview</div><div id="card-preview"></div><div class="preview-hint">Drag field chips onto header, body, or footer zones. Name is always fixed in the card body.</div></div>
    </div>
    <div style="display:flex;gap:12px;margin-top:16px"><button class="btn btn-ghost" onclick="goTo('map')">Back</button><button class="btn btn-primary" onclick="confirmCardDesign()">Continue to Filters</button></div>
  </div>
  <div class="screen" id="screen-filter">
    <div class="section-header"><div class="section-title">Set Up Filters</div><div class="section-sub">Choose up to 3 columns to use as filters. The last filter drives "Export All".</div></div>
    <div class="filter-setup"><div class="filter-counter" id="filter-counter">0 of 3 filters selected</div><div class="filter-chips" id="filter-chip-picker"></div><div id="filter-preview-area"></div></div>
    <div style="display:flex;gap:12px;margin-top:16px"><button class="btn btn-ghost" onclick="goTo('card')">Back</button><button class="btn btn-primary" onclick="launchChart()">Launch Org Chart</button></div>
  </div>
  <div class="screen" id="screen-chart">
    <div class="chart-toolbar">
      <button class="btn btn-ghost btn-sm" onclick="goTo('filter')">Setup</button><div class="tb-sep"></div>
      <div class="search-wrap" id="search-wrap"><span class="search-icon">🔍</span><input id="chart-search" type="text" placeholder="Person View" autocomplete="off"/></div><div class="tb-sep"></div>
      <div class="zoom-strip"><button class="btn-zoom" onclick="zoomBy(-0.1)">−</button><span class="zoom-label" id="zoom-level">100%</span><button class="btn-zoom" onclick="zoomBy(0.1)">+</button><button class="btn-zoom" onclick="fitToScreen(true)" title="Fit">⊡</button></div>
      <button class="btn btn-ghost btn-sm" onclick="centerView()">Center</button><button class="btn btn-ghost btn-sm" onclick="expandAll()">Expand</button><button class="btn btn-ghost btn-sm" onclick="collapseAll()">Collapse</button><div class="tb-sep"></div>
      <div class="depth-wrap"><span class="depth-label">Skip Top</span><select class="depth-select" id="depth-select" onchange="setSkipDepth(parseInt(this.value))" aria-label="Skip top N levels of hierarchy"><option value="0">None</option><option value="1">L1</option><option value="2">L2</option><option value="3">L3</option><option value="4">L4</option><option value="5">L5</option><option value="6">L6</option></select></div>
      <div class="row-size-wrap" title="Maximum number of cards in a single row. 'Auto' picks a near-square layout based on the count of siblings."><span class="row-size-label">Per row</span><select class="row-size-select" id="row-size-select" onchange="setMaxPerRow(this.value)" aria-label="Maximum cards per row"><option value="auto">Auto</option><option value="2">2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option><option value="6" selected>6</option><option value="8">8</option><option value="10">10</option><option value="12">12</option></select></div>
      <div class="tb-sep"></div>
      <div class="mgr-mode-btn" id="mgr-mode-btn" onclick="toggleManagerMode()" title="Compact ICs into a single summary list under each manager"><div class="mgr-mode-dot"></div>Manager View</div>
      <div class="grid-mode-btn" id="grid-mode-btn" onclick="toggleGridMode()" title="Switch to free-form grid: cards auto-place by depth (level=row), drag any card to any cell to override"><div class="grid-mode-dot"></div>Grid Mode</div>
      <button class="btn btn-ghost btn-sm" id="grid-lines-btn" onclick="toggleGridLines()" title="Show/hide gridlines (gridlines never appear in PNG/PPTX/print exports)" style="display:none">⊞ Gridlines</button>
      <button class="btn btn-ghost btn-sm" id="grid-reset-btn" onclick="resetGridOverrides()" title="Reset all grid position overrides — restore auto-arrangement by depth" style="display:none">↻ Auto-arrange</button>
      <div class="summary-fields-wrap" id="summary-fields-wrap" style="display:none" title="Pick up to 3 fields to display on each IC summary row"><span class="summary-fields-label">IC fields</span><select class="summary-field-select" id="summary-field1" onchange="S.summaryField1=this.value;if(S.managerMode)renderChart();persistState();" aria-label="IC summary field 1"><option value="">Field 1…</option></select><span style="font-size:0.7rem;color:#7c3aed;font-weight:700">+</span><select class="summary-field-select" id="summary-field2" onchange="S.summaryField2=this.value;if(S.managerMode)renderChart();persistState();" aria-label="IC summary field 2"><option value="">Field 2…</option></select><span style="font-size:0.7rem;color:#7c3aed;font-weight:700">+</span><select class="summary-field-select" id="summary-field3" onchange="S.summaryField3=this.value;if(S.managerMode)renderChart();persistState();" aria-label="IC summary field 3"><option value="">Field 3…</option></select></div><div class="tb-sep"></div>
      <input type="file" id="photo-folder-input" class="photo-folder-input" accept="image/*" multiple webkitdirectory/>
      <div class="bg-control-wrap" title="Pick which column value must match the photo filename (no extension)">
        <span class="bg-control-label">📁 by</span>
        <select id="photo-match-col" class="photo-match-select" onchange="S.photoMatchCol=this.value;if(S.viewData.length)renderChart();persistState();" aria-label="Column used to match photo filenames"><option value="">loading…</option></select>
      </div>
      <div class="photo-btn" id="photo-btn" onclick="openPhotoFolder()">📸 <span id="photo-btn-label">Load Photos</span><span class="photo-count" id="photo-count" style="display:none">0</span></div>
      <!-- FRO legend indicator -->
      <div class="fro-legend" id="fro-legend" style="display:none"><span class="fro-legend-line"></span>FRO line</div>
      <button class="dq-toolbar-btn" id="dq-btn" onclick="openDataQualityModal()" title="Data quality issues found in your roster (duplicates, cycles, orphans)" style="display:none">⚠ <span id="dq-count">0</span> issues</button>
      <button class="btn btn-ghost btn-sm" onclick="openInsightsModal()" title="Org-design insights (span-of-control outliers, depth, single-report managers)" id="insights-btn">💡 Insights</button>
      <button class="btn btn-ghost btn-sm" onclick="printA3()" title="Print or save as A3-landscape PDF — opens a print-ready preview window with the right paper size baked in">🖨 Print A3</button>
      <div class="tb-sep"></div>
      <div style="flex:1"></div>
      <div id="root-drop-zone" title="Drop a card here to make it a root (no manager)">⬆ Drop here to make root</div>
      <button class="btn btn-ghost btn-sm" id="undo-btn" onclick="undo()" title="Undo last change (Ctrl/Cmd+Z)" disabled>↶ Undo</button><div class="tb-sep"></div>
      <div class="bg-control-wrap" title="Background color of the chart canvas (PNG/PPTX export uses this)"><span class="bg-control-label">BG</span><input type="color" class="bg-color-input" id="bg-color-input" value="#f1f5f9" oninput="setChartBg(this.value)"/><button class="bg-transparent-btn" id="bg-transparent-btn" onclick="toggleTransparent()" title="Export with transparent background">⊘ None</button></div><div class="tb-sep"></div>
      <button class="btn btn-ghost btn-sm" onclick="downloadCSV()">CSV</button><button class="btn btn-ghost btn-sm" onclick="exportPNG()">PNG</button><button class="btn btn-ghost btn-sm" onclick="exportPPTX()">PPTX</button><button class="btn btn-sm btn-export-all" onclick="exportAll()">Export All</button>
    </div>
    <div class="stats-bar"><div class="stat-item"><div class="stat-dot"></div><strong id="stat-total">—</strong>&nbsp;employees</div><div class="stat-item"><strong id="stat-roots">—</strong>&nbsp;roots</div><div class="stat-item"><strong id="stat-vis">—</strong>&nbsp;visible</div><div class="stat-item" id="stat-photos" style="display:none;color:var(--success)">📸 <strong id="stat-photos-val">0</strong> photos</div><div class="stat-item" id="stat-mgr-mode" style="display:none;color:#7c3aed">👔 <strong id="stat-mgr-val">—</strong></div><div class="stat-item" id="stat-filtered" style="display:none;color:var(--warning)">Filtered</div></div>
    <div class="filter-bar" id="filter-bar" style="display:none"></div>
    <div class="chart-canvas-wrap" id="chart-canvas-wrap">
      <div class="chart-canvas-content" id="chart-canvas-content">
        <div class="grid-overlay" id="grid-overlay"></div>
        <svg class="grid-svg" id="grid-svg" xmlns="http://www.w3.org/2000/svg"></svg>
        <svg id="fro-svg" style="position:absolute;top:0;left:0;pointer-events:none;overflow:visible;z-index:2;display:block"></svg>
        <div class="org-tree" id="org-tree"></div>
        <div id="org-grid"></div>
      </div>
    </div>
  </div>
</main>

<!-- Multi-select align toolbar (visible when 2+ cards selected via Shift+click) -->
<div class="align-toolbar" id="align-toolbar" role="toolbar" aria-label="Align selected cards">
  <span class="at-label" id="at-count">2 selected</span>
  <span class="at-sep"></span>
  <button class="at-btn" onclick="alignSelected('left')" title="Align left edges">⬅</button>
  <button class="at-btn" onclick="alignSelected('center-h')" title="Center horizontally">⬌</button>
  <button class="at-btn" onclick="alignSelected('right')" title="Align right edges">➡</button>
  <span class="at-sep"></span>
  <button class="at-btn" onclick="alignSelected('top')" title="Align top edges">⬆</button>
  <button class="at-btn" onclick="alignSelected('middle')" title="Center vertically">⇕</button>
  <button class="at-btn" onclick="alignSelected('bottom')" title="Align bottom edges">⬇</button>
  <span class="at-sep"></span>
  <button class="at-btn" onclick="distributeSelected('h')" title="Distribute horizontally evenly">≡</button>
  <button class="at-btn" onclick="distributeSelected('v')" title="Distribute vertically evenly">⋮</button>
  <span class="at-sep"></span>
  <button class="at-btn danger" onclick="clearSelection()" title="Clear selection (Esc)">✕</button>
</div>

<!-- Search results dropdown (rendered at body level so no ancestor CSS clips it) -->
<div id="chart-search-results"></div>

<!-- Reassign Modal -->
<div class="modal-overlay hidden" id="reassign-modal">
  <div class="modal-box">
    <div class="modal-header"><div><div class="modal-title">Reassign Manager</div><div class="modal-sub" id="reassign-subject">Moving —</div></div><button class="modal-close" onclick="closeReassignModal()">✕</button></div>
    <div class="modal-body"><input class="modal-search" id="reassign-search" type="text" placeholder="Search employee name or ID..." autocomplete="off" oninput="filterReassignList()"/><div class="modal-list" id="reassign-list"></div></div>
    <div class="modal-footer"><button class="btn btn-sm" onclick="removeCurrentNode()" style="background:#fee2e2;border:1.5px solid #fca5a5;color:#dc2626;margin-right:auto">Remove</button><span class="modal-note" id="reassign-note">Select a new manager above</span><button class="btn btn-ghost btn-sm" onclick="closeReassignModal()">Cancel</button><button class="btn btn-primary btn-sm" id="reassign-confirm-btn" onclick="confirmReassign()" disabled>Reassign</button></div>
  </div>
</div>

<!-- Data Quality Modal -->
<div class="modal-overlay hidden" id="dq-modal">
  <div class="modal-box" style="width:560px;max-width:96vw">
    <div class="modal-header">
      <div><div class="modal-title">Data Quality Report</div><div class="modal-sub" id="dq-sub">Issues found in your roster</div></div>
      <button class="modal-close" onclick="closeDataQualityModal()" aria-label="Close">✕</button>
    </div>
    <div class="modal-body" id="dq-body" style="padding:16px 20px"></div>
    <div class="modal-footer">
      <span class="modal-note" id="dq-note">Removing items adds them to the chart's Removed list — Ctrl+Z undoes</span>
      <button class="btn btn-ghost btn-sm" onclick="closeDataQualityModal()">Close</button>
    </div>
  </div>
</div>

<!-- Insights Modal -->
<div class="modal-overlay hidden" id="insights-modal">
  <div class="modal-box" style="width:580px;max-width:96vw">
    <div class="modal-header">
      <div><div class="modal-title">💡 Org Design Insights</div><div class="modal-sub" id="insights-sub">Rule-based suggestions from your current chart</div></div>
      <button class="modal-close" onclick="closeInsightsModal()" aria-label="Close">✕</button>
    </div>
    <div class="modal-body" id="insights-body" style="padding:16px 20px"></div>
    <div class="modal-footer">
      <span class="modal-note">Heuristics only — review with context before acting</span>
      <button class="btn btn-ghost btn-sm" onclick="closeInsightsModal()">Close</button>
    </div>
  </div>
</div>

<!-- Person View Modal -->
<div class="modal-overlay hidden" id="person-view-modal">
  <div class="pv-modal-box">
    <div class="pv-modal-header">
      <div style="flex:1;min-width:0">
        <div class="modal-title" id="pv-title" style="font-size:1.1rem">Person View</div>
        <div class="modal-sub" id="pv-sub">Cross-filter · All raw data · FRO shown as dotted line</div>
      </div>
      <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">
        <span style="font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:var(--text3)">Depth:</span>
        <button class="pv-depth-btn selected" data-d="999" onclick="setPVDepth(999)">All</button>
        <button class="pv-depth-btn" data-d="1" onclick="setPVDepth(1)">N‑1</button>
        <button class="pv-depth-btn" data-d="2" onclick="setPVDepth(2)">N‑2</button>
        <button class="pv-depth-btn" data-d="3" onclick="setPVDepth(3)">N‑3</button>
        <button class="pv-depth-btn" data-d="4" onclick="setPVDepth(4)">N‑4</button>
        <button class="pv-depth-btn" data-d="5" onclick="setPVDepth(5)">N‑5</button>
        <div class="tb-sep"></div>
        <div class="zoom-strip"><button class="btn-zoom" onclick="pvZoomBy(-0.1)">−</button><span class="zoom-label" id="pv-zoom-level">100%</span><button class="btn-zoom" onclick="pvZoomBy(0.1)">+</button><button class="btn-zoom" onclick="pvFit()" title="Fit">⊡</button></div>
        <div class="tb-sep"></div>
        <div class="row-size-wrap" title="Max number of direct reports shown side-by-side. 'Auto' picks a near-square layout."><span class="row-size-label">Per row</span><select class="row-size-select" id="pv-row-size-select" onchange="setPVMaxPerRow(this.value)" aria-label="Person View: max cards per row"><option value="auto" selected>Auto</option><option value="2">2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option><option value="6">6</option><option value="8">8</option><option value="10">10</option><option value="12">12</option></select></div>
        <div class="tb-sep"></div>
        <button class="grid-mode-btn" id="pv-grid-btn" onclick="togglePVGrid()" title="Grid mode for this person's subtree — auto-arrange by depth, drag any card to any cell"><div class="grid-mode-dot"></div>Grid</button>
        <button class="btn btn-ghost btn-sm" onclick="printPVA3()" title="Print or save this person's chart as A3 landscape PDF">🖨 A3</button>
        <button class="btn btn-ghost btn-sm" onclick="locatePersonOnChart()" title="Close this view and highlight the person on the main chart">📌 Locate</button>
        <button class="btn btn-ghost btn-sm" onclick="exportPVPNG()" title="Save this view as a PNG image">📸 PNG</button>
        <button class="modal-close" onclick="closePV()" style="font-size:1.2rem;margin-left:4px" aria-label="Close Person View">✕</button>
      </div>
    </div>
    <div class="pv-chart-area" id="pv-chart-area">
      <div class="pv-tree-content" id="pv-tree-content">
        <svg class="grid-svg" id="pv-grid-svg" xmlns="http://www.w3.org/2000/svg"></svg>
        <svg id="pv-fro-svg" style="position:absolute;top:0;left:0;pointer-events:none;overflow:visible;z-index:2;display:block"></svg>
        <div class="org-tree" id="pv-org-tree"></div>
        <div id="pv-org-grid"></div>
      </div>
    </div>
  </div>
</div>

<script>
// Surface uncaught script errors visibly so a parse-time or runtime error doesn't silently
// disable the page (e.g. file-input change listener never gets attached). Without this,
// a single broken expression makes the whole inline <script> die quietly.
window.addEventListener('error',function(e){
  try{
    const msg='⚠ Script error: '+(e.message||'unknown')+(e.filename?' ('+e.filename+':'+e.lineno+')':'');
    console.error('OrgDesign Pro:',e.error||e.message,e);
    let bar=document.getElementById('app-error-bar');
    if(!bar){
      bar=document.createElement('div');bar.id='app-error-bar';
      bar.style.cssText='position:fixed;top:0;left:0;right:0;z-index:99999;background:#dc2626;color:#fff;padding:8px 14px;font:600 13px/1.4 -apple-system,sans-serif;display:flex;align-items:center;gap:10px';
      document.body&&document.body.appendChild(bar);
    }
    bar.textContent=msg+'  ·  Open the browser DevTools Console for details.';
  }catch(_){/* noop */}
});
const S={
  rawRows:[],columns:[],colSamples:{},
  colMap:{empId:'',empName:'',managerId:'',froId:''},
  cardSlots:{h1:'',h2:'',h3:'',b1:'',f1:'',f2:'',f3:''},
  cardAccent:'#4f46e5',
  empTypeCol:'',empTypeMap:{},empTypeLabels:{active:'',vacant:'',resigned:''},
  empTypeColors:{active:'#059669',vacant:'#dc2626',resigned:'#d97706'},
  filterCols:[],activeFilters:{},
  managerOverrides:{},removedIds:new Set(),
  viewData:[],childMap:{},descCount:{},nodeHeight:{},nodeDepth:{},
  zoom:1,highlighted:null,draggingField:null,
  reassignTarget:null,reassignPick:null,
  skipDepth:0,
  photoMap:{},photoObjUrls:[],photoSize:80,photoShape:'circle',photoPlacement:'top',photoMatchCol:'',
  managerMode:false,summaryField1:'',summaryField2:'',summaryField3:'',
  chartBgColor:'#f1f5f9',transparentExport:false,
  // Person View state
  pvPersonId:null,pvDepth:999,pvZoom:1,pvMode:false,
  // Drag-to-reassign + undo + persistence
  draggingNodeId:null,undoStack:[],_persistTimer:null,_lastFileSig:''
};
const PERSIST_KEY='orgdesign_state_v2';
const UNDO_MAX=40;

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function xe(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}

function goTo(step){document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));document.getElementById('screen-'+step).classList.add('active');const order=['upload','map','card','filter','chart'];const cur=order.indexOf(step);order.forEach((s,i)=>{const el=document.getElementById('nav-step-'+s);if(!el)return;el.className='step-item'+(i<cur?' done':i===cur?' active':'');const dot=el.querySelector('.step-dot');if(dot)dot.textContent=i<cur?'✓':String(i+1);});if(step==='chart'){setTimeout(()=>initPan(),80);setTimeout(()=>initSearch(),80);setTimeout(()=>populateSummaryFields(),120);setTimeout(()=>applyChartBg(),120);setTimeout(()=>bindRootDropZone(),140);setTimeout(()=>{const mb=document.getElementById('mgr-mode-btn');if(mb)mb.classList.toggle('active',S.managerMode);const sf=document.getElementById('summary-fields-wrap');if(sf)sf.style.display=S.managerMode?'flex':'none';const bgi=document.getElementById('bg-color-input');if(bgi)bgi.value=S.chartBgColor;const tb=document.getElementById('bg-transparent-btn');if(tb)tb.classList.toggle('active',S.transparentExport);refreshDataQualityBtn();const gb=document.getElementById('grid-mode-btn');if(gb)gb.classList.toggle('active',S.gridMode);const gl=document.getElementById('grid-lines-btn'),gr=document.getElementById('grid-reset-btn');if(gl)gl.style.display=S.gridMode?'inline-flex':'none';if(gr)gr.style.display=S.gridMode?'inline-flex':'none';const cc=document.getElementById('chart-canvas-content');if(cc)cc.classList.toggle('grid-mode',S.gridMode);const rs=document.getElementById('row-size-select');if(rs)rs.value=String(S.maxPerRow);if(S.gridMode){bindCanvasGridDND();setTimeout(()=>{applyGridOverridesToTree();redrawGridConnectorsFromTree();applyGridLines();},250);}},160);}}
function handleFile(file){
  try{
    const ext=(file.name.split('.').pop()||'').toLowerCase();
    if(ext==='csv'){
      Papa.parse(file,{header:true,skipEmptyLines:true,
        complete:r=>{try{initData(r.data);}catch(ex){console.error(ex);alert('Failed to load CSV: '+ex.message);}},
        error:e=>alert('CSV error: '+(e.message||e))});
    }else if(ext==='xlsx'||ext==='xls'){
      const reader=new FileReader();
      reader.onload=e=>{
        try{
          const wb=XLSX.read(e.target.result,{type:'array'});
          if(!wb.SheetNames.length){alert('Excel file has no sheets.');return;}
          const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
          initData(rows);
        }catch(ex){console.error(ex);alert('Failed to read Excel: '+ex.message);}
      };
      reader.onerror=()=>alert('Could not read the file.');
      reader.readAsArrayBuffer(file);
    }else{alert('Please upload a CSV or Excel file (.csv, .xlsx, .xls).');}
  }catch(ex){console.error(ex);alert('Upload failed: '+ex.message);}
}
function initData(rows){
  if(!rows||!rows.length){alert('The file looks empty — no rows were parsed.');return;}
  S.rawRows=rows;
  S.columns=Object.keys(rows[0]||{});
  if(!S.columns.length){alert('Could not detect any columns. Make sure the first row has headers.');return;}
  S.colSamples={};S.columns.forEach(col=>{S.colSamples[col]=[...new Set(rows.slice(0,25).map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='undefined'&&v!=='null'))].slice(0,3);});
  S.colMap=autoDetect(S.columns);
  S.undoStack=[];
  buildMapScreen();
  goTo('map');
  // Soft-offer: if a saved session matches this file's columns, surface a restore action.
  // This is opt-in (toast button) rather than silent auto-apply, so any bug in restoration
  // never blocks a fresh upload.
  try{
    const persisted=loadPersisted();
    if(persisted&&persisted.sig===fileSig()&&persisted.colMap&&S.columns.includes(persisted.colMap.empId)&&S.columns.includes(persisted.colMap.empName)){
      window._pendingRestore=persisted;
      setTimeout(()=>showToast('Found a saved session for this file. Restore?',true,'restore'),300);
    }
  }catch(_){/* never let restore-detection block the upload */}
}

function autoDetect(cols){const lc=cols.map(c=>c.toLowerCase().trim());function find(exact,partial){for(const p of exact){const i=lc.findIndex(c=>c===p);if(i>=0)return cols[i];}for(const p of partial){const i=lc.findIndex(c=>c.startsWith(p)||c.endsWith(p));if(i>=0)return cols[i];}for(const p of partial){const i=lc.findIndex(c=>c.includes(p));if(i>=0)return cols[i];}return '';}return{
  empId:find(['employee code','emp code','emp id','employee id','empcode','empid','staff id','employee_id','emp_id'],['employee code','emp code','employee id','emp id','empcode','empid','staff id']),
  empName:find(['employee name','emp name','full name','person name','staff name','employee_name','emp_name','full_name'],['employee name','emp name','full name','person name','staff name']),
  managerId:find(['l1 manager code','l1 manager','manager code','manager id','reports to','supervisor','mgr code','mgrid','manager_code','manager_id'],['manager code','manager id','l1 manager','reports to','supervisor','mgr code']),
  froId:find(['fro id','fro','functional manager id','functional manager code','functional reporting officer','functional reporting manager','functional_manager_id','fro_id'],['fro id','fro','functional manager','functional reporting'])
};}

async function openPhotoFolder(){if('showDirectoryPicker' in window){try{const d=await window.showDirectoryPicker({mode:'read'});await loadFromDirectoryHandle(d);}catch(e){if(e.name!=='AbortError')document.getElementById('photo-folder-input').click();}}else{document.getElementById('photo-folder-input').click();}}
async function loadFromDirectoryHandle(dirHandle){S.photoObjUrls.forEach(u=>URL.revokeObjectURL(u));S.photoObjUrls=[];const newMap={};const IMG=new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);for await(const[name,handle] of dirHandle.entries()){if(handle.kind==='file'){const ext=name.split('.').pop().toLowerCase();if(IMG.has(ext)){const f=await handle.getFile();const k=name.replace(/\.[^.]+$/,'').toLowerCase().trim();const u=URL.createObjectURL(f);newMap[k]=u;S.photoObjUrls.push(u);}}}S.photoMap=newMap;updatePhotoUI();if(S.viewData.length)renderChart();}
function loadFromFileInput(files){S.photoObjUrls.forEach(u=>URL.revokeObjectURL(u));S.photoObjUrls=[];const newMap={};const IMG=new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);Array.from(files).forEach(file=>{const ext=file.name.split('.').pop().toLowerCase();if(IMG.has(ext)){const k=file.name.replace(/\.[^.]+$/,'').toLowerCase().trim();const u=URL.createObjectURL(file);newMap[k]=u;S.photoObjUrls.push(u);}});S.photoMap=newMap;updatePhotoUI();if(S.viewData.length)renderChart();}
function updatePhotoUI(){const count=Object.keys(S.photoMap).length;document.getElementById('photo-btn').classList.toggle('loaded',count>0);document.getElementById('photo-btn-label').textContent=count>0?'Photos':'Load Photos';const badge=document.getElementById('photo-count');badge.textContent=count;badge.style.display=count>0?'':'none';const stat=document.getElementById('stat-photos');if(stat){stat.style.display=count>0?'flex':'none';document.getElementById('stat-photos-val').textContent=count;}}
function getPhotoUrl(node){if(!Object.keys(S.photoMap).length)return '';const col=S.photoMatchCol||S.colMap.empId;const val=String(node[col]||'').toLowerCase().trim();return val&&S.photoMap[val]?S.photoMap[val]:'';}

function buildMapScreen(){document.getElementById('col-count').textContent=S.columns.length;document.getElementById('detected-columns').innerHTML=S.columns.map(c=>'<div class="col-chip">'+esc(c)+(S.colSamples[c].length?'<span class="chip-sample">'+esc(S.colSamples[c].join(', '))+'</span>':'')+'</div>').join('');const blank='<option value="">— select —</option>';const opts=blank+S.columns.map(c=>'<option value="'+esc(c)+'">'+esc(c)+'</option>').join('');['empId','empName','managerId','froId'].forEach(k=>{const sel=document.getElementById('map-'+k);if(!sel)return;sel.innerHTML=opts;sel.value=S.colMap[k]||'';});const wrap=document.getElementById('data-preview-wrap');const preview=S.rawRows.slice(0,3);if(preview.length){let html='<table class="data-preview-table"><thead><tr>'+S.columns.map(c=>'<th>'+esc(c)+'</th>').join('')+'</tr></thead><tbody>';preview.forEach(row=>{html+='<tr>'+S.columns.map(c=>'<td>'+esc(String(row[c]||'').substring(0,22))+'</td>').join('')+'</tr>';});wrap.innerHTML=html+'</tbody></table>';}else{wrap.innerHTML='';}setTimeout(onMapChange,30);}

function confirmColumnMap(){S.colMap.empId=document.getElementById('map-empId').value;S.colMap.empName=document.getElementById('map-empName').value;S.colMap.managerId=document.getElementById('map-managerId').value;S.colMap.froId=document.getElementById('map-froId').value;if(!S.colMap.empId||!S.colMap.empName){alert('Please map Employee ID and Employee Name.');return;}if(S.colMap.empId===S.colMap.empName){alert('Employee ID and Employee Name must be different columns.');return;}if(S.colMap.managerId&&S.colMap.managerId===S.colMap.empId){alert('Manager ID and Employee ID must be different columns.');return;}if(S.colMap.froId&&(S.colMap.froId===S.colMap.empId||S.colMap.froId===S.colMap.empName||S.colMap.froId===S.colMap.managerId)){alert('FRO column must be different from the other roles.');return;}buildCardScreen();goTo('card');}

const AUTO_FIELDS=[{id:'__auto_reports__',icon:'📊',label:'Direct Reports',desc:'Count of direct reports'},{id:'__auto_teamsize__',icon:'👥',label:'Total Team Size',desc:'All descendants count'}];
function buildCardScreen(){const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId,S.colMap.froId].filter(Boolean));const available=S.columns.filter(c=>!core.has(c));document.getElementById('card-fields-panel').innerHTML='<div class="fields-section"><div class="fields-section-label">Column Fields</div>'+(available.length?available.map(f=>'<div class="field-chip" draggable="true" data-field="'+esc(f)+'" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)"><span class="drag-icon">⠿</span>'+esc(f)+'</div>').join(''):'<div style="font-size:0.78rem;color:var(--text3);font-style:italic">No extra columns</div>')+'</div><div class="fields-section"><div class="fields-section-label">Auto-Calculated</div>'+AUTO_FIELDS.map(f=>'<div class="field-chip" draggable="true" data-field="'+f.id+'" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)" title="'+f.desc+'"><span class="drag-icon">⠿</span>'+f.icon+' '+f.label+'</div>').join('')+'</div>';if(!S.cardSlots.f3)S.cardSlots.f3='__auto_reports__';const COLORS=['#4f46e5','#7c3aed','#db2777','#dc2626','#d97706','#059669','#0891b2','#0284c7','#374151','#0f172a'];document.getElementById('color-palette').innerHTML=COLORS.map(c=>'<div class="color-swatch'+(S.cardAccent===c?' selected':'')+'" style="background:'+c+'" onclick="setCardAccent(\''+c+'\')"></div>').join('');const empColSel=document.getElementById('emp-type-col');if(empColSel){empColSel.innerHTML='<option value="">Select column...</option>'+S.columns.filter(c=>!core.has(c)).map(c=>'<option value="'+esc(c)+'"'+(S.empTypeCol===c?' selected':'')+'>'+esc(c)+'</option>').join('');if(S.empTypeCol)populateEmpTypeValues(S.empTypeCol);}renderCardPreview();syncChipStates();}
function onEmpTypeColChange(){S.empTypeCol=document.getElementById('emp-type-col').value;if(S.empTypeCol){populateEmpTypeValues(S.empTypeCol);}else{document.getElementById('emp-type-rows').style.display='none';}if(S.viewData.length)renderChart();persistState();}
function populateEmpTypeValues(col){const vals=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();const rows=document.getElementById('emp-type-rows');rows.style.display='';['active','vacant','resigned'].forEach(key=>{const sel=document.getElementById('emp-val-'+key);sel.innerHTML='<option value="">Value...</option>'+vals.map(v=>'<option value="'+esc(v)+'"'+(S.empTypeLabels[key]===v?' selected':'')+'>'+esc(v)+'</option>').join('');sel.onchange=()=>{S.empTypeLabels[key]=sel.value;buildEmpTypeMap();if(S.viewData.length)renderChart();persistState();};const colorInput=document.getElementById('emp-color-'+key);colorInput.value=S.empTypeColors[key];colorInput.oninput=()=>{S.empTypeColors[key]=colorInput.value;buildEmpTypeMap();if(S.viewData.length)renderChart();persistState();};});buildEmpTypeMap();}
function buildEmpTypeMap(){S.empTypeMap={};['active','vacant','resigned'].forEach(key=>{const v=S.empTypeLabels[key];if(v)S.empTypeMap[v]=S.empTypeColors[key];});}
function getNodeBorderColor(node){if(S.empTypeCol&&S.empTypeMap){const val=String(node[S.empTypeCol]||'').trim();if(S.empTypeMap[val])return S.empTypeMap[val];}return S.cardAccent;}
function onDragStart(e){S.draggingField=e.currentTarget.dataset.field;e.currentTarget.classList.add('dragging');e.dataTransfer.effectAllowed='move';}
function onDragEnd(e){e.currentTarget.classList.remove('dragging');S.draggingField=null;}
function onZoneDragOver(e){e.preventDefault();e.currentTarget.classList.add('drop-target');}
function onZoneDragLeave(e){e.currentTarget.classList.remove('drop-target');}
function onZoneDrop(e,zone){e.preventDefault();e.currentTarget.classList.remove('drop-target');if(!S.draggingField)return;Object.keys(S.cardSlots).forEach(z=>{if(S.cardSlots[z]===S.draggingField)S.cardSlots[z]='';});S.cardSlots[zone]=S.draggingField;S.draggingField=null;renderCardPreview();syncChipStates();}
function clearZone(zone){S.cardSlots[zone]='';renderCardPreview();syncChipStates();}
function syncChipStates(){const placed=new Set(Object.values(S.cardSlots).filter(Boolean));document.querySelectorAll('.field-chip').forEach(c=>c.classList.toggle('placed',placed.has(c.dataset.field)));}
function fieldLabel(id){if(!id)return '';const af=AUTO_FIELDS.find(f=>f.id===id);if(af)return af.icon+' '+af.label;return id;}
function fieldSampleVal(id){if(!id)return '';if(id==='__auto_reports__')return '12';if(id==='__auto_teamsize__')return '48';const row=S.rawRows.find(r=>r[id])||S.rawRows[0]||{};return String(row[id]||'Sample').substring(0,18);}
function zoneHtml(zoneId,placeholder){const dA='ondragover="onZoneDragOver(event)" ondragleave="onZoneDragLeave(event)" ondrop="onZoneDrop(event,\''+zoneId+'\')"';const v=S.cardSlots[zoneId];if(v)return '<div class="card-zone filled" '+dA+'><span class="zone-field">'+esc(fieldLabel(v))+'</span><span class="zone-val">'+esc(fieldSampleVal(v))+'</span><span class="zone-remove" onclick="clearZone(\''+zoneId+'\')">✕</span></div>';return '<div class="card-zone" '+dA+'><span class="zone-ph">'+placeholder+'</span></div>';}
function renderCardPreview(){const sampleRow=S.rawRows.find(r=>r[S.colMap.empName])||S.rawRows[0]||{};const sampleName=String(sampleRow[S.colMap.empName]||'Employee Name').substring(0,26);const ac=S.cardAccent;const ps=S.photoSize,pr=getPhotoRadius();const photoDiv='<div style="width:'+ps+'px;height:'+ps+'px;border-radius:'+pr+';background:linear-gradient(150deg,'+ac+'18,'+ac+'30);color:'+ac+';font-size:'+Math.round(ps*0.28)+'px;font-weight:800;display:flex;align-items:center;justify-content:center;border:3px solid '+ac+'55;flex-shrink:0">AB</div>';const b1ZoneHtml='<div class="preview-b1-zone">'+zoneHtml('b1','Body slot')+'</div>';const nameBlock='<div style="width:100%;text-align:center"><div style="font-size:0.88rem;font-weight:800;color:var(--text);margin-bottom:4px">🔒 '+esc(sampleName)+'</div>'+b1ZoneHtml+'</div>';let bodyHtml;const pl=S.photoPlacement;if(pl==='none'){bodyHtml='<div style="display:flex;flex-direction:column;gap:6px">'+nameBlock+'</div>';}else if(pl==='top'){bodyHtml='<div style="display:flex;flex-direction:column;align-items:center;gap:10px">'+photoDiv+nameBlock+'</div>';}else if(pl==='left'){bodyHtml='<div style="display:flex;flex-direction:row;align-items:flex-start;gap:10px">'+photoDiv+'<div style="flex:1;min-width:0">'+nameBlock+'</div></div>';}else{bodyHtml='<div style="display:flex;flex-direction:row-reverse;align-items:flex-start;gap:10px">'+photoDiv+'<div style="flex:1;min-width:0">'+nameBlock+'</div></div>';}document.getElementById('card-preview').innerHTML='<div class="preview-card" style="border-color:'+ac+'"><div class="preview-card-header">'+zoneHtml('h1','H1')+zoneHtml('h2','H2')+zoneHtml('h3','H3')+'</div><div class="preview-card-body">'+bodyHtml+'</div><div class="preview-card-footer">'+zoneHtml('f1','F1')+zoneHtml('f2','F2')+zoneHtml('f3','F3')+'</div></div>';}
function setCardAccent(color){S.cardAccent=color;document.querySelectorAll('.color-swatch').forEach(s=>s.classList.toggle('selected',s.style.background===color));renderCardPreview();if(S.viewData.length)renderChart();persistState();}
function setPhotoShape(shape){S.photoShape=shape;document.querySelectorAll('.shape-btn').forEach(b=>b.classList.toggle('selected',b.dataset.shape===shape));renderCardPreview();if(S.viewData.length)renderChart();persistState();}
function setPhotoPlacement(p){S.photoPlacement=p;document.querySelectorAll('[data-placement]').forEach(b=>b.classList.toggle('selected',b.dataset.placement===p));renderCardPreview();if(S.viewData.length)renderChart();persistState();}
function getPhotoRadius(){if(S.photoShape==='circle')return'50%';if(S.photoShape==='rounded')return'12px';return'4px';}
function confirmCardDesign(){buildEmpTypeMap();buildFilterScreen();goTo('filter');}
function buildFilterScreen(){const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId,S.colMap.froId].filter(Boolean));const filterable=S.columns.filter(c=>!core.has(c));const cc=document.getElementById('filter-chip-picker');cc.innerHTML=filterable.map(col=>'<div class="filter-chip '+(S.filterCols.includes(col)?'selected':'')+'" data-col="'+esc(col)+'">'+esc(col)+'</div>').join('');cc.onclick=function(e){const chip=e.target.closest('.filter-chip');if(!chip)return;const col=chip.dataset.col;if(col)toggleFilterCol(col);};renderFilterPreview();}
function toggleFilterCol(col){if(S.filterCols.includes(col))S.filterCols=S.filterCols.filter(c=>c!==col);else if(S.filterCols.length<3)S.filterCols.push(col);else{S.filterCols.shift();S.filterCols.push(col);}document.querySelectorAll('.filter-chip').forEach(c=>c.classList.toggle('selected',S.filterCols.includes(c.dataset.col)));renderFilterPreview();}
function renderFilterPreview(){document.getElementById('filter-counter').textContent=S.filterCols.length+' of 3 filters selected';const area=document.getElementById('filter-preview-area');if(!S.filterCols.length){area.innerHTML='<div style="font-size:0.82rem;color:var(--text3);padding:12px 0">No filters — full chart will display.</div>';return;}area.innerHTML='<div class="filter-preview-box">'+S.filterCols.map((col,i)=>{const isLast=i===S.filterCols.length-1;const vals=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort().slice(0,10);return '<div class="fpr-row"><span class="fpr-col">'+esc(col)+(isLast?' <span style="background:var(--accent);color:#fff;border-radius:999px;padding:1px 7px;font-size:0.58rem;font-weight:700;margin-left:4px">Export All</span>':'')+'</span><div class="fpr-vals">'+vals.map(v=>'<span class="fv-pill">'+esc(v)+'</span>').join('')+(vals.length>=10?'<span style="font-size:0.7rem;color:var(--text3)">+ more</span>':'')+'</div></div>';}).join('')+'</div>';}
function launchChart(){S.activeFilters={};S.skipDepth=0;buildViewData();buildFilterBar();renderChart();goTo('chart');setTimeout(refreshDataQualityBtn,200);persistState();}

function populateSummaryFields(){const core=new Set([S.colMap.empId,S.colMap.empName,S.colMap.managerId,S.colMap.froId].filter(Boolean));const opts='<option value="">—</option><option value="__name__">Name</option>'+S.columns.filter(c=>!core.has(c)).map(c=>'<option value="'+esc(c)+'">'+esc(c)+'</option>').join('');['summary-field1','summary-field2','summary-field3'].forEach((id,i)=>{const el=document.getElementById(id);if(!el)return;el.innerHTML=opts;const v=[S.summaryField1,S.summaryField2,S.summaryField3][i];if(v)el.value=v;});document.getElementById('depth-select').value=S.skipDepth;populatePhotoMatchCol();}
function populatePhotoMatchCol(){const sel=document.getElementById('photo-match-col');if(!sel)return;sel.innerHTML=S.columns.map(c=>'<option value="'+esc(c)+'"'+(c===S.colMap.empId?' selected':'')+'>'+esc(c)+'</option>').join('');if(!S.photoMatchCol)S.photoMatchCol=S.colMap.empId;sel.value=S.photoMatchCol||S.colMap.empId;
  // show FRO legend if FRO column is mapped
  const froLegend=document.getElementById('fro-legend');
  if(froLegend)froLegend.style.display=S.colMap.froId?'inline-flex':'none';
}
function toggleManagerMode(){pushUndo();S.managerMode=!S.managerMode;document.getElementById('mgr-mode-btn').classList.toggle('active',S.managerMode);document.getElementById('summary-fields-wrap').style.display=S.managerMode?'flex':'none';const stat=document.getElementById('stat-mgr-mode');if(stat)stat.style.display=S.managerMode?'flex':'none';renderChart();persistState();}
function isManager(nodeId){return(S.childMap[nodeId]||[]).length>0;}

function buildViewData(){const{empId,empName,managerId}=S.colMap;let nodes=S.rawRows.map(row=>{const id=String(row[empId]||'').replace(/\.0$/,'').trim();const mgr=managerId?String(row[managerId]||'').replace(/\.0$/,'').trim():'';const node={id,name:String(row[empName]||'Unknown'),manager:mgr};S.columns.forEach(col=>{node[col]=String(row[col]||'');});return node;}).filter(n=>n.id&&!S.removedIds.has(n.id));const validIds=new Set(nodes.map(n=>n.id));nodes.forEach(n=>{if(S.managerOverrides.hasOwnProperty(n.id))n.manager=S.managerOverrides[n.id];});nodes.forEach(n=>{if(!validIds.has(n.manager)||n.manager===n.id)n.manager='';});const hasFilter=Object.values(S.activeFilters).some(v=>v);if(hasFilter){const matching=new Set(nodes.filter(n=>Object.entries(S.activeFilters).every(([c,v])=>!v||n[c]===v)).map(n=>n.id));const byId=Object.fromEntries(nodes.map(n=>[n.id,n]));const keep=new Set(matching);matching.forEach(id=>{let cur=byId[id];const vis=new Set();while(cur&&cur.manager&&byId[cur.manager]&&!vis.has(cur.id)){vis.add(cur.id);keep.add(cur.manager);cur=byId[cur.manager];}});nodes=nodes.filter(n=>keep.has(n.id));}S.viewData=nodes;S.childMap={};nodes.forEach(n=>{if(!S.childMap[n.manager])S.childMap[n.manager]=[];S.childMap[n.manager].push(n);});S.descCount={};function calcD(id,vis){if(vis.has(id))return 0;vis.add(id);if(S.descCount[id]!==undefined)return S.descCount[id];const kids=S.childMap[id]||[];S.descCount[id]=kids.reduce((s,k)=>s+1+calcD(k.id,vis),0);return S.descCount[id];}nodes.filter(n=>!n.manager).forEach(r=>calcD(r.id,new Set()));S.nodeHeight={};function calcH(id,vis){if(vis.has(id))return 0;vis.add(id);if(S.nodeHeight[id]!==undefined)return S.nodeHeight[id];const kids=S.childMap[id]||[];S.nodeHeight[id]=kids.length?1+Math.max(...kids.map(k=>calcH(k.id,vis))):0;return S.nodeHeight[id];}nodes.filter(n=>!n.manager).forEach(r=>calcH(r.id,new Set()));nodes.forEach(n=>{if(S.nodeHeight[n.id]===undefined)calcH(n.id,new Set());});S.nodeDepth={};function calcDepth(id,d,vis){if(vis.has(id))return;vis.add(id);S.nodeDepth[id]=d;(S.childMap[id]||[]).forEach(k=>calcDepth(k.id,d+1,vis));}nodes.filter(n=>!n.manager).forEach(r=>calcDepth(r.id,0,new Set()));nodes.forEach(n=>{if(S.nodeDepth[n.id]===undefined)S.nodeDepth[n.id]=0;});if(typeof refreshDataQualityBtn==='function')setTimeout(refreshDataQualityBtn,30);}
function childrenOf(id){return S.childMap[id]||[];}
function countDescendants(id){return S.descCount[id]||0;}

function buildFilterBar(){const bar=document.getElementById('filter-bar');if(!S.filterCols.length){bar.style.display='none';return;}bar.style.display='flex';const allVals={};S.filterCols.forEach(col=>{allVals[col]=[...new Set(S.rawRows.map(r=>String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();});bar.innerHTML='<span style="font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);flex-shrink:0">Filters</span>'+S.filterCols.map(col=>'<div class="filter-dropdown-wrap"><span class="filter-dropdown-label">'+esc(col)+'</span><select class="filter-dropdown" data-filter-col="'+esc(col)+'"><option value="">All '+esc(col)+'</option>'+allVals[col].map(v=>'<option value="'+esc(v)+'"'+(S.activeFilters[col]===v?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select></div>').join('')+(Object.values(S.activeFilters).some(v=>v)?'<button class="btn btn-ghost btn-sm" onclick="clearAllFilters()" style="margin-left:auto">Clear All</button>':'');bar.querySelectorAll('.filter-dropdown').forEach(sel=>{sel.addEventListener('change',function(){applyFilter(this.dataset.filterCol,this.value);});});}
function applyFilter(col,val){pushUndo();if(val)S.activeFilters[col]=val;else delete S.activeFilters[col];requestAnimationFrame(()=>setTimeout(()=>{buildViewData();renderChart();buildFilterBar();persistState();},0));}
function clearAllFilters(){pushUndo();S.activeFilters={};requestAnimationFrame(()=>setTimeout(()=>{buildViewData();renderChart();buildFilterBar();persistState();},0));}
function setSkipDepth(n){pushUndo();S.skipDepth=n;const ds=document.getElementById('depth-select');if(ds)ds.value=n;renderChart();persistState();}
function setChartBg(color){S.chartBgColor=color;if(!S.transparentExport)applyChartBg();persistState();}
function toggleTransparent(){S.transparentExport=!S.transparentExport;const btn=document.getElementById('bg-transparent-btn');const inp=document.getElementById('bg-color-input');if(btn)btn.classList.toggle('active',S.transparentExport);if(inp)inp.disabled=S.transparentExport;applyChartBg();persistState();}
function applyChartBg(){const wrap=document.getElementById('chart-canvas-wrap');if(!wrap)return;if(S.transparentExport){wrap.classList.add('transparent-preview');wrap.style.background='';}else{wrap.classList.remove('transparent-preview');wrap.style.background=S.chartBgColor;}}
function getSlotVal(node,slot){const f=S.cardSlots[slot];if(!f)return '';if(f==='__auto_reports__')return childrenOf(node.id).length+' reports';if(f==='__auto_teamsize__')return countDescendants(node.id)+' people';return String(node[f]||'').substring(0,28);}

/* ── mkKidsWrap: groups children into rows. Row size driven by S.maxPerRow
       (number 2-12 or 'auto' which picks a near-square layout). ── */
function mkKidsWrap(kids,depth){
  const MAX=getEffectiveMaxPerRow(kids.length);
  if(kids.length<=MAX){
    const ul=document.createElement('ul');
    kids.forEach(k=>ul.appendChild(mkNodeLI(k,depth+1)));
    return ul;
  }
  const wrap=document.createElement('div');
  wrap.className='children-rows-wrap';
  let chunkIdx=0;
  for(let i=0;i<kids.length;i+=MAX){
    const chunk=kids.slice(i,i+MAX);
    const ul=document.createElement('ul');
    ul.className='children-row-ul'+(chunkIdx>0?' row-cont':'');
    chunk.forEach(k=>ul.appendChild(mkNodeLI(k,depth+1)));
    wrap.appendChild(ul);
    chunkIdx++;
  }
  return wrap;
}

function renderChart(){
  const tree=document.getElementById('org-tree');
  tree.innerHTML='';
  tree.style.width='';
  const ds=document.getElementById('depth-select');if(ds)ds.value=S.skipDepth;
  let roots;
  if(S.skipDepth>0){roots=S.viewData.filter(n=>(S.nodeDepth[n.id]||0)===S.skipDepth);}
  else{roots=S.childMap['']||[];}
  if(!roots.length){tree.innerHTML='<div class="no-data">No nodes found. Try a lower Skip Top value.</div>';updateStats(roots);return;}
  const ul=document.createElement('ul');
  roots.forEach(r=>ul.appendChild(mkNodeLI(r,0)));
  tree.appendChild(ul);
  updateStats(roots);
  clearTimeout(window._fit);window._fit=setTimeout(()=>fitToScreen(true),180);
  clearTimeout(window._froTimer);window._froTimer=setTimeout(renderFROLines,600);
  if(S.gridMode){setTimeout(()=>{applyGridOverridesToTree();redrawGridConnectorsFromTree();applyGridLines();},220);}
}
function mkNodeLI(node,depth){depth=depth||0;const li=document.createElement('li');li.dataset.id=node.id;const ac=getNodeBorderColor(node);const acLight=ac+'18',acMid=ac+'55';const kids=childrenOf(node.id);const card=document.createElement('div');card.className='node-card'+(node.id===S.highlighted?' highlighted':'');
  // ── FULL 4-side border color + per-card accent variable ──
  card.style.borderColor=ac;
  card.style.setProperty('--card-accent',ac);
  // Drag handlers are always attached. They check context at runtime —
  // tree-mode drag in PV is blocked, but Grid-mode drag (free-move) works in
  // both main chart and Person View.
  card.draggable=!(S.pvMode?S.pvGridMode:S.gridMode);
  card.dataset.dragId=node.id;
  card.setAttribute('aria-grabbed','false');
  card.addEventListener('dragstart',onCardDragStart);
  card.addEventListener('dragover',onCardDragOver);
  card.addEventListener('dragleave',onCardDragLeave);
  card.addEventListener('drop',onCardDrop);
  card.addEventListener('pointerdown',onCardPointerDown);
  card.addEventListener('dragend',onCardDragEnd);
  card.addEventListener('click',function(e){
    if(e.shiftKey){e.preventDefault();e.stopPropagation();toggleSelectCard(node.id,card);}
  });
  if(S.selectedIds&&S.selectedIds.has(node.id))card.classList.add('selected');
  const h1=getSlotVal(node,'h1'),h2=getSlotVal(node,'h2'),h3=getSlotVal(node,'h3');const f1=getSlotVal(node,'f1'),f2=getSlotVal(node,'f2'),f3=getSlotVal(node,'f3')||node.id.substring(0,14);const b1=getSlotVal(node,'b1');const subtitle=h2;const ps=S.photoSize,pr=getPhotoRadius(),pfs=Math.round(ps*0.28)+'px';const pInline='width:'+ps+'px;height:'+ps+'px;border-radius:'+pr+';';const initials=node.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();const photoUrl=getPhotoUrl(node);let photoHtml='';if(photoUrl){photoHtml='<img class="ncard-photo" src="'+esc(photoUrl)+'" crossorigin="anonymous" style="'+pInline+'border:3px solid '+acMid+';box-shadow:0 8px 24px '+ac+'66" onerror="this.onerror=null;this.style.display=\'none\';this.nextElementSibling.style.display=\'flex\'"><div class="ncard-photo-fallback" style="display:none;'+pInline+'font-size:'+pfs+';background:linear-gradient(150deg,'+acLight+','+ac+'28);color:'+ac+';border:3px solid '+acMid+';">'+esc(initials)+'</div>';}else if(Object.keys(S.photoMap).length>0){photoHtml='<div class="ncard-photo-fallback" style="display:flex;'+pInline+'font-size:'+pfs+';background:linear-gradient(150deg,'+acLight+','+ac+'28);color:'+ac+';border:3px solid '+acMid+';">'+esc(initials)+'</div>';}
  const b1row=b1?'<div class="ncard-body-b1">'+esc(b1)+'</div>':'';const textBlock='<div class="ncard-text-wrap"><div class="ncard-name">'+esc(node.name)+'</div>'+(subtitle?'<div class="ncard-sub">'+esc(subtitle)+'</div>':'')+b1row+'</div>';let bodyHtml;const pl=S.photoPlacement;if(!photoHtml||pl==='none'){bodyHtml='<div class="ncard-body-inner" style="flex-direction:column">'+textBlock+'</div>';}else if(pl==='top'){bodyHtml='<div class="ncard-body-inner" style="flex-direction:column;align-items:center"><div style="flex-shrink:0">'+photoHtml+'</div>'+textBlock+'</div>';}else if(pl==='left'){bodyHtml='<div class="ncard-body-inner" style="flex-direction:row;align-items:flex-start"><div style="flex-shrink:0">'+photoHtml+'</div><div style="flex:1;min-width:0">'+textBlock+'</div></div>';}else{bodyHtml='<div class="ncard-body-inner" style="flex-direction:row-reverse;align-items:flex-start"><div style="flex-shrink:0">'+photoHtml+'</div><div style="flex:1;min-width:0">'+textBlock+'</div></div>';}
  const pvBtns=S.pvMode?'':
    '<div class="ncard-export-btn" onclick="exportSubtree(event,\''+esc(node.id)+'\')" style="position:absolute;top:6px;right:30px;width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:6px;font-size:0.6rem;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 0.15s;z-index:8">📸</div>'+
    '<div class="ncard-edit-btn" onclick="openReassignModal(event,\''+esc(node.id)+'\')" style="position:absolute;top:6px;right:6px;width:22px;height:22px;background:var(--bg);border:1.5px solid var(--border2);border-radius:6px;font-size:0.65rem;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 0.15s;z-index:8">✎</div>';
  card.innerHTML='<div class="ncard-header" style="background:'+acLight+';border-bottom-color:'+ac+'33"><span class="ncard-slot'+(h1?' has-val':'')+'" title="'+esc(h1)+'">'+(esc(h1)||'—')+'</span><span class="ncard-slot'+(h2?' has-val':'')+'" title="'+esc(h2)+'">'+(esc(h2)||'—')+'</span><span class="ncard-slot'+(h3?' has-val':'')+'" title="'+esc(h3)+'">'+(esc(h3)||'—')+'</span></div><div class="ncard-body">'+bodyHtml+'</div><div class="ncard-footer" style="background:'+acLight+';border-top-color:'+ac+'33"><span class="ncard-slot'+(f1?' has-val':'')+'" title="'+esc(f1)+'">'+(esc(f1)||'—')+'</span><span class="ncard-slot'+(f2?' has-val':'')+'" title="'+esc(f2)+'">'+(esc(f2)||'—')+'</span><span class="ncard-slot'+(f3?' has-val':'')+'" title="'+esc(f3)+'">'+(esc(f3)||node.id.substring(0,14))+'</span></div>'+pvBtns;
  if(!S.pvMode){card.querySelectorAll('.ncard-edit-btn,.ncard-export-btn').forEach(b=>{card.addEventListener('mouseenter',()=>b.style.opacity='1');card.addEventListener('mouseleave',()=>b.style.opacity='0');});}
  if(kids.length){const cb=document.createElement('div');cb.className='collapse-btn';cb.innerHTML='▾';cb.title='Collapse / expand';cb.addEventListener('click',e=>{e.stopPropagation();toggleCollapse(li,cb);});card.appendChild(cb);}
  li.appendChild(card);
  if(kids.length){
    if(S.managerMode){
      const managerKids=kids.filter(k=>isManager(k.id));
      const leafKids=kids.filter(k=>!isManager(k.id));
      if(managerKids.length===0&&leafKids.length>0){
        const ul=document.createElement('ul');ul.appendChild(mkLeafSummaryLI(leafKids,ac));li.appendChild(ul);
      }else if(managerKids.length>0){
        const wrap=mkKidsWrap(managerKids,depth);
        if(leafKids.length>0){const lastUl=wrap.tagName==='UL'?wrap:(wrap.querySelector('ul:last-of-type')||wrap);lastUl.appendChild(mkLeafSummaryLI(leafKids,ac));}
        li.appendChild(wrap);
      }
    }else{
      li.appendChild(mkKidsWrap(kids,depth));
    }
  }
  return li;
}

function mkLeafSummaryLI(leafNodes, ac) {
  const li = document.createElement('li');
  const f1 = S.summaryField1, f2 = S.summaryField2, f3 = S.summaryField3;
  const count = leafNodes.length;
  const AV = 28; const PAD_H = 14; const PAD_V = 8; const TEXT_LH = 16; const GAP = 10; const HEADER_LBL_H = 22;
  const FF = "font-family:'Plus Jakarta Sans',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;";
  // IC summary now uses the parent card's accent (ac) — was hardcoded purple before
  const headerBg = ac + '14';        // 8% tint
  const headerBorder = ac + '40';    // 25% tint
  const headerColor = ac;            // full
  const headerHtml =
    '<div style="position:relative;background:' + headerBg + ';border-bottom:1px solid ' + headerBorder + ';padding:9px ' + PAD_H + 'px;' + FF + '">' +
      '<div style="height:' + HEADER_LBL_H + 'px;line-height:' + HEADER_LBL_H + 'px;font-size:11px;font-weight:800;color:' + headerColor + ';text-transform:uppercase;letter-spacing:0.05em;padding-right:42px;white-space:nowrap;overflow:hidden;' + FF + '">ICs (' + count + ')</div>' +
      '<div style="position:absolute;top:8px;right:' + PAD_H + 'px;height:24px;line-height:24px;background:' + ac + ';color:#ffffff;border-radius:999px;padding:0 10px;font-size:10px;font-weight:800;text-align:center;' + FF + '">' + count + '</div>' +
    '</div>';
  let rowsHtml = '';
  leafNodes.forEach((n, idx) => {
    const initials = n.name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase();
    const borderC = getNodeBorderColor(n);
    const photoUrl = getPhotoUrl(n);
    const isLast = idx === leafNodes.length - 1;
    const nameVal = n.name.substring(0, 24);
    const f1IsName = (f1 === '__name__');
    const primaryVal = f1 ? (f1IsName ? nameVal : (String(n[f1] || '').trim() || nameVal).substring(0, 24)) : nameVal;
    const showNameSub = f1 && !f1IsName && primaryVal !== nameVal;
    const val2 = f2 ? (f2 === '__name__' ? n.name.substring(0, 22) : String(n[f2] || '').substring(0, 22)) : '';
    const val3 = f3 ? (f3 === '__name__' ? n.name.substring(0, 22) : String(n[f3] || '').substring(0, 22)) : '';
    const numLines = 1 + (showNameSub ? 1 : 0) + (val2 ? 1 : 0) + (val3 ? 1 : 0);
    const textTotalH = numLines * TEXT_LH;
    const innerH = Math.max(AV, textTotalH);
    const totalRowH = innerH + PAD_V * 2;
    const avatarTopY = PAD_V + Math.max(0, Math.round((innerH - AV) / 2));
    const textTopPad = PAD_V + Math.max(0, Math.round((innerH - textTotalH) / 2));
    const textLeftMargin = PAD_H + AV + GAP;
    let avatarHtml;
    if (photoUrl) {
      avatarHtml = '<img src="' + esc(photoUrl) + '" crossorigin="anonymous" style="position:absolute;left:' + PAD_H + 'px;top:' + avatarTopY + 'px;width:' + AV + 'px;height:' + AV + 'px;border-radius:7px;object-fit:cover;object-position:center top;border:2px solid ' + borderC + '55;box-sizing:border-box;display:block;">';
    } else {
      const innerLH = AV - 4;
      avatarHtml = '<div style="position:absolute;left:' + PAD_H + 'px;top:' + avatarTopY + 'px;width:' + AV + 'px;height:' + AV + 'px;border-radius:7px;background:' + borderC + '1f;color:' + borderC + ';border:2px solid ' + borderC + '55;box-sizing:border-box;text-align:center;font-size:12px;font-weight:800;line-height:' + innerLH + 'px;' + FF + '">' + esc(initials) + '</div>';
    }
    let textLines = '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:12px;font-weight:700;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;' + FF + '">' + esc(primaryVal) + '</div>';
    if (showNameSub) { textLines += '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:10px;color:#475569;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;' + FF + '">' + esc(nameVal) + '</div>'; }
    if (val2) { textLines += '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:10px;color:#94a3b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;' + FF + '">' + esc(val2) + '</div>'; }
    if (val3) { textLines += '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:10px;color:#94a3b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;' + FF + '">' + esc(val3) + '</div>'; }
    const rowBorder = isLast ? '' : 'border-bottom:1px solid #e2e8f0;';
    rowsHtml += '<div style="position:relative;background:#ffffff;height:' + totalRowH + 'px;' + rowBorder + 'box-sizing:border-box;">' + avatarHtml + '<div style="margin-left:' + textLeftMargin + 'px;padding-top:' + textTopPad + 'px;padding-right:' + PAD_H + 'px;">' + textLines + '</div></div>';
  });
  const card = document.createElement('div');
  card.className = 'summary-list-card';
  card.style.borderTopColor = ac;
  card.dataset.summaryParent = 'ic';
  // Synthetic ID so this IC summary block can be dragged + selected like a regular card
  // in grid mode. We store the manager IDs it belongs to so connectors can be drawn.
  const managerOfICs = (leafNodes[0] && leafNodes[0].manager) || '';
  const synthId = '__ic__' + (managerOfICs || ('row-' + Math.random().toString(36).slice(2,8)));
  card.dataset.dragId = synthId;
  card.dataset.icManager = managerOfICs;
  card.dataset.icCount = String(leafNodes.length);
  card.draggable = !(S.pvMode ? S.pvGridMode : S.gridMode);
  card.addEventListener('dragstart', onCardDragStart);
  card.addEventListener('dragover', onCardDragOver);
  card.addEventListener('dragleave', onCardDragLeave);
  card.addEventListener('drop', onCardDrop);
  card.addEventListener('dragend', onCardDragEnd);
  card.addEventListener('pointerdown', onCardPointerDown);
  card.addEventListener('click', function(e){
    if (e.shiftKey){e.preventDefault();e.stopPropagation();toggleSelectCard(synthId, card);}
  });
  if (S.selectedIds && S.selectedIds.has(synthId)) card.classList.add('selected');
  card.innerHTML = headerHtml + rowsHtml;
  li.appendChild(card);
  return li;
}

function toggleCollapse(li,btn){li.classList.toggle('collapsed');const c=li.classList.contains('collapsed');const childEl=li.querySelector(':scope > ul,:scope > .children-rows-wrap');if(childEl)childEl.style.display=c?'none':'';btn.innerHTML=c?'▸':'▾';btn.style.color=c?'var(--warning)':'';li.querySelector('.node-card').classList.toggle('collapsed-node',c);setTimeout(()=>updateStats(),60);}
function expandAll(){document.querySelectorAll('li.collapsed').forEach(li=>{li.classList.remove('collapsed');const u=li.querySelector(':scope > ul,:scope > .children-rows-wrap');if(u)u.style.display='';const card=li.querySelector('.node-card');if(card)card.classList.remove('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▾';b.style.color='';}});setTimeout(()=>updateStats(),60);}
function collapseAll(){document.querySelectorAll('li').forEach(li=>{if(!li.parentElement||!li.parentElement.parentElement||!li.parentElement.parentElement.closest('li'))return;const hasKids=li.querySelector(':scope > ul,:scope > .children-rows-wrap');if(hasKids){li.classList.add('collapsed');hasKids.style.display='none';const card=li.querySelector('.node-card');if(card)card.classList.add('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▸';b.style.color='var(--warning)';};}});setTimeout(()=>updateStats(),60);}
function updateStats(roots){if(!roots)roots=S.skipDepth>0?S.viewData.filter(n=>(S.nodeDepth[n.id]||0)===S.skipDepth):(S.childMap['']||[]);document.getElementById('stat-total').textContent=S.viewData.length;document.getElementById('stat-roots').textContent=roots.length;let visCount=0;document.querySelectorAll('.node-card').forEach(card=>{if(!card.closest('li.collapsed > ul')&&!card.closest('li.collapsed > .children-rows-wrap'))visCount++;});document.getElementById('stat-vis').textContent=visCount;document.getElementById('stat-filtered').style.display=Object.values(S.activeFilters).some(v=>v)?'flex':'none';const mgrStat=document.getElementById('stat-mgr-mode');const mgrVal=document.getElementById('stat-mgr-val');if(mgrStat){mgrStat.style.display=S.managerMode?'flex':'none';if(S.managerMode&&mgrVal){mgrVal.textContent=S.viewData.filter(n=>!isManager(n.id)).length+' ICs in lists';}}}
function cwrap(){return document.getElementById('chart-canvas-wrap');}
function ccontent(){return document.getElementById('chart-canvas-content');}
function applyZoom(z){S.zoom=Math.max(0.1,Math.min(3,z));ccontent().style.transform='scale('+S.zoom+')';document.getElementById('zoom-level').textContent=Math.round(S.zoom*100)+'%';if(S.gridMode){clearTimeout(window._gz);window._gz=setTimeout(redrawGridConnectorsFromTree,80);}}
function zoomBy(d){applyZoom(S.zoom+d);}
function fitToScreen(andCenter){requestAnimationFrame(()=>{const tree=document.getElementById('org-tree');const wrap=cwrap();if(!tree||!wrap)return;const tw=tree.scrollWidth,th=tree.scrollHeight,aw=wrap.clientWidth-100,ah=wrap.clientHeight-100;if(tw<10||th<10)return;applyZoom(Math.max(0.12,Math.min(1,aw/tw,ah/th)));if(andCenter)setTimeout(centerView,70);});}
function centerView(){const wrap=cwrap();const tree=document.getElementById('org-tree');if(!wrap||!tree)return;const sw=tree.scrollWidth*S.zoom;wrap.scrollLeft=Math.max(0,(sw-wrap.clientWidth)/2);wrap.scrollTop=0;}
let _panning=false,_px,_py,_psl,_pst;
function initPan(){const wrap=cwrap();if(!wrap)return;wrap.onmousedown=e=>{if(e.target.closest('.node-card,.summary-list-card,.collapse-btn'))return;_panning=true;_px=e.clientX;_py=e.clientY;_psl=wrap.scrollLeft;_pst=wrap.scrollTop;wrap.style.cursor='grabbing';};window.onmousemove=e=>{if(!_panning)return;cwrap().scrollLeft=_psl-(e.clientX-_px);cwrap().scrollTop=_pst-(e.clientY-_py);};window.onmouseup=()=>{_panning=false;if(cwrap())cwrap().style.cursor='';};wrap.addEventListener('wheel',e=>{if(e.ctrlKey||e.metaKey){e.preventDefault();zoomBy(e.deltaY<0?0.08:-0.08);}},{passive:false});}
function initSearch(){
  const input=document.getElementById('chart-search');
  const box=document.getElementById('chart-search-results');
  if(!input||input._initSearch)return;input._initSearch=true;
  function positionBox(){const r=input.getBoundingClientRect();box.style.top=(r.bottom+4)+'px';box.style.left=r.left+'px';box.style.width=Math.max(290,r.width)+'px';}
  let _searchTimer=null;let _lastQuery='';
  function runSearch(){
    const q=input.value.trim().toLowerCase();
    if(q===_lastQuery)return;
    _lastQuery=q;
    if(!q){box.classList.remove('visible');box.innerHTML='';return;}
    const hits=S.viewData.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)).slice(0,10);
    const rawHits=S.rawRows.map(r=>{const id=String(r[S.colMap.empId]||'').replace(/\.0$/,'').trim();const name=String(r[S.colMap.empName]||'');return{id,name};}).filter(n=>n.id&&!hits.find(h=>h.id===n.id)&&(n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q))).slice(0,5);
    const allHits=[...hits.map(n=>({...n,inChart:true})),...rawHits.map(n=>({...n,inChart:false}))];
    const newHtml=allHits.length?allHits.map(n=>'<div class="sr-item" onmousedown="event.preventDefault()" onclick="openPersonView(\''+esc(n.id)+'\')"><div class="sr-info"><div class="sr-name">'+esc(n.name)+'</div><div class="sr-sub">'+esc(n.id)+(n.inChart?'':' · <em>not in current filter</em>')+'</div></div><span class="sr-chev">›</span></div>').join(''):'<div class="sr-item" style="color:var(--text3);font-size:0.8rem;padding:12px 13px;cursor:default">No results</div>';
    if(box.innerHTML!==newHtml)box.innerHTML=newHtml;
    positionBox();box.classList.add('visible');
  }
  input.addEventListener('input',function(){
    clearTimeout(_searchTimer);
    _searchTimer=setTimeout(runSearch,120);
  });
  input.addEventListener('focus',()=>{if(input.value.trim())positionBox();});
  document.addEventListener('click',e=>{if(!e.target.closest('.search-wrap')&&!e.target.closest('#chart-search-results')){box.classList.remove('visible');}});
  window.addEventListener('resize',()=>{if(box.classList.contains('visible'))positionBox();});
}
function highlightNode(id){document.querySelectorAll('.node-card.highlighted').forEach(c=>c.classList.remove('highlighted'));S.highlighted=id;expandAll();const li=document.querySelector('#org-tree li[data-id="'+CSS.escape(id)+'"]');if(li){const card=li.querySelector('.node-card');if(card){card.classList.add('highlighted');setTimeout(()=>{const r=card.getBoundingClientRect();const w=cwrap();const wr=w.getBoundingClientRect();w.scrollTo({left:w.scrollLeft+(r.left-wr.left)-wr.width/2+r.width/2,top:w.scrollTop+(r.top-wr.top)-wr.height/2+r.height/2,behavior:'smooth'});},80);}}document.getElementById('chart-search').value='';document.getElementById('chart-search-results').classList.remove('visible');}

/* ════════════════════════════════════════════════════
   FRO LINES — main chart
   ════════════════════════════════════════════════════ */
function drawFROLine(svg,x1,y1,x2,y2,uid){
  let defs=svg.querySelector('defs');
  if(!defs){defs=document.createElementNS('http://www.w3.org/2000/svg','defs');svg.insertBefore(defs,svg.firstChild);}
  const markerId='fro-m-'+String(uid).replace(/[^a-zA-Z0-9]/g,'_').substring(0,30);
  const marker=document.createElementNS('http://www.w3.org/2000/svg','marker');
  marker.setAttribute('id',markerId);marker.setAttribute('markerWidth','8');marker.setAttribute('markerHeight','6');
  marker.setAttribute('refX','7');marker.setAttribute('refY','3');marker.setAttribute('orient','auto');
  const arr=document.createElementNS('http://www.w3.org/2000/svg','polygon');
  arr.setAttribute('points','0 0, 8 3, 0 6');arr.setAttribute('fill','#7c3aed');arr.setAttribute('opacity','0.75');
  marker.appendChild(arr);defs.appendChild(marker);
  const midY=(y1+y2)/2;const cp=Math.abs(y2-y1)*0.35+30;
  const d=Math.abs(x1-x2)<60?
    `M ${x1} ${y1} C ${x1} ${y1-cp}, ${x2} ${y2+cp}, ${x2} ${y2}`:
    `M ${x1} ${y1} C ${x1} ${midY}, ${x2} ${midY}, ${x2} ${y2}`;
  const path=document.createElementNS('http://www.w3.org/2000/svg','path');
  path.setAttribute('d',d);path.setAttribute('stroke','#7c3aed');path.setAttribute('stroke-width','2');
  path.setAttribute('stroke-dasharray','7,4');path.setAttribute('fill','none');path.setAttribute('opacity','0.65');
  path.setAttribute('marker-end','url(#'+markerId+')');svg.appendChild(path);
  // FRO label at midpoint
  const lx=(x1+x2)/2,ly=(y1+y2)/2-6;
  const lbg=document.createElementNS('http://www.w3.org/2000/svg','rect');
  lbg.setAttribute('x',String(lx-14));lbg.setAttribute('y',String(ly-8));
  lbg.setAttribute('width','28');lbg.setAttribute('height','14');lbg.setAttribute('rx','4');
  lbg.setAttribute('fill','#f5f3ff');lbg.setAttribute('stroke','#ddd6fe');lbg.setAttribute('stroke-width','1');
  svg.appendChild(lbg);
  const lt=document.createElementNS('http://www.w3.org/2000/svg','text');
  lt.setAttribute('x',String(lx));lt.setAttribute('y',String(ly));lt.setAttribute('font-size','9');
  lt.setAttribute('font-weight','800');lt.setAttribute('fill','#7c3aed');lt.setAttribute('text-anchor','middle');
  lt.setAttribute('dominant-baseline','middle');lt.setAttribute('font-family',"'Plus Jakarta Sans',sans-serif");
  lt.textContent='FRO';svg.appendChild(lt);
}

function renderFROLines(){
  const svg=document.getElementById('fro-svg');if(!svg)return;svg.innerHTML='';
  if(!S.colMap.froId||!S.viewData.length)return;
  const content=document.getElementById('chart-canvas-content');if(!content)return;
  svg.setAttribute('width',(content.scrollWidth||content.offsetWidth)+'px');
  svg.setAttribute('height',(content.scrollHeight||content.offsetHeight)+'px');
  const contentRect=content.getBoundingClientRect();
  S.viewData.forEach(node=>{
    const froId=String(node[S.colMap.froId]||'').replace(/\.0$/,'').trim();
    if(!froId||froId===node.id)return;
    const fromLi=document.querySelector('#org-tree li[data-id="'+CSS.escape(node.id)+'"]');
    const toLi=document.querySelector('#org-tree li[data-id="'+CSS.escape(froId)+'"]');
    if(!fromLi||!toLi)return;
    const fc=fromLi.querySelector(':scope>.node-card');const tc=toLi.querySelector(':scope>.node-card');if(!fc||!tc)return;
    const fr=fc.getBoundingClientRect(),tr=tc.getBoundingClientRect();
    const x1=(fr.left+fr.width/2-contentRect.left)/S.zoom;
    const y1=(fr.top-contentRect.top)/S.zoom;
    const x2=(tr.left+tr.width/2-contentRect.left)/S.zoom;
    const y2=(tr.bottom-contentRect.top)/S.zoom;
    drawFROLine(svg,x1,y1,x2,y2,node.id+'_m');
  });
}

/* ════════════════════════════════════════════════════
   PERSON VIEW — cross-filter org chart for any person
   ════════════════════════════════════════════════════ */
function openPersonView(personId){
  S.pvPersonId=personId;S.pvDepth=999;S.pvZoom=1;
  const rawRow=S.rawRows.find(r=>String(r[S.colMap.empId]||'').replace(/\.0$/,'').trim()===personId);
  const name=rawRow?String(rawRow[S.colMap.empName]||personId):personId;
  document.getElementById('pv-title').textContent=name;
  document.getElementById('pv-sub').textContent='ID: '+personId+' · Cross-filter view · All data · FRO shown as dotted line';
  document.querySelectorAll('.pv-depth-btn').forEach(b=>b.classList.toggle('selected',b.dataset.d==='999'));
  document.getElementById('person-view-modal').classList.remove('hidden');
  document.getElementById('chart-search').value='';document.getElementById('chart-search-results').classList.remove('visible');
  const pvSel=document.getElementById('pv-row-size-select');if(pvSel)pvSel.value=String(S.pvMaxPerRow);
  // Sync the PV grid-mode UI with restored state so the SVG, CSS connectors
  // and button highlight all match S.pvGridMode after a page reload.
  const pvTc=document.getElementById('pv-tree-content');if(pvTc)pvTc.classList.toggle('grid-mode',S.pvGridMode);
  const pvBtn=document.getElementById('pv-grid-btn');if(pvBtn)pvBtn.classList.toggle('active',S.pvGridMode);
  renderPersonView(personId,999);
  initPVPan();
}
function closePV(){document.getElementById('person-view-modal').classList.add('hidden');S.pvPersonId=null;}
function setPVDepth(d){S.pvDepth=d;document.querySelectorAll('.pv-depth-btn').forEach(b=>b.classList.toggle('selected',parseInt(b.dataset.d)===d||(d===999&&b.dataset.d==='999')));renderPersonView(S.pvPersonId,d);}

function buildRawChildMap(){
  const{empId,empName,managerId}=S.colMap;
  const allNodes=S.rawRows.map(row=>{
    const id=String(row[empId]||'').replace(/\.0$/,'').trim();
    const mgr=managerId?String(row[managerId]||'').replace(/\.0$/,'').trim():'';
    const node={id,name:String(row[empName]||'Unknown'),manager:mgr};
    S.columns.forEach(col=>{node[col]=String(row[col]||'');});
    return node;
  }).filter(n=>n.id);
  const validIds=new Set(allNodes.map(n=>n.id));
  allNodes.forEach(n=>{if(S.managerOverrides.hasOwnProperty(n.id))n.manager=S.managerOverrides[n.id];if(!validIds.has(n.manager)||n.manager===n.id)n.manager='';});
  const childMap={};allNodes.forEach(n=>{if(!childMap[n.manager])childMap[n.manager]=[];childMap[n.manager].push(n);});
  const byId=Object.fromEntries(allNodes.map(n=>[n.id,n]));
  return{allNodes,childMap,byId};
}

function renderPersonView(personId,maxDepth){
  if(!personId)return;
  const{allNodes,childMap,byId}=buildRawChildMap();
  // Collect subtree nodes
  const included=new Set();
  function collect(id,d){if(included.has(id))return;included.add(id);if(d<maxDepth)(childMap[id]||[]).forEach(k=>collect(k.id,d+1));}
  collect(personId,0);
  // Build temp state for rendering
  const savedVD=S.viewData,savedCM=S.childMap,savedDC=S.descCount,savedNH=S.nodeHeight,savedND=S.nodeDepth;
  const savedPvMode=S.pvMode;S.pvMode=true;
  S.viewData=allNodes.filter(n=>included.has(n.id));
  // Make personId a root
  const personNode=byId[personId];const savedMgr=personNode?personNode.manager:undefined;if(personNode)personNode.manager='';
  S.childMap={};S.viewData.forEach(n=>{if(!S.childMap[n.manager])S.childMap[n.manager]=[];S.childMap[n.manager].push(n);});
  S.descCount={};S.nodeHeight={};S.nodeDepth={};
  function cD(id){const k=S.childMap[id]||[];S.descCount[id]=k.reduce((s,c)=>s+1+cD(c.id),0);return S.descCount[id];}
  function cH(id){const k=S.childMap[id]||[];S.nodeHeight[id]=k.length?1+Math.max(...k.map(c=>cH(c.id))):0;return S.nodeHeight[id];}
  function cDep(id,d){S.nodeDepth[id]=d;(S.childMap[id]||[]).forEach(k=>cDep(k.id,d+1));}
  if(byId[personId]){cD(personId);cH(personId);cDep(personId,0);}
  const pvTree=document.getElementById('pv-org-tree');pvTree.innerHTML='';
  const root=S.viewData.find(n=>n.id===personId);
  if(root){const ul=document.createElement('ul');ul.appendChild(mkNodeLI(root,0));pvTree.appendChild(ul);}
  // Restore
  if(personNode&&savedMgr!==undefined)personNode.manager=savedMgr;
  S.viewData=savedVD;S.childMap=savedCM;S.descCount=savedDC;S.nodeHeight=savedNH;S.nodeDepth=savedND;S.pvMode=savedPvMode;
  // Re-apply PV grid translations + connectors if grid mode was on for this session
  if(S.pvGridMode){
    setPVCardDraggability(true);
    bindPVCanvasGridDND();
    setTimeout(()=>{applyPVGridOverridesToTree();redrawPVConnectors();},180);
  }
  // Fit + FRO lines
  setTimeout(()=>{pvFit();setTimeout(()=>renderPVFROLines(),350);},180);
  // Update stats in modal
  document.getElementById('pv-sub').textContent='ID: '+personId+' · '+included.size+' people · Cross-filter · FRO shown';
}

function renderPVFROLines(){
  const svg=document.getElementById('pv-fro-svg');if(!svg)return;svg.innerHTML='';
  if(!S.colMap.froId)return;
  const treeContent=document.getElementById('pv-tree-content');if(!treeContent)return;
  svg.setAttribute('width',treeContent.scrollWidth+'px');svg.setAttribute('height',treeContent.scrollHeight+'px');
  const tcRect=treeContent.getBoundingClientRect();
  document.querySelectorAll('#pv-org-tree li[data-id]').forEach(li=>{
    const nodeId=li.dataset.id;
    const rawRow=S.rawRows.find(r=>String(r[S.colMap.empId]||'').replace(/\.0$/,'').trim()===nodeId);
    if(!rawRow)return;
    const froId=String(rawRow[S.colMap.froId]||'').replace(/\.0$/,'').trim();
    if(!froId||froId===nodeId)return;
    const toLi=document.querySelector('#pv-org-tree li[data-id="'+CSS.escape(froId)+'"]');if(!toLi)return;
    const fc=li.querySelector(':scope>.node-card');const tc=toLi.querySelector(':scope>.node-card');if(!fc||!tc)return;
    const fr=fc.getBoundingClientRect(),tr=tc.getBoundingClientRect();
    const x1=(fr.left+fr.width/2-tcRect.left)/S.pvZoom;
    const y1=(fr.top-tcRect.top)/S.pvZoom;
    const x2=(tr.left+tr.width/2-tcRect.left)/S.pvZoom;
    const y2=(tr.bottom-tcRect.top)/S.pvZoom;
    drawFROLine(svg,x1,y1,x2,y2,nodeId+'_pv');
  });
}

function pvFit(){
  requestAnimationFrame(()=>{
    const tree=document.getElementById('pv-org-tree');
    const area=document.getElementById('pv-chart-area');
    if(!tree||!area)return;
    const tw=tree.scrollWidth,th=tree.scrollHeight,aw=area.clientWidth-120,ah=area.clientHeight-120;
    if(tw<10||th<10)return;
    S.pvZoom=Math.max(0.1,Math.min(1,aw/tw,ah/th));
    document.getElementById('pv-tree-content').style.transform='scale('+S.pvZoom+')';
    document.getElementById('pv-zoom-level').textContent=Math.round(S.pvZoom*100)+'%';
    setTimeout(()=>{const sw=tree.scrollWidth*S.pvZoom;area.scrollLeft=Math.max(0,(sw-area.clientWidth)/2);area.scrollTop=0;},70);
  });
}
function pvZoomBy(d){S.pvZoom=Math.max(0.1,Math.min(3,S.pvZoom+d));document.getElementById('pv-tree-content').style.transform='scale('+S.pvZoom+')';document.getElementById('pv-zoom-level').textContent=Math.round(S.pvZoom*100)+'%';clearTimeout(window._pvFroTimer);window._pvFroTimer=setTimeout(renderPVFROLines,400);if(S.pvGridMode){clearTimeout(window._pvGz);window._pvGz=setTimeout(redrawPVConnectors,80);}}

let _pvPanning=false,_pvPx,_pvPy,_pvSl,_pvSt;
function initPVPan(){const area=document.getElementById('pv-chart-area');if(!area||area._pvPanInit)return;area._pvPanInit=true;area.addEventListener('mousedown',e=>{if(e.target.closest('.node-card,.summary-list-card,.collapse-btn'))return;_pvPanning=true;_pvPx=e.clientX;_pvPy=e.clientY;_pvSl=area.scrollLeft;_pvSt=area.scrollTop;area.style.cursor='grabbing';});window.addEventListener('mousemove',e=>{if(!_pvPanning)return;const a=document.getElementById('pv-chart-area');if(a){a.scrollLeft=_pvSl-(e.clientX-_pvPx);a.scrollTop=_pvSt-(e.clientY-_pvPy);}});window.addEventListener('mouseup',()=>{_pvPanning=false;const a=document.getElementById('pv-chart-area');if(a)a.style.cursor='';});area.addEventListener('wheel',e=>{if(e.ctrlKey||e.metaKey){e.preventDefault();pvZoomBy(e.deltaY<0?0.08:-0.08);}},{passive:false});}

function locatePersonOnChart(){if(!S.pvPersonId)return;closePV();setTimeout(()=>highlightNode(S.pvPersonId),80);}

/* ── Person View Grid Mode (translate-based, mirrors main chart) ── */
S.pvGridMode=false;S.pvGridOverrides={};
function togglePVGrid(){
  S.pvGridMode=!S.pvGridMode;
  document.getElementById('pv-grid-btn').classList.toggle('active',S.pvGridMode);
  const tc=document.getElementById('pv-tree-content');
  if(tc)tc.classList.toggle('grid-mode',S.pvGridMode);
  setPVCardDraggability(S.pvGridMode);
  if(S.pvGridMode){
    bindPVCanvasGridDND();
    setTimeout(()=>{applyPVGridOverridesToTree();redrawPVConnectors();},150);
    showToast('Person View · Grid Mode on — drag any card to reposition');
  }else{
    clearPVTreeTranslations();
    const svg=document.getElementById('pv-grid-svg');if(svg)svg.innerHTML='';
    showToast('Person View · Grid Mode off');
  }
}
function applyPVGridOverridesToTree(){
  document.querySelectorAll('#pv-org-tree .node-card[data-drag-id], #pv-org-tree .summary-list-card[data-drag-id]').forEach(card=>{
    const id=card.dataset.dragId;const ovr=S.pvGridOverrides[id];
    if(ovr&&typeof ovr.dx==='number'&&typeof ovr.dy==='number'&&(ovr.dx||ovr.dy)){
      card.style.transform='translate('+ovr.dx+'px,'+ovr.dy+'px)';
      card.classList.add('grid-translated');
    }else{card.style.transform='';card.classList.remove('grid-translated');}
  });
}
function clearPVTreeTranslations(){
  document.querySelectorAll('#pv-org-tree .grid-translated').forEach(card=>{
    card.style.transform='';card.classList.remove('grid-translated');
  });
}
function redrawPVConnectors(){
  if(!S.pvGridMode)return;
  const svg=document.getElementById('pv-grid-svg');if(!svg)return;svg.innerHTML='';
  const cc=document.getElementById('pv-tree-content');if(!cc)return;
  const ccRect=cc.getBoundingClientRect();
  let maxRight=Math.max(cc.scrollWidth,cc.offsetWidth);
  let maxBottom=Math.max(cc.scrollHeight,cc.offsetHeight);
  cc.querySelectorAll('#pv-org-tree .node-card[data-drag-id], #pv-org-tree .summary-list-card[data-drag-id]').forEach(card=>{
    const r=card.getBoundingClientRect();
    const right=(r.right-ccRect.left)/S.pvZoom;
    const bottom=(r.bottom-ccRect.top)/S.pvZoom;
    if(right>maxRight)maxRight=right;
    if(bottom>maxBottom)maxBottom=bottom;
  });
  const w=Math.ceil(maxRight)+40;const h=Math.ceil(maxBottom)+40;
  svg.setAttribute('width',w+'px');svg.setAttribute('height',h+'px');
  svg.setAttribute('overflow','visible');
  svg.style.width=w+'px';svg.style.height=h+'px';svg.style.overflow='visible';svg.style.display='block';
  function addPath(d,color){
    const p=document.createElementNS('http://www.w3.org/2000/svg','path');
    p.setAttribute('d',d);p.setAttribute('stroke',color||'#94a3b8');
    p.setAttribute('stroke-width','2');p.setAttribute('fill','none');
    p.setAttribute('stroke-linejoin','miter');p.setAttribute('shape-rendering','crispEdges');
    svg.appendChild(p);
  }
  function rectToLocal(r){
    return{
      x:(r.left-ccRect.left)/S.pvZoom,y:(r.top-ccRect.top)/S.pvZoom,
      cx:(r.left+r.width/2-ccRect.left)/S.pvZoom,cy:(r.top+r.height/2-ccRect.top)/S.pvZoom,
      bottom:(r.bottom-ccRect.top)/S.pvZoom,right:(r.right-ccRect.left)/S.pvZoom,
      w:r.width/S.pvZoom,h:r.height/S.pvZoom
    };
  }
  // Group child li elements by their containing parent li. Walk UP through
  // any wrappers (.children-rows-wrap, .children-row-ul, etc.) until we find
  // an ancestor LI — using li.parentElement.parentElement only works when the
  // tree is flat; it breaks the moment we render a row-wrapped subtree, which
  // is why grid lines disappeared on people with many direct reports.
  function findAncestorLi(li){
    let p=li.parentElement;
    while(p&&p.id!=='pv-org-tree'){
      if(p.tagName==='LI')return p;
      p=p.parentElement;
    }
    return null;
  }
  const lis=Array.from(document.querySelectorAll('#pv-org-tree li'));
  const byParent=new Map();
  lis.forEach(li=>{
    const pli=findAncestorLi(li);if(!pli)return;
    if(!byParent.has(pli))byParent.set(pli,[]);
    byParent.get(pli).push(li);
  });
  byParent.forEach((kids,pli)=>{
    const pCard=pli.querySelector(':scope > .node-card');if(!pCard)return;
    const pr=rectToLocal(pCard.getBoundingClientRect());
    const childRects=kids.map(li=>{const c=li.querySelector(':scope > .node-card');return c?rectToLocal(c.getBoundingClientRect()):null;}).filter(Boolean);
    if(!childRects.length)return;
    const minChildTop=Math.min.apply(null,childRects.map(c=>c.y));
    if(minChildTop<=pr.bottom+8){
      childRects.forEach(c=>{const trunkY=Math.min(pr.bottom,c.y)-16;addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY+' H '+c.cx+' V '+c.y,'#94a3b8');});
      return;
    }
    const trunkY=pr.bottom+Math.max(20,Math.round((minChildTop-pr.bottom)/2));
    addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY,'#94a3b8');
    const xs=childRects.map(c=>c.cx).concat([pr.cx]);
    const minX=Math.min.apply(null,xs),maxX=Math.max.apply(null,xs);
    if(maxX>minX+0.5)addPath('M '+minX+' '+trunkY+' H '+maxX,'#94a3b8');
    childRects.forEach(c=>addPath('M '+c.cx+' '+trunkY+' V '+c.y,'#94a3b8'));
  });
  // IC summary connectors in PV
  document.querySelectorAll('#pv-org-tree .summary-list-card[data-ic-manager]').forEach(ic=>{
    const mgrId=ic.dataset.icManager;if(!mgrId)return;
    const mgrCard=cc.querySelector('#pv-org-tree .node-card[data-drag-id="'+CSS.escape(mgrId)+'"]');
    if(!mgrCard)return;
    const pr=rectToLocal(mgrCard.getBoundingClientRect());
    const cr=rectToLocal(ic.getBoundingClientRect());
    if(cr.y>pr.bottom+8){
      const trunkY=pr.bottom+Math.max(20,Math.round((cr.y-pr.bottom)/2));
      addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY+' H '+cr.cx+' V '+cr.y,'#7c3aed');
    }else{
      const trunkY=Math.min(pr.bottom,cr.y)-16;
      addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY+' H '+cr.cx+' V '+cr.y,'#7c3aed');
    }
  });
}
function bindPVCanvasGridDND(){
  const cc=document.getElementById('pv-tree-content');if(!cc||cc._pvGridDND)return;cc._pvGridDND=true;
  cc.addEventListener('dragover',e=>{
    if(!S.draggingNodeId||!S.pvGridMode||!S.draggingFromPV)return;
    e.preventDefault();e.dataTransfer.dropEffect='move';
  });
  cc.addEventListener('drop',e=>{
    if(!S.pvGridMode||!S.draggingNodeId||!S.draggingFromPV)return;
    e.preventDefault();
    const id=S.draggingNodeId;
    const card=cc.querySelector('#pv-org-tree [data-drag-id="'+CSS.escape(id)+'"]');
    if(!card){S.draggingNodeId=null;return;}
    const ccRect=cc.getBoundingClientRect();
    const dropX=(e.clientX-ccRect.left)/S.pvZoom;
    const dropY=(e.clientY-ccRect.top)/S.pvZoom;
    const cr=card.getBoundingClientRect();
    const cardCX=(cr.left+cr.width/2-ccRect.left)/S.pvZoom;
    const cardCY=(cr.top+cr.height/2-ccRect.top)/S.pvZoom;
    const oldOvr=S.pvGridOverrides[id]||{dx:0,dy:0};
    let dx=oldOvr.dx+(dropX-cardCX);let dy=oldOvr.dy+(dropY-cardCY);
    const SNAP=20;dx=Math.round(dx/SNAP)*SNAP;dy=Math.round(dy/SNAP)*SNAP;
    if(dx===0&&dy===0)delete S.pvGridOverrides[id];
    else S.pvGridOverrides[id]={dx,dy};
    S.draggingNodeId=null;
    applyPVGridOverridesToTree();
    redrawPVConnectors();
  });
}
function renderPVGrid(){
  if(!S.pvPersonId)return;
  const{allNodes,childMap,byId}=buildRawChildMap();
  // Collect subtree nodes within current depth
  const included=new Set();
  function collect(id,d){if(included.has(id))return;included.add(id);if(d<S.pvDepth)(childMap[id]||[]).forEach(k=>collect(k.id,d+1));}
  collect(S.pvPersonId,0);
  const visibleNodes=allNodes.filter(n=>included.has(n.id));
  // Make personId the root (clear manager)
  const personNode=byId[S.pvPersonId];const savedMgr=personNode?personNode.manager:undefined;if(personNode)personNode.manager='';
  // Compute depth from root for each subtree node
  const depth={};const queue=[{id:S.pvPersonId,d:0}];const seen=new Set();
  while(queue.length){const{id,d}=queue.shift();if(seen.has(id))continue;seen.add(id);depth[id]=d;(childMap[id]||[]).filter(k=>included.has(k.id)).forEach(k=>queue.push({id:k.id,d:d+1}));}
  // Auto positions: row = depth+1, col = order in BFS within depth
  const rowsByDepth={};
  Object.entries(depth).sort((a,b)=>a[1]-b[1]).forEach(([id,d])=>{if(!rowsByDepth[d])rowsByDepth[d]=[];rowsByDepth[d].push(byId[id]);});
  const positions={};
  Object.keys(rowsByDepth).sort((a,b)=>+a-+b).forEach(d=>{rowsByDepth[d].forEach((n,i)=>{positions[n.id]={row:+d+1,col:i+1};});});
  // Apply per-PV overrides
  Object.entries(S.pvGridOverrides).forEach(([id,pos])=>{if(visibleNodes.find(v=>v.id===id))positions[id]={row:pos.row,col:pos.col};});
  let maxCol=1,maxRow=1;Object.values(positions).forEach(p=>{if(p.col>maxCol)maxCol=p.col;if(p.row>maxRow)maxRow=p.row;});
  // Render into #pv-org-grid using mkBareCard. We swap S.viewData/childMap temporarily so mkBareCard works.
  const sv=S.viewData,scm=S.childMap,sdc=S.descCount,snh=S.nodeHeight,snd=S.nodeDepth,spv=S.pvMode;
  S.pvMode=true;S.viewData=visibleNodes;S.childMap={};
  visibleNodes.forEach(n=>{const m=(n.id===S.pvPersonId)?'':n.manager;if(!S.childMap[m])S.childMap[m]=[];S.childMap[m].push(n);});
  S.descCount={};S.nodeHeight={};S.nodeDepth={};
  function cD(id){const k=S.childMap[id]||[];S.descCount[id]=k.reduce((s,c)=>s+1+cD(c.id),0);return S.descCount[id];}
  function cH(id){const k=S.childMap[id]||[];S.nodeHeight[id]=k.length?1+Math.max(...k.map(c=>cH(c.id))):0;return S.nodeHeight[id];}
  cD(S.pvPersonId);cH(S.pvPersonId);
  Object.assign(S.nodeDepth,depth);
  const grid=document.getElementById('pv-org-grid');grid.innerHTML='';
  grid.style.gridTemplateColumns='repeat('+(maxCol+2)+',280px)';
  grid.style.gridTemplateRows='repeat('+(maxRow+1)+',minmax(200px,auto))';
  const occupied={};Object.entries(positions).forEach(([id,p])=>{occupied[p.row+'_'+p.col]=id;});
  for(let r=1;r<=maxRow+1;r++){
    for(let c=1;c<=maxCol+2;c++){
      const cell=document.createElement('div');cell.className='grid-cell';
      cell.dataset.row=r;cell.dataset.col=c;
      cell.style.gridRow=r;cell.style.gridColumn=c;
      bindPVGridCellDND(cell);
      const occId=occupied[r+'_'+c];
      if(occId){const n=visibleNodes.find(v=>v.id===occId);if(n){const card=mkBareCard(n);cell.appendChild(card);cell.dataset.id=occId;}}
      grid.appendChild(cell);
    }
  }
  // Connectors via SVG
  drawPVGridConnectors(positions,visibleNodes);
  // Restore S
  if(personNode&&savedMgr!==undefined)personNode.manager=savedMgr;
  S.viewData=sv;S.childMap=scm;S.descCount=sdc;S.nodeHeight=snh;S.nodeDepth=snd;S.pvMode=spv;
  setTimeout(pvFit,160);
}
function bindPVGridCellDND(cell){
  cell.addEventListener('dragover',e=>{if(!S.draggingNodeId||!S.pvGridMode)return;e.preventDefault();e.dataTransfer.dropEffect='move';cell.classList.add('drop-target-cell');});
  cell.addEventListener('dragleave',()=>cell.classList.remove('drop-target-cell'));
  cell.addEventListener('drop',e=>{
    cell.classList.remove('drop-target-cell');
    if(!S.pvGridMode||!S.draggingNodeId)return;
    e.preventDefault();e.stopPropagation();
    const draggedId=S.draggingNodeId;
    const newRow=parseInt(cell.dataset.row),newCol=parseInt(cell.dataset.col);
    const occupantId=cell.dataset.id;
    if(occupantId&&occupantId!==draggedId){
      // swap
      const draggedOld=S.pvGridOverrides[draggedId];
      S.pvGridOverrides[draggedId]={row:newRow,col:newCol};
      if(draggedOld)S.pvGridOverrides[occupantId]={row:draggedOld.row,col:draggedOld.col};
    }else{
      S.pvGridOverrides[draggedId]={row:newRow,col:newCol};
    }
    S.draggingNodeId=null;
    renderPVGrid();
    showToast('Card repositioned in this view');
  });
}
function drawPVGridConnectors(positions,visible){
  const svg=document.getElementById('pv-grid-svg');if(!svg)return;svg.innerHTML='';
  const grid=document.getElementById('pv-org-grid');if(!grid)return;
  svg.setAttribute('width',(grid.scrollWidth||grid.offsetWidth)+'px');
  svg.setAttribute('height',(grid.scrollHeight||grid.offsetHeight)+'px');
  svg.style.left=grid.offsetLeft+'px';svg.style.top=grid.offsetTop+'px';
  const gRect=grid.getBoundingClientRect();
  const byId=Object.fromEntries(visible.map(n=>[n.id,n]));
  visible.forEach(n=>{
    if(!n.manager||!byId[n.manager])return;
    const childCell=grid.querySelector('[data-row="'+positions[n.id].row+'"][data-col="'+positions[n.id].col+'"]');
    const pPos=positions[n.manager];if(!pPos)return;
    const parentCell=grid.querySelector('[data-row="'+pPos.row+'"][data-col="'+pPos.col+'"]');
    if(!childCell||!parentCell)return;
    const cc=childCell.querySelector('.node-card'),pc=parentCell.querySelector('.node-card');
    if(!cc||!pc)return;
    const cr=cc.getBoundingClientRect(),pr=pc.getBoundingClientRect();
    const x1=(pr.left+pr.width/2-gRect.left)/S.pvZoom;
    const y1=(pr.bottom-gRect.top)/S.pvZoom;
    const x2=(cr.left+cr.width/2-gRect.left)/S.pvZoom;
    const y2=(cr.top-gRect.top)/S.pvZoom;
    const midY=(y1+y2)/2;
    const path=document.createElementNS('http://www.w3.org/2000/svg','path');
    path.setAttribute('d','M '+x1+' '+y1+' C '+x1+' '+midY+', '+x2+' '+midY+', '+x2+' '+y2);
    path.setAttribute('stroke','#94a3b8');path.setAttribute('stroke-width','2');
    path.setAttribute('fill','none');path.setAttribute('opacity','0.7');
    svg.appendChild(path);
  });
}
async function printPVA3(){
  if(!S.pvPersonId){alert('No person view open.');return;}
  const overlay=makeOverlay('Preparing A3 PDF…','Capturing this person\'s view');document.body.appendChild(overlay);
  let svgRestorers=[];
  try{
    // pv-tree-content carries both the tree DOM and the PV grid SVG, so a single
    // target works for both tree and grid modes.
    const target=document.getElementById('pv-tree-content');
    const wasTransform=target.style.transform;
    const savedPvZoom=S.pvZoom;
    target.style.transform='scale(1)';
    S.pvZoom=1;  // keep redrawPVConnectors' coord math in sync with the rendered scale
    await new Promise(r=>setTimeout(r,250));
    if(S.pvGridMode){redrawPVConnectors();await new Promise(r=>setTimeout(r,80));}
    svgRestorers=await svgsToImagesAsync(target);
    await new Promise(r=>setTimeout(r,150));
    const bounds=S.pvGridMode?_translatedBounds(target):null;
    const h2cOpts={backgroundColor:'#ffffff',scale:2,useCORS:true,logging:false,allowTaint:true,foreignObjectRendering:false};
    if(bounds){const pw=Math.max(target.scrollWidth,bounds.width),ph=Math.max(target.scrollHeight,bounds.height);h2cOpts.width=Math.ceil(pw);h2cOpts.height=Math.ceil(ph);h2cOpts.windowWidth=Math.ceil(pw)+200;h2cOpts.windowHeight=Math.ceil(ph)+200;h2cOpts.scrollX=0;h2cOpts.scrollY=0;h2cOpts.x=0;h2cOpts.y=0;}
    const canvas=await html2canvas(target,h2cOpts);
    S.pvZoom=savedPvZoom;
    target.style.transform=wasTransform;
    if(S.pvGridMode)setTimeout(redrawPVConnectors,40);
    const dataUrl=canvas.toDataURL('image/png');
    const w=window.open('','_blank','width=1400,height=900');
    if(!w){alert('Pop-up blocked.');return;}
    w.document.open();
    w.document.write('<!DOCTYPE html><html><head><title>Person View — A3 Print</title>'+
      '<style>@page{size:A3 landscape;margin:6mm}html,body{margin:0;padding:0;background:#fff;font-family:-apple-system,BlinkMacSystemFont,sans-serif}body{display:flex;flex-direction:column;align-items:center;padding:18px}.print-bar{display:flex;gap:10px;margin-bottom:10px}.print-bar button{padding:9px 16px;background:#4f46e5;color:#fff;border:none;border-radius:8px;font-weight:700;cursor:pointer}.print-bar .hint{font-size:12px;color:#64748b;align-self:center}img{max-width:100%;display:block}@media print{body{padding:0;display:block}.print-bar{display:none!important}img{width:100%;height:auto}}</style></head><body>'+
      '<div class="print-bar"><button onclick="window.print()">🖨 Print / Save as PDF</button><span class="hint">A3 landscape</span></div>'+
      '<img src="'+dataUrl+'" alt="Person View"/>'+
      '<script>window.addEventListener(\'load\',function(){setTimeout(function(){try{window.print();}catch(_){}}, 350);});<\/script>'+
      '</body></html>');
    w.document.close();
  }catch(e){console.error(e);alert('Print failed: '+e.message);}finally{restoreSvgs(svgRestorers);overlay.remove();}
}

async function exportPVPNG(){
  if(!S.pvPersonId)return;
  const overlay=makeOverlay('Exporting Person View...','');document.body.appendChild(overlay);
  const pvContent=document.getElementById('pv-tree-content');
  const savedTransform=pvContent.style.transform;
  const savedPvZoom=S.pvZoom;
  // Sync the inline transform AND S.pvZoom — redrawPVConnectors divides rect
  // coords by S.pvZoom, so they must match the actual rendered scale or the
  // SVG paths land at 1/zoom-off positions.
  pvContent.style.transform='scale(1)';
  S.pvZoom=1;
  try{
    await new Promise(r=>setTimeout(r,300));
    if(S.pvGridMode){redrawPVConnectors();await new Promise(r=>setTimeout(r,80));}
    const pvSvg=document.getElementById('pv-fro-svg');pvSvg.setAttribute('width',pvContent.scrollWidth+'px');pvSvg.setAttribute('height',pvContent.scrollHeight+'px');
    // In grid mode, CSS transforms on cards don't expand scrollWidth/Height. Scan
    // card positions to compute the true bounds so translated cards aren't cropped.
    const bounds=_translatedBounds(pvContent);
    const w=Math.max(pvContent.scrollWidth,bounds.width);
    const h=Math.max(pvContent.scrollHeight,bounds.height);
    const canvas=await html2canvas(pvContent,{backgroundColor:S.transparentExport?null:'#f1f5f9',scale:2,useCORS:true,logging:false,allowTaint:true,width:Math.ceil(w),height:Math.ceil(h),windowWidth:Math.ceil(w)+200,windowHeight:Math.ceil(h)+200,scrollX:0,scrollY:0,x:0,y:0});
    const name=(document.getElementById('pv-title').textContent||'person').replace(/[^a-zA-Z0-9]/g,'_');
    const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');
    await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,'person_'+name+'_N'+S.pvDepth+'_'+stamp+'.png');res();},'image/png'));
  }catch(e){alert('Export failed: '+e.message);}
  finally{
    S.pvZoom=savedPvZoom;
    pvContent.style.transform=savedTransform;
    // Redraw at the user's actual zoom so the on-screen paths line up again
    if(S.pvGridMode)setTimeout(redrawPVConnectors,40);
    overlay.remove();
  }
}
/* Compute the bounding box of every card inside `container`, INCLUDING any
   CSS transform. The browser doesn't grow scrollWidth/scrollHeight to include
   translated children, so we have to walk the cards and find the extreme
   right/bottom edges ourselves. */
function _translatedBounds(container){
  const rect=container.getBoundingClientRect();
  let maxRight=Math.max(container.scrollWidth,container.offsetWidth);
  let maxBottom=Math.max(container.scrollHeight,container.offsetHeight);
  container.querySelectorAll('.node-card,.summary-list-card').forEach(c=>{
    const r=c.getBoundingClientRect();
    const right=r.right-rect.left;
    const bottom=r.bottom-rect.top;
    if(right>maxRight)maxRight=right;
    if(bottom>maxBottom)maxBottom=bottom;
  });
  return {width:Math.ceil(maxRight)+24,height:Math.ceil(maxBottom)+24};
}

function triggerDownload(blob,fname){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=fname;a.click();URL.revokeObjectURL(url);}
function csvEsc(v){return'"'+String(v||'').replace(/"/g,'""').replace(/[\r\n]+/g,' ')+'"';}
function buildCSVContent(){const cols=[S.colMap.empId,S.colMap.empName,S.colMap.managerId,...S.columns.filter(c=>c!==S.colMap.empId&&c!==S.colMap.empName&&c!==S.colMap.managerId)].filter(Boolean);return cols.map(csvEsc).join(',')+'\n'+S.viewData.map(n=>cols.map(c=>csvEsc(n[c]||'')).join(',')).join('\n');}
function downloadCSV(){triggerDownload(new Blob([buildCSVContent()],{type:'text/csv;charset=utf-8;'}),'orgchart_export.csv');}
function makeOverlay(title,sub){const o=document.createElement('div');o.className='export-overlay';o.innerHTML='<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px">'+title+'</div><div style="font-size:0.75rem;color:#94a3b8;margin-top:4px">'+sub+'</div>';return o;}
function _saveCollapsedState(){const ids=[];document.querySelectorAll('li.collapsed').forEach(li=>{if(li.dataset.id)ids.push(li.dataset.id);});return ids;}
function _restoreCollapsedState(ids){if(!ids||!ids.length)return;const s=new Set(ids);document.querySelectorAll('li[data-id]').forEach(li=>{if(s.has(li.dataset.id)){const u=li.querySelector(':scope > ul,:scope > .children-rows-wrap');if(u){li.classList.add('collapsed');u.style.display='none';const card=li.querySelector('.node-card');if(card)card.classList.add('collapsed-node');const b=li.querySelector('.collapse-btn');if(b){b.innerHTML='▸';b.style.color='var(--warning)';}}}});setTimeout(()=>updateStats(),60);}
async function buildRenderStage(){
  const savedCollapsed=_saveCollapsedState();expandAll();
  await new Promise(r=>setTimeout(r,400));
  if(document.fonts&&document.fonts.ready)await document.fonts.ready;
  await new Promise(r=>setTimeout(r,200));
  // In grid mode, redraw connectors first so the cloned grid SVG carries the
  // freshest path geometry for the export.
  if(S.gridMode){redrawGridConnectorsFromTree();await new Promise(r=>setTimeout(r,80));}
  // Clone the WHOLE chart-canvas-content when grid mode is on so card
  // transforms + SVG connector lines come along; otherwise just the org-tree.
  const source=S.gridMode?document.getElementById('chart-canvas-content'):document.getElementById('org-tree');
  const container=document.createElement('div');
  container.className='export-stage-root'+(S.gridMode?' grid-mode':'');
  const stageBg=S.transparentExport?'transparent':S.chartBgColor;
  container.style.cssText='position:fixed;top:0;left:0;background:'+stageBg+';padding:48px 64px 80px 64px;display:inline-block;z-index:9998;pointer-events:none;overflow:visible';
  const clone=source.cloneNode(true);
  // Drop the zoom transform on the cloned chart-canvas-content so card
  // transforms render at their natural 1× pixel positions in the canvas.
  if(S.gridMode){clone.style.transform='';clone.style.transformOrigin='top left';}
  clone.querySelectorAll('.collapse-btn,.ncard-edit-btn,.ncard-export-btn').forEach(el=>el.remove());
  clone.querySelectorAll('li.collapsed').forEach(li=>{
    li.classList.remove('collapsed');
    const ul=li.querySelector(':scope > ul,:scope > .children-rows-wrap');
    if(ul)ul.style.removeProperty('display');
    const card=li.querySelector('.node-card');if(card)card.classList.remove('collapsed-node');
  });
  clone.querySelectorAll('.node-card,.summary-list-card').forEach(c=>{
    c.style.removeProperty('opacity');
    // Keep card transforms in grid mode so the user's drag-arrangement is exported.
    if(!S.gridMode)c.style.removeProperty('transform');
    c.style.setProperty('overflow','visible','important');
  });
  // Hide the gridline overlay pattern in the export (per print stylesheet contract).
  if(S.gridMode){const ov=clone.querySelector('#grid-overlay, .grid-overlay');if(ov)ov.style.display='none';}
  container.appendChild(clone);
  document.body.appendChild(container);
  await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
  await new Promise(r=>setTimeout(r,300));
  _restoreCollapsedState(savedCollapsed);
  return{stage:container,wrapper:container,sourceForBounds:S.gridMode?clone:null};
}
async function renderToCanvas(stageObj){
  const el=stageObj.stage;
  let w=el.scrollWidth||el.offsetWidth;
  let h=el.scrollHeight||el.offsetHeight;
  // In grid mode, scrollWidth/scrollHeight ignores translated cards. Walk
  // the cards and grow the canvas so nothing gets clipped.
  if(stageObj.sourceForBounds){
    const b=_translatedBounds(stageObj.sourceForBounds);
    // Stage has 48/64/80 padding around the clone — keep that in the canvas size.
    w=Math.max(w,b.width+128);
    h=Math.max(h,b.height+128);
  }
  const bg=S.transparentExport?null:S.chartBgColor;
  return html2canvas(el,{backgroundColor:bg,scale:2,useCORS:true,logging:false,allowTaint:true,foreignObjectRendering:false,width:Math.ceil(w),height:Math.ceil(h),windowWidth:Math.ceil(w)+200,windowHeight:Math.ceil(h)+200,scrollX:0,scrollY:0,x:0,y:0});
}
async function exportPNG(){const overlay=makeOverlay('Rendering org chart...','Capturing full chart at 2x resolution');document.body.appendChild(overlay);const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));let stage;try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');const fp=Object.values(S.activeFilters).filter(Boolean).map(v=>v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');const mode=S.managerMode?'_mgr_view':'';await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,'orgchart_'+(fp?fp+'_':'')+mode+stamp+'.png');res();},'image/png'));}catch(e){alert('PNG export failed: '+e.message);}finally{if(stage&&stage.wrapper)stage.wrapper.remove();overlay.remove();applyZoom(savedZoom);}}
async function exportSubtree(e,nodeId){e.stopPropagation();const node=S.viewData.find(n=>n.id===nodeId);if(!node)return;const includeIds=new Set([nodeId]);function collectDesc(id){(S.childMap[id]||[]).forEach(k=>{includeIds.add(k.id);collectDesc(k.id);});}collectDesc(nodeId);const overlay=makeOverlay('Exporting '+node.name+'\'s team ('+includeIds.size+')...','');document.body.appendChild(overlay);const savedViewData=S.viewData,savedChildMap=S.childMap,savedDescCount=S.descCount,savedNodeHeight=S.nodeHeight,savedNodeDepth=S.nodeDepth;const savedSkipDepth=S.skipDepth;const hadOverride=S.managerOverrides.hasOwnProperty(nodeId);const prevOverride=S.managerOverrides[nodeId];S.viewData=savedViewData.filter(n=>includeIds.has(n.id));S.managerOverrides[nodeId]='';S.skipDepth=0;S.childMap={};S.viewData.forEach(n=>{const mgr=(n.id===nodeId)?'':n.manager;if(!S.childMap[mgr])S.childMap[mgr]=[];S.childMap[mgr].push(n);});S.descCount={};S.nodeHeight={};S.nodeDepth={};function cD(id){const k=S.childMap[id]||[];S.descCount[id]=k.reduce((s,c)=>s+1+cD(c.id),0);return S.descCount[id];}function cH(id){const k=S.childMap[id]||[];S.nodeHeight[id]=k.length?1+Math.max(...k.map(c=>cH(c.id))):0;return S.nodeHeight[id];}function cDep(id,d){S.nodeDepth[id]=d;(S.childMap[id]||[]).forEach(k=>cDep(k.id,d+1));}cD(nodeId);cH(nodeId);cDep(nodeId,0);const savedZoom=S.zoom;applyZoom(1);renderChart();await new Promise(r=>setTimeout(r,400));let stage;try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const stamp=new Date().toISOString().slice(0,10).replace(/-/g,'');const safeName=node.name.replace(/[^a-zA-Z0-9]/g,'_');await new Promise(res=>canvas.toBlob(blob=>{if(blob)triggerDownload(blob,'team_'+safeName+'_'+stamp+'.png');res();},'image/png'));}catch(ex){alert('Subtree export failed: '+ex.message);}finally{if(stage&&stage.wrapper)stage.wrapper.remove();overlay.remove();applyZoom(savedZoom);if(hadOverride)S.managerOverrides[nodeId]=prevOverride;else delete S.managerOverrides[nodeId];S.viewData=savedViewData;S.childMap=savedChildMap;S.descCount=savedDescCount;S.nodeHeight=savedNodeHeight;S.nodeDepth=savedNodeDepth;S.skipDepth=savedSkipDepth;renderChart();}}
const SW=12192000,SH=6858000;
function pptxRect(id,x,y,cx,cy,fill){return '<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="r'+id+'"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+x+'" y="'+y+'"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="'+fill+'"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';}
function pptxTxt(id,x,y,cx,cy,text,sz,bold,color,algn){algn=algn||'ctr';return '<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="t'+id+'"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+x+'" y="'+y+'"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="'+algn+'"/><a:r><a:rPr lang="en-US" sz="'+sz+'" b="'+(bold?1:0)+'" dirty="0"><a:solidFill><a:srgbClr val="'+color+'"/></a:solidFill></a:rPr><a:t>'+xe(text)+'</a:t></a:r></a:p></p:txBody></p:sp>';}
function pptxImg(id,x,y,cx,cy,rId){return '<p:pic><p:nvPicPr><p:cNvPr id="'+id+'" name="img'+id+'"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="'+rId+'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="'+x+'" y="'+y+'"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';}
function pptxSlide(bg,content,rels){return ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="'+bg+'"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+SW+'" cy="'+SH+'"/><a:chOff x="0" y="0"/><a:chExt cx="'+SW+'" cy="'+SH+'"/></a:xfrm></p:grpSpPr>'+content+'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',rels];}
async function buildPPTXBlob(imgB64,cW,cH,titleSuffix){titleSuffix=titleSuffix||'';const ac=S.cardAccent.replace('#','');const stamp=new Date().toLocaleDateString('en-IN',{day:'numeric',month:'long',year:'numeric'});const activeF=Object.entries(S.activeFilters).filter(([,v])=>v);const filterLine=activeF.map(([k,v])=>k+': '+v).join('  |  ')||(titleSuffix||'All Employees');const roots=(S.childMap['']||[]).length;const mgrCount=S.viewData.filter(n=>isManager(n.id)).length;const modeNote=S.managerMode?' | Manager View (ICs in lists)':'';const[s1xml,s1rels]=pptxSlide('F1F5F9',pptxRect(2,0,0,SW,Math.round(SH*0.52),ac)+pptxRect(3,0,Math.round(SH*0.52),SW,Math.round(SH*0.48),'FFFFFF')+pptxTxt(4,Math.round(SW*0.08),Math.round(SH*0.12),Math.round(SW*0.84),Math.round(SH*0.22),'Org Chart',7600,true,'FFFFFF','l')+pptxTxt(5,Math.round(SW*0.08),Math.round(SH*0.35),Math.round(SW*0.84),420000,filterLine+modeNote,2200,true,'FFFFFF','l')+pptxTxt(6,Math.round(SW*0.08),Math.round(SH*0.44),Math.round(SW*0.84),340000,'Generated: '+stamp,1500,false,'C7D2FE','l')+pptxTxt(7,Math.round(SW*0.08),Math.round(SH*0.59),Math.round(SW*0.38),400000,String(S.viewData.length),5200,true,ac,'l')+pptxTxt(8,Math.round(SW*0.08),Math.round(SH*0.74),Math.round(SW*0.38),310000,'Total Employees',1600,false,'64748B','l')+pptxTxt(9,Math.round(SW*0.55),Math.round(SH*0.59),Math.round(SW*0.35),400000,String(mgrCount),4000,true,'64748B','l')+pptxTxt(10,Math.round(SW*0.55),Math.round(SH*0.74),Math.round(SW*0.35),310000,'Managers',1600,false,'64748B','l'),'Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>');const imgAspect=cW/cH,slideAspect=SW/SH;let iW,iH,iX,iY;if(imgAspect>=slideAspect){iW=SW;iH=Math.round(SW/imgAspect);iX=0;iY=Math.round((SH-iH)/2);}else{iH=SH;iW=Math.round(SH*imgAspect);iX=Math.round((SW-iW)/2);iY=0;}const capY=SH-Math.round(SH*0.065);const[s2xml,]=pptxSlide('FFFFFF',pptxImg(20,iX,iY,iW,iH,'rId2')+pptxTxt(21,Math.round(SW*0.04),capY,Math.round(SW*0.92),Math.round(SH*0.055),filterLine+modeNote+' · '+stamp+' · '+S.viewData.length+' employees',1000,false,'64748B','r'),'Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>');const statItems=[{label:'Total Employees',val:S.viewData.length,color:ac},{label:'Managers',val:mgrCount,color:'7c3aed'},{label:'Roots',val:roots,color:'0891b2'},{label:'Mode',val:S.managerMode?'Manager View':'Full Tree',color:'059669'}];const boxW=Math.round(SW*0.19),boxH=Math.round(SH*0.3),gap=Math.round((SW-boxW*4)*0.2);const totalBW=boxW*4+gap*3,bStartX=Math.round((SW-totalBW)/2),bY=Math.round(SH*0.3);let sc=pptxRect(2,0,0,SW,Math.round(SH*0.2),ac)+pptxTxt(3,Math.round(SW*0.04),0,Math.round(SW*0.6),Math.round(SH*0.2),'Summary Dashboard',1800,true,'FFFFFF','l')+pptxTxt(4,Math.round(SW*0.65),0,Math.round(SW*0.3),Math.round(SH*0.2),stamp,1200,false,'C7D2FE','r')+pptxTxt(5,Math.round(SW*0.04),Math.round(SH*0.22),Math.round(SW*0.92),Math.round(SH*0.06),filterLine+modeNote,1600,false,'64748B','l');statItems.forEach((st,i)=>{const bx=bStartX+i*(boxW+gap);sc+=pptxRect(10+i*2,bx,bY,boxW,boxH,'F8FAFC')+pptxRect(11+i*2,bx,bY,boxW,Math.round(boxH*0.05),st.color)+pptxTxt(20+i*2,bx,bY+Math.round(boxH*0.1),boxW,Math.round(boxH*0.52),String(st.val),4800,true,st.color)+pptxTxt(21+i*2,bx,bY+Math.round(boxH*0.7),boxW,Math.round(boxH*0.28),st.label,1300,false,'64748B');});const[s3xml,]=pptxSlide('FFFFFF',sc,'Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>');const mkRel=r=>'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><'+r+'</Relationships>';const BP={'[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>','_rels/.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>','ppt/presentation.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst><p:sldSz cx="'+SW+'" cy="'+SH+'" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>','ppt/_rels/presentation.xml.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>','ppt/theme/theme1.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="OrgTheme"><a:themeElements><a:clrScheme name="OrgScheme"><a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1><a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1><a:dk2><a:srgbClr val="0F172A"/></a:dk2><a:lt2><a:srgbClr val="F1F5F9"/></a:lt2><a:accent1><a:srgbClr val="'+ac+'"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="EF4444"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="06B6D4"/></a:accent6><a:hlink><a:srgbClr val="'+ac+'"/></a:hlink><a:folHlink><a:srgbClr val="64748B"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>','ppt/slideMasters/slideMaster1.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lstStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></a:lstStyle></p:titleStyle><p:bodyStyle><a:lstStyle/></p:bodyStyle><p:otherStyle><a:lstStyle/></p:otherStyle></p:txStyles></p:sldMaster>','ppt/slideMasters/_rels/slideMaster1.xml.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>','ppt/slideLayouts/slideLayout1.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>','ppt/slideLayouts/_rels/slideLayout1.xml.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>'};const zip=new JSZip();Object.entries(BP).forEach(([k,v])=>zip.file(k,v));zip.file('ppt/slides/slide1.xml',s1xml);zip.file('ppt/slides/_rels/slide1.xml.rels',mkRel(s1rels));zip.file('ppt/slides/slide2.xml',s2xml);zip.file('ppt/slides/_rels/slide2.xml.rels',mkRel('Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>'));zip.file('ppt/slides/slide3.xml',s3xml);zip.file('ppt/slides/_rels/slide3.xml.rels',mkRel('Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'));zip.file('ppt/media/image1.png',imgB64,{base64:true});return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'DEFLATE'});}
async function exportPPTX(){if(typeof JSZip==='undefined'){alert('ZIP library failed to load.');return;}const overlay=makeOverlay('Building PowerPoint...','Rendering then packaging');document.body.appendChild(overlay);const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));let stage;try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const blob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height);const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');const fp=Object.values(S.activeFilters).filter(Boolean).map(v=>v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');const mode=S.managerMode?'_mgr':'';triggerDownload(blob,'orgchart_'+(fp?fp+'_':'')+mode+dp+'.pptx');}catch(e){alert('PPTX failed: '+e.message);console.error(e);}finally{if(stage&&stage.wrapper)stage.wrapper.remove();overlay.remove();applyZoom(savedZoom);}}
async function exportAll(){if(typeof JSZip==='undefined'){alert('ZIP library failed to load.');return;}const lastFilterCol=S.filterCols[S.filterCols.length-1]||null;if(!lastFilterCol){await _exportAllSingleView();return;}const parentFilters=Object.entries(S.activeFilters).filter(([k])=>k!==lastFilterCol);const relevantRows=S.rawRows.filter(row=>parentFilters.every(([col,val])=>!val||String(row[col]||'').trim()===val));const lastVals=[...new Set(relevantRows.map(r=>String(r[lastFilterCol]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined'))].sort();if(!lastVals.length){await _exportAllSingleView();return;}const overlay=document.createElement('div');overlay.className='export-overlay';overlay.innerHTML='<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;color:#0f172a;margin-top:10px" id="_ea_title">Exporting '+lastVals.length+' charts...</div><div id="_ea_step" style="font-size:0.8rem;color:#64748b;margin-top:4px">Preparing...</div><div id="_ea_prog" style="font-size:0.7rem;color:var(--text3);margin-top:2px">0 / '+lastVals.length+'</div>';document.body.appendChild(overlay);const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));const savedFilters={...S.activeFilters};const outerZip=new JSZip();let successCount=0;try{for(let i=0;i<lastVals.length;i++){const val=lastVals[i];const safeName=val.replace(/[^a-zA-Z0-9]/g,'_');document.getElementById('_ea_step').textContent='📊 '+val;document.getElementById('_ea_prog').textContent=(i+1)+' / '+lastVals.length;S.activeFilters[lastFilterCol]=val;buildViewData();renderChart();await new Promise(r=>setTimeout(r,400));outerZip.file(safeName+'/'+safeName+'.csv',buildCSVContent());let stage2;try{stage2=await buildRenderStage();const canvas2=await renderToCanvas(stage2);outerZip.file(safeName+'/'+safeName+'.png',canvas2.toDataURL('image/png').split(',')[1],{base64:true});const pptxBlob=await buildPPTXBlob(canvas2.toDataURL('image/png').split(',')[1],canvas2.width,canvas2.height,val);outerZip.file(safeName+'/'+safeName+'.pptx',pptxBlob);successCount++;}finally{if(stage2&&stage2.wrapper)stage2.wrapper.remove();}}}finally{S.activeFilters=savedFilters;buildViewData();renderChart();buildFilterBar();overlay.remove();applyZoom(savedZoom);}if(successCount>0){const zipBlob=await outerZip.generateAsync({type:'blob',compression:'DEFLATE'});const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');triggerDownload(zipBlob,'orgcharts_all_'+dp+'.zip');}}
async function _exportAllSingleView(){const overlay=makeOverlay('Exporting current view...','PNG + PPTX + CSV');document.body.appendChild(overlay);const savedZoom=S.zoom;applyZoom(1);await new Promise(r=>setTimeout(r,140));let stage;try{stage=await buildRenderStage();const canvas=await renderToCanvas(stage);const dp=new Date().toISOString().slice(0,10).replace(/-/g,'');const zip=new JSZip();zip.file('orgchart.csv',buildCSVContent());zip.file('orgchart.png',canvas.toDataURL('image/png').split(',')[1],{base64:true});const pptxBlob=await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1],canvas.width,canvas.height);zip.file('orgchart.pptx',pptxBlob);const zipBlob=await zip.generateAsync({type:'blob',compression:'DEFLATE'});triggerDownload(zipBlob,'orgchart_'+dp+'.zip');}catch(e){alert('Export failed: '+e.message);}finally{if(stage&&stage.wrapper)stage.wrapper.remove();overlay.remove();applyZoom(savedZoom);}}
let _reassignAllNodes=[];
function openReassignModal(e,nodeId){e.stopPropagation();S.reassignTarget=nodeId;S.reassignPick=null;const node=S.viewData.find(n=>n.id===nodeId);document.getElementById('reassign-subject').innerHTML='Moving <strong>'+esc(node?node.name:nodeId)+'</strong>';document.getElementById('reassign-search').value='';document.getElementById('reassign-confirm-btn').disabled=true;document.getElementById('reassign-note').textContent='Select a new manager above';_reassignAllNodes=[{id:'__root__',name:'Make Root (no manager)',manager:''},...S.viewData.filter(n=>n.id!==nodeId)];renderReassignList(_reassignAllNodes);document.getElementById('reassign-modal').classList.remove('hidden');}
function closeReassignModal(){document.getElementById('reassign-modal').classList.add('hidden');S.reassignTarget=null;S.reassignPick=null;}
function filterReassignList(){const q=document.getElementById('reassign-search').value.trim().toLowerCase();renderReassignList(q?_reassignAllNodes.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)):_reassignAllNodes);}
function renderReassignList(nodes){document.getElementById('reassign-list').innerHTML=nodes.slice(0,60).map(n=>{const isRoot=n.id==='__root__';const initials=n.name.split(' ').map(w=>w[0]||'').join('').substring(0,2).toUpperCase();return '<div class="modal-emp-row'+(S.reassignPick===n.id?' selected':'')+'" onclick="pickReassign(\''+esc(n.id)+'\',\''+esc(n.name)+'\')"><div class="modal-emp-avatar">'+(isRoot?'🔼':esc(initials))+'</div><div><div class="modal-emp-name">'+esc(n.name)+'</div><div class="modal-emp-sub">'+(isRoot?'Will appear as root node':esc(n.id))+'</div></div></div>';}).join('');}
function pickReassign(id,name){S.reassignPick=id;document.getElementById('reassign-confirm-btn').disabled=false;document.getElementById('reassign-note').textContent='→ '+name;const q=document.getElementById('reassign-search').value.trim().toLowerCase();renderReassignList(q?_reassignAllNodes.filter(n=>n.name.toLowerCase().includes(q)||n.id.toLowerCase().includes(q)):_reassignAllNodes);}
function confirmReassign(){if(!S.reassignTarget||!S.reassignPick)return;pushUndo();S.managerOverrides[S.reassignTarget]=S.reassignPick==='__root__'?'':S.reassignPick;closeReassignModal();buildViewData();renderChart();persistState();showToast('Reassigned',true);}
function removeCurrentNode(){if(!S.reassignTarget)return;if(!confirm('Remove this person from the chart? Their direct reports will be re-parented to whoever was their manager.'))return;pushUndo();S.removedIds.add(S.reassignTarget);closeReassignModal();buildViewData();renderChart();persistState();showToast('Removed from chart',true);}

/* ════════════════════════════════════════════════════════════════════
   PHASE 1 · Drag-to-reassign  +  PHASE 3 · Undo  +  PHASE 2 · Persist
   ════════════════════════════════════════════════════════════════════ */
function isDescendant(maybeDescendantId,ancestorId){
  // returns true if maybeDescendantId is a descendant (or self) of ancestorId
  if(maybeDescendantId===ancestorId)return true;
  const byId=Object.fromEntries(S.viewData.map(n=>[n.id,n]));
  let cur=byId[maybeDescendantId];const seen=new Set();
  while(cur&&cur.manager&&!seen.has(cur.id)){
    seen.add(cur.id);
    if(cur.manager===ancestorId)return true;
    cur=byId[cur.manager];
  }
  return false;
}
function onCardDragStart(e){
  const card=e.currentTarget;
  const inPV=!!card.closest('#pv-org-tree, #pv-tree-content');
  // Block drag-start when we're not allowed to translate:
  //   - PV without pvGridMode: card is read-only
  //   - Main without gridMode: only drag-to-reassign is allowed
  //     (we don't preventDefault here because reassign IS a real interaction)
  if(inPV&&!S.pvGridMode){e.preventDefault();return;}
  const id=card.dataset.dragId;if(!id)return;
  S.draggingNodeId=id;
  S.draggingFromPV=inPV;
  card.classList.add('node-dragging');
  card.setAttribute('aria-grabbed','true');
  try{e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain',id);}catch(_){/*noop*/}
  const rdz=document.getElementById('root-drop-zone');if(rdz)rdz.classList.add('dragging-active');
  if(!inPV&&!S.gridMode&&!window._dragHintShown){
    window._dragHintShown=true;
    showToast('Tree mode: drop on another card to make that the new manager · Want to free-move? Turn on Grid Mode');
  }
}
function onCardDragOver(e){
  if(!S.draggingNodeId)return;
  if(S.gridMode){return;}// In grid mode the cell handles the visual feedback
  const targetId=e.currentTarget.dataset.dragId;
  if(targetId===S.draggingNodeId)return;
  // Cannot reparent under own descendant (would create a cycle)
  const invalid=isDescendant(targetId,S.draggingNodeId);
  e.preventDefault();
  e.dataTransfer.dropEffect=invalid?'none':'move';
  e.currentTarget.classList.add(invalid?'drop-halo-bad':'drop-halo');
}
function onCardDragLeave(e){
  e.currentTarget.classList.remove('drop-halo','drop-halo-bad');
}
function onCardDrop(e){
  // In grid mode, the canvas-content handler does the repositioning
  if(S.gridMode){return;}
  e.preventDefault();
  e.currentTarget.classList.remove('drop-halo','drop-halo-bad');
  const draggingId=S.draggingNodeId;if(!draggingId)return;
  const targetId=e.currentTarget.dataset.dragId;
  if(targetId===draggingId)return;
  // IC summary cards have synthetic ids (prefix '__ic__'). Reassign via tree-drag
  // doesn't make sense for them — they represent a list of people, not a single
  // node. Hint the user instead.
  if(String(draggingId).startsWith('__ic__')||String(targetId).startsWith('__ic__')){
    showToast('IC summary cards can only be repositioned in Grid Mode — turn it on first');
    S.draggingNodeId=null;return;
  }
  if(isDescendant(targetId,draggingId)){showToast('Cannot move under own descendant',true);return;}
  if(S.managerOverrides[draggingId]===targetId){return;}// no-op
  pushUndo();
  S.managerOverrides[draggingId]=targetId;
  S.draggingNodeId=null;
  buildViewData();renderChart();persistState();
  const tName=(S.viewData.find(n=>n.id===targetId)||{}).name||targetId;
  const dName=(S.viewData.find(n=>n.id===draggingId)||{}).name||draggingId;
  showToast('Moved '+dName+' under '+tName);
}
function onCardDragEnd(e){
  e.currentTarget.classList.remove('node-dragging');
  e.currentTarget.setAttribute('aria-grabbed','false');
  document.querySelectorAll('.drop-halo,.drop-halo-bad').forEach(c=>c.classList.remove('drop-halo','drop-halo-bad'));
  document.getElementById('root-drop-zone').classList.remove('dragging-active','over');
  S.draggingNodeId=null;
}
// Root drop zone (drop a card to make it a root)
function bindRootDropZone(){
  const z=document.getElementById('root-drop-zone');if(!z||z._bound)return;z._bound=true;
  z.addEventListener('dragover',e=>{if(!S.draggingNodeId)return;e.preventDefault();e.dataTransfer.dropEffect='move';z.classList.add('over');});
  z.addEventListener('dragleave',()=>z.classList.remove('over'));
  z.addEventListener('drop',e=>{
    e.preventDefault();z.classList.remove('over','dragging-active');
    const id=S.draggingNodeId;if(!id)return;
    if(S.managerOverrides[id]==='')return;
    pushUndo();
    S.managerOverrides[id]='';
    S.draggingNodeId=null;
    buildViewData();renderChart();persistState();
    const dName=(S.viewData.find(n=>n.id===id)||{}).name||id;
    showToast(dName+' is now a root');
  });
}

/* ── Toast ── */
let _toastTimer=null;
function showToast(msg,withAction,actionType){
  let t=document.getElementById('app-toast');
  if(!t){t=document.createElement('div');t.id='app-toast';t.className='toast';document.body.appendChild(t);}
  let btn='';
  if(withAction){
    if(actionType==='restore'){btn='<button class="toast-action" onclick="restoreSavedSession()">↻ Restore</button>';}
    else if(S.undoStack.length){btn='<button class="toast-action" onclick="undo()">↶ Undo</button>';}
  }
  t.innerHTML=esc(msg)+btn;
  t.classList.add('visible');
  clearTimeout(_toastTimer);
  // Restore prompt sticks longer; other toasts fade after ~3.2s
  _toastTimer=setTimeout(()=>t.classList.remove('visible'),actionType==='restore'?9000:3200);
}
function restoreSavedSession(){
  const p=window._pendingRestore;if(!p)return;
  try{
    if(applyPersisted(p)){
      buildViewData();buildFilterBar();renderChart();goTo('chart');
      showToast('Session restored — Ctrl+Z undoes any change');
    }else{
      showToast('Saved session no longer matches — starting fresh');
    }
  }catch(ex){console.error(ex);showToast('Could not restore session — starting fresh');}
  window._pendingRestore=null;
}

/* ── Undo stack ── */
function snapshotState(){
  return{
    managerOverrides:{...S.managerOverrides},
    removedIds:new Set(S.removedIds),
    activeFilters:{...S.activeFilters},
    skipDepth:S.skipDepth,
    managerMode:S.managerMode,
    // Capture grid-mode state too — without these, Ctrl+Z after a card-drag
    // would silently reset the grid layout to natural tree positions.
    gridMode:S.gridMode,
    gridOverrides:JSON.parse(JSON.stringify(S.gridOverrides||{})),
    pvGridMode:S.pvGridMode,
    pvGridOverrides:JSON.parse(JSON.stringify(S.pvGridOverrides||{}))
  };
}
function pushUndo(){
  S.undoStack.push(snapshotState());
  if(S.undoStack.length>UNDO_MAX)S.undoStack.shift();
  const ub=document.getElementById('undo-btn');if(ub)ub.disabled=false;
}
function undo(){
  if(!S.undoStack.length){showToast('Nothing to undo');return;}
  const prev=S.undoStack.pop();
  S.managerOverrides=prev.managerOverrides;
  S.removedIds=prev.removedIds;
  S.activeFilters=prev.activeFilters;
  S.skipDepth=prev.skipDepth;
  S.managerMode=prev.managerMode;
  S.gridMode=!!prev.gridMode;
  S.gridOverrides=prev.gridOverrides||{};
  S.pvGridMode=!!prev.pvGridMode;
  S.pvGridOverrides=prev.pvGridOverrides||{};
  const mb=document.getElementById('mgr-mode-btn');if(mb)mb.classList.toggle('active',S.managerMode);
  const sf=document.getElementById('summary-fields-wrap');if(sf)sf.style.display=S.managerMode?'flex':'none';
  // Sync the grid-mode UI to the restored state
  const cc=document.getElementById('chart-canvas-content');if(cc)cc.classList.toggle('grid-mode',S.gridMode);
  const gb=document.getElementById('grid-mode-btn');if(gb)gb.classList.toggle('active',S.gridMode);
  const gl=document.getElementById('grid-lines-btn'),gr=document.getElementById('grid-reset-btn');
  if(gl)gl.style.display=S.gridMode?'inline-flex':'none';
  if(gr)gr.style.display=S.gridMode?'inline-flex':'none';
  const pvb=document.getElementById('pv-grid-btn');if(pvb)pvb.classList.toggle('active',S.pvGridMode);
  const pvtc=document.getElementById('pv-tree-content');if(pvtc)pvtc.classList.toggle('grid-mode',S.pvGridMode);
  buildViewData();renderChart();buildFilterBar();persistState();
  const ub=document.getElementById('undo-btn');if(ub)ub.disabled=S.undoStack.length===0;
  showToast('Undid last change');
}
window.addEventListener('keydown',function(e){
  if((e.ctrlKey||e.metaKey)&&!e.shiftKey&&!e.altKey&&(e.key==='z'||e.key==='Z')){
    if(/^(INPUT|TEXTAREA|SELECT)$/.test((e.target||{}).tagName||''))return;
    e.preventDefault();undo();return;
  }
  // ESC closes whatever modal is open, otherwise clears multi-select
  if(e.key==='Escape'||e.key==='Esc'){
    const order=['dq-modal','insights-modal','reassign-modal','person-view-modal'];
    for(const id of order){
      const m=document.getElementById(id);
      if(m&&!m.classList.contains('hidden')){
        if(id==='reassign-modal')closeReassignModal();
        else if(id==='person-view-modal')closePV();
        else if(id==='dq-modal')closeDataQualityModal();
        else if(id==='insights-modal')closeInsightsModal();
        e.preventDefault();return;
      }
    }
    if(S.selectedIds&&S.selectedIds.size>0){clearSelection();e.preventDefault();}
  }
});

/* ── Persistence (localStorage) ── */
function fileSig(){return(S.columns||[]).join('|');}
function persistState(){
  if(!S.rawRows.length)return;
  clearTimeout(S._persistTimer);
  S._persistTimer=setTimeout(()=>{
    try{
      const data={
        sig:fileSig(),
        colMap:S.colMap,
        managerOverrides:S.managerOverrides,
        removedIds:[...S.removedIds],
        cardSlots:S.cardSlots,
        cardAccent:S.cardAccent,
        empTypeCol:S.empTypeCol,
        empTypeLabels:S.empTypeLabels,
        empTypeColors:S.empTypeColors,
        filterCols:S.filterCols,
        activeFilters:S.activeFilters,
        photoMatchCol:S.photoMatchCol,
        photoSize:S.photoSize,
        photoShape:S.photoShape,
        photoPlacement:S.photoPlacement,
        summaryField1:S.summaryField1,
        summaryField2:S.summaryField2,
        summaryField3:S.summaryField3,
        chartBgColor:S.chartBgColor,
        transparentExport:S.transparentExport,
        skipDepth:S.skipDepth,
        managerMode:S.managerMode,
        gridMode:S.gridMode,
        gridOverrides:S.gridOverrides,
        gridShowLines:S.gridShowLines,
        pvGridMode:S.pvGridMode,
        pvGridOverrides:S.pvGridOverrides,
        maxPerRow:S.maxPerRow,
        pvMaxPerRow:S.pvMaxPerRow,
        savedAt:Date.now()
      };
      localStorage.setItem(PERSIST_KEY,JSON.stringify(data));
    }catch(e){console.warn('persist failed',e);}
  },200);
}
function loadPersisted(){
  try{const raw=localStorage.getItem(PERSIST_KEY);return raw?JSON.parse(raw):null;}catch(e){return null;}
}
function clearPersisted(){try{localStorage.removeItem(PERSIST_KEY);}catch(e){}S.undoStack=[];const ub=document.getElementById('undo-btn');if(ub)ub.disabled=true;}

/* ════════════════════════════════════════════════════════════════════
   PHASE 5 · Live column-mapping conflict detection
   ════════════════════════════════════════════════════════════════════ */
function onMapChange(){
  const get=id=>document.getElementById('map-'+id).value;
  const cur={empId:get('empId'),empName:get('empName'),managerId:get('managerId'),froId:get('froId')};
  // Conflict = same column picked for two roles (empty values are fine)
  const counts={};
  Object.values(cur).filter(Boolean).forEach(v=>{counts[v]=(counts[v]||0)+1;});
  const dupCols=new Set(Object.keys(counts).filter(k=>counts[k]>1));
  ['empId','empName','managerId','froId'].forEach(role=>{
    const sel=document.getElementById('map-'+role);
    const hint=document.getElementById('hint-'+role);
    if(!sel||!hint)return;
    const v=sel.value;
    sel.classList.toggle('conflict',!!v&&dupCols.has(v));
    // Reset hint base text
    hint.classList.remove('warn','err','ok');
    if(role==='empId')hint.textContent='Unique identifier — also used to match photos';
    if(role==='empName')hint.textContent='Full name shown on the card';
    if(role==='managerId')hint.textContent='Links employee to their direct line manager';
    if(role==='froId')hint.textContent='Functional reporting officer — shown as a purple dotted line on the chart';
    if(v&&dupCols.has(v)){hint.textContent='Conflict — also used by another role above';hint.classList.add('err');}
  });
  // Live ID-resolution check for managerId
  if(cur.empId&&cur.managerId&&!dupCols.has(cur.empId)&&!dupCols.has(cur.managerId)){
    const ids=new Set(S.rawRows.map(r=>String(r[cur.empId]||'').replace(/\.0$/,'').trim()).filter(Boolean));
    let unresolved=0,total=0;
    S.rawRows.forEach(r=>{
      const m=String(r[cur.managerId]||'').replace(/\.0$/,'').trim();
      if(!m)return;total++;if(!ids.has(m))unresolved++;
    });
    const hint=document.getElementById('hint-managerId');
    if(total===0){hint.textContent='No manager IDs in this column';hint.classList.add('warn');}
    else if(unresolved===0){hint.textContent='All '+total+' manager IDs resolve to valid employees ✓';hint.classList.add('ok');}
    else{hint.textContent=unresolved+' of '+total+' manager IDs do not match any Employee ID';hint.classList.add('warn');}
  }
}

/* ════════════════════════════════════════════════════════════════════
   PHASE 4 · Data quality validation + report
   ════════════════════════════════════════════════════════════════════ */
function validateData(){
  const issues={duplicates:[],selfRef:[],cycles:[],orphanMgrs:[],emptyIds:[]};
  const{empId,empName,managerId}=S.colMap;
  if(!empId)return issues;
  const idCounts={};const nodes={};
  S.rawRows.forEach((r,idx)=>{
    const id=String(r[empId]||'').replace(/\.0$/,'').trim();
    const name=String(r[empName]||'').trim();
    if(!id){issues.emptyIds.push({idx,name:name||'(row '+(idx+1)+')'});return;}
    idCounts[id]=(idCounts[id]||0)+1;
    if(!nodes[id])nodes[id]={id,name,mgr:managerId?String(r[managerId]||'').replace(/\.0$/,'').trim():''};
  });
  Object.entries(idCounts).filter(([,c])=>c>1).forEach(([id,c])=>{issues.duplicates.push({id,name:(nodes[id]||{}).name||'',count:c});});
  Object.values(nodes).forEach(n=>{
    if(n.mgr&&n.mgr===n.id)issues.selfRef.push(n);
    else if(n.mgr&&!nodes[n.mgr])issues.orphanMgrs.push({id:n.id,name:n.name,bogusMgr:n.mgr});
  });
  // Cycle detection (Floyd-style for each node)
  const cyclesFound=new Set();
  Object.values(nodes).forEach(start=>{
    if(cyclesFound.has(start.id))return;
    const path=[];const seen=new Set();let cur=start;
    while(cur&&cur.mgr&&nodes[cur.mgr]&&cur.mgr!==cur.id){
      if(seen.has(cur.id)){
        // Cycle found — collect cycle nodes
        const ci=path.indexOf(cur.id);
        const cycle=ci>=0?path.slice(ci):path.slice();cycle.push(cur.id);
        const key=[...new Set(cycle)].sort().join('|');
        if(!cyclesFound.has(key)){cyclesFound.add(key);issues.cycles.push(cycle);}
        cycle.forEach(id=>cyclesFound.add(id));
        break;
      }
      seen.add(cur.id);path.push(cur.id);
      cur=nodes[cur.mgr];
    }
  });
  return issues;
}
function dqIssueCount(i){return(i.duplicates.length+i.selfRef.length+i.cycles.length+i.orphanMgrs.length+i.emptyIds.length);}
function refreshDataQualityBtn(){
  const issues=validateData();const n=dqIssueCount(issues);
  const btn=document.getElementById('dq-btn');if(!btn)return;
  // Always rewrite innerHTML so we don't depend on a child #dq-count surviving prior swaps
  if(n>0){btn.style.display='inline-flex';btn.classList.remove('clean');btn.innerHTML='⚠ <span id="dq-count">'+n+'</span> issues';}
  else{btn.style.display='inline-flex';btn.classList.add('clean');btn.innerHTML='✓ Data clean';}
}
function openDataQualityModal(){
  const issues=validateData();
  const body=document.getElementById('dq-body');
  document.getElementById('dq-sub').textContent=dqIssueCount(issues)+' issues found across '+S.rawRows.length+' rows';
  const sec=(label,emoji,arr,renderRow,bulkAction)=>{
    const empty=arr.length===0;
    return '<div class="dq-section'+(empty?' empty':'')+'"><h4>'+emoji+' '+label+'<span class="dq-count">'+arr.length+'</span></h4>'+(empty?'<div class="dq-list" style="font-style:italic">None — looks good.</div>':'<div class="dq-list">'+arr.slice(0,40).map(renderRow).join('')+(arr.length>40?'<div class="dq-row" style="font-style:italic;color:var(--text3)">+ '+(arr.length-40)+' more…</div>':'')+'</div>'+(bulkAction&&!empty?bulkAction:''))+'</div>';
  };
  body.innerHTML=
    sec('Duplicate Employee IDs','♻️',issues.duplicates,
      d=>'<div class="dq-row"><span><strong>'+esc(d.id)+'</strong> — '+esc(d.name||'')+'</span><span style="color:var(--text3)">×'+d.count+' rows</span></div>')+
    sec('Empty Employee IDs','⬛',issues.emptyIds,
      d=>'<div class="dq-row"><span>row '+(d.idx+1)+(d.name?' — '+esc(d.name):'')+'</span></div>')+
    sec('Self-referencing manager','↪️',issues.selfRef,
      n=>'<div class="dq-row"><span><strong>'+esc(n.id)+'</strong> — '+esc(n.name||'')+'</span><span class="dq-fix" data-fix="self" data-id="'+esc(n.id)+'">Make root</span></div>')+
    sec('Reporting cycles','🔁',issues.cycles,
      c=>'<div class="dq-row"><span>'+c.map(id=>esc(id)).join(' → ')+'</span></div>')+
    sec('Manager IDs not in roster (orphans)','👻',issues.orphanMgrs,
      o=>'<div class="dq-row"><span><strong>'+esc(o.id)+'</strong> — '+esc(o.name||'')+' → manager <code style="background:var(--bg3);padding:1px 5px;border-radius:4px">'+esc(o.bogusMgr)+'</code> not found</span></div>',
      issues.orphanMgrs.length?'<div class="dq-bulk"><button class="btn btn-ghost btn-sm" onclick="dqMakeOrphansRoot()">Make all orphan-managed roots</button></div>':'');
  body.querySelectorAll('.dq-fix[data-fix="self"]').forEach(el=>{
    el.addEventListener('click',()=>dqFixSelfRef(el.dataset.id));
  });
  document.getElementById('dq-modal').classList.remove('hidden');
}
function closeDataQualityModal(){document.getElementById('dq-modal').classList.add('hidden');}
function dqFixSelfRef(id){pushUndo();S.managerOverrides[id]='';buildViewData();renderChart();persistState();refreshDataQualityBtn();showToast('Made root',true);openDataQualityModal();}
function dqMakeOrphansRoot(){const issues=validateData();if(!issues.orphanMgrs.length)return;pushUndo();issues.orphanMgrs.forEach(o=>{S.managerOverrides[o.id]='';});buildViewData();renderChart();persistState();refreshDataQualityBtn();showToast('Made '+issues.orphanMgrs.length+' orphan-managed nodes into roots',true);openDataQualityModal();}

/* ════════════════════════════════════════════════════════════════════
   PHASE 7 · Rule-based reorg insights
   ════════════════════════════════════════════════════════════════════ */
function computeInsights(){
  const out={wideSpan:[],deepChain:0,singleReport:[],vacantManagers:[],totalManagers:0,medianSpan:0,maxDepth:0,longestChain:[]};
  const byId=Object.fromEntries(S.viewData.map(n=>[n.id,n]));
  const managers=S.viewData.filter(n=>(S.childMap[n.id]||[]).length>0);
  out.totalManagers=managers.length;
  const spans=managers.map(m=>(S.childMap[m.id]||[]).length).sort((a,b)=>a-b);
  out.medianSpan=spans.length?spans[Math.floor(spans.length/2)]:0;
  managers.forEach(m=>{
    const span=(S.childMap[m.id]||[]).length;
    if(span>=12)out.wideSpan.push({id:m.id,name:m.name,span});
    if(span===1)out.singleReport.push({id:m.id,name:m.name});
  });
  // Max depth + longest chain
  let deepest=null,deepestD=-1;
  S.viewData.forEach(n=>{const d=S.nodeDepth[n.id]||0;if(d>deepestD){deepestD=d;deepest=n;}});
  out.deepChain=deepestD+1;out.maxDepth=deepestD+1;
  if(deepest){let cur=deepest;const chain=[];while(cur){chain.unshift(cur.name||cur.id);cur=byId[cur.manager];}out.longestChain=chain;}
  out.wideSpan.sort((a,b)=>b.span-a.span);
  return out;
}
function openInsightsModal(){
  if(!S.viewData.length){showToast('Load a chart first');return;}
  const ins=computeInsights();
  const sec=(emoji,title,sub,body)=>'<div class="dq-section"><h4>'+emoji+' '+title+'</h4><div style="font-size:0.76rem;color:var(--text3);margin-bottom:8px">'+sub+'</div>'+body+'</div>';
  let html='';
  html+='<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:14px">'+
    '<div style="padding:10px;background:var(--bg2);border-radius:8px"><div style="font-size:0.65rem;font-weight:800;color:var(--text3);text-transform:uppercase;letter-spacing:0.06em">Total Mgrs</div><div style="font-size:1.4rem;font-weight:800;color:var(--accent)">'+ins.totalManagers+'</div></div>'+
    '<div style="padding:10px;background:var(--bg2);border-radius:8px"><div style="font-size:0.65rem;font-weight:800;color:var(--text3);text-transform:uppercase;letter-spacing:0.06em">Median Span</div><div style="font-size:1.4rem;font-weight:800;color:var(--accent)">'+ins.medianSpan+'</div></div>'+
    '<div style="padding:10px;background:var(--bg2);border-radius:8px"><div style="font-size:0.65rem;font-weight:800;color:var(--text3);text-transform:uppercase;letter-spacing:0.06em">Max Depth</div><div style="font-size:1.4rem;font-weight:800;color:var(--accent)">'+ins.maxDepth+'</div></div>'+
    '</div>';
  html+=sec('🌳','Wide spans (≥12 reports)','Consider splitting these teams into sub-managers',
    ins.wideSpan.length?'<div class="dq-list">'+ins.wideSpan.slice(0,20).map(m=>'<div class="dq-row"><span><strong>'+esc(m.name)+'</strong> ('+esc(m.id)+')</span><span class="dq-count" style="background:#d97706">'+m.span+' reports</span></div>').join('')+'</div>':'<div style="font-style:italic;color:var(--text3);font-size:0.78rem">None — all spans look manageable.</div>');
  html+=sec('📏','Deep chains (>7 levels)','Layers of management between top and bottom',
    ins.maxDepth>7?'<div style="font-size:0.78rem;color:var(--danger);font-weight:700">⚠ Chain is '+ins.maxDepth+' levels deep — consider flattening.</div><div style="font-size:0.74rem;color:var(--text2);margin-top:6px">Longest chain: '+ins.longestChain.map(esc).join(' → ')+'</div>':'<div style="font-size:0.78rem;color:#059669;font-weight:700">✓ Max depth is '+ins.maxDepth+' levels — reasonable.</div>');
  html+=sec('🎯','Single-report managers','Managers with exactly 1 direct report — consider absorbing the role',
    ins.singleReport.length?'<div class="dq-list">'+ins.singleReport.slice(0,20).map(m=>'<div class="dq-row"><span><strong>'+esc(m.name)+'</strong> ('+esc(m.id)+')</span></div>').join('')+(ins.singleReport.length>20?'<div class="dq-row" style="font-style:italic;color:var(--text3)">+ '+(ins.singleReport.length-20)+' more…</div>':'')+'</div>':'<div style="font-style:italic;color:var(--text3);font-size:0.78rem">None — every manager has 2+ reports.</div>');
  document.getElementById('insights-body').innerHTML=html;
  document.getElementById('insights-sub').textContent=S.viewData.length+' employees · '+ins.totalManagers+' managers · median span '+ins.medianSpan;
  document.getElementById('insights-modal').classList.remove('hidden');
}
function closeInsightsModal(){document.getElementById('insights-modal').classList.add('hidden');}

/* ════════════════════════════════════════════════════════════════════
   PHASE 9 · Grid Mode — auto-arrange by depth + drag-to-cell override
   ════════════════════════════════════════════════════════════════════ */
S.gridMode=false;S.gridOverrides={};S.gridShowLines=true;S.selectedIds=new Set();
S.maxPerRow=6;S.pvMaxPerRow='auto';
function getEffectiveMaxPerRow(siblingCount){
  // In Person View we use the PV-specific row size so the user can decide
  // how many direct reports fit on a single row independently of the main chart.
  const setting=S.pvMode?S.pvMaxPerRow:S.maxPerRow;
  // If the user picked 'auto', aim for a near-square layout that minimizes
  // chart width. Otherwise honor the explicit value.
  if(setting==='auto'||setting==null){
    if(siblingCount<=4)return Math.max(1,siblingCount);
    if(siblingCount<=9)return Math.ceil(Math.sqrt(siblingCount));
    if(siblingCount<=16)return 4;
    if(siblingCount<=36)return 6;
    return Math.min(8,Math.ceil(Math.sqrt(siblingCount)));
  }
  const n=parseInt(setting);
  return (isFinite(n)&&n>0)?n:6;
}
function setMaxPerRow(val){
  pushUndo();
  S.maxPerRow=(val==='auto')?'auto':parseInt(val);
  const sel=document.getElementById('row-size-select');if(sel)sel.value=String(S.maxPerRow);
  renderChart();persistState();
  showToast('Max per row: '+(val==='auto'?'Auto-fit':val),true);
}
function setPVMaxPerRow(val){
  S.pvMaxPerRow=(val==='auto')?'auto':parseInt(val);
  const sel=document.getElementById('pv-row-size-select');if(sel)sel.value=String(S.pvMaxPerRow);
  if(S.pvPersonId)renderPersonView(S.pvPersonId,S.pvDepth);
  persistState();
  showToast('Person View · max per row: '+(val==='auto'?'Auto-fit':val),true);
}
/* Convert all <svg> elements inside a root to <img> data-URLs before html2canvas
   runs. html2canvas frequently renders SVGs as solid black blocks, especially
   when the SVG has computed widths or external CSS. Returns a list of restorers
   that puts the original SVGs back in place. */
async function svgsToImagesAsync(root){
  if(!root)return [];
  const restorers=[];
  const svgs=Array.from(root.querySelectorAll('svg'));
  await Promise.all(svgs.map(svg=>new Promise(resolve=>{
    try{
      const w=svg.scrollWidth||svg.clientWidth||parseInt(svg.getAttribute('width'))||0;
      const h=svg.scrollHeight||svg.clientHeight||parseInt(svg.getAttribute('height'))||0;
      if(!w||!h){resolve();return;}
      const cloned=svg.cloneNode(true);
      cloned.setAttribute('xmlns','http://www.w3.org/2000/svg');
      cloned.setAttribute('width',String(w));
      cloned.setAttribute('height',String(h));
      const xml=new XMLSerializer().serializeToString(cloned);
      const dataUrl='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(xml);
      const img=document.createElement('img');
      img.style.cssText=svg.style.cssText;
      img.style.width=w+'px';img.style.height=h+'px';
      img.style.position=svg.style.position||'absolute';
      img.style.top=svg.style.top||'0';
      img.style.left=svg.style.left||'0';
      img.style.pointerEvents='none';
      img.style.zIndex=svg.style.zIndex||'2';
      const parent=svg.parentNode;const next=svg.nextSibling;
      restorers.push(()=>{
        try{if(img.parentNode)img.parentNode.replaceChild(svg,img);
        else if(parent)parent.insertBefore(svg,next);}catch(_){/*noop*/}
      });
      img.onload=()=>resolve();
      img.onerror=()=>resolve();
      img.src=dataUrl;
      parent.replaceChild(img,svg);
    }catch(e){console.warn('svgToImage fail',e);resolve();}
  })));
  return restorers;
}
function restoreSvgs(restorers){(restorers||[]).forEach(r=>{try{r();}catch(_){}});}
async function printA3(){
  const overlay=makeOverlay('Preparing A3 PDF…','Rendering chart at high resolution');
  document.body.appendChild(overlay);
  const savedZoom=S.zoom;applyZoom(1);
  await new Promise(r=>setTimeout(r,140));
  let stage=null,canvas=null;let svgRestorers=[];
  // Hide the gridline overlay during capture (html2canvas can mis-render its background-image)
  const gridOverlay=document.getElementById('grid-overlay');
  const overlayWasVisible=gridOverlay&&gridOverlay.classList.contains('visible');
  if(overlayWasVisible)gridOverlay.classList.remove('visible');
  try{
    const cc=document.getElementById('chart-canvas-content');
    const wasTransform=cc.style.transform;cc.style.transform='scale(1)';
    await new Promise(r=>setTimeout(r,200));
    // Convert all SVGs in the chart canvas to image data-URLs first to
    // avoid html2canvas rendering them as black blocks.
    svgRestorers=await svgsToImagesAsync(cc);
    await new Promise(r=>setTimeout(r,150));
    // In grid mode, scrollWidth/scrollHeight doesn't include translated card positions.
    const bounds=S.gridMode?_translatedBounds(cc):null;
    const pw=bounds?Math.max(cc.scrollWidth,bounds.width):0;
    const ph=bounds?Math.max(cc.scrollHeight,bounds.height):0;
    const h2cOpts={backgroundColor:S.transparentExport?null:'#ffffff',scale:2,useCORS:true,logging:false,allowTaint:true,foreignObjectRendering:false};
    if(bounds){h2cOpts.width=Math.ceil(pw);h2cOpts.height=Math.ceil(ph);h2cOpts.windowWidth=Math.ceil(pw)+200;h2cOpts.windowHeight=Math.ceil(ph)+200;h2cOpts.scrollX=0;h2cOpts.scrollY=0;h2cOpts.x=0;h2cOpts.y=0;}
    canvas=await html2canvas(cc,h2cOpts);
    cc.style.transform=wasTransform;
    const dataUrl=canvas.toDataURL('image/png');
    const w=window.open('','_blank','width=1400,height=900');
    if(!w){alert('Pop-up blocked. Allow pop-ups for this site and try again.');return;}
    const stamp=new Date().toLocaleDateString();
    w.document.open();
    w.document.write(
      '<!DOCTYPE html><html><head><title>Org Chart — A3 Print</title>'+
      '<style>'+
      '@page{size:A3 landscape;margin:6mm}'+
      'html,body{margin:0;padding:0;background:#fff;font-family:-apple-system,BlinkMacSystemFont,sans-serif}'+
      'body{display:flex;flex-direction:column;align-items:center;justify-content:flex-start;min-height:100vh;padding:18px}'+
      '.print-bar{display:flex;gap:10px;align-items:center;margin-bottom:10px}'+
      '.print-bar button{padding:9px 16px;background:#4f46e5;color:#fff;border:none;border-radius:8px;font-weight:700;cursor:pointer;font-size:14px}'+
      '.print-bar button:hover{background:#4338ca}'+
      '.print-bar .hint{font-size:12px;color:#64748b}'+
      'img{max-width:100%;display:block;object-fit:contain}'+
      '@media print{body{padding:0;display:block}.print-bar{display:none!important}img{width:100%;height:auto;max-height:none;page-break-inside:avoid}}'+
      '</style></head><body>'+
      '<div class="print-bar"><button onclick="window.print()">🖨 Print / Save as PDF</button>'+
      '<span class="hint">Paper preset: <b>A3 landscape</b>. In the print dialog, choose "Save as PDF" or your printer.</span></div>'+
      '<img src="'+dataUrl+'" alt="Org Chart"/>'+
      '<script>window.addEventListener(\'load\',function(){setTimeout(function(){try{window.print();}catch(_){}}, 350);});<\/script>'+
      '</body></html>'
    );
    w.document.close();
  }catch(e){console.error(e);alert('Print failed: '+e.message);}
  finally{
    if(stage&&stage.wrapper)stage.wrapper.remove();
    restoreSvgs(svgRestorers);
    if(overlayWasVisible&&gridOverlay)gridOverlay.classList.add('visible');
    overlay.remove();
    applyZoom(savedZoom);
  }
}
/* ════════════════════════════════════════════════════════════════════
   GRID MODE v2 — tree layout stays, drag translates cards
   Replaces the old separate-grid layout. Cards keep their tree
   positions until the user drags one; per-card {dx,dy} translation is
   stored in S.gridOverrides and re-applied after every renderChart.
   ════════════════════════════════════════════════════════════════════ */
function toggleGridMode(){
  S.gridMode=!S.gridMode;
  const cc=document.getElementById('chart-canvas-content');
  if(cc)cc.classList.toggle('grid-mode',S.gridMode);
  document.getElementById('grid-mode-btn').classList.toggle('active',S.gridMode);
  document.getElementById('grid-lines-btn').style.display=S.gridMode?'inline-flex':'none';
  document.getElementById('grid-reset-btn').style.display=S.gridMode?'inline-flex':'none';
  setCardDraggability(S.gridMode);
  if(S.gridMode){
    bindCanvasGridDND();
    setTimeout(()=>{applyGridOverridesToTree();redrawGridConnectorsFromTree();applyGridLines();},120);
  }else{
    clearTreeTranslations();
    const svg=document.getElementById('grid-svg');if(svg)svg.innerHTML='';
    const ov=document.getElementById('grid-overlay');if(ov)ov.classList.remove('visible');
    clearSelection();
  }
  persistState();
  showToast(S.gridMode?'Grid Mode on — drag any card to reposition · Shift+click to multi-select':'Grid Mode off — tree view restored');
}
function toggleGridLines(){S.gridShowLines=!S.gridShowLines;applyGridLines();persistState();}
function applyGridLines(){
  const ov=document.getElementById('grid-overlay');if(!ov)return;
  ov.classList.toggle('visible',S.gridMode&&S.gridShowLines);
  // Resize overlay to canvas-content dimensions so gridlines cover the tree fully
  const cc=document.getElementById('chart-canvas-content');
  if(cc&&ov){ov.style.width=Math.max(cc.scrollWidth,cc.offsetWidth)+'px';ov.style.height=Math.max(cc.scrollHeight,cc.offsetHeight)+'px';}
}
function resetGridOverrides(){
  if(!Object.keys(S.gridOverrides).length){showToast('Nothing to reset');return;}
  if(!confirm('Reset every card to its original tree position?'))return;
  pushUndo();
  S.gridOverrides={};
  clearTreeTranslations();
  applyGridOverridesToTree();
  requestAnimationFrame(()=>{redrawGridConnectorsFromTree();});
  persistState();
  showToast('Cards returned to tree positions',true);
}
/* Apply / clear translations on the tree DOM (covers both regular cards and IC summary cards) */
function applyGridOverridesToTree(){
  document.querySelectorAll('#org-tree .node-card[data-drag-id], #org-tree .summary-list-card[data-drag-id]').forEach(card=>{
    const id=card.dataset.dragId;const ovr=S.gridOverrides[id];
    if(ovr&&typeof ovr.dx==='number'&&typeof ovr.dy==='number'&&(ovr.dx||ovr.dy)){
      card.style.transform='translate('+ovr.dx+'px,'+ovr.dy+'px)';
      card.classList.add('grid-translated');
    }else{
      card.style.transform='';
      card.classList.remove('grid-translated');
    }
  });
}
function clearTreeTranslations(){
  document.querySelectorAll('#org-tree .grid-translated').forEach(card=>{
    card.style.transform='';card.classList.remove('grid-translated');
  });
}
/* Draw shared-trunk org-chart connectors. For each manager with N children:
     - one vertical from the manager's bottom-center down to a trunk row
     - one horizontal trunk spanning from leftmost to rightmost child center
     - one vertical from the trunk down to each child's top-center
   This routes around cards instead of drawing a straight line that cuts
   through siblings (the previous bug). Plus a separate purple connector
   to the manager's IC summary card if any. */
function redrawGridConnectorsFromTree(){
  if(!S.gridMode)return;
  const svg=document.getElementById('grid-svg');if(!svg)return;svg.innerHTML='';
  const cc=document.getElementById('chart-canvas-content');if(!cc)return;
  // Compute SVG box big enough to cover all visible cards INCLUDING translated ones
  // (CSS transforms don't expand the parent's scrollWidth, so we have to scan).
  const ccRect=cc.getBoundingClientRect();
  let maxRight=Math.max(cc.scrollWidth,cc.offsetWidth);
  let maxBottom=Math.max(cc.scrollHeight,cc.offsetHeight);
  cc.querySelectorAll('#org-tree .node-card[data-drag-id], #org-tree .summary-list-card[data-drag-id]').forEach(card=>{
    const r=card.getBoundingClientRect();
    const right=(r.right-ccRect.left)/S.zoom;
    const bottom=(r.bottom-ccRect.top)/S.zoom;
    if(right>maxRight)maxRight=right;
    if(bottom>maxBottom)maxBottom=bottom;
  });
  const w=Math.ceil(maxRight)+40;const h=Math.ceil(maxBottom)+40;
  svg.setAttribute('width',w+'px');svg.setAttribute('height',h+'px');
  svg.setAttribute('overflow','visible');
  svg.style.width=w+'px';svg.style.height=h+'px';
  svg.style.overflow='visible';
  svg.style.display='block';
  function addPath(d,color){
    const p=document.createElementNS('http://www.w3.org/2000/svg','path');
    p.setAttribute('d',d);p.setAttribute('stroke',color||'#94a3b8');
    p.setAttribute('stroke-width','2');p.setAttribute('fill','none');
    p.setAttribute('stroke-linejoin','miter');p.setAttribute('shape-rendering','crispEdges');
    svg.appendChild(p);
  }
  function rectToLocal(r){
    return{
      x:(r.left-ccRect.left)/S.zoom,y:(r.top-ccRect.top)/S.zoom,
      cx:(r.left+r.width/2-ccRect.left)/S.zoom,cy:(r.top+r.height/2-ccRect.top)/S.zoom,
      bottom:(r.bottom-ccRect.top)/S.zoom,right:(r.right-ccRect.left)/S.zoom,
      w:r.width/S.zoom,h:r.height/S.zoom
    };
  }
  // Group children by manager
  const byMgr={};
  S.viewData.forEach(n=>{if(!n.manager)return;if(!byMgr[n.manager])byMgr[n.manager]=[];byMgr[n.manager].push(n);});
  Object.entries(byMgr).forEach(([mgrId,kids])=>{
    const mgrCard=cc.querySelector('#org-tree .node-card[data-drag-id="'+CSS.escape(mgrId)+'"]');
    if(!mgrCard)return;
    const childCards=kids.map(k=>cc.querySelector('#org-tree .node-card[data-drag-id="'+CSS.escape(k.id)+'"]')).filter(Boolean);
    if(!childCards.length)return;
    const pr=rectToLocal(mgrCard.getBoundingClientRect());
    const childRects=childCards.map(c=>rectToLocal(c.getBoundingClientRect()));
    const minChildTop=Math.min.apply(null,childRects.map(c=>c.y));
    // If a child sits above the manager (after manual drag), fall back to a per-child elbow that loops over the parent.
    if(minChildTop<=pr.bottom+8){
      childRects.forEach(c=>{
        const trunkY=Math.min(pr.bottom,c.y)-16;
        addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY+' H '+c.cx+' V '+c.y,'#94a3b8');
      });
      return;
    }
    // Standard org-chart shared trunk: trunk lives in the gap between parent bottom and lowest child top.
    const trunkY=pr.bottom+Math.max(20,Math.round((minChildTop-pr.bottom)/2));
    // Parent → trunk
    addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY,'#94a3b8');
    // Trunk horizontal
    const xs=childRects.map(c=>c.cx).concat([pr.cx]);
    const minX=Math.min.apply(null,xs),maxX=Math.max.apply(null,xs);
    if(maxX>minX+0.5)addPath('M '+minX+' '+trunkY+' H '+maxX,'#94a3b8');
    // Trunk → each child
    childRects.forEach(c=>{addPath('M '+c.cx+' '+trunkY+' V '+c.y,'#94a3b8');});
  });
  // Manager → IC summary connectors (purple)
  cc.querySelectorAll('#org-tree .summary-list-card[data-ic-manager]').forEach(ic=>{
    const mgrId=ic.dataset.icManager;if(!mgrId)return;
    const mgrCard=cc.querySelector('#org-tree .node-card[data-drag-id="'+CSS.escape(mgrId)+'"]');
    if(!mgrCard)return;
    const pr=rectToLocal(mgrCard.getBoundingClientRect());
    const cr=rectToLocal(ic.getBoundingClientRect());
    if(cr.y>pr.bottom+8){
      const trunkY=pr.bottom+Math.max(20,Math.round((cr.y-pr.bottom)/2));
      addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY+' H '+cr.cx+' V '+cr.y,'#7c3aed');
    }else{
      // IC sits to the side: route around with a 16px elbow above whichever is higher
      const trunkY=Math.min(pr.bottom,cr.y)-16;
      addPath('M '+pr.cx+' '+pr.bottom+' V '+trunkY+' H '+cr.cx+' V '+cr.y,'#7c3aed');
    }
  });
}
/* Pointer-event based drag for Grid Mode (main chart + Person View).
   We avoid HTML5 drag-and-drop here because it's unreliable inside
   Streamlit iframes (drop events sometimes don't fire, drag images
   are clipped, dataTransfer is flaky). With pointer capture we get
   real-time line updates as the card is dragged. */
function setCardDraggability(gridOn){
  document.querySelectorAll('#org-tree .node-card[data-drag-id], #org-tree .summary-list-card[data-drag-id]').forEach(card=>{
    card.draggable=!gridOn;
  });
}
function setPVCardDraggability(gridOn){
  document.querySelectorAll('#pv-org-tree .node-card[data-drag-id], #pv-org-tree .summary-list-card[data-drag-id]').forEach(card=>{
    card.draggable=!gridOn;
  });
}
function onCardPointerDown(e){
  const card=e.currentTarget;
  const inPV=!!card.closest('#pv-org-tree, #pv-tree-content');
  const gridOn=inPV?S.pvGridMode:S.gridMode;
  if(!gridOn)return;
  if(e.button!==0)return;
  if(e.target.closest('.ncard-edit-btn,.ncard-export-btn,.collapse-btn'))return;
  if(e.shiftKey)return;
  const id=card.dataset.dragId;if(!id)return;
  e.preventDefault();
  const overrides=inPV?S.pvGridOverrides:S.gridOverrides;
  const redraw=inPV?redrawPVConnectors:redrawGridConnectorsFromTree;
  const oldOvr=overrides[id]||{dx:0,dy:0};
  const startX=e.clientX,startY=e.clientY;
  let curDx=oldOvr.dx,curDy=oldOvr.dy;
  let moved=false;
  try{card.setPointerCapture(e.pointerId);}catch(_){}
  function onMove(ev){
    const zoom=inPV?S.pvZoom:S.zoom;
    const ddx=(ev.clientX-startX)/zoom;
    const ddy=(ev.clientY-startY)/zoom;
    if(!moved&&Math.abs(ddx)<3&&Math.abs(ddy)<3)return;
    moved=true;
    card.classList.add('node-dragging','grid-translated');
    curDx=oldOvr.dx+ddx;
    curDy=oldOvr.dy+ddy;
    card.style.transform='translate('+curDx+'px,'+curDy+'px)';
    if(!window._gridDragRAF){
      window._gridDragRAF=requestAnimationFrame(()=>{
        window._gridDragRAF=null;
        redraw();
      });
    }
  }
  function onUp(ev){
    card.removeEventListener('pointermove',onMove);
    card.removeEventListener('pointerup',onUp);
    card.removeEventListener('pointercancel',onUp);
    card.classList.remove('node-dragging');
    try{card.releasePointerCapture(e.pointerId);}catch(_){}
    if(!moved)return;
    const SNAP=20;
    curDx=Math.round(curDx/SNAP)*SNAP;
    curDy=Math.round(curDy/SNAP)*SNAP;
    pushUndo();
    if(curDx===0&&curDy===0){
      delete overrides[id];
      card.style.transform='';
      card.classList.remove('grid-translated');
    }else{
      overrides[id]={dx:curDx,dy:curDy};
      card.style.transform='translate('+curDx+'px,'+curDy+'px)';
    }
    redraw();
    persistState();
  }
  card.addEventListener('pointermove',onMove);
  card.addEventListener('pointerup',onUp);
  card.addEventListener('pointercancel',onUp);
}
/* Legacy HTML5 drag-drop hook on the canvas; kept as a fallback for browsers
   where pointer events fail. Pointer events take priority in grid mode. */
function bindCanvasGridDND(){
  const cc=document.getElementById('chart-canvas-content');if(!cc||cc._gridDND)return;cc._gridDND=true;
  cc.addEventListener('dragover',e=>{
    if(!S.draggingNodeId||!S.gridMode)return;
    e.preventDefault();e.dataTransfer.dropEffect='move';
  });
  cc.addEventListener('drop',e=>{
    if(!S.gridMode||!S.draggingNodeId)return;
    e.preventDefault();
    const id=S.draggingNodeId;
    const card=cc.querySelector('#org-tree [data-drag-id="'+CSS.escape(id)+'"]');
    if(!card){S.draggingNodeId=null;return;}
    const ccRect=cc.getBoundingClientRect();
    const dropX=(e.clientX-ccRect.left)/S.zoom;
    const dropY=(e.clientY-ccRect.top)/S.zoom;
    const cr=card.getBoundingClientRect();
    const cardCenterX=(cr.left+cr.width/2-ccRect.left)/S.zoom;
    const cardCenterY=(cr.top+cr.height/2-ccRect.top)/S.zoom;
    const oldOvr=S.gridOverrides[id]||{dx:0,dy:0};
    let newDx=oldOvr.dx+(dropX-cardCenterX);
    let newDy=oldOvr.dy+(dropY-cardCenterY);
    const SNAP=20;
    newDx=Math.round(newDx/SNAP)*SNAP;newDy=Math.round(newDy/SNAP)*SNAP;
    pushUndo();
    if(newDx===0&&newDy===0)delete S.gridOverrides[id];
    else S.gridOverrides[id]={dx:newDx,dy:newDy};
    S.draggingNodeId=null;
    applyGridOverridesToTree();
    // Wait for the next paint frame so the transform is committed before
    // measuring card rects for the connector redraw.
    requestAnimationFrame(()=>{redrawGridConnectorsFromTree();});
    persistState();
  });
}
/* Multi-select + alignment */
function toggleSelectCard(id,card){
  if(S.selectedIds.has(id)){S.selectedIds.delete(id);card&&card.classList.remove('selected');}
  else{S.selectedIds.add(id);card&&card.classList.add('selected');}
  refreshAlignToolbar();
}
function clearSelection(){
  S.selectedIds.forEach(id=>{const c=document.querySelector('.node-card[data-drag-id="'+CSS.escape(id)+'"]');if(c)c.classList.remove('selected');});
  S.selectedIds.clear();refreshAlignToolbar();
}
function refreshAlignToolbar(){
  const tb=document.getElementById('align-toolbar');if(!tb)return;
  tb.classList.toggle('visible',S.selectedIds.size>=2);
  const ct=document.getElementById('at-count');if(ct)ct.textContent=S.selectedIds.size+' selected';
}
function _measureNatural(card){
  // Measure the card's bounding box without its current translation override
  const t=card.style.transform;card.style.transform='';
  const cr=card.getBoundingClientRect();
  card.style.transform=t;
  return cr;
}
function _gridCtx(){
  // Returns the current grid context (main chart vs Person View) so multi-
  // select tools and connector redraws share one code path.
  const pvOpen=!document.getElementById('person-view-modal').classList.contains('hidden');
  return pvOpen?{
    inPV:true,
    cc:document.getElementById('pv-tree-content'),
    zoom:S.pvZoom||1,
    overrides:S.pvGridOverrides,
    treeSel:'#pv-org-tree',
    gridOn:S.pvGridMode,
    redraw:redrawPVConnectors,
    apply:applyPVGridOverridesToTree,
    toggle:togglePVGrid
  }:{
    inPV:false,
    cc:document.getElementById('chart-canvas-content'),
    zoom:S.zoom,
    overrides:S.gridOverrides,
    treeSel:'#org-tree',
    gridOn:S.gridMode,
    redraw:redrawGridConnectorsFromTree,
    apply:applyGridOverridesToTree,
    toggle:toggleGridMode
  };
}
function alignSelected(direction){
  if(S.selectedIds.size<2)return;
  const ctx=_gridCtx();
  if(!ctx.gridOn)ctx.toggle();
  const cc=ctx.cc;if(!cc)return;
  const ccRect=cc.getBoundingClientRect();
  const items=[];
  S.selectedIds.forEach(id=>{
    const card=cc.querySelector(ctx.treeSel+' [data-drag-id="'+CSS.escape(id)+'"]');if(!card)return;
    const nr=_measureNatural(card);
    items.push({id,nx:(nr.left-ccRect.left)/ctx.zoom,ny:(nr.top-ccRect.top)/ctx.zoom,nw:nr.width/ctx.zoom,nh:nr.height/ctx.zoom});
  });
  if(items.length<2)return;
  pushUndo();
  if(direction==='left'){const t=Math.min(...items.map(i=>i.nx));items.forEach(i=>{ctx.overrides[i.id]={dx:t-i.nx,dy:(ctx.overrides[i.id]||{dy:0}).dy||0};});}
  else if(direction==='right'){const t=Math.max(...items.map(i=>i.nx+i.nw));items.forEach(i=>{ctx.overrides[i.id]={dx:t-(i.nx+i.nw),dy:(ctx.overrides[i.id]||{dy:0}).dy||0};});}
  else if(direction==='center-h'){const avg=items.reduce((s,i)=>s+(i.nx+i.nw/2),0)/items.length;items.forEach(i=>{ctx.overrides[i.id]={dx:avg-(i.nx+i.nw/2),dy:(ctx.overrides[i.id]||{dy:0}).dy||0};});}
  else if(direction==='top'){const t=Math.min(...items.map(i=>i.ny));items.forEach(i=>{ctx.overrides[i.id]={dx:(ctx.overrides[i.id]||{dx:0}).dx||0,dy:t-i.ny};});}
  else if(direction==='bottom'){const t=Math.max(...items.map(i=>i.ny+i.nh));items.forEach(i=>{ctx.overrides[i.id]={dx:(ctx.overrides[i.id]||{dx:0}).dx||0,dy:t-(i.ny+i.nh)};});}
  else if(direction==='middle'){const avg=items.reduce((s,i)=>s+(i.ny+i.nh/2),0)/items.length;items.forEach(i=>{ctx.overrides[i.id]={dx:(ctx.overrides[i.id]||{dx:0}).dx||0,dy:avg-(i.ny+i.nh/2)};});}
  ctx.apply();
  requestAnimationFrame(()=>{ctx.redraw();});
  persistState();
  showToast('Aligned '+items.length+' cards',true);
}
function distributeSelected(axis){
  if(S.selectedIds.size<3){showToast('Select 3 or more cards to distribute');return;}
  const ctx=_gridCtx();
  if(!ctx.gridOn)ctx.toggle();
  const cc=ctx.cc;if(!cc)return;
  const ccRect=cc.getBoundingClientRect();
  const items=[];
  S.selectedIds.forEach(id=>{
    const card=cc.querySelector(ctx.treeSel+' [data-drag-id="'+CSS.escape(id)+'"]');if(!card)return;
    const nr=_measureNatural(card);
    items.push({id,nx:(nr.left-ccRect.left)/ctx.zoom,ny:(nr.top-ccRect.top)/ctx.zoom,nw:nr.width/ctx.zoom,nh:nr.height/ctx.zoom});
  });
  if(items.length<3)return;
  pushUndo();
  if(axis==='h'){
    items.sort((a,b)=>(a.nx+a.nw/2)-(b.nx+b.nw/2));
    const first=items[0].nx+items[0].nw/2,last=items[items.length-1].nx+items[items.length-1].nw/2;
    const step=(last-first)/(items.length-1);
    items.forEach((it,idx)=>{const target=first+step*idx;const cur=it.nx+it.nw/2;ctx.overrides[it.id]={dx:target-cur,dy:(ctx.overrides[it.id]||{dy:0}).dy||0};});
  }else{
    items.sort((a,b)=>(a.ny+a.nh/2)-(b.ny+b.nh/2));
    const first=items[0].ny+items[0].nh/2,last=items[items.length-1].ny+items[items.length-1].nh/2;
    const step=(last-first)/(items.length-1);
    items.forEach((it,idx)=>{const target=first+step*idx;const cur=it.ny+it.nh/2;ctx.overrides[it.id]={dx:(ctx.overrides[it.id]||{dx:0}).dx||0,dy:target-cur};});
  }
  ctx.apply();
  requestAnimationFrame(()=>{ctx.redraw();});
  persistState();
  showToast('Distributed '+items.length+' cards',true);
}
function mkBareCard(node){
  // Reuse mkNodeLI for full card fidelity. Temporarily suppress this node's
  // children in S.childMap so mkNodeLI doesn't recursively build the full subtree
  // (would be O(N^2) across the whole grid).
  const orig=S.childMap[node.id];
  S.childMap[node.id]=[];
  const li=mkNodeLI(node,0);
  S.childMap[node.id]=orig;
  const card=li.querySelector('.node-card');
  const cb=card?card.querySelector('.collapse-btn'):null;
  if(cb)cb.remove();
  return card;
}
function computeGridPositions(visibleNodes){
  const positions={};
  // Default: row = depth - skipDepth + 1, col = order within depth (left-to-right tree-order)
  // BFS from roots to preserve tree ordering
  const skip=S.skipDepth||0;
  const roots=skip>0?visibleNodes.filter(n=>(S.nodeDepth[n.id]||0)===skip):(S.childMap['']||[]);
  const rowsByDepth={};
  const queue=[...roots];const seen=new Set();
  while(queue.length){
    const n=queue.shift();
    if(seen.has(n.id))continue;seen.add(n.id);
    const d=(S.nodeDepth[n.id]||0)-skip+1;
    if(!rowsByDepth[d])rowsByDepth[d]=[];
    rowsByDepth[d].push(n);
    (S.childMap[n.id]||[]).forEach(k=>{if(visibleNodes.find(v=>v.id===k.id))queue.push(k);});
  }
  // Pick up any node that wasn't reached (e.g. orphan)
  visibleNodes.forEach(n=>{if(!seen.has(n.id)){const d=(S.nodeDepth[n.id]||0)-skip+1;if(!rowsByDepth[d])rowsByDepth[d]=[];rowsByDepth[d].push(n);}});
  Object.keys(rowsByDepth).sort((a,b)=>+a-+b).forEach(d=>{
    rowsByDepth[d].forEach((n,i)=>{positions[n.id]={row:+d,col:i+1};});
  });
  // Apply overrides (but only for nodes still present)
  Object.entries(S.gridOverrides).forEach(([id,pos])=>{
    if(visibleNodes.find(n=>n.id===id)&&pos&&pos.row>=1&&pos.col>=1)positions[id]={row:pos.row,col:pos.col};
  });
  return positions;
}
function renderGrid(){
  const grid=document.getElementById('org-grid');if(!grid)return;
  grid.innerHTML='';
  // Determine visible nodes (mirroring renderChart's logic)
  let visible;
  if(S.skipDepth>0){
    // include all nodes at skipDepth and below (their descendants are also visible via tree)
    visible=S.viewData.filter(n=>(S.nodeDepth[n.id]||0)>=S.skipDepth);
  }else{
    visible=S.viewData.slice();
  }
  if(!visible.length){grid.innerHTML='<div class="no-data" style="margin:24px">No nodes to display.</div>';updateStats(visible);return;}
  const positions=computeGridPositions(visible);
  let maxCol=1,maxRow=1;
  Object.values(positions).forEach(p=>{if(p.col>maxCol)maxCol=p.col;if(p.row>maxRow)maxRow=p.row;});
  // Build cells: pre-create maxRow×maxCol grid of empty cells (drop targets), then fill in occupied ones
  grid.style.gridTemplateColumns='repeat('+(maxCol+2)+',280px)';
  grid.style.gridTemplateRows='repeat('+(maxRow+1)+',minmax(200px,auto))';
  const occupied={};Object.entries(positions).forEach(([id,p])=>{occupied[p.row+'_'+p.col]=id;});
  for(let r=1;r<=maxRow+1;r++){
    for(let c=1;c<=maxCol+2;c++){
      const cell=document.createElement('div');
      cell.className='grid-cell';
      cell.dataset.row=r;cell.dataset.col=c;
      cell.style.gridRow=r;cell.style.gridColumn=c;
      bindGridCellDND(cell);
      const occId=occupied[r+'_'+c];
      if(occId){const n=visible.find(v=>v.id===occId);if(n){const card=mkBareCard(n);cell.appendChild(card);cell.dataset.id=occId;}}
      grid.appendChild(cell);
    }
  }
  drawGridConnectors(positions,visible);
  updateStats(visible);
  setTimeout(fitToScreen,160);
}
function bindGridCellDND(cell){
  cell.addEventListener('dragover',e=>{
    if(!S.draggingNodeId||!S.gridMode)return;
    e.preventDefault();e.dataTransfer.dropEffect='move';
    cell.classList.add('drop-target-cell');
  });
  cell.addEventListener('dragleave',()=>cell.classList.remove('drop-target-cell','drop-target-cell-bad'));
  cell.addEventListener('drop',e=>{
    cell.classList.remove('drop-target-cell','drop-target-cell-bad');
    if(!S.gridMode||!S.draggingNodeId)return;
    e.preventDefault();e.stopPropagation();
    const draggedId=S.draggingNodeId;
    const newRow=parseInt(cell.dataset.row),newCol=parseInt(cell.dataset.col);
    // If cell is occupied by another card, swap positions
    const occupantId=cell.dataset.id;
    pushUndo();
    if(occupantId&&occupantId!==draggedId){
      // swap: dragged → cell's coords; occupant → dragged's old coords
      const draggedOld=S.gridOverrides[draggedId]||findAutoPos(draggedId);
      S.gridOverrides[draggedId]={row:newRow,col:newCol};
      if(draggedOld)S.gridOverrides[occupantId]={row:draggedOld.row,col:draggedOld.col};
    }else{
      S.gridOverrides[draggedId]={row:newRow,col:newCol};
    }
    S.draggingNodeId=null;
    renderGrid();persistState();
    showToast('Card repositioned',true);
  });
}
function findAutoPos(id){
  const visible=S.viewData;const positions=computeGridPositions(visible);
  // Strip overrides for this id to get the natural pos
  const tmp={...S.gridOverrides};delete tmp[id];
  const orig=S.gridOverrides;S.gridOverrides=tmp;
  const p=computeGridPositions(visible)[id];
  S.gridOverrides=orig;
  return p||null;
}
function drawGridConnectors(positions,visible){
  const svg=document.getElementById('grid-svg');if(!svg)return;
  svg.innerHTML='';
  const grid=document.getElementById('org-grid');if(!grid)return;
  const gRect=grid.getBoundingClientRect();
  // Use an SVG sized to the grid
  svg.setAttribute('width',(grid.scrollWidth||grid.offsetWidth)+'px');
  svg.setAttribute('height',(grid.scrollHeight||grid.offsetHeight)+'px');
  svg.style.left=grid.offsetLeft+'px';
  svg.style.top=grid.offsetTop+'px';
  // For each visible node with a manager that's also visible, draw a connector
  const byId=Object.fromEntries(visible.map(n=>[n.id,n]));
  visible.forEach(n=>{
    if(!n.manager||!byId[n.manager])return;
    const childCell=grid.querySelector('[data-row="'+positions[n.id].row+'"][data-col="'+positions[n.id].col+'"]');
    const parentPos=positions[n.manager];if(!parentPos)return;
    const parentCell=grid.querySelector('[data-row="'+parentPos.row+'"][data-col="'+parentPos.col+'"]');
    if(!childCell||!parentCell)return;
    const cc=childCell.querySelector('.node-card'),pc=parentCell.querySelector('.node-card');
    if(!cc||!pc)return;
    const cr=cc.getBoundingClientRect(),pr=pc.getBoundingClientRect();
    const x1=(pr.left+pr.width/2-gRect.left)/S.zoom;
    const y1=(pr.bottom-gRect.top)/S.zoom;
    const x2=(cr.left+cr.width/2-gRect.left)/S.zoom;
    const y2=(cr.top-gRect.top)/S.zoom;
    const midY=(y1+y2)/2;
    const path=document.createElementNS('http://www.w3.org/2000/svg','path');
    path.setAttribute('d','M '+x1+' '+y1+' C '+x1+' '+midY+', '+x2+' '+midY+', '+x2+' '+y2);
    path.setAttribute('stroke','#94a3b8');path.setAttribute('stroke-width','2');
    path.setAttribute('fill','none');path.setAttribute('opacity','0.7');
    svg.appendChild(path);
  });
}
function applyPersisted(d){
  if(!d)return false;
  // Validate that saved colMap columns still exist in the current file
  const cols=new Set(S.columns);
  if(!d.colMap||!cols.has(d.colMap.empId)||!cols.has(d.colMap.empName))return false;
  S.colMap=d.colMap;
  S.managerOverrides=d.managerOverrides||{};
  S.removedIds=new Set(d.removedIds||[]);
  S.cardSlots=d.cardSlots||S.cardSlots;
  S.cardAccent=d.cardAccent||S.cardAccent;
  S.empTypeCol=d.empTypeCol||'';
  S.empTypeLabels=d.empTypeLabels||{active:'',vacant:'',resigned:''};
  S.empTypeColors=d.empTypeColors||S.empTypeColors;
  S.filterCols=(d.filterCols||[]).filter(c=>cols.has(c));
  S.activeFilters=Object.fromEntries(Object.entries(d.activeFilters||{}).filter(([k])=>cols.has(k)));
  S.photoMatchCol=cols.has(d.photoMatchCol||'')?d.photoMatchCol:S.colMap.empId;
  S.photoSize=d.photoSize||80;
  S.photoShape=d.photoShape||'circle';
  S.photoPlacement=d.photoPlacement||'top';
  S.summaryField1=d.summaryField1||'';
  S.summaryField2=d.summaryField2||'';
  S.summaryField3=d.summaryField3||'';
  S.chartBgColor=d.chartBgColor||'#f1f5f9';
  S.transparentExport=!!d.transparentExport;
  S.skipDepth=d.skipDepth||0;
  S.managerMode=!!d.managerMode;
  S.gridMode=!!d.gridMode;
  // Migrate v1 grid overrides ({row,col}) to v2 ({dx,dy}) by discarding them — v2
  // uses pixel translations measured at runtime, not row/col cell coordinates.
  S.gridOverrides={};
  Object.entries(d.gridOverrides||{}).forEach(([id,o])=>{
    if(o&&typeof o.dx==='number'&&typeof o.dy==='number')S.gridOverrides[id]=o;
  });
  S.gridShowLines=d.gridShowLines!==false;
  // PV grid mode + overrides survive reload too.
  S.pvGridMode=!!d.pvGridMode;
  S.pvGridOverrides={};
  Object.entries(d.pvGridOverrides||{}).forEach(([id,o])=>{
    if(o&&typeof o.dx==='number'&&typeof o.dy==='number')S.pvGridOverrides[id]=o;
  });
  S.maxPerRow=(d.maxPerRow==='auto')?'auto':(parseInt(d.maxPerRow)||6);
  S.pvMaxPerRow=(d.pvMaxPerRow==='auto'||d.pvMaxPerRow==null)?'auto':(parseInt(d.pvMaxPerRow)||'auto');
  buildEmpTypeMap();
  return true;
}

// Each binding is wrapped so a single missing element never aborts the whole init script.
function _bind(id,event,fn){try{const el=document.getElementById(id);if(el)el.addEventListener(event,fn);else console.warn('init: missing element #'+id);}catch(ex){console.error('init bind',id,ex);}}
// (file-input change is also wired inline as a defensive double-bind in case this _bind is missed)
_bind('file-input','change',function(e){if(e.target.files[0])handleFile(e.target.files[0]);});
_bind('photo-folder-input','change',function(e){if(e.target.files.length)loadFromFileInput(e.target.files);});

/* Demo data: lets the user try the app without uploading anything. Also serves as
   a smoke test that the upload pipeline is working. */
function loadDemoData(){
  const rows=[
    {'Employee ID':'E001','Employee Name':'Alex Rivera','Manager ID':'',Department:'Executive',Title:'CEO'},
    {'Employee ID':'E002','Employee Name':'Priya Shah','Manager ID':'E001',Department:'Engineering',Title:'VP Engineering'},
    {'Employee ID':'E003','Employee Name':'Marcus Liu','Manager ID':'E001',Department:'Sales',Title:'VP Sales'},
    {'Employee ID':'E004','Employee Name':'Sara Okafor','Manager ID':'E001',Department:'People',Title:'VP People'},
    {'Employee ID':'E005','Employee Name':'Diego Fernández','Manager ID':'E002',Department:'Engineering',Title:'Eng Manager'},
    {'Employee ID':'E006','Employee Name':'Yuki Tanaka','Manager ID':'E002',Department:'Engineering',Title:'Eng Manager'},
    {'Employee ID':'E007','Employee Name':'Aanya Mehta','Manager ID':'E005',Department:'Engineering',Title:'Senior Engineer'},
    {'Employee ID':'E008','Employee Name':'Tom Becker','Manager ID':'E005',Department:'Engineering',Title:'Engineer'},
    {'Employee ID':'E009','Employee Name':'Rina Patel','Manager ID':'E005',Department:'Engineering',Title:'Engineer'},
    {'Employee ID':'E010','Employee Name':'Hari Sundar','Manager ID':'E006',Department:'Engineering',Title:'Senior Engineer'},
    {'Employee ID':'E011','Employee Name':'Jane Park','Manager ID':'E006',Department:'Engineering',Title:'Engineer'},
    {'Employee ID':'E012','Employee Name':'Lucas Brown','Manager ID':'E003',Department:'Sales',Title:'Sales Manager'},
    {'Employee ID':'E013','Employee Name':'Eva Stone','Manager ID':'E012',Department:'Sales',Title:'AE'},
    {'Employee ID':'E014','Employee Name':'Omar Hassan','Manager ID':'E012',Department:'Sales',Title:'AE'},
    {'Employee ID':'E015','Employee Name':'Maya Chen','Manager ID':'E004',Department:'People',Title:'Recruiter'}
  ];
  initData(rows);
}

/* Library-status indicator: makes it obvious if a CDN library failed to load
   (e.g. corp firewall blocked cdnjs). Without this, an XLSX upload would just
   silently fail when XLSX.read runs. */
function refreshLibStatus(){
  const el=document.getElementById('lib-status');if(!el)return;
  const libs=[
    {name:'Papa (CSV)',ok:typeof Papa!=='undefined'},
    {name:'XLSX (Excel)',ok:typeof XLSX!=='undefined'},
    {name:'JSZip',ok:typeof JSZip!=='undefined'},
    {name:'html2canvas',ok:typeof html2canvas!=='undefined'}
  ];
  const allOk=libs.every(l=>l.ok);
  el.innerHTML=libs.map(l=>'<span style="color:'+(l.ok?'#059669':'#dc2626')+'">'+(l.ok?'✓':'✗')+' '+l.name+'</span>').join('');
  if(!allOk){
    el.innerHTML+='<div style="width:100%;text-align:center;color:#dc2626;margin-top:6px">⚠ One or more libraries failed to load — upload may not work. Check your network/firewall and reload.</div>';
  }
}
// Run after a tick so CDN scripts have a chance to evaluate
setTimeout(refreshLibStatus,200);
setTimeout(refreshLibStatus,1500);
(function(){
  const dz=document.getElementById('upload-dropzone');if(!dz)return;
  dz.addEventListener('dragover',function(e){e.preventDefault();dz.classList.add('drag-over');});
  dz.addEventListener('dragleave',function(){dz.classList.remove('drag-over');});
  dz.addEventListener('drop',function(e){e.preventDefault();dz.classList.remove('drag-over');const f=e.dataTransfer.files[0];if(f)handleFile(f);});
})();
_bind('reassign-modal','click',function(e){if(e.target===e.currentTarget)closeReassignModal();});
_bind('person-view-modal','click',function(e){if(e.target===e.currentTarget)closePV();});
_bind('dq-modal','click',function(e){if(e.target===e.currentTarget)closeDataQualityModal();});
_bind('insights-modal','click',function(e){if(e.target===e.currentTarget)closeInsightsModal();});
</script>
</body>
</html>'''

components.html(APP_HTML, height=900, scrolling=False)
