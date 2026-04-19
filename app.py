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
<style>
/* ═══════════════ TOKENS ═══════════════ */
:root {
  --bg:           #ffffff;
  --bg2:          #f8fafc;
  --bg3:          #f1f5f9;
  --bg4:          #e8edf5;
  --border:       #e2e8f0;
  --border2:      #cbd5e1;
  --text:         #0f172a;
  --text2:        #475569;
  --text3:        #94a3b8;
  --accent:       #4f46e5;
  --accent2:      #6366f1;
  --accent-light: #eef2ff;
  --accent-mid:   #c7d2fe;
  --success:      #059669;
  --success-light:#d1fae5;
  --warning:      #d97706;
  --warning-light:#fef3c7;
  --danger:       #dc2626;
  --shadow-xs:    0 1px 2px rgba(0,0,0,0.05);
  --shadow-sm:    0 1px 4px rgba(0,0,0,0.07), 0 1px 2px rgba(0,0,0,0.04);
  --shadow-md:    0 4px 16px rgba(0,0,0,0.08), 0 2px 4px rgba(0,0,0,0.04);
  --shadow-lg:    0 12px 40px rgba(0,0,0,0.1), 0 4px 12px rgba(0,0,0,0.05);
  --r:            10px;
  --r-lg:         14px;
  --r-xl:         18px;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body {
  height: 100%;
  background: var(--bg);
  font-family: 'Plus Jakarta Sans', sans-serif;
  color: var(--text);
  overflow: hidden;
  font-size: 14px;
}
body { display: flex; flex-direction: column; }

/* ═══════════════ TOP NAV ═══════════════ */
.topnav {
  flex-shrink: 0;
  height: 54px;
  background: var(--bg);
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  padding: 0 24px;
  gap: 12px;
  z-index: 100;
  box-shadow: var(--shadow-xs);
}
.brand {
  font-family: 'Syne', sans-serif;
  font-weight: 800;
  font-size: 1.05rem;
  color: var(--text);
  display: flex;
  align-items: center;
  gap: 9px;
  letter-spacing: -0.02em;
  flex-shrink: 0;
}
.brand-icon {
  width: 30px; height: 30px;
  background: linear-gradient(135deg, var(--accent), var(--accent2));
  border-radius: 9px;
  display: flex; align-items: center; justify-content: center;
  font-size: 0.9rem;
  box-shadow: 0 3px 10px rgba(79,70,229,0.3);
}
.nav-sep { width: 1px; height: 26px; background: var(--border); flex-shrink: 0; }
.step-trail {
  display: flex; align-items: center; gap: 2px;
  flex: 1; justify-content: center;
}
.step-item {
  display: flex; align-items: center; gap: 6px;
  font-size: 0.76rem; font-weight: 600;
  color: var(--text3);
  transition: color 0.2s;
  white-space: nowrap;
  padding: 4px 6px;
  border-radius: 6px;
}
.step-item.active { color: var(--accent); background: var(--accent-light); }
.step-item.done   { color: var(--success); }
.step-dot {
  width: 22px; height: 22px;
  border-radius: 50%;
  background: var(--bg3);
  border: 2px solid var(--border2);
  display: flex; align-items: center; justify-content: center;
  font-size: 0.62rem; font-weight: 800; color: var(--text3);
  transition: all 0.2s; flex-shrink: 0;
}
.step-item.active .step-dot { background: var(--accent); border-color: var(--accent); color: #fff; }
.step-item.done   .step-dot { background: var(--success); border-color: var(--success); color: #fff; font-size: 0.7rem; }
.step-arrow { color: var(--border2); font-size: 0.8rem; margin: 0 1px; }

/* ═══════════════ MAIN ═══════════════ */
.main { flex: 1; overflow: hidden; position: relative; }
.screen {
  position: absolute; inset: 0;
  overflow-y: auto;
  display: flex; flex-direction: column;
  padding: 32px 36px;
  background: var(--bg);
  opacity: 0; pointer-events: none;
  transform: translateX(18px);
  transition: opacity 0.22s ease, transform 0.22s ease;
}
.screen.active { opacity: 1; pointer-events: auto; transform: translateX(0); }
#screen-chart { padding: 0; overflow: hidden; }

/* ═══════════════ UPLOAD ═══════════════ */
.upload-center {
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  min-height: 100%; gap: 24px; padding: 24px;
}
.upload-hero { text-align: center; }
.upload-hero h1 {
  font-family: 'Syne', sans-serif; font-weight: 800;
  font-size: 2rem; color: var(--text);
  letter-spacing: -0.03em; margin-bottom: 8px;
}
.upload-hero p { color: var(--text3); font-size: 0.9rem; max-width: 400px; line-height: 1.6; }
.upload-zone {
  width: 520px; max-width: 100%;
  border: 2px dashed var(--border2);
  border-radius: var(--r-xl);
  padding: 48px 32px;
  text-align: center; cursor: pointer;
  transition: all 0.2s; background: var(--bg);
  position: relative;
}
.upload-zone:hover, .upload-zone.drag-over {
  border-color: var(--accent); background: var(--accent-light);
}
.upload-zone input[type="file"] {
  position: absolute; inset: 0; opacity: 0;
  cursor: pointer; width: 100%; height: 100%;
}
.upload-emoji { font-size: 2.8rem; margin-bottom: 14px; display: block; }
.upload-zone h3 { font-weight: 800; font-size: 1.15rem; color: var(--text); margin-bottom: 6px; }
.upload-zone p  { font-size: 0.84rem; color: var(--text3); line-height: 1.5; }
.upload-zone p span { color: var(--accent); font-weight: 700; }
.info-cards { display: flex; gap: 14px; width: 520px; max-width: 100%; }
.info-card {
  flex: 1; background: var(--bg2);
  border: 1px solid var(--border);
  border-radius: var(--r-lg); padding: 16px;
}
.info-card-title {
  font-size: 0.72rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
  color: var(--text3); margin-bottom: 10px;
}
.info-card-row {
  font-size: 0.8rem; color: var(--text2); font-weight: 500;
  padding: 3px 0; display: flex; align-items: center; gap: 6px;
}
.info-card-row::before { content: ''; width: 5px; height: 5px; background: var(--border2); border-radius: 50%; flex-shrink: 0; }

/* ═══════════════ SHARED SECTION HEADER ═══════════════ */
.section-header { margin-bottom: 24px; }
.section-title {
  font-family: 'Syne', sans-serif; font-weight: 700;
  font-size: 1.45rem; color: var(--text);
  letter-spacing: -0.02em; margin-bottom: 4px;
}
.section-sub { font-size: 0.84rem; color: var(--text2); }

/* ═══════════════ COLUMN MAPPING ═══════════════ */
.detected-chips { display: flex; flex-wrap: wrap; gap: 7px; margin-bottom: 24px; }
.col-chip {
  display: inline-flex; align-items: center; gap: 7px;
  padding: 5px 11px;
  background: var(--bg2); border: 1.5px solid var(--border);
  border-radius: 999px;
  font-size: 0.76rem; font-weight: 600; color: var(--text2);
}
.col-chip .chip-sample {
  color: var(--text3); font-size: 0.7rem; font-style: italic;
}
.map-grid {
  display: grid; grid-template-columns: repeat(3, 1fr);
  gap: 14px; max-width: 820px; margin-bottom: 28px;
}
.map-card {
  background: var(--bg);
  border: 1.5px solid var(--border);
  border-radius: var(--r-lg); padding: 16px;
  transition: border-color 0.2s, box-shadow 0.2s;
}
.map-card:focus-within { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(79,70,229,0.08); }
.map-card-label {
  font-size: 0.7rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
  color: var(--text3); margin-bottom: 8px;
  display: flex; align-items: center; gap: 7px;
}
.badge-req {
  background: #fee2e2; color: var(--danger);
  padding: 1px 7px; border-radius: 999px;
  font-size: 0.6rem; font-weight: 700;
}
.badge-opt {
  background: var(--bg3); color: var(--text3);
  padding: 1px 7px; border-radius: 999px;
  font-size: 0.6rem; font-weight: 700;
}
.map-select {
  width: 100%; background: var(--bg3);
  border: 1.5px solid var(--border);
  border-radius: 8px; padding: 8px 10px;
  font-size: 0.84rem; font-weight: 600;
  color: var(--text); font-family: 'Plus Jakarta Sans', sans-serif;
  outline: none; cursor: pointer; appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");
  background-repeat: no-repeat; background-position: right 10px center;
  transition: border-color 0.15s;
}
.map-select:focus { border-color: var(--accent); background-color: var(--bg); }
.map-hint { font-size: 0.72rem; color: var(--text3); margin-top: 6px; }
.data-preview-table {
  width: 100%; max-width: 820px;
  border-collapse: collapse;
  margin-bottom: 28px; font-size: 0.78rem;
}
.data-preview-table th {
  background: var(--bg3); padding: 7px 12px;
  text-align: left; font-weight: 700; color: var(--text2);
  border: 1px solid var(--border); font-size: 0.72rem;
  text-transform: uppercase; letter-spacing: 0.04em;
}
.data-preview-table td {
  padding: 6px 12px; border: 1px solid var(--border);
  color: var(--text2); max-width: 160px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.data-preview-table tr:nth-child(even) td { background: var(--bg2); }

/* ═══════════════ CARD DESIGNER ═══════════════ */
.card-design-layout {
  display: grid; grid-template-columns: 260px 1fr;
  gap: 24px; flex: 1; min-height: 0;
}
.fields-panel {
  background: var(--bg2); border: 1.5px solid var(--border);
  border-radius: var(--r-lg); padding: 18px;
  overflow-y: auto;
}
.fields-panel-title {
  font-size: 0.68rem; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.07em;
  color: var(--text3); margin-bottom: 12px;
}
.fields-section { margin-bottom: 16px; }
.fields-section-label {
  font-size: 0.66rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
  color: var(--text3); margin-bottom: 8px;
  padding-bottom: 4px; border-bottom: 1px solid var(--border);
}
.field-chip {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 6px 11px;
  background: var(--bg); border: 1.5px solid var(--border);
  border-radius: 8px; font-size: 0.78rem; font-weight: 600;
  color: var(--text); cursor: grab; margin: 3px;
  transition: all 0.15s; user-select: none;
}
.field-chip:hover { border-color: var(--accent); color: var(--accent); background: var(--accent-light); }
.field-chip.placed { background: var(--accent-light); border-color: var(--accent-mid); color: var(--accent); }
.field-chip.dragging { opacity: 0.4; transform: scale(0.93); }
.drag-icon { color: var(--text3); font-size: 0.7rem; cursor: grab; }
.auto-chip-icon { font-size: 0.8rem; }

.card-preview-area {
  display: flex; flex-direction: column; align-items: flex-start; gap: 14px;
}
.preview-label {
  font-size: 0.68rem; font-weight: 800;
  text-transform: uppercase; letter-spacing: 0.07em;
  color: var(--text3);
}
/* The actual card */
.preview-card {
  width: 300px;
  background: var(--bg);
  border: 2px solid var(--border);
  border-top: 4px solid var(--accent);
  border-radius: var(--r-lg);
  box-shadow: var(--shadow-md);
}
.preview-card-header {
  padding: 9px 12px;
  background: var(--bg2);
  border-bottom: 1px solid var(--border);
  border-radius: 12px 12px 0 0;
  display: flex; justify-content: space-between; align-items: center; gap: 8px;
}
.preview-card-body { padding: 12px 14px; }
.preview-name {
  font-weight: 800; font-size: 0.92rem; color: var(--text);
  margin-bottom: 7px; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis;
}
.preview-name-fixed {
  display: flex; align-items: center; gap: 6px;
  background: var(--bg3); border: 1px dashed var(--border2);
  border-radius: 6px; padding: 6px 10px;
}
.preview-name-fixed span { font-size: 0.8rem; color: var(--text2); font-weight: 600; }
.preview-name-fixed .lock { font-size: 0.75rem; opacity: 0.5; }
.preview-card-footer {
  padding: 8px 12px;
  border-top: 1px solid var(--border);
  border-radius: 0 0 12px 12px;
  background: var(--bg2);
  display: flex; justify-content: space-between; align-items: center; gap: 8px;
}
/* Drop zones */
.card-zone {
  flex: 1; min-height: 30px; min-width: 60px;
  border: 2px dashed var(--border2);
  border-radius: 7px; padding: 4px 8px;
  font-size: 0.7rem; color: var(--text3);
  display: flex; align-items: center; justify-content: center;
  transition: all 0.15s; position: relative;
  cursor: default;
}
.card-zone .zone-ph { opacity: 0.6; font-style: italic; }
.card-zone.drop-target { border-color: var(--accent); background: var(--accent-light); }
.card-zone.filled {
  border-style: solid; border-color: var(--accent-mid);
  background: var(--accent-light);
  flex-direction: column; gap: 2px; align-items: flex-start; justify-content: center;
}
.zone-field { font-weight: 700; font-size: 0.7rem; color: var(--accent); }
.zone-val { font-size: 0.68rem; color: var(--text2); font-style: italic; }
.zone-remove {
  position: absolute; top: 3px; right: 4px;
  font-size: 0.6rem; cursor: pointer;
  opacity: 0.5; line-height: 1;
}
.zone-remove:hover { opacity: 1; }
.card-zone-subtitle { width: 100%; }
.preview-hint {
  font-size: 0.76rem; color: var(--text3); max-width: 300px; line-height: 1.5;
}

/* ═══════════════ FILTER SETUP ═══════════════ */
.filter-setup { max-width: 640px; }
.filter-chips { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 24px; }
.filter-chip {
  padding: 7px 15px;
  background: var(--bg3); border: 1.5px solid var(--border);
  border-radius: 999px; font-size: 0.82rem; font-weight: 600;
  color: var(--text2); cursor: pointer; transition: all 0.15s; user-select: none;
}
.filter-chip:hover { border-color: var(--accent); color: var(--accent); background: var(--accent-light); }
.filter-chip.selected { background: var(--accent); border-color: var(--accent); color: #fff; }
.filter-chip.selected:hover { background: #4338ca; }
.filter-counter {
  font-size: 0.72rem; color: var(--text3); font-weight: 600; margin-bottom: 12px;
}
.filter-preview-box {
  background: var(--bg2); border: 1px solid var(--border);
  border-radius: var(--r); padding: 16px; margin-top: 8px;
}
.fpr-row {
  display: flex; align-items: flex-start; gap: 10px;
  padding: 8px 0; border-bottom: 1px solid var(--border);
  font-size: 0.8rem;
}
.fpr-row:last-child { border-bottom: none; }
.fpr-col { font-weight: 700; color: var(--text); min-width: 130px; }
.fpr-vals { display: flex; flex-wrap: wrap; gap: 5px; flex: 1; }
.fv-pill {
  padding: 2px 9px; background: var(--bg); border: 1px solid var(--border);
  border-radius: 999px; font-size: 0.72rem; font-weight: 500; color: var(--text2);
}

/* ═══════════════ BUTTONS ═══════════════ */
.btn {
  padding: 9px 20px; border-radius: var(--r);
  font-size: 0.84rem; font-weight: 700; cursor: pointer; border: none;
  transition: all 0.15s;
  display: inline-flex; align-items: center; gap: 7px;
  font-family: 'Plus Jakarta Sans', sans-serif; line-height: 1;
  white-space: nowrap;
}
.btn-primary {
  background: var(--accent); color: #fff;
  box-shadow: 0 4px 14px rgba(79,70,229,0.3);
}
.btn-primary:hover { background: #4338ca; transform: translateY(-1px); box-shadow: 0 6px 20px rgba(79,70,229,0.4); }
.btn-ghost {
  background: transparent; color: var(--text2);
  border: 1.5px solid var(--border);
}
.btn-ghost:hover { background: var(--bg3); color: var(--text); border-color: var(--border2); }
.btn-sm { padding: 6px 13px; font-size: 0.78rem; border-radius: 8px; }
.btn-row { display: flex; gap: 10px; margin-top: 28px; }

/* ═══════════════ CHART SCREEN ═══════════════ */
.chart-toolbar {
  flex-shrink: 0; height: 52px;
  background: var(--bg); border-bottom: 1px solid var(--border);
  display: flex; align-items: center; padding: 0 16px; gap: 8px;
  box-shadow: var(--shadow-xs); position: relative; z-index: 20;
}
.stats-bar {
  flex-shrink: 0; height: 34px;
  background: var(--bg2); border-bottom: 1px solid var(--border);
  display: flex; align-items: center; padding: 0 18px; gap: 18px;
  font-size: 0.73rem;
}
.stat-item { display: flex; align-items: center; gap: 6px; color: var(--text3); font-weight: 600; }
.stat-item strong { color: var(--text); font-weight: 800; }
.stat-dot { width: 6px; height: 6px; border-radius: 50%; background: var(--accent); }
.filter-bar {
  flex-shrink: 0; background: var(--bg);
  border-bottom: 1px solid var(--border);
  padding: 7px 18px; display: flex;
  align-items: center; gap: 10px; flex-wrap: wrap; min-height: 44px;
}
.filter-dropdown-wrap {
  display: flex; align-items: center; gap: 6px; font-size: 0.79rem;
}
.filter-dropdown-label { font-weight: 700; color: var(--text2); }
.filter-dropdown {
  background: var(--bg);
  border: 1.5px solid var(--border); border-radius: 8px;
  padding: 5px 28px 5px 10px; font-size: 0.79rem; font-weight: 600;
  color: var(--text); font-family: 'Plus Jakarta Sans', sans-serif;
  cursor: pointer; outline: none; appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2394a3b8'/%3E%3C/svg%3E");
  background-repeat: no-repeat; background-position: right 8px center;
  transition: border-color 0.15s;
}
.filter-dropdown:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(79,70,229,0.1); }
.filter-active-tag {
  display: inline-flex; align-items: center; gap: 5px;
  padding: 3px 10px; background: var(--accent-light);
  border: 1px solid var(--accent-mid); border-radius: 999px;
  font-size: 0.73rem; font-weight: 700; color: var(--accent);
}
.filter-active-tag .x { cursor: pointer; opacity: 0.6; font-size: 0.65rem; }
.filter-active-tag .x:hover { opacity: 1; }

/* Chart canvas */
.chart-canvas-wrap {
  flex: 1; overflow: auto;
  background: var(--bg3); cursor: grab;
  position: relative;
}
.chart-canvas-wrap:active { cursor: grabbing; }
.chart-canvas-wrap::before {
  content: ''; position: fixed; inset: 0;
  background-image:
    linear-gradient(rgba(148,163,184,0.12) 1px, transparent 1px),
    linear-gradient(90deg, rgba(148,163,184,0.12) 1px, transparent 1px);
  background-size: 32px 32px;
  pointer-events: none; z-index: 0;
}
.chart-canvas-content {
  display: inline-block;
  padding: 56px 80px 120px 80px;
  transform-origin: top left;
  position: relative; z-index: 1;
}
.org-tree { display: inline-block; }

/* Org tree connectors */
.org-tree ul {
  padding-top: 24px; position: relative;
  list-style: none; display: flex; justify-content: center;
}
.org-tree li {
  display: table-cell; vertical-align: top;
  text-align: center; position: relative;
  padding: 24px 7px 0 7px;
}
.org-tree li::before, .org-tree li::after {
  content: ''; position: absolute;
  top: 0; right: 50%;
  border-top: 2px solid #cbd5e1;
  width: 50%; height: 24px;
}
.org-tree li::after { right: auto; left: 50%; border-left: 2px solid #cbd5e1; }
.org-tree li:only-child::before,
.org-tree li:only-child::after { display: none; }
.org-tree li:first-child::before,
.org-tree li:last-child::after { display: none; }
.org-tree li:first-child::after  { border-radius: 6px 0 0 0; }
.org-tree li:last-child::before  { border-radius: 0 6px 0 0; }
.org-tree ul ul::before {
  content: ''; position: absolute; top: 0; left: 50%;
  border-left: 2px solid #cbd5e1; height: 24px;
}

/* Node card */
.node-card {
  display: inline-block; width: 220px;
  background: var(--bg);
  border: 1.5px solid var(--border);
  border-top: 3px solid var(--accent);
  border-radius: var(--r-lg);
  cursor: pointer; text-align: left;
  transition: transform 0.15s, box-shadow 0.15s, border-color 0.15s;
  box-shadow: var(--shadow-sm); position: relative;
  font-family: 'Plus Jakarta Sans', sans-serif;
}
.node-card:hover {
  transform: translateY(-3px);
  box-shadow: 0 8px 28px rgba(0,0,0,0.12), 0 0 0 2px rgba(79,70,229,0.12);
  border-color: var(--accent); z-index: 10;
}
.node-card.highlighted {
  border-color: var(--warning) !important;
  border-top-color: var(--warning) !important;
  box-shadow: 0 0 0 3px rgba(217,119,6,0.2), 0 8px 24px rgba(0,0,0,0.1) !important;
}
.node-card.collapsed-node { opacity: 0.65; }
.node-card .ncard-header {
  padding: 7px 11px 6px;
  background: var(--bg2);
  border-bottom: 1px solid var(--border);
  border-radius: 12px 12px 0 0;
  display: flex; justify-content: space-between; align-items: center; gap: 6px;
}
.ncard-badge {
  font-size: 0.59rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.04em;
  background: var(--accent-light); color: var(--accent);
  padding: 2px 8px; border-radius: 999px;
  max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  border: 1px solid var(--accent-mid);
}
.ncard-badge2 {
  font-size: 0.65rem; font-weight: 700;
  color: var(--text3); white-space: nowrap;
}
.ncard-body { padding: 10px 12px 8px; }
.ncard-name {
  font-size: 0.86rem; font-weight: 800; color: var(--text);
  margin-bottom: 3px; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis; letter-spacing: -0.01em;
}
.ncard-sub {
  font-size: 0.73rem; color: var(--text2); line-height: 1.4;
  display: -webkit-box; -webkit-line-clamp: 2;
  -webkit-box-orient: vertical; overflow: hidden;
}
.ncard-footer {
  padding: 6px 12px;
  border-top: 1px solid var(--border);
  border-radius: 0 0 12px 12px;
  background: var(--bg2);
  display: flex; justify-content: space-between; align-items: center;
  font-size: 0.67rem; font-weight: 600;
}
.ncard-fl { color: var(--text3); }
.ncard-fr {
  background: var(--bg3); padding: 2px 8px;
  border-radius: 999px; color: var(--text3); font-size: 0.64rem;
}
.ncard-fr.has-r { background: var(--accent-light); color: var(--accent); border: 1px solid var(--accent-mid); }
.collapse-btn {
  position: absolute; bottom: -11px; left: 50%;
  transform: translateX(-50%);
  width: 22px; height: 22px;
  background: var(--bg); border: 1.5px solid var(--border2);
  border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  cursor: pointer; font-size: 0.58rem; color: var(--text3);
  transition: all 0.15s; z-index: 5; box-shadow: var(--shadow-xs);
}
.collapse-btn:hover { background: var(--accent); border-color: var(--accent); color: #fff; }
li.collapsed > ul { display: none; }

/* Search */
.search-wrap { position: relative; flex: 1; max-width: 260px; }
.search-icon { position: absolute; left: 10px; top: 50%; transform: translateY(-50%); font-size: 0.8rem; pointer-events: none; opacity: 0.45; }
#chart-search {
  width: 100%; background: var(--bg2);
  border: 1.5px solid var(--border); border-radius: 8px;
  padding: 6px 10px 6px 29px;
  font-size: 0.8rem; font-weight: 500; color: var(--text);
  font-family: 'Plus Jakarta Sans', sans-serif; outline: none;
  transition: border-color 0.15s;
}
#chart-search:focus { border-color: var(--accent); background: var(--bg); }
#chart-search::placeholder { color: var(--text3); }
#chart-search-results {
  position: absolute; top: calc(100% + 5px); left: 0; right: 0;
  background: var(--bg); border: 1px solid var(--border);
  border-radius: var(--r); box-shadow: var(--shadow-lg);
  max-height: 260px; overflow-y: auto; z-index: 999; display: none;
}
#chart-search-results.visible { display: block; }
.sr-item {
  padding: 9px 13px; cursor: pointer;
  border-bottom: 1px solid var(--border); transition: background 0.1s;
}
.sr-item:last-child { border-bottom: none; }
.sr-item:hover { background: var(--bg3); }
.sr-name { font-weight: 700; font-size: 0.82rem; }
.sr-sub  { font-size: 0.72rem; color: var(--text3); margin-top: 2px; }

/* Zoom strip */
.zoom-strip {
  display: flex; align-items: center; gap: 1px;
  background: var(--bg2); border-radius: 8px;
  padding: 2px; border: 1.5px solid var(--border);
}
.btn-zoom {
  background: transparent; border: none; border-radius: 6px;
  width: 26px; height: 26px; cursor: pointer;
  font-size: 0.85rem; font-weight: 700; color: var(--text2);
  font-family: 'Plus Jakarta Sans', sans-serif;
  display: flex; align-items: center; justify-content: center;
  transition: background 0.12s;
}
.btn-zoom:hover { background: var(--bg3); color: var(--text); }
.zoom-label {
  font-size: 0.72rem; font-weight: 800; color: var(--text);
  min-width: 42px; text-align: center;
  font-variant-numeric: tabular-nums;
}

/* Export overlay */
.export-overlay {
  position: fixed; inset: 0; z-index: 9999;
  background: rgba(255,255,255,0.85);
  backdrop-filter: blur(6px);
  display: flex; flex-direction: column;
  align-items: center; justify-content: center; gap: 12px;
}
.export-spinner {
  width: 44px; height: 44px;
  border: 3px solid var(--border2);
  border-top-color: var(--accent); border-radius: 50%;
  animation: spin 0.7s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }
.no-data {
  padding: 40px; color: var(--text3); font-size: 0.9rem; font-weight: 600;
  background: var(--bg); border: 1.5px solid var(--border);
  border-radius: var(--r-lg); max-width: 440px;
}
</style>
</head>
<body>

<!-- ══════════════ TOP NAV ══════════════ -->
<nav class="topnav">
  <div class="brand">
    <div class="brand-icon">🏢</div>
    OrgDesign Pro
  </div>
  <div class="nav-sep"></div>
  <div class="step-trail">
    <div class="step-item active" id="nav-step-upload">
      <div class="step-dot">1</div><span>Upload</span>
    </div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-map">
      <div class="step-dot">2</div><span>Map Columns</span>
    </div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-card">
      <div class="step-dot">3</div><span>Design Card</span>
    </div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-filter">
      <div class="step-dot">4</div><span>Set Filters</span>
    </div>
    <div class="step-arrow">›</div>
    <div class="step-item" id="nav-step-chart">
      <div class="step-dot">5</div><span>Org Chart</span>
    </div>
  </div>
</nav>

<!-- ══════════════ MAIN ══════════════ -->
<main class="main">

  <!-- ── Screen 1: Upload ── -->
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
          <div class="info-card-title">➕ Optional Columns</div>
          <div class="info-card-row">Designation, Grade, Level</div>
          <div class="info-card-row">Location, City, Region</div>
          <div class="info-card-row">Function, Sub Function</div>
        </div>
      </div>
    </div>
  </div>

  <!-- ── Screen 2: Column Mapping ── -->
  <div class="screen" id="screen-map">
    <div class="section-header">
      <div class="section-title">Map Your Columns</div>
      <div class="section-sub">We detected <span id="col-count">0</span> columns. Auto-mapped where possible — review and adjust below.</div>
    </div>
    <div style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);margin-bottom:9px">Detected Columns</div>
    <div class="detected-chips" id="detected-columns"></div>
    <div class="map-grid">
      <div class="map-card">
        <div class="map-card-label">👤 Employee ID <span class="badge-req">Required</span></div>
        <select class="map-select" id="map-empId"></select>
        <div class="map-hint">Unique identifier for each person</div>
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

  <!-- ── Screen 3: Card Designer ── -->
  <div class="screen" id="screen-card">
    <div class="section-header" style="margin-bottom:18px">
      <div class="section-title">Design Your Card</div>
      <div class="section-sub">Drag fields from the left panel into the card zones. The name is locked to your mapped column.</div>
    </div>
    <div class="card-design-layout">
      <div class="fields-panel">
        <div class="fields-panel-title">Available Fields</div>
        <div id="card-fields-panel"></div>
      </div>
      <div class="card-preview-area">
        <div class="preview-label">Live Card Preview</div>
        <div id="card-preview"></div>
        <div class="preview-hint">
          Drag a field chip onto any zone. Drop another on it to swap. Click ✕ to clear a zone.
        </div>
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" onclick="goTo('map')">← Back</button>
      <button class="btn btn-primary" onclick="confirmCardDesign()">Continue to Filters →</button>
    </div>
  </div>

  <!-- ── Screen 4: Filter Setup ── -->
  <div class="screen" id="screen-filter">
    <div class="section-header">
      <div class="section-title">Set Up Filters</div>
      <div class="section-sub">Choose up to 3 columns to use as filters — for example, Location, Function, or Sub Function.</div>
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

  <!-- ── Screen 5: Chart View ── -->
  <div class="screen" id="screen-chart">
    <!-- Toolbar -->
    <div class="chart-toolbar">
      <button class="btn btn-ghost btn-sm" onclick="goTo('filter')">← Setup</button>
      <div style="width:1px;height:22px;background:var(--border);flex-shrink:0"></div>
      <div class="search-wrap">
        <span class="search-icon">🔍</span>
        <input id="chart-search" type="text" placeholder="Search by name or ID…" autocomplete="off"/>
        <div id="chart-search-results"></div>
      </div>
      <div style="width:1px;height:22px;background:var(--border);flex-shrink:0"></div>
      <div class="zoom-strip">
        <button class="btn-zoom" onclick="zoomBy(-0.1)">−</button>
        <span class="zoom-label" id="zoom-level">100%</span>
        <button class="btn-zoom" onclick="zoomBy(0.1)">+</button>
        <button class="btn-zoom" onclick="fitToScreen(true)" title="Fit to screen">⊡</button>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="centerView()">🧭</button>
      <button class="btn btn-ghost btn-sm" onclick="expandAll()">⊞ Expand</button>
      <button class="btn btn-ghost btn-sm" onclick="collapseAll()">⊟ Collapse</button>
      <div style="flex:1"></div>
      <button class="btn btn-ghost btn-sm" onclick="downloadCSV()">💾 CSV</button>
      <button class="btn btn-primary btn-sm" onclick="exportPNG()">🖼️ Save PNG</button>
    </div>
    <!-- Stats -->
    <div class="stats-bar">
      <div class="stat-item"><div class="stat-dot"></div><strong id="stat-total">—</strong>&nbsp;employees</div>
      <div class="stat-item"><strong id="stat-roots">—</strong>&nbsp;root nodes</div>
      <div class="stat-item"><strong id="stat-vis">—</strong>&nbsp;visible cards</div>
      <div class="stat-item" id="stat-filtered" style="display:none;color:var(--warning)">⚠️ Filtered view</div>
    </div>
    <!-- Filter bar -->
    <div class="filter-bar" id="filter-bar" style="display:none"></div>
    <!-- Canvas -->
    <div class="chart-canvas-wrap" id="chart-canvas-wrap">
      <div class="chart-canvas-content" id="chart-canvas-content">
        <div class="org-tree" id="org-tree"></div>
      </div>
    </div>
  </div>

</main>

<script>
// ════════════════════════════════════════════════
// APP STATE
// ════════════════════════════════════════════════
const S = {
  rawRows:      [],
  columns:      [],
  colSamples:   {},
  colMap:       { empId: '', empName: '', managerId: '' },
  cardSlots:    { badge1: '', badge2: '', subtitle: '', footer1: '', footer2: '' },
  filterCols:   [],
  activeFilters:{},
  viewData:     [],
  zoom:         1,
  highlighted:  null,
  draggingField:null,
};

// ════════════════════════════════════════════════
// NAVIGATION
// ════════════════════════════════════════════════
function goTo(step) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  const el = document.getElementById('screen-' + step);
  if (el) el.classList.add('active');

  const order = ['upload','map','card','filter','chart'];
  const cur = order.indexOf(step);
  order.forEach((s, i) => {
    const el = document.getElementById('nav-step-' + s);
    if (!el) return;
    el.className = 'step-item' + (i < cur ? ' done' : i === cur ? ' active' : '');
    const dot = el.querySelector('.step-dot');
    if (dot) dot.textContent = i < cur ? '✓' : String(i + 1);
  });

  if (step === 'chart') {
    setTimeout(() => initPan(), 80);
    setTimeout(() => initSearch(), 80);
  }
}

// ════════════════════════════════════════════════
// FILE HANDLING
// ════════════════════════════════════════════════
function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: r => initData(r.data),
      error:    e => alert('CSV error: ' + e.message)
    });
  } else if (['xlsx','xls'].includes(ext)) {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      initData(XLSX.utils.sheet_to_json(ws, { defval: '' }));
    };
    reader.readAsArrayBuffer(file);
  } else {
    alert('Please upload a CSV or Excel (.xlsx/.xls) file.');
  }
}

function initData(rows) {
  S.rawRows = rows;
  S.columns = rows.length ? Object.keys(rows[0]) : [];
  S.colSamples = {};
  S.columns.forEach(col => {
    S.colSamples[col] = [...new Set(
      rows.slice(0,25).map(r => String(r[col]||'').trim()).filter(v=>v&&v!=='undefined'&&v!=='null')
    )].slice(0,3);
  });
  S.colMap = autoDetect(S.columns);
  buildMapScreen();
  goTo('map');
}

function autoDetect(cols) {
  const lc = cols.map(c => c.toLowerCase());
  const find = patterns => {
    for (const p of patterns) {
      const i = lc.findIndex(c => c.includes(p));
      if (i >= 0) return cols[i];
    }
    return '';
  };
  return {
    empId:     find(['employee code','emp code','emp id','employee id','empcode','empid','staff id','staff code','person id','worker id']),
    empName:   find(['employee name','emp name','full name','person name','staff name','worker name','name']),
    managerId: find(['l1 manager code','l1 manager','manager code','manager id','reports to','supervisor','mgr code','mgrid']),
  };
}

// ════════════════════════════════════════════════
// SCREEN 2 — COLUMN MAP
// ════════════════════════════════════════════════
function buildMapScreen() {
  document.getElementById('col-count').textContent = S.columns.length;

  // Chips
  const chips = document.getElementById('detected-columns');
  chips.innerHTML = S.columns.map(c => {
    const s = S.colSamples[c].join(', ');
    return `<div class="col-chip">📋 ${esc(c)}${s?`<span class="chip-sample">${esc(s)}</span>`:''}`;
  }).join('');

  // Selects
  const blankOpt = '<option value="">— select —</option>';
  const opts = blankOpt + S.columns.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
  ['empId','empName','managerId'].forEach(k => {
    const sel = document.getElementById('map-'+k);
    sel.innerHTML = opts;
    sel.value = S.colMap[k] || '';
  });

  // Data preview table
  const wrap = document.getElementById('data-preview-wrap');
  const preview = S.rawRows.slice(0,3);
  if (!preview.length) { wrap.innerHTML = ''; return; }
  let html = '<table class="data-preview-table"><thead><tr>' +
    S.columns.map(c=>`<th>${esc(c)}</th>`).join('') + '</tr></thead><tbody>';
  preview.forEach(row => {
    html += '<tr>' + S.columns.map(c=>`<td title="${esc(String(row[c]||''))}">${esc(String(row[c]||'').substring(0,22))}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function confirmColumnMap() {
  S.colMap.empId     = document.getElementById('map-empId').value;
  S.colMap.empName   = document.getElementById('map-empName').value;
  S.colMap.managerId = document.getElementById('map-managerId').value;
  if (!S.colMap.empId || !S.colMap.empName) {
    alert('Please map the Employee ID and Employee Name columns.');
    return;
  }
  buildCardScreen();
  goTo('card');
}

// ════════════════════════════════════════════════
// SCREEN 3 — CARD DESIGNER
// ════════════════════════════════════════════════
const AUTO_FIELDS = [
  { id: '__auto_reports__',  icon: '📊', label: 'Direct Reports',  desc: 'Count of direct reports' },
  { id: '__auto_teamsize__', icon: '👥', label: 'Total Team Size',  desc: 'All descendants count' },
];

function buildCardScreen() {
  const core = new Set([S.colMap.empId, S.colMap.empName, S.colMap.managerId].filter(Boolean));
  const available = S.columns.filter(c => !core.has(c));

  const panel = document.getElementById('card-fields-panel');
  panel.innerHTML =
    '<div class="fields-section">' +
      '<div class="fields-section-label">Column Fields</div>' +
      (available.length
        ? available.map(f =>
            `<div class="field-chip" draggable="true" data-field="${esc(f)}"
              ondragstart="onDragStart(event)" ondragend="onDragEnd(event)">
              <span class="drag-icon">⠿</span>${esc(f)}</div>`
          ).join('')
        : '<div style="font-size:0.78rem;color:var(--text3);font-style:italic">No extra columns</div>'
      ) +
    '</div>' +
    '<div class="fields-section">' +
      '<div class="fields-section-label">Auto-Calculated</div>' +
      AUTO_FIELDS.map(f =>
        `<div class="field-chip" draggable="true" data-field="${f.id}"
          ondragstart="onDragStart(event)" ondragend="onDragEnd(event)"
          title="${f.desc}">
          <span class="drag-icon">⠿</span><span class="auto-chip-icon">${f.icon}</span>${f.label}
        </div>`
      ).join('') +
    '</div>';

  // Default: auto-reports in footer2
  if (!S.cardSlots.footer2) S.cardSlots.footer2 = '__auto_reports__';

  renderCardPreview();
  syncChipStates();
}

function onDragStart(e) {
  S.draggingField = e.currentTarget.dataset.field;
  e.currentTarget.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
}
function onDragEnd(e) { e.currentTarget.classList.remove('dragging'); }

function onZoneDragOver(e)  { e.preventDefault(); e.currentTarget.classList.add('drop-target'); }
function onZoneDragLeave(e) { e.currentTarget.classList.remove('drop-target'); }

function onZoneDrop(e, zone) {
  e.preventDefault();
  e.currentTarget.classList.remove('drop-target');
  if (!S.draggingField) return;
  // If field was already in another zone, clear it
  Object.keys(S.cardSlots).forEach(z => { if (S.cardSlots[z] === S.draggingField) S.cardSlots[z] = ''; });
  S.cardSlots[zone] = S.draggingField;
  S.draggingField = null;
  renderCardPreview();
  syncChipStates();
}

function clearZone(zone) {
  S.cardSlots[zone] = '';
  renderCardPreview();
  syncChipStates();
}

function syncChipStates() {
  const placed = new Set(Object.values(S.cardSlots).filter(Boolean));
  document.querySelectorAll('.field-chip').forEach(c => {
    c.classList.toggle('placed', placed.has(c.dataset.field));
  });
}

function fieldLabel(id) {
  if (!id) return '';
  const af = AUTO_FIELDS.find(f => f.id === id);
  if (af) return af.icon + ' ' + af.label;
  return id;
}

function fieldSampleVal(id) {
  if (!id) return '';
  if (id === '__auto_reports__') return '12';
  if (id === '__auto_teamsize__') return '48';
  const row = S.rawRows.find(r => r[id]) || S.rawRows[0] || {};
  return String(row[id] || 'Sample').substring(0, 22);
}

function zoneHtml(zoneId, placeholder, extraClass = '') {
  const v = S.cardSlots[zoneId];
  const dAttrs = `ondragover="onZoneDragOver(event)" ondragleave="onZoneDragLeave(event)" ondrop="onZoneDrop(event,'${zoneId}')"`;
  if (v) {
    return `<div class="card-zone filled ${extraClass}" ${dAttrs}>
      <span class="zone-field">${esc(fieldLabel(v))}</span>
      <span class="zone-val">${esc(fieldSampleVal(v))}</span>
      <span class="zone-remove" onclick="clearZone('${zoneId}')">✕</span>
    </div>`;
  }
  return `<div class="card-zone ${extraClass}" ${dAttrs}>
    <span class="zone-ph">${placeholder}</span>
  </div>`;
}

function renderCardPreview() {
  const sampleName = (() => {
    const row = S.rawRows.find(r => r[S.colMap.empName]) || S.rawRows[0] || {};
    return String(row[S.colMap.empName] || 'Employee Name').substring(0, 26);
  })();

  document.getElementById('card-preview').innerHTML = `
    <div class="preview-card">
      <div class="preview-card-header">
        ${zoneHtml('badge1','+ Badge Left')}
        ${zoneHtml('badge2','+ Badge Right')}
      </div>
      <div class="preview-card-body">
        <div class="preview-name-fixed">
          <span class="lock">🔒</span>
          <span>${esc(sampleName)}</span>
        </div>
        <div style="margin-top:6px">${zoneHtml('subtitle','+ Subtitle / Designation','card-zone-subtitle')}</div>
      </div>
      <div class="preview-card-footer">
        ${zoneHtml('footer1','+ Footer Left')}
        ${zoneHtml('footer2','+ Footer Right')}
      </div>
    </div>`;
}

function confirmCardDesign() {
  buildFilterScreen();
  goTo('filter');
}

// ════════════════════════════════════════════════
// SCREEN 4 — FILTER SETUP
// ════════════════════════════════════════════════
function buildFilterScreen() {
  const core = new Set([S.colMap.empId, S.colMap.empName, S.colMap.managerId].filter(Boolean));
  const filterable = S.columns.filter(c => !core.has(c));

  const picker = document.getElementById('filter-chip-picker');
  picker.innerHTML = filterable.map(col =>
    `<div class="filter-chip ${S.filterCols.includes(col)?'selected':''}"
      data-col="${esc(col)}" onclick="toggleFilterCol('${esc(col)}')">
      ${esc(col)}</div>`
  ).join('');

  renderFilterPreview();
}

function toggleFilterCol(col) {
  if (S.filterCols.includes(col)) {
    S.filterCols = S.filterCols.filter(c => c !== col);
  } else if (S.filterCols.length < 3) {
    S.filterCols.push(col);
  } else {
    // Bump oldest, add new
    S.filterCols.shift();
    S.filterCols.push(col);
  }
  document.querySelectorAll('.filter-chip').forEach(c => {
    c.classList.toggle('selected', S.filterCols.includes(c.dataset.col));
  });
  renderFilterPreview();
}

function renderFilterPreview() {
  const counter = document.getElementById('filter-counter');
  counter.textContent = `${S.filterCols.length} of 3 filters selected`;

  const area = document.getElementById('filter-preview-area');
  if (!S.filterCols.length) {
    area.innerHTML = `<div style="font-size:0.82rem;color:var(--text3);padding:12px 0">
      No filters selected — the full org chart will be displayed.</div>`;
    return;
  }
  area.innerHTML = `<div class="filter-preview-box">
    ${S.filterCols.map(col => {
      const vals = [...new Set(
        S.rawRows.map(r => String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined')
      )].sort().slice(0, 10);
      return `<div class="fpr-row">
        <span class="fpr-col">${esc(col)}</span>
        <div class="fpr-vals">
          ${vals.map(v=>`<span class="fv-pill">${esc(v)}</span>`).join('')}
          ${vals.length >= 10 ? '<span style="font-size:0.7rem;color:var(--text3)">+ more</span>' : ''}
        </div>
      </div>`;
    }).join('')}
  </div>`;
}

function launchChart() {
  S.activeFilters = {};
  buildViewData();
  buildFilterBar();
  renderChart();
  goTo('chart');
}

// ════════════════════════════════════════════════
// SCREEN 5 — DATA PREP
// ════════════════════════════════════════════════
function buildViewData() {
  const { empId, empName, managerId } = S.colMap;
  let nodes = S.rawRows.map(row => {
    const id  = String(row[empId]||'').replace(/\.0$/,'').trim();
    const mgr = managerId ? String(row[managerId]||'').replace(/\.0$/,'').trim() : '';
    const node = { id, name: String(row[empName]||'Unknown'), manager: mgr };
    S.columns.forEach(col => { node[col] = String(row[col]||''); });
    return node;
  }).filter(n => n.id);

  const validIds = new Set(nodes.map(n => n.id));
  nodes.forEach(n => { if (!validIds.has(n.manager)) n.manager = ''; });

  // Apply active filters: keep matching + all their ancestors
  const hasFilter = Object.values(S.activeFilters).some(v => v);
  if (hasFilter) {
    const matching = new Set(
      nodes.filter(n => Object.entries(S.activeFilters).every(([c,v]) => !v || n[c] === v))
           .map(n => n.id)
    );
    const byId = Object.fromEntries(nodes.map(n => [n.id, n]));
    const keep = new Set(matching);
    matching.forEach(id => {
      let cur = byId[id];
      while (cur && cur.manager && byId[cur.manager]) {
        keep.add(cur.manager);
        cur = byId[cur.manager];
      }
    });
    nodes = nodes.filter(n => keep.has(n.id));
  }

  S.viewData = nodes;
}

function childrenOf(id) { return S.viewData.filter(n => n.manager === id); }

function countDescendants(id, memo) {
  if (!memo) memo = {};
  if (memo[id] !== undefined) return memo[id];
  const kids = childrenOf(id);
  memo[id] = kids.reduce((s, k) => s + 1 + countDescendants(k.id, memo), 0);
  return memo[id];
}

// ════════════════════════════════════════════════
// FILTER BAR
// ════════════════════════════════════════════════
function buildFilterBar() {
  const bar = document.getElementById('filter-bar');
  if (!S.filterCols.length) { bar.style.display = 'none'; return; }
  bar.style.display = 'flex';

  // Collect unique values from raw data for each filter col
  const allVals = {};
  S.filterCols.forEach(col => {
    allVals[col] = [...new Set(
      S.rawRows.map(r => String(r[col]||'').trim()).filter(v=>v&&v!=='null'&&v!=='undefined')
    )].sort();
  });

  bar.innerHTML =
    '<span style="font-size:0.68rem;font-weight:800;text-transform:uppercase;letter-spacing:0.07em;color:var(--text3);flex-shrink:0">Filters</span>' +
    S.filterCols.map(col => {
      const cur = S.activeFilters[col] || '';
      const opts = `<option value="">All ${esc(col)}</option>` +
        allVals[col].map(v => `<option value="${esc(v)}" ${cur===v?'selected':''}>${esc(v)}</option>`).join('');
      return `<div class="filter-dropdown-wrap">
        <span class="filter-dropdown-label">${esc(col)}</span>
        <select class="filter-dropdown" onchange="applyFilter('${esc(col)}',this.value)">${opts}</select>
      </div>`;
    }).join('') +
    (Object.values(S.activeFilters).some(v=>v)
      ? `<button class="btn btn-ghost btn-sm" onclick="clearAllFilters()" style="margin-left:auto">✕ Clear All</button>`
      : '');
}

function applyFilter(col, val) {
  if (val) S.activeFilters[col] = val;
  else delete S.activeFilters[col];
  buildViewData();
  renderChart();
  buildFilterBar();
}

function clearAllFilters() {
  S.activeFilters = {};
  buildViewData();
  renderChart();
  buildFilterBar();
}

// ════════════════════════════════════════════════
// ORG CHART RENDER
// ════════════════════════════════════════════════
function getSlotVal(node, slot) {
  const f = S.cardSlots[slot];
  if (!f) return '';
  if (f === '__auto_reports__') return childrenOf(node.id).length + ' reports';
  if (f === '__auto_teamsize__') return countDescendants(node.id) + ' people';
  return String(node[f]||'').substring(0,28);
}

function renderChart() {
  const tree = document.getElementById('org-tree');
  tree.innerHTML = '';

  const roots = S.viewData.filter(n => !n.manager);
  if (!roots.length) {
    tree.innerHTML = `<div class="no-data">No root nodes found.<br>Check column mapping or clear active filters.</div>`;
    updateStats(roots);
    return;
  }

  const ul = document.createElement('ul');
  roots.forEach(r => ul.appendChild(mkNodeLI(r)));
  tree.appendChild(ul);

  updateStats(roots);
  clearTimeout(window._fit);
  window._fit = setTimeout(() => fitToScreen(true), 180);
}

function mkNodeLI(node) {
  const li = document.createElement('li');
  li.dataset.id = node.id;

  const card = document.createElement('div');
  card.className = 'node-card' + (node.id === S.highlighted ? ' highlighted' : '');

  const reports  = childrenOf(node.id).length;
  const badge1   = getSlotVal(node, 'badge1');
  const badge2   = getSlotVal(node, 'badge2');
  const subtitle = getSlotVal(node, 'subtitle');
  const footer1  = getSlotVal(node, 'footer1') || node.id.substring(0,14);
  const footer2  = getSlotVal(node, 'footer2');
  const hasR     = reports > 0;

  card.innerHTML =
    '<div class="ncard-header">' +
      (badge1 ? `<span class="ncard-badge" title="${esc(badge1)}">${esc(badge1)}</span>` : '<span></span>') +
      (badge2 ? `<span class="ncard-badge2">${esc(badge2)}</span>` : '') +
    '</div>' +
    '<div class="ncard-body">' +
      `<div class="ncard-name" title="${esc(node.name)}">${esc(node.name)}</div>` +
      (subtitle ? `<div class="ncard-sub" title="${esc(subtitle)}">${esc(subtitle)}</div>` : '') +
    '</div>' +
    '<div class="ncard-footer">' +
      `<span class="ncard-fl">${esc(footer1)}</span>` +
      (footer2
        ? `<span class="ncard-fr has-r">${esc(footer2)}</span>`
        : `<span class="ncard-fr${hasR?' has-r':''}">${reports} ${reports===1?'report':'reports'}</span>`) +
    '</div>';

  if (hasR) {
    const cb = document.createElement('div');
    cb.className = 'collapse-btn';
    cb.innerHTML = '▾';
    cb.title = 'Collapse / expand';
    cb.addEventListener('click', e => { e.stopPropagation(); toggleCollapse(li, cb); });
    card.appendChild(cb);
  }

  li.appendChild(card);
  const kids = childrenOf(node.id);
  if (kids.length) {
    const ul = document.createElement('ul');
    kids.forEach(k => ul.appendChild(mkNodeLI(k)));
    li.appendChild(ul);
  }
  return li;
}

function toggleCollapse(li, btn) {
  li.classList.toggle('collapsed');
  const c = li.classList.contains('collapsed');
  btn.innerHTML = c ? '▸' : '▾';
  btn.style.color = c ? 'var(--warning)' : '';
  li.querySelector('.node-card')?.classList.toggle('collapsed-node', c);
  setTimeout(() => updateStats(), 60);
}

function expandAll() {
  document.querySelectorAll('li.collapsed').forEach(li => {
    li.classList.remove('collapsed');
    li.querySelector('.node-card')?.classList.remove('collapsed-node');
    const b = li.querySelector('.collapse-btn');
    if (b) { b.innerHTML = '▾'; b.style.color = ''; }
  });
  setTimeout(() => updateStats(), 60);
}

function collapseAll() {
  document.querySelectorAll('li').forEach(li => {
    if (!li.parentElement?.parentElement?.closest('li')) return;
    if (li.querySelector(':scope > ul')) {
      li.classList.add('collapsed');
      li.querySelector('.node-card')?.classList.add('collapsed-node');
      const b = li.querySelector('.collapse-btn');
      if (b) { b.innerHTML = '▸'; b.style.color = 'var(--warning)'; }
    }
  });
  setTimeout(() => updateStats(), 60);
}

function updateStats(roots) {
  if (!roots) roots = S.viewData.filter(n => !n.manager);
  document.getElementById('stat-total').textContent = S.viewData.length;
  document.getElementById('stat-roots').textContent = roots.length;
  document.getElementById('stat-vis').textContent   = document.querySelectorAll('.node-card').length;
  const hasFilter = Object.values(S.activeFilters).some(v=>v);
  document.getElementById('stat-filtered').style.display = hasFilter ? 'flex' : 'none';
}

// ════════════════════════════════════════════════
// ZOOM & PAN
// ════════════════════════════════════════════════
function cwrap() { return document.getElementById('chart-canvas-wrap'); }
function ccontent() { return document.getElementById('chart-canvas-content'); }

function applyZoom(z) {
  S.zoom = Math.max(0.1, Math.min(3, z));
  ccontent().style.transform = 'scale(' + S.zoom + ')';
  document.getElementById('zoom-level').textContent = Math.round(S.zoom*100)+'%';
}
function zoomBy(d) { applyZoom(S.zoom + d); }

function fitToScreen(andCenter) {
  requestAnimationFrame(() => {
    const tree = document.getElementById('org-tree');
    const wrap = cwrap();
    if (!tree||!wrap) return;
    const tw=tree.scrollWidth, th=tree.scrollHeight;
    const aw=wrap.clientWidth-100, ah=wrap.clientHeight-100;
    if (tw<10||th<10) return;
    applyZoom(Math.max(0.12, Math.min(1, aw/tw, ah/th)));
    if (andCenter) setTimeout(centerView, 70);
  });
}

function centerView() {
  const wrap = cwrap();
  const tree = document.getElementById('org-tree');
  if (!wrap||!tree) return;
  const sw = tree.scrollWidth * S.zoom;
  wrap.scrollLeft = Math.max(0, (sw - wrap.clientWidth)/2);
  wrap.scrollTop  = 0;
}

let _panning=false, _px,_py,_psl,_pst;
function initPan() {
  const wrap = cwrap();
  if (!wrap) return;
  wrap.onmousedown = e => {
    if (e.target.closest('.node-card,.collapse-btn')) return;
    _panning=true; _px=e.clientX; _py=e.clientY;
    _psl=wrap.scrollLeft; _pst=wrap.scrollTop;
    wrap.style.cursor='grabbing';
  };
  window.onmousemove = e => {
    if (!_panning) return;
    cwrap().scrollLeft = _psl-(e.clientX-_px);
    cwrap().scrollTop  = _pst-(e.clientY-_py);
  };
  window.onmouseup = () => { _panning=false; if(cwrap()) cwrap().style.cursor=''; };
  wrap.addEventListener('wheel', e => {
    if (e.ctrlKey||e.metaKey) { e.preventDefault(); zoomBy(e.deltaY<0?0.08:-0.08); }
  }, {passive:false});
}

// ════════════════════════════════════════════════
// SEARCH
// ════════════════════════════════════════════════
function initSearch() {
  const input = document.getElementById('chart-search');
  const box   = document.getElementById('chart-search-results');
  if (!input) return;
  input.addEventListener('input', function() {
    const q = this.value.trim().toLowerCase();
    if (!q) { box.classList.remove('visible'); return; }
    const hits = S.viewData.filter(n =>
      n.name.toLowerCase().includes(q) || n.id.toLowerCase().includes(q)
    ).slice(0,8);
    box.innerHTML = hits.length
      ? hits.map(n =>
          `<div class="sr-item" onclick="highlightNode('${esc(n.id)}')">
            <div class="sr-name">${esc(n.name)}</div>
            <div class="sr-sub">${esc(getSlotVal(n,'subtitle')||n.id)}</div>
          </div>`).join('')
      : '<div class="sr-item" style="color:var(--text3);font-size:0.8rem">No results</div>';
    box.classList.add('visible');
  });
  document.addEventListener('click', e => {
    if (!e.target.closest('.search-wrap')) box.classList.remove('visible');
  });
}

function highlightNode(id) {
  document.querySelectorAll('.node-card.highlighted').forEach(c=>c.classList.remove('highlighted'));
  S.highlighted = id;
  expandAll();
  const li = document.querySelector(`li[data-id="${CSS.escape(id)}"]`);
  if (li) {
    const card = li.querySelector('.node-card');
    if (card) {
      card.classList.add('highlighted');
      setTimeout(() => {
        const r = card.getBoundingClientRect();
        const w = cwrap(); const wr = w.getBoundingClientRect();
        w.scrollTo({ left: w.scrollLeft+(r.left-wr.left)-wr.width/2+r.width/2,
                     top:  w.scrollTop +(r.top -wr.top )-wr.height/2+r.height/2, behavior:'smooth' });
      }, 80);
    }
  }
  document.getElementById('chart-search').value = '';
  document.getElementById('chart-search-results').classList.remove('visible');
}

// ════════════════════════════════════════════════
// EXPORT
// ════════════════════════════════════════════════
function csvEsc(v) { return '"' + String(v??'').replace(/"/g,'""') + '"'; }

function downloadCSV() {
  const cols = [S.colMap.empId, S.colMap.empName, S.colMap.managerId,
    ...S.columns.filter(c => c!==S.colMap.empId && c!==S.colMap.empName && c!==S.colMap.managerId)
  ].filter(Boolean);
  let csv = cols.map(csvEsc).join(',') + '\n';
  S.viewData.forEach(n => { csv += cols.map(c=>csvEsc(n[c]||'')).join(',') + '\n'; });
  const url = URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
  const a = document.createElement('a'); a.href=url; a.download='orgchart_export.csv'; a.click();
  URL.revokeObjectURL(url);
}

async function exportPNG() {
  const overlay = document.createElement('div');
  overlay.className = 'export-overlay';
  overlay.innerHTML = '<div class="export-spinner"></div><div style="font-weight:700;font-size:0.9rem;margin-top:10px">Rendering org chart…</div><div style="font-size:0.75rem;color:var(--text3);margin-top:4px">Large charts may take a moment</div>';
  document.body.appendChild(overlay);

  const savedZoom = S.zoom; applyZoom(1);
  await new Promise(r=>setTimeout(r,100));

  const stage = document.createElement('div');
  stage.style.cssText = 'position:absolute;top:0;left:0;visibility:hidden;pointer-events:none;background:#f8fafc;padding:48px;display:inline-block;white-space:nowrap;';

  const cloned = document.getElementById('org-tree').cloneNode(true);
  cloned.querySelectorAll('*').forEach(el => { el.style.overflow='visible'; el.style.webkitLineClamp='unset'; el.classList.remove('collapsed'); });
  cloned.querySelectorAll('ul').forEach(ul => ul.style.display='');
  cloned.querySelectorAll('.collapse-btn').forEach(b=>b.remove());
  stage.appendChild(cloned);
  document.body.appendChild(stage);

  await new Promise(r=>setTimeout(r,120));
  if (document.fonts?.ready) await document.fonts.ready;
  await new Promise(r=>setTimeout(r,80));

  try {
    const canvas = await html2canvas(stage, {
      backgroundColor:'#f8fafc', scale:2, useCORS:true, logging:false, allowTaint:true,
      width:stage.scrollWidth, height:stage.scrollHeight,
      windowWidth:stage.scrollWidth+200, windowHeight:stage.scrollHeight+200, x:0, y:0
    });
    canvas.toBlob(blob => {
      const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href=url; a.download=`orgchart_${stamp}.png`; a.click();
      URL.revokeObjectURL(url);
    }, 'image/png');
  } catch(err) {
    alert('Export failed: ' + err.message);
  } finally {
    stage.remove(); overlay.remove(); applyZoom(savedZoom);
  }
}

// ════════════════════════════════════════════════
// UTILS
// ════════════════════════════════════════════════
function esc(s) {
  return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}

// ════════════════════════════════════════════════
// INIT
// ════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  const zone  = document.getElementById('upload-dropzone');
  const input = document.getElementById('file-input');

  zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop',      e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  input.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
});
</script>
</body>
</html>"""

components.html(APP_HTML, height=870, scrolling=False)
