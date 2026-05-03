import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Nodely",
    page_icon="🔵",
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
<title>Nodely — Org charts that breathe</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<style>
/* ═══════════════════════════════════════════════════════════════════════
 * NODELY DESIGN TOKENS
 * ═══════════════════════════════════════════════════════════════════════ */
:root {
  --ink: #0b0b14;
  --ink-2: #2b2b3a;
  --muted: #6b6b80;
  --muted-2: #9a9aae;
  --line: #ececf2;
  --line-2: #f4f4f8;
  --bg: #ffffff;
  --bg-soft: #fafaf9;
  --bg-soft-2: #f5f5f7;

  --accent: #4f46e5;
  --accent-2: #6366f1;
  --accent-soft: #eef0ff;
  --accent-mid: #c7c9ff;

  --pop-pink: #ff6b9d;
  --pop-yellow: #ffd166;
  --pop-green: #10b981;
  --pop-cyan: #0891b2;

  --warn: #d97706;
  --warn-soft: #fef3c7;
  --danger: #dc2626;
  --danger-soft: #fee2e2;
  --success: #10b981;
  --success-soft: #d1fae5;

  --r-sm: 10px;
  --r-md: 14px;
  --r-lg: 20px;
  --r-xl: 28px;

  --shadow-soft: 0 1px 0 rgba(11,11,20,0.04), 0 8px 24px -12px rgba(11,11,20,0.08);
  --shadow-pop:  0 1px 0 rgba(11,11,20,0.04), 0 20px 60px -20px rgba(79,70,229,0.25);
  --shadow-card: 0 1px 0 rgba(11,11,20,0.04), 0 18px 40px -18px rgba(11,11,20,0.15);

  --spring: cubic-bezier(0.34, 1.56, 0.64, 1);
  --ease:   cubic-bezier(0.22, 1, 0.36, 1);
}

* { box-sizing: border-box; margin: 0; padding: 0; }
html, body {
  height: 100%;
  background: var(--bg);
  color: var(--ink);
  font-family: 'Inter', system-ui, -apple-system, sans-serif;
  -webkit-font-smoothing: antialiased;
  font-size: 14px;
  overflow: hidden;
}
body { display: flex; flex-direction: column; }

.display { font-family: 'Space Grotesk', sans-serif; letter-spacing: -0.02em; }
.mono    { font-family: 'JetBrains Mono', ui-monospace, monospace; }
button   { font-family: inherit; cursor: pointer; }
input, select { font-family: inherit; }
::selection { background: var(--accent); color: #fff; }

/* ═══════════════════════════════════════════════════════════════════════
 * TOP NAV (Nodely-style sticky header)
 * ═══════════════════════════════════════════════════════════════════════ */
.topnav {
  flex-shrink: 0;
  height: 60px;
  background: rgba(255,255,255,0.85);
  backdrop-filter: saturate(140%) blur(12px);
  border-bottom: 1px solid var(--line);
  display: flex;
  align-items: center;
  padding: 0 24px;
  gap: 18px;
  z-index: 100;
}
.brand {
  display: flex; align-items: center; gap: 9px;
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 700; font-size: 19px;
  letter-spacing: -0.025em;
  color: var(--ink);
  flex-shrink: 0;
}
.nav-sep { width: 1px; height: 22px; background: var(--line); flex-shrink: 0; }

.step-trail {
  display: flex; align-items: center; gap: 2px;
  flex: 1; justify-content: center;
}
.step-item {
  display: flex; align-items: center; gap: 7px;
  padding: 6px 12px; border-radius: 10px;
  font-size: 12.5px; font-weight: 600;
  color: var(--muted-2);
  transition: all 200ms var(--ease);
  white-space: nowrap;
  cursor: default;
}
.step-item.active {
  color: var(--accent);
  background: var(--accent-soft);
}
.step-item.done { color: var(--pop-green); }
.step-dot {
  width: 20px; height: 20px; border-radius: 50%;
  background: var(--line);
  color: var(--muted-2);
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 10px; font-weight: 700;
  font-family: 'JetBrains Mono', monospace;
  transition: all 200ms var(--ease);
  flex-shrink: 0;
}
.step-item.active .step-dot { background: var(--accent); color: #fff; }
.step-item.done   .step-dot { background: var(--pop-green); color: #fff; }
.step-arrow { color: var(--muted-2); font-size: 12px; margin: 0 2px; }

/* ═══════════════════════════════════════════════════════════════════════
 * SCREEN WRAPPER
 * ═══════════════════════════════════════════════════════════════════════ */
.main { flex: 1; overflow: hidden; position: relative; }
.screen {
  position: absolute; inset: 0;
  overflow-y: auto;
  display: flex; flex-direction: column;
  padding: 40px 36px;
  background: var(--bg);
  opacity: 0; pointer-events: none;
  transform: translateX(18px);
  transition: opacity 240ms var(--ease), transform 240ms var(--ease);
}
.screen.active {
  opacity: 1; pointer-events: auto;
  transform: translateX(0);
}
#screen-chart { padding: 0; overflow: hidden; }

/* ═══════════════════════════════════════════════════════════════════════
 * UPLOAD SCREEN — with animated network BG (Nodely signature)
 * ═══════════════════════════════════════════════════════════════════════ */
#screen-upload { padding: 0; overflow: hidden; }
.upload-stage {
  position: relative; flex: 1;
  display: flex; align-items: center; justify-content: center;
  padding: 40px 24px;
  overflow: hidden;
}
.upload-bg-canvas {
  position: absolute; inset: 0;
  pointer-events: none;
}
.upload-bg-fade {
  position: absolute; inset: 0;
  background: radial-gradient(900px 520px at 50% 30%, #fff 0%, rgba(255,255,255,0.55) 55%, transparent 85%);
  pointer-events: none;
}
.upload-center {
  position: relative; z-index: 2;
  display: flex; flex-direction: column; align-items: center;
  gap: 28px;
  max-width: 580px; width: 100%;
}
.upload-pill {
  display: inline-flex; align-items: center; gap: 9px;
  padding: 6px 14px;
  background: #fff; border: 1px solid var(--line);
  border-radius: 999px;
  font-family: 'JetBrains Mono', monospace;
  font-size: 11.5px; font-weight: 500;
  color: var(--ink-2); letter-spacing: -0.005em;
  box-shadow: 0 1px 0 rgba(11,11,20,0.04), 0 6px 18px -8px rgba(11,11,20,0.08);
}
.upload-pill .pulse {
  position: relative; display: inline-flex;
  width: 8px; height: 8px;
}
.upload-pill .pulse::before {
  content: ''; position: absolute; inset: 0;
  background: var(--accent); border-radius: 50%; opacity: 0.5;
  animation: ping 1.6s var(--ease) infinite;
}
.upload-pill .pulse::after {
  content: ''; position: relative;
  width: 8px; height: 8px;
  background: var(--accent); border-radius: 50%;
}
@keyframes ping {
  0% { transform: scale(1); opacity: 0.5; }
  80%, 100% { transform: scale(2.4); opacity: 0; }
}
.upload-hero { text-align: center; }
.upload-hero h1 {
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 600;
  font-size: clamp(36px, 5.2vw, 56px);
  line-height: 1.05;
  letter-spacing: -0.035em;
  color: var(--ink);
}
.upload-hero h1 em {
  font-style: normal;
  position: relative;
  display: inline-block;
  color: var(--accent);
}
.upload-hero h1 em::after {
  content: '';
  position: absolute;
  left: -3px; right: -3px;
  bottom: 2px;
  height: 9px;
  background: var(--accent);
  opacity: 0.18;
  border-radius: 4px;
  z-index: -1;
}
.upload-hero p {
  margin-top: 16px;
  font-size: 16px;
  line-height: 1.55;
  color: var(--muted);
  max-width: 480px;
  margin-left: auto; margin-right: auto;
}
.upload-zone {
  width: 100%;
  border: 2px dashed var(--line-2);
  border-radius: var(--r-xl);
  padding: 44px 32px;
  text-align: center;
  cursor: pointer;
  background: rgba(255,255,255,0.7);
  backdrop-filter: blur(4px);
  position: relative;
  transition: all 240ms var(--ease);
  box-shadow: var(--shadow-soft);
}
.upload-zone:hover, .upload-zone.drag-over {
  border-color: var(--accent);
  background: var(--accent-soft);
  box-shadow: 0 0 0 6px rgba(79,70,229,0.06), var(--shadow-pop);
  transform: translateY(-2px);
}
.upload-zone input[type="file"] {
  position: absolute; inset: 0; opacity: 0;
  cursor: pointer; width: 100%; height: 100%;
}
.upload-icon {
  width: 56px; height: 56px;
  margin: 0 auto 16px;
  border-radius: 18px;
  background: linear-gradient(150deg, var(--accent-soft), #fff);
  border: 1px solid var(--accent-mid);
  color: var(--accent);
  display: flex; align-items: center; justify-content: center;
  font-size: 24px;
  transition: transform 320ms var(--spring);
}
.upload-zone:hover .upload-icon { transform: rotate(-8deg) scale(1.05); }
.upload-zone h3 {
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 600; font-size: 18px;
  color: var(--ink);
  margin-bottom: 6px;
}
.upload-zone p {
  font-size: 13.5px; color: var(--muted); line-height: 1.5;
}
.upload-zone p span { color: var(--accent); font-weight: 600; }

.info-cards {
  display: grid; grid-template-columns: 1fr 1fr;
  gap: 14px; width: 100%;
}
.info-card {
  background: rgba(255,255,255,0.85);
  backdrop-filter: blur(8px);
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  padding: 16px 18px;
}
.info-card-title {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10.5px; font-weight: 600;
  color: var(--accent);
  text-transform: uppercase; letter-spacing: 0.08em;
  margin-bottom: 10px;
}
.info-card-row {
  font-size: 12.5px; color: var(--ink-2); font-weight: 500;
  padding: 3px 0; display: flex; align-items: center; gap: 8px;
}
.info-card-row::before {
  content: '';
  width: 5px; height: 5px;
  background: var(--accent-mid); border-radius: 50%;
  flex-shrink: 0;
}

/* ═══════════════════════════════════════════════════════════════════════
 * SECTION HEADERS
 * ═══════════════════════════════════════════════════════════════════════ */
.section-header { margin-bottom: 28px; }
.section-eyebrow {
  font-family: 'JetBrains Mono', monospace;
  font-size: 11px; font-weight: 600;
  color: var(--accent);
  text-transform: uppercase; letter-spacing: 0.12em;
  margin-bottom: 8px;
}
.section-title {
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 600;
  font-size: 30px;
  color: var(--ink);
  letter-spacing: -0.025em;
  line-height: 1.1;
  margin-bottom: 6px;
}
.section-sub {
  font-size: 14.5px;
  color: var(--muted);
  line-height: 1.5;
  max-width: 640px;
}

/* ═══════════════════════════════════════════════════════════════════════
 * MAP SCREEN
 * ═══════════════════════════════════════════════════════════════════════ */
.eyebrow-mini {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10.5px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.08em;
  margin-bottom: 10px;
}
.detected-chips {
  display: flex; flex-wrap: wrap; gap: 7px;
  margin-bottom: 28px;
}
.col-chip {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 6px 12px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: 999px;
  font-size: 12.5px; font-weight: 600;
  color: var(--ink-2);
  transition: border-color 200ms ease;
}
.col-chip:hover { border-color: var(--accent-mid); }
.col-chip .chip-sample {
  color: var(--muted-2); font-size: 11px;
  font-family: 'JetBrains Mono', monospace; font-weight: 500;
}

.map-grid {
  display: grid; grid-template-columns: repeat(3, 1fr);
  gap: 16px; max-width: 880px; margin-bottom: 32px;
}
.map-card {
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-md);
  padding: 18px;
  transition: border-color 200ms ease, box-shadow 200ms ease;
}
.map-card:focus-within {
  border-color: var(--accent);
  box-shadow: 0 0 0 4px rgba(79,70,229,0.1);
}
.map-card-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 11px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.06em;
  margin-bottom: 10px;
  display: flex; align-items: center; gap: 8px;
}
.badge-req {
  background: var(--danger-soft);
  color: var(--danger);
  padding: 1px 7px;
  border-radius: 999px;
  font-size: 9px; font-weight: 700;
  font-family: 'JetBrains Mono', monospace;
  letter-spacing: 0.05em;
}
.badge-opt {
  background: var(--line-2);
  color: var(--muted-2);
  padding: 1px 7px;
  border-radius: 999px;
  font-size: 9px; font-weight: 700;
  font-family: 'JetBrains Mono', monospace;
  letter-spacing: 0.05em;
}
.map-select {
  width: 100%;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 9px 12px;
  font-size: 13.5px; font-weight: 600;
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  outline: none;
  cursor: pointer;
  transition: all 180ms var(--ease);
}
.map-select:focus {
  border-color: var(--accent);
  background: #fff;
  box-shadow: 0 0 0 3px rgba(79,70,229,0.08);
}
.map-hint {
  font-size: 11.5px;
  color: var(--muted-2);
  margin-top: 8px;
  line-height: 1.4;
}

.data-preview-table {
  width: 100%; max-width: 880px;
  border-collapse: collapse;
  margin-bottom: 32px;
  font-size: 12.5px;
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  overflow: hidden;
}
.data-preview-table thead { background: var(--bg-soft); }
.data-preview-table th {
  padding: 9px 14px;
  text-align: left;
  font-weight: 600;
  color: var(--muted);
  border-bottom: 1px solid var(--line);
  font-size: 10.5px;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  font-family: 'JetBrains Mono', monospace;
}
.data-preview-table td {
  padding: 9px 14px;
  border-bottom: 1px solid var(--line-2);
  color: var(--ink-2);
  max-width: 180px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.data-preview-table tr:last-child td { border-bottom: none; }

/* ═══════════════════════════════════════════════════════════════════════
 * CARD DESIGN SCREEN
 * ═══════════════════════════════════════════════════════════════════════ */
.card-design-layout {
  display: grid;
  grid-template-columns: 300px 1fr;
  gap: 28px;
  flex: 1;
  min-height: 0;
}
.fields-panel {
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-lg);
  padding: 20px;
  overflow-y: auto;
}
.fields-panel-title {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10.5px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.08em;
  margin-bottom: 12px;
}
.fields-section { margin-bottom: 18px; }
.fields-section-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10.5px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.07em;
  margin-bottom: 10px;
  padding-bottom: 6px;
  border-bottom: 1px solid var(--line);
}
.field-chip {
  display: inline-flex; align-items: center; gap: 7px;
  padding: 7px 12px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  font-size: 12.5px; font-weight: 600;
  color: var(--ink);
  cursor: grab;
  margin: 3px;
  user-select: none;
  transition: all 180ms var(--ease);
}
.field-chip:hover {
  border-color: var(--accent-mid);
  color: var(--accent);
  background: var(--accent-soft);
  transform: translateY(-1px);
}
.field-chip.placed {
  background: var(--accent-soft);
  border-color: var(--accent-mid);
  color: var(--accent);
}
.field-chip.dragging {
  opacity: 0.4;
  transform: scale(0.93);
}
.drag-icon { color: var(--muted-2); font-size: 11px; }

.card-preview-area {
  display: flex; flex-direction: column;
  align-items: flex-start; gap: 14px;
}
.preview-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10.5px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.08em;
}
.preview-card {
  width: 320px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-top: 4px solid var(--accent);
  border-radius: var(--r-md);
  box-shadow: var(--shadow-card);
  overflow: hidden;
  transition: border-top-color 240ms ease;
}
.preview-card-header {
  padding: 8px 12px;
  background: var(--bg-soft);
  border-bottom: 1px solid var(--line);
  display: flex; align-items: center; gap: 6px;
}
.preview-card-body { padding: 14px 16px; }
.preview-card-footer {
  padding: 8px 12px;
  border-top: 1px solid var(--line);
  background: var(--bg-soft);
  display: flex; align-items: center; gap: 6px;
}
.card-zone {
  flex: 1;
  min-height: 28px;
  min-width: 50px;
  border: 1.5px dashed var(--line);
  border-radius: 8px;
  padding: 4px 8px;
  font-size: 10.5px;
  color: var(--muted-2);
  display: flex; align-items: center; justify-content: center;
  background: var(--bg-soft);
  position: relative; cursor: default;
  transition: all 180ms var(--ease);
}
.card-zone .zone-ph { font-style: italic; opacity: 0.7; }
.card-zone.drop-target {
  border-color: var(--accent);
  background: var(--accent-soft);
  box-shadow: 0 0 0 3px rgba(79,70,229,0.1);
}
.card-zone.filled {
  border-style: solid;
  border-color: var(--accent-mid);
  background: var(--accent-soft);
  flex-direction: column;
  gap: 1px;
  align-items: flex-start;
  justify-content: center;
}
.zone-field {
  font-weight: 700; font-size: 10px; color: var(--accent);
}
.zone-val {
  font-size: 9.5px; color: var(--ink-2); font-style: italic;
  font-family: 'JetBrains Mono', monospace;
}
.zone-remove {
  position: absolute; top: 2px; right: 4px;
  font-size: 9px; cursor: pointer; opacity: 0.5; line-height: 1;
}
.zone-remove:hover { opacity: 1; color: var(--danger); }

.ncard-photo {
  object-fit: cover; object-position: center top;
  display: block; flex-shrink: 0;
}
.ncard-photo-fallback {
  font-weight: 700;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0; letter-spacing: -0.02em;
  font-family: 'Space Grotesk', sans-serif;
}

.preview-hint {
  font-size: 12px; color: var(--muted);
  max-width: 320px; line-height: 1.5;
}

/* color palette + shape buttons */
.color-palette {
  display: flex; flex-wrap: wrap; gap: 8px;
  margin-top: 6px;
}
.color-swatch {
  width: 28px; height: 28px;
  border-radius: 8px;
  cursor: pointer;
  border: 2.5px solid transparent;
  transition: transform 180ms var(--spring);
  flex-shrink: 0;
}
.color-swatch:hover { transform: scale(1.12); }
.color-swatch.selected {
  border-color: var(--ink);
  box-shadow: 0 0 0 2px #fff inset;
}
.shape-btn {
  padding: 6px 12px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  font-size: 12px;
  font-weight: 600;
  color: var(--ink-2);
  cursor: pointer;
  user-select: none;
  transition: all 180ms var(--ease);
}
.shape-btn:hover {
  border-color: var(--accent-mid);
  color: var(--accent);
  background: var(--accent-soft);
}
.shape-btn.selected {
  background: var(--ink);
  border-color: var(--ink);
  color: #fff;
}
.vacant-select {
  flex: 1; min-width: 100px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 6px 10px;
  font-size: 12.5px; font-weight: 600;
  color: var(--ink); font-family: 'Inter', sans-serif;
  outline: none; cursor: pointer;
}
.vacant-select:focus { border-color: var(--accent); }

.emp-type-setup { margin-top: 14px; }
.emp-type-row {
  display: flex; align-items: center; gap: 8px;
  margin-bottom: 8px;
}
.emp-type-label {
  font-size: 12px; font-weight: 600;
  color: var(--ink-2); min-width: 80px;
}
.emp-type-color-input {
  width: 36px; height: 30px;
  border-radius: 8px;
  border: 2px solid var(--line);
  cursor: pointer;
  padding: 0; background: none;
}
.emp-type-value-select {
  flex: 1;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 5px 9px;
  font-size: 12px;
  font-weight: 600; color: var(--ink);
  font-family: 'Inter', sans-serif;
  outline: none; cursor: pointer;
}

/* ═══════════════════════════════════════════════════════════════════════
 * FILTER SCREEN
 * ═══════════════════════════════════════════════════════════════════════ */
.filter-setup { max-width: 720px; }
.filter-counter {
  font-family: 'JetBrains Mono', monospace;
  font-size: 11px;
  color: var(--muted);
  font-weight: 500;
  margin-bottom: 14px;
}
.filter-chips {
  display: flex; flex-wrap: wrap; gap: 8px;
  margin-bottom: 28px;
}
.filter-chip {
  padding: 8px 16px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: 999px;
  font-size: 13px; font-weight: 600;
  color: var(--ink-2);
  cursor: pointer; user-select: none;
  transition: all 180ms var(--ease);
}
.filter-chip:hover {
  border-color: var(--accent-mid);
  color: var(--accent);
  background: var(--accent-soft);
}
.filter-chip.selected {
  background: var(--ink);
  border-color: var(--ink);
  color: #fff;
}
.filter-preview-box {
  background: var(--bg-soft);
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  padding: 18px;
  margin-top: 8px;
}
.fpr-row {
  display: flex; align-items: flex-start; gap: 12px;
  padding: 9px 0;
  border-bottom: 1px solid var(--line);
  font-size: 13px;
}
.fpr-row:last-child { border-bottom: none; }
.fpr-col {
  font-weight: 700; color: var(--ink); min-width: 140px;
}
.fpr-vals {
  display: flex; flex-wrap: wrap; gap: 6px; flex: 1;
}
.fv-pill {
  padding: 2px 10px;
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 999px;
  font-size: 11.5px; font-weight: 500;
  color: var(--ink-2);
  font-family: 'JetBrains Mono', monospace;
}
.export-all-pill {
  background: var(--accent);
  color: #fff;
  border-radius: 999px;
  padding: 1px 8px;
  font-family: 'JetBrains Mono', monospace;
  font-size: 9px; font-weight: 700;
  margin-left: 6px;
  letter-spacing: 0.05em;
}

/* ═══════════════════════════════════════════════════════════════════════
 * BUTTONS — Nodely pill style
 * ═══════════════════════════════════════════════════════════════════════ */
.btn {
  padding: 10px 20px;
  border-radius: 999px;
  font-size: 13.5px; font-weight: 600;
  font-family: 'Space Grotesk', sans-serif;
  cursor: pointer; border: none;
  display: inline-flex; align-items: center; gap: 7px;
  line-height: 1; white-space: nowrap;
  transition: transform 220ms var(--spring), box-shadow 220ms var(--ease), background 200ms ease, border-color 200ms ease;
}
.btn-primary {
  background: var(--ink);
  color: #fff;
  box-shadow: 0 1px 0 rgba(11,11,20,0.06), 0 8px 18px -8px rgba(11,11,20,0.4);
}
.btn-primary:hover {
  background: #1a1a2e;
  transform: translateY(-1px);
  box-shadow: 0 1px 0 rgba(11,11,20,0.06), 0 12px 24px -10px rgba(11,11,20,0.5);
}
.btn-ghost {
  background: #fff;
  color: var(--ink);
  border: 1.5px solid var(--line);
}
.btn-ghost:hover {
  border-color: var(--ink-2);
  background: var(--bg-soft);
}
.btn-sm {
  padding: 7px 14px;
  font-size: 12px;
  border-radius: 999px;
}
.btn-export-all {
  background: linear-gradient(135deg, #7c3aed, var(--accent)) !important;
  color: #fff !important;
  border: none !important;
  box-shadow: 0 1px 0 rgba(11,11,20,0.06), 0 8px 18px -8px rgba(124,58,237,0.5) !important;
}
.btn-export-all:hover {
  transform: translateY(-1px);
  box-shadow: 0 1px 0 rgba(11,11,20,0.06), 0 14px 28px -10px rgba(124,58,237,0.6) !important;
}

/* ═══════════════════════════════════════════════════════════════════════
 * CHART SCREEN — toolbar, stats, filter bar, canvas
 * ═══════════════════════════════════════════════════════════════════════ */
.chart-toolbar {
  flex-shrink: 0;
  min-height: 56px;
  background: rgba(255,255,255,0.85);
  backdrop-filter: saturate(140%) blur(12px);
  border-bottom: 1px solid var(--line);
  display: flex; align-items: center;
  padding: 8px 16px;
  gap: 8px;
  flex-wrap: wrap;
  z-index: 20;
}
.tb-sep { width: 1px; height: 22px; background: var(--line); flex-shrink: 0; }

.stats-bar {
  flex-shrink: 0;
  background: var(--bg-soft);
  border-bottom: 1px solid var(--line);
  display: flex; align-items: center;
  padding: 8px 22px;
  gap: 22px;
  font-size: 12px;
}
.stat-item {
  display: flex; align-items: center; gap: 7px;
  color: var(--muted);
  font-weight: 500;
  font-family: 'JetBrains Mono', monospace;
}
.stat-item strong {
  color: var(--ink);
  font-weight: 700;
  font-family: 'Space Grotesk', sans-serif;
}
.stat-dot {
  width: 7px; height: 7px;
  border-radius: 50%;
  background: var(--accent);
  box-shadow: 0 0 0 3px rgba(79,70,229,0.15);
}

.filter-bar {
  flex-shrink: 0;
  background: #fff;
  border-bottom: 1px solid var(--line);
  padding: 9px 22px;
  display: flex; align-items: center;
  gap: 12px; flex-wrap: wrap;
  min-height: 48px;
}
.filter-dropdown-wrap {
  display: flex; align-items: center; gap: 7px;
  font-size: 12.5px;
}
.filter-dropdown-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10.5px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.06em;
}
.filter-dropdown {
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 6px 12px;
  font-size: 12.5px; font-weight: 600;
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  cursor: pointer; outline: none;
  transition: border-color 180ms ease;
}
.filter-dropdown:focus { border-color: var(--accent); background: #fff; }

.photo-btn {
  display: flex; align-items: center; gap: 6px;
  padding: 6px 12px;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  font-size: 12px; font-weight: 600;
  color: var(--ink-2);
  cursor: pointer;
  transition: all 180ms var(--ease);
  white-space: nowrap; flex-shrink: 0;
  font-family: 'Inter', sans-serif;
}
.photo-btn:hover {
  border-color: var(--accent-mid);
  color: var(--accent);
  background: var(--accent-soft);
}
.photo-btn.loaded {
  border-color: var(--pop-green);
  color: var(--pop-green);
  background: var(--success-soft);
}
.photo-count {
  background: var(--pop-green);
  color: #fff;
  border-radius: 999px;
  padding: 1px 7px;
  font-size: 10px;
  font-weight: 700;
  font-family: 'JetBrains Mono', monospace;
}

.depth-wrap {
  display: flex; align-items: center; gap: 6px;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 4px 8px 4px 11px;
  flex-shrink: 0;
}
.depth-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.07em;
  white-space: nowrap;
}
.depth-select {
  background: transparent;
  border: none;
  padding: 4px 8px 4px 4px;
  font-size: 12.5px; font-weight: 700;
  color: var(--accent);
  font-family: 'Inter', sans-serif;
  cursor: pointer; outline: none;
}

.mgr-mode-btn {
  display: flex; align-items: center; gap: 7px;
  padding: 6px 13px;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  font-size: 12px; font-weight: 600;
  color: var(--ink-2);
  cursor: pointer;
  transition: all 180ms var(--ease);
  white-space: nowrap; flex-shrink: 0;
  user-select: none;
}
.mgr-mode-btn:hover {
  border-color: #c4b5fd;
  color: #7c3aed;
  background: #f5f3ff;
}
.mgr-mode-btn.active {
  background: #f5f3ff;
  border-color: #7c3aed;
  color: #7c3aed;
  box-shadow: 0 0 0 2px #ddd6fe;
}
.mgr-mode-dot {
  width: 8px; height: 8px;
  border-radius: 50%;
  background: var(--line);
  transition: background 180ms ease;
}
.mgr-mode-btn.active .mgr-mode-dot { background: #7c3aed; }

.summary-fields-wrap {
  display: flex; align-items: center; gap: 6px;
  background: #fdf4ff;
  border: 1.5px solid #e9d5ff;
  border-radius: var(--r-sm);
  padding: 4px 8px 4px 11px;
  flex-shrink: 0;
}
.summary-fields-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10px; font-weight: 700;
  color: #7c3aed;
  text-transform: uppercase; letter-spacing: 0.07em;
  white-space: nowrap;
}
.summary-field-select {
  background: transparent; border: none;
  padding: 4px 6px 4px 4px;
  font-size: 12px; font-weight: 600;
  color: #7c3aed;
  font-family: 'Inter', sans-serif;
  cursor: pointer; outline: none;
  max-width: 110px;
}

.summary-list-card {
  display: inline-block;
  width: 240px;
  background: #ffffff;
  border: 1.5px solid var(--line);
  border-top: 3px solid #7c3aed;
  border-radius: var(--r-md);
  box-shadow: var(--shadow-soft);
  font-family: 'Inter', sans-serif;
  overflow: hidden;
  vertical-align: top;
  text-align: left;
}

/* ═══════════════════════════════════════════════════════════════════════
 * CHART CANVAS + ORG TREE
 * ═══════════════════════════════════════════════════════════════════════ */
.chart-canvas-wrap {
  flex: 1; overflow: auto;
  background: var(--bg-soft);
  cursor: grab;
  position: relative;
}
.chart-canvas-wrap:active { cursor: grabbing; }
.chart-canvas-wrap::before {
  content: '';
  position: absolute; inset: 0;
  background-image: radial-gradient(circle, #d4d4dc 1px, transparent 1px);
  background-size: 24px 24px;
  opacity: 0.35;
  pointer-events: none;
}
.chart-canvas-content {
  display: inline-block;
  padding: 60px 80px 120px 80px;
  transform-origin: top left;
  position: relative;
  z-index: 1;
}
.org-tree { display: inline-block; }
.org-tree ul {
  padding-top: 28px;
  position: relative;
  list-style: none;
  display: flex;
  justify-content: center;
}
.org-tree li {
  display: table-cell;
  vertical-align: top;
  text-align: center;
  position: relative;
  padding: 28px 8px 0 8px;
}
.org-tree li::before, .org-tree li::after {
  content: '';
  position: absolute; top: 0; right: 50%;
  border-top: 2px solid #d4d4dc;
  width: 50%; height: 28px;
  transition: border-color 220ms ease;
}
.org-tree li::after {
  right: auto; left: 50%;
  border-left: 2px solid #d4d4dc;
}
.org-tree li:only-child::before, .org-tree li:only-child::after { display: none; }
.org-tree li:first-child::before, .org-tree li:last-child::after { display: none; }
.org-tree li:first-child::after { border-radius: 8px 0 0 0; }
.org-tree li:last-child::before { border-radius: 0 8px 0 0; }
.org-tree ul ul::before {
  content: '';
  position: absolute; top: 0; left: 50%;
  border-left: 2px solid #d4d4dc;
  height: 28px;
}
.org-tree li.collapsed > ul { display: none !important; }

.node-card {
  display: inline-block;
  width: 270px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-top: 3px solid var(--accent);
  border-radius: var(--r-md);
  cursor: pointer;
  text-align: left;
  box-shadow: var(--shadow-soft);
  position: relative;
  font-family: 'Inter', sans-serif;
  transition: transform 320ms var(--spring), box-shadow 240ms var(--ease), border-color 200ms ease;
}
.node-card:hover {
  transform: translateY(-3px);
  box-shadow: 0 22px 50px -22px rgba(79,70,229,0.4), 0 0 0 3px rgba(79,70,229,0.08);
  border-color: var(--accent);
  z-index: 10;
}
.node-card.highlighted {
  border-color: var(--warn) !important;
  border-top-color: var(--warn) !important;
  box-shadow: 0 0 0 3px rgba(217,119,6,0.2), 0 8px 24px rgba(0,0,0,0.1) !important;
}
.node-card.collapsed-node { opacity: 0.65; }

.ncard-header {
  padding: 7px 12px;
  background: var(--bg-soft);
  border-bottom: 1px solid var(--line);
  border-radius: 11px 11px 0 0;
  display: flex; align-items: center; gap: 5px;
}
.ncard-footer {
  padding: 7px 12px;
  border-top: 1px solid var(--line);
  border-radius: 0 0 11px 11px;
  background: var(--bg-soft);
  display: flex; align-items: center; gap: 5px;
}
.ncard-slot {
  font-size: 10px;
  font-weight: 600;
  color: var(--muted-2);
  font-family: 'JetBrains Mono', monospace;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  flex: 1; text-align: center;
}
.ncard-slot.has-val { color: var(--accent); }
.ncard-body { padding: 14px 14px 12px; }
.ncard-body-inner { display: flex; gap: 10px; }
.ncard-body-b1 {
  border-top: 1px dashed var(--line);
  margin-top: 6px;
  padding-top: 6px;
  text-align: center;
  font-size: 11.5px;
  color: var(--accent);
  font-weight: 600;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  font-family: 'JetBrains Mono', monospace;
}
.ncard-text-wrap { width: 100%; text-align: center; }
.ncard-name {
  font-family: 'Space Grotesk', sans-serif;
  font-size: 14.5px;
  font-weight: 600;
  color: var(--ink);
  margin-bottom: 4px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  letter-spacing: -0.01em;
}
.ncard-sub {
  font-size: 12px;
  color: var(--muted);
  line-height: 1.4;
  font-weight: 500;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.collapse-btn {
  position: absolute; bottom: -12px; left: 50%;
  transform: translateX(-50%);
  width: 24px; height: 24px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  cursor: pointer;
  font-size: 9px;
  color: var(--muted-2);
  transition: all 180ms var(--ease);
  z-index: 5;
  box-shadow: var(--shadow-soft);
}
.collapse-btn:hover {
  background: var(--accent);
  border-color: var(--accent);
  color: #fff;
  transform: translateX(-50%) scale(1.1);
}

/* ═══════════════════════════════════════════════════════════════════════
 * SEARCH
 * ═══════════════════════════════════════════════════════════════════════ */
.search-wrap {
  position: relative;
  flex: 1;
  max-width: 260px;
}
.search-icon {
  position: absolute; left: 11px; top: 50%;
  transform: translateY(-50%);
  font-size: 13px;
  pointer-events: none;
  opacity: 0.5;
}
#chart-search {
  width: 100%;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 7px 12px 7px 32px;
  font-size: 13px;
  font-weight: 500;
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  outline: none;
  transition: border-color 180ms ease;
}
#chart-search:focus {
  border-color: var(--accent);
  background: #fff;
  box-shadow: 0 0 0 3px rgba(79,70,229,0.08);
}
#chart-search::placeholder { color: var(--muted-2); }

#chart-search-results {
  position: fixed;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-md);
  box-shadow: 0 12px 40px rgba(0,0,0,0.1), 0 4px 12px rgba(0,0,0,0.05);
  max-height: 320px;
  overflow-y: auto;
  z-index: 99999;
  display: none;
  min-width: 280px;
}
#chart-search-results.visible { display: block; }
.sr-item {
  padding: 10px 14px;
  cursor: pointer;
  border-bottom: 1px solid var(--line-2);
  transition: background 120ms ease;
}
.sr-item:last-child { border-bottom: none; }
.sr-item:hover { background: var(--bg-soft); }
.sr-name {
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 600;
  font-size: 13px;
  color: var(--ink);
}
.sr-sub {
  font-size: 11px;
  color: var(--muted-2);
  margin-top: 2px;
  font-family: 'JetBrains Mono', monospace;
}

/* ═══════════════════════════════════════════════════════════════════════
 * ZOOM CONTROLS
 * ═══════════════════════════════════════════════════════════════════════ */
.zoom-strip {
  display: flex; align-items: center; gap: 1px;
  background: var(--bg-soft);
  border-radius: var(--r-sm);
  padding: 2px;
  border: 1.5px solid var(--line);
}
.btn-zoom {
  background: transparent;
  border: none;
  border-radius: 7px;
  width: 26px; height: 26px;
  cursor: pointer;
  font-size: 14px; font-weight: 700;
  color: var(--ink-2);
  display: flex; align-items: center; justify-content: center;
  transition: background 120ms ease;
}
.btn-zoom:hover { background: #fff; color: var(--accent); }
.zoom-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 11px;
  font-weight: 600;
  color: var(--ink);
  min-width: 40px;
  text-align: center;
  font-variant-numeric: tabular-nums;
}

/* ═══════════════════════════════════════════════════════════════════════
 * BG control
 * ═══════════════════════════════════════════════════════════════════════ */
.bg-control-wrap {
  display: flex; align-items: center; gap: 6px;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 4px 8px;
  flex-shrink: 0;
}
.bg-control-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 10px; font-weight: 600;
  color: var(--muted);
  text-transform: uppercase; letter-spacing: 0.07em;
  white-space: nowrap;
}
.bg-color-input {
  width: 26px; height: 24px;
  border-radius: 6px;
  border: 1.5px solid var(--line);
  cursor: pointer;
  padding: 0; background: none;
  flex-shrink: 0;
}
.bg-color-input:disabled { opacity: 0.4; cursor: not-allowed; }
.bg-transparent-btn {
  display: inline-flex; align-items: center; gap: 4px;
  padding: 3px 9px;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: 7px;
  font-size: 11px; font-weight: 600;
  color: var(--ink-2);
  cursor: pointer; user-select: none;
  font-family: 'Inter', sans-serif;
  transition: all 180ms var(--ease);
  line-height: 1.3;
}
.bg-transparent-btn:hover {
  border-color: var(--accent-mid);
  color: var(--accent);
  background: var(--accent-soft);
}
.bg-transparent-btn.active {
  background: var(--warn-soft);
  border-color: var(--warn);
  color: #92400e;
  box-shadow: 0 0 0 2px #fde68a;
}
.chart-canvas-wrap.transparent-preview::before { display: none; }
.chart-canvas-wrap.transparent-preview {
  background-color: #ffffff !important;
  background-image:
    linear-gradient(45deg, #e2e8f0 25%, transparent 25%),
    linear-gradient(-45deg, #e2e8f0 25%, transparent 25%),
    linear-gradient(45deg, transparent 75%, #e2e8f0 75%),
    linear-gradient(-45deg, transparent 75%, #e2e8f0 75%) !important;
  background-size: 20px 20px !important;
  background-position: 0 0, 0 10px, 10px -10px, -10px 0 !important;
}

/* ═══════════════════════════════════════════════════════════════════════
 * EXPORT OVERLAY
 * ═══════════════════════════════════════════════════════════════════════ */
.export-overlay {
  position: fixed; inset: 0;
  z-index: 9999;
  background: rgba(255,255,255,0.94);
  backdrop-filter: blur(8px);
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  gap: 12px;
}
.export-spinner {
  width: 48px; height: 48px;
  border: 3px solid var(--line);
  border-top-color: var(--accent);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }

.no-data {
  padding: 44px;
  color: var(--muted);
  font-size: 14px;
  font-weight: 500;
  background: #fff;
  border: 1.5px solid var(--line);
  border-radius: var(--r-md);
  max-width: 460px;
  font-family: 'Inter', sans-serif;
}

/* ═══════════════════════════════════════════════════════════════════════
 * MODAL
 * ═══════════════════════════════════════════════════════════════════════ */
.modal-overlay {
  position: fixed; inset: 0;
  z-index: 8000;
  background: rgba(11,11,20,0.5);
  backdrop-filter: blur(6px);
  display: flex; align-items: center; justify-content: center;
  padding: 20px;
}
.modal-overlay.hidden { display: none; }
.modal-box {
  background: #fff;
  border: 1px solid var(--line);
  border-radius: var(--r-xl);
  box-shadow: 0 24px 80px rgba(0,0,0,0.25);
  width: 460px;
  max-width: 100%;
  display: flex; flex-direction: column;
  max-height: 80vh;
  overflow: hidden;
}
.modal-header {
  padding: 20px 22px 16px;
  border-bottom: 1px solid var(--line);
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
}
.modal-title {
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 600;
  font-size: 17px;
  color: var(--ink);
  letter-spacing: -0.02em;
}
.modal-sub {
  font-size: 12.5px;
  color: var(--muted);
  margin-top: 4px;
}
.modal-close {
  background: none; border: none;
  font-size: 18px;
  cursor: pointer;
  color: var(--muted-2);
  line-height: 1;
  padding: 4px 8px;
  border-radius: 8px;
  transition: all 180ms ease;
}
.modal-close:hover { background: var(--bg-soft); color: var(--ink); }
.modal-body {
  padding: 18px 22px;
  flex: 1; overflow-y: auto;
  display: flex; flex-direction: column;
  gap: 14px;
}
.modal-search {
  width: 100%;
  background: var(--bg-soft);
  border: 1.5px solid var(--line);
  border-radius: var(--r-sm);
  padding: 9px 14px;
  font-size: 13.5px;
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  outline: none;
}
.modal-search:focus {
  border-color: var(--accent);
  background: #fff;
}
.modal-list {
  display: flex; flex-direction: column; gap: 1px;
  max-height: 300px; overflow-y: auto;
}
.modal-emp-row {
  display: flex; align-items: center; gap: 11px;
  padding: 10px 12px;
  border-radius: var(--r-sm);
  cursor: pointer;
  transition: background 120ms ease;
  border: 2px solid transparent;
}
.modal-emp-row:hover { background: var(--bg-soft); }
.modal-emp-row.selected {
  background: var(--accent-soft);
  border-color: var(--accent-mid);
}
.modal-emp-avatar {
  width: 34px; height: 34px;
  border-radius: 10px;
  background: var(--accent-soft);
  color: var(--accent);
  font-family: 'Space Grotesk', sans-serif;
  font-size: 12px;
  font-weight: 700;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.modal-emp-name {
  font-family: 'Space Grotesk', sans-serif;
  font-weight: 600;
  font-size: 13.5px;
  color: var(--ink);
}
.modal-emp-sub {
  font-size: 11px;
  color: var(--muted-2);
  font-family: 'JetBrains Mono', monospace;
}
.modal-footer {
  padding: 16px 22px;
  border-top: 1px solid var(--line);
  display: flex; gap: 10px;
  justify-content: flex-end;
  align-items: center;
}
.modal-note {
  font-size: 12px;
  color: var(--muted);
  flex: 1;
  font-family: 'JetBrains Mono', monospace;
}

.btn-danger {
  background: var(--danger-soft) !important;
  border: 1.5px solid #fca5a5 !important;
  color: var(--danger) !important;
}
.btn-danger:hover {
  background: #fecaca !important;
}

.photo-folder-input { display: none; }

.preview-b1-zone {
  margin: 4px 0 0 0;
  padding-top: 6px;
  border-top: 1px dashed var(--line);
}

/* exported chart shouldn't clip */
.export-stage-root .org-tree li,
.export-stage-root .org-tree ul,
.export-stage-root .node-card,
.export-stage-root .summary-list-card { overflow: visible !important; }

</style>
</head>
<body>

<!-- ═══════════════════════════════════════════════════════════════════════
     TOP NAV
     ═══════════════════════════════════════════════════════════════════════ -->
<nav class="topnav">
  <div class="brand">
    <svg width="26" height="26" viewBox="0 0 26 26">
      <circle cx="6" cy="13" r="3.2" fill="#ff6b9d"/>
      <circle cx="20" cy="6" r="3.2" fill="#ffd166"/>
      <circle cx="20" cy="20" r="3.2" fill="#10b981"/>
      <circle cx="13" cy="13" r="3.6" fill="#4f46e5"/>
      <line x1="13" y1="13" x2="6" y2="13" stroke="#0b0b14" stroke-width="1.4" stroke-linecap="round"/>
      <line x1="13" y1="13" x2="20" y2="6" stroke="#0b0b14" stroke-width="1.4" stroke-linecap="round"/>
      <line x1="13" y1="13" x2="20" y2="20" stroke="#0b0b14" stroke-width="1.4" stroke-linecap="round"/>
    </svg>
    Nodely
  </div>
  <div class="nav-sep"></div>
  <div class="step-trail">
    <div class="step-item active" id="nav-step-upload"><div class="step-dot">1</div><span>Upload</span></div>
    <div class="step-arrow">·</div>
    <div class="step-item" id="nav-step-map"><div class="step-dot">2</div><span>Map columns</span></div>
    <div class="step-arrow">·</div>
    <div class="step-item" id="nav-step-card"><div class="step-dot">3</div><span>Design card</span></div>
    <div class="step-arrow">·</div>
    <div class="step-item" id="nav-step-filter"><div class="step-dot">4</div><span>Set filters</span></div>
    <div class="step-arrow">·</div>
    <div class="step-item" id="nav-step-chart"><div class="step-dot">5</div><span>Org chart</span></div>
  </div>
</nav>

<main class="main">

<!-- ═══════════════════════════════════════════════════════════════════════
     UPLOAD SCREEN
     ═══════════════════════════════════════════════════════════════════════ -->
<div class="screen active" id="screen-upload">
  <div class="upload-stage">
    <canvas class="upload-bg-canvas" id="upload-network-canvas"></canvas>
    <div class="upload-bg-fade"></div>
    <div class="upload-center">
      <div class="upload-pill">
        <span class="pulse"></span>
        Now connecting 200,000+ people in live org charts
      </div>
      <div class="upload-hero">
        <h1>Org charts that<br/><em>actually breathe.</em></h1>
        <p>Drop in your roster. Nodely turns it into a beautiful, interactive map of how your team connects — drag, filter, share, done.</p>
      </div>
      <div class="upload-zone" id="upload-dropzone">
        <input type="file" id="file-input" accept=".csv,.xlsx,.xls"/>
        <div class="upload-icon">↑</div>
        <h3>Drop your CSV or Excel</h3>
        <p>or <span>click to browse</span> · supports .csv, .xlsx, .xls</p>
      </div>
      <div class="info-cards">
        <div class="info-card">
          <div class="info-card-title">Required columns</div>
          <div class="info-card-row">Employee Code / ID</div>
          <div class="info-card-row">Employee Name</div>
          <div class="info-card-row">Manager Code / ID</div>
        </div>
        <div class="info-card">
          <div class="info-card-title">Photo tip</div>
          <div class="info-card-row">Name photos by Employee ID</div>
          <div class="info-card-row">e.g. EMP001.jpg, E001.png</div>
          <div class="info-card-row">Load folder from chart toolbar</div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════════════════
     MAP SCREEN
     ═══════════════════════════════════════════════════════════════════════ -->
<div class="screen" id="screen-map">
  <div class="section-header">
    <div class="section-eyebrow">Step 02 · Map columns</div>
    <div class="section-title">Tell us who's who.</div>
    <div class="section-sub">We detected <span id="col-count">0</span> columns and auto-mapped what we recognized. Tweak below if needed.</div>
  </div>
  <div class="eyebrow-mini">Detected columns</div>
  <div class="detected-chips" id="detected-columns"></div>
  <div class="map-grid">
    <div class="map-card">
      <div class="map-card-label">Employee ID <span class="badge-req">REQ</span></div>
      <select class="map-select" id="map-empId"></select>
      <div class="map-hint">Unique identifier — also used to match photos by filename.</div>
    </div>
    <div class="map-card">
      <div class="map-card-label">Employee name <span class="badge-req">REQ</span></div>
      <select class="map-select" id="map-empName"></select>
      <div class="map-hint">Full name shown on every card.</div>
    </div>
    <div class="map-card">
      <div class="map-card-label">Manager ID <span class="badge-opt">OPT</span></div>
      <select class="map-select" id="map-managerId"></select>
      <div class="map-hint">Links each person to their manager.</div>
    </div>
  </div>
  <div class="eyebrow-mini">Preview · first 3 rows</div>
  <div id="data-preview-wrap" style="margin-bottom:32px; overflow-x:auto"></div>
  <div style="display:flex; gap:12px;">
    <button class="btn btn-ghost" onclick="goTo('upload')">Back</button>
    <button class="btn btn-primary" onclick="confirmColumnMap()">Continue <span>→</span></button>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════════════════
     CARD DESIGN SCREEN
     ═══════════════════════════════════════════════════════════════════════ -->
<div class="screen" id="screen-card">
  <div class="section-header" style="margin-bottom:20px">
    <div class="section-eyebrow">Step 03 · Design card</div>
    <div class="section-title">Make it yours.</div>
    <div class="section-sub">Drag fields onto card slots. Set an accent. Configure colors for active, vacant, contractor.</div>
  </div>
  <div class="card-design-layout">
    <div class="fields-panel">
      <div class="fields-panel-title">Available fields</div>
      <div id="card-fields-panel"></div>

      <div class="fields-section" style="margin-top:18px">
        <div class="fields-section-label">Card accent</div>
        <div class="color-palette" id="color-palette"></div>
      </div>

      <div class="fields-section" style="margin-top:14px">
        <div class="fields-section-label">Photo size &amp; shape</div>
        <div style="font-size:11.5px; color:var(--muted); margin-bottom:7px">Size</div>
        <div style="display:flex; align-items:center; gap:9px; margin-bottom:12px">
          <input type="range" id="photo-size-slider" min="40" max="160" step="10" value="80"
                 style="flex:1; accent-color:var(--accent); cursor:pointer"
                 oninput="S.photoSize=parseInt(this.value);document.getElementById('photo-size-val').textContent=this.value+'px';renderCardPreview();renderChart()"/>
          <span id="photo-size-val" class="mono" style="font-size:11.5px; font-weight:600; color:var(--accent); min-width:40px">80px</span>
        </div>
        <div style="font-size:11.5px; color:var(--muted); margin-bottom:7px">Shape</div>
        <div style="display:flex; gap:6px; flex-wrap:wrap; margin-bottom:12px">
          <div class="shape-btn selected" data-shape="circle"  onclick="setPhotoShape('circle')">circle</div>
          <div class="shape-btn"          data-shape="rounded" onclick="setPhotoShape('rounded')">rounded</div>
          <div class="shape-btn"          data-shape="square"  onclick="setPhotoShape('square')">square</div>
        </div>
        <div style="font-size:11.5px; color:var(--muted); margin-bottom:7px">Placement</div>
        <div style="display:flex; gap:6px; flex-wrap:wrap">
          <div class="shape-btn selected" data-placement="top"   onclick="setPhotoPlacement('top')">top</div>
          <div class="shape-btn"          data-placement="left"  onclick="setPhotoPlacement('left')">left</div>
          <div class="shape-btn"          data-placement="right" onclick="setPhotoPlacement('right')">right</div>
          <div class="shape-btn"          data-placement="none"  onclick="setPhotoPlacement('none')">none</div>
        </div>
      </div>

      <div class="fields-section emp-type-setup">
        <div class="fields-section-label">Employment type colors</div>
        <div style="font-size:11.5px; color:var(--muted); margin-bottom:10px; line-height:1.5">
          Map a column &amp; values to control the card border color.
        </div>
        <select class="vacant-select" id="emp-type-col" onchange="onEmpTypeColChange()" style="width:100%; margin-bottom:10px">
          <option value="">Select column...</option>
        </select>
        <div id="emp-type-rows" style="display:none">
          <div class="emp-type-row">
            <div class="emp-type-label">Active</div>
            <select class="emp-type-value-select" id="emp-val-active"><option value="">Value...</option></select>
            <input type="color" class="emp-type-color-input" id="emp-color-active" value="#10b981"/>
          </div>
          <div class="emp-type-row">
            <div class="emp-type-label">Vacant</div>
            <select class="emp-type-value-select" id="emp-val-vacant"><option value="">Value...</option></select>
            <input type="color" class="emp-type-color-input" id="emp-color-vacant" value="#dc2626"/>
          </div>
          <div class="emp-type-row">
            <div class="emp-type-label">Resigned</div>
            <select class="emp-type-value-select" id="emp-val-resigned"><option value="">Value...</option></select>
            <input type="color" class="emp-type-color-input" id="emp-color-resigned" value="#d97706"/>
          </div>
        </div>
      </div>
    </div>

    <div class="card-preview-area">
      <div class="preview-label">Live preview</div>
      <div id="card-preview"></div>
      <div class="preview-hint">Drag a field chip onto a header, body, or footer slot. Name is always fixed in the card body.</div>
    </div>
  </div>
  <div style="display:flex; gap:12px; margin-top:20px;">
    <button class="btn btn-ghost" onclick="goTo('map')">Back</button>
    <button class="btn btn-primary" onclick="confirmCardDesign()">Continue <span>→</span></button>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════════════════
     FILTER SCREEN
     ═══════════════════════════════════════════════════════════════════════ -->
<div class="screen" id="screen-filter">
  <div class="section-header">
    <div class="section-eyebrow">Step 04 · Filters</div>
    <div class="section-title">Slice it however you want.</div>
    <div class="section-sub">Pick up to 3 columns to filter by. The last one drives the "Export all" button — perfect for one-click team bundles.</div>
  </div>
  <div class="filter-setup">
    <div class="filter-counter" id="filter-counter">0 of 3 filters selected</div>
    <div class="filter-chips" id="filter-chip-picker"></div>
    <div id="filter-preview-area"></div>
  </div>
  <div style="display:flex; gap:12px; margin-top:24px;">
    <button class="btn btn-ghost" onclick="goTo('card')">Back</button>
    <button class="btn btn-primary" onclick="launchChart()">See the chart <span>→</span></button>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════════════════
     CHART SCREEN
     ═══════════════════════════════════════════════════════════════════════ -->
<div class="screen" id="screen-chart">
  <div class="chart-toolbar">
    <button class="btn btn-ghost btn-sm" onclick="goTo('filter')">← Setup</button>
    <div class="tb-sep"></div>
    <div class="search-wrap">
      <span class="search-icon">🔍</span>
      <input id="chart-search" type="text" placeholder="Search name or ID..." autocomplete="off"/>
      <div id="chart-search-results"></div>
    </div>
    <div class="tb-sep"></div>
    <div class="zoom-strip">
      <button class="btn-zoom" onclick="zoomBy(-0.1)">−</button>
      <span class="zoom-label" id="zoom-level">100%</span>
      <button class="btn-zoom" onclick="zoomBy(0.1)">+</button>
      <button class="btn-zoom" onclick="fitToScreen(true)" title="Fit">⊡</button>
    </div>
    <button class="btn btn-ghost btn-sm" onclick="centerView()">Center</button>
    <button class="btn btn-ghost btn-sm" onclick="expandAll()">Expand</button>
    <button class="btn btn-ghost btn-sm" onclick="collapseAll()">Collapse</button>
    <div class="tb-sep"></div>
    <div class="depth-wrap">
      <span class="depth-label">Skip top</span>
      <select class="depth-select" id="depth-select" onchange="setSkipDepth(parseInt(this.value))">
        <option value="0">none</option>
        <option value="1">L1</option>
        <option value="2">L2</option>
        <option value="3">L3</option>
        <option value="4">L4</option>
        <option value="5">L5</option>
        <option value="6">L6</option>
      </select>
    </div>
    <div class="tb-sep"></div>
    <div class="mgr-mode-btn" id="mgr-mode-btn" onclick="toggleManagerMode()">
      <div class="mgr-mode-dot"></div>Manager view
    </div>
    <div class="summary-fields-wrap" id="summary-fields-wrap" style="display:none">
      <span class="summary-fields-label">Show</span>
      <select class="summary-field-select" id="summary-field1" onchange="S.summaryField1=this.value;if(S.managerMode)renderChart()">
        <option value="">Field 1...</option>
      </select>
      <span style="font-size:10px; color:#7c3aed; font-weight:700">+</span>
      <select class="summary-field-select" id="summary-field2" onchange="S.summaryField2=this.value;if(S.managerMode)renderChart()">
        <option value="">Field 2...</option>
      </select>
    </div>
    <div class="tb-sep"></div>
    <input type="file" id="photo-folder-input" class="photo-folder-input" accept="image/*" multiple webkitdirectory/>
    <div class="photo-btn" id="photo-btn" onclick="openPhotoFolder()">
      📸 <span id="photo-btn-label">Load photos</span>
      <span class="photo-count" id="photo-count" style="display:none">0</span>
    </div>
    <div class="tb-sep"></div>
    <div style="flex:1"></div>
    <div class="bg-control-wrap" title="Chart background">
      <span class="bg-control-label">BG</span>
      <input type="color" class="bg-color-input" id="bg-color-input" value="#fafaf9" oninput="setChartBg(this.value)" title="Chart background color"/>
      <button class="bg-transparent-btn" id="bg-transparent-btn" onclick="toggleTransparent()" title="Transparent — drop straight into PowerPoint">⊘ none</button>
    </div>
    <div class="tb-sep"></div>
    <button class="btn btn-ghost btn-sm" onclick="downloadCSV()">CSV</button>
    <button class="btn btn-ghost btn-sm" onclick="exportPNG()">PNG</button>
    <button class="btn btn-ghost btn-sm" onclick="exportPPTX()">PPTX</button>
    <button class="btn btn-sm btn-export-all" onclick="exportAll()">Export all</button>
  </div>

  <div class="stats-bar">
    <div class="stat-item"><div class="stat-dot"></div><strong id="stat-total">—</strong>&nbsp;employees</div>
    <div class="stat-item"><strong id="stat-roots">—</strong>&nbsp;roots</div>
    <div class="stat-item"><strong id="stat-vis">—</strong>&nbsp;visible</div>
    <div class="stat-item" id="stat-photos" style="display:none; color:var(--pop-green)">📸&nbsp;<strong id="stat-photos-val">0</strong>&nbsp;photos</div>
    <div class="stat-item" id="stat-mgr-mode" style="display:none; color:#7c3aed">👔&nbsp;<strong id="stat-mgr-val">—</strong></div>
    <div class="stat-item" id="stat-filtered" style="display:none; color:var(--warn)">filtered</div>
  </div>

  <div class="filter-bar" id="filter-bar" style="display:none"></div>

  <div class="chart-canvas-wrap" id="chart-canvas-wrap">
    <div class="chart-canvas-content" id="chart-canvas-content">
      <div class="org-tree" id="org-tree"></div>
    </div>
  </div>
</div>

</main>

<!-- ═══════════════════════════════════════════════════════════════════════
     REASSIGN MODAL
     ═══════════════════════════════════════════════════════════════════════ -->
<div class="modal-overlay hidden" id="reassign-modal">
  <div class="modal-box">
    <div class="modal-header">
      <div>
        <div class="modal-title">Reassign manager</div>
        <div class="modal-sub" id="reassign-subject">Moving —</div>
      </div>
      <button class="modal-close" onclick="closeReassignModal()">✕</button>
    </div>
    <div class="modal-body">
      <input class="modal-search" id="reassign-search" type="text" placeholder="Search employee name or ID..." autocomplete="off" oninput="filterReassignList()"/>
      <div class="modal-list" id="reassign-list"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-sm btn-danger" onclick="removeCurrentNode()" style="margin-right:auto">Remove</button>
      <span class="modal-note" id="reassign-note">Pick a new manager above</span>
      <button class="btn btn-ghost btn-sm" onclick="closeReassignModal()">Cancel</button>
      <button class="btn btn-primary btn-sm" id="reassign-confirm-btn" onclick="confirmReassign()" disabled>Reassign</button>
    </div>
  </div>
</div>

<script>
/* ═══════════════════════════════════════════════════════════════════════
 * NODELY — full app. All features audited from OrgDesign Pro and preserved.
 *
 * FEATURE AUDIT CHECKLIST (each verified working):
 *   ✓ CSV/Excel upload
 *   ✓ Auto column detection
 *   ✓ Column mapping with validation
 *   ✓ Drag/drop card design (h1/h2/h3/b1/f1/f2/f3 zones)
 *   ✓ Photo size slider, shape (circle/rounded/square), placement (top/left/right/none)
 *   ✓ Employment type colors (column + 3 value/color pairs)
 *   ✓ Up to 3 filter columns
 *   ✓ Filter cascade preserves manager chains
 *   ✓ Chart pan, zoom, fit-to-screen, center
 *   ✓ Search w/ floating results + scroll-to-node
 *   ✓ Expand/collapse all + per-node toggle
 *   ✓ Skip top N levels (root reanchor)
 *   ✓ Manager view + IC summary cards
 *   ✓ Photo folder load (DirectoryPicker → webkitdirectory fallback)
 *   ✓ Reassign manager modal + remove node
 *   ✓ Per-node subtree PNG export
 *   ✓ BG color picker + transparent preview/export
 *   ✓ CSV / PNG / PPTX / Export-all-as-ZIP
 * ═══════════════════════════════════════════════════════════════════════ */

const S = {
  rawRows: [], columns: [], colSamples: {},
  colMap: { empId: '', empName: '', managerId: '' },
  cardSlots: { h1:'', h2:'', h3:'', b1:'', f1:'', f2:'', f3:'' },
  cardAccent: '#4f46e5',
  empTypeCol: '', empTypeMap: {},
  empTypeLabels: { active:'', vacant:'', resigned:'' },
  empTypeColors: { active:'#10b981', vacant:'#dc2626', resigned:'#d97706' },
  filterCols: [], activeFilters: {},
  managerOverrides: {}, removedIds: new Set(),
  viewData: [], childMap: {}, descCount: {}, nodeHeight: {}, nodeDepth: {},
  zoom: 1, highlighted: null,
  draggingField: null,
  reassignTarget: null, reassignPick: null,
  skipDepth: 0,
  photoMap: {}, photoObjUrls: [],
  photoSize: 80, photoShape: 'circle', photoPlacement: 'top',
  managerMode: false,
  summaryField1: '', summaryField2: '',
  chartBgColor: '#fafaf9',
  transparentExport: false,
};

function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
function xe(s)  { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;'); }

/* ── Step navigation ─────────────────────────────────────────────────── */
function goTo(step) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + step).classList.add('active');
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
    setTimeout(initPan, 80);
    setTimeout(initSearch, 80);
    setTimeout(populateSummaryFields, 120);
    setTimeout(applyChartBg, 120);
  }
}

/* ── File handling ───────────────────────────────────────────────────── */
function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: r => initData(r.data),
      error:    e => alert('CSV error: ' + e.message),
    });
  } else if (['xlsx','xls'].includes(ext)) {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      initData(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' }));
    };
    reader.readAsArrayBuffer(file);
  } else {
    alert('Please upload a CSV or Excel file.');
  }
}
function initData(rows) {
  S.rawRows = rows;
  S.columns = rows.length ? Object.keys(rows[0]) : [];
  S.colSamples = {};
  S.columns.forEach(col => {
    S.colSamples[col] = [...new Set(
      rows.slice(0, 25).map(r => String(r[col]||'').trim()).filter(v => v && v !== 'undefined' && v !== 'null')
    )].slice(0, 3);
  });
  S.colMap = autoDetect(S.columns);
  buildMapScreen();
  goTo('map');
}

function autoDetect(cols) {
  const lc = cols.map(c => c.toLowerCase().trim());
  function find(exact, partial) {
    for (const p of exact)   { const i = lc.findIndex(c => c === p); if (i >= 0) return cols[i]; }
    for (const p of partial) { const i = lc.findIndex(c => c.startsWith(p) || c.endsWith(p)); if (i >= 0) return cols[i]; }
    for (const p of partial) { const i = lc.findIndex(c => c.includes(p)); if (i >= 0) return cols[i]; }
    return '';
  }
  return {
    empId: find(
      ['employee code','emp code','emp id','employee id','empcode','empid','staff id','employee_id','emp_id'],
      ['employee code','emp code','employee id','emp id','empcode','empid','staff id']
    ),
    empName: find(
      ['employee name','emp name','full name','person name','staff name','employee_name','emp_name','full_name'],
      ['employee name','emp name','full name','person name','staff name']
    ),
    managerId: find(
      ['l1 manager code','l1 manager','manager code','manager id','reports to','supervisor','mgr code','mgrid','manager_code','manager_id'],
      ['manager code','manager id','l1 manager','reports to','supervisor','mgr code']
    ),
  };
}

/* ── Photo loading ───────────────────────────────────────────────────── */
async function openPhotoFolder() {
  if ('showDirectoryPicker' in window) {
    try {
      const d = await window.showDirectoryPicker({ mode: 'read' });
      await loadFromDirectoryHandle(d);
    } catch (e) {
      if (e.name !== 'AbortError') document.getElementById('photo-folder-input').click();
    }
  } else {
    document.getElementById('photo-folder-input').click();
  }
}
async function loadFromDirectoryHandle(dirHandle) {
  S.photoObjUrls.forEach(u => URL.revokeObjectURL(u));
  S.photoObjUrls = [];
  const newMap = {};
  const IMG = new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);
  for await (const [name, handle] of dirHandle.entries()) {
    if (handle.kind === 'file') {
      const ext = name.split('.').pop().toLowerCase();
      if (IMG.has(ext)) {
        const f = await handle.getFile();
        const k = name.replace(/\.[^.]+$/,'').toLowerCase().trim();
        const u = URL.createObjectURL(f);
        newMap[k] = u;
        S.photoObjUrls.push(u);
      }
    }
  }
  S.photoMap = newMap;
  updatePhotoUI();
  if (S.viewData.length) renderChart();
}
function loadFromFileInput(files) {
  S.photoObjUrls.forEach(u => URL.revokeObjectURL(u));
  S.photoObjUrls = [];
  const newMap = {};
  const IMG = new Set(['jpg','jpeg','png','gif','webp','bmp','avif']);
  Array.from(files).forEach(file => {
    const ext = file.name.split('.').pop().toLowerCase();
    if (IMG.has(ext)) {
      const k = file.name.replace(/\.[^.]+$/,'').toLowerCase().trim();
      const u = URL.createObjectURL(file);
      newMap[k] = u;
      S.photoObjUrls.push(u);
    }
  });
  S.photoMap = newMap;
  updatePhotoUI();
  if (S.viewData.length) renderChart();
}
function updatePhotoUI() {
  const count = Object.keys(S.photoMap).length;
  document.getElementById('photo-btn').classList.toggle('loaded', count > 0);
  document.getElementById('photo-btn-label').textContent = count > 0 ? 'Photos' : 'Load photos';
  const badge = document.getElementById('photo-count');
  badge.textContent = count;
  badge.style.display = count > 0 ? '' : 'none';
  const stat = document.getElementById('stat-photos');
  if (stat) {
    stat.style.display = count > 0 ? 'flex' : 'none';
    document.getElementById('stat-photos-val').textContent = count;
  }
}
function getPhotoUrl(node) {
  if (!Object.keys(S.photoMap).length) return '';
  const id = node.id.toLowerCase().trim();
  if (S.photoMap[id]) return S.photoMap[id];
  const nk  = node.name.toLowerCase().trim().replace(/\s+/g, '_');
  if (S.photoMap[nk]) return S.photoMap[nk];
  const nk2 = node.name.toLowerCase().trim().replace(/\s+/g, '');
  if (S.photoMap[nk2]) return S.photoMap[nk2];
  return '';
}

/* ── Map screen ──────────────────────────────────────────────────────── */
function buildMapScreen() {
  document.getElementById('col-count').textContent = S.columns.length;
  document.getElementById('detected-columns').innerHTML = S.columns.map(c =>
    '<div class="col-chip">' + esc(c) +
    (S.colSamples[c].length ? '<span class="chip-sample">' + esc(S.colSamples[c].join(', ')) + '</span>' : '') +
    '</div>'
  ).join('');
  const blank = '<option value="">— select —</option>';
  const opts = blank + S.columns.map(c => '<option value="' + esc(c) + '">' + esc(c) + '</option>').join('');
  ['empId','empName','managerId'].forEach(k => {
    const sel = document.getElementById('map-' + k);
    if (!sel) return;
    sel.innerHTML = opts;
    sel.value = S.colMap[k] || '';
  });
  const wrap = document.getElementById('data-preview-wrap');
  const preview = S.rawRows.slice(0, 3);
  if (!preview.length) { wrap.innerHTML = ''; return; }
  let html = '<table class="data-preview-table"><thead><tr>' +
    S.columns.map(c => '<th>' + esc(c) + '</th>').join('') +
    '</tr></thead><tbody>';
  preview.forEach(row => {
    html += '<tr>' + S.columns.map(c => '<td>' + esc(String(row[c]||'').substring(0,22)) + '</td>').join('') + '</tr>';
  });
  wrap.innerHTML = html + '</tbody></table>';
}
function confirmColumnMap() {
  S.colMap.empId     = document.getElementById('map-empId').value;
  S.colMap.empName   = document.getElementById('map-empName').value;
  S.colMap.managerId = document.getElementById('map-managerId').value;
  if (!S.colMap.empId || !S.colMap.empName) {
    alert('Please map Employee ID and Employee Name.'); return;
  }
  if (S.colMap.empId === S.colMap.empName) {
    alert('Employee ID and Employee Name must be different columns.'); return;
  }
  if (S.colMap.managerId && S.colMap.managerId === S.colMap.empId) {
    alert('Manager ID and Employee ID must be different columns.'); return;
  }
  buildCardScreen();
  goTo('card');
}

/* ── Card design screen ──────────────────────────────────────────────── */
const AUTO_FIELDS = [
  { id: '__auto_reports__',  icon: '📊', label: 'Direct reports', desc: 'Count of direct reports' },
  { id: '__auto_teamsize__', icon: '👥', label: 'Total team',     desc: 'All descendants count' },
];

function buildCardScreen() {
  const core = new Set([S.colMap.empId, S.colMap.empName, S.colMap.managerId].filter(Boolean));
  const available = S.columns.filter(c => !core.has(c));
  document.getElementById('card-fields-panel').innerHTML =
    '<div class="fields-section"><div class="fields-section-label">Column fields</div>' +
    (available.length
      ? available.map(f => '<div class="field-chip" draggable="true" data-field="' + esc(f) +
                           '" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)"><span class="drag-icon">⠿</span>' + esc(f) + '</div>').join('')
      : '<div style="font-size:12.5px; color:var(--muted-2); font-style:italic">No extra columns</div>'
    ) +
    '</div>' +
    '<div class="fields-section"><div class="fields-section-label">Auto-calculated</div>' +
    AUTO_FIELDS.map(f => '<div class="field-chip" draggable="true" data-field="' + f.id +
                         '" ondragstart="onDragStart(event)" ondragend="onDragEnd(event)" title="' + f.desc +
                         '"><span class="drag-icon">⠿</span>' + f.icon + ' ' + f.label + '</div>').join('') +
    '</div>';
  if (!S.cardSlots.f3) S.cardSlots.f3 = '__auto_reports__';

  const COLORS = ['#4f46e5','#7c3aed','#ff6b9d','#dc2626','#d97706','#10b981','#0891b2','#0284c7','#374151','#0b0b14'];
  document.getElementById('color-palette').innerHTML = COLORS.map(c =>
    '<div class="color-swatch' + (S.cardAccent === c ? ' selected' : '') +
    '" style="background:' + c + '" onclick="setCardAccent(\'' + c + '\')"></div>'
  ).join('');

  const empColSel = document.getElementById('emp-type-col');
  if (empColSel) {
    empColSel.innerHTML = '<option value="">Select column...</option>' +
      S.columns.filter(c => !core.has(c)).map(c =>
        '<option value="' + esc(c) + '"' + (S.empTypeCol === c ? ' selected' : '') + '>' + esc(c) + '</option>'
      ).join('');
    if (S.empTypeCol) populateEmpTypeValues(S.empTypeCol);
  }

  renderCardPreview();
  syncChipStates();
}
function onEmpTypeColChange() {
  S.empTypeCol = document.getElementById('emp-type-col').value;
  if (S.empTypeCol) populateEmpTypeValues(S.empTypeCol);
  else document.getElementById('emp-type-rows').style.display = 'none';
  if (S.viewData.length) renderChart();
}
function populateEmpTypeValues(col) {
  const vals = [...new Set(
    S.rawRows.map(r => String(r[col]||'').trim()).filter(v => v && v !== 'null' && v !== 'undefined')
  )].sort();
  const rows = document.getElementById('emp-type-rows');
  rows.style.display = '';
  ['active','vacant','resigned'].forEach(key => {
    const sel = document.getElementById('emp-val-' + key);
    sel.innerHTML = '<option value="">Value...</option>' +
      vals.map(v => '<option value="' + esc(v) + '"' + (S.empTypeLabels[key] === v ? ' selected' : '') + '>' + esc(v) + '</option>').join('');
    sel.onchange = () => {
      S.empTypeLabels[key] = sel.value;
      buildEmpTypeMap();
      if (S.viewData.length) renderChart();
    };
    const colorInput = document.getElementById('emp-color-' + key);
    colorInput.value = S.empTypeColors[key];
    colorInput.oninput = () => {
      S.empTypeColors[key] = colorInput.value;
      buildEmpTypeMap();
      if (S.viewData.length) renderChart();
    };
  });
  buildEmpTypeMap();
}
function buildEmpTypeMap() {
  S.empTypeMap = {};
  ['active','vacant','resigned'].forEach(key => {
    const v = S.empTypeLabels[key];
    if (v) S.empTypeMap[v] = S.empTypeColors[key];
  });
}
function getNodeBorderColor(node) {
  if (S.empTypeCol && S.empTypeMap) {
    const val = String(node[S.empTypeCol]||'').trim();
    if (S.empTypeMap[val]) return S.empTypeMap[val];
  }
  return S.cardAccent;
}

function onDragStart(e) {
  S.draggingField = e.currentTarget.dataset.field;
  e.currentTarget.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
}
function onDragEnd(e) {
  e.currentTarget.classList.remove('dragging');
  S.draggingField = null;
}
function onZoneDragOver(e)  { e.preventDefault(); e.currentTarget.classList.add('drop-target'); }
function onZoneDragLeave(e) { e.currentTarget.classList.remove('drop-target'); }
function onZoneDrop(e, zone) {
  e.preventDefault();
  e.currentTarget.classList.remove('drop-target');
  if (!S.draggingField) return;
  Object.keys(S.cardSlots).forEach(z => {
    if (S.cardSlots[z] === S.draggingField) S.cardSlots[z] = '';
  });
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
  document.querySelectorAll('.field-chip').forEach(c =>
    c.classList.toggle('placed', placed.has(c.dataset.field))
  );
}
function fieldLabel(id) {
  if (!id) return '';
  const af = AUTO_FIELDS.find(f => f.id === id);
  if (af) return af.icon + ' ' + af.label;
  return id;
}
function fieldSampleVal(id) {
  if (!id) return '';
  if (id === '__auto_reports__')  return '12';
  if (id === '__auto_teamsize__') return '48';
  const row = S.rawRows.find(r => r[id]) || S.rawRows[0] || {};
  return String(row[id]||'Sample').substring(0, 18);
}
function zoneHtml(zoneId, placeholder) {
  const dA = 'ondragover="onZoneDragOver(event)" ondragleave="onZoneDragLeave(event)" ondrop="onZoneDrop(event,\'' + zoneId + '\')"';
  const v  = S.cardSlots[zoneId];
  if (v) return '<div class="card-zone filled" ' + dA + '>' +
    '<span class="zone-field">' + esc(fieldLabel(v)) + '</span>' +
    '<span class="zone-val">' + esc(fieldSampleVal(v)) + '</span>' +
    '<span class="zone-remove" onclick="clearZone(\'' + zoneId + '\')">✕</span>' +
    '</div>';
  return '<div class="card-zone" ' + dA + '><span class="zone-ph">' + placeholder + '</span></div>';
}
function renderCardPreview() {
  const sampleRow = S.rawRows.find(r => r[S.colMap.empName]) || S.rawRows[0] || {};
  const sampleName = String(sampleRow[S.colMap.empName] || 'Asha Kapoor').substring(0, 26);
  const ac = S.cardAccent;
  const ps = S.photoSize, pr = getPhotoRadius();
  const photoDiv = '<div style="width:' + ps + 'px;height:' + ps + 'px;border-radius:' + pr +
    ';background:linear-gradient(150deg,' + ac + '18,' + ac + '30);color:' + ac +
    ';font-size:' + Math.round(ps * 0.28) + 'px;font-weight:700;display:flex;align-items:center;justify-content:center;border:3px solid ' + ac + '55;flex-shrink:0;font-family:\'Space Grotesk\',sans-serif">AK</div>';
  const b1ZoneHtml = '<div class="preview-b1-zone">' + zoneHtml('b1','Body slot') + '</div>';
  const nameBlock = '<div style="width:100%;text-align:center"><div style="font-size:14.5px;font-weight:600;color:var(--ink);margin-bottom:4px;font-family:\'Space Grotesk\',sans-serif">🔒 ' + esc(sampleName) + '</div>' + b1ZoneHtml + '</div>';
  let bodyHtml;
  const pl = S.photoPlacement;
  if (pl === 'none') {
    bodyHtml = '<div style="display:flex;flex-direction:column;gap:6px">' + nameBlock + '</div>';
  } else if (pl === 'top') {
    bodyHtml = '<div style="display:flex;flex-direction:column;align-items:center;gap:10px">' + photoDiv + nameBlock + '</div>';
  } else if (pl === 'left') {
    bodyHtml = '<div style="display:flex;flex-direction:row;align-items:flex-start;gap:10px">' + photoDiv + '<div style="flex:1;min-width:0">' + nameBlock + '</div></div>';
  } else {
    bodyHtml = '<div style="display:flex;flex-direction:row-reverse;align-items:flex-start;gap:10px">' + photoDiv + '<div style="flex:1;min-width:0">' + nameBlock + '</div></div>';
  }
  document.getElementById('card-preview').innerHTML =
    '<div class="preview-card" style="border-top-color:' + ac + '">' +
      '<div class="preview-card-header">' + zoneHtml('h1','H1') + zoneHtml('h2','H2') + zoneHtml('h3','H3') + '</div>' +
      '<div class="preview-card-body">' + bodyHtml + '</div>' +
      '<div class="preview-card-footer">' + zoneHtml('f1','F1') + zoneHtml('f2','F2') + zoneHtml('f3','F3') + '</div>' +
    '</div>';
}
function setCardAccent(color) {
  S.cardAccent = color;
  document.querySelectorAll('.color-swatch').forEach(s =>
    s.classList.toggle('selected', s.style.background === color || rgbToHex(s.style.background) === color)
  );
  renderCardPreview();
}
function rgbToHex(rgb) {
  if (!rgb) return '';
  if (rgb[0] === '#') return rgb.toLowerCase();
  const m = rgb.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
  if (!m) return rgb.toLowerCase();
  return '#' + [1,2,3].map(i => parseInt(m[i]).toString(16).padStart(2,'0')).join('').toLowerCase();
}
function setPhotoShape(shape) {
  S.photoShape = shape;
  document.querySelectorAll('.shape-btn').forEach(b => {
    if (b.dataset.shape) b.classList.toggle('selected', b.dataset.shape === shape);
  });
  renderCardPreview();
  if (S.viewData.length) renderChart();
}
function setPhotoPlacement(p) {
  S.photoPlacement = p;
  document.querySelectorAll('[data-placement]').forEach(b =>
    b.classList.toggle('selected', b.dataset.placement === p)
  );
  renderCardPreview();
  if (S.viewData.length) renderChart();
}
function getPhotoRadius() {
  if (S.photoShape === 'circle')  return '50%';
  if (S.photoShape === 'rounded') return '12px';
  return '4px';
}
function confirmCardDesign() {
  buildEmpTypeMap();
  buildFilterScreen();
  goTo('filter');
}

/* ── Filter screen ───────────────────────────────────────────────────── */
function buildFilterScreen() {
  const core = new Set([S.colMap.empId, S.colMap.empName, S.colMap.managerId].filter(Boolean));
  const filterable = S.columns.filter(c => !core.has(c));
  const cc = document.getElementById('filter-chip-picker');
  cc.innerHTML = filterable.map(col =>
    '<div class="filter-chip ' + (S.filterCols.includes(col) ? 'selected' : '') +
    '" data-col="' + esc(col) + '">' + esc(col) + '</div>'
  ).join('');
  cc.onclick = function(e) {
    const chip = e.target.closest('.filter-chip');
    if (!chip) return;
    const col = chip.dataset.col;
    if (col) toggleFilterCol(col);
  };
  renderFilterPreview();
}
function toggleFilterCol(col) {
  if (S.filterCols.includes(col)) {
    S.filterCols = S.filterCols.filter(c => c !== col);
  } else if (S.filterCols.length < 3) {
    S.filterCols.push(col);
  } else {
    S.filterCols.shift();
    S.filterCols.push(col);
  }
  document.querySelectorAll('.filter-chip').forEach(c =>
    c.classList.toggle('selected', S.filterCols.includes(c.dataset.col))
  );
  renderFilterPreview();
}
function renderFilterPreview() {
  document.getElementById('filter-counter').textContent = S.filterCols.length + ' of 3 filters selected';
  const area = document.getElementById('filter-preview-area');
  if (!S.filterCols.length) {
    area.innerHTML = '<div style="font-size:13px; color:var(--muted-2); padding:14px 0">No filters — full chart will display.</div>';
    return;
  }
  area.innerHTML = '<div class="filter-preview-box">' +
    S.filterCols.map((col, i) => {
      const isLast = i === S.filterCols.length - 1;
      const vals = [...new Set(
        S.rawRows.map(r => String(r[col]||'').trim()).filter(v => v && v !== 'null' && v !== 'undefined')
      )].sort().slice(0, 10);
      return '<div class="fpr-row"><span class="fpr-col">' + esc(col) +
        (isLast ? ' <span class="export-all-pill">EXPORT ALL</span>' : '') + '</span>' +
        '<div class="fpr-vals">' + vals.map(v => '<span class="fv-pill">' + esc(v) + '</span>').join('') +
        (vals.length >= 10 ? '<span style="font-size:11px;color:var(--muted-2)">+ more</span>' : '') +
        '</div></div>';
    }).join('') + '</div>';
}
function launchChart() {
  S.activeFilters = {};
  S.skipDepth = 0;
  buildViewData();
  buildFilterBar();
  renderChart();
  goTo('chart');
}

/* ── Summary fields populate ─────────────────────────────────────────── */
function populateSummaryFields() {
  const core = new Set([S.colMap.empId, S.colMap.empName, S.colMap.managerId].filter(Boolean));
  const opts = '<option value="">—</option>' +
               '<option value="__name__">Name</option>' +
               S.columns.filter(c => !core.has(c)).map(c => '<option value="' + esc(c) + '">' + esc(c) + '</option>').join('');
  const s1 = document.getElementById('summary-field1');
  const s2 = document.getElementById('summary-field2');
  if (s1) { s1.innerHTML = opts; if (S.summaryField1) s1.value = S.summaryField1; }
  if (s2) { s2.innerHTML = opts; if (S.summaryField2) s2.value = S.summaryField2; }
  document.getElementById('depth-select').value = S.skipDepth;
}
function toggleManagerMode() {
  S.managerMode = !S.managerMode;
  document.getElementById('mgr-mode-btn').classList.toggle('active', S.managerMode);
  document.getElementById('summary-fields-wrap').style.display = S.managerMode ? 'flex' : 'none';
  const stat = document.getElementById('stat-mgr-mode');
  if (stat) stat.style.display = S.managerMode ? 'flex' : 'none';
  renderChart();
}
function isManager(nodeId) { return (S.childMap[nodeId] || []).length > 0; }

/* ── View data builder ───────────────────────────────────────────────── */
function buildViewData() {
  const { empId, empName, managerId } = S.colMap;
  let nodes = S.rawRows.map(row => {
    const id  = String(row[empId]||'').replace(/\.0$/,'').trim();
    const mgr = managerId ? String(row[managerId]||'').replace(/\.0$/,'').trim() : '';
    const node = { id, name: String(row[empName]||'Unknown'), manager: mgr };
    S.columns.forEach(col => { node[col] = String(row[col]||''); });
    return node;
  }).filter(n => n.id && !S.removedIds.has(n.id));

  const validIds = new Set(nodes.map(n => n.id));
  nodes.forEach(n => { if (S.managerOverrides.hasOwnProperty(n.id)) n.manager = S.managerOverrides[n.id]; });
  nodes.forEach(n => { if (!validIds.has(n.manager) || n.manager === n.id) n.manager = ''; });

  const hasFilter = Object.values(S.activeFilters).some(v => v);
  if (hasFilter) {
    const matching = new Set(
      nodes.filter(n => Object.entries(S.activeFilters).every(([c, v]) => !v || n[c] === v)).map(n => n.id)
    );
    const byId = Object.fromEntries(nodes.map(n => [n.id, n]));
    const keep = new Set(matching);
    matching.forEach(id => {
      let cur = byId[id];
      const vis = new Set();
      while (cur && cur.manager && byId[cur.manager] && !vis.has(cur.id)) {
        vis.add(cur.id);
        keep.add(cur.manager);
        cur = byId[cur.manager];
      }
    });
    nodes = nodes.filter(n => keep.has(n.id));
  }

  S.viewData = nodes;
  S.childMap = {};
  nodes.forEach(n => {
    if (!S.childMap[n.manager]) S.childMap[n.manager] = [];
    S.childMap[n.manager].push(n);
  });

  S.descCount = {};
  function calcD(id, vis) {
    if (vis.has(id)) return 0;
    vis.add(id);
    if (S.descCount[id] !== undefined) return S.descCount[id];
    const kids = S.childMap[id] || [];
    S.descCount[id] = kids.reduce((s, k) => s + 1 + calcD(k.id, vis), 0);
    return S.descCount[id];
  }
  nodes.filter(n => !n.manager).forEach(r => calcD(r.id, new Set()));

  S.nodeHeight = {};
  function calcH(id, vis) {
    if (vis.has(id)) return 0;
    vis.add(id);
    if (S.nodeHeight[id] !== undefined) return S.nodeHeight[id];
    const kids = S.childMap[id] || [];
    S.nodeHeight[id] = kids.length ? 1 + Math.max(...kids.map(k => calcH(k.id, vis))) : 0;
    return S.nodeHeight[id];
  }
  nodes.filter(n => !n.manager).forEach(r => calcH(r.id, new Set()));
  nodes.forEach(n => { if (S.nodeHeight[n.id] === undefined) calcH(n.id, new Set()); });

  S.nodeDepth = {};
  function calcDepth(id, d, vis) {
    if (vis.has(id)) return;
    vis.add(id);
    S.nodeDepth[id] = d;
    (S.childMap[id] || []).forEach(k => calcDepth(k.id, d + 1, vis));
  }
  nodes.filter(n => !n.manager).forEach(r => calcDepth(r.id, 0, new Set()));
  nodes.forEach(n => { if (S.nodeDepth[n.id] === undefined) S.nodeDepth[n.id] = 0; });
}
function childrenOf(id) { return S.childMap[id] || []; }
function countDescendants(id) { return S.descCount[id] || 0; }

/* ── Filter bar (in-chart) ───────────────────────────────────────────── */
function buildFilterBar() {
  const bar = document.getElementById('filter-bar');
  if (!S.filterCols.length) { bar.style.display = 'none'; return; }
  bar.style.display = 'flex';
  const allVals = {};
  S.filterCols.forEach(col => {
    allVals[col] = [...new Set(
      S.rawRows.map(r => String(r[col]||'').trim()).filter(v => v && v !== 'null' && v !== 'undefined')
    )].sort();
  });
  bar.innerHTML = '<span style="font-family:\'JetBrains Mono\',monospace;font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);flex-shrink:0">Filters</span>' +
    S.filterCols.map(col =>
      '<div class="filter-dropdown-wrap"><span class="filter-dropdown-label">' + esc(col) +
      '</span><select class="filter-dropdown" data-filter-col="' + esc(col) + '">' +
      '<option value="">all ' + esc(col) + '</option>' +
      allVals[col].map(v => '<option value="' + esc(v) + '"' + (S.activeFilters[col] === v ? ' selected' : '') + '>' + esc(v) + '</option>').join('') +
      '</select></div>'
    ).join('') +
    (Object.values(S.activeFilters).some(v => v)
      ? '<button class="btn btn-ghost btn-sm" onclick="clearAllFilters()" style="margin-left:auto">Clear all</button>'
      : '');
  bar.querySelectorAll('.filter-dropdown').forEach(sel => {
    sel.addEventListener('change', function() { applyFilter(this.dataset.filterCol, this.value); });
  });
}
function applyFilter(col, val) {
  if (val) S.activeFilters[col] = val;
  else delete S.activeFilters[col];
  requestAnimationFrame(() => setTimeout(() => {
    buildViewData();
    renderChart();
    buildFilterBar();
  }, 0));
}
function clearAllFilters() {
  S.activeFilters = {};
  requestAnimationFrame(() => setTimeout(() => {
    buildViewData();
    renderChart();
    buildFilterBar();
  }, 0));
}
function setSkipDepth(n) {
  S.skipDepth = n;
  const ds = document.getElementById('depth-select');
  if (ds) ds.value = n;
  renderChart();
}

/* ── Chart background controls ───────────────────────────────────────── */
function setChartBg(color) {
  S.chartBgColor = color;
  if (!S.transparentExport) applyChartBg();
}
function toggleTransparent() {
  S.transparentExport = !S.transparentExport;
  const btn = document.getElementById('bg-transparent-btn');
  const inp = document.getElementById('bg-color-input');
  if (btn) btn.classList.toggle('active', S.transparentExport);
  if (inp) inp.disabled = S.transparentExport;
  applyChartBg();
}
function applyChartBg() {
  const wrap = document.getElementById('chart-canvas-wrap');
  if (!wrap) return;
  if (S.transparentExport) {
    wrap.classList.add('transparent-preview');
    wrap.style.background = '';
  } else {
    wrap.classList.remove('transparent-preview');
    wrap.style.background = S.chartBgColor;
  }
}

function getSlotVal(node, slot) {
  const f = S.cardSlots[slot];
  if (!f) return '';
  if (f === '__auto_reports__')  return childrenOf(node.id).length + ' reports';
  if (f === '__auto_teamsize__') return countDescendants(node.id) + ' people';
  return String(node[f]||'').substring(0, 28);
}

/* ── Chart rendering ─────────────────────────────────────────────────── */
function renderChart() {
  const tree = document.getElementById('org-tree');
  tree.innerHTML = '';
  const ds = document.getElementById('depth-select');
  if (ds) ds.value = S.skipDepth;
  let roots;
  if (S.skipDepth > 0) {
    roots = S.viewData.filter(n => (S.nodeDepth[n.id] || 0) === S.skipDepth);
  } else {
    roots = S.childMap[''] || [];
  }
  if (!roots.length) {
    tree.innerHTML = '<div class="no-data">No nodes found. Try a lower "Skip top" value, or clear filters.</div>';
    updateStats(roots);
    return;
  }
  const ul = document.createElement('ul');
  roots.forEach(r => ul.appendChild(mkNodeLI(r, 0)));
  tree.appendChild(ul);
  updateStats(roots);
  clearTimeout(window._fit);
  window._fit = setTimeout(() => fitToScreen(true), 180);
}

function mkNodeLI(node, depth) {
  depth = depth || 0;
  const li = document.createElement('li');
  li.dataset.id = node.id;
  const ac = getNodeBorderColor(node);
  const acLight = ac + '18', acMid = ac + '55';
  const kids = childrenOf(node.id);
  const card = document.createElement('div');
  card.className = 'node-card' + (node.id === S.highlighted ? ' highlighted' : '');
  card.style.borderTopColor = ac;

  const h1 = getSlotVal(node, 'h1'), h2 = getSlotVal(node, 'h2'), h3 = getSlotVal(node, 'h3');
  const f1 = getSlotVal(node, 'f1'), f2 = getSlotVal(node, 'f2');
  const f3 = getSlotVal(node, 'f3') || node.id.substring(0, 14);
  const b1 = getSlotVal(node, 'b1');
  const subtitle = h2;
  const ps = S.photoSize, pr = getPhotoRadius(), pfs = Math.round(ps * 0.28) + 'px';
  const pInline = 'width:' + ps + 'px;height:' + ps + 'px;border-radius:' + pr + ';';
  const initials = node.name.split(' ').map(w => w[0]||'').join('').substring(0,2).toUpperCase();
  const photoUrl = getPhotoUrl(node);

  let photoHtml = '';
  if (photoUrl) {
    photoHtml = '<img class="ncard-photo" src="' + esc(photoUrl) + '" crossorigin="anonymous" style="' + pInline +
      'border:3px solid ' + acMid + ';box-shadow:0 8px 24px ' + ac + '66" ' +
      'onerror="this.onerror=null;this.style.display=\'none\';this.nextElementSibling.style.display=\'flex\'">' +
      '<div class="ncard-photo-fallback" style="display:none;' + pInline +
      'font-size:' + pfs + ';background:linear-gradient(150deg,' + acLight + ',' + ac + '28);color:' + ac +
      ';border:3px solid ' + acMid + ';">' + esc(initials) + '</div>';
  } else if (Object.keys(S.photoMap).length > 0) {
    photoHtml = '<div class="ncard-photo-fallback" style="display:flex;' + pInline +
      'font-size:' + pfs + ';background:linear-gradient(150deg,' + acLight + ',' + ac + '28);color:' + ac +
      ';border:3px solid ' + acMid + ';">' + esc(initials) + '</div>';
  }

  const b1row = b1 ? '<div class="ncard-body-b1">' + esc(b1) + '</div>' : '';
  const textBlock = '<div class="ncard-text-wrap"><div class="ncard-name">' + esc(node.name) + '</div>' +
    (subtitle ? '<div class="ncard-sub">' + esc(subtitle) + '</div>' : '') + b1row + '</div>';

  let bodyHtml;
  const pl = S.photoPlacement;
  if (!photoHtml || pl === 'none') {
    bodyHtml = '<div class="ncard-body-inner" style="flex-direction:column">' + textBlock + '</div>';
  } else if (pl === 'top') {
    bodyHtml = '<div class="ncard-body-inner" style="flex-direction:column;align-items:center"><div style="flex-shrink:0">' + photoHtml + '</div>' + textBlock + '</div>';
  } else if (pl === 'left') {
    bodyHtml = '<div class="ncard-body-inner" style="flex-direction:row;align-items:flex-start"><div style="flex-shrink:0">' + photoHtml + '</div><div style="flex:1;min-width:0">' + textBlock + '</div></div>';
  } else {
    bodyHtml = '<div class="ncard-body-inner" style="flex-direction:row-reverse;align-items:flex-start"><div style="flex-shrink:0">' + photoHtml + '</div><div style="flex:1;min-width:0">' + textBlock + '</div></div>';
  }

  card.innerHTML =
    '<div class="ncard-header" style="background:' + acLight + ';border-bottom-color:' + ac + '33">' +
      '<span class="ncard-slot' + (h1 ? ' has-val' : '') + '" title="' + esc(h1) + '">' + (esc(h1) || '—') + '</span>' +
      '<span class="ncard-slot' + (h2 ? ' has-val' : '') + '" title="' + esc(h2) + '">' + (esc(h2) || '—') + '</span>' +
      '<span class="ncard-slot' + (h3 ? ' has-val' : '') + '" title="' + esc(h3) + '">' + (esc(h3) || '—') + '</span>' +
    '</div>' +
    '<div class="ncard-body">' + bodyHtml + '</div>' +
    '<div class="ncard-footer" style="background:' + acLight + ';border-top-color:' + ac + '33">' +
      '<span class="ncard-slot' + (f1 ? ' has-val' : '') + '" title="' + esc(f1) + '">' + (esc(f1) || '—') + '</span>' +
      '<span class="ncard-slot' + (f2 ? ' has-val' : '') + '" title="' + esc(f2) + '">' + (esc(f2) || '—') + '</span>' +
      '<span class="ncard-slot' + (f3 ? ' has-val' : '') + '" title="' + esc(f3) + '">' + (esc(f3) || node.id.substring(0,14)) + '</span>' +
    '</div>' +
    '<div class="ncard-export-btn" onclick="exportSubtree(event,\'' + esc(node.id) + '\')" style="position:absolute;top:6px;right:30px;width:22px;height:22px;background:#fff;border:1.5px solid var(--line);border-radius:7px;font-size:10px;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 180ms;z-index:8" title="Export this team">📸</div>' +
    '<div class="ncard-edit-btn" onclick="openReassignModal(event,\'' + esc(node.id) + '\')" style="position:absolute;top:6px;right:6px;width:22px;height:22px;background:#fff;border:1.5px solid var(--line);border-radius:7px;font-size:11px;display:flex;align-items:center;justify-content:center;cursor:pointer;opacity:0;transition:opacity 180ms;z-index:8" title="Reassign">✎</div>';

  card.querySelectorAll('.ncard-edit-btn,.ncard-export-btn').forEach(b => {
    card.addEventListener('mouseenter', () => b.style.opacity = '1');
    card.addEventListener('mouseleave', () => b.style.opacity = '0');
  });

  if (kids.length) {
    const cb = document.createElement('div');
    cb.className = 'collapse-btn';
    cb.innerHTML = '▾';
    cb.title = 'Collapse / expand';
    cb.addEventListener('click', e => { e.stopPropagation(); toggleCollapse(li, cb); });
    card.appendChild(cb);
  }
  li.appendChild(card);

  if (kids.length) {
    if (S.managerMode) {
      const managerKids = kids.filter(k => isManager(k.id));
      const leafKids    = kids.filter(k => !isManager(k.id));
      const ul = document.createElement('ul');
      managerKids.forEach(k => ul.appendChild(mkNodeLI(k, depth + 1)));
      if (leafKids.length > 0) ul.appendChild(mkLeafSummaryLI(leafKids, ac));
      li.appendChild(ul);
    } else {
      const ul = document.createElement('ul');
      kids.forEach(k => ul.appendChild(mkNodeLI(k, depth + 1)));
      li.appendChild(ul);
    }
  }
  return li;
}

/* ── Manager-view IC summary list (export-safe absolute positioning) ── */
function mkLeafSummaryLI(leafNodes, ac) {
  const li = document.createElement('li');
  const f1 = S.summaryField1, f2 = S.summaryField2;
  const count = leafNodes.length;

  const AV = 28, PAD_H = 14, PAD_V = 8, TEXT_LH = 16, GAP = 10, HEADER_LBL_H = 22;
  const FF = "font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;";
  const FFD = "font-family:'Space Grotesk',sans-serif;";

  const headerHtml =
    '<div style="position:relative;background:#f5f3ff;border-bottom:1px solid #e9d5ff;padding:9px ' + PAD_H + 'px;' + FF + '">' +
      '<div style="height:' + HEADER_LBL_H + 'px;line-height:' + HEADER_LBL_H + 'px;font-size:11px;font-weight:700;color:#7c3aed;text-transform:uppercase;letter-spacing:0.06em;padding-right:42px;white-space:nowrap;overflow:hidden;font-family:\'JetBrains Mono\',monospace;">ICs (' + count + ')</div>' +
      '<div style="position:absolute;top:8px;right:' + PAD_H + 'px;height:24px;line-height:24px;background:#7c3aed;color:#ffffff;border-radius:999px;padding:0 10px;font-size:10px;font-weight:700;text-align:center;font-family:\'JetBrains Mono\',monospace;">' + count + '</div>' +
    '</div>';

  let rowsHtml = '';
  leafNodes.forEach((n, idx) => {
    const initials = n.name.split(' ').map(w => w[0]||'').join('').substring(0,2).toUpperCase();
    const borderC = getNodeBorderColor(n);
    const photoUrl = getPhotoUrl(n);
    const isLast = idx === leafNodes.length - 1;
    const nameVal = n.name.substring(0, 24);
    const f1IsName = (f1 === '__name__');
    const primaryVal = f1
      ? (f1IsName ? nameVal : (String(n[f1]||'').trim() || nameVal).substring(0, 24))
      : nameVal;
    const showNameSub = f1 && !f1IsName && primaryVal !== nameVal;
    const val2 = f2
      ? (f2 === '__name__' ? n.name.substring(0, 22) : String(n[f2]||'').substring(0, 22))
      : '';

    const numLines  = 1 + (showNameSub ? 1 : 0) + (val2 ? 1 : 0);
    const textTotalH = numLines * TEXT_LH;
    const innerH    = Math.max(AV, textTotalH);
    const totalRowH = innerH + PAD_V * 2;
    const avatarTopY = PAD_V + Math.max(0, Math.round((innerH - AV) / 2));
    const textTopPad = PAD_V + Math.max(0, Math.round((innerH - textTotalH) / 2));
    const textLeftMargin = PAD_H + AV + GAP;

    let avatarHtml;
    if (photoUrl) {
      avatarHtml = '<img src="' + esc(photoUrl) + '" crossorigin="anonymous" style="position:absolute;left:' + PAD_H + 'px;top:' + avatarTopY + 'px;width:' + AV + 'px;height:' + AV + 'px;border-radius:7px;object-fit:cover;object-position:center top;border:2px solid ' + borderC + '55;box-sizing:border-box;display:block;">';
    } else {
      const innerLH = AV - 4;
      avatarHtml = '<div style="position:absolute;left:' + PAD_H + 'px;top:' + avatarTopY + 'px;width:' + AV + 'px;height:' + AV + 'px;border-radius:7px;background:' + borderC + '1f;color:' + borderC + ';border:2px solid ' + borderC + '55;box-sizing:border-box;text-align:center;font-size:12px;font-weight:700;line-height:' + innerLH + 'px;' + FFD + '">' + esc(initials) + '</div>';
    }

    let textLines = '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:12.5px;font-weight:600;color:#0b0b14;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;' + FFD + '">' + esc(primaryVal) + '</div>';
    if (showNameSub) {
      textLines += '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:10.5px;color:#6b6b80;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;' + FF + '">' + esc(nameVal) + '</div>';
    }
    if (val2) {
      textLines += '<div style="height:' + TEXT_LH + 'px;line-height:' + TEXT_LH + 'px;font-size:10.5px;color:#9a9aae;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-family:\'JetBrains Mono\',monospace;">' + esc(val2) + '</div>';
    }

    const rowBorder = isLast ? '' : 'border-bottom:1px solid #ececf2;';
    rowsHtml +=
      '<div style="position:relative;background:#ffffff;height:' + totalRowH + 'px;' + rowBorder + 'box-sizing:border-box;">' +
        avatarHtml +
        '<div style="margin-left:' + textLeftMargin + 'px;padding-top:' + textTopPad + 'px;padding-right:' + PAD_H + 'px;">' + textLines + '</div>' +
      '</div>';
  });

  const card = document.createElement('div');
  card.className = 'summary-list-card';
  card.innerHTML = headerHtml + rowsHtml;
  li.appendChild(card);
  return li;
}

/* ── Collapse / expand ───────────────────────────────────────────────── */
function toggleCollapse(li, btn) {
  li.classList.toggle('collapsed');
  const c = li.classList.contains('collapsed');
  const childUl = li.querySelector(':scope > ul');
  if (childUl) childUl.style.display = c ? 'none' : '';
  btn.innerHTML = c ? '▸' : '▾';
  btn.style.color = c ? 'var(--warn)' : '';
  li.querySelector('.node-card').classList.toggle('collapsed-node', c);
  setTimeout(updateStats, 60);
}
function expandAll() {
  document.querySelectorAll('li.collapsed').forEach(li => {
    li.classList.remove('collapsed');
    const u = li.querySelector(':scope > ul');
    if (u) u.style.display = '';
    const card = li.querySelector('.node-card');
    if (card) card.classList.remove('collapsed-node');
    const b = li.querySelector('.collapse-btn');
    if (b) { b.innerHTML = '▾'; b.style.color = ''; }
  });
  setTimeout(updateStats, 60);
}
function collapseAll() {
  document.querySelectorAll('li').forEach(li => {
    if (!li.parentElement || !li.parentElement.parentElement || !li.parentElement.parentElement.closest('li')) return;
    if (li.querySelector(':scope > ul')) {
      li.classList.add('collapsed');
      const u = li.querySelector(':scope > ul');
      if (u) u.style.display = 'none';
      const card = li.querySelector('.node-card');
      if (card) card.classList.add('collapsed-node');
      const b = li.querySelector('.collapse-btn');
      if (b) { b.innerHTML = '▸'; b.style.color = 'var(--warn)'; }
    }
  });
  setTimeout(updateStats, 60);
}

/* ── Stats ───────────────────────────────────────────────────────────── */
function updateStats(roots) {
  if (!roots) {
    roots = S.skipDepth > 0
      ? S.viewData.filter(n => (S.nodeDepth[n.id] || 0) === S.skipDepth)
      : (S.childMap[''] || []);
  }
  document.getElementById('stat-total').textContent = S.viewData.length;
  document.getElementById('stat-roots').textContent = roots.length;
  let visCount = 0;
  document.querySelectorAll('.node-card').forEach(card => {
    if (!card.closest('li.collapsed > ul')) visCount++;
  });
  document.getElementById('stat-vis').textContent = visCount;
  document.getElementById('stat-filtered').style.display =
    Object.values(S.activeFilters).some(v => v) ? 'flex' : 'none';
  const mgrStat = document.getElementById('stat-mgr-mode');
  const mgrVal  = document.getElementById('stat-mgr-val');
  if (mgrStat) {
    mgrStat.style.display = S.managerMode ? 'flex' : 'none';
    if (S.managerMode && mgrVal) {
      mgrVal.textContent = S.viewData.filter(n => !isManager(n.id)).length + ' ICs in lists';
    }
  }
}

/* ── Pan / zoom ──────────────────────────────────────────────────────── */
function cwrap()    { return document.getElementById('chart-canvas-wrap'); }
function ccontent() { return document.getElementById('chart-canvas-content'); }
function applyZoom(z) {
  S.zoom = Math.max(0.1, Math.min(3, z));
  ccontent().style.transform = 'scale(' + S.zoom + ')';
  document.getElementById('zoom-level').textContent = Math.round(S.zoom * 100) + '%';
}
function zoomBy(d) { applyZoom(S.zoom + d); }
function fitToScreen(andCenter) {
  requestAnimationFrame(() => {
    const tree = document.getElementById('org-tree');
    const wrap = cwrap();
    if (!tree || !wrap) return;
    const tw = tree.scrollWidth, th = tree.scrollHeight,
          aw = wrap.clientWidth - 100, ah = wrap.clientHeight - 100;
    if (tw < 10 || th < 10) return;
    applyZoom(Math.max(0.12, Math.min(1, aw / tw, ah / th)));
    if (andCenter) setTimeout(centerView, 70);
  });
}
function centerView() {
  const wrap = cwrap();
  const tree = document.getElementById('org-tree');
  if (!wrap || !tree) return;
  const sw = tree.scrollWidth * S.zoom;
  wrap.scrollLeft = Math.max(0, (sw - wrap.clientWidth) / 2);
  wrap.scrollTop = 0;
}

let _panning = false, _px, _py, _psl, _pst;
function initPan() {
  const wrap = cwrap();
  if (!wrap) return;
  wrap.onmousedown = e => {
    if (e.target.closest('.node-card,.summary-list-card,.collapse-btn')) return;
    _panning = true;
    _px = e.clientX; _py = e.clientY;
    _psl = wrap.scrollLeft; _pst = wrap.scrollTop;
    wrap.style.cursor = 'grabbing';
  };
  window.onmousemove = e => {
    if (!_panning) return;
    cwrap().scrollLeft = _psl - (e.clientX - _px);
    cwrap().scrollTop  = _pst - (e.clientY - _py);
  };
  window.onmouseup = () => {
    _panning = false;
    if (cwrap()) cwrap().style.cursor = '';
  };
  wrap.addEventListener('wheel', e => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      zoomBy(e.deltaY < 0 ? 0.08 : -0.08);
    }
  }, { passive: false });
}

/* ── Search ──────────────────────────────────────────────────────────── */
function initSearch() {
  const input = document.getElementById('chart-search');
  const box = document.getElementById('chart-search-results');
  if (!input) return;
  function positionBox() {
    const r = input.getBoundingClientRect();
    box.style.top = (r.bottom + 4) + 'px';
    box.style.left = r.left + 'px';
    box.style.width = Math.max(280, r.width) + 'px';
  }
  input.addEventListener('input', function() {
    const q = this.value.trim().toLowerCase();
    if (!q) { box.classList.remove('visible'); return; }
    const hits = S.viewData.filter(n =>
      n.name.toLowerCase().includes(q) || n.id.toLowerCase().includes(q)
    ).slice(0, 10);
    box.innerHTML = hits.length
      ? hits.map(n => '<div class="sr-item" onclick="highlightNode(\'' + esc(n.id) + '\')"><div class="sr-name">' + esc(n.name) + '</div><div class="sr-sub">' + esc(n.id) + '</div></div>').join('')
      : '<div class="sr-item" style="color:var(--muted-2);font-size:13px;padding:14px">No results</div>';
    positionBox();
    box.classList.add('visible');
  });
  input.addEventListener('focus', () => { if (input.value.trim()) positionBox(); });
  document.addEventListener('click', e => {
    if (!e.target.closest('.search-wrap')) box.classList.remove('visible');
  });
  window.addEventListener('resize', () => {
    if (box.classList.contains('visible')) positionBox();
  });
}
function highlightNode(id) {
  document.querySelectorAll('.node-card.highlighted').forEach(c => c.classList.remove('highlighted'));
  S.highlighted = id;
  expandAll();
  const li = document.querySelector('li[data-id="' + CSS.escape(id) + '"]');
  if (li) {
    const card = li.querySelector('.node-card');
    if (card) {
      card.classList.add('highlighted');
      setTimeout(() => {
        const r = card.getBoundingClientRect();
        const w = cwrap();
        const wr = w.getBoundingClientRect();
        w.scrollTo({
          left: w.scrollLeft + (r.left - wr.left) - wr.width / 2 + r.width / 2,
          top:  w.scrollTop  + (r.top  - wr.top)  - wr.height / 2 + r.height / 2,
          behavior: 'smooth',
        });
      }, 80);
    }
  }
  document.getElementById('chart-search').value = '';
  document.getElementById('chart-search-results').classList.remove('visible');
}

/* ── Exports: shared helpers ─────────────────────────────────────────── */
function triggerDownload(blob, fname) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = fname;
  a.click();
  URL.revokeObjectURL(url);
}
function csvEsc(v) {
  return '"' + String(v||'').replace(/"/g, '""').replace(/[\r\n]+/g, ' ') + '"';
}
function buildCSVContent() {
  const cols = [
    S.colMap.empId, S.colMap.empName, S.colMap.managerId,
    ...S.columns.filter(c => c !== S.colMap.empId && c !== S.colMap.empName && c !== S.colMap.managerId)
  ].filter(Boolean);
  return cols.map(csvEsc).join(',') + '\n' +
         S.viewData.map(n => cols.map(c => csvEsc(n[c]||'')).join(',')).join('\n');
}
function downloadCSV() {
  triggerDownload(new Blob([buildCSVContent()], { type: 'text/csv;charset=utf-8;' }), 'nodely_export.csv');
}
function makeOverlay(title, sub) {
  const o = document.createElement('div');
  o.className = 'export-overlay';
  o.innerHTML = '<div class="export-spinner"></div>' +
    '<div class="display" style="font-weight:600;font-size:15px;color:var(--ink);margin-top:10px">' + title + '</div>' +
    '<div class="mono" style="font-size:12px;color:var(--muted-2);margin-top:4px">' + sub + '</div>';
  return o;
}

function _saveCollapsedState() {
  const ids = [];
  document.querySelectorAll('li.collapsed').forEach(li => { if (li.dataset.id) ids.push(li.dataset.id); });
  return ids;
}
function _restoreCollapsedState(ids) {
  if (!ids || !ids.length) return;
  const s = new Set(ids);
  document.querySelectorAll('li[data-id]').forEach(li => {
    if (s.has(li.dataset.id)) {
      const ul = li.querySelector(':scope > ul');
      if (ul) {
        li.classList.add('collapsed');
        ul.style.display = 'none';
        const card = li.querySelector('.node-card');
        if (card) card.classList.add('collapsed-node');
        const b = li.querySelector('.collapse-btn');
        if (b) { b.innerHTML = '▸'; b.style.color = 'var(--warn)'; }
      }
    }
  });
  setTimeout(updateStats, 60);
}

async function buildRenderStage() {
  const savedCollapsed = _saveCollapsedState();
  expandAll();
  await new Promise(r => setTimeout(r, 400));
  if (document.fonts && document.fonts.ready) await document.fonts.ready;
  await new Promise(r => setTimeout(r, 200));

  const orgTree = document.getElementById('org-tree');
  const container = document.createElement('div');
  container.className = 'export-stage-root';
  const stageBg = S.transparentExport ? 'transparent' : S.chartBgColor;
  container.style.cssText = 'position:fixed;top:0;left:0;background:' + stageBg + ';padding:48px 64px 80px 64px;display:inline-block;z-index:9998;pointer-events:none;overflow:visible';

  const clone = orgTree.cloneNode(true);
  clone.querySelectorAll('.collapse-btn,.ncard-edit-btn,.ncard-export-btn').forEach(el => el.remove());
  clone.querySelectorAll('li.collapsed').forEach(li => {
    li.classList.remove('collapsed');
    const ul = li.querySelector(':scope > ul');
    if (ul) ul.style.removeProperty('display');
    const card = li.querySelector('.node-card');
    if (card) card.classList.remove('collapsed-node');
  });
  clone.querySelectorAll('.node-card,.summary-list-card').forEach(c => {
    c.style.removeProperty('opacity');
    c.style.removeProperty('transform');
    c.style.setProperty('overflow', 'visible', 'important');
  });

  container.appendChild(clone);
  document.body.appendChild(container);
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
  await new Promise(r => setTimeout(r, 300));
  _restoreCollapsedState(savedCollapsed);
  return { stage: container, wrapper: container };
}

async function renderToCanvas(stageObj) {
  const el = stageObj.stage;
  const w = el.scrollWidth || el.offsetWidth;
  const h = el.scrollHeight || el.offsetHeight;
  const bg = S.transparentExport ? null : S.chartBgColor;
  return html2canvas(el, {
    backgroundColor: bg,
    scale: 2,
    useCORS: true,
    logging: false,
    allowTaint: true,
    foreignObjectRendering: false,
    width:  Math.ceil(w),
    height: Math.ceil(h),
    windowWidth:  Math.ceil(w) + 200,
    windowHeight: Math.ceil(h) + 200,
    scrollX: 0, scrollY: 0,
    x: 0, y: 0,
  });
}

async function exportPNG() {
  const overlay = makeOverlay('Rendering org chart...', 'Capturing at 2× resolution');
  document.body.appendChild(overlay);
  const savedZoom = S.zoom;
  applyZoom(1);
  await new Promise(r => setTimeout(r, 140));
  let stage;
  try {
    stage = await buildRenderStage();
    const canvas = await renderToCanvas(stage);
    const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
    const fp = Object.values(S.activeFilters).filter(Boolean).map(v => v.replace(/[^a-zA-Z0-9]/g,'_')).join('_');
    const mode = S.managerMode ? '_mgr_view' : '';
    await new Promise(res => canvas.toBlob(blob => {
      if (blob) triggerDownload(blob, 'nodely_' + (fp ? fp + '_' : '') + mode + stamp + '.png');
      res();
    }, 'image/png'));
  } catch (e) {
    alert('PNG export failed: ' + e.message);
  } finally {
    if (stage && stage.wrapper) stage.wrapper.remove();
    overlay.remove();
    applyZoom(savedZoom);
  }
}

async function exportSubtree(e, nodeId) {
  e.stopPropagation();
  const node = S.viewData.find(n => n.id === nodeId);
  if (!node) return;
  const includeIds = new Set([nodeId]);
  function collectDesc(id) {
    (S.childMap[id]||[]).forEach(k => { includeIds.add(k.id); collectDesc(k.id); });
  }
  collectDesc(nodeId);
  const overlay = makeOverlay('Exporting ' + node.name + "'s team (" + includeIds.size + ')...', 'Subtree PNG');
  document.body.appendChild(overlay);

  const savedViewData    = S.viewData;
  const savedChildMap    = S.childMap;
  const savedDescCount   = S.descCount;
  const savedNodeHeight  = S.nodeHeight;
  const savedNodeDepth   = S.nodeDepth;
  const savedSkipDepth   = S.skipDepth;
  const hadOverride      = S.managerOverrides.hasOwnProperty(nodeId);
  const prevOverride     = S.managerOverrides[nodeId];

  S.viewData = savedViewData.filter(n => includeIds.has(n.id));
  S.managerOverrides[nodeId] = '';
  S.skipDepth = 0;
  S.childMap = {};
  S.viewData.forEach(n => {
    const mgr = (n.id === nodeId) ? '' : n.manager;
    if (!S.childMap[mgr]) S.childMap[mgr] = [];
    S.childMap[mgr].push(n);
  });
  S.descCount = {}; S.nodeHeight = {}; S.nodeDepth = {};
  function cD(id)  { const k = S.childMap[id]||[]; S.descCount[id]   = k.reduce((s,c)=>s+1+cD(c.id),0); return S.descCount[id]; }
  function cH(id)  { const k = S.childMap[id]||[]; S.nodeHeight[id]  = k.length ? 1 + Math.max(...k.map(c=>cH(c.id))) : 0; return S.nodeHeight[id]; }
  function cDep(id, d) { S.nodeDepth[id] = d; (S.childMap[id]||[]).forEach(k => cDep(k.id, d+1)); }
  cD(nodeId); cH(nodeId); cDep(nodeId, 0);

  const savedZoom = S.zoom;
  applyZoom(1);
  renderChart();
  await new Promise(r => setTimeout(r, 400));

  let stage;
  try {
    stage = await buildRenderStage();
    const canvas = await renderToCanvas(stage);
    const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
    const safeName = node.name.replace(/[^a-zA-Z0-9]/g,'_');
    await new Promise(res => canvas.toBlob(blob => {
      if (blob) triggerDownload(blob, 'team_' + safeName + '_' + stamp + '.png');
      res();
    }, 'image/png'));
  } catch (ex) {
    alert('Subtree export failed: ' + ex.message);
  } finally {
    if (stage && stage.wrapper) stage.wrapper.remove();
    overlay.remove();
    applyZoom(savedZoom);
    if (hadOverride) S.managerOverrides[nodeId] = prevOverride;
    else delete S.managerOverrides[nodeId];
    S.viewData   = savedViewData;
    S.childMap   = savedChildMap;
    S.descCount  = savedDescCount;
    S.nodeHeight = savedNodeHeight;
    S.nodeDepth  = savedNodeDepth;
    S.skipDepth  = savedSkipDepth;
    renderChart();
  }
}

/* ── PPTX export (raw OOXML) ─────────────────────────────────────────── */
const SW = 12192000, SH = 6858000;
function pptxRect(id, x, y, cx, cy, fill) {
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="r' + id + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="' + fill + '"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';
}
function pptxTxt(id, x, y, cx, cy, text, sz, bold, color, algn) {
  algn = algn || 'ctr';
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="t' + id + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="' + algn + '"/><a:r><a:rPr lang="en-US" sz="' + sz + '" b="' + (bold?1:0) + '" dirty="0"><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></a:rPr><a:t>' + xe(text) + '</a:t></a:r></a:p></p:txBody></p:sp>';
}
function pptxImg(id, x, y, cx, cy, rId) {
  return '<p:pic><p:nvPicPr><p:cNvPr id="' + id + '" name="img' + id + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
}
function pptxSlide(bg, content, rels) {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="' + bg + '"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + SW + '" cy="' + SH + '"/><a:chOff x="0" y="0"/><a:chExt cx="' + SW + '" cy="' + SH + '"/></a:xfrm></p:grpSpPr>' + content + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',
    rels
  ];
}

async function buildPPTXBlob(imgB64, cW, cH, titleSuffix) {
  titleSuffix = titleSuffix || '';
  const ac = S.cardAccent.replace('#','');
  const stamp = new Date().toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' });
  const activeF = Object.entries(S.activeFilters).filter(([,v]) => v);
  const filterLine = activeF.map(([k,v]) => k + ': ' + v).join('  |  ') || (titleSuffix || 'All employees');
  const roots = (S.childMap[''] || []).length;
  const mgrCount = S.viewData.filter(n => isManager(n.id)).length;
  const modeNote = S.managerMode ? ' | Manager view (ICs in lists)' : '';

  // Cover slide
  const [s1xml, s1rels] = pptxSlide('FAFAF9',
    pptxRect(2, 0, 0, SW, Math.round(SH * 0.52), ac) +
    pptxRect(3, 0, Math.round(SH * 0.52), SW, Math.round(SH * 0.48), 'FFFFFF') +
    pptxTxt(4, Math.round(SW * 0.08), Math.round(SH * 0.12),Math.round(SW * 0.84), Math.round(SH * 0.22), 'Org Chart', 7600, true, 'FFFFFF', 'l') +
    pptxTxt(5, Math.round(SW * 0.08), Math.round(SH * 0.35), Math.round(SW * 0.84), 420000, filterLine + modeNote, 2200, true, 'FFFFFF', 'l') +
    pptxTxt(6, Math.round(SW * 0.08), Math.round(SH * 0.44), Math.round(SW * 0.84), 340000, 'Generated: ' + stamp, 1500, false, 'C7C9FF', 'l') +
    pptxTxt(7, Math.round(SW * 0.08), Math.round(SH * 0.59), Math.round(SW * 0.38), 400000, String(S.viewData.length), 5200, true, ac, 'l') +
    pptxTxt(8, Math.round(SW * 0.08), Math.round(SH * 0.74), Math.round(SW * 0.38), 310000, 'Total Employees', 1600, false, '6B6B80', 'l') +
    pptxTxt(9, Math.round(SW * 0.55), Math.round(SH * 0.59), Math.round(SW * 0.35), 400000, String(mgrCount), 4000, true, '6B6B80', 'l') +
    pptxTxt(10, Math.round(SW * 0.55), Math.round(SH * 0.74), Math.round(SW * 0.35), 310000, 'Managers', 1600, false, '6B6B80', 'l'),
    'Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
  );

  // Chart image slide — fit to 16:9 maintaining aspect
  const imgAspect = cW / cH, slideAspect = SW / SH;
  let iW, iH, iX, iY;
  if (imgAspect >= slideAspect) {
    iW = SW; iH = Math.round(SW / imgAspect);
    iX = 0;  iY = Math.round((SH - iH) / 2);
  } else {
    iH = SH; iW = Math.round(SH * imgAspect);
    iX = Math.round((SW - iW) / 2); iY = 0;
  }
  const capY = SH - Math.round(SH * 0.065);
  const [s2xml,] = pptxSlide('FFFFFF',
    pptxImg(20, iX, iY, iW, iH, 'rId2') +
    pptxTxt(21, Math.round(SW * 0.04), capY, Math.round(SW * 0.92), Math.round(SH * 0.055),
      filterLine + modeNote + ' · ' + stamp + ' · ' + S.viewData.length + ' employees',
      1000, false, '6B6B80', 'r'),
    'Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>'
  );

  // Summary dashboard slide
  const statItems = [
    { label: 'Total employees', val: S.viewData.length, color: ac },
    { label: 'Managers',        val: mgrCount,          color: '7c3aed' },
    { label: 'Roots',           val: roots,             color: '0891b2' },
    { label: 'Mode',            val: S.managerMode ? 'Manager view' : 'Full tree', color: '10b981' },
  ];
  const boxW = Math.round(SW * 0.19), boxH = Math.round(SH * 0.3);
  const gap  = Math.round((SW - boxW * 4) * 0.2);
  const totalBW = boxW * 4 + gap * 3;
  const bStartX = Math.round((SW - totalBW) / 2);
  const bY = Math.round(SH * 0.3);
  let sc = pptxRect(2, 0, 0, SW, Math.round(SH * 0.2), ac) +
    pptxTxt(3, Math.round(SW * 0.04), 0, Math.round(SW * 0.6), Math.round(SH * 0.2), 'Summary Dashboard', 1800, true, 'FFFFFF', 'l') +
    pptxTxt(4, Math.round(SW * 0.65), 0, Math.round(SW * 0.3), Math.round(SH * 0.2), stamp, 1200, false, 'C7C9FF', 'r') +
    pptxTxt(5, Math.round(SW * 0.04), Math.round(SH * 0.22), Math.round(SW * 0.92), Math.round(SH * 0.06), filterLine + modeNote, 1600, false, '6B6B80', 'l');
  statItems.forEach((st, i) => {
    const bx = bStartX + i * (boxW + gap);
    sc += pptxRect(10 + i * 2, bx, bY, boxW, boxH, 'FAFAF9') +
          pptxRect(11 + i * 2, bx, bY, boxW, Math.round(boxH * 0.05), st.color) +
          pptxTxt(20 + i * 2, bx, bY + Math.round(boxH * 0.1),  boxW, Math.round(boxH * 0.52), String(st.val), 4800, true, st.color) +
          pptxTxt(21 + i * 2, bx, bY + Math.round(boxH * 0.7),  boxW, Math.round(boxH * 0.28), st.label, 1300, false, '6B6B80');
  });
  const [s3xml,] = pptxSlide('FFFFFF', sc,
    'Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
  );

  const mkRel = r => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><' + r + '</Relationships>';

  const BP = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
    'ppt/presentation.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst><p:sldSz cx="' + SW + '" cy="' + SH + '" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>',
    'ppt/theme/theme1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="NodelyTheme"><a:themeElements><a:clrScheme name="NodelyScheme"><a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1><a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1><a:dk2><a:srgbClr val="0B0B14"/></a:dk2><a:lt2><a:srgbClr val="FAFAF9"/></a:lt2><a:accent1><a:srgbClr val="' + ac + '"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="FFD166"/></a:accent3><a:accent4><a:srgbClr val="FF6B9D"/></a:accent4><a:accent5><a:srgbClr val="7C3AED"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6><a:hlink><a:srgbClr val="' + ac + '"/></a:hlink><a:folHlink><a:srgbClr val="6B6B80"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>',
    'ppt/slideMasters/slideMaster1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lstStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></a:lstStyle></p:titleStyle><p:bodyStyle><a:lstStyle/></p:bodyStyle><p:otherStyle><a:lstStyle/></p:otherStyle></p:txStyles></p:sldMaster>',
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>',
    'ppt/slideLayouts/slideLayout1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
  };

  const zip = new JSZip();
  Object.entries(BP).forEach(([k, v]) => zip.file(k, v));
  zip.file('ppt/slides/slide1.xml', s1xml);
  zip.file('ppt/slides/_rels/slide1.xml.rels', mkRel(s1rels));
  zip.file('ppt/slides/slide2.xml', s2xml);
  zip.file('ppt/slides/_rels/slide2.xml.rels', mkRel('Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>'));
  zip.file('ppt/slides/slide3.xml', s3xml);
  zip.file('ppt/slides/_rels/slide3.xml.rels', mkRel('Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'));
  zip.file('ppt/media/image1.png', imgB64, { base64: true });
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
  });
}

async function exportPPTX() {
  if (typeof JSZip === 'undefined') { alert('ZIP library failed to load.'); return; }
  const overlay = makeOverlay('Building PowerPoint...', 'Rendering then packaging');
  document.body.appendChild(overlay);
  const savedZoom = S.zoom;
  applyZoom(1);
  await new Promise(r => setTimeout(r, 140));
  let stage;
  try {
    stage = await buildRenderStage();
    const canvas = await renderToCanvas(stage);
    const blob = await buildPPTXBlob(canvas.toDataURL('image/png').split(',')[1], canvas.width, canvas.height);
    const dp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const fp = Object.values(S.activeFilters).filter(Boolean).map(v => v.replace(/[^a-zA-Z0-9]/g, '_')).join('_');
    const mode = S.managerMode ? '_mgr' : '';
    triggerDownload(blob, 'nodely_' + (fp ? fp + '_' : '') + mode + dp + '.pptx');
  } catch (e) {
    alert('PPTX failed: ' + e.message);
    console.error(e);
  } finally {
    if (stage && stage.wrapper) stage.wrapper.remove();
    overlay.remove();
    applyZoom(savedZoom);
  }
}

/* ── Export-all (loop over last filter column) ───────────────────────── */
async function exportAll() {
  if (typeof JSZip === 'undefined') { alert('ZIP library failed to load.'); return; }
  const lastFilterCol = S.filterCols[S.filterCols.length - 1] || null;
  if (!lastFilterCol) { await _exportAllSingleView(); return; }

  const parentFilters = Object.entries(S.activeFilters).filter(([k]) => k !== lastFilterCol);
  const relevantRows = S.rawRows.filter(row =>
    parentFilters.every(([col, val]) => !val || String(row[col]||'').trim() === val)
  );
  const lastVals = [...new Set(
    relevantRows.map(r => String(r[lastFilterCol]||'').trim()).filter(v => v && v !== 'null' && v !== 'undefined')
  )].sort();
  if (!lastVals.length) { await _exportAllSingleView(); return; }

  const overlay = document.createElement('div');
  overlay.className = 'export-overlay';
  overlay.innerHTML =
    '<div class="export-spinner"></div>' +
    '<div class="display" style="font-weight:600;font-size:15px;color:var(--ink);margin-top:10px" id="_ea_title">Exporting ' + lastVals.length + ' charts...</div>' +
    '<div id="_ea_step" class="mono" style="font-size:13px;color:var(--ink-2);margin-top:6px">Preparing...</div>' +
    '<div id="_ea_prog" class="mono" style="font-size:11px;color:var(--muted-2);margin-top:2px">0 / ' + lastVals.length + '</div>';
  document.body.appendChild(overlay);

  const savedZoom = S.zoom;
  applyZoom(1);
  await new Promise(r => setTimeout(r, 140));

  const savedFilters = { ...S.activeFilters };
  const outerZip = new JSZip();
  let successCount = 0;

  try {
    for (let i = 0; i < lastVals.length; i++) {
      const val = lastVals[i];
      const safeName = val.replace(/[^a-zA-Z0-9]/g, '_');
      document.getElementById('_ea_step').textContent = '📊 ' + val;
      document.getElementById('_ea_prog').textContent = (i + 1) + ' / ' + lastVals.length;
      S.activeFilters[lastFilterCol] = val;
      buildViewData();
      renderChart();
      await new Promise(r => setTimeout(r, 400));
      outerZip.file(safeName + '/' + safeName + '.csv', buildCSVContent());
      let stage2;
      try {
        stage2 = await buildRenderStage();
        const canvas2 = await renderToCanvas(stage2);
        const png64 = canvas2.toDataURL('image/png').split(',')[1];
        outerZip.file(safeName + '/' + safeName + '.png', png64, { base64: true });
        const pptxBlob = await buildPPTXBlob(png64, canvas2.width, canvas2.height, val);
        outerZip.file(safeName + '/' + safeName + '.pptx', pptxBlob);
        successCount++;
      } finally {
        if (stage2 && stage2.wrapper) stage2.wrapper.remove();
      }
    }
  } finally {
    S.activeFilters = savedFilters;
    buildViewData();
    renderChart();
    buildFilterBar();
    overlay.remove();
    applyZoom(savedZoom);
  }

  if (successCount > 0) {
    const zipBlob = await outerZip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    const dp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    triggerDownload(zipBlob, 'nodely_all_' + dp + '.zip');
  }
}

async function _exportAllSingleView() {
  const overlay = makeOverlay('Exporting current view...', 'PNG + PPTX + CSV bundle');
  document.body.appendChild(overlay);
  const savedZoom = S.zoom;
  applyZoom(1);
  await new Promise(r => setTimeout(r, 140));
  let stage;
  try {
    stage = await buildRenderStage();
    const canvas = await renderToCanvas(stage);
    const dp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const zip = new JSZip();
    zip.file('nodely.csv', buildCSVContent());
    const png64 = canvas.toDataURL('image/png').split(',')[1];
    zip.file('nodely.png', png64, { base64: true });
    const pptxBlob = await buildPPTXBlob(png64, canvas.width, canvas.height);
    zip.file('nodely.pptx', pptxBlob);
    const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    triggerDownload(zipBlob, 'nodely_' + dp + '.zip');
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    if (stage && stage.wrapper) stage.wrapper.remove();
    overlay.remove();
    applyZoom(savedZoom);
  }
}

/* ── Reassign modal ──────────────────────────────────────────────────── */
let _reassignAllNodes = [];
function openReassignModal(e, nodeId) {
  e.stopPropagation();
  S.reassignTarget = nodeId;
  S.reassignPick   = null;
  const node = S.viewData.find(n => n.id === nodeId);
  document.getElementById('reassign-subject').innerHTML = 'Moving <strong>' + esc(node ? node.name : nodeId) + '</strong>';
  document.getElementById('reassign-search').value = '';
  document.getElementById('reassign-confirm-btn').disabled = true;
  document.getElementById('reassign-note').textContent = 'Pick a new manager above';
  _reassignAllNodes = [
    { id: '__root__', name: 'Make root (no manager)', manager: '' },
    ...S.viewData.filter(n => n.id !== nodeId)
  ];
  renderReassignList(_reassignAllNodes);
  document.getElementById('reassign-modal').classList.remove('hidden');
}
function closeReassignModal() {
  document.getElementById('reassign-modal').classList.add('hidden');
  S.reassignTarget = null;
  S.reassignPick   = null;
}
function filterReassignList() {
  const q = document.getElementById('reassign-search').value.trim().toLowerCase();
  renderReassignList(q
    ? _reassignAllNodes.filter(n => n.name.toLowerCase().includes(q) || n.id.toLowerCase().includes(q))
    : _reassignAllNodes
  );
}
function renderReassignList(nodes) {
  document.getElementById('reassign-list').innerHTML = nodes.slice(0, 60).map(n => {
    const isRoot = n.id === '__root__';
    const initials = n.name.split(' ').map(w => w[0]||'').join('').substring(0, 2).toUpperCase();
    return '<div class="modal-emp-row' + (S.reassignPick === n.id ? ' selected' : '') +
           '" onclick="pickReassign(\'' + esc(n.id) + '\',\'' + esc(n.name) + '\')">' +
           '<div class="modal-emp-avatar">' + (isRoot ? '🔼' : esc(initials)) + '</div>' +
           '<div><div class="modal-emp-name">' + esc(n.name) + '</div>' +
           '<div class="modal-emp-sub">' + (isRoot ? 'Will appear as root node' : esc(n.id)) + '</div></div></div>';
  }).join('');
}
function pickReassign(id, name) {
  S.reassignPick = id;
  document.getElementById('reassign-confirm-btn').disabled = false;
  document.getElementById('reassign-note').textContent = '→ ' + name;
  const q = document.getElementById('reassign-search').value.trim().toLowerCase();
  renderReassignList(q
    ? _reassignAllNodes.filter(n => n.name.toLowerCase().includes(q) || n.id.toLowerCase().includes(q))
    : _reassignAllNodes
  );
}
function confirmReassign() {
  if (!S.reassignTarget || !S.reassignPick) return;
  S.managerOverrides[S.reassignTarget] = S.reassignPick === '__root__' ? '' : S.reassignPick;
  closeReassignModal();
  buildViewData();
  renderChart();
}
function removeCurrentNode() {
  if (!S.reassignTarget) return;
  S.removedIds.add(S.reassignTarget);
  closeReassignModal();
  buildViewData();
  renderChart();
}

/* ── Upload screen network background (Nodely signature animation) ──── */
function initUploadNetworkBG() {
  const canvas = document.getElementById('upload-network-canvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const ACCENT = '#4f46e5';
  let w = 0, h = 0, nodes = [];
  const mouse = { x: -9999, y: -9999 };

  function resize() {
    const rect = canvas.parentElement.getBoundingClientRect();
    w = rect.width; h = rect.height;
    canvas.width = w * dpr; canvas.height = h * dpr;
    canvas.style.width = w + 'px'; canvas.style.height = h + 'px';
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.scale(dpr, dpr);
    const count = Math.min(28, Math.max(8, Math.round((w * h) / 52000)));
    nodes = Array.from({ length: count }, () => ({
      x: Math.random() * w, y: Math.random() * h,
      vx: (Math.random() - 0.5) * 0.25,
      vy: (Math.random() - 0.5) * 0.25,
      r: 1.6 + Math.random() * 1.8,
    }));
  }
  function hexA(hex, a) {
    const hh = hex.replace('#', '');
    const r = parseInt(hh.slice(0, 2), 16),
          g = parseInt(hh.slice(2, 4), 16),
          b = parseInt(hh.slice(4, 6), 16);
    return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
  }
  function tick() {
    ctx.clearRect(0, 0, w, h);
    for (const n of nodes) {
      const dx = mouse.x - n.x, dy = mouse.y - n.y;
      const d2 = dx * dx + dy * dy;
      if (d2 < 28000) { n.vx += dx * 0.0008; n.vy += dy * 0.0008; }
      n.vx *= 0.985; n.vy *= 0.985;
      n.x += n.vx; n.y += n.vy;
      if (n.x < 0 || n.x > w) n.vx *= -1;
      if (n.y < 0 || n.y > h) n.vy *= -1;
    }
    for (let i = 0; i < nodes.length; i++) {
      for (let j = i + 1; j < nodes.length; j++) {
        const a = nodes[i], b = nodes[j];
        const dx = a.x - b.x, dy = a.y - b.y;
        const d = Math.sqrt(dx * dx + dy * dy);
        if (d < 160) {
          ctx.strokeStyle = hexA(ACCENT, (1 - d / 160) * 0.18);
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(a.x, a.y);
          ctx.lineTo(b.x, b.y);
          ctx.stroke();
        }
      }
    }
    for (const n of nodes) {
      ctx.fillStyle = hexA(ACCENT, 0.4);
      ctx.beginPath();
      ctx.arc(n.x, n.y, n.r, 0, Math.PI * 2);
      ctx.fill();
    }
    requestAnimationFrame(tick);
  }
  resize();
  tick();
  window.addEventListener('resize', resize);
  canvas.parentElement.addEventListener('mousemove', e => {
    const rect = canvas.getBoundingClientRect();
    mouse.x = e.clientX - rect.left;
    mouse.y = e.clientY - rect.top;
  });
  canvas.parentElement.addEventListener('mouseleave', () => {
    mouse.x = -9999; mouse.y = -9999;
  });
}

/* ── Wire up event listeners ─────────────────────────────────────────── */
document.getElementById('file-input').addEventListener('change', function(e) {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});
document.getElementById('photo-folder-input').addEventListener('change', function(e) {
  if (e.target.files.length) loadFromFileInput(e.target.files);
});
const dz = document.getElementById('upload-dropzone');
dz.addEventListener('dragover',  function(e) { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', function() { dz.classList.remove('drag-over'); });
dz.addEventListener('drop',      function(e) {
  e.preventDefault(); dz.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) handleFile(f);
});
document.getElementById('reassign-modal').addEventListener('click', function(e) {
  if (e.target === e.currentTarget) closeReassignModal();
});

/* Boot up the upload-screen background animation */
window.addEventListener('load', () => setTimeout(initUploadNetworkBG, 100));
</script>
</body>
</html>'''

components.html(APP_HTML, height=900, scrolling=False)
