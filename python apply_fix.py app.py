#!/usr/bin/env python3
"""Apply the Streamlit Cloud export fix to app.py"""
import sys

def patch(src):
    patches_ok = 0

    # ── PATCH 1: buildRenderStage – fix position:fixed → position:absolute wrapper
    p1_old = "  // Step 4: Place clone in an on-screen stage behind the export overlay\n  const stage=document.createElement('div');\n  stage.style.cssText=`position:fixed;top:0;left:0;z-index:9998;background:#f8fafc;padding:${PAD}px;display:inline-block;white-space:nowrap;pointer-events:none;font-family:'Plus Jakarta Sans',sans-serif`;\n  stage.appendChild(cloned);\n  document.body.appendChild(stage);\n\n  // Step 5: Allow one full paint cycle + font load\n  await new Promise(r=>setTimeout(r,350));\n  if(document.fonts?.ready)await document.fonts.ready;\n  await new Promise(r=>setTimeout(r,120));\n\n  // Step 6: Freeze computed styles for node cards (IC summary cards are skipped — already inline-styled)\n  inlineStyles(stage);\n  await new Promise(r=>setTimeout(r,80));\n  return stage;\n}"
    p1_new = "  // FIX (Streamlit Cloud iframe): position:absolute wrapper, not position:fixed.\n  // position:fixed clips to the iframe viewport — html2canvas misses IC summary\n  // list content that overflows below the fold (renders as blank boxes in PNG/PPTX).\n  const wrapper=document.createElement('div');\n  wrapper.style.cssText='position:absolute;top:-99999px;left:0;width:100%;overflow:visible;pointer-events:none;z-index:9998';\n  const stage=document.createElement('div');\n  stage.style.cssText=`background:#f8fafc;padding:${PAD}px;display:inline-block;white-space:nowrap;font-family:'Plus Jakarta Sans',sans-serif;overflow:visible`;\n  stage.appendChild(cloned);\n  wrapper.appendChild(stage);\n  document.body.appendChild(wrapper);\n\n  await new Promise(r=>setTimeout(r,400));\n  if(document.fonts?.ready)await document.fonts.ready;\n  await new Promise(r=>setTimeout(r,150));\n\n  // FIX: force summary list cards to overflow:visible before inlineStyles freezes\n  // computed styles — otherwise the card body height clips IC rows during capture.\n  stage.querySelectorAll('.summary-list-card').forEach(card=>{\n    card.style.overflow='visible';card.style.maxHeight='none';\n  });\n  stage.querySelectorAll('.summary-list-card *').forEach(el=>{\n    el.style.overflow='visible';el.style.maxHeight='none';\n  });\n\n  inlineStyles(stage);\n  await new Promise(r=>setTimeout(r,100));\n  return {stage, wrapper};\n}"

    if p1_old in src:
        src = src.replace(p1_old, p1_new, 1); patches_ok += 1
        print("  [OK] Patch 1: buildRenderStage stage placement")
    else:
        print("  [!!] Patch 1 NOT FOUND — check whitespace in source")

    # fix PAD value
    src = src.replace(
        "async function buildRenderStage(){\n  const PAD=20;",
        "async function buildRenderStage(){\n  const PAD=40;", 1)

    # ── PATCH 2: renderToCanvas – accept object, drop windowWidth/Height
    p2_old = "async function renderToCanvas(stage){\n  const W=Math.ceil(stage.scrollWidth),H=Math.ceil(stage.scrollHeight);\n  return html2canvas(stage,{\n    backgroundColor:'#f8fafc',scale:3,useCORS:true,logging:false,\n    allowTaint:true,foreignObjectRendering:false,\n    width:W,height:H,scrollX:0,scrollY:0,\n    windowWidth:Math.max(W+200,window.innerWidth),\n    windowHeight:Math.max(H+200,window.innerHeight),\n    x:0,y:0,\n  });\n}"
    p2_new = "async function renderToCanvas(stageObj){\n  // FIX: accept {stage,wrapper} from buildRenderStage\n  const stage=stageObj.stage||stageObj;\n  const W=Math.ceil(stage.scrollWidth),H=Math.ceil(stage.scrollHeight);\n  // FIX: omit windowWidth/windowHeight — inside a Streamlit iframe these cause\n  // html2canvas to clip output to iframe viewport height, hiding IC summary rows.\n  return html2canvas(stage,{\n    backgroundColor:'#f8fafc',scale:3,useCORS:true,logging:false,\n    allowTaint:true,foreignObjectRendering:false,\n    width:W,height:H,scrollX:0,scrollY:0,x:0,y:0,\n  });\n}"

    if p2_old in src:
        src = src.replace(p2_old, p2_new, 1); patches_ok += 1
        print("  [OK] Patch 2: renderToCanvas")
    else:
        print("  [!!] Patch 2 NOT FOUND")

    # ── PATCH 3: all stage cleanup sites → wrapper-aware
    cleanup_old = "if(stage)stage.remove();"
    cleanup_new = "if(stage?.wrapper)stage.wrapper.remove();else if(stage?.stage)stage.stage.remove();else if(stage)stage.remove();"
    n = src.count(cleanup_old)
    if n:
        src = src.replace(cleanup_old, cleanup_new); patches_ok += 1
        print(f"  [OK] Patch 3: stage.remove() → wrapper-aware ({n} sites)")
    else:
        print("  [!!] Patch 3: no stage.remove() found")

    # stage2 in exportAll loop
    src = src.replace(
        "finally{if(stage2)stage2.remove();}",
        "finally{if(stage2?.wrapper)stage2.wrapper.remove();else if(stage2)stage2.remove();}"
    )

    return src, patches_ok

path = sys.argv[1] if len(sys.argv) > 1 else 'app.py'
print(f"Reading {path}…")
with open(path, 'r', encoding='utf-8') as f:
    src = f.read()
print(f"  {len(src):,} bytes read")

result, n = patch(src)

out = path.replace('.py', '_fixed.py')
with open(out, 'w', encoding='utf-8') as f:
    f.write(result)
print(f"\nSaved → {out}  ({n}/3 patch groups applied)")
print("Rename to app.py and redeploy to Streamlit Cloud.")
