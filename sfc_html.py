"""
SFC INTERACTIVE HTML EXPORT.

Produces a single self-contained .html file containing, per SFC block:
  - the SFC drawn as inline SVG (each step / transition is a real element)
  - hover tooltips with full action / condition detail
  - click a step -> expand its actions in a side panel
  - zoom / pan the diagram
  - a searchable, filterable actions table below the diagram

Reuses the same layout logic as the matplotlib renderer (grid snap, spine
placement, explicit FHX connections) so the picture matches the Excel/PNG.
"""

import html
import json

COL_W = 210
ROW_H = 70
SW, SH = 96, 36
TBAR = 18


def _snap_to_grid(steps, trans):
    nodes = [(n, d['x'], d['y']) for n, d in steps] + \
            [(n, d['x'], d['y']) for n, d in trans.items()]
    xs = sorted(set(x for _, x, _ in nodes))
    ys = sorted(set(y for _, _, y in nodes))

    def cluster(vals, tol):
        clusters = []
        for v in vals:
            placed = False
            for c in clusters:
                if abs(v - c[0]) <= tol:
                    c.append(v); placed = True; break
            if not placed:
                clusters.append([v])
        idx = {}
        for i, c in enumerate(sorted(clusters, key=lambda g: g[0])):
            for v in c:
                idx[v] = i
        return idx

    col_idx = cluster(xs, 90)
    row_idx = cluster(ys, 30)
    pos = {n: (col_idx[x], row_idx[y]) for n, x, y in nodes}
    return pos, max(col_idx.values()) + 1, max(row_idx.values()) + 1


def _geo_t2s(steps, trans):
    t2s = {}
    sp = {n: (d['x'], d['y']) for n, d in steps}
    for tname, td in trans.items():
        tx, ty = td['x'], td['y']
        best, bd = None, 1e18
        for sn, (sx, sy) in sp.items():
            if sy > ty:
                dd = (sy - ty) + abs(sx - tx) * 0.5
                if dd < bd:
                    bd, best = dd, sn
        if best:
            t2s[tname] = best
    return t2s


def _orth_path(x1, y1, x2, y2, loopback=False):
    """Return an SVG polyline points string for an orthogonal connector."""
    if loopback:
        side = x1 + 90 if x2 >= x1 else x1 - 90
        return f"{x1},{y1} {side},{y1} {side},{y2} {x2},{y2}"
    if abs(x1 - x2) < 1 or abs(y1 - y2) < 1:
        return f"{x1},{y1} {x2},{y2}"
    midy = (y1 + y2) / 2
    return f"{x1},{y1} {x1},{midy} {x2},{midy} {x2},{y2}"


def _layout(data):
    steps = data['ordered_steps']
    trans = data['transitions']
    s2t   = data.get('step_to_trans', {})
    if not steps:
        return None

    t2s_exp = data.get('trans_to_step', {})
    t2s = {}
    for tn in trans:
        if t2s_exp.get(tn):
            t2s[tn] = list(t2s_exp[tn])
        else:
            g = _geo_t2s(steps, {tn: trans[tn]})
            if g.get(tn):
                t2s[tn] = [g[tn]]

    pos, ncols, nrows = _snap_to_grid(steps, trans)

    def cx(c): return c * COL_W + COL_W / 2
    def cy(r): return (r + 1) * ROW_H        # top-down (SVG y grows down)

    row_of = {n: pos[n][1] for n in pos}

    # primary forward transition onto spine column
    col_override = {}
    for sn, tlist in s2t.items():
        if sn not in pos:
            continue
        s_col, s_row = pos[sn]
        fwd = []
        for tn in tlist:
            if tn not in pos:
                continue
            srows = [row_of[s] for s in t2s.get(tn, []) if s in pos]
            if not (srows and min(srows) <= s_row):
                fwd.append(tn)
        fwd.sort(key=lambda t: pos[t][1])
        if fwd:
            col_override[fwd[0]] = s_col
    for tn, col in col_override.items():
        r = pos[tn][1]; pos[tn] = (col, r)

    def center(n):
        c, r = pos[n]; return cx(c), cy(r)

    span_x = ncols * COL_W + COL_W
    span_y = (nrows + 2) * ROW_H

    return dict(steps=steps, trans=trans, s2t=s2t, t2s=t2s, pos=pos,
                row_of=row_of, center=center, span_x=span_x, span_y=span_y,
                init=steps[0][0])


def _svg_for_block(L):
    """Build the inline SVG string for one block layout."""
    steps, trans = L['steps'], L['trans']
    s2t, t2s = L['s2t'], L['t2s']
    pos, row_of, center = L['pos'], L['row_of'], L['center']
    init = L['init']
    parts = [f'<svg class="sfc" viewBox="0 0 {L["span_x"]:.0f} {L["span_y"]:.0f}" '
             f'preserveAspectRatio="xMidYMin meet">']

    # connectors
    def succ_rows(tn):
        return [row_of[s] for s in t2s.get(tn, []) if s in pos]

    for sn, tlist in s2t.items():
        if sn not in pos:
            continue
        scx, scy = center(sn)
        s_row = row_of[sn]
        fwd, loop = [], []
        for tn in tlist:
            if tn not in pos:
                continue
            sr = succ_rows(tn)
            (loop if (sr and min(sr) <= s_row) else fwd).append(tn)
        fwd.sort(key=lambda t: pos[t][1])
        for i, tn in enumerate(fwd):
            tcx, tcy = center(tn)
            x1 = scx; y1 = scy + SH/2
            if i == 0:
                pts = _orth_path(scx, y1, scx, tcy - TBAR/2)
            else:
                pts = _orth_path(scx, y1, tcx, tcy - TBAR/2)
            parts.append(f'<polyline class="lnk" points="{pts}"/>')
        for tn in loop:
            tcx, tcy = center(tn)
            pts = _orth_path(scx, scy + SH/2, tcx, tcy - TBAR/2)
            parts.append(f'<polyline class="lnk" points="{pts}"/>')

    for tn, snlist in t2s.items():
        if tn not in pos:
            continue
        tcx, tcy = center(tn)
        for sn in snlist:
            if sn not in pos:
                continue
            scx, scy = center(sn)
            if row_of[sn] <= row_of[tn]:
                pts = _orth_path(tcx, tcy + TBAR/2, scx, scy - SH/2, loopback=True)
                parts.append(f'<polyline class="lnk loop" points="{pts}"/>')
            else:
                pts = _orth_path(tcx, tcy + TBAR/2, scx, scy - SH/2)
                parts.append(f'<polyline class="lnk" points="{pts}"/>')

    # step boxes (interactive)
    for idx, (sn, sd) in enumerate(steps):
        ccx, ccy = center(sn)
        w = max(SW, len(sn) * 8 + 16)
        x, y = ccx - w/2, ccy - SH/2
        cls = "step init" if sn == init else "step"
        desc = (sd.get('description', '') or '')
        parts.append(
            f'<g class="{cls}" data-step="{html.escape(sn, quote=True)}" '
            f'tabindex="0">'
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{SH}" rx="3"/>'
            f'<text x="{ccx:.1f}" y="{ccy-2:.1f}" class="sid">{html.escape(sn)}</text>'
            f'<text x="{ccx:.1f}" y="{ccy+9:.1f}" class="sdesc">{html.escape(desc[:26])}</text>'
            f'</g>'
        )

    # transition bars + ids (interactive)
    for tn, td in trans.items():
        if tn not in pos:
            continue
        ccx, ccy = center(tn)
        is_end = td.get('termination', 'F') == 'T'
        cls = "trans end" if is_end else "trans"
        parts.append(
            f'<g class="{cls}" data-trans="{html.escape(tn, quote=True)}" tabindex="0">'
            f'<line x1="{ccx-10:.1f}" y1="{ccy:.1f}" x2="{ccx+10:.1f}" y2="{ccy:.1f}"/>'
            f'<text x="{ccx+15:.1f}" y="{ccy+3:.1f}" class="tid">{html.escape(tn)}</text>'
            f'</g>'
        )

    parts.append('</svg>')
    return "\n".join(parts)


def _block_data_json(L):
    """Per-step and per-transition detail used by the JS panel/table/tooltips."""
    steps_d = {}
    for sn, sd in L['steps']:
        acts = []
        for a in sd.get('actions', []):
            acts.append({
                'id': a.get('action') or a.get('action_id') or a.get('name') or '',
                'desc': a.get('description', ''),
                'qual': a.get('qualifier', ''),
                'expr': a.get('expression', ''),
                'delay': a.get('delay_time', '') or a.get('delay_expression', ''),
                'confirm': a.get('confirm_expression', ''),
            })
        steps_d[sn] = {'desc': sd.get('description', ''), 'actions': acts}
    trans_d = {}
    for tn, td in L['trans'].items():
        trans_d[tn] = {
            'desc': td.get('description', ''),
            'expr': td.get('expression', ''),
            'term': td.get('termination', 'F') == 'T',
        }
    return {'steps': steps_d, 'trans': trans_d}


_CSS = """
:root{--bd:#cbd2da;--ink:#0f172a;--mut:#64748b;--accent:#2563eb;--bg:#f8fafc;}
*{box-sizing:border-box}
body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:var(--bg);}
header{padding:14px 20px;background:#fff;border-bottom:1px solid var(--bd);position:sticky;top:0;z-index:5;}
header h1{margin:0;font-size:16px}
header .sub{color:var(--mut);font-size:12px;margin-top:2px}
.tabs{display:flex;gap:6px;padding:8px 20px;background:#fff;border-bottom:1px solid var(--bd);flex-wrap:wrap}
.tab{padding:5px 12px;border:1px solid var(--bd);border-radius:6px;background:#fff;cursor:pointer;font-size:12px}
.tab.on{background:var(--accent);color:#fff;border-color:var(--accent)}
.main{display:flex;flex-direction:column;height:calc(100vh - 96px)}
.wrap{display:flex;gap:0;flex:1 1 auto;min-height:120px;overflow:hidden}
.diagram{flex:1 1 auto;overflow:auto;position:relative;background:#fff;min-width:160px}
.vsplit{flex:0 0 6px;cursor:col-resize;background:var(--bd)}
.vsplit:hover,.vsplit.drag{background:var(--accent)}
.hsplit{flex:0 0 6px;cursor:row-resize;background:var(--bd)}
.hsplit:hover,.hsplit.drag{background:var(--accent)}
.controls{position:absolute;top:10px;right:10px;display:flex;gap:6px;z-index:3}
.controls button{width:30px;height:30px;border:1px solid var(--bd);background:#fff;border-radius:6px;cursor:pointer;font-size:15px}
.panel{flex:0 0 340px;overflow:auto;background:#fff;padding:14px;min-width:160px;border-left:1px solid var(--bd)}
.panel h3{margin:0 0 8px;font-size:14px}
.panel .empty{color:var(--mut);font-size:13px}
.act{border:1px solid var(--bd);border-radius:6px;padding:8px 10px;margin-bottom:8px}
.act .h{font-weight:600}
.act .q{display:inline-block;font-size:11px;color:#fff;background:var(--accent);border-radius:4px;padding:0 5px;margin-left:6px}
.act .expr{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12px;white-space:pre-wrap;background:#f1f5f9;border-radius:4px;padding:5px;margin-top:5px}
.act .meta{font-size:11px;color:var(--mut);margin-top:4px}
svg.sfc{min-width:100%;transform-origin:0 0}
.step rect{fill:#fff;stroke:#000;stroke-width:1}
.step.init rect{stroke-width:1;}
.step.init{filter:drop-shadow(0 0 0 #000)}
.step .sid{font-size:8px;font-weight:700;text-anchor:middle}
.step .sdesc{font-size:6.5px;text-anchor:middle;fill:#333}
.step{cursor:pointer}
.step:hover rect,.step:focus rect{fill:#dbeafe;stroke:var(--accent)}
.step.sel rect{fill:#bfdbfe;stroke:var(--accent);stroke-width:2}
.trans line{stroke:#000;stroke-width:2}
.trans.end line{stroke:#9a3b1c;stroke-width:3}
.trans .tid{font-size:7px;font-weight:700}
.trans{cursor:pointer}
.trans:hover .tid,.trans:focus .tid{fill:var(--accent)}
.lnk{fill:none;stroke:#000;stroke-width:.8}
.lnk.loop{stroke:#475569}
#tip{position:fixed;pointer-events:none;z-index:20;max-width:360px;background:#0f172a;color:#fff;border-radius:6px;padding:8px 10px;font-size:12px;display:none;box-shadow:0 4px 14px rgba(0,0,0,.25)}
#tip .tt{font-weight:600;margin-bottom:3px}
#tip .ln{white-space:pre-wrap;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:11px}
.tab.special{margin-left:4px;background:#fff7ed;border-color:#fdba74}
.tab.special.on{background:#ea580c;border-color:#ea580c;color:#fff}
.special{padding:16px 20px;overflow:auto}
.special.hidden{display:none}
.special h2{font-size:15px;margin:0 0 10px}
.special input{width:300px;max-width:70%;padding:7px 10px;border:1px solid var(--bd);border-radius:6px;font-size:13px;margin-bottom:10px}
.special table{border-collapse:collapse;width:100%;font-size:12px}
.special th,.special td{border:1px solid var(--bd);padding:5px 8px;text-align:left;vertical-align:top}
.special th{background:#eff6ff;position:sticky;top:0}
.special tr.grp td{background:#f1f5f9;font-weight:600}
.special td.expr{font-family:ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap}
.kind{display:inline-block;font-size:10px;color:#fff;border-radius:4px;padding:1px 6px}
.kind.hold{background:#b45309}.kind.sentinel{background:#7c3aed}.kind.failure{background:#b91c1c}
.special-mode .vsplit,.special-mode .panel{display:none}
.tablewrap{padding:10px 20px 14px;background:#fff;border-top:1px solid var(--bd);flex:0 0 280px;overflow:auto;min-height:90px}
.tablewrap input{width:280px;max-width:60%;padding:7px 10px;border:1px solid var(--bd);border-radius:6px;font-size:13px;margin-bottom:10px}
table{border-collapse:collapse;width:100%;font-size:12px}
th,td{border:1px solid var(--bd);padding:5px 8px;text-align:left;vertical-align:top}
th{background:#eff6ff;position:sticky;top:0;cursor:pointer}
td.expr{font-family:ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap}
tr.step-row td{background:#f1f5f9;font-weight:600}
.hidden{display:none}
"""

_JS = """
const DATA = __DATA__;
const EXTRA = __EXTRA__;
let curBlock = Object.keys(DATA)[0];
let zoom = 1;

function esc(s){return (s||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));}

function showSpecial(which){
  document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('on',t.dataset.k===which));
  document.querySelectorAll('.block').forEach(b=>b.classList.add('hidden'));
  const dia=document.querySelector('.diagram'), panel=document.getElementById('panel'), sp=document.getElementById('special');
  sp.classList.remove('hidden');
  document.querySelector('.wrap').classList.add('special-mode');
  if(which==='__params__'){
    const byGroup={};
    EXTRA.params.forEach(p=>{(byGroup[p.group||'(ungrouped)']=byGroup[p.group||'(ungrouped)']||[]).push(p);});
    let h='<h2>Phase Parameters ('+EXTRA.params.length+')</h2>';
    h+='<input id="pq" placeholder="Filter parameters…" oninput="filterSpecial(\\'param\\')">';
    h+='<table id="ptable"><thead><tr><th>Name</th><th>ID</th><th>Group</th><th>Description</th></tr></thead><tbody>';
    Object.keys(byGroup).sort().forEach(g=>{
      h+='<tr class="grp"><td colspan="4">'+esc(g)+'</td></tr>';
      byGroup[g].forEach(p=>{h+='<tr class="prow"><td>'+esc(p.name)+'</td><td>'+esc(p.id)+'</td><td>'+esc(p.group)+'</td><td>'+esc(p.desc)+'</td></tr>';});
    });
    h+='</tbody></table>';
    sp.innerHTML=h;
  } else if(which==='__monitors__'){
    const byKind={Hold:[],Sentinel:[],Failure:[]};
    EXTRA.monitors.forEach(m=>{(byKind[m.kind]=byKind[m.kind]||[]).push(m);});
    let h='<h2>Monitor Conditions ('+EXTRA.monitors.length+')</h2>';
    h+='<input id="mq" placeholder="Filter conditions…" oninput="filterSpecial(\\'mon\\')">';
    h+='<table id="mtable"><thead><tr><th>Type</th><th>Name</th><th>Condition</th></tr></thead><tbody>';
    ['Hold','Sentinel','Failure'].forEach(k=>{
      const items=byKind[k]||[]; if(!items.length)return;
      h+='<tr class="grp"><td colspan="3">'+k+' Monitor ('+items.length+')</td></tr>';
      items.forEach(m=>{h+='<tr class="mrow"><td><span class="kind '+k.toLowerCase()+'">'+k+'</span></td><td>'+esc(m.name)+'</td><td class="expr">'+esc(m.condition)+'</td></tr>';});
    });
    h+='</tbody></table>';
    sp.innerHTML=h;
  }
}

function filterSpecial(kind){
  const q=(document.getElementById(kind==='param'?'pq':'mq').value||'').toLowerCase();
  const sel=kind==='param'?'#ptable tr.prow':'#mtable tr.mrow';
  document.querySelectorAll(sel).forEach(tr=>{
    tr.classList.toggle('hidden', q && !tr.textContent.toLowerCase().includes(q));
  });
}

function showBlock(key){
  curBlock = key;
  const sp=document.getElementById('special'); if(sp){sp.classList.add('hidden');}
  document.querySelector('.wrap').classList.remove('special-mode');
  document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('on',t.dataset.k===key));
  document.querySelectorAll('.block').forEach(b=>b.classList.toggle('hidden',b.dataset.k!==key));
  clearPanel();
  buildTable();
  zoom=1; applyZoom();
}

function clearPanel(){
  document.getElementById('panel').innerHTML='<h3>Step actions</h3><div class="empty">Click a step in the diagram to see its full actions here.</div>';
  document.querySelectorAll('.step.sel').forEach(s=>s.classList.remove('sel'));
}

function selectStep(sn){
  const blk = DATA[curBlock];
  const sd = blk.steps[sn]; if(!sd) return;
  document.querySelectorAll('.block:not(.hidden) .step.sel').forEach(s=>s.classList.remove('sel'));
  const g = document.querySelector('.block:not(.hidden) .step[data-step="'+CSS.escape(sn)+'"]');
  if(g) g.classList.add('sel');
  let h = '<h3>'+esc(sn)+(sd.desc?' — '+esc(sd.desc):'')+'</h3>';
  if(!sd.actions.length){ h+='<div class="empty">No actions.</div>'; }
  sd.actions.forEach(a=>{
    h+='<div class="act"><div class="h">'+esc(a.id)+(a.qual?'<span class="q">'+esc(a.qual)+'</span>':'')+'</div>';
    if(a.desc) h+='<div>'+esc(a.desc)+'</div>';
    if(a.expr) h+='<div class="expr">'+esc(a.expr)+'</div>';
    let meta=[]; if(a.delay && a.delay!=='0') meta.push('delay: '+esc(a.delay));
    if(a.confirm) meta.push('confirm: '+esc(a.confirm));
    if(meta.length) h+='<div class="meta">'+meta.join('  ·  ')+'</div>';
    h+='</div>';
  });
  document.getElementById('panel').innerHTML=h;
}

// tooltips
const tip=document.getElementById('tip');
function tipShow(html,x,y){tip.innerHTML=html;tip.style.display='block';tip.style.left=(x+14)+'px';tip.style.top=(y+14)+'px';}
function tipHide(){tip.style.display='none';}

function wireBlock(blockEl){
  blockEl.querySelectorAll('.step').forEach(g=>{
    const sn=g.dataset.step;
    g.addEventListener('click',()=>selectStep(sn));
    g.addEventListener('keypress',e=>{if(e.key==='Enter')selectStep(sn);});
    g.addEventListener('mousemove',e=>{
      const sd=DATA[curBlock].steps[sn]; if(!sd)return;
      let h='<div class="tt">'+esc(sn)+(sd.desc?' — '+esc(sd.desc):'')+'</div>';
      h+='<div class="ln">'+(sd.actions.length?sd.actions.map(a=>'• '+esc(a.id)+' ['+esc(a.qual)+'] '+esc(a.desc)).join('\\n'):'(no actions)')+'</div>';
      tipShow(h,e.clientX,e.clientY);
    });
    g.addEventListener('mouseleave',tipHide);
  });
  blockEl.querySelectorAll('.trans').forEach(g=>{
    const tn=g.dataset.trans;
    g.addEventListener('mousemove',e=>{
      const td=DATA[curBlock].trans[tn]; if(!td)return;
      let h='<div class="tt">'+esc(tn)+(td.term?' (terminating)':'')+(td.desc?' — '+esc(td.desc):'')+'</div>';
      if(td.expr) h+='<div class="ln">'+esc(td.expr)+'</div>';
      tipShow(h,e.clientX,e.clientY);
    });
    g.addEventListener('mouseleave',tipHide);
  });
}

function buildTable(){
  const blk=DATA[curBlock];
  let rows='';
  Object.keys(blk.steps).forEach(sn=>{
    const sd=blk.steps[sn];
    rows+='<tr class="step-row"><td colspan="6">'+esc(sn)+(sd.desc?' — '+esc(sd.desc):'')+' ('+sd.actions.length+' actions)</td></tr>';
    sd.actions.forEach(a=>{
      rows+='<tr class="arow"><td>'+esc(sn)+'</td><td>'+esc(a.id)+'</td><td>'+esc(a.desc)+'</td><td>'+esc(a.qual)+'</td><td class="expr">'+esc(a.expr)+'</td><td>'+[a.delay&&a.delay!=='0'?'delay '+esc(a.delay):'',a.confirm?'confirm '+esc(a.confirm):''].filter(Boolean).join(' / ')+'</td></tr>';
    });
  });
  document.getElementById('tbody').innerHTML=rows;
  document.getElementById('search').value='';
}

function filterTable(){
  const q=document.getElementById('search').value.toLowerCase();
  document.querySelectorAll('#tbody tr.arow').forEach(tr=>{
    tr.classList.toggle('hidden', q && !tr.textContent.toLowerCase().includes(q));
  });
  // hide step header rows whose actions are all hidden
  document.querySelectorAll('#tbody tr.step-row').forEach(h=>{
    let n=h.nextElementSibling, any=false;
    while(n && n.classList.contains('arow')){ if(!n.classList.contains('hidden')) any=true; n=n.nextElementSibling; }
    h.classList.toggle('hidden', q && !any);
  });
}

function applyZoom(){document.querySelectorAll('.block:not(.hidden) svg.sfc').forEach(s=>s.style.transform='scale('+zoom+')');}
function zoomIn(){zoom=Math.min(4,zoom*1.2);applyZoom();}
function zoomOut(){zoom=Math.max(.3,zoom/1.2);applyZoom();}
function zoomReset(){zoom=1;applyZoom();}

function initSplitters(){
  // vertical splitter: resize side panel width
  const vs=document.getElementById('vsplit'), panel=document.getElementById('panel'), wrap=document.querySelector('.wrap');
  let vdrag=false;
  vs.addEventListener('mousedown',e=>{vdrag=true;vs.classList.add('drag');e.preventDefault();document.body.style.userSelect='none';});
  window.addEventListener('mousemove',e=>{
    if(!vdrag)return;
    const r=wrap.getBoundingClientRect();
    let w=r.right-e.clientX;
    w=Math.max(160,Math.min(r.width-200,w));
    panel.style.flex='0 0 '+w+'px';
  });
  window.addEventListener('mouseup',()=>{if(vdrag){vdrag=false;vs.classList.remove('drag');document.body.style.userSelect='';}});

  // horizontal splitter: resize bottom table height
  const hs=document.getElementById('hsplit'), tw=document.getElementById('tablewrap'), main=document.querySelector('.main');
  let hdrag=false;
  hs.addEventListener('mousedown',e=>{hdrag=true;hs.classList.add('drag');e.preventDefault();document.body.style.userSelect='none';});
  window.addEventListener('mousemove',e=>{
    if(!hdrag)return;
    const r=main.getBoundingClientRect();
    let h=r.bottom-e.clientY;
    h=Math.max(90,Math.min(r.height-160,h));
    tw.style.flex='0 0 '+h+'px';
  });
  window.addEventListener('mouseup',()=>{if(hdrag){hdrag=false;hs.classList.remove('drag');document.body.style.userSelect='';}});
}

window.addEventListener('DOMContentLoaded',()=>{
  document.querySelectorAll('.block').forEach(wireBlock);
  showBlock(curBlock);
  initSplitters();
});
"""


def build_sfc_html(blocks, fname, opts=None):
    """Return a single self-contained HTML string for all blocks in the phase."""
    opts = opts or {}
    params  = blocks.get('__parameters__', [])
    monitors = blocks.get('__monitors__', [])

    block_layouts = {}
    friendly = {}     # key -> display name
    for name, data in blocks.items():
        if name in ('__parameters__', '__monitors__'):
            continue
        L = _layout(data)
        if L:
            block_layouts[name] = L
            iname = data.get('instance_name', '') or name
            friendly[name] = iname
    if not block_layouts:
        return "<html><body><p>No SFC logic found.</p></body></html>"

    # tabs + block sections
    tabs, sections, data_json = [], [], {}
    for i, (name, L) in enumerate(block_layouts.items()):
        label = friendly.get(name, name)
        tabs.append(f'<button class="tab{" on" if i==0 else ""}" '
                    f'data-k="{html.escape(name, quote=True)}" '
                    f'onclick="showBlock(\'{html.escape(name)}\')">{html.escape(label)}</button>')
        svg = _svg_for_block(L)
        sections.append(
            f'<div class="block{" hidden" if i>0 else ""}" '
            f'data-k="{html.escape(name, quote=True)}">{svg}</div>'
        )
        data_json[name] = _block_data_json(L)

    # special tabs: Parameters, Monitors
    tabs.append('<button class="tab special" data-k="__params__" '
                'onclick="showSpecial(\'__params__\')">⚙ Parameters</button>')
    tabs.append('<button class="tab special" data-k="__monitors__" '
                'onclick="showSpecial(\'__monitors__\')">⚠ Monitors</button>')

    extra = {'params': params, 'monitors': monitors}
    js = _JS.replace('__DATA__', json.dumps(data_json)).replace('__EXTRA__', json.dumps(extra))

    htmldoc = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SFC — {html.escape(fname)}</title>
<style>{_CSS}</style></head>
<body>
<header><h1>SFC Diagram — {html.escape(fname)}</h1>
<div class="sub">Hover a step or transition for detail · click a step for full actions · zoom/pan the diagram · search the table below</div></header>
<div class="tabs">{''.join(tabs)}</div>
<div class="main">
<div class="wrap">
  <div class="diagram">
    <div class="controls">
      <button onclick="zoomIn()" title="Zoom in">+</button>
      <button onclick="zoomOut()" title="Zoom out">−</button>
      <button onclick="zoomReset()" title="Reset">⤢</button>
    </div>
    {''.join(sections)}
    <div id="special" class="special hidden"></div>
  </div>
  <div class="vsplit" id="vsplit" title="Drag to resize"></div>
  <div class="panel" id="panel"></div>
</div>
<div class="hsplit" id="hsplit" title="Drag to resize"></div>
<div id="tip"></div>
<div class="tablewrap" id="tablewrap">
  <input id="search" type="text" placeholder="Search actions, expressions, steps…" oninput="filterTable()">
  <div style="overflow:auto;border:1px solid var(--bd);border-radius:6px">
  <table><thead><tr><th>Step</th><th>Action</th><th>Description</th><th>Qual</th><th>Expression</th><th>Delay / Confirm</th></tr></thead>
  <tbody id="tbody"></tbody></table>
  </div>
</div>
</div>
<script>{js}</script>
</body></html>"""
    return htmldoc
