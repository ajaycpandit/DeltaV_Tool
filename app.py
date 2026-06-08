from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io, re, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app)

# ── FHX Decoder ───────────────────────────────────────────────────────────────
def decode_fhx(raw_bytes):
    if raw_bytes[:2] == b'\xff\xfe':
        return raw_bytes.decode('utf-16-le', errors='replace').lstrip('\ufeff')
    if raw_bytes[:2] == b'\xfe\xff':
        return raw_bytes.decode('utf-16-be', errors='replace').lstrip('\ufeff')
    return raw_bytes.decode('utf-8', errors='replace')

# ── Block extractor ───────────────────────────────────────────────────────────
def extract_block(text, start):
    depth, i = 0, start
    while i < len(text):
        if text[i] == '{': depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0: return text[start:i+1]
        i += 1
    return ''

# ── Build FB-definition → instance name map ───────────────────────────────────
# Scans for patterns like:
#   FUNCTION_BLOCK NAME="RUN_LOGIC" DEFINITION="__63699255_143E0A8C__"
# to map the GUID definition name → the human-readable instance name.
def build_instance_map(text):
    """Returns {definition_name: instance_name}"""
    mapping = {}
    for m in re.finditer(
        r'FUNCTION_BLOCK\s+NAME="([^"]+)"\s+DEFINITION="([^"]+)"',
        text
    ):
        inst_name = m.group(1)   # e.g. RUN_LOGIC, HOLD_LOGIC
        defn_name = m.group(2)   # e.g. __63699255_143E0A8C__
        # Keep first occurrence (most specific context)
        if defn_name not in mapping:
            mapping[defn_name] = inst_name
    return mapping

# ── Action parser ─────────────────────────────────────────────────────────────
def parse_actions(sc):
    actions = []
    for m in re.finditer(r'ACTION NAME="([^"]+)"\s*\{', sc):
        b = extract_block(sc, m.end()-1)
        def g(p): r = re.search(p, b); return r.group(1) if r else ''
        expr_m = re.search(r'EXPRESSION="(.*?)"(?=\s+(?:DELAY|CONFIRM)|\s*\})', b, re.DOTALL)
        actions.append({
            'action':             m.group(1),
            'description':        g('DESCRIPTION="([^"]+)"'),
            'type':               g('ACTION_TYPE=(\\S+)'),
            'qualifier':          g('QUALIFIER=(\\S+)'),
            'expression':         expr_m.group(1).strip() if expr_m else '',
            'delay_time':         g('DELAY_TIME=(\\S+)'),
            'delay_expression':   g('DELAY_EXPRESSION="([^"]+)"'),
            'confirm_expression': g('CONFIRM_EXPRESSION="([^"]+)"'),
            'confirm_timeout':    g('CONFIRM_TIME_OUT=(\\S+)'),
        })
    return actions

# ── FHX parser ────────────────────────────────────────────────────────────────
def parse_fhx(text):
    # First pass: map every GUID definition → its instance name
    instance_map = build_instance_map(text)

    blocks = {}
    for m in re.finditer(r'FUNCTION_BLOCK_DEFINITION NAME="([^"]+)"[^\{]*\{', text):
        name  = m.group(1)
        block = extract_block(text, m.end()-1)
        desc  = re.search(r'DESCRIPTION="([^"]+)"', block)
        sfc_m = re.search(r'SFC_ALGORITHM\s*\{', block)
        sfc   = extract_block(block, sfc_m.end()-1) if sfc_m else ''
        steps, trans = {}, {}

        for sm in re.finditer(r'STEP NAME="([^"]+)"\s*\{', sfc):
            sb = extract_block(sfc, sm.end()-1)
            rc = re.search(r'RECTANGLE\s*=\s*\{\s*X=(\d+)\s*Y=(\d+)', sb)
            steps[sm.group(1)] = {
                'x': int(rc.group(1)) if rc else 0,
                'y': int(rc.group(2)) if rc else 0,
                'actions': parse_actions(sb)
            }

        for tm in re.finditer(r'TRANSITION NAME="([^"]+)"\s*\{', sfc):
            tb  = extract_block(sfc, tm.end()-1)
            pos = re.search(r'POSITION\s*=\s*\{\s*X=(\d+)\s*Y=(\d+)', tb)
            td  = re.search(r'DESCRIPTION="([^"]+)"', tb)
            te  = re.search(r'EXPRESSION="([^"]*)"', tb, re.DOTALL)
            tt  = re.search(r'TERMINATION=(\w+)', tb)
            trans[tm.group(1)] = {
                'x':           int(pos.group(1)) if pos else 0,
                'y':           int(pos.group(2)) if pos else 0,
                'description': td.group(1).strip() if td else '',
                'expression':  te.group(1).strip() if te else '',
                'termination': tt.group(1) if tt else 'F'
            }

        s2t = {s: [] for s in steps}
        for tn, tp in trans.items():
            cands = [(tp['y']-v['y'], sn) for sn,v in steps.items()
                     if v['y'] < tp['y'] and abs(v['x']-tp['x']) < 150]
            if not cands:
                cands = [(tp['y']-v['y'], sn) for sn,v in steps.items() if v['y'] < tp['y']]
            if cands:
                s2t[sorted(cands)[0][1]].append(tn)

        blocks[name] = {
            'description':   desc.group(1) if desc else '',
            'instance_name': instance_map.get(name, ''),   # e.g. RUN_LOGIC, HOLD_LOGIC
            'ordered_steps': sorted(steps.items(), key=lambda x: x[1]['y']),
            'transitions':   trans,
            'step_to_trans': s2t
        }
    return blocks

# ── Sheet label derivation ─────────────────────────────────────────────────────
def clean_label(raw):
    """Convert an instance name or description into a clean sheet tab label."""
    s = raw.strip().upper()
    # Replace separators with space
    s = re.sub(r'[\-_/\\]+', ' ', s)
    # Remove noise suffixes
    for noise in [' LOGIC', ' PHASE', ' SEQUENCE', ' SUB', ' FUNCTION']:
        s = s.replace(noise, '')
    # Take first 28 chars, convert spaces back to underscore
    s = re.sub(r'\s+', '_', s.strip())[:28]
    # Strip invalid Excel sheet name chars
    s = re.sub(r'[\/\\\?\*\[\]\:\']', '_', s)
    return s or raw[:28]

def derive_label(fb_name, instance_name, description, used_labels):
    """
    Priority:
    1. Instance name (e.g. RUN_LOGIC, HOLD_LOGIC) — most reliable
    2. Description keyword match
    3. Last meaningful words of description
    4. FB name itself (fallback)
    """
    # 1 — instance name is the gold standard
    if instance_name and instance_name.upper() not in ('', 'NONE'):
        candidate = clean_label(instance_name)

    else:
        # 2 — keyword match in description
        desc_up = description.upper()
        _KEYWORDS = [
            ('ABORT',    'ABORT_STOP'), ('STOP',     'ABORT_STOP'),
            ('RUNNING',  'RUNNING'),    ('RUN ',     'RUNNING'),
            ('HOLDING',  'HOLD'),       ('HOLD',     'HOLD'),
            ('RESTART',  'RESTART'),    ('CONDITION','CONDITION'),
            ('PROMPT',   'PROMPT'),     ('INIT',     'INIT'),
            ('START',    'START'),      ('COMPLETE', 'COMPLETE'),
            ('FAIL',     'FAIL'),       ('PAUSE',    'PAUSE'),
            ('CHARGE',   'CHARGE'),     ('DRAIN',    'DRAIN'),
            ('HEAT',     'HEAT'),       ('COOL',     'COOL'),
            ('TRANSFER', 'TRANSFER'),   ('WASH',     'WASH'),
            ('RINSE',    'RINSE'),      ('CLEAN',    'CLEAN'),
            ('SAMPLE',   'SAMPLE'),     ('AGITATE',  'AGITATE'),
            ('WAIT',     'WAIT'),
        ]
        candidate = None
        for kw, lbl in _KEYWORDS:
            if kw in desc_up:
                candidate = lbl
                break

        if not candidate:
            # 3 — last 2 meaningful words of description
            words = re.sub(r'[^A-Z0-9\s]', ' ', desc_up).split()
            noise = {'LOGIC','PHASE','THE','FOR','AND','OR','A','AN','OF',
                     'FUNCTION','BLOCK','STEP','SUB','SEQUENCE'}
            words = [w for w in words if w not in noise and len(w) > 1]
            candidate = '_'.join(words[-2:]) if words else fb_name[:20]

    # Make unique
    base, n = candidate[:28], 2
    label = base
    while label in used_labels:
        label = f"{base[:25]}_{n}"
        n += 1
    used_labels.add(label)
    return label

# ── Excel styles ──────────────────────────────────────────────────────────────
NAVY   = PatternFill('solid', start_color='0D1B4B')
BLUE_H = PatternFill('solid', start_color='1F3864')
BLUE_S = PatternFill('solid', start_color='2E75B6')
GREEN_S= PatternFill('solid', start_color='4E7F2C')
ALT    = PatternFill('solid', start_color='D9E1F2')
ALT_G  = PatternFill('solid', start_color='E2EFDA')
ALTG2  = PatternFill('solid', start_color='F2F9EE')
WHITE  = PatternFill('solid', start_color='FFFFFF')
THIN   = Side(style='thin', color='AAAAAA')
BORD   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NCOLS  = 10
COL_W  = [14, 22, 32, 10, 10, 58, 10, 38, 34, 14]

def wf(bold=False, sz=10, color='000000'):
    return Font(name='Calibri', bold=bold, size=sz, color=color)

def wa(h='left', wrap=False):
    return Alignment(horizontal=h, vertical='top', wrap_text=wrap)

def sc(ws, r, c, val='', bold=False, sz=10, fc='000000', fill=None, h='left', wrap=False, merge_to=None):
    cell = ws.cell(row=r, column=c, value=str(val) if val is not None else '')
    cell.font      = wf(bold, sz, fc)
    cell.fill      = fill or WHITE
    cell.alignment = wa(h, wrap)
    cell.border    = BORD
    if merge_to:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=merge_to)
    return cell

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(blocks, fname, opts):
    wb = openpyxl.Workbook()
    used_labels = set()

    # Pre-compute labels so SUMMARY and logic sheets use identical names
    block_labels = {}
    for fb_name, data in blocks.items():
        block_labels[fb_name] = derive_label(
            fb_name,
            data.get('instance_name', ''),
            data.get('description', ''),
            used_labels
        )

    # ── SUMMARY ────────────────────────────────────────────────────────────────
    if opts.get('summary', True):
        ws = wb.active
        ws.title = 'SUMMARY'
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
        t = ws.cell(row=1, column=1, value=f"  {fname}  —  DeltaV Phase Export Logic Summary")
        t.font = wf(True, 13, 'FFFFFF'); t.fill = NAVY
        t.alignment = Alignment(horizontal='left', vertical='center'); t.border = BORD
        ws.row_dimensions[1].height = 28

        for ci, h in enumerate(['Logic Block','FB Name','Description','Steps',
                                 'Actions','Transitions','Orphan Trans.','Sheet Tab'], 1):
            sc(ws, 2, ci, h, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        ws.row_dimensions[2].height = 17

        for ci, w in enumerate([18,30,50,8,10,13,14,18], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A3'

        for i, (fb_name, data) in enumerate(blocks.items()):
            lbl    = block_labels[fb_name]
            acts   = sum(len(s[1]['actions']) for s in data['ordered_steps'])
            mapped = sum(len(v) for v in data['step_to_trans'].values())
            orphan = len(data['transitions']) - mapped
            f      = ALT if i % 2 == 0 else WHITE
            r      = i + 3
            sc(ws, r, 1, lbl,                         bold=True, fill=f)
            sc(ws, r, 2, fb_name,                     fill=f)
            sc(ws, r, 3, data['description'],          fill=f, wrap=True)
            sc(ws, r, 4, len(data['ordered_steps']),   fill=f, h='center')
            sc(ws, r, 5, acts,                         fill=f, h='center')
            sc(ws, r, 6, len(data['transitions']),     fill=f, h='center')
            sc(ws, r, 7, orphan or '-',                fill=f, h='center')
            sc(ws, r, 8, lbl,                          fill=f)
            ws.row_dimensions[r].height = 16
    else:
        wb.active.title = '_temp'

    # ── Logic sheets ────────────────────────────────────────────────────────────
    for fb_name, data in blocks.items():
        lbl = block_labels[fb_name]
        ws  = wb.create_sheet(title=lbl)

        for ci, w in enumerate(COL_W, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A3'

        # Title
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
        t = ws.cell(row=1, column=1,
                    value=f"  Phase Logic: {lbl}   |   {data['description']}")
        t.font = wf(True, 12, 'FFFFFF'); t.fill = NAVY
        t.alignment = Alignment(horizontal='left', vertical='center'); t.border = BORD
        ws.row_dimensions[1].height = 24

        # Column headers
        hdrs = ['Row Type','Step / Transition','Description','Action','Qualifier',
                'Expression' if opts.get('expressions', True) else '(hidden)',
                'Delay','Delay Expression','Confirm Expression','Confirm Timeout']
        for ci, h in enumerate(hdrs, 1):
            sc(ws, 2, ci, h, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        ws.row_dimensions[2].height = 17

        row = 3
        for si, (sn, sd) in enumerate(data['ordered_steps']):

            # Step header row
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
            sh = ws.cell(row=row, column=1,
                         value=f"  STEP {si+1}:  {sn}   ({len(sd['actions'])} action{'s' if len(sd['actions'])!=1 else ''})")
            sh.font = wf(True, 10, 'FFFFFF'); sh.fill = BLUE_S
            sh.alignment = Alignment(horizontal='left', vertical='center'); sh.border = BORD
            ws.row_dimensions[row].height = 17
            row += 1

            # Action rows
            for ai, a in enumerate(sd['actions']):
                f        = ALT if ai % 2 == 0 else WHITE
                inc_expr = opts.get('expressions', True)
                sc(ws, row, 1,  'ACTION',                              bold=True, fill=f, h='center')
                sc(ws, row, 2,  sn,                                    fill=f)
                sc(ws, row, 3,  a['description'],                      fill=f, wrap=True)
                sc(ws, row, 4,  a['action'],                           fill=f, h='center')
                sc(ws, row, 5,  a['qualifier'],                        fill=f, h='center')
                sc(ws, row, 6,  a['expression'] if inc_expr else '—',  fill=f, wrap=True)
                sc(ws, row, 7,  a['delay_time'],                       fill=f, h='center')
                sc(ws, row, 8,  a['delay_expression'],                 fill=f, wrap=True)
                sc(ws, row, 9,  a['confirm_expression'],               fill=f, wrap=True)
                sc(ws, row, 10, a['confirm_timeout'],                  fill=f, h='center')
                ws.row_dimensions[row].height = max(15, min(75,
                    15 * max(1, len(a.get('expression','')) // 55 + 1)))
                row += 1

            # Transitions embedded under this step
            if opts.get('transitions', True):
                tl = data['step_to_trans'].get(sn, [])
                if tl:
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
                    th = ws.cell(row=row, column=1,
                                 value=f"  ↓  TRANSITION{'S' if len(tl)>1 else ''} FROM  {sn}  ({len(tl)})")
                    th.font = wf(True, 9, 'FFFFFF'); th.fill = GREEN_S
                    th.alignment = Alignment(horizontal='left', vertical='center'); th.border = BORD
                    ws.row_dimensions[row].height = 15
                    row += 1

                    for ti, tn in enumerate(tl):
                        tr = data['transitions'][tn]
                        f  = ALT_G if ti % 2 == 0 else ALTG2
                        sc(ws, row, 1, 'TRANSITION', bold=True, fill=f, h='center')
                        sc(ws, row, 2, tn,           bold=True, fill=f)
                        sc(ws, row, 3, tr['description'], fill=f, wrap=True)
                        sc(ws, row, 4, '⏹ END' if tr['termination']=='T' else '→ NEXT',
                                       fill=f, h='center')
                        sc(ws, row, 5, '', fill=f)
                        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=NCOLS)
                        ec = ws.cell(row=row, column=6,
                                     value=tr['expression'] if opts.get('expressions', True) else '—')
                        ec.font = wf(); ec.fill = f
                        ec.alignment = wa('left', True); ec.border = BORD
                        ws.row_dimensions[row].height = max(15, min(60,
                            15 * max(1, len(tr.get('expression','')) // 80 + 1)))
                        row += 1

            # Spacer row
            for ci in range(1, NCOLS + 1):
                sc(ws, row, ci, '', fill=WHITE)
            ws.row_dimensions[row].height = 6
            row += 1

    if '_temp' in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb['_temp']

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return send_file('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.fhx'):
        return jsonify({'error': 'File must be a .fhx file'}), 400
    opts = {
        'summary':     request.form.get('summary',     'true') == 'true',
        'transitions': request.form.get('transitions', 'true') == 'true',
        'expressions': request.form.get('expressions', 'true') == 'true',
    }
    raw    = f.read()
    text   = decode_fhx(raw)
    blocks = parse_fhx(text)
    fname  = re.sub(r'\.fhx$', '', f.filename, flags=re.IGNORECASE)
    buf    = build_excel(blocks, fname, opts)
    return send_file(buf, as_attachment=True,
                     download_name=fname + '_Logic.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
