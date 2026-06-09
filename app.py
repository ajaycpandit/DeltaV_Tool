from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io, re, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app, expose_headers=['X-Sheet-Names'])

# ═══════════════════════════════════════════════════════════════════════════════
# SHARED UTILITIES
# ═══════════════════════════════════════════════════════════════════════════════

def decode_fhx(raw):
    if raw[:2] == b'\xff\xfe': return raw.decode('utf-16-le', errors='replace').lstrip('\ufeff')
    if raw[:2] == b'\xfe\xff': return raw.decode('utf-16-be', errors='replace').lstrip('\ufeff')
    return raw.decode('utf-8', errors='replace')

def extract_block(text, start):
    depth, i = 0, start
    while i < len(text):
        if text[i] == '{': depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0: return text[start:i+1]
        i += 1
    return ''

def parse_actions(sc_text):
    actions = []
    for m in re.finditer(r'ACTION\s+NAME="([^"]+)"\s*\{', sc_text):
        b = extract_block(sc_text, m.end()-1)
        g = lambda p: (re.search(p, b) or type('x',(),{'group':lambda s,n:''})()).group(1)
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

def parse_sfc(sfc_block):
    steps, trans = {}, {}
    for sm in re.finditer(r'STEP\s+NAME="([^"]+)"\s*\{', sfc_block):
        sb = extract_block(sfc_block, sm.end()-1)
        rc = re.search(r'RECTANGLE\s*=\s*\{\s*X=(\d+)\s*Y=(\d+)', sb)
        steps[sm.group(1)] = {
            'x': int(rc.group(1)) if rc else 0,
            'y': int(rc.group(2)) if rc else 0,
            'actions': parse_actions(sb)
        }
    for tm in re.finditer(r'TRANSITION\s+NAME="([^"]+)"\s*\{', sfc_block):
        tb  = extract_block(sfc_block, tm.end()-1)
        pos = re.search(r'POSITION\s*=\s*\{\s*X=(\d+)\s*Y=(\d+)', tb)
        td  = re.search(r'DESCRIPTION="([^"]+)"', tb)
        te  = re.search(r'EXPRESSION="([^"]*)"', tb, re.DOTALL)
        tt  = re.search(r'TERMINATION=(\w+)', tb)
        trans[tm.group(1)] = {
            'x': int(pos.group(1)) if pos else 0,
            'y': int(pos.group(2)) if pos else 0,
            'description': td.group(1).strip() if td else '',
            'expression':  te.group(1).strip() if te else '',
            'termination': tt.group(1) if tt else 'F',
        }
    s2t = {s: [] for s in steps}
    for tn, tp in trans.items():
        cands = [(tp['y']-v['y'], sn) for sn,v in steps.items()
                 if v['y'] < tp['y'] and abs(v['x']-tp['x']) < 150]
        if not cands:
            cands = [(tp['y']-v['y'], sn) for sn,v in steps.items() if v['y'] < tp['y']]
        if cands:
            s2t[sorted(cands)[0][1]].append(tn)
    return {
        'ordered_steps': sorted(steps.items(), key=lambda x: x[1]['y']),
        'transitions':   trans,
        'step_to_trans': s2t,
    }

# ═══════════════════════════════════════════════════════════════════════════════
# SHARED EXCEL STYLES
# ═══════════════════════════════════════════════════════════════════════════════

NAVY    = PatternFill('solid', start_color='0D1B4B')
BLUE_H  = PatternFill('solid', start_color='1F3864')
BLUE_S  = PatternFill('solid', start_color='2E75B6')
TEAL_S  = PatternFill('solid', start_color='1F5C6B')
GREEN_S = PatternFill('solid', start_color='4E7F2C')
ORANGE_H= PatternFill('solid', start_color='7B3F00')
OPEN_F  = PatternFill('solid', start_color='C6EFCE')
CLOSE_F = PatternFill('solid', start_color='FFCCCC')
DC_F    = PatternFill('solid', start_color='FFFFCC')
ALT     = PatternFill('solid', start_color='D9E1F2')
ALT_G   = PatternFill('solid', start_color='E2EFDA')
ALTG2   = PatternFill('solid', start_color='F2F9EE')
ALT_ROW = PatternFill('solid', start_color='F2F2F2')
WHITE   = PatternFill('solid', start_color='FFFFFF')
DIS_F   = PatternFill('solid', start_color='DDDDDD')
THIN    = Side(style='thin', color='AAAAAA')
BORD    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NCOLS   = 10
COL_W   = [14, 22, 32, 10, 10, 58, 10, 38, 34, 14]

def wf(bold=False, sz=10, color='000000'):
    return Font(name='Calibri', bold=bold, size=sz, color=color)
def wa(h='left', wrap=False):
    return Alignment(horizontal=h, vertical='top', wrap_text=wrap)
def sc(ws, r, c, val='', bold=False, sz=10, fc='000000', fill=None,
       h='left', wrap=False, merge_to=None):
    cell = ws.cell(row=r, column=c, value=str(val) if val is not None else '')
    cell.font = wf(bold, sz, fc); cell.fill = fill or WHITE
    cell.alignment = wa(h, wrap); cell.border = BORD
    if merge_to:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=merge_to)
    return cell

def write_sfc_sheet(wb, label, title_fill, step_fill, data, opts, fname_desc):
    ws = wb.create_sheet(title=label[:31])
    for ci, w in enumerate(COL_W, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A3'

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    t = ws.cell(row=1, column=1, value=f"  {fname_desc}")
    t.font = wf(True,12,'FFFFFF'); t.fill = NAVY
    t.alignment = Alignment(horizontal='left', vertical='center'); t.border = BORD
    ws.row_dimensions[1].height = 24

    # Column headers
    hdrs = ['Row Type','Step / Transition','Description','Action','Qualifier',
            'Expression' if opts.get('expressions',True) else '(hidden)',
            'Delay','Delay Expression','Confirm Expression','Confirm Timeout']
    for ci, h in enumerate(hdrs, 1):
        sc(ws, 2, ci, h, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
    ws.row_dimensions[2].height = 17

    row = 3
    for si, (sn, sd) in enumerate(data['ordered_steps']):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
        sh = ws.cell(row=row, column=1,
                     value=f"  STEP {si+1}:  {sn}   "
                           f"({len(sd['actions'])} action{'s' if len(sd['actions'])!=1 else ''})")
        sh.font = wf(True,10,'FFFFFF'); sh.fill = step_fill
        sh.alignment = Alignment(horizontal='left', vertical='center'); sh.border = BORD
        ws.row_dimensions[row].height = 17; row += 1

        for ai, a in enumerate(sd['actions']):
            f = ALT if ai%2==0 else WHITE
            ie = opts.get('expressions', True)
            sc(ws,row,1, 'ACTION',                              bold=True, fill=f, h='center')
            sc(ws,row,2, sn,                                    fill=f)
            sc(ws,row,3, a['description'],                      fill=f, wrap=True)
            sc(ws,row,4, a['action'],                           fill=f, h='center')
            sc(ws,row,5, a['qualifier'],                        fill=f, h='center')
            sc(ws,row,6, a['expression'] if ie else '—',        fill=f, wrap=True)
            sc(ws,row,7, a['delay_time'],                       fill=f, h='center')
            sc(ws,row,8, a['delay_expression'],                 fill=f, wrap=True)
            sc(ws,row,9, a['confirm_expression'],               fill=f, wrap=True)
            sc(ws,row,10,a['confirm_timeout'],                  fill=f, h='center')
            ws.row_dimensions[row].height = max(15, min(75,
                15*max(1, len(a.get('expression',''))//55+1)))
            row += 1

        if opts.get('transitions', True):
            tl = data['step_to_trans'].get(sn, [])
            if tl:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
                th = ws.cell(row=row, column=1,
                             value=f"  ↓  TRANSITION{'S' if len(tl)>1 else ''} "
                                   f"FROM  {sn}  ({len(tl)})")
                th.font = wf(True,9,'FFFFFF'); th.fill = GREEN_S
                th.alignment = Alignment(horizontal='left', vertical='center'); th.border = BORD
                ws.row_dimensions[row].height = 15; row += 1

                for ti, tn in enumerate(tl):
                    tr = data['transitions'][tn]
                    f  = ALT_G if ti%2==0 else ALTG2
                    sc(ws,row,1,'TRANSITION', bold=True, fill=f, h='center')
                    sc(ws,row,2,tn,           bold=True, fill=f)
                    sc(ws,row,3,tr['description'], fill=f, wrap=True)
                    sc(ws,row,4,'⏹ END' if tr['termination']=='T' else '→ NEXT',
                       fill=f, h='center')
                    sc(ws,row,5,'',fill=f)
                    ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=NCOLS)
                    ec = ws.cell(row=row, column=6,
                                 value=tr['expression'] if opts.get('expressions',True) else '—')
                    ec.font = wf(); ec.fill = f
                    ec.alignment = wa('left',True); ec.border = BORD
                    ws.row_dimensions[row].height = max(15, min(60,
                        15*max(1, len(tr.get('expression',''))//80+1)))
                    row += 1

        for ci in range(1, NCOLS+1): sc(ws, row, ci, '', fill=WHITE)
        ws.row_dimensions[row].height = 6; row += 1

    return ws

# ═══════════════════════════════════════════════════════════════════════════════
# PHASE FHX PARSER
# ═══════════════════════════════════════════════════════════════════════════════

def build_instance_map(text):
    mapping = {}
    for m in re.finditer(r'FUNCTION_BLOCK\s+NAME="([^"]+)"\s+DEFINITION="([^"]+)"', text):
        inst, defn = m.group(1), m.group(2)
        if defn not in mapping:
            mapping[defn] = inst
    return mapping

def clean_label(raw):
    s = raw.strip().upper()
    s = re.sub(r'[\-_/\\]+', ' ', s)
    for noise in [' LOGIC',' PHASE',' SEQUENCE',' SUB',' FUNCTION']:
        s = s.replace(noise, '')
    s = re.sub(r'\s+', '_', s.strip())[:28]
    return re.sub(r'[\/\\\?\*\[\]\:\']', '_', s) or raw[:28]

def derive_phase_label(fb_name, instance_name, description, used_labels):
    if instance_name and instance_name.upper() not in ('','NONE'):
        candidate = clean_label(instance_name)
    else:
        desc_up = description.upper()
        _KW = [
            ('ABORT','ABORT_STOP'),('STOP','ABORT_STOP'),
            ('RUNNING','RUNNING'),('RUN ','RUNNING'),
            ('HOLDING','HOLD'),('HOLD','HOLD'),
            ('RESTART','RESTART'),('CONDITION','CONDITION'),
            ('PROMPT','PROMPT'),('INIT','INIT'),
            ('START','START'),('COMPLETE','COMPLETE'),
            ('FAIL','FAIL'),('PAUSE','PAUSE'),
        ]
        candidate = None
        for kw, lbl in _KW:
            if kw in desc_up:
                candidate = lbl; break
        if not candidate:
            words = re.sub(r'[^A-Z0-9\s]', ' ', desc_up).split()
            noise = {'LOGIC','PHASE','THE','FOR','AND','OR','A','AN','OF','FUNCTION','BLOCK','STEP','SUB','SEQUENCE'}
            words = [w for w in words if w not in noise and len(w)>1]
            candidate = '_'.join(words[-2:]) if words else fb_name[:20]
    label = re.sub(r'[\/\\\?\*\[\]\:\']', '_', candidate)[:28]
    base, n = label, 2
    while label in used_labels:
        label = f"{base[:25]}_{n}"; n += 1
    used_labels.add(label)
    return label

def parse_phase_fhx(text):
    instance_map = build_instance_map(text)
    blocks = {}
    for m in re.finditer(r'FUNCTION_BLOCK_DEFINITION\s+NAME="([^"]+)"[^\{]*\{', text):
        name  = m.group(1)
        block = extract_block(text, m.end()-1)
        desc  = re.search(r'DESCRIPTION="([^"]+)"', block)
        sfc_m = re.search(r'SFC_ALGORITHM\s*\{', block)
        sfc   = extract_block(block, sfc_m.end()-1) if sfc_m else ''
        sfc_data = parse_sfc(sfc)
        blocks[name] = {
            'description':   desc.group(1) if desc else '',
            'instance_name': instance_map.get(name, ''),
            **sfc_data,
        }
    return blocks

def build_phase_excel(blocks, fname, opts):
    wb = openpyxl.Workbook()
    used_labels = set()
    block_labels = {
        fb: derive_phase_label(fb, d.get('instance_name',''), d.get('description',''), used_labels)
        for fb, d in blocks.items()
    }

    if opts.get('summary', True):
        ws = wb.active; ws.title = 'SUMMARY'
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
        t = ws.cell(row=1, column=1, value=f"  {fname}  —  DeltaV Phase Export Logic Summary")
        t.font = wf(True,13,'FFFFFF'); t.fill = NAVY
        t.alignment = Alignment(horizontal='left', vertical='center'); t.border = BORD
        ws.row_dimensions[1].height = 28
        for ci, h in enumerate(['Logic Block','FB Name','Description','Steps',
                                 'Actions','Transitions','Orphan Trans.','Sheet Tab'], 1):
            sc(ws, 2, ci, h, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        ws.row_dimensions[2].height = 17
        for ci, w in enumerate([18,30,50,8,10,13,14,18], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A3'
        for i, (fb, data) in enumerate(blocks.items()):
            lbl    = block_labels[fb]
            acts   = sum(len(s[1]['actions']) for s in data['ordered_steps'])
            mapped = sum(len(v) for v in data['step_to_trans'].values())
            orphan = len(data['transitions']) - mapped
            f      = ALT if i%2==0 else WHITE
            r      = i+3
            sc(ws,r,1,lbl,bold=True,fill=f); sc(ws,r,2,fb,fill=f)
            sc(ws,r,3,data['description'],fill=f,wrap=True)
            sc(ws,r,4,len(data['ordered_steps']),fill=f,h='center')
            sc(ws,r,5,acts,fill=f,h='center')
            sc(ws,r,6,len(data['transitions']),fill=f,h='center')
            sc(ws,r,7,orphan or '-',fill=f,h='center')
            sc(ws,r,8,lbl,fill=f)
            ws.row_dimensions[r].height = 16
    else:
        wb.active.title = '_temp'

    for fb, data in blocks.items():
        lbl = block_labels[fb]
        write_sfc_sheet(wb, lbl, NAVY, BLUE_S, data, opts,
                        f"Phase Logic: {lbl}   |   {data['description']}")

    if '_temp' in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb['_temp']
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════════════════
# COMMAND DRIVEN EM PARSER
# ═══════════════════════════════════════════════════════════════════════════════

def parse_cdem_fhx(text):
    results = []
    for mc_m in re.finditer(r'MODULE_CLASS\s+NAME="([^"]+)"[^\{]*\{', text):
        em_name  = mc_m.group(1)
        mc_block = extract_block(text, mc_m.end()-1)
        if 'COMMAND_DRIVEN_ALGORITHM' not in mc_block:
            continue
        em_desc_m = re.search(r'DESCRIPTION="([^"]+)"', mc_block)
        em_desc   = em_desc_m.group(1) if em_desc_m else em_name

        cda_m = re.search(r'COMMAND_DRIVEN_ALGORITHM\s*\{', mc_block)
        cda   = extract_block(mc_block, cda_m.end()-1)

        cmd_index_to_name = {}
        for tm in re.finditer(r'TRANSITION\s+NAME="T_IN_(\d+)"\s*\{', cda):
            idx = int(tm.group(1))
            tb  = extract_block(cda, tm.end()-1)
            expr = re.search(r'EXPRESSION="([^"]+)"', tb)
            if expr:
                nm = re.search(r"'[^']+:([^']+)'", expr.group(1))
                if nm: cmd_index_to_name[idx] = nm.group(1).strip()

        cmd_inst_to_def = {}
        for fb_m in re.finditer(
            r'FUNCTION_BLOCK\s+NAME="(COMMAND_\d+)"\s+DEFINITION="([^"]+)"', mc_block):
            idx = int(re.search(r'(\d+)$', fb_m.group(1)).group(1))
            cmd_inst_to_def[idx] = (fb_m.group(1), fb_m.group(2))

        for idx in sorted(cmd_inst_to_def.keys()):
            inst, defn = cmd_inst_to_def[idx]
            cmd_name   = cmd_index_to_name.get(idx, inst)
            fb_m = re.search(
                r'FUNCTION_BLOCK_DEFINITION\s+NAME="' + re.escape(defn) + r'"[^\{]*\{', text)
            if not fb_m: continue
            fb_block = extract_block(text, fb_m.end()-1)
            fb_desc  = re.search(r'DESCRIPTION="([^"]+)"', fb_block)
            sfc_m    = re.search(r'SFC_ALGORITHM\s*\{', fb_block)
            if not sfc_m: continue
            sfc_data = parse_sfc(extract_block(fb_block, sfc_m.end()-1))
            results.append({
                'em_name':        em_name,
                'em_description': em_desc,
                'command_name':   cmd_name,
                'command_index':  idx,
                'fb_definition':  defn,
                'fb_description': fb_desc.group(1) if fb_desc else '',
                **sfc_data,
            })
    return results

def build_cdem_excel(commands, fname, opts):
    wb = openpyxl.Workbook()

    if opts.get('summary', True):
        ws = wb.active; ws.title = 'SUMMARY'
        em_name = commands[0]['em_name'] if commands else fname
        em_desc = commands[0]['em_description'] if commands else ''
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
        t = ws.cell(row=1, column=1, value=f"  {em_name}  —  Command Driven EM Logic Summary")
        t.font = wf(True,13,'FFFFFF'); t.fill = NAVY
        t.alignment = Alignment(horizontal='left', vertical='center'); t.border = BORD
        ws.row_dimensions[1].height = 28
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
        s = ws.cell(row=2, column=1, value=f"  {em_desc}")
        s.font = wf(False,10,'CCCCCC'); s.fill = NAVY
        s.alignment = Alignment(horizontal='left', vertical='center'); s.border = BORD
        ws.row_dimensions[2].height = 16
        for ci, h in enumerate(['#','Command','Description','Steps',
                                 'Actions','Transitions','Sheet Tab','FB Definition'], 1):
            sc(ws, 3, ci, h, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        ws.row_dimensions[3].height = 17
        for ci, w in enumerate([5,18,42,7,10,12,16,30], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A4'
        for i, cmd in enumerate(commands):
            acts = sum(len(s[1]['actions']) for s in cmd['ordered_steps'])
            f    = ALT if i%2==0 else WHITE; r = i+4
            sc(ws,r,1,cmd['command_index'],     fill=f,h='center')
            sc(ws,r,2,cmd['command_name'],       fill=f,bold=True)
            sc(ws,r,3,cmd['fb_description'],     fill=f,wrap=True)
            sc(ws,r,4,len(cmd['ordered_steps']), fill=f,h='center')
            sc(ws,r,5,acts,                      fill=f,h='center')
            sc(ws,r,6,len(cmd['transitions']),   fill=f,h='center')
            sc(ws,r,7,cmd['command_name'][:31],  fill=f)
            sc(ws,r,8,cmd['fb_definition'],      fill=f)
            ws.row_dimensions[r].height = 16
    else:
        wb.active.title = '_temp'

    used_tabs = set()
    for cmd in commands:
        base = re.sub(r'[\/\\\?\*\[\]\:\s]', '_', cmd['command_name'])[:28]
        tab  = base; n = 2
        while tab in used_tabs: tab = f"{base[:25]}_{n}"; n += 1
        used_tabs.add(tab)
        write_sfc_sheet(wb, tab, NAVY, TEAL_S, cmd, opts,
                        f"EM Command: {cmd['command_name']}   |   "
                        f"{cmd['em_name']}  —  {cmd['fb_description']}")

    if '_temp' in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb['_temp']
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════════════════
# STATE DRIVEN EM PARSER
# ═══════════════════════════════════════════════════════════════════════════════

def parse_sdem_fhx(text):
    results = []
    for mc_m in re.finditer(r'MODULE_CLASS\s+NAME="([^"]+)"[^\{]*\{', text):
        em_name  = mc_m.group(1)
        mc_block = extract_block(text, mc_m.end()-1)
        if 'STATE_DRIVEN_ALGORITHM' not in mc_block:
            continue
        em_desc_m = re.search(r'DESCRIPTION="([^"]+)"', mc_block)
        em_desc   = em_desc_m.group(1) if em_desc_m else em_name

        sda_m = re.search(r'STATE_DRIVEN_ALGORITHM\s*\{', mc_block)
        sda   = extract_block(mc_block, sda_m.end()-1)
        sfc_m = re.search(r'SFC_ALGORITHM\s*\{', sda)
        if not sfc_m: continue
        sfc = extract_block(sda, sfc_m.end()-1)

        index_to_name = {}
        for tm in re.finditer(r'TRANSITION\s+NAME="T_IN_(\d+)"\s*\{', sfc):
            idx  = int(tm.group(1))
            tb   = extract_block(sfc, tm.end()-1)
            expr = re.search(r'EXPRESSION="([^"]+)"', tb)
            if expr:
                nm = re.search(r"'[^']+:([^']+)'", expr.group(1))
                if nm: index_to_name[idx] = nm.group(1).strip()

        state_def = None
        for fb_m in re.finditer(
            r'FUNCTION_BLOCK\s+NAME="STATE_\d+"\s+DEFINITION="([^"]+)"', mc_block):
            state_def = fb_m.group(1); break
        if not state_def: continue

        fb_def_m = re.search(
            r'FUNCTION_BLOCK_DEFINITION\s+NAME="' + re.escape(state_def) + r'"[^\{]*\{', text)
        if not fb_def_m: continue
        fb_def_block = extract_block(text, fb_def_m.end()-1)

        devices, default_vals, default_dc = [], {}, {}
        for am in re.finditer(r'ATTRIBUTE\s+NAME="([^"]+)"[^\{]*\{', fb_def_block):
            ab = extract_block(fb_def_block, am.end()-1)
            if re.search(r'GROUP="I/O"', ab):
                devices.append(am.group(1))
        for ai_m in re.finditer(r'ATTRIBUTE_INSTANCE\s+NAME="([^"]+)"\s*\{', fb_def_block):
            ai_b  = extract_block(fb_def_block, ai_m.end()-1)
            dname = ai_m.group(1)
            val_m = re.search(r'STRING_VALUE="([^"]+)"', ai_b)
            cv_m  = re.search(r'CV=([TF\d\.]+)', ai_b)
            dc_m  = re.search(r'SDA_DONT_CARE=([TF])', ai_b)
            default_vals[dname] = val_m.group(1) if val_m else (cv_m.group(1) if cv_m else '')
            default_dc[dname]   = dc_m.group(1)=='T' if dc_m else False

        state_overrides, enabled_overrides = {}, {}
        for ai_m in re.finditer(
            r'ATTRIBUTE_INSTANCE\s+NAME="STATE_(\d+)/([^"]+)"\s*\{', mc_block):
            idx, dname = int(ai_m.group(1)), ai_m.group(2)
            ai_b  = extract_block(mc_block, ai_m.end()-1)
            val_m = re.search(r'STRING_VALUE="([^"]+)"', ai_b)
            cv_m  = re.search(r'CV=([TF\d\.]+)', ai_b)
            dc_m  = re.search(r'SDA_DONT_CARE=([TF])', ai_b)
            val   = val_m.group(1) if val_m else (cv_m.group(1) if cv_m else '')
            dc    = dc_m.group(1)=='T' if dc_m else False
            if dname == 'ENABLED':
                enabled_overrides[idx] = (cv_m.group(1) != 'F') if cv_m else True
            else:
                state_overrides.setdefault(idx, {})[dname] = (val, dc)

        states = []
        for idx in sorted(index_to_name.keys()):
            ov  = state_overrides.get(idx, {})
            row_vals, row_dc = {}, {}
            for dev in devices:
                if dev in ov: row_vals[dev], row_dc[dev] = ov[dev]
                else: row_vals[dev] = default_vals.get(dev,''); row_dc[dev] = default_dc.get(dev,False)
            states.append({
                'index':      idx,
                'state_name': index_to_name[idx],
                'enabled':    enabled_overrides.get(idx, True),
                'values':     row_vals,
                'dont_care':  row_dc,
            })

        results.append({
            'em_name':      em_name,
            'em_description': em_desc,
            'devices':      devices,
            'default_vals': default_vals,
            'states':       states,
        })
    return results

def build_sdem_excel(em_list, fname, opts):
    wb = openpyxl.Workbook(); first = True

    for em in em_list:
        devices, states = em['devices'], em['states']
        em_name, em_desc = em['em_name'], em['em_description']
        tab = re.sub(r'[\/\\\?\*\[\]\:\s]', '_', em_name)[:31]
        ws  = wb.active if first else wb.create_sheet(title=tab)
        if first: ws.title = tab; first = False
        NCOLS_T = 3 + len(devices) + 1

        # Title
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS_T)
        t = ws.cell(row=1, column=1,
                    value=f"  {em_name}  —  State Driven EM  |  {em_desc}")
        t.font = wf(True,13,'FFFFFF'); t.fill = NAVY
        t.alignment = Alignment(horizontal='left', vertical='center'); t.border = BORD
        ws.row_dimensions[1].height = 28

        # Device header row
        for ci in range(1, 4):
            sc(ws, 2, ci, '', fill=BLUE_H)
        for ci, dev in enumerate(devices, 4):
            sc(ws, 2, ci, dev, bold=True, fc='FFFFFF', fill=ORANGE_H, sz=11, h='center')
        sc(ws, 2, 4+len(devices), 'Min Time', bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        ws.row_dimensions[2].height = 20

        # Column headers
        for ci, h in enumerate(['#','State Name','Enabled'], 1):
            sc(ws, 3, ci, h, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        for ci, dev in enumerate(devices, 4):
            sc(ws, 3, ci, dev, bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        sc(ws, 3, 4+len(devices), 'CFM Min Time', bold=True, fc='FFFFFF', fill=BLUE_H, h='center')
        ws.row_dimensions[3].height = 17
        ws.freeze_panes = 'A4'

        # State rows
        for ri, state in enumerate(states):
            r   = ri+4
            bg  = ALT_ROW if ri%2==0 else WHITE
            dis = not state['enabled']
            sc(ws,r,1, state['index'],      fill=DIS_F if dis else bg, h='center')
            sc(ws,r,2, state['state_name'], bold=True, fill=DIS_F if dis else bg,
               h='left', fc='888888' if dis else '000000')
            sc(ws,r,3, 'No' if dis else 'Yes',
               fill=DIS_F if dis else bg,
               fc='888888' if dis else '006100', h='center')
            for ci, dev in enumerate(devices, 4):
                val, dc = state['values'].get(dev,''), state['dont_care'].get(dev,False)
                if dis:     fill_c = DIS_F
                elif dc:    fill_c = DC_F
                elif val.upper()=='OPEN':  fill_c = OPEN_F
                elif val.upper()=='CLOSE': fill_c = CLOSE_F
                else:       fill_c = bg
                display = '—' if dc else val
                sc(ws,r,ci, display, fill=fill_c, h='center',
                   fc='888888' if dis or dc else '000000',
                   bold=(val.upper()=='OPEN' and not dc and not dis))
            sc(ws,r, 4+len(devices), '', fill=DIS_F if dis else bg, h='center')
            ws.row_dimensions[r].height = 16

        # Legend
        lr = len(states)+5
        ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=NCOLS_T)
        lh = ws.cell(row=lr, column=1, value='  LEGEND')
        lh.font = wf(True,10,'FFFFFF'); lh.fill = BLUE_H
        lh.alignment = Alignment(horizontal='left', vertical='center'); lh.border = BORD
        ws.row_dimensions[lr].height = 16
        for li, (fill_c, lbl) in enumerate([
            (OPEN_F,  'OPEN  —  device commanded open / active'),
            (CLOSE_F, 'CLOSE  —  device commanded closed / inactive'),
            (DC_F,    '—  (Don\'t Care)  —  device not controlled in this state'),
            (DIS_F,   'State disabled / not used'),
        ]):
            rr = lr+1+li
            sc(ws, rr, 1, '', fill=fill_c)
            ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=NCOLS_T)
            lc = ws.cell(row=rr, column=2, value=f'  {lbl}')
            lc.font = wf(False,10); lc.fill = WHITE
            lc.alignment = Alignment(horizontal='left', vertical='center'); lc.border = BORD
            ws.row_dimensions[rr].height = 15

        # Column widths
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 36
        ws.column_dimensions['C'].width = 9
        for ci in range(4, 4+len(devices)):
            ws.column_dimensions[get_column_letter(ci)].width = 14
        ws.column_dimensions[get_column_letter(4+len(devices))].width = 14

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

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

    fhx_type = request.form.get('fhx_type', 'phase')   # phase | em_cd | em_sd
    opts = {
        'summary':     request.form.get('summary',     'true') == 'true',
        'transitions': request.form.get('transitions', 'true') == 'true',
        'expressions': request.form.get('expressions', 'true') == 'true',
    }
    raw   = f.read()
    text  = decode_fhx(raw)
    fname = re.sub(r'\.fhx$', '', f.filename, flags=re.IGNORECASE)

    try:
        if fhx_type == 'phase':
            blocks = parse_phase_fhx(text)
            if not blocks:
                return jsonify({'error': 'No FUNCTION_BLOCK_DEFINITION blocks found. Is this a Phase FHX?'}), 400
            buf = build_phase_excel(blocks, fname, opts)

        elif fhx_type == 'em_cd':
            commands = parse_cdem_fhx(text)
            if not commands:
                return jsonify({'error': 'No COMMAND_DRIVEN_ALGORITHM found. Is this a Command Driven EM FHX?'}), 400
            buf = build_cdem_excel(commands, fname, opts)

        elif fhx_type == 'em_sd':
            em_list = parse_sdem_fhx(text)
            if not em_list:
                return jsonify({'error': 'No STATE_DRIVEN_ALGORITHM found. Is this a State Driven EM FHX?'}), 400
            buf = build_sdem_excel(em_list, fname, opts)

        else:
            return jsonify({'error': f'Unknown fhx_type: {fhx_type}'}), 400

    except Exception as e:
        return jsonify({'error': f'Parse error: {str(e)}'}), 500

    # Build response with sheet names in header so client can show pills
    response = send_file(buf, as_attachment=True,
                         download_name=fname + '_Logic.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Read workbook sheet names before the buffer is consumed
    buf.seek(0)
    import openpyxl as _ox
    _wb = _ox.load_workbook(buf, read_only=True)
    response.headers['X-Sheet-Names'] = '|'.join(_wb.sheetnames)
    _wb.close()
    buf.seek(0)
    return response

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
