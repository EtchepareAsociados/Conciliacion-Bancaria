import os, re, json, io
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'etchepare2026conciliacion')

# ── Usuarios permitidos ──
USUARIOS = {
    os.environ.get('USER1_NAME', 'admin'):    os.environ.get('USER1_PASS', 'etchepare2026'),
    os.environ.get('USER2_NAME', 'compañera'): os.environ.get('USER2_PASS', 'etchepare2026b'),
}

TOLERANCIA = 10000
MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
         'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

MES_BG = {'ENERO':'DBEAFE','FEBRERO':'BAE6FD','MARZO':'BBF7D0','ABRIL':'FEF08A',
           'MAYO':'FED7AA','JUNIO':'FECACA','JULIO':'E9D5FF','AGOSTO':'FBCFE8',
           'SEPTIEMBRE':'A7F3D0','OCTUBRE':'FDE68A','NOVIEMBRE':'BAE6FD','DICIEMBRE':'DDD6FE'}
MES_FG = {'ENERO':'1E3A8A','FEBRERO':'075985','MARZO':'14532D','ABRIL':'713F12',
           'MAYO':'7C2D12','JUNIO':'7F1D1D','JULIO':'4C1D95','AGOSTO':'831843',
           'SEPTIEMBRE':'064E3B','OCTUBRE':'78350F','NOVIEMBRE':'0C4A6E','DICIEMBRE':'3B0764'}

C_NAVY  = '1F3864'
C_WHITE = 'FFFFFF'
C_GRAY_H= 'F8FAFC'
C_GRAY_A= 'F1F5F9'
C_GRAY_B= 'CBD5E1'
C_BLACK = '1E293B'
C_MUTED = '475569'

def thin_border(c=C_GRAY_B):
    s = Side(style='thin', color=c)
    return Border(left=s, right=s, top=s, bottom=s)

def medium_bottom(bc=C_NAVY, sc=C_GRAY_B):
    return Border(left=Side(style='thin',color=sc), right=Side(style='thin',color=sc),
                  top=Side(style='thin',color=sc), bottom=Side(style='medium',color=bc))

# ── Helpers ──
def extract_rut(desc):
    if not desc or pd.isna(desc): return None
    m = re.match(r'^(\d{7,10}[Kk]?)\s', str(desc))
    return m.group(1).upper() if m else None

def norm_rut(r):
    if not r: return None
    return str(r).strip().replace('.','').replace('-','').upper().lstrip('0')

def es_efectivo(desc):
    if not desc: return False
    return any(x in str(desc).upper() for x in ['DEPOS.DOCTO','DEPOSITO EFECTIVO','DOCTO.O.BANCOS'])

def sig_mes(p):
    if not p: return ''
    u = str(p).strip().upper()
    return MESES[(MESES.index(u)+1)%12] if u in MESES else ''

# ── Load BD ──
def load_bd():
    bd_path = os.path.join(os.path.dirname(__file__), 'bd_arrendatarios.json')
    with open(bd_path, encoding='utf-8') as f:
        records = json.load(f)
    lookup = {}
    for r in records:
        k = norm_rut(r.get('RESPONSABLE',''))
        if k:
            lookup.setdefault(k, []).append(r)
    return lookup

BD_LOOKUP = load_bd()

def find_match(rut_norm, monto):
    if not rut_norm or rut_norm not in BD_LOOKUP: return None
    cands = BD_LOOKUP[rut_norm]
    for c in cands:
        if c['MONTO_ESP'] == monto:
            return {**c, 'tipo':'OK', 'diff':0}
    for c in cands:
        if c.get('MONEDA','CLP')=='UF' and abs(c['MONTO_ESP']-monto)<=TOLERANCIA:
            return {**c, 'tipo':'OK', 'diff':monto-c['MONTO_ESP']}
    best = min(cands, key=lambda c: abs(c['MONTO_ESP']-monto))
    diff = monto - best['MONTO_ESP']
    return {**best, 'tipo':'DIF+' if diff>0 else 'DIF-', 'diff':diff}

# ── Parse historial ──
def parse_historial(wb):
    keys = set()
    ultimo_mes = {}
    for sname in wb.sheetnames:
        if sname == 'RESUMEN': continue
        ws = wb[sname]
        rows = list(ws.iter_rows(values_only=True))
        # Find header row
        data_start = 1
        for i, row in enumerate(rows[:5]):
            if row and str(row[0]).strip().upper() == 'FECHA':
                data_start = i + 1
                break
        for row in rows[data_start:]:
            if not row or not row[0]: continue
            fecha  = str(row[0]).strip() if row[0] else ''
            desc   = str(row[3]).strip() if row[3] else ''
            monto  = row[4]
            periodo= str(row[5]).strip().upper() if row[5] else ''
            rut    = norm_rut(extract_rut(desc))
            if fecha and monto:
                try:
                    keys.add(f"{rut}|{int(float(monto))}|{fecha}")
                except: pass
            if rut and periodo in MESES:
                ultimo_mes[rut] = periodo
    return keys, ultimo_mes

# ── Parse cartola ──
def parse_cartola(wb):
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    abonos = []
    for row in rows[13:]:
        if not row or len(row) < 8: continue
        try: monto = int(float(row[0]))
        except: continue
        if str(row[7]).strip() != 'A': continue
        desc  = str(row[1]).strip() if row[1] else ''
        fecha = str(row[3]).strip() if row[3] else ''
        ndoc  = str(row[4]).strip() if row[4] else ''
        abonos.append({
            'monto': monto, 'desc': desc, 'fecha': fecha, 'ndoc': ndoc,
            'rut': extract_rut(desc), 'rut_norm': norm_rut(extract_rut(desc)),
            'ef': es_efectivo(desc)
        })
    return abonos

# ── Process ──
def procesar(hist_wb, cartola_wb):
    hist_keys, ultimo_mes = parse_historial(hist_wb)
    abonos = parse_cartola(cartola_wb)

    # Detect month
    mes_cartola = ''
    fechas = [a['fecha'] for a in abonos if a['fecha']]
    if fechas:
        try: mes_cartola = MESES[int(fechas[0].split('/')[1])-1]
        except: pass

    res_ok, res_dif, res_res, res_efec = [], [], [], []

    for a in abonos:
        key = f"{a['rut_norm']}|{a['monto']}|{a['fecha']}"
        if key in hist_keys: continue

        base = {
            'fecha': a['fecha'], 'rut': a['rut'] or 'Sin RUT',
            'monto': a['monto'], 'desc': a['desc'][:55],
            'ndoc': a['ndoc'], 'mes': mes_cartola
        }

        if a['ef'] and not a['rut']:
            res_efec.append({**base, 'carpeta':'', 'estado':'EFECTIVO'})
            continue

        match = find_match(a['rut_norm'], a['monto'])
        if not match:
            res_res.append({**base, 'carpeta':'', 'estado':'RESERVA'})
            continue

        ult = ultimo_mes.get(a['rut_norm'], '')
        mes = sig_mes(ult) or mes_cartola
        carpeta = str(match['CARPETA'])

        if match['tipo'] == 'OK':
            res_ok.append({**base, 'carpeta':carpeta, 'mes':mes, 'estado':'OK', 'diff':match['diff']})
        else:
            diff = match['diff']
            label = f"▲ DE MÁS +${diff:,.0f}" if diff > 0 else f"▼ DE MENOS -${abs(diff):,.0f}"
            res_dif.append({**base, 'carpeta':carpeta, 'mes':mes, 'estado':label, 'diff':diff})

    return res_ok, res_dif, res_res, res_efec, mes_cartola

# ── Generate Excel with full formatting ──
def generar_excel(hist_wb, res_ok, res_dif, res_res, res_efec):
    SHEET_MAP = {
        'ARRIENDOS':      ([*res_ok, *res_dif], True,  '2563EB'),
        'RESERVAS':       (res_res,              False, 'D97706'),
        'EFECTIVO-CHEQUE':(res_efec,             False, 'CA8A04'),
    }

    for sheet_name, (rows, has_carpeta, accent) in SHEET_MAP.items():
        if not rows: continue
        if sheet_name not in hist_wb.sheetnames: continue
        ws = hist_wb[sheet_name]

        # Find last row with data
        last_row = ws.max_row
        while last_row > 3 and ws.cell(last_row, 1).value in [None, '']:
            last_row -= 1

        for row_data in rows:
            last_row += 1
            row_bg = C_GRAY_H if last_row % 2 == 0 else C_WHITE

            if has_carpeta:
                vals = [row_data['fecha'], row_data.get('carpeta',''),
                        row_data['rut'], row_data['desc'],
                        row_data['monto'], row_data.get('mes',''),
                        row_data['estado'], '']
            else:
                vals = [row_data['fecha'], '',
                        row_data['rut'], row_data['desc'],
                        row_data['monto'], row_data.get('mes',''),
                        row_data['estado'], '']

            for col, val in enumerate(vals, 1):
                c = ws.cell(row=last_row, column=col, value=val)
                c.fill   = PatternFill('solid', fgColor=row_bg)
                c.border = Border(bottom=Side(style='thin',color='E2E8F0'),
                                  left=Side(style='thin',color='E2E8F0'),
                                  right=Side(style='thin',color='E2E8F0'))
                c.font   = Font(name='Calibri', size=9, color=C_BLACK)
                c.alignment = Alignment(vertical='center')

            # FECHA
            ws.cell(last_row,1).alignment = Alignment(horizontal='center',vertical='center')
            # CARPETA
            ws.cell(last_row,2).font = Font(name='Calibri',bold=True,size=10,color=C_NAVY)
            ws.cell(last_row,2).alignment = Alignment(horizontal='center',vertical='center')
            # RUT
            ws.cell(last_row,3).font = Font(name='Courier New',size=9,color=C_BLACK)
            ws.cell(last_row,3).alignment = Alignment(horizontal='center',vertical='center')
            # MONTO
            mc = ws.cell(last_row,5)
            mc.font = Font(name='Courier New',bold=True,size=9,color=C_BLACK)
            mc.number_format = '#,##0'
            mc.alignment = Alignment(horizontal='right',vertical='center')
            # PERÍODO
            pc = ws.cell(last_row,6)
            per = str(pc.value).strip().upper() if pc.value else ''
            if per in MES_BG:
                pc.fill = PatternFill('solid',fgColor=MES_BG[per])
                pc.font = Font(name='Calibri',bold=True,size=8,color=MES_FG.get(per,C_BLACK))
            else:
                pc.font = Font(name='Calibri',size=8,color=C_BLACK)
            pc.alignment = Alignment(horizontal='center',vertical='center')
            # ESTADO
            ec = ws.cell(last_row,7)
            estado = str(ec.value) if ec.value else ''
            if estado.startswith('▲'):
                ec.font = Font(name='Calibri',bold=True,size=8,color='1D4ED8')
                ec.fill = PatternFill('solid',fgColor='DBEAFE')
            elif estado.startswith('▼'):
                ec.font = Font(name='Calibri',bold=True,size=8,color='991B1B')
                ec.fill = PatternFill('solid',fgColor='FEE2E2')
            elif estado == 'OK':
                ec.font = Font(name='Calibri',bold=True,size=8,color='14532D')
                ec.fill = PatternFill('solid',fgColor='DCFCE7')
            ec.alignment = Alignment(horizontal='center',vertical='center')
            # OBS
            ws.cell(last_row,8).font = Font(name='Calibri',size=8,color=C_MUTED,italic=True)

    output = io.BytesIO()
    hist_wb.save(output)
    output.seek(0)
    return output

# ── Routes ──
@app.route('/')
def index():
    if not session.get('user'):
        return render_template('login.html')
    return render_template('index.html', user=session['user'])

@app.route('/login', methods=['POST'])
def login():
    user = request.form.get('usuario','').strip()
    pwd  = request.form.get('password','').strip()
    if USUARIOS.get(user) == pwd:
        session['user'] = user
        return jsonify({'ok': True})
    return jsonify({'ok': False, 'msg': 'Usuario o contraseña incorrectos'})

@app.route('/logout')
def logout():
    session.clear()
    return render_template('login.html')

@app.route('/procesar', methods=['POST'])
def procesar_route():
    if not session.get('user'):
        return jsonify({'error': 'No autorizado'}), 401
    try:
        hist_file    = request.files.get('historial')
        cartola_file = request.files.get('cartola')
        if not hist_file or not cartola_file:
            return jsonify({'error': 'Faltan archivos'}), 400

        hist_wb    = load_workbook(hist_file)
        cartola_wb = load_workbook(cartola_file)

        res_ok, res_dif, res_res, res_efec, mes = procesar(hist_wb, cartola_wb)

        return jsonify({
            'ok':     res_ok,
            'dif':    res_dif,
            'res':    res_res,
            'efec':   res_efec,
            'mes':    mes,
            'total':  len(res_ok)+len(res_dif)+len(res_res)+len(res_efec)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/descargar', methods=['POST'])
def descargar_route():
    if not session.get('user'):
        return jsonify({'error': 'No autorizado'}), 401
    try:
        hist_file = request.files.get('historial')
        data      = json.loads(request.form.get('data','{}'))
        if not hist_file:
            return jsonify({'error': 'Falta historial'}), 400

        hist_wb = load_workbook(hist_file)
        res_ok  = data.get('ok',  [])
        res_dif = data.get('dif', [])
        res_res = data.get('res', [])
        res_efec= data.get('efec',[])

        output = generar_excel(hist_wb, res_ok, res_dif, res_res, res_efec)

        from datetime import date
        filename = f"Historial_2026_Actualizado_{date.today().isoformat()}.xlsx"
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    return 'OK', 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
