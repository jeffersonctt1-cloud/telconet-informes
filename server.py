#!/usr/bin/env python3
import os, sys, json, base64, shutil, traceback
from flask import Flask, request, jsonify, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
T = os.path.join(BASE_DIR, 'tmp')
os.makedirs(T, exist_ok=True)

app = Flask(__name__)

_initialized = False

def init_files():
    global _initialized
    if _initialized:
        return None
    errors = []
    files_to_copy = [
        ('fill_template.py',             'fill_template.py'),
        ('fill_template_gis.py',         'fill_template_gis.py'),
        ('fill_template_opu.py',         'fill_template_opu.py'),
        ('FOR_OPU_06.docx',              'opu.docx'),
        ('INFORME_TEMPLATE_LATLON.docx', 't.docx'),
        ('FOR_GIS_08_.docx',             'gis.docx'),
        ('index.html',                   'index.html'),
    ]
    for src_name, dst_name in files_to_copy:
        src = os.path.join(BASE_DIR, src_name)
        dst = os.path.join(T, dst_name)
        if os.path.exists(src):
            try:
                shutil.copy2(src, dst)
            except Exception as e:
                errors.append(f'copy {src_name}: {e}')
        else:
            errors.append(f'NOT FOUND: {src_name}')
    if T not in sys.path:
        sys.path.insert(0, T)
    _initialized = True
    return errors if errors else None

@app.route('/test')
def test():
    errors = init_files()
    files = os.listdir(T) if os.path.exists(T) else []
    base_files = os.listdir(BASE_DIR)
    return jsonify({
        'status': 'ok',
        'T': T,
        'BASE_DIR': BASE_DIR,
        'tmp_files': files,
        'base_files': base_files,
        'init_errors': errors
    })

@app.route('/')
def index():
    errors = init_files()
    if errors:
        return jsonify({'startup_errors': errors}), 500
    return send_from_directory(T, 'index.html')

@app.route('/generar-pdf', methods=['POST', 'OPTIONS'])
def generar_pdf():
    if request.method == 'OPTIONS':
        return '', 200
    errors = init_files()
    if errors:
        return jsonify({'error': 'Init failed', 'details': errors}), 500
    try:
        data = request.get_json()
        dfile    = os.path.join(T, 'd.json')
        docx_out = os.path.join(T, 'i.docx')
        with open(dfile, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        if os.path.exists(docx_out):
            os.remove(docx_out)
        import fill_template as ft
        import importlib; importlib.reload(ft)
        ft.fill_template(dfile, docx_out, os.path.join(T, 't.docx'))
        if not os.path.exists(docx_out):
            return jsonify({'error': 'No se generó el .docx'}), 500
        with open(docx_out, 'rb') as f:
            data_b64 = base64.b64encode(f.read()).decode()
        fecha = str(data.get('fecha', 'x')).replace('/', '-')[:10]
        tarea = str(data.get('tarea', 'inf'))[:8].replace(' ', '_')
        fname = f'TCN_{tarea}_{fecha}.docx'
        return jsonify({'file': data_b64, 'filename': fname,
                        'ext': 'docx',
                        'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'warning': ''})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/generar-opu', methods=['POST', 'OPTIONS'])
def generar_opu():
    if request.method == 'OPTIONS':
        return '', 200
    errors = init_files()
    if errors:
        return jsonify({'error': 'Init failed', 'details': errors}), 500
    try:
        data = request.get_json()
        shutil.copy2(os.path.join(BASE_DIR, 'FOR_OPU_06.docx'),      os.path.join(T, 'opu.docx'))
        shutil.copy2(os.path.join(BASE_DIR, 'fill_template_opu.py'), os.path.join(T, 'fill_template_opu.py'))
        dfile    = os.path.join(T, 'd.json')
        docx_out = os.path.join(T, 'o.docx')
        with open(dfile, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        if os.path.exists(docx_out):
            os.remove(docx_out)
        import fill_template_opu as fto
        import importlib; importlib.reload(fto)
        fto.fill_template_opu(dfile, docx_out, os.path.join(T, 'opu.docx'))
        if not os.path.exists(docx_out):
            return jsonify({'error': 'No se generó el .docx'}), 500
        with open(docx_out, 'rb') as f:
            data_b64 = base64.b64encode(f.read()).decode()
        fecha   = str(data.get('fecha', 'x')).replace('/', '-').replace(':', '-').replace(' ', '_')[:19]
        cliente = str(data.get('cliente', 'opu'))[:10].replace(' ', '_')
        fname   = f'OPU_{cliente}_{fecha}.docx'
        return jsonify({'file': data_b64, 'filename': fname,
                        'ext': 'docx',
                        'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'warning': ''})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/generar-gis', methods=['POST', 'OPTIONS'])
def generar_gis():
    if request.method == 'OPTIONS':
        return '', 200
    errors = init_files()
    if errors:
        return jsonify({'error': 'Init failed', 'details': errors}), 500
    try:
        data = request.get_json()
        dfile    = os.path.join(T, 'd.json')
        docx_out = os.path.join(T, 'g.docx')
        with open(dfile, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        if os.path.exists(docx_out):
            os.remove(docx_out)
        import fill_template_gis as ftg
        import importlib; importlib.reload(ftg)
        ftg.fill_template_gis(dfile, docx_out, os.path.join(T, 'gis.docx'))
        if not os.path.exists(docx_out):
            return jsonify({'error': 'No se generó el .docx'}), 500
        with open(docx_out, 'rb') as f:
            data_b64 = base64.b64encode(f.read()).decode()
        fecha = str(data.get('fecha', 'x')).replace('/', '-')[:10]
        punto = str(data.get('nombre_punto', 'gis'))[:8].replace(' ', '_')
        fname = f'GIS_{punto}_{fecha}.docx'
        return jsonify({'file': data_b64, 'filename': fname,
                        'ext': 'docx',
                        'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'warning': ''})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.after_request
def cors(r):
    r.headers['Access-Control-Allow-Origin'] = '*'
    r.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    r.headers['Access-Control-Allow-Methods'] = 'POST,GET,OPTIONS'
    return r

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(host='0.0.0.0', port=port, debug=False)
