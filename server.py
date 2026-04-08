#!/usr/bin/env python3
import os, sys, json, base64, threading, webbrowser, shutil, traceback
from flask import Flask, request, jsonify, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
T = 'C:\\t'
os.makedirs(T, exist_ok=True)

# Copy everything needed to C:\t at startup
shutil.copy2(os.path.join(BASE_DIR, 'fill_template.py'), T + '\\fill_template.py')
shutil.copy2(os.path.join(BASE_DIR, 'fill_template_gis.py'), T + '\\fill_template_gis.py')
shutil.copy2(os.path.join(BASE_DIR, 'fill_template_opu.py'), T + '\\fill_template_opu.py')
shutil.copy2(os.path.join(BASE_DIR, 'FOR_OPU_06.docx'), T + '\\opu.docx')
shutil.copy2(os.path.join(BASE_DIR, 'INFORME_TEMPLATE_LATLON.docx'), T + '\\t.docx')
shutil.copy2(os.path.join(BASE_DIR, 'FOR_GIS_08_.docx'), T + '\\gis.docx')
shutil.copy2(os.path.join(BASE_DIR, 'index.html'), T + '\\index.html')

sys.path.insert(0, T)

app = Flask(__name__)

@app.route('/test')
def test():
    import os
    files = os.listdir(T) if os.path.exists(T) else []
    return jsonify({'status':'ok','T':T,'files':files,'BASE_DIR':BASE_DIR})

@app.route('/')
def index():
    return send_from_directory(T, 'index.html')

@app.route('/generar-pdf', methods=['POST','OPTIONS'])
def generar_pdf():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        data = request.get_json()

        dfile    = T + '\\d.json'
        docx_out = T + '\\i.docx'
        pdf_out  = T + '\\i.pdf'

        with open(dfile, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)

        for p in [docx_out, pdf_out]:
            if os.path.exists(p): os.remove(p)

        import fill_template as ft
        import importlib; importlib.reload(ft)
        ft.fill_template(dfile, docx_out, T + '\\t.docx')

        if not os.path.exists(docx_out):
            return jsonify({'error': 'No se generó el .docx'}), 500

        # Return DOCX directly
        with open(docx_out, 'rb') as f:
            data_b64 = base64.b64encode(f.read()).decode()

        fecha = str(data.get('fecha','x')).replace('/','-')[:10]
        tarea = str(data.get('tarea','inf'))[:8].replace(' ','_')
        fname = f'TCN_{tarea}_{fecha}.docx'

        return jsonify({'file': data_b64, 'filename': fname,
                        'ext': 'docx',
                        'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'warning': ''})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/generar-opu', methods=['POST','OPTIONS'])
def generar_opu():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        data = request.get_json()
        os.makedirs(T, exist_ok=True)
        # Re-copy template and script on each request to ensure fresh versions
        shutil.copy2(os.path.join(BASE_DIR, 'FOR_OPU_06.docx'), T + '\\opu.docx')
        shutil.copy2(os.path.join(BASE_DIR, 'fill_template_opu.py'), T + '\\fill_template_opu.py')
        dfile    = T + '\\d.json'
        docx_out = T + '\\o.docx'
        with open(dfile, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        if os.path.exists(docx_out): os.remove(docx_out)
        sys.path.insert(0, T)
        import fill_template_opu as fto
        import importlib; importlib.reload(fto)
        fto.fill_template_opu(dfile, docx_out, T + '\\opu.docx')
        if not os.path.exists(docx_out):
            return jsonify({'error': 'No se generó el .docx'}), 500
        with open(docx_out, 'rb') as f:
            data_b64 = base64.b64encode(f.read()).decode()
        fecha = str(data.get('fecha','x')).replace('/','-').replace(':','-').replace(' ','_')[:19]
        cliente = str(data.get('cliente','opu'))[:10].replace(' ','_')
        fname = f'OPU_{cliente}_{fecha}.docx'
        return jsonify({'file': data_b64, 'filename': fname,
                        'ext': 'docx',
                        'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'warning': ''})
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/generar-gis', methods=['POST','OPTIONS'])
def generar_gis():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        data = request.get_json()
        os.makedirs(T, exist_ok=True)
        dfile    = T + '\\d.json'
        docx_out = T + '\\g.docx'
        with open(dfile, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        if os.path.exists(docx_out): os.remove(docx_out)

        sys.path.insert(0, T)
        import fill_template_gis as ftg
        import importlib; importlib.reload(ftg)
        ftg.fill_template_gis(dfile, docx_out, T + '\\gis.docx')

        if not os.path.exists(docx_out):
            return jsonify({'error': 'No se generó el .docx'}), 500

        with open(docx_out, 'rb') as f:
            data_b64 = base64.b64encode(f.read()).decode()

        fecha = str(data.get('fecha','x')).replace('/','-')[:10]
        punto = str(data.get('nombre_punto','gis'))[:8].replace(' ','_')
        fname = f'GIS_{punto}_{fecha}.docx'
        return jsonify({'file': data_b64, 'filename': fname,
                        'ext': 'docx',
                        'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'warning': ''})
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.after_request
def cors(r):
    r.headers['Access-Control-Allow-Origin'] = '*'
    r.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    r.headers['Access-Control-Allow-Methods'] = 'POST,GET,OPTIONS'
    return r

def get_local_ip():
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return '0.0.0.0'

if __name__ == '__main__':
    print('\n' + '='*60)
    print('  TELCONET — Sistema de Informes')
    print('='*60)
    print('  Local:    http://localhost:5050')
    print('  Red WiFi: http://' + get_local_ip() + ':5050')
    print()
    print('  Para acceso desde CUALQUIER red (internet):')
    print('  1. Descarga ngrok: https://ngrok.com/download')
    print('  2. En otra terminal ejecuta: ngrok http 5050')
    print('  3. Usa la URL https://xxxx.ngrok.io que aparece')
    print('='*60)
    print('  Ctrl+C para detener')
    print('='*60 + '\n')
    threading.Timer(1.2, lambda: webbrowser.open('http://localhost:5050')).start()
    app.run(host='0.0.0.0', port=5050, debug=False)
