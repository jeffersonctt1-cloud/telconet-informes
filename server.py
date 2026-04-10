import os
from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/')
def index():
    return jsonify({'status': 'ok', 'message': 'Flask funcionando'})

@app.route('/test')
def test():
    import sys
    base = os.path.dirname(os.path.abspath(__file__))
    files = os.listdir(base)
    return jsonify({'files': files, 'python': sys.version})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(host='0.0.0.0', port=port)