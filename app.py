import os
import sys
import time
import threading
import subprocess
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, render_template, jsonify, send_file

app = Flask(__name__)

_lock = threading.Lock()
_state = {
    'running': False,
    'finished': False,
    'error': None,
    'output_file': None,
    'logs': [],
}

TIMEOUT_SECS = 30 * 60  # 30 minutos


def _run():
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"/tmp/catalogo_{fecha}.xlsx"
    script = str(Path(__file__).parent / 'procesar_catalogo.py')

    def log(line):
        with _lock:
            _state['logs'].append(line)

    try:
        proc = subprocess.Popen(
            [sys.executable, script, '--descargar', '--output', output_file],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd='/tmp',
            env=os.environ.copy(),
        )

        start = time.time()
        for line in iter(proc.stdout.readline, ''):
            if time.time() - start > TIMEOUT_SECS:
                proc.kill()
                log('⏱️ Timeout de 30 minutos alcanzado. El proceso fue detenido.')
                with _lock:
                    _state['error'] = 'Timeout'
                break
            log(line.rstrip())

        proc.wait()

        if proc.returncode == 0 and Path(output_file).exists():
            with _lock:
                _state['output_file'] = output_file
        else:
            with _lock:
                if not _state['error']:
                    _state['error'] = f'El proceso terminó con código {proc.returncode}.'

    except Exception as exc:
        with _lock:
            _state['error'] = str(exc)

    finally:
        with _lock:
            _state['finished'] = True
            _state['running'] = False


@app.route('/')
def index():
    with _lock:
        running = _state['running']
        has_file = bool(_state['output_file'] and Path(_state['output_file']).exists())
    return render_template('index.html', running=running, has_file=has_file)


@app.route('/health')
def health():
    return 'OK', 200


@app.route('/generar', methods=['POST'])
def generar():
    with _lock:
        if _state['running']:
            return jsonify({'error': 'already_running'}), 409
        _state['running'] = True
        _state['finished'] = False
        _state['error'] = None
        _state['output_file'] = None
        _state['logs'] = []

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({'ok': True})


@app.route('/stream')
def stream():
    def generate():
        sent = 0
        while True:
            with _lock:
                current_logs = list(_state['logs'])
                finished = _state['finished']
                error = _state['error']
                output_file = _state['output_file']

            while sent < len(current_logs):
                line = current_logs[sent].replace('\n', ' ')
                yield f"data: {line}\n\n"
                sent += 1

            if finished:
                if output_file and Path(output_file).exists():
                    yield "event: done\ndata: ok\n\n"
                else:
                    msg = (error or 'Error desconocido').replace('\n', ' ')
                    yield f"event: job_error\ndata: {msg}\n\n"
                return

            yield ": keepalive\n\n"
            time.sleep(0.5)

    return Response(
        generate(),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',
            'Connection': 'keep-alive',
        },
    )


@app.route('/descargar')
def descargar():
    with _lock:
        f = _state['output_file']
    if not f or not Path(f).exists():
        return 'No hay fichero disponible', 404
    return send_file(f, as_attachment=True, download_name='catalogo_bicimarket.xlsx')


if __name__ == '__main__':
    app.run(debug=True, threaded=True)
