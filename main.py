"""
Desktop launcher for Daily Activity Tracker.
Starts Flask in a background thread, then opens a native window via pywebview.
"""
import threading
import time
import sys
import os
import socket
import signal

# ── Paths ──────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
os.chdir(BASE_DIR)

PORT = 5050

# ── Check port availability ────────────────────────────────────────────────
def port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def wait_for_server(port, timeout=15):
    start = time.time()
    while time.time() - start < timeout:
        if port_in_use(port):
            return True
        time.sleep(0.1)
    return False

# ── Flask startup ──────────────────────────────────────────────────────────
def start_flask():
    from app import app
    # Suppress Flask startup banner
    import logging
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)
    app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False, threaded=True)

# ── Main ───────────────────────────────────────────────────────────────────
def main():
    already_running = port_in_use(PORT)

    if not already_running:
        flask_thread = threading.Thread(target=start_flask, daemon=True)
        flask_thread.start()
        if not wait_for_server(PORT):
            print("ERROR: Flask server did not start in time.", file=sys.stderr)
            sys.exit(1)

    import webview

    url = f'http://127.0.0.1:{PORT}'

    window = webview.create_window(
        title='Daily Activity Tracker',
        url=url,
        width=1200,
        height=780,
        min_size=(360, 600),
        resizable=True,
        text_select=True,
    )

    # Clean shutdown when window is closed
    def on_closed():
        os.kill(os.getpid(), signal.SIGTERM)

    window.events.closed += on_closed

    webview.start(debug=False)

if __name__ == '__main__':
    main()
