"""
Desktop launcher for Daily Activity Tracker.
Starts Flask in a background thread, then opens a native GTK+WebKit2 window.
Falls back to pywebview on non-Linux platforms.
"""
import threading
import time
import sys
import os
import socket
import signal

PORT = 5050

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
os.chdir(BASE_DIR)


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


def start_flask():
    from app import app
    import logging
    logging.getLogger('werkzeug').setLevel(logging.ERROR)
    app.run(host='127.0.0.1', port=PORT, debug=False,
            use_reloader=False, threaded=True)


def run_gtk(url):
    import gi
    gi.require_version('Gtk', '3.0')
    try:
        gi.require_version('WebKit2', '4.1')
    except ValueError:
        gi.require_version('WebKit2', '4.0')
    from gi.repository import Gtk, WebKit2

    win = Gtk.Window()
    win.set_title('Daily Activity Tracker')
    win.set_default_size(1200, 780)
    win.set_size_request(360, 600)

    webview = WebKit2.WebView()

    # Explicitly enable everything needed for a fully interactive app
    s = webview.get_settings()
    s.set_enable_javascript(True)
    s.set_enable_javascript_markup(True)
    s.set_javascript_can_access_clipboard(True)
    s.set_javascript_can_open_windows_automatically(False)
    s.set_allow_universal_access_from_file_urls(True)
    s.set_allow_file_access_from_file_urls(True)
    s.set_enable_developer_extras(False)

    webview.load_uri(url)
    win.add(webview)
    win.show_all()
    win.connect('destroy', Gtk.main_quit)
    Gtk.main()


def run_pywebview(url):
    import webview
    window = webview.create_window(
        title='Daily Activity Tracker',
        url=url,
        width=1200,
        height=780,
        min_size=(360, 600),
        resizable=True,
        text_select=True,
    )

    def on_closed():
        os.kill(os.getpid(), signal.SIGTERM)

    window.events.closed += on_closed
    webview.start(debug=False)


def main():
    already_running = port_in_use(PORT)
    if not already_running:
        flask_thread = threading.Thread(target=start_flask, daemon=True)
        flask_thread.start()
        if not wait_for_server(PORT):
            print("ERROR: Flask server did not start in time.", file=sys.stderr)
            sys.exit(1)

    url = f'http://127.0.0.1:{PORT}'

    if sys.platform.startswith('linux'):
        try:
            run_gtk(url)
            return
        except Exception as e:
            print(f"GTK launch failed ({e}), falling back to pywebview", file=sys.stderr)

    run_pywebview(url)


if __name__ == '__main__':
    main()
