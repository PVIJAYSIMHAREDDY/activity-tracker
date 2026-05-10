"""
Desktop launcher for Daily Activity Tracker.
Starts Flask in a background thread, then opens a native window.

Launch strategy (Linux):
  1. Chromium / Chrome in --app mode  (most reliable, full JS support)
  2. GTK3 + WebKit2 direct            (lightweight, no browser needed)
  3. pywebview fallback               (Windows / macOS default)
"""
import threading, time, sys, os, socket, signal, shutil, subprocess, pwd

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


# ── Strategy 1: Chromium / Chrome --app mode ─────────────────────────────────
def _display_owner():
    """Return (username, xauthority_path) of the user who owns the X display."""
    display = os.environ.get('DISPLAY', ':0')
    num = display.lstrip(':').split('.')[0]
    try:
        uid = os.stat(f'/tmp/.X11-unix/X{num}').st_uid
        pw = pwd.getpwuid(uid)
        xauth = os.path.join(pw.pw_dir, '.Xauthority')
        return pw.pw_name, xauth
    except Exception:
        return None, None


def run_chromium(url):
    """Open the app in Chrome/Chromium's --app mode for a native-like window."""
    candidates = [
        'google-chrome', 'google-chrome-stable',
        'chromium-browser', 'chromium',
        'brave-browser',
    ]
    binary = next((b for b in candidates if shutil.which(b)), None)
    if not binary:
        return False

    env = os.environ.copy()

    # If running as root but display is owned by another user, launch Chrome
    # as that user so X11 input events are properly forwarded.
    if os.geteuid() == 0:
        display_user, xauth = _display_owner()
        if display_user and display_user != 'root':
            user_data_dir = f'/home/{display_user}/.config/activity-tracker-chrome'
            chrome_args = [
                binary,
                f'--app={url}',
                '--window-size=1200,780',
                f'--user-data-dir={user_data_dir}',
                '--disable-background-mode',
                '--disable-extensions',
                '--no-first-run',
                '--disable-default-apps',
            ]
            env['XAUTHORITY'] = xauth
            cmd = ['sudo', '-u', display_user,
                   'env',
                   f'DISPLAY={env.get("DISPLAY", ":0")}',
                   f'XAUTHORITY={xauth}',
                   ] + chrome_args
        else:
            chrome_args = [
                binary,
                f'--app={url}',
                '--window-size=1200,780',
                '--user-data-dir=/tmp/activity-tracker-chrome',
                '--disable-background-mode',
                '--disable-extensions',
                '--no-first-run',
                '--disable-default-apps',
                '--no-sandbox',
            ]
            cmd = chrome_args
    else:
        chrome_args = [
            binary,
            f'--app={url}',
            '--window-size=1200,780',
            '--user-data-dir=/tmp/activity-tracker-chrome',
            '--disable-background-mode',
            '--disable-extensions',
            '--no-first-run',
            '--disable-default-apps',
        ]
        cmd = chrome_args

    proc = subprocess.Popen(cmd, env=env)
    proc.wait()
    os.kill(os.getpid(), signal.SIGTERM)
    return True


# ── Strategy 2: GTK3 + WebKit2 direct ────────────────────────────────────────
def run_gtk(url):
    # Force X11 backend — fixes Wayland input-forwarding issues with WebKit2
    os.environ['GDK_BACKEND'] = 'x11'

    import gi
    gi.require_version('Gtk', '3.0')
    try:
        gi.require_version('WebKit2', '4.1')
    except ValueError:
        gi.require_version('WebKit2', '4.0')
    from gi.repository import Gtk, WebKit2, GLib

    # Disable process sandbox so WebKit2 can reach localhost
    ctx = WebKit2.WebContext.get_default()
    try:
        ctx.set_sandbox_enabled(False)
    except AttributeError:
        pass

    win = Gtk.Window()
    win.set_title('Daily Activity Tracker')
    win.set_default_size(1200, 780)
    win.set_size_request(360, 600)
    win.set_position(Gtk.WindowPosition.CENTER)

    webview = WebKit2.WebView.new_with_context(ctx)

    s = webview.get_settings()
    s.set_enable_javascript(True)
    s.set_enable_javascript_markup(True)
    s.set_javascript_can_access_clipboard(True)
    s.set_javascript_can_open_windows_automatically(False)
    s.set_allow_universal_access_from_file_urls(True)
    s.set_allow_file_access_from_file_urls(True)
    s.set_enable_developer_extras(False)
    s.set_enable_page_cache(False)

    def on_load_changed(wv, event):
        if event == WebKit2.LoadEvent.FINISHED:
            # Give the renderer a moment then force focus into the view
            GLib.timeout_add(150, lambda: (wv.grab_focus(), False)[1])

    webview.connect('load-changed', on_load_changed)
    webview.load_uri(url)

    win.add(webview)
    win.show_all()
    win.present()
    webview.grab_focus()

    def on_closed():
        os.kill(os.getpid(), signal.SIGTERM)

    win.connect('destroy', lambda *_: on_closed())
    Gtk.main()


# ── Strategy 3: pywebview (Windows / macOS) ───────────────────────────────────
def run_pywebview(url):
    import webview
    window = webview.create_window(
        title='Daily Activity Tracker',
        url=url,
        width=1200, height=780,
        min_size=(360, 600),
        resizable=True,
        text_select=True,
    )
    window.events.closed += lambda: os.kill(os.getpid(), signal.SIGTERM)
    webview.start(debug=False)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not port_in_use(PORT):
        t = threading.Thread(target=start_flask, daemon=True)
        t.start()
        if not wait_for_server(PORT):
            print("ERROR: Flask did not start.", file=sys.stderr)
            sys.exit(1)

    url = f'http://127.0.0.1:{PORT}'

    if sys.platform.startswith('linux'):
        # 1. Try Chromium app-mode (best input reliability)
        if run_chromium(url):
            return
        # 2. GTK + WebKit2 with X11 + focus fixes
        try:
            run_gtk(url)
            return
        except Exception as e:
            print(f"GTK failed ({e}), trying pywebview…", file=sys.stderr)

    # 3. pywebview (Windows / macOS, or Linux last-resort)
    run_pywebview(url)


if __name__ == '__main__':
    main()
