# Daily Activity Tracker

A cross-platform desktop app to track your daily tasks, habits, work hours, goals, diet, and journal — with Excel export and rich dashboard charts.

Built with **Flask + PyWebView** (native window, no browser needed) and packaged as a one-click installer for Linux, Windows, and macOS.

---

## Features

- **Tasks** — add, complete, and delete daily to-dos
- **Habits** — define custom habits and log them daily with a calendar heatmap
- **Work Hours** — log work sessions with notes and see daily totals
- **Goals** — set long-term goals and track progress with a visual bar
- **Diet** — log meals with calories, protein, carbs, and fat; see daily macro totals
- **Journal / Notes** — write journal entries, ideas, gratitude notes with mood, tags, and pin support
- **Dashboard** — Chart.js charts for tasks, habits, work hours, and goals at a glance
- **Weekly & Monthly Summary** — trend charts and aggregated stats across any week or month
- **Excel Export** — one-click export to `.xlsx` with 8 styled sheets (including Journal)
- **Dark Mode** — toggleable, persisted across sessions
- **Mobile Responsive** — works on small screens too

---

## Installation

### Linux (Ubuntu / Debian)

Download `activity-tracker_*_all.deb` from the [Releases](../../releases) page, then:

```bash
sudo dpkg -i activity-tracker_*.deb
sudo apt-get install -f
```

Launch from your app menu or run `activity-tracker` in a terminal.

### Windows

Download `ActivityTracker_Setup.exe` from the [Releases](../../releases) page and run the installer. Optionally installs a desktop shortcut and startup entry.

> Requires [Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) (already present on Windows 11 / Office installs).

### macOS

Download `ActivityTracker_macOS.dmg` from the [Releases](../../releases) page, open it, and drag **ActivityTracker.app** to your Applications folder.

---

## Run from Source

**Requirements:** Python 3.8+, pip

```bash
git clone https://github.com/PVIJAYSIMHAREDDY/activity-tracker.git
cd activity-tracker

# Install dependencies
sudo apt-get install python3-flask python3-webview python3-openpyxl   # Linux (apt)
# or: pip install flask pywebview openpyxl

# Run as desktop app
python3 main.py

# Or run as plain web app (opens in browser at http://localhost:5050)
python3 app.py
```

Data is stored as JSON files in `~/.activity-tracker/data/` (desktop mode) or `data/` (source mode).

---

## Building Installers

| Platform | Script |
|---|---|
| Linux `.deb` | `bash build/build_linux.sh` |
| Windows `.exe` | `build\build_windows.bat` |
| macOS `.dmg` | `bash build/build_mac.sh` |

Or push a git tag to trigger the GitHub Actions release workflow:

```bash
git tag v1.0.0
git push origin v1.0.0
```

This automatically builds all three installers and creates a GitHub Release.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Frontend | Vanilla JS, Chart.js v4, CSS custom properties |
| Desktop window | PyWebView (GTK/WebKit2 on Linux, WinForms/Edge on Windows, Cocoa on macOS) |
| Data storage | JSON files |
| Excel export | openpyxl |
| Packaging | PyInstaller + dpkg / Inno Setup / create-dmg |
| CI/CD | GitHub Actions |

---

## License

MIT
