# Daily Activity Tracker

A cross-platform desktop app to track your daily tasks, habits, work hours, goals, diet, and journal — with Excel export and rich dashboard charts.

Built with **Flask + GTK/WebKit2** (native window, no browser needed) and packaged as a one-click installer for Linux, Windows, and macOS.

---

## Features

- **Tasks** — add, complete, and delete daily to-dos with priority and category
- **Habits** — define custom habits and log them daily with a calendar heatmap
- **Work Hours** — log work sessions with notes and see daily totals
- **Goals** — set long-term goals and track progress with a visual bar
- **Diet & Nutrition** *(v1.1.0)* — full fitness diet tracking:
  - Set a daily nutrition plan: calories, protein, carbs, fat, fiber, sugar limit, sodium limit
  - Animated macro rings (Calories / Protein / Carbs / Fat) showing % of daily goal
  - Color-coded micro bars for Fiber, Sugar, Sodium (✅ on track / ⚠️ near limit / ❌ exceeded)
  - Log meals across 7 slots: Breakfast, Morning Snack, Lunch, Pre-Workout, Post-Workout, Dinner, Evening Snack
  - Each food entry stores calories, protein, carbs, fat, fiber, sugar, and sodium
  - Per-meal subtotals in the meal list header
- **Journal / Notes** — write journal entries, ideas, gratitude notes with mood, tags, and pin support
- **Dashboard** — Chart.js charts for tasks, habits, work hours, and goals at a glance
- **Weekly & Monthly Summary** — trend charts and aggregated stats across any week or month
- **Excel Export** — one-click export to `.xlsx` with 8 styled sheets (including Journal)
- **Dark Mode** — toggleable, persisted across sessions
- **Mobile Responsive** — works on small screens too

---

## What's New in v1.1.0

- **Fitness diet system** — set calorie and macro/micro targets, track progress with rings and bars
- **7 meal slots** — Pre-Workout and Post-Workout slots added alongside the standard meals
- **Fiber, Sugar & Sodium tracking** — log and monitor all key micro-nutrients per meal
- **Linux desktop fix** — switched from pywebview to direct GTK+WebKit2 for reliable click handling
- **Offline Chart.js** — bundled locally so the app works without internet access

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

**Requirements:** Python 3.8+

```bash
git clone https://github.com/PVIJAYSIMHAREDDY/activity-tracker.git
cd activity-tracker

# Install dependencies (Linux)
sudo apt-get install python3-flask python3-webview python3-openpyxl
# or via pip (all platforms):
pip install flask pywebview openpyxl

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
git tag v1.1.0
git push origin v1.1.0
```

This automatically builds all three installers and creates a GitHub Release.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Frontend | Vanilla JS, Chart.js v4 (bundled), CSS custom properties |
| Desktop window | GTK3 + WebKit2 on Linux; pywebview on Windows/macOS |
| Data storage | JSON files |
| Excel export | openpyxl |
| Packaging | PyInstaller + dpkg / Inno Setup / create-dmg |
| CI/CD | GitHub Actions |

---

## License

MIT
