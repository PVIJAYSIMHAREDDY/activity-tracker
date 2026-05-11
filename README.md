# Daily Activity Tracker

A cross-platform desktop app to track your daily tasks, habits, work hours, goals, diet, and journal — with a personal nutrition calculator, AI coaching system, smart nutrition planning, and Excel export.

Built with **Flask + GTK/WebKit2** (native window, no browser needed) and packaged as a one-click installer for Linux, Windows, and macOS.

---

## Features

### ✅ Task Tracking
- Add, complete, and delete daily to-dos with priority and category
- Task completion rate shown on dashboard and weekly/monthly summaries

### 💪 Habits
- Define custom habits with target values and units
- Daily logging with calendar heatmap showing streaks

### ⏱ Work Hours
- Log work sessions with project name and notes
- Daily total and 7-day trend chart on dashboard

### 🎯 Goals
- Set long-term goals with progress bar tracking
- Update progress percentage at any time

### 🥗 Diet & Nutrition
- **7 meal slots**: Breakfast, Morning Snack, Lunch, Pre-Workout, Post-Workout, Dinner, Evening Snack
- Log **7 nutrients per meal**: Calories, Protein, Carbs, Fat, Fiber, Sugar, Sodium
- **Animated macro rings** (Calories / Protein / Carbs / Fat) showing % of daily goal
- **Color-coded micro bars** for Fiber, Sugar, Sodium (✅ on track / ⚠️ near limit / ❌ exceeded)
- Per-meal subtotals in the meals list

### 🏋️ AI Coaching System *(v1.2.0)*
- **Body Profile** — enter age, gender, height, weight, activity level, and goal
- **TDEE Calculator** — Mifflin-St Jeor formula → Total Daily Energy Expenditure
- **Auto Macro Targets** — protein / carbs / fat calculated per goal (cut / maintain / bulk / athletic)
- **Weight Log** — daily weigh-ins with 60-day trend line chart and ▼/▲ change indicator
- **Weekly Nutrition Score** (0–100) with animated gauge ring and tracking streak
- **Coaching Insights** — up to 6 personalised cards per week covering:
  - Calorie surplus/deficit (goal-aware logic for cut vs bulk)
  - Protein adequacy
  - Sugar and sodium limit warnings
  - Fiber deficit nudges
  - Tracking consistency
- **Diet Plans Library** — create unlimited named plans per month, switch active plan instantly
- **Monthly Review & Plan Suggestion** — analyses last 30 days of adherence + weight change and recommends adjusted calorie/protein targets for next month with one-click plan creation

### 📓 Journal / Notes
- Write journal entries, ideas, gratitude notes with mood emoji, tags, and pin support
- Search across all dates

### 📋 My Plan — Personal Nutrition Calculator *(v1.3.0)*
- **Metric / Imperial toggle** — switch between kg·cm and lbs·ft instantly
- **Visual activity cards** — tap your level (Sedentary → Athlete)
- **Goal cards** — Lose Weight / Maintain / Gain Muscle / Athletic Performance
- **Real-time TDEE** — Mifflin-St Jeor BMR × activity multiplier, updates as you type
- **Hero display** — daily calorie target with BMR, Maintenance, and Adjustment breakdown
- **SVG macro rings** — Protein / Carbs / Fat with grams and % of total energy
- **Goal Timeline** — estimated weeks to target weight + projected completion date
- **Save & Apply** — one click saves profile and pushes targets to active diet plan
- **Auto-start on login** — launches automatically on every session

### 📏 Body Measurements *(v1.3.0)*
- Log weight, body fat %, waist, chest, hips, neck, bicep, thigh, calf with date
- History table with **+/− delta** vs previous entry for every metric
- Form pre-fills from latest saved entry

### ✏️ Edit Any Entry *(v1.3.0)*
- Edit button on every **task** (text, category, priority)
- Edit button on every **work session** (project, hours, notes)
- Edit button on every **diet entry** (food name + all 7 nutrients)
- Edit button on every **goal** (title, target, unit, deadline, category)

### 📊 Dashboard & Reports
- Chart.js charts: task completion, habit heatmap, work hours, diet macros
- **Weekly Summary** — 7-day trend charts + aggregated stats
- **Monthly Summary** — full month breakdown with calendar view
- **Excel Export** — one-click `.xlsx` with 10 styled sheets (Tasks, Habits, Work, Goals, Diet, Summary, Journal, Dashboard, **Measurements**, **Progress Report**)

### 🎨 UI & UX
- **Dark mode** — toggleable, persisted across sessions
- **Mobile responsive** — scrollable tab strip, stacked forms, touch-friendly
- **Offline-first** — Chart.js bundled locally, no internet required

---

## What's New

### v1.3.0 — Personal Nutrition Calculator & Measurements
- **My Plan tab** — dedicated nutrition calculator with real-time TDEE, visual goal/activity cards, macro rings, goal timeline, metric/imperial toggle, one-click apply to diet plan
- **Body Measurements tracking** — log 9 metrics (weight, body fat %, waist, chest, hips, neck, bicep, thigh, calf) with delta history
- **Edit any entry** — tasks, work sessions, diet entries, and goals all have inline edit buttons
- **Excel: Measurements + Progress Report sheets** — full history with delta columns and first-vs-latest comparison
- **Auto-start on login** — desktop autostart entry installed automatically
- **Bug fixes** — resolved JS crash (duplicate `todayStr`), monthly tab date navigation, Chrome app-mode click handling

### v1.2.0 — AI Coaching & Smart Nutrition
- Full coaching system: body profile, TDEE, auto-macros, weight log
- Diet Plans Library with monthly plan management
- Rule-based coaching insights with priority ranking and action steps
- Monthly review engine: analyses adherence + weight change, suggests next plan

### v1.1.0 — Fitness Diet Tracking
- Macro progress rings + micro progress bars
- Fiber, Sugar, Sodium tracking per meal
- 7 meal slots including Pre/Post-Workout
- Bundled Chart.js (offline support), direct GTK+WebKit2 launcher

### v1.0.0 — Initial Release
- Task, habit, work, goals, diet, journal tracking
- Dashboard charts, weekly/monthly summaries
- Excel export, dark mode, mobile responsive
- Cross-platform installers (Linux, Windows, macOS)

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
git tag v1.3.0
git push origin v1.3.0
```

This automatically builds all three installers and creates a GitHub Release.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Frontend | Vanilla JS, Chart.js v4 (bundled), CSS custom properties |
| Desktop window | GTK3 + WebKit2 on Linux; pywebview on Windows/macOS |
| Nutrition engine | Mifflin-St Jeor BMR, goal-based macro calculation |
| Coaching engine | Rule-based insights with priority ranking |
| Data storage | JSON files |
| Excel export | openpyxl |
| Packaging | PyInstaller + dpkg / Inno Setup / create-dmg |
| CI/CD | GitHub Actions |

---

## License

MIT
