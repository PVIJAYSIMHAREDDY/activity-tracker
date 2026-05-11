# Daily Activity Tracker

An elite AI-powered fitness and nutrition coaching app — track tasks, habits, work hours, goals, diet, workouts, and wellness with a **real-time TDEE calculator**, **adaptive AI coaching dashboard**, body measurements tracking, and Excel export.

Built with **Flask + Chrome app-mode** (native window, no browser needed) and packaged as a one-click installer for Linux, Windows, and macOS.

---

## Features

### 🤖 AI Coaching Dashboard *(v1.4.0)*
- **Overall fitness score** (0–99) with animated SVG ring and letter grade (S / A / B / C / D)
- **6-stat overview** — diet adherence %, recovery score, sessions/week, weight trend (kg/wk), target kcal, target protein
- **Alert banners** — danger / warning / info for plateaus, overeating, critical recovery, overtraining
- **Today's Focus** — 3 personalised priority action cards generated from your live data and goal
- **Insight cards** with High / Medium priority chips: plateau detection, protein gaps, undereating, recovery alerts, goal achievements
- **AI adjustment panel** — recommended calorie and protein deltas with one-click context
- Refreshes automatically when you log a workout or save wellness; manual refresh button

### 🏋️ Training Log *(v1.4.0)*
- **6 workout types**: Strength, Cardio, HIIT, Yoga/Flex, Sport, Other — select with visual type cards
- Per-session **exercise tracker**: name, sets, reps, weight (kg) with add/remove rows
- Duration, date, and free-text notes per session
- **Weekly training summary** — total sessions, total minutes, total sets, distinct types
- Full scrollable **workout history** with exercise breakdown and delete support

### 🌟 Wellness Tracking *(v1.4.0)*
- **Daily check-in** with 5 metrics via visual sliders: sleep hours (3–12 h), sleep quality, energy, stress, muscle soreness
- Pre-fills sliders from today's saved entry for frictionless updates
- Recent **5-day history strip** with colour-coded mini progress bars
- Wellness data feeds directly into the AI recovery score and coaching insights

### 📋 My Plan — Personal Nutrition Calculator *(v1.3.0)*
- **Metric / Imperial toggle** — switch between kg·cm and lbs·ft with one click
- **Visual activity cards** — Sedentary, Light, Moderate, Very Active, Athlete
- **Goal cards** — Lose Weight / Maintain / Gain Muscle / Athletic Performance
- **Real-time TDEE** — Mifflin-St Jeor BMR × activity multiplier, updates as you type
- **Hero display** — daily calorie target with BMR, Maintenance, and Adjustment at a glance
- **SVG macro rings** — Protein / Carbs / Fat with grams and % of total energy
- **Micro-nutrient targets** — Fiber, Sugar, and Sodium limits
- **Goal Timeline** — estimated weeks to target weight + projected completion date
- **Save & Apply** — one click saves profile and pushes targets to the active diet plan
- **Create Plan** — save calculated targets as a named plan in the Plans Library

### 📏 Body Measurements *(v1.3.0)*
- Log **9 metrics** with date: weight, body fat %, waist, chest, hips, neck, bicep, thigh, calf
- History table shows **+/− delta** vs previous entry for every metric
- Form pre-fills from latest saved entry for quick updates
- Full **Measurements sheet** in Excel export with delta columns

### 🥗 Diet & Nutrition
- **7 meal slots**: Breakfast, Morning Snack, Lunch, Pre-Workout, Post-Workout, Dinner, Evening Snack
- Log **7 nutrients per meal**: Calories, Protein, Carbs, Fat, Fiber, Sugar, Sodium
- **Edit** any food entry inline with the ✏️ button
- **Animated macro rings** (Calories / Protein / Carbs / Fat) showing % of daily goal
- **Colour-coded micro bars** for Fiber, Sugar, Sodium (✅ on track / ⚠️ near limit / ❌ exceeded)
- Per-meal subtotals in the meals list

### 📊 Coach Tab — Body Profile & Plans *(v1.2.0)*
- **Body Profile** — age, gender, height, weight, activity level, goal
- **TDEE Calculator** — Mifflin-St Jeor formula → Total Daily Energy Expenditure
- **Auto Macro Targets** — protein / carbs / fat per goal (cut / maintain / bulk / athletic)
- **Weight Log** — daily weigh-ins with 60-day trend line chart
- **Diet Plans Library** — create unlimited named plans per month, switch active plan instantly
- **Monthly Review & Plan Suggestion** — analyses 30-day adherence + weight change, suggests next month's plan with one-click creation

### ✅ Task Tracking
- Add, complete, and delete daily to-dos with priority and category
- **Edit** any task inline (text, category, priority) with the ✏️ button
- Task completion rate on dashboard and weekly/monthly summaries

### 💪 Habits
- Define custom habits with target values and units
- Daily logging with progress bar and percentage indicator
- Calendar heatmap showing streaks

### ⏱ Work Hours
- Log work sessions with project name and notes
- **Edit** any session inline (project, hours, notes) with the ✏️ button
- Daily total and 7-day trend chart on dashboard

### 🎯 Goals
- Set long-term goals with progress bar tracking
- **Edit** any goal inline (title, target, unit, deadline, category) with the ✏️ button
- Update current progress at any time via the inline number input

### 📓 Journal / Notes
- Write journal entries, ideas, gratitude notes with mood emoji, tags, and pin support
- Search across all dates

### 📊 Dashboard & Reports
- Chart.js charts: task completion, habit heatmap, work hours, diet macros
- **Weekly Summary** — 7-day trend charts + aggregated stats
- **Monthly Summary** — full month breakdown with calendar view
- **Excel Export** — one-click `.xlsx` with 10 styled sheets:
  Tasks, Habits, Work, Goals, Diet, 7-Day Summary, Journal, Dashboard, Measurements, Progress Report

### 🎨 UI & UX
- **Dark mode** — toggleable, persisted across sessions
- **Mobile responsive** — scrollable tab strip, stacked forms, touch-friendly
- **Offline-first** — Chart.js bundled locally, no internet required
- **Auto-start on login** — launches automatically on every session (Linux)

---

## What's New

### v1.4.0 — Elite AI Coaching Platform
- **AI Coaching Dashboard** — animated score ring (0–99), letter grade (S–D), 6-stat overview, refresh on data change
- **Training tab** — workout logging with 6 type cards, exercise tracker (sets/reps/weight), weekly summary stats
- **Daily Wellness Check-in** — 5 slider metrics, pre-fills from saved entry, 5-day history strip; feeds AI recovery score
- **AI Insight Engine** — plateau detection, adherence/undereating/overeating alerts, protein gap warnings, overtraining detection, body composition tracking, goal achievement recognition
- **Today's Focus** — 3 priority action cards generated live from your personal data and goal
- **Alert Banners** — danger/warning/info for critical issues requiring immediate attention
- **AI Calorie/Protein Adjustment Panel** — data-driven deltas recommended on the dashboard
- New backend route: `GET /api/workouts/all`, `GET /api/wellness/history`, `GET /api/ai/analysis`

### v1.3.0 — Personal Nutrition Calculator & Measurements
- **My Plan tab** — real-time TDEE, visual activity/goal cards, SVG macro rings, goal timeline, metric/imperial toggle, one-click apply to diet plan
- **Body Measurements** — log 9 body metrics with date; history table shows +/− deltas
- **Edit any entry** — ✏️ buttons on every task, work session, diet entry, and goal
- **Excel: 2 new sheets** — Measurements (delta columns) + Progress Report (first vs latest comparison)
- **Auto-start on login** — desktop autostart entry installed automatically on Linux

### v1.2.0 — AI Coaching & Smart Nutrition
- Full coaching system: body profile, TDEE, auto-macros, weight log
- Diet Plans Library with monthly plan management
- Rule-based coaching insights with priority ranking and action steps
- Monthly review engine: analyses adherence + weight change, suggests next plan

### v1.1.0 — Fitness Diet Tracking
- Macro progress rings + micro progress bars
- Fiber, Sugar, Sodium tracking per meal
- 7 meal slots including Pre/Post-Workout
- Bundled Chart.js (offline support)

### v1.0.0 — Initial Release
- Task, habit, work, goals, diet, journal tracking
- Dashboard charts, weekly/monthly summaries
- Excel export, dark mode, mobile responsive
- Cross-platform installers (Linux, Windows, macOS)

---

## Installation

### Linux (Ubuntu / Debian)

Download `activity-tracker_1.4.0_all.deb` from the [Releases](../../releases) page, then:

```bash
sudo dpkg -i activity-tracker_1.4.0_all.deb
sudo apt-get install -f
```

Launch from your app menu or run `activity-tracker` in a terminal. The app auto-starts on login.

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
git tag v1.4.0
git push origin v1.4.0
```

This automatically builds all three installers and creates a GitHub Release.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Frontend | Vanilla JS, Chart.js v4 (bundled), CSS custom properties |
| Desktop window | Chrome `--app` mode on Linux; pywebview on Windows/macOS |
| Nutrition engine | Mifflin-St Jeor BMR, goal-based macro calculation, real-time client-side TDEE |
| AI coaching engine | Linear regression weight trend, rule-based insight system with priority ranking |
| Wellness engine | Composite recovery score from sleep / energy / stress / soreness |
| Data storage | JSON files |
| Excel export | openpyxl (10 sheets) |
| Packaging | PyInstaller + dpkg / Inno Setup / create-dmg |
| CI/CD | GitHub Actions |

---

## License

MIT
