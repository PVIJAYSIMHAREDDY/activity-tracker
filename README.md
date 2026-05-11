# Daily Activity Tracker

An elite AI-powered fitness and nutrition coaching app for tracking tasks, habits, work hours, goals, diet, workouts, and wellness — with a **real-time TDEE calculator**, **adaptive AI coaching system**, body measurements tracking, and Excel export.

Built with **Flask + Chrome app-mode** (native window, no browser needed) and packaged as a one-click installer for Linux, Windows, and macOS.

---

## Features

### 📋 My Plan — Personal Nutrition Calculator *(v1.3.0)*
- **Metric / Imperial toggle** — switch between kg·cm and lbs·ft with one click
- **Visual activity cards** — select your level: Sedentary, Light, Moderate, Very Active, Athlete
- **Goal cards** — Lose Weight / Maintain / Gain Muscle / Athletic Performance
- **Real-time TDEE** — Mifflin-St Jeor BMR × activity multiplier, updates as you type
- **Hero display** — daily calorie target with BMR, Maintenance, and Adjustment at a glance
- **SVG macro rings** — Protein / Carbs / Fat with grams and % of total energy
- **Micro-nutrient targets** — Fiber, Sugar, and Sodium limits
- **Goal Timeline** — estimated weeks to target weight + projected completion date
- **Save & Apply** — one click saves profile and pushes targets to active diet plan
- **Create Plan** — save calculated targets as a named plan in the Plans Library

### 📏 Body Measurements *(v1.3.0)*
- Log **9 metrics** with date: weight, body fat %, waist, chest, hips, neck, bicep, thigh, calf
- History table shows **+/− delta** vs previous entry for every metric
- Form pre-fills from latest saved entry for quick updates
- Full **Measurements sheet** in Excel export with delta columns

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

### 🥗 Diet & Nutrition
- **7 meal slots**: Breakfast, Morning Snack, Lunch, Pre-Workout, Post-Workout, Dinner, Evening Snack
- Log **7 nutrients per meal**: Calories, Protein, Carbs, Fat, Fiber, Sugar, Sodium
- **Edit** any food entry inline with the ✏️ button
- **Animated macro rings** (Calories / Protein / Carbs / Fat) showing % of daily goal
- **Color-coded micro bars** for Fiber, Sugar, Sodium (✅ on track / ⚠️ near limit / ❌ exceeded)
- Per-meal subtotals in the meals list

### 🏋️ Training Log *(v1.4.0)*
- Log workouts with **6 types**: Strength, Cardio, HIIT, Yoga/Flex, Sport, Other
- Track **exercises per session**: name, sets, reps, weight
- **Weekly training summary** — total sessions, minutes, sets, types
- Full workout history with per-session exercise breakdown and delete support

### 🌟 Wellness Tracking *(v1.4.0)*
- **Daily check-in** with 5 metrics via visual sliders: sleep hours, sleep quality, energy, stress, muscle soreness
- Pre-fills with today's saved entry for quick updates
- Recent 5-day history strip with mini progress bars

### 🤖 AI Coaching Dashboard *(v1.4.0)*
- **Overall fitness score** (0–99) with animated SVG ring and letter grade (S/A/B/C/D)
- **6 key stats**: diet adherence, recovery score, training sessions, weight trend, target kcal, target protein
- **Alert banners** — danger/warning/info for plateau, overeating, critical recovery, overtraining
- **Today's Focus** — 3 personalised priority action cards based on your data
- **Insight cards** with priority chips (High/Medium): plateau detection, protein gaps, recovery alerts, goal achievements
- **AI adjustments panel** — calorie and protein delta recommendations

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

### 📊 Dashboard & Reports
- Chart.js charts: task completion, habit heatmap, work hours, diet macros
- **Weekly Summary** — 7-day trend charts + aggregated stats
- **Monthly Summary** — full month breakdown with calendar view
- **Excel Export** — one-click `.xlsx` with 10 styled sheets:
  Tasks, Habits, Work, Goals, Diet, 7-Day Summary, Journal, Dashboard, **Measurements**, **Progress Report**

### 🎨 UI & UX
- **Dark mode** — toggleable, persisted across sessions
- **Mobile responsive** — scrollable tab strip, stacked forms, touch-friendly
- **Offline-first** — Chart.js bundled locally, no internet required
- **Auto-start on login** — launches automatically on every session (Linux)

---

## What's New

### v1.4.0 — Elite AI Coaching Platform
- **AI Coaching Dashboard** — overall fitness score (0–99) with letter grade (S/A/B/C/D), animated score ring, 6 stats at a glance (diet adherence, recovery, sessions/week, weight trend, target kcal/protein)
- **Training tab** — log workouts with type (Strength/Cardio/HIIT/Yoga/Sport/Other), duration, notes, and per-exercise sets/reps/weight tracking; weekly training summary with total sessions/minutes/sets
- **Daily Wellness Check-in** — log sleep hours, sleep quality, energy, stress, and muscle soreness with visual sliders; pre-fills with today's entry; recent 5-day history strip
- **AI Insight Engine** — plateau detection, adherence/undereating/overeating alerts, protein gap warnings, recovery analysis, overtraining detection, body composition tracking, goal achievement recognition
- **Today's Focus** — 3 priority action cards generated from your personal data and goal
- **Alert Banners** — danger/warning/info banners for critical issues requiring immediate attention
- **AI Calorie/Protein Recommendations** — data-driven adjustments displayed on the dashboard and available as one-click apply

### v1.3.0 — Personal Nutrition Calculator & Measurements
- **My Plan tab** — dedicated nutrition calculator: real-time TDEE, visual activity/goal cards, SVG macro rings, goal timeline with projected date, metric/imperial toggle, one-click apply to diet plan
- **Body Measurements** — log 9 body metrics with date; history table shows +/− deltas vs previous entry
- **Edit any entry** — ✏️ buttons on every task, work session, diet entry, and goal
- **Excel: 2 new sheets** — Measurements (delta columns) + Progress Report (first vs latest comparison for weight, body metrics, and all goals)
- **Auto-start on login** — desktop autostart entry installed automatically on Linux
- **Bug fixes** — resolved JS crash (duplicate identifier), monthly tab navigation, Chrome app-mode input handling

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
| Desktop window | Chrome `--app` mode on Linux; pywebview on Windows/macOS |
| Nutrition engine | Mifflin-St Jeor BMR, goal-based macro calculation, real-time client-side TDEE |
| Coaching engine | Rule-based insights with priority ranking |
| Data storage | JSON files |
| Excel export | openpyxl (10 sheets) |
| Packaging | PyInstaller + dpkg / Inno Setup / create-dmg |
| CI/CD | GitHub Actions |

---

## License

MIT
