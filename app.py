from flask import Flask, render_template, request, jsonify, send_file
from datetime import date, datetime, timedelta
import json, os, io, sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.chart import BarChart, Reference, LineChart, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation

# ── Path resolution: works from source AND from installed/PyInstaller bundle ─
if getattr(sys, 'frozen', False):
    _BUNDLE = sys._MEIPASS
    app = Flask(__name__,
                template_folder=os.path.join(_BUNDLE, 'templates'),
                static_folder=os.path.join(_BUNDLE, 'static'))
    DATA_DIR = os.path.join(os.path.expanduser('~'), '.activity-tracker', 'data')
else:
    app = Flask(__name__)
    _local_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    # If the app is installed somewhere not writable by this user (e.g. /opt/),
    # store data in the user's home directory instead.
    try:
        os.makedirs(_local_data, exist_ok=True)
        test_f = os.path.join(_local_data, '.write_test')
        open(test_f, 'w').close()
        os.remove(test_f)
        DATA_DIR = _local_data
    except OSError:
        DATA_DIR = os.path.join(os.path.expanduser('~'), '.activity-tracker', 'data')

os.makedirs(DATA_DIR, exist_ok=True)

def data_file(name):
    return os.path.join(DATA_DIR, f"{name}.json")

def load_json(name, default):
    path = data_file(name)
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return default

def save_json(name, data):
    with open(data_file(name), "w") as f:
        json.dump(data, f, indent=2, default=str)

today = lambda: date.today().isoformat()

# ── routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

# ── Tasks ────────────────────────────────────────────────────────────────────

@app.route("/api/tasks", methods=["GET"])
def get_tasks():
    d = request.args.get("date", today())
    tasks = load_json("tasks", {})
    return jsonify(tasks.get(d, []))

@app.route("/api/tasks", methods=["POST"])
def add_task():
    d = request.json.get("date", today())
    tasks = load_json("tasks", {})
    day_tasks = tasks.get(d, [])
    task = {
        "id": int(datetime.now().timestamp() * 1000),
        "text": request.json["text"],
        "priority": request.json.get("priority", "Medium"),
        "category": request.json.get("category", "General"),
        "done": False,
        "created": datetime.now().isoformat()
    }
    day_tasks.append(task)
    tasks[d] = day_tasks
    save_json("tasks", tasks)
    return jsonify(task)

@app.route("/api/tasks/<int:tid>", methods=["PATCH"])
def update_task(tid):
    d = request.json.get("date", today())
    tasks = load_json("tasks", {})
    for t in tasks.get(d, []):
        if t["id"] == tid:
            t.update({k: v for k, v in request.json.items() if k != "date"})
    save_json("tasks", tasks)
    return jsonify({"ok": True})

@app.route("/api/tasks/<int:tid>", methods=["DELETE"])
def delete_task(tid):
    d = request.args.get("date", today())
    tasks = load_json("tasks", {})
    tasks[d] = [t for t in tasks.get(d, []) if t["id"] != tid]
    save_json("tasks", tasks)
    return jsonify({"ok": True})

# ── Habits ───────────────────────────────────────────────────────────────────

@app.route("/api/habits/list", methods=["GET"])
def get_habit_list():
    return jsonify(load_json("habit_list", [
        {"id": 1, "name": "Exercise", "target": 1, "unit": "session", "icon": "🏃"},
        {"id": 2, "name": "Water",    "target": 8, "unit": "glasses", "icon": "💧"},
        {"id": 3, "name": "Reading",  "target": 30,"unit": "minutes", "icon": "📚"},
        {"id": 4, "name": "Sleep",    "target": 8, "unit": "hours",   "icon": "😴"},
        {"id": 5, "name": "Meditation","target": 10,"unit": "minutes","icon": "🧘"},
    ]))

@app.route("/api/habits/list", methods=["POST"])
def add_habit():
    habits = load_json("habit_list", [])
    h = request.json
    h["id"] = int(datetime.now().timestamp() * 1000)
    habits.append(h)
    save_json("habit_list", habits)
    return jsonify(h)

@app.route("/api/habits/list/<int:hid>", methods=["DELETE"])
def delete_habit(hid):
    habits = load_json("habit_list", [])
    save_json("habit_list", [h for h in habits if h["id"] != hid])
    return jsonify({"ok": True})

@app.route("/api/habits/log", methods=["GET"])
def get_habit_log():
    d = request.args.get("date", today())
    logs = load_json("habit_log", {})
    return jsonify(logs.get(d, {}))

@app.route("/api/habits/log", methods=["POST"])
def set_habit_log():
    d = request.json.get("date", today())
    hid = str(request.json["habit_id"])
    val = request.json["value"]
    logs = load_json("habit_log", {})
    if d not in logs:
        logs[d] = {}
    logs[d][hid] = val
    save_json("habit_log", logs)
    return jsonify({"ok": True})

# ── Work Hours ────────────────────────────────────────────────────────────────

@app.route("/api/work", methods=["GET"])
def get_work():
    d = request.args.get("date", today())
    work = load_json("work", {})
    return jsonify(work.get(d, []))

@app.route("/api/work", methods=["POST"])
def add_work():
    d = request.json.get("date", today())
    work = load_json("work", {})
    entry = {
        "id": int(datetime.now().timestamp() * 1000),
        "project": request.json["project"],
        "hours": float(request.json["hours"]),
        "notes": request.json.get("notes", ""),
        "created": datetime.now().isoformat()
    }
    work.setdefault(d, []).append(entry)
    save_json("work", work)
    return jsonify(entry)

@app.route("/api/work/<int:wid>", methods=["DELETE"])
def delete_work(wid):
    d = request.args.get("date", today())
    work = load_json("work", {})
    work[d] = [w for w in work.get(d, []) if w["id"] != wid]
    save_json("work", work)
    return jsonify({"ok": True})

# ── Goals ─────────────────────────────────────────────────────────────────────

@app.route("/api/goals", methods=["GET"])
def get_goals():
    return jsonify(load_json("goals", []))

@app.route("/api/goals", methods=["POST"])
def add_goal():
    goals = load_json("goals", [])
    g = {
        "id": int(datetime.now().timestamp() * 1000),
        "title": request.json["title"],
        "target": float(request.json["target"]),
        "current": float(request.json.get("current", 0)),
        "unit": request.json.get("unit", ""),
        "deadline": request.json.get("deadline", ""),
        "category": request.json.get("category", "Personal"),
        "created": today()
    }
    goals.append(g)
    save_json("goals", goals)
    return jsonify(g)

@app.route("/api/goals/<int:gid>", methods=["PATCH"])
def update_goal(gid):
    goals = load_json("goals", [])
    for g in goals:
        if g["id"] == gid:
            g.update({k: v for k, v in request.json.items()})
    save_json("goals", goals)
    return jsonify({"ok": True})

@app.route("/api/goals/<int:gid>", methods=["DELETE"])
def delete_goal(gid):
    goals = load_json("goals", [])
    save_json("goals", [g for g in goals if g["id"] != gid])
    return jsonify({"ok": True})

# ── Diet ──────────────────────────────────────────────────────────────────────

@app.route("/api/diet", methods=["GET"])
def get_diet():
    d = request.args.get("date", today())
    diet = load_json("diet", {})
    return jsonify(diet.get(d, []))

@app.route("/api/diet", methods=["POST"])
def add_diet():
    d = request.json.get("date", today())
    diet = load_json("diet", {})
    entry = {
        "id": int(datetime.now().timestamp() * 1000),
        "meal": request.json["meal"],
        "food": request.json["food"],
        "calories": int(request.json.get("calories", 0)),
        "protein": float(request.json.get("protein", 0)),
        "carbs": float(request.json.get("carbs", 0)),
        "fat": float(request.json.get("fat", 0)),
        "time": request.json.get("time", "")
    }
    diet.setdefault(d, []).append(entry)
    save_json("diet", diet)
    return jsonify(entry)

@app.route("/api/diet/<int:did>", methods=["DELETE"])
def delete_diet(did):
    d = request.args.get("date", today())
    diet = load_json("diet", {})
    diet[d] = [x for x in diet.get(d, []) if x["id"] != did]
    save_json("diet", diet)
    return jsonify({"ok": True})

# ── Week Summary ─────────────────────────────────────────────────────────────

@app.route("/api/summary/week")
def week_summary():
    all_tasks = load_json("tasks", {})
    all_work  = load_json("work", {})
    all_diet  = load_json("diet", {})
    result = []
    for i in range(6, -1, -1):
        day = (date.today() - timedelta(days=i)).isoformat()
        dt = all_tasks.get(day, [])
        dw = all_work.get(day, [])
        dd = all_diet.get(day, [])
        result.append({
            "date":    day,
            "label":   day[5:],
            "done":    sum(1 for t in dt if t.get("done")),
            "total":   len(dt),
            "hours":   round(sum(w["hours"] for w in dw), 2),
            "calories": sum(e.get("calories", 0) for e in dd),
        })
    return jsonify(result)

# ── Notes / Journal ──────────────────────────────────────────────────────────

@app.route("/api/notes", methods=["GET"])
def get_notes():
    d = request.args.get("date", today())
    return jsonify(load_json("notes", {}).get(d, []))

@app.route("/api/notes", methods=["POST"])
def add_note():
    d = request.json.get("date", today())
    notes = load_json("notes", {})
    note = {
        "id":      int(datetime.now().timestamp() * 1000),
        "type":    request.json.get("type", "note"),
        "title":   request.json.get("title", ""),
        "content": request.json.get("content", ""),
        "mood":    request.json.get("mood", ""),
        "tags":    request.json.get("tags", []),
        "pinned":  False,
        "created": datetime.now().isoformat(),
        "updated": datetime.now().isoformat(),
    }
    notes.setdefault(d, []).insert(0, note)
    save_json("notes", notes)
    return jsonify(note)

@app.route("/api/notes/<int:nid>", methods=["PATCH"])
def update_note(nid):
    d = request.json.get("date", today())
    notes = load_json("notes", {})
    for n in notes.get(d, []):
        if n["id"] == nid:
            n.update({k: v for k, v in request.json.items() if k != "date"})
            n["updated"] = datetime.now().isoformat()
    save_json("notes", notes)
    return jsonify({"ok": True})

@app.route("/api/notes/<int:nid>", methods=["DELETE"])
def delete_note(nid):
    d = request.args.get("date", today())
    notes = load_json("notes", {})
    notes[d] = [n for n in notes.get(d, []) if n["id"] != nid]
    save_json("notes", notes)
    return jsonify({"ok": True})

@app.route("/api/notes/search")
def search_notes():
    q = request.args.get("q", "").lower().strip()
    if not q:
        return jsonify([])
    all_notes = load_json("notes", {})
    results = []
    for day, day_notes in sorted(all_notes.items(), reverse=True):
        for n in day_notes:
            if (q in n.get("title","").lower() or
                q in n.get("content","").lower() or
                any(q in t.lower() for t in n.get("tags", []))):
                results.append({**n, "date": day})
    return jsonify(results[:50])

# ── Weekly Summary ───────────────────────────────────────────────────────────

@app.route("/api/summary/weekly")
def weekly_summary():
    import calendar as cal_mod
    d        = request.args.get("date", today())
    dt       = date.fromisoformat(d)
    wstart   = dt - timedelta(days=dt.weekday())   # Monday

    all_tasks   = load_json("tasks", {})
    all_work    = load_json("work", {})
    all_habits  = load_json("habit_log", {})
    all_diet    = load_json("diet", {})
    habit_list  = load_json("habit_list", [])

    days = []
    tot_done = tot_tasks = tot_hours = tot_cal = 0
    work_by_project = {}

    for i in range(7):
        day    = (wstart + timedelta(days=i)).isoformat()
        dtasks = all_tasks.get(day, [])
        dwork  = all_work.get(day, [])
        ddiet  = all_diet.get(day, [])
        dhab   = all_habits.get(day, {})

        done  = sum(1 for t in dtasks if t.get("done"))
        total = len(dtasks)
        hours = round(sum(w["hours"] for w in dwork), 2)
        cal   = sum(e.get("calories", 0) for e in ddiet)

        pcts = [min(float(dhab.get(str(h["id"]), 0)) / float(h.get("target", 1)) * 100, 100)
                for h in habit_list if float(h.get("target", 1)) > 0]
        hab_score = round(sum(pcts) / len(pcts)) if pcts else 0

        for w in dwork:
            work_by_project[w["project"]] = round(
                work_by_project.get(w["project"], 0) + w["hours"], 2)

        tot_done  += done
        tot_tasks += total
        tot_hours += hours
        tot_cal   += cal

        days.append({
            "date": day, "label": (wstart + timedelta(days=i)).strftime("%a"),
            "short": (wstart + timedelta(days=i)).strftime("%m/%d"),
            "done": done, "total": total, "hours": hours,
            "calories": cal, "habit_score": hab_score,
        })

    habit_breakdown = []
    for h in habit_list:
        pcts, completed = [], 0
        for i in range(7):
            day  = (wstart + timedelta(days=i)).isoformat()
            dhab = all_habits.get(day, {})
            logged = float(dhab.get(str(h["id"]), 0))
            tgt    = float(h.get("target", 1))
            pct    = min(logged / tgt * 100, 100) if tgt else 0
            pcts.append(pct)
            if pct >= 100: completed += 1
        habit_breakdown.append({
            "name": h["name"], "icon": h.get("icon", ""),
            "days_completed": completed,
            "avg_pct": round(sum(pcts) / len(pcts)) if pcts else 0,
        })

    all_scores = [d["habit_score"] for d in days]
    return jsonify({
        "week_start": wstart.isoformat(),
        "week_end":   (wstart + timedelta(days=6)).isoformat(),
        "days": days,
        "totals": {
            "done": tot_done, "total": tot_tasks,
            "hours": round(tot_hours, 2), "calories": tot_cal,
            "avg_habit_score": round(sum(all_scores) / len(all_scores)) if all_scores else 0,
        },
        "work_by_project": work_by_project,
        "habit_breakdown": habit_breakdown,
    })

# ── Monthly Summary ───────────────────────────────────────────────────────────

@app.route("/api/summary/monthly")
def monthly_summary():
    import calendar as cal_mod
    d     = request.args.get("date", today())
    dt    = date.fromisoformat(d)
    year, month = dt.year, dt.month
    _, dim      = cal_mod.monthrange(year, month)
    mstart      = date(year, month, 1)

    all_tasks   = load_json("tasks", {})
    all_work    = load_json("work", {})
    all_habits  = load_json("habit_log", {})
    all_diet    = load_json("diet", {})
    habit_list  = load_json("habit_list", [])

    days = []
    work_by_project = {}
    tot_done = tot_tasks = tot_hours = tot_cal = 0

    for i in range(dim):
        day    = (mstart + timedelta(days=i)).isoformat()
        dtasks = all_tasks.get(day, [])
        dwork  = all_work.get(day, [])
        ddiet  = all_diet.get(day, [])
        dhab   = all_habits.get(day, {})

        done  = sum(1 for t in dtasks if t.get("done"))
        total = len(dtasks)
        hours = round(sum(w["hours"] for w in dwork), 2)
        cal   = sum(e.get("calories", 0) for e in ddiet)

        pcts = [min(float(dhab.get(str(h["id"]), 0)) / float(h.get("target", 1)) * 100, 100)
                for h in habit_list if float(h.get("target", 1)) > 0]
        hab_score = round(sum(pcts) / len(pcts)) if pcts else 0

        for w in dwork:
            work_by_project[w["project"]] = round(
                work_by_project.get(w["project"], 0) + w["hours"], 2)

        tot_done  += done
        tot_tasks += total
        tot_hours += hours
        tot_cal   += cal

        days.append({
            "date": day, "day_num": i + 1,
            "weekday": (mstart + timedelta(days=i)).weekday(),
            "done": done, "total": total, "hours": hours,
            "calories": cal, "habit_score": hab_score,
            "completion_pct": round(done / total * 100) if total else 0,
        })

    # Weekly buckets (for the bar chart)
    weeks = []
    for w in range(0, dim, 7):
        chunk = days[w:w + 7]
        weeks.append({
            "label": f"W{w//7+1}",
            "done":  sum(d["done"] for d in chunk),
            "total": sum(d["total"] for d in chunk),
            "hours": round(sum(d["hours"] for d in chunk), 2),
        })

    habit_breakdown = []
    for h in habit_list:
        pcts, completed = [], 0
        for i in range(dim):
            day  = (mstart + timedelta(days=i)).isoformat()
            dhab = all_habits.get(day, {})
            logged = float(dhab.get(str(h["id"]), 0))
            tgt    = float(h.get("target", 1))
            pct    = min(logged / tgt * 100, 100) if tgt else 0
            pcts.append(pct)
            if pct >= 100: completed += 1
        habit_breakdown.append({
            "name": h["name"], "icon": h.get("icon", ""),
            "days_completed": completed, "days_in_month": dim,
            "avg_pct": round(sum(pcts) / len(pcts)) if pcts else 0,
        })

    all_scores = [d["habit_score"] for d in days]
    return jsonify({
        "year": year, "month": month,
        "month_name": mstart.strftime("%B %Y"),
        "days_in_month": dim,
        "month_start_weekday": mstart.weekday(),
        "days": days, "weeks": weeks,
        "totals": {
            "done": tot_done, "total": tot_tasks,
            "hours": round(tot_hours, 2), "calories": tot_cal,
            "avg_habit_score": round(sum(all_scores) / len(all_scores)) if all_scores else 0,
        },
        "work_by_project": work_by_project,
        "habit_breakdown": habit_breakdown,
    })

# ── Excel Export ──────────────────────────────────────────────────────────────

def style_header(ws, row, col, text, bg="1E3A5F", fg="FFFFFF", bold=True, size=12):
    cell = ws.cell(row=row, column=col, value=text)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(color=fg, bold=bold, size=size)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def pct_bar(ws, row, col, pct):
    """Write a text-based progress bar."""
    filled = int(pct / 10)
    bar = "█" * filled + "░" * (10 - filled)
    cell = ws.cell(row=row, column=col, value=f"{bar} {pct:.0f}%")
    cell.font = Font(color="2E86AB" if pct < 100 else "27AE60", bold=True)
    cell.alignment = Alignment(horizontal="center")

def add_thick_border(ws, min_row, min_col, max_row, max_col):
    thick = Side(style="medium", color="1E3A5F")
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            b = Border(
                left=thick if cell.column == min_col else Side(style="thin", color="CCCCCC"),
                right=thick if cell.column == max_col else Side(style="thin", color="CCCCCC"),
                top=thick if cell.row == min_row else Side(style="thin", color="CCCCCC"),
                bottom=thick if cell.row == max_row else Side(style="thin", color="CCCCCC"),
            )
            cell.border = b

@app.route("/api/export")
def export_excel():
    d = request.args.get("date", today())
    tasks     = load_json("tasks", {}).get(d, [])
    habit_list= load_json("habit_list", [])
    habit_log = load_json("habit_log", {}).get(d, {})
    work_list = load_json("work", {}).get(d, [])
    goals     = load_json("goals", [])
    diet_list = load_json("diet", {}).get(d, [])
    notes_list= load_json("notes", {}).get(d, [])

    wb = Workbook()

    # ── color palette ──────────────────────────────────────────────────────
    C = {
        "navy":    "1E3A5F",
        "blue":    "2E86AB",
        "green":   "27AE60",
        "orange":  "F39C12",
        "red":     "E74C3C",
        "purple":  "8E44AD",
        "teal":    "16A085",
        "light":   "EBF5FB",
        "light2":  "F0FFF4",
        "light3":  "FFF9E6",
        "light4":  "F9F0FF",
        "white":   "FFFFFF",
        "gray":    "F8F9FA",
        "dgray":   "6C757D",
    }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 1 – DASHBOARD
    # ════════════════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "📊 Dashboard"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 3

    # title banner
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 50
    ws.row_dimensions[3].height = 10
    ws.merge_cells("B2:E2")
    c = ws["B2"]
    c.value = f"📅  DAILY ACTIVITY TRACKER  —  {d}"
    c.fill = PatternFill("solid", fgColor=C["navy"])
    c.font = Font(color=C["white"], bold=True, size=20)
    c.alignment = Alignment(horizontal="center", vertical="center")

    # summary cards (row 4-8)
    done   = sum(1 for t in tasks if t.get("done"))
    total  = len(tasks)
    task_pct = (done / total * 100) if total else 0

    hab_pct_list = []
    for h in habit_list:
        logged = float(habit_log.get(str(h["id"]), 0))
        target = float(h.get("target", 1))
        hab_pct_list.append(min(logged / target * 100, 100) if target else 0)
    hab_pct = sum(hab_pct_list) / len(hab_pct_list) if hab_pct_list else 0

    total_work = sum(w["hours"] for w in work_list)
    total_cal  = sum(e["calories"] for e in diet_list)

    cards = [
        ("✅  Tasks",    f"{done}/{total} done",   f"{task_pct:.0f}%",  C["blue"],   C["light"]),
        ("💪  Habits",   f"{len(hab_pct_list)} tracked", f"{hab_pct:.0f}% avg", C["green"],  C["light2"]),
        ("⏱  Work",      f"{total_work:.1f} hrs",  "Today",             C["orange"], C["light3"]),
        ("🥗  Calories", f"{total_cal} kcal",       f"{len(diet_list)} meals", C["purple"], C["light4"]),
    ]

    ws.row_dimensions[4].height = 14
    for ci, (title, sub, val, color, bg) in enumerate(cards):
        col = ci + 2  # B=2 C=3 D=4 E=5
        ws.row_dimensions[5].height = 20
        ws.row_dimensions[6].height = 30
        ws.row_dimensions[7].height = 24
        ws.row_dimensions[8].height = 14

        for r in range(5, 9):
            ws.cell(r, col).fill = PatternFill("solid", fgColor=bg)
            ws.cell(r, col).border = thin_border()

        t = ws.cell(5, col, title)
        t.fill = PatternFill("solid", fgColor=color)
        t.font = Font(color=C["white"], bold=True, size=11)
        t.alignment = Alignment(horizontal="center", vertical="center")

        v = ws.cell(6, col, val)
        v.font = Font(color=color, bold=True, size=22)
        v.alignment = Alignment(horizontal="center", vertical="center")

        s = ws.cell(7, col, sub)
        s.font = Font(color=C["dgray"], size=10)
        s.alignment = Alignment(horizontal="center")

    # quick task list
    ws.row_dimensions[10].height = 24
    ws.merge_cells("B10:E10")
    h = ws["B10"]
    h.value = "📝  TODAY'S TASKS OVERVIEW"
    h.fill = PatternFill("solid", fgColor=C["navy"])
    h.font = Font(color=C["white"], bold=True, size=13)
    h.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Task", "Category", "Priority", "Status"]
    for ci, hdr in enumerate(headers):
        c = ws.cell(11, ci + 2, hdr)
        c.fill = PatternFill("solid", fgColor=C["blue"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")
        ws.row_dimensions[11].height = 18

    for ri, t in enumerate(tasks):
        row = ri + 12
        ws.row_dimensions[row].height = 18
        status = "✅ Done" if t.get("done") else "⏳ Pending"
        bg = "E8F8F5" if t.get("done") else "FEFEFE"
        vals = [t["text"], t.get("category",""), t.get("priority",""), status]
        for ci, v in enumerate(vals):
            cell = ws.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if ci > 0 else "left",
                                        vertical="center", wrap_text=True)

    add_thick_border(ws, 10, 2, 11 + len(tasks), 5)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 2 – TASKS
    # ════════════════════════════════════════════════════════════════════════
    wt = wb.create_sheet("✅ Tasks")
    wt.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFG", [3, 5, 35, 16, 14, 14, 14]):
        wt.column_dimensions[col].width = w

    wt.row_dimensions[1].height = 10
    wt.row_dimensions[2].height = 45
    wt.merge_cells("B2:G2")
    c = wt["B2"]
    c.value = f"✅  TASKS  —  {d}"
    c.fill = PatternFill("solid", fgColor=C["blue"])
    c.font = Font(color=C["white"], bold=True, size=16)
    c.alignment = Alignment(horizontal="center", vertical="center")

    # stats row
    wt.row_dimensions[4].height = 22
    stats = [("Total", total), ("Completed", done),
             ("Pending", total - done), ("Completion %", f"{task_pct:.1f}%")]
    for i, (lbl, val) in enumerate(stats):
        bg = [C["navy"], C["green"], C["orange"], C["blue"]][i]
        lc = wt.cell(4, i + 2, f"{lbl}: {val}")
        lc.fill = PatternFill("solid", fgColor=bg)
        lc.font = Font(color=C["white"], bold=True, size=11)
        lc.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["#", "Task Description", "Category", "Priority", "Status", "Created"]
    wt.row_dimensions[6].height = 22
    for ci, h in enumerate(headers):
        c = wt.cell(6, ci + 2, h)
        c.fill = PatternFill("solid", fgColor=C["navy"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")

    pri_colors = {"High": "E74C3C", "Medium": "F39C12", "Low": "27AE60"}
    for ri, t in enumerate(tasks):
        row = ri + 7
        wt.row_dimensions[row].height = 20
        bg = "E8F8F5" if t.get("done") else ("FFF9E6" if t.get("priority") == "High" else "FEFEFE")
        vals = [ri + 1, t["text"], t.get("category","General"),
                t.get("priority","Medium"),
                "✅ Done" if t.get("done") else "⏳ Pending",
                t.get("created","")[:10]]
        for ci, v in enumerate(vals):
            cell = wt.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if ci != 1 else "left",
                                        vertical="center")
            if ci == 3:  # priority
                pc = pri_colors.get(str(v), C["dgray"])
                cell.font = Font(color=pc, bold=True)

    if tasks:
        add_thick_border(wt, 6, 2, 6 + len(tasks), 7)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 – HABITS
    # ════════════════════════════════════════════════════════════════════════
    wh = wb.create_sheet("💪 Habits")
    wh.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [3, 5, 20, 14, 14, 14, 28, 3]):
        wh.column_dimensions[col].width = w

    wh.row_dimensions[2].height = 45
    wh.merge_cells("B2:G2")
    c = wh["B2"]
    c.value = f"💪  HABITS  —  {d}"
    c.fill = PatternFill("solid", fgColor=C["green"])
    c.font = Font(color=C["white"], bold=True, size=16)
    c.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["", "Habit", "Target", "Logged", "Unit", "Progress"]
    wh.row_dimensions[4].height = 22
    for ci, h in enumerate(headers):
        c = wh.cell(4, ci + 2, h)
        c.fill = PatternFill("solid", fgColor=C["navy"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")

    for ri, h in enumerate(habit_list):
        row = ri + 5
        wh.row_dimensions[row].height = 24
        logged = float(habit_log.get(str(h["id"]), 0))
        target = float(h.get("target", 1))
        pct = min(logged / target * 100, 100) if target else 0
        bg = "E8F8F5" if pct >= 100 else ("FFF9E6" if pct >= 50 else "FFF0F0")

        vals = [h.get("icon",""), h["name"], target, logged, h.get("unit","")]
        for ci, v in enumerate(vals):
            cell = wh.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center", vertical="center")

        pct_bar(wh, row, 7, pct)
        wh.cell(row, 7).fill = PatternFill("solid", fgColor=bg)
        wh.cell(row, 7).border = thin_border()

    add_thick_border(wh, 4, 2, 4 + len(habit_list), 7)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 4 – WORK HOURS
    # ════════════════════════════════════════════════════════════════════════
    ww = wb.create_sheet("⏱ Work Hours")
    ww.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [3, 30, 14, 36, 3]):
        ww.column_dimensions[col].width = w

    ww.row_dimensions[2].height = 45
    ww.merge_cells("B2:D2")
    c = ww["B2"]
    c.value = f"⏱  WORK HOURS  —  {d}"
    c.fill = PatternFill("solid", fgColor=C["orange"])
    c.font = Font(color=C["white"], bold=True, size=16)
    c.alignment = Alignment(horizontal="center", vertical="center")

    total_h = sum(w["hours"] for w in work_list)
    ww.row_dimensions[4].height = 22
    c = ww.cell(4, 2, f"Total Hours Today: {total_h:.1f} hrs")
    c.fill = PatternFill("solid", fgColor=C["navy"])
    c.font = Font(color=C["white"], bold=True, size=12)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ww.merge_cells("B4:D4")

    headers = ["Project / Task", "Hours", "Notes"]
    ww.row_dimensions[6].height = 22
    for ci, h in enumerate(headers):
        c = ww.cell(6, ci + 2, h)
        c.fill = PatternFill("solid", fgColor=C["navy"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")

    for ri, w in enumerate(work_list):
        row = ri + 7
        ww.row_dimensions[row].height = 20
        bg = C["light3"] if ri % 2 == 0 else C["white"]
        vals = [w["project"], w["hours"], w.get("notes","")]
        for ci, v in enumerate(vals):
            cell = ww.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if ci == 1 else "left",
                                        vertical="center", wrap_text=True)

    if work_list:
        add_thick_border(ww, 6, 2, 6 + len(work_list), 4)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 5 – GOALS
    # ════════════════════════════════════════════════════════════════════════
    wg = wb.create_sheet("🎯 Goals")
    wg.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [3, 28, 16, 14, 14, 14, 28, 3]):
        wg.column_dimensions[col].width = w

    wg.row_dimensions[2].height = 45
    wg.merge_cells("B2:G2")
    c = wg["B2"]
    c.value = "🎯  GOALS & PROGRESS"
    c.fill = PatternFill("solid", fgColor=C["purple"])
    c.font = Font(color=C["white"], bold=True, size=16)
    c.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Goal", "Category", "Current", "Target", "Deadline", "Progress"]
    wg.row_dimensions[4].height = 22
    for ci, h in enumerate(headers):
        c = wg.cell(4, ci + 2, h)
        c.fill = PatternFill("solid", fgColor=C["navy"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")

    for ri, g in enumerate(goals):
        row = ri + 5
        wg.row_dimensions[row].height = 24
        cur = float(g.get("current", 0))
        tgt = float(g.get("target", 1))
        pct = min(cur / tgt * 100, 100) if tgt else 0
        bg = "E8F8F5" if pct >= 100 else C["light4"]
        vals = [g["title"], g.get("category",""), cur, tgt, g.get("deadline","")]
        for ci, v in enumerate(vals):
            cell = wg.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if ci > 0 else "left",
                                        vertical="center")
        pct_bar(wg, row, 7, pct)
        wg.cell(row, 7).fill = PatternFill("solid", fgColor=bg)
        wg.cell(row, 7).border = thin_border()

    if goals:
        add_thick_border(wg, 4, 2, 4 + len(goals), 7)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 6 – DIET
    # ════════════════════════════════════════════════════════════════════════
    wd = wb.create_sheet("🥗 Diet")
    wd.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", [3, 16, 28, 12, 12, 12, 12, 12, 3]):
        wd.column_dimensions[col].width = w

    wd.row_dimensions[2].height = 45
    wd.merge_cells("B2:H2")
    c = wd["B2"]
    c.value = f"🥗  DIET LOG  —  {d}"
    c.fill = PatternFill("solid", fgColor=C["teal"])
    c.font = Font(color=C["white"], bold=True, size=16)
    c.alignment = Alignment(horizontal="center", vertical="center")

    # macro totals
    totals = {"calories": 0, "protein": 0, "carbs": 0, "fat": 0}
    for e in diet_list:
        for k in totals:
            totals[k] += e.get(k, 0)

    wd.row_dimensions[4].height = 22
    macro_labels = [
        (f"🔥 {totals['calories']} kcal", C["orange"]),
        (f"💪 {totals['protein']:.1f}g protein", C["blue"]),
        (f"🍞 {totals['carbs']:.1f}g carbs", C["green"]),
        (f"🧈 {totals['fat']:.1f}g fat", C["purple"]),
    ]
    for i, (lbl, color) in enumerate(macro_labels):
        c = wd.cell(4, i + 2, lbl)
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(color=C["white"], bold=True, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Meal", "Food Item", "Calories", "Protein(g)", "Carbs(g)", "Fat(g)", "Time"]
    wd.row_dimensions[6].height = 22
    for ci, h in enumerate(headers):
        c = wd.cell(6, ci + 2, h)
        c.fill = PatternFill("solid", fgColor=C["navy"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")

    meal_colors = {
        "Breakfast": "FFF9E6", "Lunch": "E8F8F5",
        "Dinner": "F0F0FF", "Snack": "FFF0F5"
    }
    for ri, e in enumerate(diet_list):
        row = ri + 7
        wd.row_dimensions[row].height = 20
        bg = meal_colors.get(e.get("meal",""), C["white"])
        vals = [e.get("meal",""), e.get("food",""), e.get("calories",0),
                e.get("protein",0), e.get("carbs",0), e.get("fat",0),
                e.get("time","")]
        for ci, v in enumerate(vals):
            cell = wd.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center" if ci != 1 else "left",
                                        vertical="center")

    if diet_list:
        add_thick_border(wd, 6, 2, 6 + len(diet_list), 8)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 7 – 7-DAY SUMMARY (work + tasks trend)
    # ════════════════════════════════════════════════════════════════════════
    ws7 = wb.create_sheet("📈 7-Day Summary")
    ws7.sheet_view.showGridLines = False
    ws7.column_dimensions["A"].width = 3
    ws7.column_dimensions["B"].width = 16
    ws7.column_dimensions["C"].width = 16
    ws7.column_dimensions["D"].width = 16
    ws7.column_dimensions["E"].width = 16
    ws7.column_dimensions["F"].width = 3

    ws7.row_dimensions[2].height = 45
    ws7.merge_cells("B2:E2")
    c = ws7["B2"]
    c.value = "📈  7-DAY SUMMARY"
    c.fill = PatternFill("solid", fgColor=C["navy"])
    c.font = Font(color=C["white"], bold=True, size=16)
    c.alignment = Alignment(horizontal="center", vertical="center")

    all_tasks = load_json("tasks", {})
    all_work  = load_json("work", {})

    day_labels, day_done, day_total, day_work = [], [], [], []
    for i in range(6, -1, -1):
        day = (date.today() - timedelta(days=i)).isoformat()
        dt = all_tasks.get(day, [])
        dw = all_work.get(day, [])
        day_labels.append(day[5:])  # MM-DD
        day_done.append(sum(1 for t in dt if t.get("done")))
        day_total.append(len(dt))
        day_work.append(sum(w["hours"] for w in dw))

    headers = ["Date", "Tasks Done", "Total Tasks", "Work Hours"]
    ws7.row_dimensions[4].height = 22
    for ci, h in enumerate(headers):
        c = ws7.cell(4, ci + 2, h)
        c.fill = PatternFill("solid", fgColor=C["navy"])
        c.font = Font(color=C["white"], bold=True)
        c.alignment = Alignment(horizontal="center")

    for ri in range(7):
        row = ri + 5
        ws7.row_dimensions[row].height = 20
        bg = C["light"] if ri % 2 == 0 else C["white"]
        vals = [day_labels[ri], day_done[ri], day_total[ri], day_work[ri]]
        for ci, v in enumerate(vals):
            cell = ws7.cell(row, ci + 2, v)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center", vertical="center")

    add_thick_border(ws7, 4, 2, 11, 5)

    # bar chart – tasks
    chart = BarChart()
    chart.type = "col"
    chart.title = "Tasks Completed (7 days)"
    chart.y_axis.title = "Count"
    chart.style = 10
    chart.width = 18
    chart.height = 12
    data_ref  = Reference(ws7, min_col=3, max_col=4, min_row=4, max_row=11)
    cats_ref  = Reference(ws7, min_col=2, min_row=5, max_row=11)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws7.add_chart(chart, "B14")

    # line chart – work hours
    lc = LineChart()
    lc.title = "Work Hours (7 days)"
    lc.y_axis.title = "Hours"
    lc.style = 10
    lc.width = 18
    lc.height = 12
    ldata = Reference(ws7, min_col=5, min_row=4, max_row=11)
    lcats = Reference(ws7, min_col=2, min_row=5, max_row=11)
    lc.add_data(ldata, titles_from_data=True)
    lc.set_categories(lcats)
    ws7.add_chart(lc, "B30")

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 8 – JOURNAL / NOTES
    # ════════════════════════════════════════════════════════════════════════
    wn = wb.create_sheet("📓 Journal")
    wn.sheet_view.showGridLines = False
    wn.column_dimensions["A"].width = 12
    wn.column_dimensions["B"].width = 20
    wn.column_dimensions["C"].width = 14
    wn.column_dimensions["D"].width = 14
    wn.column_dimensions["E"].width = 25
    wn.column_dimensions["F"].width = 60

    wn.merge_cells("A1:F1")
    t = wn["A1"]
    t.value = f"📓 Journal & Notes — {d}"
    t.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=C["purple"])
    t.alignment = Alignment(horizontal="center", vertical="center")
    wn.row_dimensions[1].height = 36

    NOTE_COLORS = {
        "journal":   C["blue"],
        "idea":      C["orange"],
        "gratitude": C["purple"],
        "note":      C["green"],
    }
    MOOD_LABELS = {"": "", "😊": "😊", "😐": "😐", "😔": "😔", "🤩": "🤩", "😤": "😤"}

    hdrs = ["#", "Title", "Type", "Mood", "Tags", "Content"]
    hdr_fills = [C["purple"]] * 6
    for ci, (h, fc) in enumerate(zip(hdrs, hdr_fills), 1):
        c = wn.cell(2, ci, h)
        c.font = Font(name="Calibri", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=fc)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(bottom=Side(style="medium", color="FFFFFF"))

    if notes_list:
        for ri, note in enumerate(notes_list, 1):
            row = ri + 2
            wn.row_dimensions[row].height = 50
            note_type = note.get("type", "note")
            row_color = NOTE_COLORS.get(note_type, C["green"])

            vals = [
                ri,
                note.get("title") or "(untitled)",
                note_type.capitalize(),
                note.get("mood", ""),
                ", ".join(note.get("tags", [])),
                note.get("content", ""),
            ]
            for ci, val in enumerate(vals, 1):
                c = wn.cell(row, ci, val)
                c.font = Font(name="Calibri", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
                c.border = Border(
                    left=Side(style="thin", color="DDDDDD"),
                    right=Side(style="thin", color="DDDDDD"),
                    bottom=Side(style="thin", color="DDDDDD"),
                )
                if ci == 1:
                    c.font = Font(name="Calibri", bold=True, color=row_color)
                    c.alignment = Alignment(horizontal="center", vertical="top")
                if ci == 3:
                    c.font = Font(name="Calibri", bold=True, color=row_color)
    else:
        wn.merge_cells("A3:F3")
        emp = wn["A3"]
        emp.value = "No notes recorded for this date."
        emp.font = Font(name="Calibri", italic=True, color="999999")
        emp.alignment = Alignment(horizontal="center", vertical="center")
        wn.row_dimensions[3].height = 30

    # ── freeze panes & save ────────────────────────────────────────────────
    for sheet in wb.worksheets:
        sheet.freeze_panes = sheet.cell(3, 1)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"ActivityTracker_{d}.xlsx"
    return send_file(buf, as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True, port=5050)
