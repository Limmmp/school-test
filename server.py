"""Тест по программированию для школьников.

Запуск:  python server.py            (по умолчанию 30 минут, порт 8000)
         python server.py --minutes 25 --port 8080
         python server.py --ip       (только вывести IP-адреса и выйти)

Ученики открывают http://<IP-ноутбука>:8000 (ссылка печатается при старте).
Учитель — http://localhost:8000/admin?key=<ключ из консоли>.
Результаты пишутся в results/results.csv (открывается в Excel) и results/results.json.
"""
import argparse
import csv
import html
import json
import random
import secrets
import socket
import sys
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from questions import QUESTIONS, SECTIONS

BASE = Path(__file__).parent
RESULTS_DIR = BASE / "results"
RESULTS_DIR.mkdir(exist_ok=True)
JSON_PATH = RESULTS_DIR / "results.json"
CSV_PATH = RESULTS_DIR / "results.csv"
PROGRESS_PATH = RESULTS_DIR / "progress.json"   # черновики тех, кто ещё решает
SETTINGS_PATH = RESULTS_DIR / "settings.json"   # добавленное время — переживает перезапуск

LOCK = threading.RLock()
BY_ID = {q["id"]: q for q in QUESTIONS}
NUM = {q["id"]: i for i, q in enumerate(QUESTIONS, 1)}
ADMIN_KEY = secrets.token_hex(3)
MINUTES = 30      # базовое время из командной строки
PORT = 8000


def load_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


EXTRA = int(load_json(SETTINGS_PATH, {}).get("extra", 0))  # минуты, добавленные кнопкой
PROGRESS = load_json(PROGRESS_PATH, {})                     # attempt -> черновик
SHORT = {"Лёгкие": "Л", "Средние": "С", "Сложные": "Сл"}
e = html.escape


def load_results():
    if not JSON_PATH.exists():
        return []
    rows = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    for i, r in enumerate(rows):
        r.setdefault("attempt", f"old{i}")
        r.setdefault("marks", {})
    return rows


RESULTS = load_results()


def norm(s):
    return " ".join(str(s).strip().lower().replace("ё", "е").split())


def questions_for():
    """Вопросы для ученика: уровни по порядку, внутри уровня — перемешаны."""
    out = []
    for sec in SECTIONS:
        qs = [q for q in QUESTIONS if q["section"] == sec]
        random.shuffle(qs)
        for q in qs:
            item = {k: q[k] for k in ("id", "type", "text", "section", "points")}
            if "code" in q:
                item["code"] = q["code"]
            if q["type"] == "choice":
                opts = q["options"][:]
                random.shuffle(opts)
                item["options"] = opts
            out.append(item)
    return out


def check(q, given):
    if q["type"] == "code":
        return None if given.strip() else False  # пустой код проверять нечего
    if q["type"] == "choice":
        return given == q["answer"]
    return norm(given) in {norm(a) for a in q["answer"]}


def status(r, x):
    """Итог по заданию: True / False / None (ждёт проверки учителем)."""
    if x["ok"] is not None:
        return x["ok"]
    return r.get("marks", {}).get(x["id"])


def summarize(r):
    score = maxp = pending = 0
    levels = {s: [0, 0] for s in SECTIONS}
    for x in r["detail"]:
        q = BY_ID.get(x["id"])
        if not q:
            continue
        ok = status(r, x)
        maxp += q["points"]
        levels[q["section"]][1] += 1
        if ok is None:
            pending += 1
        elif ok:
            score += q["points"]
            levels[q["section"]][0] += 1
    return {"score": score, "max": maxp, "pending": pending, "levels": levels,
            "percent": round(100 * score / maxp) if maxp else 0}


def sorted_results():
    return sorted(RESULTS, key=lambda r: (r["grade"], r["name"].lower()))


def build_detail(answers):
    detail = []
    for q in QUESTIONS:
        given = str(answers.get(q["id"], ""))[:5000]
        detail.append({"id": q["id"], "given": given, "ok": check(q, given)})
    return detail


def live_record(p):
    """Черновик решающего ученика в том же виде, что и сданная работа."""
    now = time.time() * 1000
    started = p.get("startedAt") or now
    answered = sum(1 for v in p["answers"].values() if str(v).strip())
    return {
        "attempt": p["attempt"], "name": p["name"], "grade": p["grade"], "group": p["group"],
        "ip": p["ip"], "started": fmt_time(started), "finished": "",
        "minutes": round((now - started) / 60000, 1),
        "tab_leaves": p["tab_leaves"], "timeout": False, "marks": {},
        "detail": build_detail(p["answers"]), "live": True,
        "answered": answered, "cur_id": p.get("cur_id", ""), "seen": p["seen"],
        "left": max(0, round((started + (MINUTES + EXTRA) * 60000 - now) / 60000)),
    }


def live_results():
    return sorted((live_record(p) for p in PROGRESS.values()),
                  key=lambda r: (r["grade"], r["name"].lower()))


def find(attempt):
    r = next((r for r in RESULTS if r["attempt"] == attempt), None)
    if r is None and attempt in PROGRESS:
        r = live_record(PROGRESS[attempt])
    return r


def save_progress():
    PROGRESS_PATH.write_text(json.dumps(PROGRESS, ensure_ascii=False), encoding="utf-8")


# ───────────────────────────── файлы ─────────────────────────────

def write_files():
    JSON_PATH.write_text(json.dumps(RESULTS, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        write_csv(CSV_PATH)
    except PermissionError:
        # файл открыт в Excel — ответы уже в JSON, CSV обновится при следующей записи
        print("  ! results.csv открыт в Excel — закройте его, чтобы файл обновлялся")


def csv_bytes():
    tmp = RESULTS_DIR / ".export.csv"
    write_csv(tmp)
    data = tmp.read_bytes()
    tmp.unlink()
    return data


def write_csv(path):
    # utf-8-sig и ";" — чтобы русский Excel открыл файл двойным кликом
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        head = ["Фамилия и имя", "Класс", "Группа", "IP", "Начал", "Сдал", "Минут",
                "Баллы", "Из", "Процент"]
        head += [f"{s} (верно)" for s in SECTIONS]
        head += ["Код не проверен", "Уходы со вкладки", "Сдано по таймеру"]
        for q in QUESTIONS:
            head += [f"№{NUM[q['id']]} ответ", f"№{NUM[q['id']]} верно"]
        w.writerow(head)
        for r in sorted_results():
            s = summarize(r)
            row = [r["name"], r["grade"], r["group"], r.get("ip", ""), r["started"],
                   r["finished"], r["minutes"], s["score"], s["max"], s["percent"]]
            row += [f"{a} из {b}" for a, b in s["levels"].values()]
            row += [s["pending"], r["tab_leaves"], "да" if r["timeout"] else "нет"]
            d = {x["id"]: x for x in r["detail"]}
            for q in QUESTIONS:
                x = d.get(q["id"])
                if not x:
                    row += ["", ""]
                    continue
                ok = status(r, x)
                row += [x["given"], "проверить" if ok is None else ("1" if ok else "0")]
            w.writerow(row)


def lan_ips():
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ips.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass

    # адреса домашней/школьной Wi-Fi — первыми, VPN и прочие — в конце
    def rank(ip):
        a, b = (int(x) for x in ip.split(".")[:2])
        if a == 192 and b == 168:
            return 0
        if a == 10:
            return 1
        if a == 172 and 16 <= b <= 31:
            return 2
        return 3

    return sorted((ip for ip in ips if not ip.startswith(("127.", "169.254."))), key=rank)


def fmt_time(ts):
    return datetime.fromtimestamp(ts / 1000).strftime("%H:%M:%S") if ts else ""


# ───────────────────────────── страницы учителя ─────────────────────────────

ADMIN_CSS = """
:root{--ground:#f1f4f2;--surface:#fff;--sunk:#f6f8f7;--ink:#18211d;--muted:#5a6a63;--line:#d4ddd8;
--accent:#1f6f54;--accent-ink:#fff;--accent-soft:#e1efe8;
--ok:#1f7a4d;--ok-soft:#e3f3ea;--bad:#b3261e;--bad-soft:#fbeae8;--chk:#8a5d00;--chk-soft:#fbf1dc;
--code-bg:#1b2420;--code-ink:#e2ece6}
@media (prefers-color-scheme: dark){:root{--ground:#121816;--surface:#1a2320;--sunk:#151d1a;--ink:#e5ede9;
--muted:#95a69e;--line:#2d3a35;--accent:#4fb58c;--accent-ink:#0d1512;--accent-soft:#1f3a30;
--ok:#5fd19a;--ok-soft:#173226;--bad:#ff8a7e;--bad-soft:#3a1e1b;--chk:#e8b84a;--chk-soft:#352a12;--code-bg:#0c1210}}
*{box-sizing:border-box}
body{font:14px/1.5 "Segoe UI",system-ui,sans-serif;margin:0;padding-inline:24px;padding-block:0 60px;
background:var(--ground);color:var(--ink)}
a{color:inherit}
nav{display:flex;gap:4px;flex-wrap:wrap;padding-block:14px;border-bottom:1px solid var(--line);margin-bottom:20px}
nav a{padding:8px 14px;border-radius:6px;text-decoration:none;font-weight:600}
nav a.on{background:var(--accent-soft)}
.picker{margin-left:auto;font:inherit;color:var(--ink);background:var(--surface);border:1px solid var(--line);
border-radius:6px;padding:7px 10px;max-width:100%;min-width:0}
.live-dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--ok);margin-right:6px}
.live-dot.stale{background:var(--muted)}
.prog{display:inline-block;width:90px;height:6px;border-radius:3px;background:var(--line);vertical-align:middle;margin-right:6px;overflow:hidden}
.prog i{display:block;height:100%;background:var(--accent)}
nav a:hover{background:var(--surface)}
h1{font-size:22px;margin:0 0 4px;text-wrap:balance} h2{font-size:18px;margin:28px 0 10px}
h3{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:28px 0 10px}
.sub{color:var(--muted);margin:0 0 16px}
.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:16px}
.btn{display:inline-block;background:var(--surface);color:var(--ink);border:1px solid var(--line);
padding:7px 14px;border-radius:6px;text-decoration:none;font:inherit;font-weight:600;cursor:pointer}
.btn.primary,.btn.on{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}
.btn.good{background:var(--ok-soft);color:var(--ok);border-color:var(--ok)}
.btn.nope{background:var(--bad-soft);color:var(--bad);border-color:var(--bad)}
.wrap{overflow-x:auto;background:var(--surface);border:1px solid var(--line);border-radius:8px}
table{border-collapse:collapse;width:100%}
th,td{padding:9px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:middle}
th{font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);white-space:nowrap}
tbody tr:hover{background:var(--sunk)}
.num{font-variant-numeric:tabular-nums;white-space:nowrap}
.warn{color:var(--bad);font-weight:700}
.name a{font-weight:600;text-decoration:none;border-bottom:1px solid var(--line)}
.name a:hover{border-color:var(--ink)}
.strip{white-space:nowrap}
.strip a{display:inline-block;width:14px;height:18px;margin-right:2px;border-radius:3px;vertical-align:middle;text-decoration:none}
.strip .gap{display:inline-block;width:8px}
.s-ok{background:var(--ok)} .s-bad{background:var(--bad);opacity:.8} .s-chk{background:var(--chk);opacity:.7}
.s-empty{background:var(--line)}
.pill{display:inline-block;padding:1px 8px;border-radius:999px;font-size:12px;font-weight:600;white-space:nowrap}
.p-ok{background:var(--ok-soft);color:var(--ok)} .p-bad{background:var(--bad-soft);color:var(--bad)}
.p-chk{background:var(--chk-soft);color:var(--chk)} .p-lvl{background:var(--sunk);color:var(--muted);border:1px solid var(--line)}
pre{background:var(--code-bg);color:var(--code-ink);padding:10px 12px;border-radius:6px;white-space:pre-wrap;
word-break:break-word;font:13px/1.5 Consolas,"Cascadia Mono",monospace;margin:0}
.empty{color:var(--muted);text-align:center;padding:32px}
.card{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--line);
border-radius:8px;padding:14px 16px;margin-bottom:12px;max-width:980px;scroll-margin-top:16px}
.card.ok{border-left-color:var(--ok)} .card.bad{border-left-color:var(--bad)} .card.chk{border-left-color:var(--chk)}
.card:target{outline:2px solid var(--accent)}
.ch{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:6px}
.qn{font-weight:700;font-variant-numeric:tabular-nums}
.qt{font-weight:600;margin:0 0 8px;max-width:75ch}
.cq{margin:0 0 10px}
.cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px;margin-top:6px}
.lbl{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:0 0 4px}
.given{background:var(--sunk);border:1px solid var(--line);border-radius:6px;padding:8px 10px;
font:14px/1.5 Consolas,"Cascadia Mono",monospace;white-space:pre-wrap;word-break:break-word;min-height:38px}
.given.ok{background:var(--ok-soft);border-color:var(--ok)} .given.bad{background:var(--bad-soft);border-color:var(--bad)}
.given.none{color:var(--muted);font-style:italic;font-family:inherit}
.right{background:var(--ok-soft);border:1px solid var(--ok);border-radius:6px;padding:8px 10px;
font:14px/1.5 Consolas,"Cascadia Mono",monospace;white-space:pre-wrap}
.opts{list-style:none;padding:0;margin:0;display:grid;gap:4px}
.opts li{padding:5px 10px;border-radius:6px;border:1px solid var(--line);display:flex;gap:8px;align-items:baseline}
.opts li.correct{border-color:var(--ok);background:var(--ok-soft)}
.opts li.picked-bad{border-color:var(--bad);background:var(--bad-soft)}
.opts .tag{margin-left:auto;font-size:12px;font-weight:600;white-space:nowrap}
.exp{border-left:3px solid var(--accent);padding:2px 0 2px 12px;margin:12px 0 0;max-width:75ch;color:var(--ink)}
.marker{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:12px;padding-top:12px;border-top:1px dashed var(--line)}
.marker form{margin:0}
.timebar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;background:var(--surface);
border:1px solid var(--line);border-radius:8px;padding:10px 14px;margin-bottom:18px;width:fit-content;max-width:100%}
.timebar form{margin:0}
.timebar b{font-variant-numeric:tabular-nums}
.summary{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0 18px}
.tile{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:10px 14px;min-width:120px}
.tile b{display:block;font-size:22px;font-variant-numeric:tabular-nums}
.tile span{font-size:12px;color:var(--muted)}
.pager{display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;margin-bottom:14px}
details.others{margin-top:12px}
details.others summary{cursor:pointer;color:var(--muted);font-weight:600}
.others table{margin-top:8px}
.others td{vertical-align:top}
.ipbox{display:grid;gap:18px;padding-block:6vh}
.ip{font:700 clamp(40px,8vw,110px)/1.1 Consolas,"Cascadia Mono",monospace;word-break:break-all}
.ip.second{font-size:clamp(20px,3vw,34px);color:var(--muted)}
.big-lead{font-size:clamp(18px,2.4vw,28px);color:var(--muted);margin:0}
"""


def url(path, **params):
    q = "&".join(f"{k}={quote(str(v))}" for k, v in params.items() if v not in (None, ""))
    return f"{path}?key={ADMIN_KEY}" + (f"&{q}" if q else "")


def student_picker(current):
    """Выпадающий список всех учеников — есть на каждой странице админки."""
    with LOCK:
        groups = [("Сейчас решают", live_results()), ("Сдали", sorted_results())]
    opts = ['<option value="">Открыть ученика…</option>']
    for title, rows in groups:
        if not rows:
            continue
        opts.append(f'<optgroup label="{title} ({len(rows)})">')
        for r in rows:
            sel = " selected" if r["attempt"] == current else ""
            extra = f' · {r["answered"]}/{len(QUESTIONS)}' if r.get("live") else f' · {summarize(r)["score"]} б.'
            opts.append(f'<option value="{url("/admin/student", id=r["attempt"])}"{sel}>'
                        f'{e(r["name"])} · {r["grade"]} кл.{extra}</option>')
        opts.append("</optgroup>")
    return (f'<select class="picker" aria-label="Открыть ученика" '
            f'onchange="if(this.value)location.href=this.value">{"".join(opts)}</select>')


def admin_shell(title, active, body, refresh=False, current=None):
    links = [("results", "Результаты", "/admin"),
             ("answers", "Разбор заданий", "/admin/answers"),
             ("screen", "Адрес для учеников", "/admin/screen")]
    nav = "".join(
        f'<a href="{url(href)}"{" class=on" if k == active else ""}>{label}</a>'
        for k, label, href in links) + student_picker(current)
    meta = '<meta http-equiv="refresh" content="15">' if refresh else ""
    return f"""<!doctype html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">{meta}
<title>{e(title)}</title><style>{ADMIN_CSS}</style></head><body>
<nav>{nav}</nav>{body}</body></html>"""


def state_of(ok, given, qtype, marked):
    """(css-класс, текст плашки)"""
    if ok is None:
        return "chk", "Ждёт проверки"
    if not given.strip():
        return "bad", "Нет ответа"
    if qtype == "code":
        return ("ok", "Засчитано") if ok else ("bad", "Не засчитано")
    return ("ok", "Верно") if ok else ("bad", "Неверно")


def results_page():
    with LOCK:
        rows = sorted_results()
        trs = []
        for r in rows:
            s = summarize(r)
            strip, prev = [], None
            for x in r["detail"]:
                q = BY_ID.get(x["id"])
                if not q:
                    continue
                if prev and q["section"] != prev:
                    strip.append('<span class="gap"></span>')
                prev = q["section"]
                ok = status(r, x)
                cls = "s-chk" if ok is None else ("s-ok" if ok else ("s-bad" if x["given"].strip() else "s-empty"))
                tip = f'№{NUM[q["id"]]}: {x["given"][:60] or "нет ответа"}'
                strip.append(f'<a class="{cls}" href="{url("/admin/student", id=r["attempt"])}#{q["id"]}" title="{e(tip)}"></a>')
            lv = " · ".join(f"{SHORT[k]} {a}/{b}" for k, (a, b) in s["levels"].items())
            pend = (f'<a class="pill p-chk" href="{url("/admin/student", id=r["attempt"], show="pending")}">'
                    f'{s["pending"]} на проверке</a>') if s["pending"] else ""
            warn = "num warn" if r["tab_leaves"] >= 3 else "num"
            trs.append(f"""<tr>
<td class="name"><a href="{url('/admin/student', id=r['attempt'])}">{e(r['name'])}</a></td>
<td>{r['grade']}</td><td>{e(r['group'])}</td>
<td class="num"><b>{s['score']}</b> / {s['max']}</td><td class="num">{s['percent']}%</td>
<td class="num">{lv}</td><td>{pend}</td>
<td class="strip">{''.join(strip)}</td>
<td class="num">{r['minutes']}</td><td class="{warn}">{r['tab_leaves']}</td>
<td class="num">{e(r.get('ip', ''))}</td></tr>""")
    body = "".join(trs) or '<tr><td colspan="11" class="empty">Пока никто не сдал. Страница обновляется сама.</td></tr>'
    with LOCK:
        live = live_results()
    ltrs = []
    for r in live:
        ago = round(time.time() - r["seen"])
        pc = round(100 * r["answered"] / len(QUESTIONS))
        cur = f'№{NUM[r["cur_id"]]}' if r["cur_id"] in NUM else "—"
        warn = "num warn" if r["tab_leaves"] >= 3 else "num"
        ltrs.append(f"""<tr>
<td class="name"><span class="live-dot{' stale' if ago > 60 else ''}"></span><a href="{url('/admin/student', id=r['attempt'])}">{e(r['name'])}</a></td>
<td>{r['grade']}</td><td>{e(r['group'])}</td>
<td class="num"><span class="prog"><i style="width:{pc}%"></i></span>{r['answered']} / {len(QUESTIONS)}</td>
<td class="num">{cur}</td><td class="num">{r['left']} мин</td>
<td class="{warn}">{r['tab_leaves']}</td>
<td class="num">{'только что' if ago < 15 else f'{ago} с назад'}</td>
<td class="num">{e(r['ip'])}</td></tr>""")
    live_html = ""
    if ltrs:
        live_html = f"""<h2>Сейчас решают — {len(live)}</h2>
<p class="sub">Нажмите на имя, чтобы посмотреть ответы, которые ученик уже дал. Серая точка — телефон больше минуты не выходил на связь.</p>
<div class="wrap"><table><thead><tr><th>Ученик</th><th>Класс</th><th>Группа</th><th>Ответил</th><th>Сейчас на</th>
<th>Осталось</th><th>Уходы</th><th>На связи</th><th>IP</th></tr></thead><tbody>{''.join(ltrs)}</tbody></table></div>
<h2>Сдали — {len(rows)}</h2>"""
    added = f" (добавлено {EXTRA:+d})" if EXTRA else ""
    minus = time_form(-5, "−5 мин") if EXTRA > 0 else ""
    timebar = f"""<div class="timebar"><span>Время на тест: <b>{MINUTES + EXTRA} мин</b>{added}</span>
{time_form(5, "+5 минут всем", "primary")}{minus}</div>"""
    return admin_shell("Результаты теста", "results", f"""
{timebar}
<h1>Результаты</h1>
{live_html}
<p class="sub">Сдали: {len(rows)} · обновляется каждые 15 секунд · нажмите на имя, чтобы открыть все ответы ученика ·
в полоске: зелёный — верно, красный — неверно, серый — нет ответа, жёлтый — код ждёт проверки (квадратик кликабельный) ·
баллы: лёгкое 1, среднее 2, сложное 3 · «Уходы» красным — ушёл со вкладки 3+ раза</p>
<div class="bar"><a class="btn primary" href="{url('/results.csv')}">Скачать для Excel (CSV)</a>
<span class="sub" style="margin:0">Файл также сохраняется здесь: {e(str(CSV_PATH))}</span></div>
<div class="wrap"><table><thead><tr><th>Ученик</th><th>Класс</th><th>Группа</th><th>Баллы</th><th>%</th>
<th>Верно по уровням</th><th>Код</th><th>Задания</th><th>Мин.</th><th>Уходы</th><th>IP</th></tr></thead>
<tbody>{body}</tbody></table></div>""", refresh=True)


def time_form(delta, label, cls=""):
    return (f'<form method="post" action="/admin/time">'
            f'<input type="hidden" name="key" value="{ADMIN_KEY}">'
            f'<input type="hidden" name="delta" value="{delta}">'
            f'<button class="btn {cls}">{label}</button></form>')


def answer_block(q, x, ok):
    """Ответ ученика рядом с правильным."""
    given = x["given"] if x else ""
    if q["type"] == "choice":
        lis = []
        for o in q["options"]:
            if o == q["answer"] and o == given:
                cls, tag = "correct", '<span class="tag" style="color:var(--ok)">✓ выбрал, верно</span>'
            elif o == q["answer"]:
                cls, tag = "correct", '<span class="tag" style="color:var(--ok)">правильный ответ</span>'
            elif o == given:
                cls, tag = "picked-bad", '<span class="tag" style="color:var(--bad)">✗ выбрал ученик</span>'
            else:
                cls, tag = "", ""
            lis.append(f'<li class="{cls}"><span>{e(o)}</span>{tag}</li>')
        none = '' if given else '<p class="given none" style="margin:0 0 6px">Ученик ничего не выбрал</p>'
        return f'{none}<ul class="opts">{"".join(lis)}</ul>'
    if q["type"] == "text":
        g = (f'<div class="given {"ok" if ok else "bad"}">{e(given)}</div>' if given.strip()
             else '<div class="given none">нет ответа</div>')
        right = " или ".join(e(a) for a in q["answer"])
        return (f'<div class="cols"><div><p class="lbl">Ответ ученика</p>{g}</div>'
                f'<div><p class="lbl">Правильный ответ</p><div class="right">{right}</div></div></div>')
    g = f"<pre>{e(given)}</pre>" if given.strip() else '<div class="given none">ученик ничего не написал</div>'
    return (f'<div class="cols"><div><p class="lbl">Код ученика</p>{g}</div>'
            f'<div><p class="lbl">Пример решения</p><pre>{e(q["solution"])}</pre></div></div>')


def mark_form(r, q, value, label, cls, show):
    return (f'<form method="post" action="/admin/mark">'
            f'<input type="hidden" name="key" value="{ADMIN_KEY}">'
            f'<input type="hidden" name="id" value="{e(r["attempt"])}">'
            f'<input type="hidden" name="qid" value="{q["id"]}">'
            f'<input type="hidden" name="v" value="{value}">'
            f'<input type="hidden" name="show" value="{e(show)}">'
            f'<button class="btn {cls}">{label}</button></form>')


def student_page(attempt, show):
    with LOCK:
        r = find(attempt)
        if not r:
            return admin_shell("Нет такого ученика", "results",
                               f'<h1>Ученик не найден</h1><p><a href="{url("/admin")}">← к результатам</a></p>')
        live = r.get("live", False)
        rows = live_results() if live else sorted_results()
        s = summarize(r)
        idx = [x["attempt"] for x in rows].index(attempt)
        prev_r = rows[idx - 1] if idx > 0 else None
        next_r = rows[idx + 1] if idx < len(rows) - 1 else None
        d = {x["id"]: x for x in r["detail"]}
        cards, prev_sec, shown = [], None, 0
        counts = {"all": 0, "wrong": 0, "pending": 0}
        for q in QUESTIONS:
            x = d.get(q["id"])
            if not x:
                continue
            ok = status(r, x)
            counts["all"] += 1
            counts["wrong"] += ok is False
            counts["pending"] += ok is None
            if show == "wrong" and ok is not False:
                continue
            if show == "pending" and not (q["type"] == "code"):
                continue
            if q["section"] != prev_sec:
                cards.append(f"<h3>{e(q['section'])} · {q['points']} б. за задание</h3>")
                prev_sec = q["section"]
            cls, label = state_of(ok, x["given"], q["type"], q["id"] in r["marks"])
            if live and not x["given"].strip():
                cls, label = "", "Пока не ответил"
            if live and r["cur_id"] == q["id"]:
                label += " · сейчас здесь"
            code = f'<pre class="cq">{e(q["code"])}</pre>' if q.get("code") else ""
            marker = ""
            if q["type"] == "code" and x["given"].strip() and not live:
                cur = r["marks"].get(q["id"])
                marker = ('<div class="marker"><b>Оценка:</b>'
                          + mark_form(r, q, "1", f"Засчитать (+{q['points']})", "good" if cur is True else "", show)
                          + mark_form(r, q, "0", "Не засчитывать", "nope" if cur is False else "", show)
                          + (mark_form(r, q, "", "Сбросить", "", show) if cur is not None else "")
                          + "</div>")
            shown += 1
            cards.append(f"""<div class="card {cls}" id="{q['id']}">
<div class="ch"><span class="qn">№{NUM[q['id']]}</span><span class="pill p-lvl">{e(q['section'])}</span>
<span class="pill p-{cls}">{label}</span></div>
<p class="qt">{e(q['text'])}</p>{code}{answer_block(q, x, ok)}
<p class="exp">{e(q['explain'])}</p>{marker}</div>""")
    if not shown:
        cards.append('<p class="empty">В этом фильтре заданий нет.</p>')
    filters = "".join(
        f'<a class="btn{" on" if show == k else ""}" href="{url("/admin/student", id=attempt, show=k if k != "all" else "")}">{t} ({counts[k]})</a>'
        for k, t in (("all", "Все"), ("wrong", "Ошибки"), ("pending", "Код")))
    pager = (f'<a class="btn" href="{url("/admin/student", id=prev_r["attempt"], show=show)}">← {e(prev_r["name"])}</a>' if prev_r else "<span></span>")
    pager += (f'<a class="btn" href="{url("/admin/student", id=next_r["attempt"], show=show)}">{e(next_r["name"])} →</a>' if next_r else "<span></span>")
    lv = "".join(f'<div class="tile"><b>{a} / {b}</b><span>{e(k)}</span></div>' for k, (a, b) in s["levels"].items())
    pend = f'<div class="tile"><b style="color:var(--chk)">{s["pending"]}</b><span>кода не проверено</span></div>' if s["pending"] else ""
    leaves = "warn" if r["tab_leaves"] >= 3 else ""
    if live:
        ago = round(time.time() - r["seen"])
        state = (f'<span class="pill p-chk"><span class="live-dot{" stale" if ago > 60 else ""}"></span>'
                 f'решает сейчас · ответил на {r["answered"]} из {len(QUESTIONS)} · осталось ~{r["left"]} мин · '
                 f'был на связи {ago} с назад</span>')
        when = f"начал в {r['started']} · решает {r['minutes']} мин."
        first_tile = f'<div class="tile"><b>{s["score"]}</b><span>баллов на данный момент</span></div>'
        note = ('<p class="sub">Это черновик: ответы обновляются, пока ученик решает (страница обновляется сама). '
                'Код можно будет засчитать после сдачи.</p>')
    else:
        state = ""
        when = f"сдал в {r['finished']} · {r['minutes']} мин.{' · по таймеру' if r['timeout'] else ''}"
        first_tile = f'<div class="tile"><b>{s["score"]} / {s["max"]}</b><span>баллов · {s["percent"]}%</span></div>'
        note = ""
    return admin_shell(f"{r['name']} — ответы", "results", f"""
<div class="pager">{pager}</div>
<h1>{e(r['name'])}</h1>
<p class="sub">{r['grade']} класс{(' · ' + e(r['group'])) if r['group'] else ''} · {when}
· <span class="{leaves}">уходов со вкладки: {r['tab_leaves']}</span> · IP {e(r.get('ip', '—'))}</p>
{state}{note}
<div class="summary">{first_tile}{lv}{pend if not live else ''}</div>
<div class="bar">{filters}</div>
{''.join(cards)}""", refresh=live, current=attempt)


def answers_page(grade):
    with LOCK:
        rows = [r for r in sorted_results() if grade is None or r["grade"] == grade]
        parts, prev_sec = [], None
        for q in QUESTIONS:
            if q["section"] != prev_sec:
                parts.append(f"<h3>{e(q['section'])} · {q['points']} б. за задание</h3>")
                prev_sec = q["section"]
            given_rows, ok_n, tot = [], 0, 0
            for r in rows:
                x = next((y for y in r["detail"] if y["id"] == q["id"]), None)
                if not x:
                    continue
                ok = status(r, x)
                tot += 1
                ok_n += bool(ok)
                cls, label = state_of(ok, x["given"], q["type"], False)
                g = x["given"]
                shown = (f"<pre>{e(g)}</pre>" if q["type"] == "code" and g.strip()
                         else (f'<span class="given" style="display:inline-block;min-height:0;padding:2px 8px">{e(g)}</span>' if g.strip()
                               else '<span style="color:var(--muted)">—</span>'))
                given_rows.append(
                    f'<tr><td class="name"><a href="{url("/admin/student", id=r["attempt"])}#{q["id"]}">{e(r["name"])}</a>'
                    f' <span style="color:var(--muted)">{r["grade"]}</span></td>'
                    f'<td><span class="pill p-{cls}">{label}</span></td><td>{shown}</td></tr>')
            if q["type"] == "choice":
                lis = "".join(f'<li class="{"correct" if o == q["answer"] else ""}"><span>{e(o)}</span>'
                              f'{"<span class=tag style=color:var(--ok)>правильный ответ</span>" if o == q["answer"] else ""}</li>'
                              for o in q["options"])
                ans = f'<ul class="opts">{lis}</ul>'
            elif q["type"] == "text":
                ans = f'<p class="lbl">Правильный ответ</p><div class="right">{" или ".join(e(a) for a in q["answer"])}</div>'
            else:
                ans = f'<p class="lbl">Пример решения</p><pre>{e(q["solution"])}</pre>'
            if q["type"] == "code":
                stat = f'<span class="pill p-lvl">засчитано {ok_n} из {tot}</span>' if tot else ""
            elif tot:
                pc = round(100 * ok_n / tot)
                stat = f'<span class="pill {"p-ok" if pc >= 60 else "p-bad"}">решили {ok_n} из {tot} · {pc}%</span>'
            else:
                stat = '<span class="pill p-lvl">пока нет ответов</span>'
            code = f'<pre class="cq">{e(q["code"])}</pre>' if q.get("code") else ""
            others = (f'<details class="others"><summary>Ответы учеников ({tot})</summary>'
                      f'<div class="wrap"><table><tbody>{"".join(given_rows)}</tbody></table></div></details>'
                      if given_rows else "")
            parts.append(f"""<div class="card" id="{q['id']}"><div class="ch"><span class="qn">№{NUM[q['id']]}</span>
<span class="pill p-lvl">{e(q['section'])}</span>{stat}</div>
<p class="qt">{e(q['text'])}</p>{code}{ans}<p class="exp">{e(q['explain'])}</p>{others}</div>""")
    tabs = "".join(
        f'<a class="btn{" on" if grade == g else ""}" href="{url("/admin/answers", grade=g or "")}">{t}</a>'
        for g, t in ((None, "Все классы"), (10, "10 класс"), (11, "11 класс")))
    return admin_shell("Разбор заданий", "answers", f"""
<h1>Разбор заданий</h1>
<p class="sub">Правильные ответы, объяснения и ответы всех учеников по каждому заданию. Ученики эту страницу не видят.</p>
<div class="bar">{tabs}</div>{''.join(parts)}""")


def screen_page():
    ips = lan_ips()
    if ips:
        first = f'<div class="ip">{e(ips[0])}:{PORT}</div>'
        rest = "".join(f'<div class="ip second">или {e(ip)}:{PORT}</div>' for ip in ips[1:])
    else:
        first, rest = '<div class="ip">нет сети</div>', ""
    return admin_shell("Адрес для учеников", "screen", f"""
<div class="ipbox"><p class="big-lead">Откройте в браузере телефона:</p>
{first}{rest}</div>""")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def send(self, code, body, ctype):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def html(self, body):
        self.send(200, body, "text/html; charset=utf-8")

    def json(self, obj, code=200):
        self.send(code, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8")

    def body(self):
        length = min(int(self.headers.get("Content-Length", 0)), 200_000)
        return self.rfile.read(length).decode("utf-8")

    def do_GET(self):
        u = urlparse(self.path)
        qs = parse_qs(u.query)
        one = lambda k, d="": qs.get(k, [d])[0]
        if u.path in ("/", "/index.html"):
            return self.html((BASE / "index.html").read_text(encoding="utf-8"))
        if u.path == "/api/config":
            return self.json({"minutes": MINUTES + EXTRA, "count": len(QUESTIONS)})
        if u.path == "/api/questions":
            return self.json({"questions": questions_for(), "minutes": MINUTES + EXTRA})
        if u.path.startswith("/admin") or u.path == "/results.csv":
            if one("key") != ADMIN_KEY:
                return self.send(403, "Нужен ключ учителя: /admin?key=...", "text/plain; charset=utf-8")
        if u.path == "/admin":
            return self.html(results_page())
        if u.path == "/admin/student":
            show = one("show", "all")
            return self.html(student_page(one("id"), show if show in ("all", "wrong", "pending") else "all"))
        if u.path == "/admin/answers":
            g = one("grade")
            return self.html(answers_page(int(g) if g in ("10", "11") else None))
        if u.path == "/admin/screen":
            return self.html(screen_page())
        if u.path == "/results.csv":
            with LOCK:
                data = csv_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            name = f"results_{datetime.now():%Y-%m-%d_%H-%M}.csv"
            self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send(404, "Не найдено", "text/plain; charset=utf-8")

    def do_POST(self):
        u = urlparse(self.path)
        if u.path == "/admin/mark":
            return self.mark()
        if u.path == "/admin/time":
            return self.add_time()
        if u.path not in ("/api/submit", "/api/progress"):
            return self.send(404, "Не найдено", "text/plain; charset=utf-8")
        try:
            sub = json.loads(self.body())
            grade = int(sub.get("grade", 0))
            started = int(sub.get("startedAt", 0) or 0)
            tab_leaves = int(sub.get("tabLeaves", 0) or 0)
        except (ValueError, TypeError, json.JSONDecodeError):
            return self.json({"error": "Не удалось прочитать ответы"}, 400)
        name = str(sub.get("name", "")).strip()[:80]
        if not name or grade not in (10, 11):
            return self.json({"error": "Не указаны имя или класс"}, 400)
        attempt = str(sub.get("attempt", "")).strip()[:40] or secrets.token_hex(6)
        answers = sub.get("answers", {}) if isinstance(sub.get("answers"), dict) else {}
        answers = {k: str(v)[:5000] for k, v in answers.items() if k in BY_ID}
        group = str(sub.get("group", "")).strip()[:40]

        if u.path == "/api/progress":
            with LOCK:
                if any(r["attempt"] == attempt for r in RESULTS):
                    return self.json({"ok": True, "submitted": True})
                PROGRESS[attempt] = {
                    "attempt": attempt, "name": name, "grade": grade, "group": group,
                    "ip": self.client_address[0], "startedAt": started, "answers": answers,
                    "tab_leaves": tab_leaves, "cur_id": str(sub.get("curId", ""))[:10],
                    "seen": time.time(),
                }
                save_progress()
            return self.json({"ok": True})

        detail = build_detail(answers)
        now = int(time.time() * 1000)
        rec = {
            "attempt": attempt,
            "name": name,
            "grade": grade,
            "group": group,
            "ip": self.client_address[0],
            "started": fmt_time(started),
            "finished": fmt_time(now),
            "minutes": round((now - started) / 60000, 1) if started else "",
            "tab_leaves": tab_leaves,
            "timeout": bool(sub.get("timeout")),
            "detail": detail,
            "marks": {},
        }
        with LOCK:
            # повторная отправка той же попытки не создаёт дубль
            if any(r["attempt"] == attempt for r in RESULTS):
                return self.json({"ok": True, "duplicate": True})
            RESULTS.append(rec)
            write_files()
            if PROGRESS.pop(attempt, None) is not None:
                save_progress()
            s = summarize(rec)
        print(f"  ✓ {rec['finished']}  {name} ({grade} кл., {rec['ip']}) — {s['score']}/{s['max']}")
        self.json({"ok": True})

    def add_time(self):
        global EXTRA
        form = parse_qs(self.body())
        if form.get("key", [""])[0] != ADMIN_KEY:
            return self.send(403, "Нет доступа", "text/plain; charset=utf-8")
        try:
            delta = int(form.get("delta", ["0"])[0])
        except ValueError:
            delta = 0
        if delta in (5, -5):
            EXTRA = max(0, EXTRA + delta)
            SETTINGS_PATH.write_text(json.dumps({"extra": EXTRA}), encoding="utf-8")
            print(f"  ⏱ время на тест: {MINUTES + EXTRA} мин")
        self.send_response(303)
        self.send_header("Location", url("/admin"))
        self.send_header("Content-Length", "0")
        self.end_headers()

    def mark(self):
        form = parse_qs(self.body(), keep_blank_values=True)
        one = lambda k: form.get(k, [""])[0]
        if one("key") != ADMIN_KEY:
            return self.send(403, "Нет доступа", "text/plain; charset=utf-8")
        attempt, qid, v, show = one("id"), one("qid"), one("v"), one("show")
        with LOCK:
            r = find(attempt)
            if r and not r.get("live") and qid in BY_ID and BY_ID[qid]["type"] == "code":
                if v == "":
                    r["marks"].pop(qid, None)
                else:
                    r["marks"][qid] = v == "1"
                write_files()
        self.send_response(303)
        self.send_header("Location", url("/admin/student", id=attempt,
                                         show=show if show in ("wrong", "pending") else "") + f"#{qid}")
        self.send_header("Content-Length", "0")
        self.end_headers()


def main():
    global MINUTES, PORT
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    p = argparse.ArgumentParser()
    p.add_argument("--port", type=int, default=8000)
    p.add_argument("--minutes", type=int, default=30)
    p.add_argument("--ip", action="store_true", help="только вывести IP-адреса и выйти")
    a = p.parse_args()
    if a.ip:
        for ip in lan_ips():
            print(ip)
        return
    MINUTES, PORT = a.minutes, a.port
    srv = ThreadingHTTPServer(("0.0.0.0", a.port), Handler)
    ips = lan_ips()
    added = f" (в том числе добавлено кнопкой: {EXTRA})" if EXTRA else ""
    print(f"\n  Тест запущен: {len(QUESTIONS)} заданий, {MINUTES + EXTRA} мин.{added}\n")
    print("  Ссылка для учеников (телефоны в той же Wi-Fi сети):")
    for ip in ips:
        print(f"      http://{ip}:{a.port}")
    print("\n  Страница учителя (только на этом ноутбуке):")
    print(f"      http://localhost:{a.port}/admin?key={ADMIN_KEY}")
    print(f"      адрес крупно для проектора: http://localhost:{a.port}/admin/screen?key={ADMIN_KEY}")
    print(f"\n  Результаты сохраняются в {CSV_PATH}")
    print("  Остановить: Ctrl+C\n")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  Остановлено.")


if __name__ == "__main__":
    main()
