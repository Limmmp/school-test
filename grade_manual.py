# -*- coding: utf-8 -*-
"""Разбор бумажных работ + результатов дашборда → Excel.

Бумажные работы ученики писали по номерам заданий. У большинства порядок один и тот же
(CLUSTER_ORDER) — он восстановлен по тем, кто писал ответы словами, а не номерами вариантов.
Там, где ученик написал только номер варианта (1–4), проверить ответ невозможно: порядок
вариантов на экране не сохраняется. Такие ответы помечены «?» и не входят в баллы.
"""
import json
import sys
from pathlib import Path

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))
from questions import QUESTIONS, POINTS  # noqa: E402

BY_ID = {q["id"]: q for q in QUESTIONS}
NUM = {q["id"]: i for i, q in enumerate(QUESTIONS, 1)}

OK, BAD, NONE, UNK = "✓", "✗", "—", "?"

# порядок заданий на экране, по которому отвечало большинство
CLUSTER_ORDER = ["e11", "e9", "e7", "e10", "e5", "e6", "e12", "e8", "e13", "e14",
                 "e4", "e2", "e3", "e15", "e1", "m8", "m4", "m10", "m3", "m7",
                 "m2", "m9", "m6", "m1", "m5", "h1", "h2", "h3", "h4", "h5"]

# ученики, отвечавшие по этому порядку: номер задания -> что написал
CLUSTER = {
    "Галицких": {1: "4", 2: "4", 3: "ab", 4: "", 5: "3, 1", 6: "32", 7: "", 8: "6", 9: "B",
                 10: "", 11: "11", 12: "3", 13: "14", 14: "3", 15: "2", 16: "6", 17: "", 18: "",
                 19: "7", 20: "10", 21: "3", 22: "4", 23: "", 24: "программа [2:5]", 25: "2",
                 26: "", 27: ""},
    "Хакимов, Малиновский": {1: "", 2: "4", 3: "ab", 4: "147", 5: "3,1", 6: "32", 7: "1", 8: "6",
                             9: "B", 10: "", 11: "11", 12: "3", 13: "14", 14: "4", 15: "4",
                             16: "6", 17: "", 18: "", 19: "7", 20: "10", 21: "3", 22: "4",
                             23: "8", 24: "гра", 25: "2", 26: "[1,2]", 27: "", 28: "", 29: "", 30: ""},
    "Дёнина, Сёмкина, Зацепин, Бабушкин": {1: "4", 2: "4", 3: "ab", 4: "1 4 7", 5: "3,1", 6: "32",
                                           7: "5", 8: "6", 9: "8", 10: "2", 11: "11", 12: "3",
                                           13: "14", 14: "3", 15: "4", 16: "5", 17: "4", 18: "",
                                           19: "6", 20: "3", 21: "21", 22: "4", 23: "8", 24: "гра",
                                           25: "2", 26: "[1, 2]", 27: "", 28: "", 29: "", 30: ""},
    "Чистякова": {1: "3", 2: "4", 3: "ababab", 4: "147", 5: "3,1", 6: "32", 7: "1", 8: "6", 9: "B",
                  10: "2", 11: "11", 12: "class", 13: "14", 14: "3", 15: "2", 16: "5", 17: "",
                  18: "S=0 / n = int(input()) / if n < 0 or n%1 != 0: ... (не дописано)",
                  19: "7", 20: "10", 21: "3", 22: "4", 23: "", 24: "", 25: "3", 26: "", 27: ""},
    "Павлов Тимофей, Галимзянов Андрей": {1: "false", 2: "for i in range(5)", 3: "3ab (ababab)",
                                          5: "", 6: "", 7: "", 8: "6", 9: "B", 10: "", 11: "11",
                                          12: "", 13: "14", 14: "", 15: "", 16: ""},
    "Корчевский, Каленик": {1: "False", 2: "For in rage(5)", 3: "3ab", 4: "For in rage(14)",
                            5: "0,14", 6: "2*5*5=50", 7: "201", 8: "Hello", 9: "False", 10: "False",
                            11: "X=2", 12: "My_var", 13: "14",
                            14: "Range(1, 11) должен быть range(1, 10)",
                            15: "Зависит от того, что ввел пользователь", 16: "", 17: "True",
                            18: "Make_12_with_=4512", 19: "6", 20: "1+2+3+4", 21: "3",
                            22: "Ошибки нет, выведет 10 20 30", 23: "D=2,5 A=2", 24: "2,5",
                            25: "[1, 2, 3]", 26: "DOWLAND add"},
    "Без фамилии (файл «Тест (2)»)": {1: "3", 2: "4", 3: "3ab", 4: "", 5: "", 6: "10", 7: "",
                                      8: "Привет", 9: "А", 10: "4", 11: "", 12: "3", 13: "14",
                                      14: "2", 15: "2", 16: "", 17: "2", 18: "", 19: "", 20: "",
                                      21: "17", 22: "3", 23: ""},
    "Без фамилии (файл «тест»)": {1: "3", 2: "1", 3: "3ab", 4: "", 6: "10", 11: "9", 12: "2",
                                  13: "14", 14: "4", 15: "4"},
}

# эти трое отвечали в своём порядке и словами — задание определено по смыслу ответа
INDIVIDUAL = {
    "Худынцев (14 шк., 11А)": {
        "e9": "for i in range(5):", "e7": "ab в степени 3", "e6": "2 в степени 5",
        "e15": "Нет отступа у строки внутри цикла", "e4": "11", "e8": "6", "e5": "3   1",
        "e12": "5", "e2": "Class", "e1": "Int", "e10": "147", "e13": "B", "e11": "False",
        "e14": "Nety else", "m9": "Len нельзя использовать внутри range", "m4": "True",
        "m1": "огр", "m3": "6", "m2": "21", "h2": "29", "h3": "114"},
    "Потехин, Галимзянов": {
        "e11": "False", "e9": "For I in range(5):", "e7": "Ababab", "e5": "3    5", "e6": "32",
        "e8": "6", "e13": "B", "e14": "Int(input)", "e4": "11", "e2": "My_var", "e3": "14",
        "e15": "Net Tstypa v cikle", "e1": "Str", "m7": "3", "m2": "21"},
    "Умар, Шарикалов Андрей": {
        "e11": "false", "e9": "for i in range (5)", "e7": "ababab", "e5": "3.5", "e6": "32",
        "e8": "6", "e14": "int", "e2": "class", "e3": "14", "e1": "зависит от п...",
        "m10": "a = int(input(введите любое число)); a = b; print(b + b = b)",
        "e4": "11", "m2": "21", "m5": "4 1 2 3"},
}

# вердикты, которые нельзя получить сравнением строк (ответ по смыслу, опечатки, код)
MANUAL_VERDICT = {
    ("Павлов Тимофей, Галимзянов Андрей", "e7"): OK,    # написал «ababab» в скобках
    ("Худынцев (14 шк., 11А)", "e7"): BAD,              # «ab в степени 3» — не вывод программы
    ("Худынцев (14 шк., 11А)", "e6"): BAD,              # «2 в степени 5» — не посчитал
    ("Худынцев (14 шк., 11А)", "e15"): OK,
    ("Худынцев (14 шк., 11А)", "e14"): BAD,
    ("Худынцев (14 шк., 11А)", "m9"): BAD,
    ("Потехин, Галимзянов", "e14"): OK,                 # «Int(input)» — суть верна
    ("Потехин, Галимзянов", "e15"): OK,                 # «Net Tstypa v cikle» — нет отступа
    ("Умар, Шарикалов Андрей", "e14"): OK,              # «int» — нужен int(input())
    ("Умар, Шарикалов Андрей", "m10"): BAD,             # код нерабочий
    ("Чистякова", "m10"): BAD,                          # код не дописан
    ("Корчевский, Каленик", "e9"): OK,                  # «For in rage(5)» — опечатка, ответ верный
    ("Корчевский, Каленик", "m10"): BAD,
    ("Корчевский, Каленик", "e15"): BAD,
    ("Корчевский, Каленик", "e1"): BAD,
    ("Корчевский, Каленик", "e14"): BAD,
    ("Корчевский, Каленик", "m9"): BAD,
    ("Корчевский, Каленик", "m5"): BAD,
}


def norm(s):
    s = str(s).strip().lower().replace("ё", "е").replace(",", " ").rstrip(":")
    return " ".join(s.split())


def loose(s):
    """без пробелов и двоеточий — чтобы «147» и «for i in range(5)» засчитывались"""
    return norm(s).replace(" ", "").replace(":", "")


def auto_verdict(qid, given):
    """✓ / ✗ / — / ? для ответа, записанного словами."""
    q = BY_ID[qid]
    g = given.strip()
    if not g:
        return NONE
    if q["type"] == "code":
        return UNK
    if g in "1234" and len(g) == 1 and q["type"] == "choice":
        return UNK                      # записан только номер варианта
    answers = [q["answer"]] if q["type"] == "choice" else q["answer"]
    if loose(g) in {loose(a) for a in answers}:
        return OK                       # «147» вместо «1 4 7», «range(5)» без двоеточия
    return BAD


def grade_paper():
    out = {}
    for name, by_pos in CLUSTER.items():
        answers = {}
        for pos, given in by_pos.items():
            if 1 <= pos <= len(CLUSTER_ORDER):
                answers[CLUSTER_ORDER[pos - 1]] = given
        out[name] = answers
    out.update(INDIVIDUAL)
    graded = {}
    for name, answers in out.items():
        res = {}
        for q in QUESTIONS:
            given = answers.get(q["id"], "")
            v = MANUAL_VERDICT.get((name, q["id"])) or auto_verdict(q["id"], given)
            res[q["id"]] = (given, v)
        graded[name] = res
    return graded


def grade_dashboard():
    rows = json.loads((BASE / "results" / "results.json").read_text(encoding="utf-8"))
    graded = {}
    for r in rows:
        ids = {x["id"] for x in r["detail"]}
        if not ids & set(BY_ID):
            continue                    # работа по старому набору заданий
        if not any(x["given"].strip() for x in r["detail"]):
            continue                    # пустая попытка
        res = {}
        for q in QUESTIONS:
            x = next((y for y in r["detail"] if y["id"] == q["id"]), None)
            given = x["given"] if x else ""
            ok = x["ok"] if x else False
            if ok is None:
                ok = r.get("marks", {}).get(q["id"])
            v = NONE if not given.strip() else (UNK if ok is None else (OK if ok else BAD))
            res[q["id"]] = (given, v)
        graded[r["name"]] = res
    return graded, {r["name"]: r for r in rows}


def score(res):
    pts = mx = unk = 0
    for qid, (_given, v) in res.items():
        p = POINTS[BY_ID[qid]["section"]]
        if v == UNK:
            unk += 1
            continue
        mx += p
        pts += p if v == OK else 0
    return pts, mx, unk


if __name__ == "__main__":
    paper = grade_paper()
    dash, meta = grade_dashboard()
    for name, res in list(paper.items()) + list(dash.items()):
        s, m, u = score(res)
        print(f"{name:42} {s:3} / {m:3}   не проверено: {u}")
