# -*- coding: utf-8 -*-
"""Собирает Excel со всеми результатами: бумажные работы + дашборд."""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))
from questions import QUESTIONS, POINTS  # noqa: E402
from grade_manual import BAD, NONE, OK, UNK, grade_dashboard, grade_paper  # noqa: E402

OUT = BASE / "results" / "Результаты теста.xlsx"
FILES = {
    "Галицких": "Галицких 14 школа.docx",
    "Хакимов, Малиновский": "хакимов малиновский.docx",
    "Дёнина, Сёмкина, Зацепин, Бабушкин": "денина семкина зацепин бабушкин.docx",
    "Чистякова": "Чистякова.docx",
    "Павлов Тимофей, Галимзянов Андрей": "павлов тимофей галимзянов андрей.docx",
    "Корчевский, Каленик": "Корчевский и Каленик.docx",
    "Без фамилии (файл «Тест (2)»)": "Тест (2).docx",
    "Без фамилии (файл «тест»)": "тест.docx",
    "Худынцев (14 шк., 11А)": "Худынцев 14шк 11А.docx",
    "Потехин, Галимзянов": "потехин галимзямов.docx",
    "Умар, Шарикалов Андрей": "UMAR- SHIRIKALOV ANDREI.txt",
}

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F6F54")
LVL_FILL = {"Лёгкие": "E8F2EC", "Средние": "FFF4DC", "Сложные": "FBE8E6"}
OK_FILL = PatternFill("solid", fgColor="D6EFDF")
BAD_FILL = PatternFill("solid", fgColor="FBDBD7")
UNK_FILL = PatternFill("solid", fgColor="FFF0C9")
NONE_FILL = PatternFill("solid", fgColor="F0F0F0")
THIN = Side(style="thin", color="D4DDD8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_BY_MARK = {OK: OK_FILL, BAD: BAD_FILL, UNK: UNK_FILL, NONE: NONE_FILL}


def style_header(ws, row=1, upto=None):
    for c in ws[row][:upto]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        c.fill = HEAD_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 34


def build():
    paper = grade_paper()
    dash, meta = grade_dashboard()
    students = []            # (имя, вердикты, источник, инфо)
    for name, res in paper.items():
        students.append((name, res, "Бумага", FILES.get(name, "")))
    for name, res in dash.items():
        r = meta[name]
        info = f"сдал в {r['finished']}, {r['minutes']} мин., уходов со вкладки: {r['tab_leaves']}"
        students.append((name, res, "Дашборд", info))

    wb = Workbook()

    # ───── лист «По заданиям» (матрица вердиктов) ─────
    ws = wb.create_sheet("По заданиям")
    head = ["№", "Уровень", "Балл", "Задание", "Правильный ответ"] + [s[0] for s in students]
    ws.append(head)
    for i, q in enumerate(QUESTIONS, 1):
        answer = q["answer"] if q["type"] == "choice" else (
            " / ".join(q["answer"]) if q["type"] == "text" else "проверяет учитель")
        text = q["text"] + (("  " + q["code"].replace("\n", " ⏎ ")) if q.get("code") else "")
        row = [i, q["section"], POINTS[q["section"]], text, answer]
        row += [students[j][1][q["id"]][1] for j in range(len(students))]
        ws.append(row)
    style_header(ws)
    for r in range(2, len(QUESTIONS) + 2):
        lvl = ws.cell(r, 2).value
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT, size=10)
            cell.border = BORDER
            cell.fill = PatternFill("solid", fgColor=LVL_FILL[lvl])
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 4))
        for c in range(6, 6 + len(students)):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT, size=11, bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
            cell.fill = FILL_BY_MARK.get(cell.value, NONE_FILL)
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 6
    ws.column_dimensions["D"].width = 62
    ws.column_dimensions["E"].width = 26
    for j in range(len(students)):
        ws.column_dimensions[get_column_letter(6 + j)].width = 6
    ws.freeze_panes = "F2"
    last = len(QUESTIONS) + 1
    ws.cell(last + 2, 4, "Обозначения:").font = Font(name=FONT, bold=True, size=10)
    for k, (mark, txt) in enumerate([
            (OK, "верно"), (BAD, "неверно"), (NONE, "нет ответа"),
            (UNK, "проверить нельзя: ученик записал только номер варианта, "
                  "а порядок вариантов на его экране не сохранился")]):
        ws.cell(last + 3 + k, 3, mark).font = Font(name=FONT, bold=True)
        ws.cell(last + 3 + k, 3).fill = FILL_BY_MARK[mark]
        ws.cell(last + 3 + k, 4, txt).font = Font(name=FONT, size=10)

    # ───── лист «Сводка» ─────
    s = wb["Sheet"]
    s.title = "Сводка"
    s.append(["№", "Ученик", "Класс", "Источник", "Баллы", "Из", "Процент",
              "Ответов дано", "Нельзя проверить", "Файл / как сдал"])
    for i, (name, res, src, info) in enumerate(students, 1):
        col = get_column_letter(6 + i - 1)
        rng = f"'По заданиям'!{col}$2:{col}${len(QUESTIONS) + 1}"
        pts = f"'По заданиям'!$C$2:$C${len(QUESTIONS) + 1}"
        grade = ""
        if src == "Дашборд":
            grade = next(r["grade"] for n, r in meta.items() if n == name)
        s.append([
            i, name, grade, src,
            f'=SUMPRODUCT(({rng}="{OK}")*{pts})',
            f'=SUMPRODUCT(({rng}<>"{UNK}")*{pts})',
            f"=IF(F{i + 1}=0,0,E{i + 1}/F{i + 1})",
            f'=COUNTIF({rng},"{OK}")+COUNTIF({rng},"{BAD}")+COUNTIF({rng},"{UNK}")',
            f'=COUNTIF({rng},"{UNK}")',
            info,
        ])
    style_header(s)
    for r in range(2, len(students) + 2):
        for c in range(1, 11):
            cell = s.cell(r, c)
            cell.font = Font(name=FONT, size=10)
            cell.border = BORDER
            if c in (1, 3, 5, 6, 7, 8, 9):
                cell.alignment = Alignment(horizontal="center")
        s.cell(r, 7).number_format = "0%"
        s.cell(r, 2).font = Font(name=FONT, size=10, bold=True)
    widths = {"A": 4, "B": 36, "C": 7, "D": 11, "E": 8, "F": 7, "G": 9, "H": 13, "I": 16, "J": 46}
    for col, w in widths.items():
        s.column_dimensions[col].width = w
    s.freeze_panes = "A2"
    s.auto_filter.ref = f"A1:J{len(students) + 1}"
    note = len(students) + 3
    s.cell(note, 2, "Как считались баллы").font = Font(name=FONT, bold=True, size=11)
    lines = [
        "Лёгкое задание — 1 балл, среднее — 2, сложное — 3. Всего 50 баллов.",
        "«Из» — максимум по тем заданиям, которые удалось проверить. Если ученик записал только "
        "номер варианта (1–4), задание исключено из баллов: порядок вариантов у каждого свой и не сохранился.",
        "Бумажные работы: номер задания сопоставлен с заданием по тем ученикам, кто писал ответы словами. "
        "У Худынцева, Потехина и Умара свой порядок — их ответы определены по смыслу.",
        "Опечатки засчитывались, если ответ по существу верный (например «for in rage(5)» или «147» вместо «1 4 7»).",
        "Задачи на код в дашборде засчитываются вручную в админке; незачтённые помечены «?».",
    ]
    for k, line in enumerate(lines):
        c = s.cell(note + 1 + k, 2, "• " + line)
        c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        s.merge_cells(start_row=note + 1 + k, start_column=2, end_row=note + 1 + k, end_column=10)
        s.row_dimensions[note + 1 + k].height = 28

    # ───── лист «Ответы учеников» ─────
    a = wb.create_sheet("Ответы учеников")
    a.append(["№", "Уровень", "Задание", "Правильный ответ"] + [st[0] for st in students])
    for i, q in enumerate(QUESTIONS, 1):
        answer = q["answer"] if q["type"] == "choice" else (
            " / ".join(q["answer"]) if q["type"] == "text" else "проверяет учитель")
        row = [i, q["section"], q["text"], answer]
        row += [st[1][q["id"]][0] for st in students]
        a.append(row)
    style_header(a)
    for r in range(2, len(QUESTIONS) + 2):
        for c in range(1, 5 + len(students)):
            cell = a.cell(r, c)
            cell.font = Font(name=FONT, size=10)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for c in range(5, 5 + len(students)):
            mark = students[c - 5][1][QUESTIONS[r - 2]["id"]][1]
            a.cell(r, c).fill = FILL_BY_MARK.get(mark, NONE_FILL)
    a.column_dimensions["A"].width = 4
    a.column_dimensions["B"].width = 10
    a.column_dimensions["C"].width = 46
    a.column_dimensions["D"].width = 22
    for j in range(len(students)):
        a.column_dimensions[get_column_letter(5 + j)].width = 20
    a.freeze_panes = "E2"

    # ───── лист «Статистика по заданиям» ─────
    t = wb.create_sheet("Трудные задания")
    t.append(["№", "Уровень", "Задание", "Правильный ответ", "Решили", "Всего проверено",
              "Доля верных", "Что объяснить"])
    n = len(students)
    for i, q in enumerate(QUESTIONS, 1):
        r = i + 1
        rng = f"'По заданиям'!$F{r}:${get_column_letter(5 + n)}{r}"
        answer = q["answer"] if q["type"] == "choice" else (
            " / ".join(q["answer"]) if q["type"] == "text" else "проверяет учитель")
        t.append([i, q["section"], q["text"], answer,
                  f'=COUNTIF({rng},"{OK}")',
                  f'=COUNTIF({rng},"{OK}")+COUNTIF({rng},"{BAD}")+COUNTIF({rng},"{NONE}")',
                  f"=IF(F{r}=0,0,E{r}/F{r})", q["explain"]])
    style_header(t)
    for r in range(2, len(QUESTIONS) + 2):
        for c in range(1, 9):
            cell = t.cell(r, c)
            cell.font = Font(name=FONT, size=10)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=c in (3, 8))
        t.cell(r, 7).number_format = "0%"
        t.cell(r, 2).fill = PatternFill("solid", fgColor=LVL_FILL[t.cell(r, 2).value])
    for col, w in {"A": 4, "B": 10, "C": 44, "D": 20, "E": 8, "F": 14, "G": 11, "H": 70}.items():
        t.column_dimensions[col].width = w
    t.freeze_panes = "A2"
    t.auto_filter.ref = f"A1:H{len(QUESTIONS) + 1}"

    wb.save(OUT)
    print("готово:", OUT, "| учеников:", len(students))


if __name__ == "__main__":
    build()
