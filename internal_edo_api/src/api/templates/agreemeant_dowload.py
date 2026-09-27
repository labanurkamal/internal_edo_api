"""
Генератор Excel-шаблона: листы 'Объект', 'Договора', 'Инструкция'.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from io import BytesIO

# ── Стили ─────────────────────────────────────────────────────────────────────
YELLOW  = PatternFill("solid", fgColor="FFF2CC")
GREEN   = PatternFill("solid", fgColor="E2EFDA")
BLUE    = PatternFill("solid", fgColor="D6E4F0")
HEADER  = PatternFill("solid", fgColor="2E4057")
GRAY    = PatternFill("solid", fgColor="F2F2F2")
HIDDEN  = PatternFill("solid", fgColor="EFEFEF")

thin   = Side(style="thin", color="CCCCCC")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

hfont  = Font(bold=True, size=10, name="Arial", color="FFFFFF")
bfont  = Font(bold=True, size=11, name="Arial")
nfont  = Font(size=10, name="Arial")
sfont  = Font(size=9,  name="Arial", color="808080", italic=True)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def _cell(ws, coord, value="", fill=None, font=None, align=None, brd=True):
    c = ws[coord]
    c.value = value
    if fill:  c.fill  = fill

    if font:  c.font  = font
    if align: c.alignment = align
    if brd:   c.border = border
    return c


# ── Лист "Объект" ─────────────────────────────────────────────────────────────

OBJECT_ROWS = [
    # (ключ, пример_значения, обязательный)
    ("Название здания (KZ)",              "«Евразия» СО",                                        True),
    ("Название здания (RU)",              "ТЦ «Евразия»",                                        True),
    ("Тип помещения (KZ)",                "тұрғын емес",                                         True),
    ("Тип помещения (RU)",                "нежилое",                                             True),
    ("Адрес здания (KZ)",                 "Алматы қаласы, Медеу ауданы, Самал-1 шағын ауданы, 9 А", True),
    ("Адрес здания (RU)",                 "г. Алматы, микрорайон Самал-1, д.9а",                True),
    ("Собственник (KZ)",                  "«Айсер» ЖШС",                                        True),
    ("Собственник (RU)",                  "ТОО «Айсер»",                                        True),
    ("Госакт номер",                      "0214702",                                             True),
    ("Госакт дата",                       "16 января 2004",                                      True),
    ("Договор аренды номер",              "17",                                                  True),
    ("Договор аренды дата",               "01.01.2026",                                          True),
    ("",                                  "",                                                    False),  # разделитель
    ("Субарендодатель — Тип орг.",        "ИП",                                                  True),
    ("Субарендодатель — Название",        "Сайланбаева",                                         True),
    ("Субарендодатель — ФИО (KZ)",        "Сайланбаева А.Б.",                                    True),
    ("Субарендодатель — ФИО (RU)",        "Сайланбаевой А.Б.,",                                  True),
    ("Субарендодатель — ИИН",             "990404400077",                                        True),
    ("Субарендодатель — Талон",           "KZ43TWQ03888390",                                     True),
    ("Субарендодатель — Уд. личн.",       "048161731",                                           True),
    ("Субарендодатель — Дата уд.",        "01.04.2021",                                          True),
    ("Субарендодатель — Адрес (KZ)",      "Астана қ, Жилой Массив Ақ-Бұлақ-2, Переулок Бағлан 5/18", True),
    ("Субарендодатель — Адрес (RU)",      "г. Астана, Жилой Массив Ақ-Бұлақ-2, Переулок Бағлан 5/18", True),
    ("Субарендодатель — БИК",             "CASPKZKA",                                            True),
    ("Субарендодатель — Счёт IBAN",       "KZ33722S000035874795",                                True),
]


def _build_object_sheet(wb):
    ws = wb.create_sheet("Объект")
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 58

    # Заголовок
    ws.merge_cells("A1:B1")
    _cell(ws, "A1", "Настройки объекта и субарендодателя (заполняется один раз)",
          fill=HEADER, font=Font(bold=True, size=13, name="Arial", color="FFFFFF"),
          align=center, brd=False)
    ws.row_dimensions[1].height = 28

    # Подзаголовки колонок
    ws.row_dimensions[2].height = 20
    _cell(ws, "A2", "Поле",    fill=BLUE, font=bfont, align=center)
    _cell(ws, "B2", "Значение", fill=BLUE, font=bfont, align=center)

    for i, (key, example, required) in enumerate(OBJECT_ROWS, start=3):
        ws.row_dimensions[i].height = 18
        if not key:  # разделитель
            ws.merge_cells(f"A{i}:B{i}")
            _cell(ws, f"A{i}", "── Субарендодатель ──",
                  fill=BLUE, font=Font(bold=True, size=10, name="Arial", color="000000"),
                  align=center)
            continue
        fill = YELLOW if required else GREEN
        _cell(ws, f"A{i}", key,     fill=GRAY,  font=nfont, align=left)
        _cell(ws, f"B{i}", example, fill=fill,  font=nfont, align=left)

    # Валидация типа орг субарендодателя
    dv = DataValidation(type="list", formula1='"ИП,ТОО,АО"', showDropDown=False)
    ws.add_data_validation(dv)
    # Строка 17 = "Субарендодатель — Тип орг." (3 + 13 - 1 = 15... считаем)
    for i, (key, _, _) in enumerate(OBJECT_ROWS, start=3):
        if key == "Субарендодатель — Тип орг.":
            dv.sqref = f"B{i}"
            break


# ── Лист "Договора" ───────────────────────────────────────────────────────────

ДОГОВОРА_COLS = [
    # (заголовок, подсказка, пример, обязательный, скрытый)
    ("Номер договора",           "1/118",                    "1/118",                      True,  False),
    ("Дата подписания",          "дд.мм.гггг",               "15.02.2026",                 True,  False),
    ("Дата начала аренды",       "дд.мм.гггг",               "15.02.2026",                 True,  False),
    ("Дата окончания аренды",    "дд.мм.гггг",               "14.02.2027",                 True,  False),
    ("Этаж",                     "цифра",                    "1",                          True,  False),
    ("Тип помещения",            "бутик / офис / қойма",     "бутик",                      True,  False),
    ("Площадь (м²)",             "цифры",                    "55",                         True,  False),
    ("Ставка за м² (тг)",        "цифры без пробелов",       "12000",                      True,  False),
    ("Коммунальные (тг)",        "цифры без пробелов",       "95000",                      True,  False),
    ("Тип орг. арендатора",      "ИП / ТОО / АО / ЖСШ",     "ИП",                         True,  False),
    ("Название арендатора",      "без ИП/ТОО",               "Даурен",                     True,  False),
    ("ФИО арендатора (KZ)",      "Фамилия Имя Отчество",     "Әбенов Дәурен Қайратович",   True,  False),
    ("ИИН арендатора",           "12 цифр",                  "920730401856",               True,  False),
    ("Уд. личн. арендатора",     "номер без пробелов",       "073456219",                  True,  False),
    ("Дата уд. арендатора",      "дд.мм.гггг",               "05.11.2020",                 True,  False),
    ("Талон арендатора",         "KZ10TWQ...",               "KZ10TWQ07812340",            True,  False),
    ("Адрес арендатора (KZ)",    "Назарбаев к. 44 (без үй)", "Назарбаев к. 44",            True,  False),
    ("Адрес арендатора (RU)",    "Назарбаева д.44 (без ул.)", "Назарбаева д.44",           True,  False),
    ("БИК арендатора",           "CASPKZKA / IRTYKZKA",      "CASPKZKA",                   True,  False),
    ("Счёт IBAN арендатора",     "KZ...",                    "KZ07722S000054321098",       True,  False),
    ("Телефон арендатора",       "+ 7 XXX XXX XX XX",        "+ 7 707 456 78 90",          True,  False),
]

COL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _build_contracts_sheet(wb):
    ws = wb.create_sheet("Договора")
    ws.freeze_panes = "A4"

    # Строка 1 — заголовок
    ws.merge_cells(f"A1:{COL_LETTERS[len(ДОГОВОРА_COLS)-1]}1")
    _cell(ws, "A1", "Договора субаренды — заполните начиная с 5-й строки (строка 4 — пример)",
          fill=HEADER, font=Font(bold=True, size=13, name="Arial", color="FFFFFF"),
          align=center, brd=False)
    ws.row_dimensions[1].height = 28

    # Строка 2 — легенда
    ws.merge_cells("A2:U2")
    ws["A2"].value = "🟡 Жёлтые — обязательные поля"
    ws["A2"].fill  = YELLOW
    ws["A2"].font  = Font(size=10, name="Arial", color="806000")
    ws["A2"].alignment = left
    ws.row_dimensions[2].height = 18

    # Строка 3 — заголовки колонок
    ws.row_dimensions[3].height = 52
    for i, (header, hint, _, required, hidden) in enumerate(ДОГОВОРА_COLS):
        col = COL_LETTERS[i]
        fill = HIDDEN if hidden else HEADER
        c = ws[f"{col}3"]
        c.value = f"{header}\n({hint})" if hint else header
        c.fill  = fill
        c.font  = hfont
        c.alignment = center
        c.border = border

    # Строка 4 — пример
    ws.row_dimensions[4].height = 17
    for i, (_, _, example, _, hidden) in enumerate(ДОГОВОРА_COLS):
        col = COL_LETTERS[i]
        c = ws[f"{col}4"]
        c.value = example
        c.fill  = GRAY
        c.font  = sfont
        c.alignment = left
        c.border = border
    ws["A4"].value = "← ПРИМЕР: " + ДОГОВОРА_COLS[0][2]

    # Строки 5-104 — данные
    for row in range(5, 105):
        ws.row_dimensions[row].height = 17
        for i, (_, _, _, required, hidden) in enumerate(ДОГОВОРА_COLS):
            col = COL_LETTERS[i]
            c = ws[f"{col}{row}"]
            c.fill   = HIDDEN if hidden else (YELLOW if required else GREEN)
            c.font   = nfont
            c.alignment = left
            c.border = border

    # Валидация типа орг
    dv = DataValidation(type="list", formula1='"ИП,ТОО,АО,ЖСШ"', showDropDown=False,
                        showErrorMessage=True, errorTitle="Неверный тип",
                        error="Выберите: ИП, ТОО, АО или ЖСШ")
    ws.add_data_validation(dv)
    dv.sqref = "J5:J104"

    # Валидация типа помещения
    dv2 = DataValidation(type="list", formula1='"бутик,офис,қойма,павильон,дүкен"',
                         showDropDown=False)
    ws.add_data_validation(dv2)
    dv2.sqref = "F5:F104"

    # Ширины колонок
    widths = [14,14,14,14, 7,14, 10,12,12, 10,16,28, 16,14,14,18, 22,22,12,26, 18,
              12,18,8,18,18]
    for i, w in enumerate(widths):
        ws.column_dimensions[COL_LETTERS[i]].width = w

    # Скрываем авто-колонки V-Z
    for i in range(21, len(ДОГОВОРА_COLS)):
        ws.column_dimensions[COL_LETTERS[i]].hidden = True


# ── Лист "Инструкция" ─────────────────────────────────────────────────────────

def _build_instruction_sheet(wb):
    ws = wb.create_sheet("Инструкция")
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 56

    rows = [
        ("📋 ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ", "", True),
        ("", "", False),
        ("1. Начните с листа «Объект»", "Заполните данные здания и субарендодателя один раз", False),
        ("2. Перейдите на лист «Договора»", "Каждая строка начиная с 5-й = один договор", False),
        ("", "", False),
        ("ФОРМАТЫ ПОЛЕЙ", "", True),
        ("Даты",              "дд.мм.гггг  →  15.02.2026", False),
        ("Суммы",             "только цифры без пробелов  →  12000", False),
        ("Тип организации",   "выбрать из выпадающего списка: ИП / ТОО / АО / ЖСШ", False),
        ("Тип помещения",     "выбрать из списка: бутик / офис / қойма / павильон", False),
        ("Название орг.",     "только название без типа  →  Даурен  (НЕ «ИП Даурен»)", False),
        ("ФИО",               "полностью: Фамилия Имя Отчество", False),
        ("Адрес KZ",          "улица и дом без «үй»  →  Назарбаев к. 44  (үй добавится авто)", False),
        ("Адрес RU",          "улица и дом без «ул.»  →  Назарбаева д.44  (ул. добавится авто)", False),
        ("БИК",               "CASPKZKA = Kaspi Bank / IRTYKZKA = Халык Банк", False),
        ("", "", False),
        ("ЧТО ВЫЧИСЛЯЕТСЯ АВТОМАТИЧЕСКИ", "", True),
        ("✅ Итого аренда",   "Площадь × Ставка за м²", False),
        ("✅ Прописью (KZ/RU)", "по сумме через num2words", False),
        ("✅ Месяц из даты",  "15.02.2026 → ақпан / февраля", False),
        ("✅ ЖК / ЖШС / АҚ", "по типу орг ИП / ТОО / АО", False),
        ("✅ КБе",            "ИП → 19, ТОО/АО → 17", False),
        ("✅ Банк",           "по БИК", False),
        ("✅ Сокр. ФИО",      "Әбенов Дәурен Қайратович → Әбенов Д.Қ.", False),
        ("✅ үй / ул.",       "добавляется к адресу автоматически", False),
    ]

    ws.row_dimensions[1].height = 10
    for i, (a, b, is_header) in enumerate(rows, start=2):
        ws.row_dimensions[i].height = 20
        ca = ws[f"A{i}"]
        cb = ws[f"B{i}"]
        ca.value = a
        cb.value = b
        if is_header:
            ca.fill = BLUE
            cb.fill = BLUE
            ca.font = Font(bold=True, size=11, name="Arial", color="FFFFFF")
            cb.font = Font(bold=True, size=11, name="Arial", color="FFFFFF")
            ws.merge_cells(f"A{i}:B{i}") if not b else None
        else:
            ca.font = Font(bold=bool(a and not b), size=10, name="Arial")
            cb.font = Font(size=10, name="Arial")
        ca.alignment = left
        cb.alignment = left


# ── Главная функция ───────────────────────────────────────────────────────────

def create_template(out_path=None) -> BytesIO | None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # удаляем дефолтный лист

    _build_object_sheet(wb)
    _build_contracts_sheet(wb)
    # _build_instruction_sheet(wb)

    if out_path is None or isinstance(out_path, BytesIO):
        buf = out_path or BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    wb.save(out_path)
    print(f"✓ Шаблон создан: {out_path}")

#
# if __name__ == "__main__":
#     create_template("Договора_Шаблон.xlsx")