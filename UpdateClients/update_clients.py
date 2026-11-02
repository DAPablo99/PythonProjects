import argparse
import calendar
import logging
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

INPUT_DIR = Path("data/input")
PATTERN = "*.xlsx"

OUTPUT = Path("data/output/Active_clients2026.xlsx")
PIVOT_OUTPUT = Path("data/output/Active_clients2026(pivot).xlsx")
BASE_OUTPUT = Path("data/output/Active_clients2026(base).xlsx")
PIVOT_SHEET = "Зведена"

MERGED_NAME = OUTPUT.name
PIVOT_NAME = PIVOT_OUTPUT.name
BASE_NAME = BASE_OUTPUT.name

GOOGLE_CREDENTIALS = "credentials.json"
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit"
GOOGLE_WORKSHEET = "План(поточний місяць)"
GOOGLE_BASE_WORKSHEET = "Поточна динаміка"
GOOGLE_MANAGER_WORKSHEET = "Виконання(поточний місяць)"
MANAGER_COL = "Менеджер продажу"
SUPPORT_COL = "Менеджер підтримки"
TOTAL_MANAGER = "Прізвище менеджера"

STATUS_DAYS = 90

MONTHS = {1: "Січень", 2: "Лютий", 3: "Березень", 4: "Квітень", 5: "Травень",
          6: "Червень", 7: "Липень", 8: "Серпень", 9: "Вересень",
          10: "Жовтень", 11: "Листопад", 12: "Грудень"}

INDEX_COLS = ["Код агента", "Платник контрагент", "Головний контрагент", "Організація",
              "Актуальний статус клієнта", "АВС клас", "Бізнес-сегмент",
              "ЕДРПОУ", "Менеджер продажу", "Менеджер підтримки"]
SUM_COLS = ["Вартість послуги з ПДВ, грн", "Вага відправлення, кг", "Кількість місць"]
INV = "Кількість накладних"
VALUE_COLS = SUM_COLS + [INV]

CLIENT_ATTRS = ["АВС клас", "Головний контрагент", "Бізнес-сегмент",
                "Менеджер продажу", "Менеджер підтримки"]

PLAN_AMOUNT = 10_000_000
PLAN_MONTH = None
PLAN_BASE_MONTHS = 3
PLAN_EXCLUDE_STATUSES = {"новий"}

REVENUE = "Вартість послуги з ПДВ, грн"
STATUS_COL = "Актуальний статус клієнта"
PLAN_BASE_COL = "Факт за базу, грн"
PLAN_SHARE_COL = "Питома вага"
PLAN_COL = "План, грн"
FORECAST_COL = "Прогноз, грн"
FC_VALUE_COLS = [REVENUE, "Вага відправлення, кг", INV]
FC_PCT_COL = "Прогноз до попер. місяця, %"

log = logging.getLogger("update_clients")


def load_and_merge(input_dir: Path) -> pd.DataFrame:
    files = [f for f in sorted(input_dir.glob(PATTERN))
             if f.name not in (MERGED_NAME, PIVOT_NAME, BASE_NAME) and not f.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"У папці {input_dir} немає файлів за маскою {PATTERN}")

    frames = []
    for f in files:
        df = pd.read_excel(
            f, skiprows=4, header=0,
            dtype={"Платник контрагент.Код": str, "Платник контрагент.ЕДРПОУ": str},
        )
        df = df.iloc[:-1]
        df = df.dropna(axis=1, how="all")
        df.columns = (df.columns.astype(str)
                      .str.replace(r"\s+", " ", regex=True)
                      .str.strip())
        df["Файл-джерело"] = f.name
        frames.append(df)
        log.info("%s: %s рядків", f.name, len(df))

    result = pd.concat(frames, ignore_index=True)
    before = len(result)
    result = result.drop_duplicates(subset=["Накладная"], keep="last")
    log.info("Видалено дублів по накладних: %s", before - len(result))

    result["Дата відправлення"] = pd.to_datetime(
        result["Дата відправлення"], format="%d.%m.%Y", errors="coerce"
    )
    bad_dates = int(result["Дата відправлення"].isna().sum())
    if bad_dates:
        log.warning("Не розпізнано дат: %s", bad_dates)

    result.columns = result.columns.str.replace("Платник контрагент.", "", regex=False)
    return result.rename(columns={"Код": "Код агента"})


def add_current_status(result: pd.DataFrame) -> pd.DataFrame:
    cutoff = result["Дата відправлення"].max() - pd.Timedelta(days=STATUS_DAYS)
    last_period = result[result["Дата відправлення"] > cutoff]

    ref = (last_period
           .groupby(["Платник контрагент", "Код агента", "Статус клієнта"], as_index=False)
           ["Дата відправлення"].max())
    ref = (ref.sort_values("Дата відправлення")
              .drop_duplicates(subset="Код агента", keep="last"))
    status_map = ref.set_index("Код агента")["Статус клієнта"]

    result["Актуальний статус клієнта"] = (
        result["Код агента"].map(status_map).fillna(result["Статус клієнта"])
    )
    return result


def normalize_client_attrs(result: pd.DataFrame) -> pd.DataFrame:
    latest = (result.sort_values("Дата відправлення", kind="stable")
                    .drop_duplicates(subset="Код агента", keep="last")
                    .set_index("Код агента")[CLIENT_ATTRS])
    for a in CLIENT_ATTRS:
        result[a] = result["Код агента"].map(latest[a])
    return result


def build_pivot(result: pd.DataFrame):
    src = result.copy()
    src["Місяць"] = src["Дата відправлення"].dt.month
    src[INV] = src["Накладная"]
    src[INDEX_COLS] = src[INDEX_COLS].fillna("")

    aggfunc = {c: "sum" for c in SUM_COLS}
    aggfunc[INV] = "nunique"

    pt = src.pivot_table(index=INDEX_COLS, columns="Місяць", values=VALUE_COLS,
                         aggfunc=aggfunc, fill_value=0)

    months_present = sorted(int(m) for m in src["Місяць"].dropna().unique())
    ordered = [(v, m) for m in months_present for v in VALUE_COLS]
    pt = pt[ordered]
    pt.columns = pd.MultiIndex.from_tuples([(MONTHS[m], v) for v, m in ordered])

    log.info("Рядків у зведеній: %s | унікальних кодів: %s",
             len(pt), src["Код агента"].nunique())
    return pt, months_present


def weighted_days(year: int, month: int, last_day: int) -> float:
    return sum(1.0 if date(year, month, d).weekday() < 5 else 0.25
               for d in range(1, last_day + 1))


def forecast_factor(month: int) -> float:
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    year = today.year
    total_w = weighted_days(year, month, calendar.monthrange(year, month)[1])

    if month == yesterday.month:
        done_w = weighted_days(year, month, yesterday.day)
        log.info("Прогноз %s: пройшло %.2f з %.2f зважених днів (коеф. %.4f)",
                 MONTHS[month], done_w, total_w, total_w / done_w)
        return total_w / done_w
    if month == today.month:
        log.warning("Перший день місяця - прогноз ще не можна порахувати")
        return float("nan")
    return 1.0


def calc_forecast(pt: pd.DataFrame, plan_month: int) -> pd.Series:
    col = (MONTHS[plan_month], REVENUE)

    if col not in pt.columns:
        log.warning("Немає даних за %s - прогноз дорівнює 0", MONTHS[plan_month])
        return pd.Series(0.0, index=pt.index)

    factor = forecast_factor(plan_month)
    if pd.isna(factor):
        return pd.Series(0.0, index=pt.index)
    return pt[col] * factor


def add_plan(pt: pd.DataFrame, plan_amount: float, plan_month: int) -> pd.DataFrame:
    base_months = [(plan_month - k - 1) % 12 + 1 for k in range(PLAN_BASE_MONTHS, 0, -1)]
    names = [MONTHS[m] for m in base_months]
    present = [n for n in names if (n, REVENUE) in pt.columns]
    missing = [n for n in names if n not in present]
    if missing:
        log.warning("Для бази плану немає даних за: %s", ", ".join(missing))
    if not present:
        raise ValueError(f"Немає жодного місяця бази для плану на {MONTHS[plan_month]}")

    fact = pt[[(n, REVENUE) for n in present]].sum(axis=1)

    status = pt.index.get_level_values(STATUS_COL).astype(str).str.strip().str.lower()
    eligible = ~status.isin(PLAN_EXCLUDE_STATUSES)
    base = fact.where(eligible, 0).clip(lower=0)

    total = base.sum()
    if total <= 0:
        raise ValueError("Сума факту за базу дорівнює нулю - план рознести не можна")
    share = base / total

    top = f"План {MONTHS[plan_month]} (база: {present[0]}–{present[-1]})"
    pt = pt.copy()
    pt[(top, PLAN_BASE_COL)] = fact
    pt[(top, PLAN_SHARE_COL)] = share
    pt[(top, PLAN_COL)] = share * plan_amount
    pt[(top, FORECAST_COL)] = calc_forecast(pt, plan_month)

    log.info("План %s грн на %s рознесено на %s рядків (без плану: %s)",
             f"{plan_amount:,.0f}", MONTHS[plan_month], int((base > 0).sum()),
             int((~eligible).sum()))
    return pt


def add_weighted_forecast(pt: pd.DataFrame, month: int) -> pd.DataFrame:
    factor = forecast_factor(month)

    top = f"Прогноз {MONTHS[month]}"
    pt = pt.copy()
    for v in FC_VALUE_COLS:
        col = (MONTHS[month], v)
        fact = pt[col] if col in pt.columns else 0.0
        pt[(top, v)] = fact * factor

    prev_month = (month - 2) % 12 + 1
    prev_col = (MONTHS[prev_month], REVENUE)
    if prev_col in pt.columns:
        prev_fact = pt[prev_col]
        pt[(top, FC_PCT_COL)] = pt[(top, REVENUE)] / prev_fact.where(prev_fact > 0) - 1
    else:
        log.warning("Немає факту за %s - % до попереднього місяця порожній", MONTHS[prev_month])
        pt[(top, FC_PCT_COL)] = float("nan")
    return pt


def column_blocks(labels: list, start: int) -> list:
    blocks, prev = [], None
    for i in range(start, len(labels)):
        lab = labels[i]
        if lab not in ("", None) and lab != prev:
            blocks.append([i, i])
            prev = lab
        elif blocks:
            blocks[-1][1] = i
    return blocks


def save_pivot_excel(pt: pd.DataFrame, months_present: list, path: Path) -> None:
    n_idx = len(INDEX_COLS)
    path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pt.to_excel(writer, sheet_name=PIVOT_SHEET, merge_cells=False)
        ws = writer.sheets[PIVOT_SHEET]

        header_fill = PatternFill("solid", fgColor="D9E1F2")
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        tops = [""] * n_idx + [c[0] for c in pt.columns]
        for c0, c1 in column_blocks(tops, n_idx):
            if c1 > c0:
                ws.merge_cells(start_row=1, start_column=c0 + 1, end_row=1, end_column=c1 + 1)

        for row in ws.iter_rows(min_row=1, max_row=3):
            for cell in row:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = center
        ws.row_dimensions[2].height = 45

        for row in ws.iter_rows(min_row=4, min_col=n_idx + 1):
            for cell in row:
                header = str(ws.cell(2, cell.column).value)
                if header == PLAN_SHARE_COL:
                    cell.number_format = "0.00%"
                elif header == FC_PCT_COL:
                    cell.number_format = "0.0%"
                elif "Вага" in header:
                    cell.number_format = "#,##0.0"
                else:
                    cell.number_format = "#,##0"

        for col in range(1, ws.max_column + 1):
            width = max(10, len(str(ws.cell(3, col).value or "")) + 2)
            for r in range(3, ws.max_row + 1):
                v = ws.cell(r, col).value
                if v is None:
                    continue
                if isinstance(v, float):
                    s = f"{v:,.1f}"
                elif isinstance(v, int):
                    s = f"{v:,}"
                else:
                    s = str(v)
                width = max(width, len(s) + 2)
            ws.column_dimensions[get_column_letter(col)].width = min(width, 45)

        thick = Side(style="medium", color="7F7F7F")
        dotted = Side(style="dotted", color="D9D9D9")
        thin = Side(style="thin", color="E6E6E6")
        month_head = ["D9E1F2", "BDD7EE"]
        month_body = [None, "F5F9FD"]

        def fill(hex_):
            return PatternFill("solid", fgColor=hex_)

        month_i = 0
        for c0, c1 in column_blocks(tops, n_idx):
            first, last = c0 + 1, c1 + 1
            is_plan = str(tops[c0]).startswith(("План", "Прогноз"))
            if is_plan:
                head, body = "FFE699", "FFF8E1"
            else:
                head, body = month_head[month_i % 2], month_body[month_i % 2]
                month_i += 1

            for r in range(1, ws.max_row + 1):
                for c in range(first, last + 1):
                    cell = ws.cell(r, c)
                    cell.border = Border(
                        left=thick if c == first else dotted,
                        right=thick if c == last else dotted,
                        top=None,
                        bottom=thick if r == 3 else thin)
                    if r <= 3:
                        cell.fill = fill(head)
                    elif body:
                        cell.fill = fill(body)
                    if is_plan and r >= 4 and ws.cell(2, c).value in (PLAN_COL, FORECAST_COL):
                        cell.font = Font(bold=True)

        for r in range(1, ws.max_row + 1):
            for c in range(1, n_idx + 1):
                ws.cell(r, c).border = Border(
                    right=thick if c == n_idx else None,
                    bottom=thick if r == 3 else None)

        ws.freeze_panes = ws.cell(row=4, column=3)
        ws.auto_filter.ref = f"A3:{get_column_letter(ws.max_column)}{ws.max_row}"


def _clean(v):
    return v.item() if hasattr(v, "item") else v


def _gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_file(str(GOOGLE_CREDENTIALS), scopes=scopes)
    return gspread.authorize(credentials)


def read_pivot_for_google(path: Path):
    n_idx = len(INDEX_COLS)
    raw = pd.read_excel(path, sheet_name=PIVOT_SHEET, header=None).fillna("")

    months_row = raw.iloc[0].tolist()
    header_row = raw.iloc[2, :n_idx].tolist() + raw.iloc[1, n_idx:].tolist()
    body = raw.iloc[3:].values.tolist()

    updated = f"Оновлено: {datetime.now():%d.%m.%Y %H:%M}"
    data = [[updated] + [""] * (raw.shape[1] - 1), months_row, header_row] + body
    data = [[_clean(v) for v in row] for row in data]
    return data, updated


def _format_requests(sid: int, data: list, n_idx: int, n_val: int) -> list:
    n_rows, n_cols = len(data), len(data[0])

    def rng(r0, r1, c0, c1):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def rgb(h):
        return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255,
                "blue": int(h[4:6], 16) / 255}

    def px(text_len):
        return int(min(max(text_len * 7 + 30, 70), 330))

    reqs = [
        {"repeatCell": {"range": rng(0, n_rows, 0, n_cols),
                        "cell": {}, "fields": "userEnteredFormat"}},
        {"repeatCell": {
            "range": rng(0, 1, 0, 1),
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "wrapStrategy": "OVERFLOW_CELL"}},
            "fields": "userEnteredFormat(textFormat,wrapStrategy)"}},
        {"repeatCell": {
            "range": rng(1, 3, 0, n_cols),
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb("D9E1F2"),
                "textFormat": {"bold": True},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,"
                      "horizontalAlignment,verticalAlignment,wrapStrategy)"}},
    ]

    headers = data[2]
    for c in range(n_idx, n_cols):
        if headers[c] == PLAN_SHARE_COL:
            pattern = "0.00%"
        elif headers[c] == FC_PCT_COL:
            pattern = "0.0%"
        elif "Вага" in str(headers[c]):
            pattern = "#,##0.0"
        else:
            pattern = "#,##0"
        reqs.append({"repeatCell": {
            "range": rng(3, n_rows, c, c + 1),
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": pattern}}},
            "fields": "userEnteredFormat.numberFormat"}})

    for c in range(n_idx):
        longest = max(len(str(row[c])) for row in data[2:])
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c, "endIndex": c + 1},
            "properties": {"pixelSize": px(longest)}, "fields": "pixelSize"}})

    reqs += [
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": n_idx, "endIndex": n_cols},
            "properties": {"pixelSize": 115}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 50}, "fields": "pixelSize"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid,
                           "gridProperties": {"frozenRowCount": 3, "frozenColumnCount": 2}},
            "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}},
    ]

    def border(style, color):
        return {"style": style, "color": color}

    DARK = rgb("7F7F7F")
    LIGHT = rgb("D9D9D9")

    MONTH_HEAD = ["D9E1F2", "BDD7EE"]
    MONTH_BODY = [None, "F5F9FD"]
    PLAN_HEAD, PLAN_BODY = "FFE699", "FFF8E1"

    reqs.append({"updateBorders": {
        "range": rng(3, n_rows, 0, n_cols),
        "innerHorizontal": border("SOLID", rgb("E6E6E6"))}})

    reqs.append({"updateBorders": {
        "range": rng(2, 3, 0, n_cols),
        "bottom": border("SOLID_MEDIUM", DARK)}})

    reqs.append({"updateBorders": {
        "range": rng(1, n_rows, 0, n_idx),
        "right": border("SOLID_MEDIUM", DARK)}})

    month_i = 0
    for c0, c1 in column_blocks(data[1], n_idx):
        is_plan = str(data[1][c0]).startswith(("План", "Прогноз"))
        if is_plan:
            head, body = PLAN_HEAD, PLAN_BODY
        else:
            head, body = MONTH_HEAD[month_i % 2], MONTH_BODY[month_i % 2]
            month_i += 1

        reqs.append({"repeatCell": {
            "range": rng(1, 3, c0, c1 + 1),
            "cell": {"userEnteredFormat": {"backgroundColor": rgb(head)}},
            "fields": "userEnteredFormat.backgroundColor"}})
        if body:
            reqs.append({"repeatCell": {
                "range": rng(3, n_rows, c0, c1 + 1),
                "cell": {"userEnteredFormat": {"backgroundColor": rgb(body)}},
                "fields": "userEnteredFormat.backgroundColor"}})
        reqs.append({"updateBorders": {
            "range": rng(1, n_rows, c0, c1 + 1),
            "left": border("SOLID_MEDIUM", DARK),
            "right": border("SOLID_MEDIUM", DARK),
            "innerVertical": border("DOTTED", LIGHT)}})
        if is_plan:
            for c in range(c0, c1 + 1):
                if headers[c] in (PLAN_COL, FORECAST_COL):
                    reqs.append({"repeatCell": {
                        "range": rng(3, n_rows, c, c + 1),
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                        "fields": "userEnteredFormat.textFormat.bold"}})

    reqs.append({"updateBorders": {
        "range": rng(n_rows - 1, n_rows, 0, n_cols),
        "bottom": border("SOLID_MEDIUM", DARK)}})
    return reqs


def _gradient_requests(spreadsheet, sid: int, data: list) -> list:
    headers = data[2]
    if FC_PCT_COL not in headers:
        return []
    c = headers.index(FC_PCT_COL)

    def rgb(h):
        return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255,
                "blue": int(h[4:6], 16) / 255}

    meta = spreadsheet.fetch_sheet_metadata()
    n_old = next((len(s.get("conditionalFormats", [])) for s in meta["sheets"]
                  if s["properties"]["sheetId"] == sid), 0)
    reqs = [{"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}}
            for _ in range(n_old)]

    GRADIENT_LIMIT = "0,3"
    reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": 3, "endRowIndex": len(data),
                    "startColumnIndex": c, "endColumnIndex": c + 1}],
        "gradientRule": {
            "minpoint": {"color": rgb("F8696B"), "type": "NUMBER", "value": f"-{GRADIENT_LIMIT}"},
            "midpoint": {"color": rgb("FFFFFF"), "type": "NUMBER", "value": "0"},
            "maxpoint": {"color": rgb("63BE7B"), "type": "NUMBER", "value": GRADIENT_LIMIT}}}}})
    return reqs


def upload_to_google(data: list, months_present: list, updated: str,
                     worksheet_name: str = GOOGLE_WORKSHEET) -> None:
    n_idx, n_val = len(INDEX_COLS), len(VALUE_COLS)
    n_rows, n_cols = len(data), len(data[0])

    worksheet = _gspread_client().open_by_url(GOOGLE_SHEET_URL).worksheet(worksheet_name)

    if worksheet.row_count < n_rows:
        worksheet.add_rows(n_rows - worksheet.row_count)
    if worksheet.col_count < n_cols:
        worksheet.add_cols(n_cols - worksheet.col_count)

    try:
        worksheet.clear_basic_filter()
    except Exception:
        pass

    worksheet.clear()
    worksheet.unmerge_cells(f"A2:{rowcol_to_a1(2, worksheet.col_count)}")
    worksheet.update(range_name="A1", values=data)

    for c0, c1 in column_blocks(data[1], n_idx):
        if c1 > c0:
            worksheet.merge_cells(f"{rowcol_to_a1(2, c0 + 1)}:{rowcol_to_a1(2, c1 + 1)}")

    worksheet.set_basic_filter(f"A3:{rowcol_to_a1(n_rows, n_cols)}")

    reqs = _format_requests(worksheet.id, data, n_idx, n_val)
    reqs += _gradient_requests(worksheet.spreadsheet, worksheet.id, data)
    worksheet.spreadsheet.batch_update({"requests": reqs})

    log.info("Аркуш «%s» оновлено! %s", worksheet_name, updated)


def build_manager_summary(pt: pd.DataFrame, plan_month: int,
                          manager_col: str = MANAGER_COL) -> pd.DataFrame:
    plan_top = next((c[0] for c in pt.columns if c[1] == PLAN_COL), None)
    if plan_top is None:
        raise ValueError("У зведеній немає блоку плану - зведену по менеджерах не побудувати")

    fact_col = (MONTHS[plan_month], REVENUE)
    fact = pt[fact_col] if fact_col in pt.columns else 0

    df = pd.DataFrame({
        "Факт, грн": fact,
        "План, грн": pt[(plan_top, PLAN_COL)],
        "Прогноз, грн": pt[(plan_top, FORECAST_COL)],
    })
    df = df.groupby(level=manager_col).sum()
    df.index = df.index.where(df.index != "", "(без менеджера)")
    df = df.sort_values("План, грн", ascending=False)
    df.loc["Разом"] = df.sum()

    plan = df["План, грн"]
    df["Прогноз до плану, %"] = (df["Прогноз, грн"] / plan).where(plan > 0, 0)

    df.index.name = manager_col
    return df.reset_index()


def upload_manager_summary(df: pd.DataFrame, plan_month: int,
                           col_off: int = 0, first: bool = True) -> None:
    spreadsheet = _gspread_client().open_by_url(GOOGLE_SHEET_URL)

    n_cols = len(df.columns)
    n_rows = len(df) + 2
    need_cols = col_off + n_cols
    try:
        ws = spreadsheet.worksheet(GOOGLE_MANAGER_WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(GOOGLE_MANAGER_WORKSHEET, rows=n_rows + 5, cols=need_cols + 1)

    updated = f"Оновлено: {datetime.now():%d.%m.%Y %H:%M} | Факт за {MONTHS[plan_month]}"
    values = ([[updated] + [""] * (n_cols - 1), df.columns.tolist()]
              + [[_clean(v) for v in row] for row in df.values.tolist()])

    if first:
        ws.clear()
    if ws.row_count < n_rows:
        ws.add_rows(n_rows - ws.row_count)
    if ws.col_count < need_cols:
        ws.add_cols(need_cols - ws.col_count)
    ws.update(range_name=rowcol_to_a1(1, col_off + 1), values=values)

    sid = ws.id

    def rng(r0, r1, c0, c1):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0 + col_off, "endColumnIndex": c1 + col_off}

    reqs = [
        {"repeatCell": {"range": rng(0, n_rows, 0, n_cols),
                        "cell": {}, "fields": "userEnteredFormat"}},
        {"repeatCell": {"range": rng(0, 1, 0, 1),
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                                       "wrapStrategy": "OVERFLOW_CELL"}},
                        "fields": "userEnteredFormat(textFormat,wrapStrategy)"}},
        {"repeatCell": {"range": rng(1, 2, 0, n_cols),
                        "cell": {"userEnteredFormat": {
                            "backgroundColor": {"red": 0.851, "green": 0.882, "blue": 0.949},
                            "textFormat": {"bold": True},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP"}},
                        "fields": "userEnteredFormat(backgroundColor,textFormat,"
                                  "horizontalAlignment,verticalAlignment,wrapStrategy)"}},
        {"repeatCell": {"range": rng(n_rows - 1, n_rows, 0, n_cols),
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                        "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {"range": rng(2, n_rows, 1, 4),
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": rng(2, n_rows, 4, n_cols),
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.0%"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": col_off, "endIndex": col_off + 1},
            "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": col_off + 1, "endIndex": col_off + n_cols},
            "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"}},
    ]

    def rgb(h):
        return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255,
                "blue": int(h[4:6], 16) / 255}

    def border(style, color):
        return {"style": style, "color": color}

    DARK, LIGHT = rgb("7F7F7F"), rgb("D9D9D9")

    groups = [
        (1, 2, "BDD7EE", None),
        (2, 4, "FFE699", "FFF8E1"),
        (4, n_cols, "E2EFDA", "F4FAF0"),
    ]

    reqs.append({"updateBorders": {"range": rng(2, n_rows, 0, n_cols),
                                   "innerHorizontal": border("SOLID", rgb("E6E6E6"))}})
    reqs.append({"updateBorders": {"range": rng(1, 2, 0, n_cols),
                                   "bottom": border("SOLID_MEDIUM", DARK)}})
    reqs.append({"updateBorders": {"range": rng(1, n_rows, 0, 1),
                                   "right": border("SOLID_MEDIUM", DARK)}})

    for c0, c1, head, body in groups:
        reqs.append({"repeatCell": {
            "range": rng(1, 2, c0, c1),
            "cell": {"userEnteredFormat": {"backgroundColor": rgb(head)}},
            "fields": "userEnteredFormat.backgroundColor"}})
        if body:
            reqs.append({"repeatCell": {
                "range": rng(2, n_rows, c0, c1),
                "cell": {"userEnteredFormat": {"backgroundColor": rgb(body)}},
                "fields": "userEnteredFormat.backgroundColor"}})
        reqs.append({"updateBorders": {
            "range": rng(1, n_rows, c0, c1),
            "left": border("SOLID_MEDIUM", DARK),
            "right": border("SOLID_MEDIUM", DARK),
            "innerVertical": border("DOTTED", LIGHT)}})

    reqs.append({"repeatCell": {
        "range": rng(n_rows - 1, n_rows, 0, n_cols),
        "cell": {"userEnteredFormat": {"backgroundColor": rgb("EDEDED"),
                                       "textFormat": {"bold": True}}},
        "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append({"updateBorders": {"range": rng(n_rows - 1, n_rows, 0, n_cols),
                                   "top": border("SOLID_MEDIUM", DARK),
                                   "bottom": border("SOLID_MEDIUM", DARK)}})

    if first:
        meta = spreadsheet.fetch_sheet_metadata()
        n_old = next((len(s.get("conditionalFormats", [])) for s in meta["sheets"]
                      if s["properties"]["sheetId"] == sid), 0)
        reqs = [{"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}}
                for _ in range(n_old)] + reqs

    THRESHOLD = 90
    first_cell = rowcol_to_a1(3, 5 + col_off)

    for formula, bg, fg in [
        (f"={first_cell}*100>={THRESHOLD}", "C6EFCE", "006100"),
        (f"={first_cell}*100<{THRESHOLD}", "FFC7CE", "9C0006"),
    ]:
        reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [rng(2, n_rows, 4, n_cols)],
            "booleanRule": {
                "condition": {"type": "CUSTOM_FORMULA",
                              "values": [{"userEnteredValue": formula}]},
                "format": {"backgroundColor": rgb(bg),
                           "textFormat": {"foregroundColor": rgb(fg)}}}}}})
    spreadsheet.batch_update({"requests": reqs})
    log.info("Аркуш %s оновлено (зсув колонок: %s)", GOOGLE_MANAGER_WORKSHEET, col_off)


def build_manager_total(manager_df: pd.DataFrame, manager_part: str):
    total = manager_df[manager_df[MANAGER_COL] == "Разом"]
    names = manager_df.loc[
        manager_df[MANAGER_COL].astype(str).str.contains(manager_part, case=False, regex=False),
        MANAGER_COL]
    if total.empty or names.empty:
        log.warning("Менеджера '%s' не знайдено - загальний рядок пропущено", manager_part)
        return None
    out = total.copy()
    out[MANAGER_COL] = names.iloc[0]
    return out.reset_index(drop=True)


def upload_manager_total(df: pd.DataFrame, title: str, start_row: int) -> None:
    ws = _gspread_client().open_by_url(GOOGLE_SHEET_URL).worksheet(GOOGLE_MANAGER_WORKSHEET)

    n_cols = len(df.columns)
    head_r, data_r, end = start_row + 1, start_row + 2, start_row + 3
    if ws.row_count < end:
        ws.add_rows(end - ws.row_count)

    values = ([[title] + [""] * (n_cols - 1), df.columns.tolist()]
              + [[_clean(v) for v in row] for row in df.values.tolist()])
    ws.update(range_name=rowcol_to_a1(start_row + 1, 1), values=values)

    sid = ws.id

    def rng(r0, r1, c0, c1):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    def rgb(h):
        return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255,
                "blue": int(h[4:6], 16) / 255}

    def border(style, color):
        return {"style": style, "color": color}

    DARK, LIGHT = rgb("7F7F7F"), rgb("D9D9D9")

    reqs = [
        {"repeatCell": {"range": rng(start_row, end, 0, n_cols),
                        "cell": {}, "fields": "userEnteredFormat"}},
        {"repeatCell": {"range": rng(start_row, start_row + 1, 0, 1),
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                                       "wrapStrategy": "OVERFLOW_CELL"}},
                        "fields": "userEnteredFormat(textFormat,wrapStrategy)"}},
        {"repeatCell": {"range": rng(head_r, head_r + 1, 0, n_cols),
                        "cell": {"userEnteredFormat": {
                            "backgroundColor": rgb("D9E1F2"),
                            "textFormat": {"bold": True},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP"}},
                        "fields": "userEnteredFormat(backgroundColor,textFormat,"
                                  "horizontalAlignment,verticalAlignment,wrapStrategy)"}},
        {"repeatCell": {"range": rng(data_r, end, 1, 4),
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": rng(data_r, end, 4, n_cols),
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.0%"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"updateBorders": {"range": rng(head_r, head_r + 1, 0, n_cols),
                           "bottom": border("SOLID_MEDIUM", DARK)}},
        {"updateBorders": {"range": rng(head_r, end, 0, 1),
                           "right": border("SOLID_MEDIUM", DARK)}},
    ]

    for c0, c1, head in [(1, 2, "BDD7EE"), (2, 4, "FFE699"), (4, n_cols, "E2EFDA")]:
        reqs.append({"repeatCell": {
            "range": rng(head_r, head_r + 1, c0, c1),
            "cell": {"userEnteredFormat": {"backgroundColor": rgb(head)}},
            "fields": "userEnteredFormat.backgroundColor"}})
        reqs.append({"updateBorders": {
            "range": rng(head_r, end, c0, c1),
            "left": border("SOLID_MEDIUM", DARK),
            "right": border("SOLID_MEDIUM", DARK),
            "innerVertical": border("DOTTED", LIGHT)}})

    reqs.append({"repeatCell": {
        "range": rng(data_r, end, 0, n_cols),
        "cell": {"userEnteredFormat": {"backgroundColor": rgb("D9D9D9"),
                                       "textFormat": {"bold": True}}},
        "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append({"updateBorders": {
        "range": rng(data_r, end, 0, n_cols),
        "top": border("SOLID_MEDIUM", rgb("000000")),
        "bottom": border("SOLID_MEDIUM", rgb("000000"))}})

    THRESHOLD = 90
    first_cell = rowcol_to_a1(data_r + 1, 5)
    for formula, bg, fg in [(f"={first_cell}*100>={THRESHOLD}", "C6EFCE", "006100"),
                            (f"={first_cell}*100<{THRESHOLD}", "FFC7CE", "9C0006")]:
        reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [rng(data_r, end, 4, n_cols)],
            "booleanRule": {
                "condition": {"type": "CUSTOM_FORMULA",
                              "values": [{"userEnteredValue": formula}]},
                "format": {"backgroundColor": rgb(bg),
                           "textFormat": {"foregroundColor": rgb(fg)}}}}}})

    ws.spreadsheet.batch_update({"requests": reqs})
    log.info("Загальний рядок додано на аркуш %s", GOOGLE_MANAGER_WORKSHEET)


def parse_args():
    p = argparse.ArgumentParser(description="Оновлення зведеної по клієнтах")
    p.add_argument("--input-dir", type=Path, default=INPUT_DIR)
    p.add_argument("--no-google", action="store_true")
    p.add_argument("--plan", type=float, default=PLAN_AMOUNT)
    p.add_argument("--plan-month", type=int, choices=range(1, 13), default=PLAN_MONTH)
    return p.parse_args()


def setup_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(output_dir / "update_clients.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def main() -> int:
    args = parse_args()
    setup_logging(OUTPUT.parent)

    try:
        result = load_and_merge(args.input_dir)
        result = add_current_status(result)
        result = normalize_client_attrs(result)

        pt_base, months_present = build_pivot(result)
        fc_month = args.plan_month or datetime.now().month
        pt_fc = add_weighted_forecast(pt_base, fc_month)
        pt_plan = manager_df = support_df = None
        if args.plan:
            plan_month = args.plan_month or datetime.now().month
            pt_plan = add_plan(pt_base, args.plan, plan_month)
            manager_df = build_manager_summary(pt_plan, plan_month)
            support_df = build_manager_summary(pt_plan, plan_month, SUPPORT_COL)

        save_pivot_excel(pt_fc, months_present, BASE_OUTPUT)
        log.info("Збережено: %s", BASE_OUTPUT)
        if pt_plan is not None:
            save_pivot_excel(pt_plan, months_present, PIVOT_OUTPUT)
            log.info("Збережено: %s", PIVOT_OUTPUT)

        if not args.no_google:
            data, updated = read_pivot_for_google(BASE_OUTPUT)
            upload_to_google(data, months_present, updated, GOOGLE_BASE_WORKSHEET)

            if pt_plan is not None:
                data, updated = read_pivot_for_google(PIVOT_OUTPUT)
                upload_to_google(data, months_present, updated, GOOGLE_WORKSHEET)

                upload_manager_summary(manager_df, plan_month)
                upload_manager_summary(support_df, plan_month,
                                       col_off=len(manager_df.columns) + 1, first=False)
                total_df = build_manager_total(manager_df, TOTAL_MANAGER)
                if total_df is not None:
                    upload_manager_total(total_df, "Загальний результат",
                                         start_row=len(manager_df) + 4)
    except Exception:
        log.exception("Оновлення завершилось помилкою")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
