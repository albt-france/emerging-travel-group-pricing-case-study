"""Build excel/ETG_Commercial_Pricing_Cockpit.xlsx with Excel itself (COM automation, pywin32).

Why COM and not only openpyxl: the Power Pivot data model, DAX measures, PivotTables connected
to the model, slicers and CUBE formulas can only be created by Excel.

What the script does:
  1. prepares small aggregated tables with pandas (daily, monthly, price-match log, ...);
  2. opens Excel invisibly, loads 6 CSV files into the data model through Power Query,
     creates the relationships and ~40 DAX measures;
  3. writes every sheet: formulas are written in English with .Formula2 (Excel translates them
     to the user's language), conditional formats only reference helper cells, so the workbook
     works in any Excel language;
  4. clears personal information (author, last modified by, company) and saves.

All booking data in the workbook is SIMULATED (see data/ASSUMPTIONS.md).
Run from the repository root:  python src/build_excel.py   (Windows + Excel 365 required)
"""
from __future__ import annotations

import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import win32com.client as win32

from common import OUT, ROOT, SIM

XLSX = ROOT / "excel" / "ETG_Commercial_Pricing_Cockpit.xlsx"
TMP = ROOT / ".xltmp"            # neutral working folder for Excel (deleted at the end)
CONN = "ThisWorkbookDataModel"
OUT_SIM = OUT / "sim"
DEFAULT_DAY = "2026-09-15"     # day shown when the workbook opens (day after a partner was lost)
MARGIN_FLOOR = 0.04


def rgb(hex_color: str) -> int:
    """Excel colours are BGR integers."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


C = {k: rgb(v) for k, v in dict(blue="#2A78D6", blue_dark="#1C5CAB", blue_pale="#CDE2FB", grey="#898781",
                                    grey_light="#E1E0D9", page="#F4F4F2", card="#FCFCFB", ink="#0B0B0B",
                                    ink_soft="#52514E", red="#D03B3B", red_pale="#F6D5D5", green="#0CA30C",
                                    yellow_input="#FFF4CE", white="#FFFFFF").items()}



# --------------------------------------------------------------------------- language-proof helpers
# Late-bound COM calls from Python reach Excel with the user's locale: on a French Excel,
# NumberFormat is read as a French format code and Range.Resize does not behave as in VBA.
# These three helpers make the build independent of the Excel language.
_INTL: dict = {}


def localize(code: str) -> str:
    """Translate an English number-format code to the format code of the running Excel."""
    dec, thou = _INTL.get("dec", "."), _INTL.get("thou", ",")
    y, d = _INTL.get("y", "y"), _INTL.get("d", "d")
    out = code.replace(",", "\x00").replace(".", "\x01")
    out = out.replace("\x00", thou).replace("\x01", dec)
    if "yy" in code:
        out = out.replace("y", y).replace("d", d)
    return out


def nfmt(rng, code: str) -> None:
    if code == "General":
        return
    try:
        rng.NumberFormatLocal = localize(code)
    except Exception:                      # PivotField has no NumberFormatLocal
        rng.NumberFormat = localize(code)


def area(rng, nrows: int, ncols: int):
    """Equivalent of VBA Range.Resize(nrows, ncols)."""
    ws = rng.Worksheet
    return ws.Range(ws.Cells(rng.Row, rng.Column), ws.Cells(rng.Row + nrows - 1, rng.Column + ncols - 1))


def xl_date(ts) -> float:
    """Excel serial date (avoids the time-zone shift of COM datetime conversion)."""
    return float((pd.Timestamp(ts).normalize() - pd.Timestamp("1899-12-30")).days)


# =========================================================================== data preparation
def prepare() -> dict[str, pd.DataFrame]:
    c = pd.read_csv(OUT_SIM / "bookings_clean.csv", parse_dates=["booking_date", "cancel_date"])
    keys = ["brand", "partner_type", "line_of_business", "client_country"]
    daily = c.groupby(["booking_date"] + keys).agg(
        bookings=("booking_id", "size"), room_nights=("room_nights", "sum"), sales_eur=("sales_eur", "sum"),
        net_cost_eur=("net_cost_eur", "sum"), margin_eur=("margin_eur", "sum")).reset_index()
    canc = (c[c["status"] == "Cancelled"].groupby(["cancel_date"] + keys).size()
            .rename("cancellations").reset_index().rename(columns={"cancel_date": "booking_date"}))
    daily = daily.merge(canc, on=["booking_date"] + keys, how="outer").fillna(
        {"bookings": 0, "room_nights": 0, "sales_eur": 0, "net_cost_eur": 0, "margin_eur": 0, "cancellations": 0})
    daily = daily[daily["booking_date"] <= c["booking_date"].max()]
    daily = daily.rename(columns={"booking_date": "date"}).sort_values(["date"] + keys)

    c["year_month"] = c["booking_date"].dt.strftime("%Y-%m")
    c["sales_local"] = c["sell_price_local"] - c["discount_eur"] * c["fx_rate_to_eur"]
    mkeys = ["brand", "line_of_business", "client_country"]
    monthly = c.groupby(["year_month"] + mkeys).agg(
        bookings=("booking_id", "size"), sales_eur=("sales_eur", "sum"), sales_local=("sales_local", "sum"),
        net_cost_eur=("net_cost_eur", "sum")).reset_index()
    monthly.insert(4, "segment", monthly[mkeys].agg(" | ".join, axis=1))

    pm = pd.read_csv(SIM / "price_match_requests.csv")
    partners = pd.read_csv(SIM / "partners.csv")
    pm = pm.merge(partners[["partner_id", "partner_type"]], on="partner_id", how="left")
    return {"daily": daily, "monthly": monthly, "pm": pm,
            "dq": pd.read_csv(OUT_SIM / "dq_issues.csv"),
            "raw_clean": pd.read_csv(OUT_SIM / "daily_raw_vs_clean.csv", parse_dates=["booking_date"]),
            "reco": pd.read_csv(OUT_SIM / "pricing_recommendations.csv"),
            "incidents": pd.read_csv(OUT_SIM / "incidents.csv"),
            "dates": pd.date_range(c["booking_date"].min(), c["booking_date"].max()),
            "months": sorted(c["year_month"].unique())}


# =========================================================================== Excel helpers
def start_excel():
    """Start Excel with a neutral TEMP folder inside the project and attach to it through COM.

    Why: the Power Pivot engine writes its working-folder path (by default the user's TEMP folder,
    which contains the Windows user name) into the data model saved in the file. A neutral TEMP keeps
    the workbook free of personal information.
    """
    import os
    import subprocess
    import winreg
    from openpyxl import Workbook

    TMP.mkdir(parents=True, exist_ok=True)
    seed = TMP / "seed.xlsx"
    Workbook().save(seed)
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe") as k:
        exe = winreg.QueryValue(k, None)
    env = dict(os.environ, TEMP=str(TMP), TMP=str(TMP))
    subprocess.Popen([exe, "/x", str(seed)], env=env)
    for _ in range(120):                      # wait until the workbook is registered
        try:
            wb = win32.GetObject(str(seed))
            return wb.Application, wb
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("Excel did not start")


class Book:
    def __init__(self):
        self.xl, self.wb = start_excel()
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.xl.ScreenUpdating = False
        intl = self.xl.International
        _INTL.update(dec=intl[2], thou=intl[3], y=intl[18], d=intl[20])
        self.wb.Styles("Normal").Font.Name = "Segoe UI"
        self.wb.Styles("Normal").Font.Size = 10
        while self.wb.Worksheets.Count > 1:
            self.wb.Worksheets(self.wb.Worksheets.Count).Delete()
        # Guard: dynamic arrays (Excel 365) must be available, otherwise stop with a clear message.
        probe = self.wb.Worksheets(1).Range("A1")
        probe.Formula = "=CONCAT(\"o\",\"k\")"
        ok = probe.Text == "ok"
        probe.ClearContents()
        # Manual calculation while building: CUBE formulas would otherwise keep Excel busy
        # (calls rejected). Everything is calculated once at the end (wait_cube).
        if not ok:
            self.wb.Close(False)
            self.xl.Quit()
            raise RuntimeError("Excel 365 functions are not available (unlicensed or older Excel). "
                               "Sign in to Microsoft 365 in Excel (File > Account) and run again.")

    def sheet(self, name: str, tab: str | None = None):
        if self.wb.Worksheets.Count == 1 and self.wb.Worksheets(1).Name in ("Feuil1", "Sheet1", "Sheet"):
            ws = self.wb.Worksheets(1)
        else:
            ws = self.wb.Worksheets.Add(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        ws.Name = name
        ws.Activate()
        self.xl.ActiveWindow.DisplayGridlines = False
        ws.Cells.Interior.Color = C["page"]
        if tab:
            ws.Tab.Color = C[tab]
        return ws

    # ---------------------------------------------------------------- data model
    def add_query(self, name: str, path: Path, types: dict[str, str], keep: list[str] | None = None,
                  extra_steps: str = ""):
        type_list = ", ".join(f'{{"{k}", {v}}}' for k, v in types.items())
        select = f'Selected = Table.SelectColumns(Promoted, {{{", ".join(chr(34) + k + chr(34) for k in keep)}}}),' \
            if keep else "Selected = Promoted,"
        m = f'''let
    Source = Csv.Document(File.Contents("{path}"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    {select}
    {extra_steps}
    Typed = Table.TransformColumnTypes({"Extra" if extra_steps else "Selected"}, {{{type_list}}}, "en-US")
in
    Typed'''
        self.wb.Queries.Add(name, m)
        self.wb.Connections.Add2(
            f"Query - {name}", f"Power Query load of {path.name}",
            f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""',
            f'"{name}"', 6, True, False)

    def relate(self, many_table, many_col, one_table, one_col):
        m = self.wb.Model
        m.ModelRelationships.Add(m.ModelTables(many_table).ModelTableColumns(many_col),
                                 m.ModelTables(one_table).ModelTableColumns(one_col))

    def measure(self, name: str, formula: str, fmt: str = "number", desc: str = ""):
        m = self.wb.Model
        f = {"number": m.ModelFormatDecimalNumber, "whole": m.ModelFormatWholeNumber,
             "pct": m.ModelFormatPercentageNumber, "eur": m.ModelFormatDecimalNumber}[fmt]
        m.ModelMeasures.Add(name, m.ModelTables("Bookings"), formula, f, desc)


def write_frame(ws, top_left: str, df: pd.DataFrame, table_name: str | None = None, style: str = "TableStyleLight1"):
    """Write a DataFrame (header + values) and optionally turn it into an Excel Table."""
    def conv(v):
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return None
        if isinstance(v, (pd.Timestamp, datetime)):
            return xl_date(v)
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating, float)):
            return None if np.isnan(v) else float(v)
        if isinstance(v, (np.bool_, bool)):
            return bool(v)
        return v
    values = [list(df.columns)] + [[conv(v) for v in row] for row in df.itertuples(index=False)]
    start = ws.Range(top_left)
    r0, c0, ncol = start.Row, start.Column, len(df.columns)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(values) - 1, c0 + ncol - 1))
    chunk = 20000
    for i in range(0, len(values), chunk):
        part = values[i:i + chunk]
        ws.Range(ws.Cells(r0 + i, c0), ws.Cells(r0 + i + len(part) - 1, c0 + ncol - 1)).Value = part
    if table_name:
        lo = ws.ListObjects.Add(1, rng, None, 1)
        lo.Name = table_name
        lo.TableStyle = style
    return rng


def title(ws, text: str, subtitle: str, tag: str, width_cols: str = "N"):
    ws.Range("B1").Value = text
    ws.Range("B1").Font.Size = 16
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Color = C["ink"]
    ws.Range("B2").Value = subtitle
    ws.Range("B2").Font.Color = C["ink_soft"]
    ws.Range(f"{width_cols}1").Value = tag
    ws.Range(f"{width_cols}1").Font.Bold = True
    ws.Range(f"{width_cols}1").Font.Size = 9
    ws.Range(f"{width_cols}1").Font.Color = C["ink_soft"]
    ws.Range(f"{width_cols}1").HorizontalAlignment = -4152  # right
    ws.Rows(1).RowHeight = 28
    ws.Columns("A").ColumnWidth = 2


def header(rng):
    rng.Font.Bold = True
    rng.Font.Color = C["ink_soft"]
    rng.Interior.Color = C["card"]
    rng.Borders(9).LineStyle = 1        # bottom border
    rng.Borders(9).Color = C["grey"]


def card(rng):
    rng.Interior.Color = C["card"]
    for b in (7, 8, 9, 10):              # left, top, bottom, right
        rng.Borders(b).LineStyle = 1
        rng.Borders(b).Color = C["grey_light"]


def input_cell(rng):
    rng.Interior.Color = C["yellow_input"]
    for b in (7, 8, 9, 10):
        rng.Borders(b).LineStyle = 1
        rng.Borders(b).Color = C["grey"]


def red_if(rng, flag_ref: str, font_only: bool = False):
    """Conditional format: red when the helper cell is TRUE (reference only, language-proof)."""
    fc = rng.FormatConditions.Add(2, None, f"={flag_ref}")
    if font_only:
        fc.Font.Color = C["red"]
        fc.Font.Bold = True
    else:
        fc.Interior.Color = C["red_pale"]
        fc.Font.Color = C["red"]


def dropdown(rng, source: str):
    rng.Validation.Delete()
    rng.Validation.Add(3, 1, 1, source)


def footer(ws, row: int, text: str = ""):
    base = ("SIMULATED data (seed 42), not Emerging Travel Group data. Independent case study based on public data "
            "- not affiliated with or endorsed by Emerging Travel Group.")
    ws.Range(f"B{row}").Value = (text + "  |  " if text else "") + base
    ws.Range(f"B{row}").Font.Size = 8
    ws.Range(f"B{row}").Font.Color = C["grey"]


# =========================================================================== build
def build_model(bk: Book) -> None:
    sim, osim = SIM, OUT_SIM
    keep = ["booking_id", "booking_date", "checkin_date", "brand", "line_of_business", "client_country",
            "destination_country", "supplier_id", "partner_id", "partner_type", "channel", "room_nights", "currency",
            "sell_price_local", "fx_rate_to_eur", "sell_price_eur", "net_cost_eur", "discount_eur", "sales_eur",
            "margin_eur", "status", "cancel_date", "cost_imputed"]
    types = {k: "type text" for k in keep}
    types.update({k: "type date" for k in ["booking_date", "checkin_date", "cancel_date"]})
    types.update({k: "type number" for k in ["sell_price_local", "fx_rate_to_eur", "sell_price_eur", "net_cost_eur",
                                             "discount_eur", "sales_eur", "margin_eur"]})
    types["room_nights"] = "Int64.Type"
    types["fx_key"] = "type text"
    bk.add_query("Bookings", osim / "bookings_clean.csv", types, keep,
                 extra_steps='Extra = Table.AddColumn(Selected, "fx_key", each [booking_date] & "|" & [currency]),')
    bk.add_query("Calendar", sim / "calendar.csv",
                 {"date": "type date", "date_key": "type text", "year": "Int64.Type", "month_num": "Int64.Type",
                  "year_month": "type text", "month_name": "type text", "iso_week": "Int64.Type",
                  "weekday": "type text", "weekday_num": "Int64.Type", "is_weekend": "type logical"})
    bk.add_query("Suppliers", sim / "suppliers.csv",
                 {"supplier_id": "type text", "supplier_name": "type text", "supplier_type": "type text",
                  "coverage": "type text", "payment_terms_days": "Int64.Type", "line_of_business": "type text"})
    bk.add_query("Partners", sim / "partners.csv",
                 {"partner_id": "type text", "partner_name": "type text", "partner_type": "type text",
                  "brand": "type text", "country": "type text", "tier": "type text", "commercial_terms": "type text"})
    bk.add_query("FX", sim / "fx_rates.csv",
                 {"date": "type date", "currency": "type text", "units_per_eur": "type number", "fx_key": "type text",
                  "fx_month": "type text"},
                 extra_steps='Extra = Table.AddColumn(Selected, "fx_month", each Text.Start([date], 7)),')
    bk.relate("Bookings", "booking_date", "Calendar", "date")
    bk.relate("Bookings", "supplier_id", "Suppliers", "supplier_id")
    bk.relate("Bookings", "partner_id", "Partners", "partner_id")
    bk.relate("Bookings", "fx_key", "FX", "fx_key")

    base = {
        "Bookings": ("COUNTROWS(Bookings)", "whole", "Bookings created on the date (all statuses)"),
        "Sales EUR": ("SUM(Bookings[sales_eur])", "eur", "Selling price after discount, EUR, booking date"),
        "Net Cost EUR": ("SUM(Bookings[net_cost_eur])", "eur", "Supplier net cost, EUR"),
        "Margin EUR": ("SUM(Bookings[margin_eur])", "eur", "Sales - net cost"),
        "Room Nights": ("SUM(Bookings[room_nights])", "whole", "Hotel room nights"),
        "Discounts EUR": ("SUM(Bookings[discount_eur])", "eur", "Price-match discounts granted"),
        "Cancellations": ("CALCULATE(COUNTROWS(Bookings), ALL('Calendar'), TREATAS(VALUES('Calendar'[date]), "
                          "Bookings[cancel_date]))", "whole", "Bookings cancelled ON the date (cancel date)"),
    }
    for name, (f, fmt, desc) in base.items():
        bk.measure(name, f, fmt, desc)
    bk.measure("Margin %", "DIVIDE([Margin EUR], [Sales EUR])", "pct", "Margin / sales")
    bk.measure("Avg Booking Value", "DIVIDE([Sales EUR], [Bookings])", "eur", "Sales / bookings")
    bk.measure("Avg FX Units per EUR", "AVERAGE(FX[units_per_eur])", "number", "ECB reference rate (real data)")
    bk.measure("Cancellation Rate", "DIVIDE(CALCULATE(COUNTROWS(Bookings), Bookings[status] = \"Cancelled\"), [Bookings])",
               "pct", "Share of bookings made on the date that are cancelled today")

    shifts = {
        "D-1": "'Calendar'[date] = d - 1",
        "LY": "'Calendar'[date] = d - 364",
        "7d Total": "'Calendar'[date] >= d - 7 && 'Calendar'[date] <= d - 1",
        "Exp Total": "'Calendar'[date] = d - 7 || 'Calendar'[date] = d - 14 || 'Calendar'[date] = d - 21 "
                     "|| 'Calendar'[date] = d - 28",
    }
    for mname in ["Bookings", "Sales EUR", "Margin EUR", "Room Nights", "Cancellations"]:
        for sfx, cond in shifts.items():
            bk.measure(f"{mname} {sfx}",
                       f"VAR d = MAX('Calendar'[date]) RETURN CALCULATE([{mname}], FILTER(ALL('Calendar'), {cond}))",
                       base[mname][1], f"{mname}, {sfx}")
        bk.measure(f"{mname} 7d Avg", f"DIVIDE([{mname} 7d Total], 7)", base[mname][1], "Average of the previous 7 days")
        bk.measure(f"{mname} Expected", f"DIVIDE([{mname} Exp Total], 4)", base[mname][1],
                   "Average of the same weekday over the previous 4 weeks")
    for sfx, (num, den) in {"D-1": ("D-1", "D-1"), "LY": ("LY", "LY"), "7d Avg": ("7d Total", "7d Total"),
                            "Expected": ("Exp Total", "Exp Total")}.items():
        bk.measure(f"Margin % {sfx}", f"DIVIDE([Margin EUR {num}], [Sales EUR {den}])", "pct", f"Margin %, {sfx}")
        bk.measure(f"Avg Booking Value {sfx}", f"DIVIDE([Sales EUR {num}], [Bookings {den}])", "eur", f"ABV, {sfx}")
    bk.wb.Model.Refresh()


def build_lists(bk: Book, d: dict) -> None:
    ws = bk.sheet("Lists", "grey")
    ws.Range("A1").Value = "date_key"
    keys = [[x.strftime("%Y-%m-%d")] for x in d["dates"]]
    nfmt(area(ws.Range("A2"), len(keys), 1), "@")
    area(ws.Range("A2"), len(keys), 1).Value = keys
    lists = {
        "B": ["brand", "All", "RateHawk", "ZenHotels", "Roundtrip"],
        "C": ["line_of_business", "All", "Hotels", "Transfers", "Car rentals", "Flights"],
        "D": ["dimension", "Total", "brand", "partner_type", "line_of_business", "client_country"],
        "E": ["metric", "bookings", "sales_eur", "margin_rate"],
        "F": ["year_month"] + d["months"],
        "G": ["yes_no", "Yes", "No"],
        "I": ["meal", "Room only", "Breakfast"],
        "J": ["cancellation", "Free cancellation", "Non-refundable"],
    }
    for col, vals in lists.items():
        nfmt(area(ws.Range(f"{col}1"), len(vals), 1), "@")
        area(ws.Range(f"{col}1"), len(vals), 1).Value = [[v] for v in vals]
    ws.Range("H1").Value = "members (depends on Alerts!C5)"   # formula written once Alerts exists
    header(ws.Range("A1:J1"))
    ws.Columns("A:J").AutoFit()


def build_data_sheets(bk: Book, d: dict) -> None:
    ws = bk.sheet("Data_Daily", "grey")
    write_frame(ws, "A1", d["daily"], "tblDaily")
    nfmt(ws.Range("A:A"), "yyyy-mm-dd")
    nfmt(ws.Range("F:J"), "#,##0.00")
    ws.Columns("A:K").AutoFit()
    ws = bk.sheet("Data_Monthly", "grey")
    write_frame(ws, "A1", d["monthly"], "tblMonthly")
    nfmt(ws.Range("G:I"), "#,##0.00")
    ws.Columns("A:I").AutoFit()


def build_cockpit(bk: Book, d: dict) -> None:
    ws = bk.sheet("Cockpit", "blue")
    title(ws, "Daily commercial cockpit", "Pick a day, a brand and a line of business. Every number below is a CUBE "
          "formula reading the Power Pivot data model (measures in DAX).", "SIMULATED DATA", "L")
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C:L").ColumnWidth = 13
    # ---- selectors
    for r, (label, default, src) in enumerate([("Day", DEFAULT_DAY, "=Lists!$A$2:$A$" + str(len(d["dates"]) + 1)),
                                               ("Brand", "All", "=Lists!$B$2:$B$5"),
                                               ("Line of business", "All", "=Lists!$C$2:$C$6")], start=4):
        ws.Range(f"B{r}").Value = label
        ws.Range(f"B{r}").Font.Bold = True
        nfmt(ws.Range(f"C{r}"), "@")
        ws.Range(f"C{r}").Value = default
        input_cell(ws.Range(f"C{r}"))
        dropdown(ws.Range(f"C{r}"), src)
    ws.Range("D4").Formula2 = '=INDEX({"Mon","Tue","Wed","Thu","Fri","Sat","Sun"},WEEKDAY(DATE(LEFT(C4,4),MID(C4,6,2),RIGHT(C4,2)),2))'
    ws.Range("D4").Font.Color = C["grey"]
    # helper members (column O, hidden)
    ws.Range("O3").Value = "CUBE members (helpers)"
    ws.Range("O4").Formula2 = f'=CUBEMEMBER("{CONN}","[Calendar].[date_key].&["&$C$4&"]")'
    ws.Range("O5").Formula2 = (f'=IF($C$5="All",CUBEMEMBER("{CONN}","[Bookings].[brand].[All]"),'
                               f'CUBEMEMBER("{CONN}","[Bookings].[brand].&["&$C$5&"]"))')
    ws.Range("O6").Formula2 = (f'=IF($C$6="All",CUBEMEMBER("{CONN}","[Bookings].[line_of_business].[All]"),'
                               f'CUBEMEMBER("{CONN}","[Bookings].[line_of_business].&["&$C$6&"]"))')
    ws.Range("B7").Value = "Thresholds:"
    ws.Range("B7").Font.Color = C["grey"]
    ws.Range("C7").Value = 0.25
    nfmt(ws.Range("C7"), "0%")
    ws.Range("D7").Value = "volume gap vs expected"
    ws.Range("F7").Value = 0.015
    nfmt(ws.Range("F7"), "0.0%")
    ws.Range("G7").Value = "margin-rate drop (pts)"
    ws.Range("I7").Value = MARGIN_FLOOR
    nfmt(ws.Range("I7"), "0.0%")
    ws.Range("J7").Value = "margin floor"
    ws.Range("B7:L7").Font.Size = 8
    ws.Range("B7:L7").Font.Color = C["grey"]
    for cell in ("C7", "F7", "I7"):
        input_cell(ws.Range(cell))

    # ---- KPI table
    hdr = ["Metric", "Selected day", "Day before", "vs D-1", "Same day LY", "vs LY", "Prev. 7-day avg",
           "Expected*", "vs expected", "Status"]
    area(ws.Range("B9"), 1, len(hdr)).Value = [hdr]
    header(area(ws.Range("B9"), 1, len(hdr)))
    kpis = [("Bookings", "Bookings", "#,##0"), ("Sales (EUR)", "Sales EUR", "#,##0"),
            ("Margin (EUR)", "Margin EUR", "#,##0"), ("Margin rate", "Margin %", "0.0%"),
            ("Room nights", "Room Nights", "#,##0"), ("Avg booking value (EUR)", "Avg Booking Value", "#,##0"),
            ("Cancellations (by cancel date)", "Cancellations", "#,##0")]
    flt = "$O$4,$O$5,$O$6"
    for i, (label, m, fmt) in enumerate(kpis):
        r = 10 + i
        ws.Range(f"B{r}").Value = label
        cv = lambda meas: f'=CUBEVALUE("{CONN}","[Measures].[{meas}]",{flt})'
        ws.Range(f"C{r}").Formula2 = cv(m)
        ws.Range(f"D{r}").Formula2 = cv(f"{m} D-1")
        ws.Range(f"F{r}").Formula2 = cv(f"{m} LY")
        ws.Range(f"H{r}").Formula2 = cv(f"{m} 7d Avg")
        ws.Range(f"I{r}").Formula2 = cv(f"{m} Expected")
        if fmt == "0.0%":
            ws.Range(f"E{r}").Formula2 = f'=IFERROR(C{r}-D{r},"")'
            ws.Range(f"G{r}").Formula2 = f'=IFERROR(C{r}-F{r},"")'
            ws.Range(f"J{r}").Formula2 = f'=IFERROR(C{r}-I{r},"")'
            for col in "EGJ":
                nfmt(ws.Range(f"{col}{r}"), '+0.0%;-0.0%;0.0%')
            ws.Range(f"K{r}").Formula2 = (f'=IF(C{r}="","",IF(C{r}<$I$7,"Below floor",'
                                          f'IF(C{r}-I{r}<=-$F$7,"Alert","OK")))')
        else:
            ws.Range(f"E{r}").Formula2 = f'=IFERROR(C{r}/D{r}-1,"")'
            ws.Range(f"G{r}").Formula2 = f'=IFERROR(C{r}/F{r}-1,"")'
            ws.Range(f"J{r}").Formula2 = f'=IFERROR(C{r}/I{r}-1,"")'
            for col in "EGJ":
                nfmt(ws.Range(f"{col}{r}"), '+0%;-0%;0%')
            ws.Range(f"K{r}").Formula2 = f'=IF(J{r}="","",IF(ABS(J{r})>$C$7,"Alert","OK"))'
        for col in "CDFHI":
            nfmt(ws.Range(f"{col}{r}"), fmt)
        ws.Range(f"P{r}").Formula2 = f'=OR(K{r}="Alert",K{r}="Below floor")'
        red_if(ws.Range(f"K{r}"), f"$P{r}")
    card(ws.Range("B9:K16"))
    ws.Range("B17").Value = ("* Expected = average of the same weekday over the previous 4 weeks (handles the weekly "
                             "pattern). LY = same weekday 52 weeks earlier (364 days).")
    ws.Range("B17").Font.Size = 8
    ws.Range("B17").Font.Color = C["grey"]

    # ---- segment tables
    def segment_block(top: int, caption: str, hierarchy: str, members: list[str], use_brand=True, use_lob=True):
        ws.Range(f"B{top}").Value = caption
        ws.Range(f"B{top}").Font.Bold = True
        ws.Range(f"B{top}").Font.Size = 11
        h = ["Segment", "Bookings", "Sales (EUR)", "Margin rate", "Sales expected", "Sales vs exp.",
             "Margin exp.", "Margin vs exp.", "Status"]
        area(ws.Range(f"B{top + 1}"), 1, len(h)).Value = [h]
        header(area(ws.Range(f"B{top + 1}"), 1, len(h)))
        f = "$O$4" + (",$O$5" if use_brand else "") + (",$O$6" if use_lob else "")
        for i, mem in enumerate(members):
            r = top + 2 + i
            ws.Range(f"B{r}").Formula2 = f'=CUBEMEMBER("{CONN}","{hierarchy}.&[{mem}]")'
            cv = lambda meas: f'=CUBEVALUE("{CONN}","[Measures].[{meas}]",{f},$B{r})'
            ws.Range(f"C{r}").Formula2 = cv("Bookings")
            ws.Range(f"D{r}").Formula2 = cv("Sales EUR")
            ws.Range(f"E{r}").Formula2 = cv("Margin %")
            ws.Range(f"F{r}").Formula2 = cv("Sales EUR Expected")
            ws.Range(f"G{r}").Formula2 = f'=IFERROR(D{r}/F{r}-1,"")'
            ws.Range(f"H{r}").Formula2 = cv("Margin % Expected")
            ws.Range(f"I{r}").Formula2 = f'=IFERROR(E{r}-H{r},"")'
            ws.Range(f"J{r}").Formula2 = (f'=IF(D{r}="","",IF(AND(E{r}<>"",E{r}<$I$7),"Below floor",'
                                          f'IF(OR(AND(G{r}<>"",ABS(G{r})>$C$7),AND(I{r}<>"",I{r}<=-$F$7)),"Alert","OK")))')
            ws.Range(f"P{r}").Formula2 = f'=OR(J{r}="Alert",J{r}="Below floor")'
            red_if(ws.Range(f"J{r}"), f"$P{r}")
        last = top + 1 + len(members)
        nfmt(ws.Range(f"C{top + 2}:D{last}"), "#,##0")
        nfmt(ws.Range(f"F{top + 2}:F{last}"), "#,##0")
        nfmt(ws.Range(f"E{top + 2}:E{last}"), "0.0%")
        nfmt(ws.Range(f"H{top + 2}:H{last}"), "0.0%")
        nfmt(ws.Range(f"G{top + 2}:G{last}"), "+0%;-0%;0%")
        nfmt(ws.Range(f"I{top + 2}:I{last}"), "+0.0%;-0.0%;0.0%")
        card(ws.Range(f"B{top + 1}:J{last}"))
        return last + 2

    r = segment_block(19, "By brand", "[Bookings].[brand]", ["RateHawk", "ZenHotels", "Roundtrip"], use_brand=False)
    r = segment_block(r, "By line of business", "[Bookings].[line_of_business]",
                      ["Hotels", "Transfers", "Car rentals", "Flights"], use_lob=False)
    r = segment_block(r, "By partner type", "[Bookings].[partner_type]",
                      ["Agency", "API partner", "White label", "Corporate", "B2C"])
    r = segment_block(r, "By supplier type (Suppliers table)", "[Suppliers].[supplier_type]",
                      ["Direct", "Bedbank", "Wholesaler", "Ground", "Air"])
    r = segment_block(r, "By source market (client country)", "[Bookings].[client_country]",
                      ["DE", "GB", "FR", "IT", "ES", "NL", "PL", "TR", "AE", "US", "BR", "MX", "IN"])
    r = segment_block(r, "By destination", "[Bookings].[destination_country]",
                      ["ES", "IT", "FR", "EL", "PT", "DE", "HR", "TR", "AE", "TH"])

    # ---- top partners with CUBESET + CUBERANKEDMEMBER
    ws.Range(f"B{r}").Value = "Top 10 partners of the selected day (CUBESET sorted by sales + CUBERANKEDMEMBER)"
    ws.Range(f"B{r}").Font.Bold = True
    ws.Range(f"B{r}").Font.Size = 11
    ws.Range(f"O{r}").Formula2 = (f'=CUBESET("{CONN}",'
                                  f'"TOPCOUNT([Bookings].[partner_id].children,10,([Measures].[Sales EUR],'
                                  f'[Calendar].[date_key].&["&$C$4&"]))","Top partners")')
    h = ["Partner", "Bookings", "Sales (EUR)", "Margin rate", "Sales expected", "Sales vs exp.", "Tier"]
    area(ws.Range(f"B{r + 1}"), 1, len(h)).Value = [h]
    header(area(ws.Range(f"B{r + 1}"), 1, len(h)))
    for i in range(10):
        rr = r + 2 + i
        ws.Range(f"B{rr}").Formula2 = f'=CUBERANKEDMEMBER("{CONN}",$O${r},{i + 1})'
        ws.Range(f"C{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Bookings]",$O$4,$B{rr})'
        ws.Range(f"D{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Sales EUR]",$O$4,$B{rr})'
        ws.Range(f"E{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Margin %]",$O$4,$B{rr})'
        ws.Range(f"F{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Sales EUR Expected]",$O$4,$B{rr})'
        ws.Range(f"G{rr}").Formula2 = f'=IFERROR(D{rr}/F{rr}-1,"")'
        ws.Range(f"H{rr}").Formula2 = f'=XLOOKUP(B{rr},tblPartners[partner_id],tblPartners[tier],"")'
    nfmt(ws.Range(f"C{r + 2}:D{r + 11}"), "#,##0")
    nfmt(ws.Range(f"F{r + 2}:F{r + 11}"), "#,##0")
    nfmt(ws.Range(f"E{r + 2}:E{r + 11}"), "0.0%")
    nfmt(ws.Range(f"G{r + 2}:G{r + 11}"), "+0%;-0%;0%")
    card(ws.Range(f"B{r + 1}:H{r + 11}"))
    r += 13

    # ---- 28-day trend (CUBEVALUE per date) + chart
    ws.Range(f"B{r}").Value = "Last 28 days: sales vs same weekday last year"
    ws.Range(f"B{r}").Font.Bold = True
    ws.Range(f"B{r}").Font.Size = 11
    area(ws.Range(f"B{r + 1}"), 1, 4).Value = [["Day", "Sales (EUR)", "Sales LY (EUR)", "Margin rate"]]
    header(area(ws.Range(f"B{r + 1}"), 1, 4))
    for i in range(28):
        rr = r + 2 + i
        ws.Range(f"B{rr}").Formula2 = (f'=INDEX(Lists!$A:$A,MATCH($C$4,Lists!$A:$A,0)-{27 - i})')
        ws.Range(f"N{rr}").Formula2 = f'=CUBEMEMBER("{CONN}","[Calendar].[date_key].&["&$B{rr}&"]")'
        ws.Range(f"C{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Sales EUR]",$N{rr},$O$5,$O$6)'
        ws.Range(f"D{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Sales EUR LY]",$N{rr},$O$5,$O$6)'
        ws.Range(f"E{rr}").Formula2 = f'=CUBEVALUE("{CONN}","[Measures].[Margin %]",$N{rr},$O$5,$O$6)'
    nfmt(ws.Range(f"C{r + 2}:D{r + 29}"), "#,##0")
    nfmt(ws.Range(f"E{r + 2}:E{r + 29}"), "0.0%")
    ch = ws.ChartObjects().Add(ws.Range(f"G{r}").Left, ws.Range(f"G{r}").Top, 520, 300).Chart
    ch.ChartType = 4  # xlLine
    ch.SetSourceData(ws.Range(f"B{r + 1}:D{r + 29}"))
    ch.HasTitle = True
    ch.ChartTitle.Text = "Daily sales (EUR): selected period vs last year"
    style_chart(ch, [C["blue"], C["grey"]])
    footer(ws, r + 31, "Cockpit = CUBE formulas on the data model")
    ws.Columns("N:P").Hidden = True
    ws.Range("B9").Select()
    bk.xl.ActiveWindow.FreezePanes = False


def style_chart(ch, colors):
    ch.ChartArea.Format.Fill.ForeColor.RGB = C["card"]
    ch.ChartArea.Format.Line.ForeColor.RGB = C["grey_light"]
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
    if ch.HasTitle:
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
    for i, col in enumerate(colors, start=1):
        if i <= ch.SeriesCollection().Count:
            s = ch.SeriesCollection(i)
            try:
                s.Format.Line.ForeColor.RGB = col
            except Exception:
                pass
            try:
                s.Format.Fill.ForeColor.RGB = col
            except Exception:
                pass
    try:
        ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = C["grey_light"]
    except Exception:
        pass
    ch.HasLegend = True
    ch.Legend.Position = -4107  # bottom


def build_pivots(bk: Book) -> None:
    ws = bk.sheet("Pivots", "blue")
    title(ws, "PivotTables on the data model", "Use the slicers: they filter both PivotTables. "
          "Measures come from the Power Pivot model (DAX).", "SIMULATED DATA", "P")
    wb = bk.wb
    pc = wb.PivotCaches().Create(2, wb.Connections(CONN), 6)
    pt = pc.CreatePivotTable(ws.Range("B6"), "ptMonthlyBrand", True, 6)
    pt.CubeFields("[Calendar].[year_month]").Orientation = 1
    pt.CubeFields("[Bookings].[brand]").Orientation = 2
    nfmt(pt.AddDataField(pt.CubeFields("[Measures].[Sales EUR]")), "#,##0")
    nfmt(pt.AddDataField(pt.CubeFields("[Measures].[Margin %]")), "0.0%")
    pt.TableStyle2 = "PivotStyleLight16"
    pt.CompactLayoutRowHeader = "Month"
    pt.CompactLayoutColumnHeader = "Brand"
    pc2 = wb.PivotCaches().Create(2, wb.Connections(CONN), 6)
    pt2 = pc2.CreatePivotTable(ws.Range("B34"), "ptSupplierMargin", True, 6)
    pt2.CubeFields("[Suppliers].[supplier_type]").Orientation = 1
    pt2.CubeFields("[Suppliers].[supplier_name]").Orientation = 1
    pt2.CubeFields("[Calendar].[year_month]").Orientation = 2
    nfmt(pt2.AddDataField(pt2.CubeFields("[Measures].[Margin %]")), "0.0%")
    pt2.TableStyle2 = "PivotStyleLight16"
    pt2.CompactLayoutRowHeader = "Supplier"
    pt2.CompactLayoutColumnHeader = "Month"
    slicers = [("[Bookings].[line_of_business]", "Line of business"), ("[Bookings].[client_country]", "Source market"),
               ("[Bookings].[partner_type]", "Partner type"), ("[Bookings].[destination_country]", "Destination"),
               ("[Partners].[tier]", "Partner tier"), ("[Bookings].[channel]", "Channel")]
    left = ws.Range("N4").Left
    for i, (field, cap) in enumerate(slicers):
        sc = wb.SlicerCaches.Add2(pt, field)
        level = field + "." + field.split(".")[1]
        sc.Slicers.Add(ws, level, f"sl_{i}", cap, 60 + (i // 2) * 175, left + (i % 2) * 150, 140, 165)
        sc.PivotTables.AddPivotTable(pt2)
    # FX pivot (real ECB rates) for the currencies in the ledger
    pc3 = wb.PivotCaches().Create(2, wb.Connections(CONN), 6)
    ws.Range("B70").Value = "Real ECB reference rates (units per EUR), monthly average - FX table of the model (PUBLIC DATA)"
    ws.Range("B70").Font.Bold = True
    pt3 = pc3.CreatePivotTable(ws.Range("B72"), "ptFX", True, 6)
    pt3.CubeFields("[FX].[fx_month]").Orientation = 1
    pt3.CubeFields("[FX].[currency]").Orientation = 2
    nfmt(pt3.AddDataField(pt3.CubeFields("[Measures].[Avg FX Units per EUR]")), "#,##0.000")
    pt3.TableStyle2 = "PivotStyleLight16"
    pt3.CompactLayoutRowHeader = "Month"
    pt3.CompactLayoutColumnHeader = "Currency"
    footer(ws, 4, "Slicers on the right")


def build_alerts(bk: Book, d: dict) -> None:
    ws = bk.sheet("Alerts", "red")
    bk.wb.Worksheets("Lists").Range("H2").Formula2 = (
        '=IF(Alerts!$C$5="Total",{"All"},SORT(UNIQUE(CHOOSECOLS(tblDaily,MATCH(Alerts!$C$5,tblDaily[#Headers],0)))))')
    title(ws, "Alert band: is today normal?", "Expected = average of the same weekday over the previous 4 weeks. "
          "Noise = standard deviation of the daily deviation over the previous 56 days. Alert when |z| >= threshold.",
          "SIMULATED DATA", "P")
    labels = [("Dimension", "client_country", "=Lists!$D$2:$D$6"), ("Member", "AE", "=Lists!$H$2#"),
              ("Metric", "bookings", "=Lists!$E$2:$E$4")]
    for i, (lab, val, src) in enumerate(labels):
        r = 5 + i
        ws.Range(f"B{r}").Value = lab
        ws.Range(f"B{r}").Font.Bold = True
        nfmt(ws.Range(f"C{r}"), "@")
        ws.Range(f"C{r}").Value = val
        input_cell(ws.Range(f"C{r}"))
        dropdown(ws.Range(f"C{r}"), src)
    ws.Range("B8").Value = "Alert threshold |z|"
    ws.Range("C8").Value = 3.5
    ws.Range("B9").Value = "Min. expected bookings/day"
    ws.Range("C9").Value = 20
    input_cell(ws.Range("C8:C9"))
    ws.Range("E5").Value = "Criteria (helper)"
    ws.Range("F5").Formula2 = '=IF($C$5="Total","*",$C$6)'
    ws.Range("E6").Value = "Column (helper)"
    ws.Range("F6").Formula2 = '=IF($C$5="Total","brand",$C$5)'
    ws.Range("E5:F6").Font.Color = C["grey"]
    ws.Range("H5").Value = "Alerts in the period"
    ws.Range("H5").Font.Bold = True
    ws.Range("I5").Formula2 = "=COUNTIF($L$37:$L$9999,TRUE)"
    ws.Range("I5").Font.Size = 14
    ws.Range("I5").Font.Bold = True
    ws.Range("I5").Font.Color = C["red"]
    ws.Range("H6").Value = "Alert days (spill)"
    n = len(d["dates"])
    first, last = 37, 37 + n - 1                     # table below the chart
    ws.Range("I6").Formula2 = f'=IFERROR(TRANSPOSE(FILTER($B${first}:$B${last},$L${first}:$L${last})),"none")'
    nfmt(area(ws.Range("I6"), 1, 20), "yyyy-mm-dd")
    area(ws.Range("I6"), 1, 20).Font.Color = C["red"]

    h = ["Date", "Bookings", "Sales (EUR)", "Margin (EUR)", "Margin rate", "Value", "Expected", "Deviation",
         "Noise (sd)", "z", "Alert", "Low band", "High band", "Bookings exp."]
    area(ws.Range("B36"), 1, len(h)).Value = [h]
    header(area(ws.Range("B36"), 1, len(h)))
    area(ws.Range(f"B{first}"), n, 1).Value = [[xl_date(x)] for x in d["dates"]]
    nfmt(ws.Range(f"B{first}:B{last}"), "yyyy-mm-dd")
    sumifs = lambda col: (f'=LET(dimcol,INDEX(tblDaily,0,MATCH($F$6,tblDaily[#Headers],0)),'
                          f'SUMIFS(tblDaily[{col}],tblDaily[date],$B{first},dimcol,$F$5))')
    ws.Range(f"C{first}:C{last}").Formula2 = sumifs("bookings")
    ws.Range(f"D{first}:D{last}").Formula2 = sumifs("sales_eur")
    ws.Range(f"E{first}:E{last}").Formula2 = sumifs("margin_eur")
    ws.Range(f"F{first}:F{last}").Formula2 = f'=IFERROR(E{first}/D{first},"")'
    ws.Range(f"G{first}:G{last}").Formula2 = (f'=CHOOSE(MATCH($C$7,{{"bookings","sales_eur","margin_rate"}},0),'
                                              f'C{first},D{first},F{first})')
    s = first + 28                                   # first row with 4 previous same weekdays
    ws.Range(f"H{s}:H{last}").Formula2 = f'=IFERROR(AVERAGE(G{s - 7},G{s - 14},G{s - 21},G{s - 28}),"")'
    ws.Range(f"O{s}:O{last}").Formula2 = f'=AVERAGE(C{s - 7},C{s - 14},C{s - 21},C{s - 28})'
    ws.Range(f"I{s}:I{last}").Formula2 = (f'=IF(OR(G{s}="",H{s}=""),"",IF($C$7="margin_rate",G{s}-H{s},'
                                          f'G{s}/H{s}-1))')
    ws.Range(f"J{s + 1}:J{last}").Formula2 = (f'=IF(COUNT(I${s}:I{s})<28,"",'
                                              f'STDEV.S(INDEX(I:I,MAX({s},ROW()-56)):I{s}))')
    ws.Range(f"K{s + 1}:K{last}").Formula2 = f'=IF(OR(J{s + 1}="",I{s + 1}=""),"",I{s + 1}/J{s + 1})'
    ws.Range(f"L{first}:L{last}").Formula2 = (f'=IF(K{first}="",FALSE,AND(ABS(K{first})>=$C$8,O{first}>=$C$9,'
                                              f'IF($C$7="margin_rate",ABS(I{first})>=0.015,TRUE)))')
    ws.Range(f"M{s + 1}:M{last}").Formula2 = (f'=IF(J{s + 1}="","",IF($C$7="margin_rate",H{s + 1}-$C$8*J{s + 1},'
                                              f'H{s + 1}*(1-$C$8*J{s + 1})))')
    ws.Range(f"N{s + 1}:N{last}").Formula2 = (f'=IF(J{s + 1}="","",IF($C$7="margin_rate",H{s + 1}+$C$8*J{s + 1},'
                                              f'H{s + 1}*(1+$C$8*J{s + 1})))')
    nfmt(ws.Range(f"C{first}:E{last}"), "#,##0")
    nfmt(ws.Range(f"F{first}:F{last}"), "0.0%")
    nfmt(ws.Range(f"G{first}:H{last}"), "#,##0.0##")
    nfmt(ws.Range(f"M{first}:N{last}"), "#,##0.0##")
    nfmt(ws.Range(f"I{first}:I{last}"), "+0.0%;-0.0%")
    nfmt(ws.Range(f"J{first}:J{last}"), "0.0%")
    nfmt(ws.Range(f"K{first}:K{last}"), "0.0")
    nfmt(ws.Range(f"O{first}:O{last}"), "0.0")
    red_if(ws.Range(f"B{first}:L{last}"), f"$L{first}")
    ws.Columns("B:O").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 19
    # chart: last 200 days of value, expected and band (above the table)
    r0 = last - 199
    ch = ws.ChartObjects().Add(ws.Range("B12").Left, ws.Range("B12").Top, 1000, 330).Chart
    ch.ChartType = 4
    for col, name, color in [("G", "Value", C["blue"]), ("H", "Expected", C["grey"]),
                             ("M", "Low band", C["red"]), ("N", "High band", C["red"])]:
        srs = ch.SeriesCollection().NewSeries()
        srs.Name = name
        srs.Values = ws.Range(f"{col}{r0}:{col}{last}")
        srs.XValues = ws.Range(f"B{r0}:B{last}")
        srs.Format.Line.ForeColor.RGB = color
        if name.endswith("band"):
            srs.Format.Line.DashStyle = 4   # dash
            srs.Format.Line.Weight = 1
    ch.HasTitle = True
    ch.ChartTitle.Text = "Last 200 days: value, expected and alert band (dashed)"
    style_chart(ch, [])
    footer(ws, 10, "Same rule as src/analysis.py and sql/03_alerts_same_weekday.sql")


def build_bridge(bk: Book, d: dict) -> None:
    ws = bk.sheet("Bridge", "blue")
    title(ws, "Margin bridge: volume, mix, price, FX, cost", "Pick two months. Segments = brand x line of business x "
          "source market. The five effects add up exactly to the margin change.", "SIMULATED DATA", "R")
    ws.Range("B4").Value = "Month 0 (base)"
    ws.Range("B5").Value = "Month 1 (compared)"
    for cell, val in (("C4", "2025-09"), ("C5", "2026-09")):
        nfmt(ws.Range(cell), "@")
        ws.Range(cell).Value = val
        input_cell(ws.Range(cell))
        dropdown(ws.Range(cell), f"=Lists!$F$2:$F${len(d['months']) + 1}")
    ws.Range("B4:B5").Font.Bold = True
    # segment table (spilled arrays)
    top = 32
    h = ["Segment", "Bookings 0", "Bookings 1", "Sales 0", "Sales 1", "Local sales 0", "Local sales 1", "Cost 0",
         "Cost 1", "In both", "Unit margin 0", "Volume", "Mix", "Price", "FX", "Cost"]
    area(ws.Range(f"B{top - 1}"), 1, len(h)).Value = [h]
    header(area(ws.Range(f"B{top - 1}"), 1, len(h)))
    A = f"B{top}#"
    sm = lambda col, month: f'=SUMIFS(tblMonthly[{col}],tblMonthly[segment],{A},tblMonthly[year_month],{month})'
    ws.Range(f"B{top}").Formula2 = "=SORT(UNIQUE(tblMonthly[segment]))"
    ws.Range(f"C{top}").Formula2 = sm("bookings", "$C$4")
    ws.Range(f"D{top}").Formula2 = sm("bookings", "$C$5")
    ws.Range(f"E{top}").Formula2 = sm("sales_eur", "$C$4")
    ws.Range(f"F{top}").Formula2 = sm("sales_eur", "$C$5")
    ws.Range(f"G{top}").Formula2 = sm("sales_local", "$C$4")
    ws.Range(f"H{top}").Formula2 = sm("sales_local", "$C$5")
    ws.Range(f"I{top}").Formula2 = sm("net_cost_eur", "$C$4")
    ws.Range(f"J{top}").Formula2 = sm("net_cost_eur", "$C$5")
    N0, N1, S0, S1, L0, L1, C0, C1, V = (f"{c}{top}#" for c in "CDEFGHIJK")
    ws.Range(f"K{top}").Formula2 = f"=({N0}>0)*({N1}>0)"
    ws.Range(f"L{top}").Formula2 = f"=IF({V}=1,({S0}-{C0})/{N0},0)"
    # totals (only segments present in both months)
    tot = {"N0": ("Bookings month 0", f"=SUMPRODUCT({N0},{V})"), "N1": ("Bookings month 1", f"=SUMPRODUCT({N1},{V})"),
           "M0": ("Margin month 0", f"=SUMPRODUCT({S0}-{C0},{V})"), "M1": ("Margin month 1", f"=SUMPRODUCT({S1}-{C1},{V})")}
    for i, (k, (lab, f)) in enumerate(tot.items()):
        ws.Range(f"E{4 + i}").Value = lab
        ws.Range(f"F{4 + i}").Formula2 = f
        nfmt(ws.Range(f"F{4 + i}"), "#,##0")
    ws.Range("E8").Value = "Avg unit margin month 0"
    ws.Range("F8").Formula2 = "=F6/F4"
    nfmt(ws.Range("F8"), "#,##0.00")
    ws.Range("E4:E8").Font.Color = C["ink_soft"]
    ws.Range(f"M{top}").Formula2 = f"={V}*({N1}-{N0})*$F$8"
    ws.Range(f"N{top}").Formula2 = f"={V}*$F$5*({N1}/$F$5-{N0}/$F$4)*({'L' + str(top) + '#'}-$F$8)"
    ws.Range(f"O{top}").Formula2 = f"=IF({V}=1,{N1}*(({L1}/{N1})/({L0}/{S0})-{S0}/{N0}),0)"
    ws.Range(f"P{top}").Formula2 = f"=IF({V}=1,{N1}*({S1}/{N1}-({L1}/{N1})/({L0}/{S0})),0)"
    ws.Range(f"Q{top}").Formula2 = f"=IF({V}=1,-{N1}*({C1}/{N1}-{C0}/{N0}),0)"
    nfmt(ws.Range(f"C{top}:Q{top + 400}"), "#,##0")
    nfmt(ws.Range(f"L{top}:L{top + 400}"), "#,##0.00")
    # bridge summary + waterfall helper
    area(ws.Range("H4"), 1, 5).Value = [["Step", "EUR", "Base", "Up", "Down"]]
    header(ws.Range("H4:L4"))
    steps = [("=\"Margin \"&$C$4", "=F6"), ("Volume", f"=SUM(M{top}#)"), ("Mix", f"=SUM(N{top}#)"),
             ("Price", f"=SUM(O{top}#)"), ("FX", f"=SUM(P{top}#)"), ("Cost", f"=SUM(Q{top}#)"),
             ("=\"Margin \"&$C$5", "=F7")]
    for i, (lab, f) in enumerate(steps):
        r = 5 + i
        ws.Range(f"H{r}").Formula2 = lab
        ws.Range(f"I{r}").Formula2 = f
        if i in (0, 6):
            ws.Range(f"J{r}").Value = 0
            ws.Range(f"K{r}").Formula2 = f"=I{r}"
            ws.Range(f"L{r}").Value = 0
        else:
            # running total before this step = start + previous effects
            ws.Range(f"J{r}").Formula2 = f"=MIN(SUM($I$5:I{r - 1}),SUM($I$5:I{r}))"
            ws.Range(f"K{r}").Formula2 = f"=MAX(I{r},0)"
            ws.Range(f"L{r}").Formula2 = f"=MAX(-I{r},0)"
    ws.Range("H12").Value = "Check (must be 0)"
    ws.Range("I12").Formula2 = "=ROUND(SUM(I5:I10)-I11,6)"
    nfmt(ws.Range("I5:L11"), "#,##0")
    card(ws.Range("H4:L11"))
    ch = ws.ChartObjects().Add(ws.Range("N3").Left, ws.Range("N3").Top, 520, 300).Chart
    ch.ChartType = 52   # stacked column: invisible base + up + down = waterfall
    for col, name in (("J", "Base"), ("K", "Increase"), ("L", "Decrease")):
        srs = ch.SeriesCollection().NewSeries()
        srs.Name = name
        srs.Values = ws.Range(f"{col}5:{col}11")
        srs.XValues = ws.Range("H5:H11")
    ch.SeriesCollection(1).Format.Fill.Visible = False
    ch.SeriesCollection(1).Format.Line.Visible = False
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = C["blue"]
    ch.SeriesCollection(3).Format.Fill.ForeColor.RGB = C["red"]
    ch.ChartGroups(1).GapWidth = 60
    ch.HasTitle = True
    ch.ChartTitle.Text = "Margin bridge (EUR)"
    style_chart(ch, [])
    ch.Legend.LegendEntries(1).Delete()
    ws.Range(f"B{top - 3}").Value = ("Per segment s: Volume = (N1-N0) x avg unit margin 0; Mix = N1 x (w1-w0) x (m0_s - avg m0); "
                                     "Price = N1_s x (P1 at month-0 FX - P0); FX = N1_s x (P1 - P1 at month-0 FX); "
                                     "Cost = -N1_s x (C1 - C0). N = bookings, w = share, P/C = sales/cost per booking.")
    ws.Range(f"B{top - 3}").Font.Size = 8
    ws.Range(f"B{top - 3}").Font.Color = C["grey"]
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C:Q").ColumnWidth = 11
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("H").ColumnWidth = 16
    footer(ws, 27, "Same method as src/analysis.py margin_bridge()")


def build_price_match(bk: Book, d: dict) -> None:
    ws = bk.sheet("PriceMatch_Calc", "green")
    title(ws, "Price-match request calculator", "Enter the request (yellow cells). The sheet checks the evidence, "
          "makes the offers like-for-like, then decides within the margin floor.", "SIMULATED DATA", "H")
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 70
    ex = d["pm"][d["pm"]["case_type_truth"] == "non_refundable"].iloc[0]
    inputs = [
        ("ETG selling price (EUR)", float(ex["etg_price_eur"]), None, "0.00"),
        ("Competitor price reported (EUR)", float(ex["competitor_price_eur"]), None, "0.00"),
        ("Room nights", int(ex["room_nights"]), None, "0"),
        ("Same room type?", "Yes", "=Lists!$G$2:$G$3", "@"),
        ("Same dates?", "Yes", "=Lists!$G$2:$G$3", "@"),
        ("Screenshot provided?", "Yes", "=Lists!$G$2:$G$3", "@"),
        ("Screenshot age (hours)", int(ex["screenshot_age_hours"]), None, "0"),
        ("Our meal plan", ex["etg_meal"], "=Lists!$I$2:$I$3", "@"),
        ("Competitor meal plan", ex["competitor_meal"], "=Lists!$I$2:$I$3", "@"),
        ("Our cancellation policy", ex["etg_cancellation"], "=Lists!$J$2:$J$3", "@"),
        ("Competitor cancellation policy", ex["competitor_cancellation"], "=Lists!$J$2:$J$3", "@"),
        ("Competitor price includes taxes?", ex["competitor_taxes_included"], "=Lists!$G$2:$G$3", "@"),
        ("Our net cost from the supplier (EUR)", float(ex["net_cost_eur"]), None, "0.00"),
        ("Cheapest alternative supplier net cost (EUR, blank if none)",
         None if pd.isna(ex["alt_supplier_net_eur"]) else float(ex["alt_supplier_net_eur"]), None, "0.00"),
    ]
    ws.Range("B4").Value = "1. The request"
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Font.Size = 12
    for i, (lab, val, src, fmt) in enumerate(inputs):
        r = 5 + i
        ws.Range(f"B{r}").Value = lab
        nfmt(ws.Range(f"C{r}"), fmt)
        ws.Range(f"C{r}").Value = val
        input_cell(ws.Range(f"C{r}"))
        if src:
            dropdown(ws.Range(f"C{r}"), src)
    params = [("Margin floor (on selling price)", MARGIN_FLOOR, "0.0%"), ("Breakfast value per guest-night (EUR)", 12, "0.00"),
              ("Guests per room", 2, "0"), ("Taxes to add when excluded", 0.10, "0%"),
              ("Value of free cancellation vs non-refundable", 0.08, "0%"), ("Max screenshot age (hours)", 48, "0")]
    ws.Range("B20").Value = "2. Rules (parameters)"
    ws.Range("B20").Font.Bold = True
    ws.Range("B20").Font.Size = 12
    for i, (lab, val, fmt) in enumerate(params):
        r = 21 + i
        ws.Range(f"B{r}").Value = lab
        ws.Range(f"C{r}").Value = val
        nfmt(ws.Range(f"C{r}"), fmt)
        input_cell(ws.Range(f"C{r}"))
    calc = [
        ("3. Checks and decision", None, None, None),
        ("Evidence complete? (same room, same dates, screenshot, recent)",
         '=AND(C8="Yes",C9="Yes",C10="Yes",C11<=C26)', "General",
         "Without evidence we ask for details before any price change."),
        ("+ Taxes the competitor left out (EUR)", '=IF(C16="No",C24*C6,0)', "0.00",
         "A price without taxes is not comparable: add them back."),
        ("+ Value of our breakfast (EUR)", '=IF(AND(C12="Breakfast",C13="Room only"),C22*C23*MAX(C7,1),0)', "0.00",
         "Our rate includes breakfast, the competitor's does not."),
        ("+ Value of our free cancellation (EUR)", '=IF(AND(C14="Free cancellation",C15="Non-refundable"),C25*C5,0)',
         "0.00", "A non-refundable rate is cheaper because the guest takes the risk."),
        ("Competitor price, like-for-like (EUR)", "=C6+C30+C31+C32", "0.00", "What the competitor would charge for our product."),
        ("Real gap (EUR)", "=C5-C33", "0.00", "Positive = we are really more expensive."),
        ("Lowest price that keeps the margin floor (EUR)", "=C17/(1-C21)", "0.00", "Net cost / (1 - floor)."),
        ("Maximum discount allowed (EUR)", "=MAX(C5-C35,0)", "0.00", "Selling price - floor price."),
        ("Price at the floor with the alternative supplier (EUR)", '=IF(C18="","",C18/(1-C21))', "0.00",
         "Can another supplier let us match?"),
        ("DECISION", '=IF(NOT(C29),"Ask for details",IF(C34<=0,"Refuse",IF(C34<=C36,"Discount",'
                     'IF(AND(C37<>"",N(C37)<=C33),"Switch supplier","Refuse"))))', "General", ""),
        ("Discount to grant (EUR)", '=IF(C38="Discount",ROUND(C34,2),0)', "0.00", "We match the like-for-like price, not more."),
        ("Margin rate after decision", '=IFERROR((C5-C39-IF(C38="Switch supplier",C18,C17))/(C5-C39),"")', "0.0%",
         "Must stay above the floor."),
    ]
    for i, (lab, f, fmt, note) in enumerate(calc):
        r = 28 + i
        ws.Range(f"B{r}").Value = lab
        if f is None:
            ws.Range(f"B{r}").Font.Bold = True
            ws.Range(f"B{r}").Font.Size = 12
            continue
        ws.Range(f"C{r}").Formula2 = f
        if fmt != "General":
            nfmt(ws.Range(f"C{r}"), fmt)
        ws.Range(f"D{r}").Value = note
        ws.Range(f"D{r}").Font.Color = C["grey"]
    card(ws.Range("B29:D40"))
    ws.Range("B38:C38").Font.Bold = True
    ws.Range("B38:C38").Font.Size = 12
    ws.Range("C38").Font.Color = C["blue_dark"]
    ws.Range("B42").Value = "4. Reply to the partner (template, filled automatically)"
    ws.Range("B42").Font.Bold = True
    ws.Range("B42").Font.Size = 12
    reply = ('=SWITCH(C38,'
             '"Discount","Thank you for flagging this. We checked the offer: same room, dates and conditions. '
             'We have adjusted our price by EUR "&FIXED(C39,2)&" to EUR "&FIXED(C5-C39,2)&". The new rate is valid for 24 hours.",'
             '"Switch supplier","Thank you for flagging this. We found the same room through another of our suppliers '
             'and can now offer EUR "&FIXED(C33,2)&" with the same conditions. The new rate is valid for 24 hours.",'
             '"Ask for details","Thank you for your message. To review the price, please send a screenshot taken in the '
             'last 48 hours showing the hotel, room type, dates, meal plan, cancellation policy and the total with taxes.",'
             '"Refuse",IF(C34>0,"Thank you for flagging this. We checked the offer carefully. The difference comes from '
             'the supplier cost for these dates, and we cannot go lower without selling at a loss. Our rate includes '
             '"&LOWER(C12)&" and "&LOWER(C14)&". We are happy to look at alternative hotels nearby at a better price.",'
             '"Thank you for flagging this. We checked the offer: it is not the same product. "'
             '&IF(C30>0,"The competitor price excludes taxes (about EUR "&FIXED(C30,2)&"). ","")'
             '&IF(C31>0,"Our rate includes breakfast (worth about EUR "&FIXED(C31,2)&"). ","")'
             '&IF(C32>0,"Our rate can be cancelled free of charge, theirs is non-refundable (worth about EUR "&FIXED(C32,2)&"). ","")'
             '&"Compared like-for-like, our price of EUR "&FIXED(C5,2)&" is the better offer, so we keep it."))')
    ws.Range("B43").Formula2 = reply
    ws.Range("B43:D47").Merge()
    ws.Range("B43").WrapText = True
    ws.Range("B43").VerticalAlignment = -4160
    card(ws.Range("B43:D47"))
    footer(ws, 49, "Same rules as src/simulate.py decide() and playbook/price_match_rules.md")

    # ---------------- log of requests + KPIs
    ws = bk.sheet("PriceMatch_Log", "green")
    title(ws, "Price-match requests: log and KPIs", "1,347 simulated requests. Column 'Rule check' re-applies the "
          "calculator rules to every row.", "SIMULATED DATA", "P")
    cols = ["request_id", "request_date", "partner_id", "partner_type", "booking_id", "destination_country",
            "competitor", "etg_price_eur", "competitor_price_eur", "room_nights", "same_room_type", "same_dates",
            "screenshot", "screenshot_age_hours", "etg_meal", "competitor_meal", "etg_cancellation",
            "competitor_cancellation", "competitor_taxes_included", "net_cost_eur", "alt_supplier_net_eur",
            "decision", "discount_eur", "handling_hours", "booked"]
    pm = d["pm"][cols].copy()
    pm["request_date"] = pd.to_datetime(pm["request_date"])
    top = 22
    rng = write_frame(ws, f"B{top}", pm, "tblPM")
    lo = ws.ListObjects("tblPM")
    lo.ListColumns.Add().Name = "rule_check"
    rule = ('=LET(p,[@[etg_price_eur]],cp,[@[competitor_price_eur]],n,MAX([@[room_nights]],1),'
            'ev,AND([@screenshot]="Yes",[@[screenshot_age_hours]]<=48,[@[same_dates]]="Yes"),'
            'adj,cp+IF([@[competitor_taxes_included]]="No",0.1*cp,0)'
            '+IF(AND([@[etg_meal]]="Breakfast",[@[competitor_meal]]="Room only"),24*n,0)'
            '+IF(AND([@[etg_cancellation]]="Free cancellation",[@[competitor_cancellation]]="Non-refundable"),0.08*p,0),'
            'gap,p-adj,mx,MAX(p-[@[net_cost_eur]]/(1-0.04),0),alt,[@[alt_supplier_net_eur]],'
            'IF(NOT(ev),"Ask for details",IF(gap<=0,"Refuse",IF(gap<=mx,"Discount",'
            'IF(AND(alt<>"",N(alt)/(1-0.04)<=adj),"Switch supplier","Refuse")))))')
    lo.ListColumns("rule_check").DataBodyRange.Formula2 = rule
    nfmt(ws.Range(f"C{top + 1}:C{top + len(pm)}"), "yyyy-mm-dd")
    kp = [("Requests", "=ROWS(tblPM[request_id])", "#,##0"),
          ("Requests checked with complete evidence", '=COUNTIFS(tblPM[decision],"<>Ask for details")', "#,##0"),
          ("Discount granted", '=COUNTIFS(tblPM[decision],"Discount")/ROWS(tblPM[decision])', "0%"),
          ("Switched to a cheaper supplier", '=COUNTIFS(tblPM[decision],"Switch supplier")/ROWS(tblPM[decision])', "0%"),
          ("Refused with arguments", '=COUNTIFS(tblPM[decision],"Refuse")/ROWS(tblPM[decision])', "0%"),
          ("Asked for details", '=COUNTIFS(tblPM[decision],"Ask for details")/ROWS(tblPM[decision])', "0%"),
          ("Cost of discounts on booked requests (EUR)", '=SUMIFS(tblPM[discount_eur],tblPM[booked],"Yes")', "#,##0"),
          ("Sales saved (discount or switch, then booked) (EUR)",
           '=SUMPRODUCT((tblPM[decision]="Discount")+(tblPM[decision]="Switch supplier"),--(tblPM[booked]="Yes"),'
           'tblPM[etg_price_eur]-tblPM[discount_eur])', "#,##0"),
          ("Median handling time (hours)", "=MEDIAN(tblPM[handling_hours])", "0.0"),
          ("Rule check: rows where the formula gives the logged decision",
           '=SUMPRODUCT(--(tblPM[rule_check]=tblPM[decision]))&" / "&ROWS(tblPM[decision])', "@")]
    for i, (lab, f, fmt) in enumerate(kp):
        r = 4 + i
        ws.Range(f"B{r}").Value = lab
        ws.Range(f"F{r}").Formula2 = f
        nfmt(ws.Range(f"F{r}"), fmt)
        ws.Range(f"F{r}").Font.Bold = True
    card(ws.Range("B4:F13"))
    area(ws.Range("H4"), 1, 4).Value = [["Decision", "Requests", "Booked rate", "Avg handling (h)"]]
    header(ws.Range("H4:K4"))
    for i, dec in enumerate(["Discount", "Switch supplier", "Refuse", "Ask for details"]):
        r = 5 + i
        ws.Range(f"H{r}").Value = dec
        ws.Range(f"I{r}").Formula2 = f'=COUNTIFS(tblPM[decision],H{r})'
        ws.Range(f"J{r}").Formula2 = f'=COUNTIFS(tblPM[decision],H{r},tblPM[booked],"Yes")/I{r}'
        ws.Range(f"K{r}").Formula2 = f'=AVERAGEIFS(tblPM[handling_hours],tblPM[decision],H{r})'
    nfmt(ws.Range("J5:J8"), "0%")
    nfmt(ws.Range("K5:K8"), "0.0")
    card(ws.Range("H4:K8"))
    ch = ws.ChartObjects().Add(ws.Range("M3").Left, ws.Range("M3").Top, 380, 230).Chart
    ch.ChartType = 57   # bar clustered
    ch.SetSourceData(ws.Range("H4:I8"))
    ch.HasTitle = True
    ch.ChartTitle.Text = "Decisions on 1,347 requests"
    style_chart(ch, [C["blue"]])
    ch.HasLegend = False
    footer(ws, 19, "")
    ws.Columns("B:Z").AutoFit()
    ws.Columns("B").ColumnWidth = 50


def build_pricing(bk: Book, d: dict) -> None:
    ws = bk.sheet("Pricing", "blue")
    title(ws, "Pricing recommendations within the current model", "B2B hotel segments, last 90 days. Rule: hold near "
          "the floor, -0.5 pt where valid price-match pressure is high and growth weak, +0.5 pt where pressure is "
          "low and growth strong.", "SIMULATED DATA", "R")
    ws.Range("B4").Value = "Price elasticity, prudent case"
    ws.Range("C4").Value = -1.0
    ws.Range("B5").Value = "Price elasticity, severe case"
    ws.Range("C5").Value = -3.0
    ws.Range("B6").Value = "Margin floor"
    ws.Range("C6").Value = MARGIN_FLOOR
    nfmt(ws.Range("C6"), "0.0%")
    input_cell(ws.Range("C4:C6"))
    r = d["reco"][["partner_type", "destination_country", "bookings", "sales_eur", "cost_eur", "margin_eur",
                   "yoy_volume", "markup", "margin_rate", "valid_requests_per_1000", "markup_change",
                   "recommendation"]].copy()
    top = 9
    write_frame(ws, f"B{top}", r, "tblPricing")
    lo = ws.ListObjects("tblPricing")
    for name, f in [
        ("price_change_pct", "=[@[markup_change]]/(1+[@markup])"),
        ("margin_effect_prudent", "=LET(dp,[@[price_change_pct]],dv,$C$4*dp,"
                                  "[@[sales_eur]]*(1+dp)*(1+dv)-[@[cost_eur]]*(1+dv)-[@[margin_eur]])"),
        ("margin_effect_severe", "=LET(dp,[@[price_change_pct]],dv,$C$5*dp,"
                                 "[@[sales_eur]]*(1+dp)*(1+dv)-[@[cost_eur]]*(1+dv)-[@[margin_eur]])"),
        ("below_floor_flag", "=[@[margin_rate]]<$C$6+0.01"),
    ]:
        lo.ListColumns.Add().Name = name
        lo.ListColumns(name).DataBodyRange.Formula2 = f
    n = len(r)
    nfmt(ws.Range(f"E{top + 1}:G{top + n}"), "#,##0")
    nfmt(ws.Range(f"H{top + 1}:J{top + n}"), "0.0%")
    nfmt(ws.Range(f"K{top + 1}:K{top + n}"), "0.0")
    nfmt(ws.Range(f"L{top + 1}:L{top + n}"), "+0.0%;-0.0%;0.0%")
    nfmt(ws.Range(f"N{top + 1}:N{top + n}"), "+0.00%;-0.00%;0.00%")
    nfmt(ws.Range(f"O{top + 1}:P{top + n}"), "+#,##0;-#,##0;0")
    red_if(ws.Range(f"B{top + 1}:J{top + n}"), f"$Q{top + 1}", font_only=True)
    ws.Range("E4").Value = "Expected 90-day margin effect (EUR), prudent / severe elasticity"
    ws.Range("E4").Font.Bold = True
    ws.Range("E5").Formula2 = "=SUM(tblPricing[margin_effect_prudent])"
    ws.Range("F5").Formula2 = "=SUM(tblPricing[margin_effect_severe])"
    nfmt(ws.Range("E5:F5"), "+#,##0;-#,##0")
    ws.Range("E5:F5").Font.Size = 14
    ws.Range("E5:F5").Font.Bold = True
    ws.Columns("B:Q").AutoFit()
    ws.Columns("M").ColumnWidth = 60
    footer(ws, top + n + 2, "Elasticities are assumptions (range), to be validated with a price test")


def build_dq(bk: Book, d: dict) -> None:
    ws = bk.sheet("Data_Quality", "red")
    title(ws, "Data quality first", "Checks run on the raw ledger before any conclusion (src/analysis.py, "
          "sql/06_data_quality_checks.sql).", "SIMULATED DATA", "J")
    write_frame(ws, "B4", d["dq"], "tblDQ")
    rc = d["raw_clean"]
    rc = rc[rc["booking_date"].between("2026-07-14", "2026-07-28") | rc["booking_date"].between("2026-08-01", "2026-08-12")
            | rc["booking_date"].between("2026-10-09", "2026-10-21")].copy()
    top = 6 + len(d["dq"]) + 2
    ws.Range(f"B{top - 1}").Value = "Daily totals, raw vs clean, around the three issues"
    ws.Range(f"B{top - 1}").Font.Bold = True
    write_frame(ws, f"B{top}", rc, "tblRawClean")
    lo = ws.ListObjects("tblRawClean")
    lo.ListColumns.Add().Name = "sales_gap_pct"
    lo.ListColumns("sales_gap_pct").DataBodyRange.Formula2 = "=[@[sales_eur_raw]]/[@[sales_eur_clean]]-1"
    lo.ListColumns.Add().Name = "flag"
    lo.ListColumns("flag").DataBodyRange.Formula2 = ("=OR(ABS([@[sales_gap_pct]])>0.02,"
                                                     "ABS([@[margin_rate_raw]]-[@[margin_rate_clean]])>0.005)")
    n = len(rc)
    nfmt(ws.Range(f"B{top + 1}:B{top + n}"), "yyyy-mm-dd")
    nfmt(ws.Range(f"C{top + 1}:H{top + n}"), "#,##0")
    for col in ("E", "H"):
        nfmt(ws.Range(f"{col}{top + 1}:{col}{top + n}"), "0.0%")
    nfmt(ws.Range(f"I{top + 1}:I{top + n}"), "+0.0%;-0.0%;0.0%")
    red_if(ws.Range(f"B{top + 1}:I{top + n}"), f"$J{top + 1}")
    ws.Columns("B:J").AutoFit()
    ws.Columns("G").ColumnWidth = 60
    ws.Columns("H").ColumnWidth = 60
    footer(ws, 2 + 1, "")


def build_dims(bk: Book) -> None:
    ws = bk.sheet("Partners_Suppliers", "grey")
    title(ws, "Partners and suppliers (fictional codes)", "Dimension tables, also loaded in the data model.",
          "SIMULATED DATA", "J")
    p = pd.read_csv(SIM / "partners.csv")
    s = pd.read_csv(SIM / "suppliers.csv")
    write_frame(ws, "B4", s, "tblSuppliers")
    write_frame(ws, "J4", p, "tblPartners")
    ws.Columns("B:P").AutoFit()


def build_readme(bk: Book, d: dict) -> None:
    ws = bk.sheet("README", "ink")
    bk.wb.Worksheets("README").Move(Before=bk.wb.Worksheets(1))
    title(ws, "ETG commercial & pricing cockpit (case study)", "A daily monitoring and pricing workbook for a multi-brand "
          "B2B travel platform. Independent case study - not affiliated with or endorsed by Emerging Travel Group.",
          "SIMULATED DATA", "F")
    rows = [
        ("What is real, what is simulated", ""),
        ("SIMULATED", "All bookings, partners, suppliers, prices, margins and price-match requests (fixed seed 42). "
                      "They are NOT Emerging Travel Group data; names are fictional codes."),
        ("PUBLIC DATA", "FX rates in the FX table are real ECB reference rates. Demand seasonality used by the "
                        "simulation comes from Eurostat. Sources: data/SOURCES.md."),
        ("", ""),
        ("Sheet", "What it does"),
        ("Cockpit", "Daily KPIs vs day before, same day last year, previous 7-day average and an expected value "
                    "(same weekday, 4 weeks). By brand, line of business, partner type, supplier type, source market, "
                    "destination, top 10 partners. Built only with CUBEMEMBER / CUBEVALUE / CUBESET / "
                    "CUBERANKEDMEMBER on the data model. Change the yellow cells."),
        ("Pivots", "PivotTables connected to the data model with slicers (line of business, market, partner type, "
                   "destination, partner tier, channel). Supplier margin by month. Real ECB FX by month."),
        ("Alerts", "Statistical alert band with SUMIFS / LET / INDEX / STDEV.S on tblDaily. Choose dimension, member "
                   "and metric. Same rule as the Python and SQL code. Alert days listed with FILTER."),
        ("Bridge", "Margin change between two months split into volume, mix, price, FX and cost, with dynamic "
                   "arrays (UNIQUE, SORT, SUMIFS on spilled ranges) and a waterfall chart. Check cell must be 0."),
        ("PriceMatch_Calc", "Calculator for a partner's price-match request: evidence, like-for-like adjustments "
                            "(taxes, breakfast, cancellation), margin floor, maximum discount, decision, reply text."),
        ("PriceMatch_Log", "1,347 requests with KPIs (COUNTIFS, SUMIFS, SUMPRODUCT, MEDIAN). A LET formula re-applies "
                           "the rules to every row: 'Rule check' shows how many rows agree."),
        ("Pricing", "Markup recommendations by B2B segment with the expected margin effect for a range of price "
                    "elasticities (change the yellow cells)."),
        ("Data_Quality", "Issues found in the raw ledger (duplicates, missing costs, time-zone shift) and their effect "
                         "on daily totals."),
        ("Data_Daily / Data_Monthly", "Aggregated tables (tblDaily, tblMonthly) used by the formula sheets."),
        ("Partners_Suppliers", "Dimension tables (fictional codes)."),
        ("Lists", "Values for the drop-down lists."),
        ("", ""),
        ("Data model (Power Pivot)", "Tables: Bookings (clean ledger, ~192k rows), Calendar, Suppliers, Partners, "
                                     "FX. Relationships: Bookings[booking_date] -> Calendar[date], [supplier_id] -> "
                                     "Suppliers, [partner_id] -> Partners, [fx_key] -> FX. About 50 DAX measures "
                                     "(Data > Manage Data Model > Measures)."),
        ("Definitions", "Sales = selling price after discount, EUR, on the booking date. Margin = sales - supplier net "
                        "cost. Margin rate = margin / sales. Cancellations are counted on the cancellation date. "
                        "Expected = average of the same weekday over the previous 4 weeks."),
        ("Refresh", "The data is stored in the file: no refresh needed. To reload the CSV files, open Data > "
                    "Queries & Connections and update the folder path in each query, then Refresh All."),
    ]
    for i, (a, b) in enumerate(rows):
        r = 4 + i
        ws.Range(f"B{r}").Value = a
        ws.Range(f"C{r}").Value = b
        ws.Range(f"C{r}").WrapText = True
        if a in ("What is real, what is simulated", "Sheet", "Data model (Power Pivot)"):
            ws.Range(f"B{r}:C{r}").Font.Bold = True
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 120
    ws.Range("B4:C24").VerticalAlignment = -4160
    footer(ws, 26, "")


# =========================================================================== checks & save
def wait_cube(bk: Book) -> None:
    bk.xl.Calculation = -4105          # automatic again
    bk.xl.CalculateFullRebuild()
    for _ in range(60):
        bk.xl.CalculateUntilAsyncQueriesDone()
        time.sleep(0.5)
        if bk.wb.Worksheets("Cockpit").Range("C11").Text not in ("", "#GETTING_DATA", "#N/A"):
            break


def verify(bk: Book) -> list[str]:
    """Scan used ranges for formula errors and run a few cross-checks with the Python outputs."""
    msgs = []
    for ws in bk.wb.Worksheets:
        try:
            errs = ws.UsedRange.SpecialCells(-4123, 16)   # formulas returning errors
            msgs.append(f"{ws.Name}: {errs.Count} formula cells in error, e.g. {errs.Cells(1, 1).Address}")
        except Exception:
            pass
    br = bk.wb.Worksheets("Bridge")
    py = pd.read_csv(OUT_SIM / "margin_bridge_sep.csv")
    xl = [br.Range(f"I{r}").Value for r in range(5, 12)]
    diff = max(abs(a - b) for a, b in zip(xl, py["eur"]))
    msgs.append(f"Bridge Excel vs Python max diff: {diff:.6f} EUR; check cell = {br.Range('I12').Value}")
    al = bk.wb.Worksheets("Alerts")
    msgs.append(f"Alerts AE bookings: {al.Range('I5').Value} alerts; first = {al.Range('I6').Text}")
    pm = bk.wb.Worksheets("PriceMatch_Log")
    msgs.append(f"Price-match rule check: {pm.Range('F13').Text}")
    ck = bk.wb.Worksheets("Cockpit")
    msgs.append("Cockpit " + ", ".join(f"{ck.Range(f'B{r}').Text}={ck.Range(f'C{r}').Text}" for r in range(10, 17)))
    return msgs


def snapshot(bk: Book, sheet: str, address: str, name: str) -> None:
    """Export a picture of a sheet range (cells + charts) to docs/img/excel_<name>.png for the report."""
    out = ROOT / "docs" / "img" / f"excel_{name}.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    ws = bk.wb.Worksheets(sheet)
    ws.Activate()
    rng = ws.Range(address)
    for attempt in range(3):
        try:
            rng.CopyPicture(1, 2)                       # as shown on screen, bitmap
            co = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
            co.Activate()
            co.Chart.ChartArea.Format.Line.Visible = False
            co.Chart.Paste()
            co.Chart.Export(str(out), "PNG")
            co.Delete()
            return
        except Exception:
            time.sleep(1.5)
    print(f"snapshot failed: {sheet}!{address}")


def main() -> None:
    d = prepare()
    XLSX.parent.mkdir(parents=True, exist_ok=True)
    bk = Book()
    try:
        bk.sheet("Lists_tmp")
        build_model(bk)
        build_dims(bk)
        build_data_sheets(bk, d)
        build_lists(bk, d)
        bk.wb.Worksheets("Lists_tmp").Delete()
        build_pivots(bk)               # PivotTables first, while calculation is automatic
        bk.xl.Calculation = -4135      # then manual: CUBE formulas would keep Excel busy
        build_cockpit(bk, d)
        build_alerts(bk, d)
        build_bridge(bk, d)
        build_price_match(bk, d)
        build_pricing(bk, d)
        build_dq(bk, d)
        build_readme(bk, d)
        order = ["README", "Cockpit", "Pivots", "Alerts", "Bridge", "PriceMatch_Calc", "PriceMatch_Log", "Pricing",
                 "Data_Quality", "Partners_Suppliers", "Data_Daily", "Data_Monthly", "Lists"]
        for name in reversed(order):
            bk.wb.Worksheets(name).Move(Before=bk.wb.Worksheets(1))
        wait_cube(bk)
        pv = bk.wb.Worksheets("Pivots")
        pv.Columns("C:P").ColumnWidth = 11
        pv.Columns("B").ColumnWidth = 22
        for m in verify(bk):
            print(m)
        bk.xl.ScreenUpdating = True
        bk.xl.Visible = True              # pictures need the sheets to be rendered
        bk.xl.WindowState = -4140         # minimised
        snapshot(bk, "Cockpit", "B1:L41", "cockpit")
        snapshot(bk, "Alerts", "B4:N34", "alerts")
        snapshot(bk, "Bridge", "H3:W22", "bridge")
        snapshot(bk, "Pivots", "B1:V30", "pivots")
        snapshot(bk, "PriceMatch_Calc", "B1:D48", "pricematch")
        bk.xl.Visible = False
        # personal information: cleared on save
        for prop in ("Author", "Last author", "Company", "Manager"):
            try:
                bk.wb.BuiltinDocumentProperties(prop).Value = ""
            except Exception:
                pass
        bk.wb.RemovePersonalInformation = True
        bk.wb.Worksheets("README").Activate()
        bk.wb.Worksheets("README").Range("A1").Select()
        bk.xl.ScreenUpdating = True
        if XLSX.exists():
            XLSX.unlink()
        bk.wb.SaveAs(str(XLSX), 51)
    finally:
        bk.wb.Close(False)
        bk.xl.Quit()
    time.sleep(3)
    import shutil
    shutil.rmtree(TMP, ignore_errors=True)
    print("saved", XLSX)


if __name__ == "__main__":
    main()
