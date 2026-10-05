"""HFG frame standard: the layout every model sheet follows.

Layout (see docs/frame-standard.md):
  Column A        navigation links (back to the table of contents, to the checks)
  Columns B to G  labels, indented by column for sub-items
  Column H        units or a single driver value
  Column I        row total, closing value or check flag
  Column J on     monthly timeline (one column per month)
  Rows 1 to 3     sheet title, model name, entity name
  Rows 5 to 15    timeline block (time series sheets only)
  Row 17 on       content, in sections
"""

from __future__ import annotations

from .xlsx.cells import col_letter
from .xlsx.sheet import Sheet
from .xlsx.styles import Derived

LABEL_COL = 2  # B
UNIT_COL = 8  # H
TOTAL_COL = 9  # I
FIRST_TS_COL = 10  # J
CONTENT_START_ROW = 17

# Cell formats taken from the Budget Template. Each index was checked against
# the template's named styles (Heading 1., Number., Assumption Number. and so on).
TEMPLATE_STYLES = {
    "nav": 69,              # Hyperlink, Wingdings arrow and cross
    "title": 124,           # bold 11
    "model_name": 123,      # bold 10
    "entity": 106,          # Segoe UI Semibold 9
    "tl_top_label": 130,    # Period Title.
    "tl_top_value": 131,
    "tl_2_label": 132,
    "tl_2_value": 133,
    "tl_label": 134,        # Heading 4.
    "tl_blank": 271,
    "tl_date": 135,         # Date. d-mmm-yy
    "tl_num": 136,          # Number. #,##0
    "tl_year": 137,         # Year.
    "tl_key": 138,          # Heading 4. right aligned
    "tl_last_label": 139,
    "tl_last_blank": 272,
    "tl_last_key": 140,
    "section": 465,         # Heading 1. bar in theme accent 1
    "section_c": 466,
    "sub": 151,             # Heading 2. grey band
    "sub_c": 426,
    "h3": 932,              # Heading 3. bold label
    "label": 134,           # Heading 4. label
    "unit": 138,            # Heading 4. right aligned
    "num": 136,             # Number. #,##0
    "num1": 144,            # Number. #,##0.0
    "num_dash": 152,        # Number. with dashed bottom border
    "total1": 158,          # Number. bold, top border, #,##0.0
    "pct": 165,             # Percentage. 0.0%
    "date": 154,            # Date. mmm-yy
    "date_d": 135,          # Date. d-mmm-yy
    "in_num": 145,          # Assumption Number. #,##0
    "in_num1": 174,         # Assumption Number. #,##0.0
    "in_pct": 163,          # Assumption Percentage. 0.0%
    "in_pct2": 150,         # Assumption Percentage. 0.00%
    "in_text": 933,         # Assumption Heading. text input
    "in_year": 949,         # Assumption Year.
    "link_text": 895,       # Hyperlink Text.
}

DERIVED_STYLES = {
    "total": Derived(base="total1", num_fmt='_(#,##0_);\\(#,##0\\);_("-"_);_)@_)'),
    "in_date": Derived(base="in_num", num_fmt='_)d\\-mmm\\-yy_);_)d\\-mmm\\-yy_);_)"-"_);_)@_)', h_align="right"),
    "in_dec": Derived(base="in_num", num_fmt='_(#,##0.00_);\\(#,##0.00\\);_("-"_);_)@_)'),
    "dec": Derived(base="num", num_fmt='_(#,##0.00_);\\(#,##0.00\\);_("-"_);_)@_)'),
    "pct_total": Derived(base="pct", bold=True, border_top="thin"),
    "check": Derived(base="num", font_rgb="FFC00000"),
    "check_flag": Derived(base="num", font_rgb="FFC00000", bold=True, h_align="center"),
    "label_bold": Derived(base="h3"),
    "text": Derived(base="label"),
    "text_c": Derived(base="label", h_align="center"),
    "text_r": Derived(base="label", h_align="right"),
    "head_r": Derived(base="h3", h_align="right"),
    "multiple": Derived(base="num", num_fmt='_(#,##0.00"x"_);\\(#,##0.00"x"\\);_("-"_);_)@_)'),
}

TIMELINE_ROWS = {
    "month_ending": 5, "month": 6, "start": 7, "end": 8, "counter": 9, "fin_year": 10,
    "active": 11, "month_no": 12, "month_key": 13, "qtr_key": 14, "half_key": 15,
}


def ts_col(t: int) -> int:
    """Column index for period t (1 based)."""
    return FIRST_TS_COL + t - 1


def ts_letter(t: int) -> str:
    return col_letter(ts_col(t))


def write_header(sh: Sheet, st, title: str, n_periods: int | None, home: str = "HL_Home",
                 checks: str = "HL_Err_Chk"):
    """Rows 1 to 3, plus the timeline block when n_periods is given."""
    sh.set(1, 1, "±", style=st["nav"])
    sh.set(2, 1, "x", style=st["nav"])
    sh.hyperlinks.append(("A1", home, "±"))
    sh.hyperlinks.append(("A2", checks, "x"))
    sh.set(1, LABEL_COL, title, style=st["title"])
    sh.set(2, LABEL_COL, formula="Model_Name", style=st["model_name"])
    sh.set(3, LABEL_COL, formula="Cover!B1", style=st["entity"])
    sh.cols.append((1, 1, 3.5, None))
    sh.cols.append((2, 6, 2.5, None))
    sh.cols.append((7, 7, 34.0, None))
    if not n_periods:
        return  # non-timeline sheets set their own widths from column H
    sh.cols.append((8, 8, 9.0, None))
    sh.cols.append((9, 9, 12.5, None))
    last = ts_col(n_periods)
    sh.cols.append((FIRST_TS_COL, last, 11.625, None))
    labels = [
        (5, "Month Ending", "tl_top_label", "tl_top_label"),
        (6, "Month", "tl_2_label", "tl_2_label"),
        (7, "Period Start Date", "tl_label", "tl_blank"),
        (8, "Period End Date", "tl_label", "tl_blank"),
        (9, "Counter", "tl_label", "tl_blank"),
        (10, "Financial Year", "tl_label", "tl_blank"),
        (11, "Active Column Number", "tl_label", "tl_blank"),
        (12, "Month Number", "tl_label", "tl_blank"),
        (13, "Month Key", "tl_label", "tl_blank"),
        (14, "Quarter Key", "tl_label", "tl_blank"),
        (15, "Half Key", "tl_last_label", "tl_last_blank"),
    ]
    for r, text, lab, blank in labels:
        sh.set(r, LABEL_COL, text, style=st[lab])
        for c in range(LABEL_COL + 1, TOTAL_COL + 1):
            sh.set(r, c, style=st[blank])
        if r >= 7:
            sh.row_props(r, outline=1)
    for t in range(1, n_periods + 1):
        c = ts_col(t)
        L = col_letter(c)
        sh.set(5, c, formula=f'TEXT({L}8,"mmm-yy")', style=st["tl_top_value"])
        sh.set(6, c, formula=f'"M"&{L}12&IF({L}9>DD_Ts_Last_Hist_Mth,Ts_Fcast_Per_ID,Ts_Hist_Per_ID)', style=st["tl_2_value"])
        sh.set(7, c, formula=f"EDATE(Ts_Start_Date,{L}9-1)", style=st["tl_date"])
        sh.set(8, c, formula=f"EDATE(Ts_Start_Date,{L}9)-1", style=st["tl_date"])
        sh.set(9, c, formula=f"COLUMNS($J9:{L}9)", style=st["tl_num"])
        sh.set(10, c, formula=f"YEAR({L}8)+IF(MONTH({L}8)>DD_Ts_Fin_Yr_End_Mth,1,0)", style=st["tl_year"])
        sh.set(11, c, formula=f"IF(AND({L}9>DD_Ts_Last_Hist_Mth,{L}9<=Ts_Term),{L}9-DD_Ts_Last_Hist_Mth,0)", style=st["tl_num"])
        sh.set(12, c, formula=f"MOD(MONTH({L}8)-DD_Ts_Fin_Yr_End_Mth-1,Ts_Mths_In_Yr)+1", style=st["tl_num"])
        sh.set(13, c, formula=f'"M"&{L}12&"-"&{L}10', style=st["tl_key"])
        sh.set(14, c, formula=f'"Q"&INT(({L}12-1)/Ts_Mths_In_Qtr)+1&"-"&{L}10', style=st["tl_key"])
        sh.set(15, c, formula=f'"H"&INT(({L}12-1)/Ts_Mths_In_Half)+1&"-"&{L}10', style=st["tl_last_key"])
    sh.freeze = (16, FIRST_TS_COL)
