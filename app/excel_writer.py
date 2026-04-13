"""
excel_writer.py
---------------
All openpyxl logic for writing AI review results into the M&V Review Sheet Template.
The AI never sees or touches this file — Python handles everything.

Sheet  : "1. M&V Report Compliance Check"
Columns written (1-based):
  H  col 8  = Active Status
  I  col 9  = Consultant Comments

Colour coding:
  Approved     -> bg #C6EFCE  font #375623  bold  centered
  Not Approved -> bg #FFC7CE  font #9C0006  bold  centered
  Incomplete   -> bg #FFEB9C  font #9C6500  bold  centered

  Comment non-empty -> bg #FFFF99  font #000000  wrap  top-aligned
  Comment empty     -> bg #FFFFFF
  Row height        -> max(50, (len//80 + newlines + 1) * 15)

  All written cells: Calibri 11, thin border #BFBFBF.
"""

import io

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Column indices (1-based) ─────────────────────────────────────────────────
COL_SN      = 2   # B
COL_STATUS  = 8   # H  Active Status
COL_COMMENT = 9   # I  Consultant Comments
DATA_START  = 24  # First data row to scan

FONT_NAME    = "Calibri"
FONT_SIZE    = 11
BORDER_COLOR = "BFBFBF"

# ── Style maps ───────────────────────────────────────────────────────────────
STATUS_STYLES = {
    "Approved":     {"bg": "C6EFCE", "fc": "375623"},
    "Not Approved": {"bg": "FFC7CE", "fc": "9C0006"},
    "Incomplete":   {"bg": "FFEB9C", "fc": "9C6500"},
}

COMMENT_BG = "FFFF99"
COMMENT_FC = "000000"


def _make_border() -> Border:
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def _make_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)


def _style_cell(cell, value: str, bg: str, fc: str,
                bold: bool = False, wrap: bool = False,
                align: str = "left") -> None:
    cell.value = value
    cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color=fc, bold=bold)
    cell.fill = _make_fill(bg)
    vertical = "center" if align == "center" else "top"
    cell.alignment = Alignment(horizontal=align, vertical=vertical, wrap_text=wrap)
    cell.border = _make_border()


def _is_section_header(value) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    try:
        f = float(s)
        return f == int(f) and "." not in s
    except (ValueError, TypeError):
        return False


def write_review(template_bytes: bytes, review_by_sn: dict) -> bytes:
    """
    Load the Excel template from bytes, write review results, return as bytes.

    Parameters
    ----------
    template_bytes : bytes
        Raw bytes of the M&V Review Sheet Template.
    review_by_sn  : dict
        Dict keyed by SN string -> {"status": ..., "comment": ...}

    Returns
    -------
    bytes
        In-memory bytes of the filled workbook (xlsx).
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes), keep_links=False)

    # Drop external links and all defined names (named ranges) — the output
    # file doesn't need them and broken/external ones cause Excel repair warnings.
    wb._external_links.clear()
    for name in list(wb.defined_names):
        del wb.defined_names[name]

    for sheet in wb.worksheets:
        sheet._charts.clear()
        sheet._images.clear()
        if hasattr(sheet, "_drawing") and sheet._drawing is not None:
            sheet._drawing = None

    ws = wb["1. M&V Report Compliance Check"]

    for row_idx in range(DATA_START, ws.max_row + 1):
        sn_val = ws.cell(row=row_idx, column=COL_SN).value
        if _is_section_header(sn_val):
            continue

        sn = str(sn_val).strip()
        if sn not in review_by_sn:
            continue

        item    = review_by_sn[sn]
        status  = item.get("status", "")
        comment = item.get("comment", "") or ""

        # ── Active Status (col H) ────────────────────────────────────────────
        st_s = STATUS_STYLES.get(status, {"bg": "FFFFFF", "fc": "000000"})
        _style_cell(
            ws.cell(row=row_idx, column=COL_STATUS),
            status, bg=st_s["bg"], fc=st_s["fc"],
            bold=True, align="center"
        )

        # ── Consultant Comments (col I) ──────────────────────────────────────
        if comment:
            _style_cell(
                ws.cell(row=row_idx, column=COL_COMMENT),
                comment, bg=COMMENT_BG, fc=COMMENT_FC,
                wrap=True, align="left"
            )
            lines = len(comment) // 80 + comment.count("\n") + 1
            ws.row_dimensions[row_idx].height = max(50, lines * 15)
        else:
            _style_cell(
                ws.cell(row=row_idx, column=COL_COMMENT),
                "", bg="FFFFFF", fc=COMMENT_FC,
                wrap=True, align="left"
            )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
