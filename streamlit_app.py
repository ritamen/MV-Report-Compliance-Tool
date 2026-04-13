# streamlit_app.py  –  M&V Report vs Calculation Sheet Sanity Check Tool
import base64
import io
import json
import logging
import os
import re
import sys
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent / "app"
PROMPT_PATH   = BASE_DIR / "assets" / "MV_Report_Calc_Sanity_Check_Prompt.txt"
TEMPLATE_PATH = BASE_DIR / "assets" / "M_V_Review_Sheet_Template_260407_V1_1.xlsx"
LOGO_PATH     = BASE_DIR / "static" / "arklogo2.png"

load_dotenv(BASE_DIR.parent / ".env")

sys.path.insert(0, str(BASE_DIR))
from excel_writer import write_review
from sn_extractor import extract_expected_sns

# ── Load assets once ──────────────────────────────────────────────────────────
REVIEWER_PROMPT: str   = PROMPT_PATH.read_text(encoding="utf-8")
TEMPLATE_BYTES:  bytes = TEMPLATE_PATH.read_bytes()
EXPECTED_SNS           = extract_expected_sns(str(TEMPLATE_PATH))

# ── API constants ─────────────────────────────────────────────────────────────
MODEL           = "claude-sonnet-4-6"
MAX_TOKENS      = 4000
THINKING_TOKENS = 3000
TIMEOUT_SECS    = 120

VALID_STATUS = {"Approved", "Not Approved", "Incomplete"}

logging.basicConfig(level=logging.INFO)


# ── Backend helpers ────────────────────────────────────────────────────────────

def img_to_base64(path: str) -> str:
    return base64.b64encode(Path(path).read_bytes()).decode("utf-8")

def _encode_bytes(data: bytes) -> str:
    return base64.standard_b64encode(data).decode("utf-8")

def _strip_fences(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()

def _extract_json_text(response) -> str:
    for block in response.content:
        if block.type == "text":
            return block.text
    raise ValueError("No text content block found in API response.")

def _validate_items(items: list) -> list:
    errors = []
    for i, item in enumerate(items):
        if not isinstance(item, dict):
            errors.append(f"Item {i} is not a dict"); continue
        for field in ("sn", "status", "comment"):
            if field not in item:
                errors.append(f"Item {i} missing field '{field}'")
        if "status" in item and item["status"] not in VALID_STATUS:
            errors.append(f"Item {i} invalid status={item['status']!r}")
    if isinstance(items, list) and len(items) != 2:
        errors.append(f"Expected exactly 2 items, got {len(items)}")
    return errors

def _parse_and_validate(raw: str):
    cleaned = _strip_fences(raw)
    try:
        items = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        return [], [f"JSON parse error: {exc}"]
    if not isinstance(items, list):
        return [], ["Response is not a JSON array"]
    return items, _validate_items(items)

def _xlsx_to_text(xlsx_bytes: bytes) -> str:
    """Convert an xlsx workbook to a plain-text CSV representation."""
    import io
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    parts = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        parts.append(f"=== Sheet: {sheet_name} ===")
        for row in ws.iter_rows(values_only=True):
            parts.append(",".join("" if v is None else str(v) for v in row))
    return "\n".join(parts)

def _build_user_content(pdf_bytes: bytes, xlsx_bytes: bytes) -> list:
    calc_text = _xlsx_to_text(xlsx_bytes)
    return [
        {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": _encode_bytes(pdf_bytes),
            },
            "title": "M&V Report",
        },
        {
            "type": "text",
            "text": f"M&V Calculation Sheet (converted from Excel):\n\n{calc_text}",
        },
        {
            "type": "text",
            "text": (
                "Compare the M&V Report against the M&V Calculation Sheet using every "
                "question in your instructions. Return a single valid JSON array only — "
                "no markdown, no explanation, no text outside the array. "
                "Each element must have exactly three fields: sn, status, comment."
            ),
        },
    ]

def _call_claude(client, user_content: list) -> str:
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        temperature=1,
        thinking={"type": "enabled", "budget_tokens": THINKING_TOKENS},
        system=REVIEWER_PROMPT,
        messages=[{"role": "user", "content": user_content}],
        timeout=TIMEOUT_SECS,
    )
    return _extract_json_text(response)

def _call_claude_retry(client, user_content: list, first_raw: str) -> str:
    retry_messages = [
        {"role": "user",      "content": user_content},
        {"role": "assistant", "content": first_raw},
        {"role": "user",      "content": (
            "Your previous response was invalid. Return ONLY a valid JSON array with "
            "exactly two elements for SNs 0.1 and 0.2. Each item must have:\n"
            "- sn: '0.1' or '0.2'\n"
            "- status: exactly one of: Approved / Not Approved / Incomplete\n"
            "  (do NOT use Approved as Noted — this is Round 1)\n"
            "- comment: string (never blank — include confirmation or discrepancy detail)\n"
            "No markdown, no explanation, no text outside the JSON array."
        )},
    ]
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        temperature=1,
        thinking={"type": "enabled", "budget_tokens": THINKING_TOKENS},
        system=REVIEWER_PROMPT,
        messages=retry_messages,
        timeout=TIMEOUT_SECS,
    )
    return _extract_json_text(response)

def run_sanity_check(pdf_bytes: bytes, xlsx_bytes: bytes, pdf_filename: str):
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set.")

    client = anthropic.Anthropic(api_key=api_key)
    user_content = _build_user_content(pdf_bytes, xlsx_bytes)

    try:
        raw = _call_claude(client, user_content)
    except anthropic.APITimeoutError:
        raise RuntimeError(
            "The API call timed out after 120 seconds. Please re-run the check."
        )
    except anthropic.AuthenticationError:
        raise RuntimeError(
            "API authentication failed. Please check the ANTHROPIC_API_KEY in your .env file."
        )

    items, errors = _parse_and_validate(raw)

    if errors:
        logging.warning("First response invalid: %s — retrying.", errors)
        raw = _call_claude_retry(client, user_content, raw)
        items, errors = _parse_and_validate(raw)
        if errors:
            raise RuntimeError(
                "The AI returned an invalid response after retry. Please re-run the check."
            )

    review_by_sn = {str(item["sn"]).strip(): item for item in items}
    missing_sns  = sorted(set(EXPECTED_SNS) - set(review_by_sn.keys()))

    approved     = sum(1 for it in items if it.get("status") == "Approved")
    not_approved = sum(1 for it in items if it.get("status") == "Not Approved")
    incomplete   = sum(1 for it in items if it.get("status") == "Incomplete")
    total        = len(items)

    filled_bytes = write_review(TEMPLATE_BYTES, review_by_sn)

    base_name = pdf_filename.replace(".pdf", "").replace(".PDF", "")
    output_filename = f"MV_Report_Sanity_Check_{base_name}.xlsx"

    return {
        "total":        total,
        "approved":     approved,
        "not_approved": not_approved,
        "incomplete":   incomplete,
        "missing_sns":  missing_sns,
        "excel_bytes":  filled_bytes,
        "filename":     output_filename,
    }


# ============================================================
# UI
# ============================================================

st.set_page_config(page_title="ARK Energy | M&V Report Sanity Check", layout="wide")

st.markdown(
    """
    <style>
    :root{
      --ark-blue: #0D6079;
      --ark-orange: #F79428;
      --ark-black: #000000;
    }

    html, body, [class*="css"], * {
        font-family: "Trebuchet MS", Arial, sans-serif !important;
        color: var(--ark-black);
    }
    button, input, textarea, select, label, p, div, span, li {
        font-family: "Trebuchet MS", Arial, sans-serif !important;
    }

    header[data-testid="stHeader"] { display: none; }
    footer { display: none; }

    .block-container {
        padding-top: 6.2rem !important;
        padding-bottom: 1.2rem !important;
        max-width: 98vw !important;
    }

    .ark-nav {
        position: fixed;
        top: 0; left: 0; right: 0;
        z-index: 9999;
        background: linear-gradient(90deg,#060C2E 0%,#08133A 45%,#0B1A4A 100%);
        padding: 12px 18px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.35);
    }
    .ark-nav-inner{
        width: 98vw; margin: 0 auto; border-radius: 14px; padding: 10px 14px;
        display: flex; align-items: center; justify-content: space-between;
        background: linear-gradient(90deg,#060C2E 0%,#08133A 45%,#0B1A4A 100%);
    }
    .ark-nav-left { display:flex; align-items:center; gap:14px; }
    .ark-nav-title { color:white !important; font-size:22px !important; font-weight:900; line-height:1.2; margin:0; }
    .ark-nav-subtitle { color:rgba(255,255,255,0.65) !important; font-size:13px !important; font-weight:400; margin:0; }
    .ark-nav-right { display:flex; align-items:center; gap:10px; }
    .pill {
        border-radius:999px; padding:8px 14px; font-size:14px; font-weight:900;
        border:1px solid rgba(255,255,255,0.25); color:white !important; background:transparent; white-space:nowrap;
    }

    .ark-section { margin-top:10px; margin-bottom:6px; display:flex; align-items:baseline; gap:10px; }
    .ark-section-title { font-size:18px; font-weight:900; color:var(--ark-blue); margin:0; line-height:1; }
    .ark-section-rule { height:2px; background:rgba(13,96,121,0.25); width:100%; margin-top:8px; margin-bottom:14px; }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #ebebeb !important;
        border: 1px solid #c8c8c8 !important;
        border-radius: 6px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #e2e2e2 !important;
        border-color: #b0b0b0 !important;
    }

    [data-testid="stFileUploaderDropzone"] button [data-testid="stIconMaterial"] {
        display: none !important;
    }

    label { font-size:15px !important; font-weight:700 !important; }

    div.stButton > button[kind="primary"],
    div.stButton > button[kind="primary"] * { color:#FFFFFF !important; }
    div.stButton > button[kind="primary"] {
        background-color: var(--ark-orange) !important;
        color: #FFFFFF !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
        border: none !important;
        box-shadow: 0 8px 20px rgba(247,148,40,0.25) !important;
        width: 100% !important;
    }
    div.stButton > button[kind="primary"]:hover { background-color:var(--ark-blue) !important; color:#FFFFFF !important; }

    .stat-card { background:white; border-radius:12px; padding:18px 16px; text-align:center; box-shadow:0 3px 10px rgba(0,0,0,0.06); margin-bottom:8px; }
    .stat-number { font-size:36px; font-weight:900; line-height:1; margin-bottom:6px; }
    .stat-label  { font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; color:#555; }
    .color-blue   { color:#0D6079; }
    .color-green  { color:#375623; }
    .color-red    { color:#9C0006; }
    .color-orange { color:#9C6500; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Fix "uploadupload" icon text in file uploaders
st.markdown("""
<script>
(function fixUploaders() {
    function hide() {
        document.querySelectorAll('[data-testid="stFileUploaderDropzone"] button').forEach(btn => {
            btn.querySelectorAll('[data-testid="stIconMaterial"]').forEach(el => {
                el.style.cssText = 'display:none!important;width:0!important;height:0!important;overflow:hidden!important;font-size:0!important;';
            });
            btn.childNodes.forEach(node => {
                if (node.nodeType === Node.TEXT_NODE && node.textContent.trim().toLowerCase() === 'upload') {
                    node.textContent = '';
                }
            });
        });
    }
    hide();
    new MutationObserver(hide).observe(document.body, { childList: true, subtree: true });
})();
</script>
""", unsafe_allow_html=True)

# ── Fixed Header ──────────────────────────────────────────────────────────────
logo_path = str(LOGO_PATH)
logo_b64  = img_to_base64(logo_path) if Path(logo_path).exists() else ""

st.markdown(
    f"""
    <div class="ark-nav">
      <div class="ark-nav-inner">
        <div class="ark-nav-left">
          <img src="data:image/png;base64,{logo_b64}"
            style="height:68px; width:auto; display:block;" />
          <div>
            <div class="ark-nav-title">M&amp;V Report Sanity Check</div>
            <div class="ark-nav-subtitle">Report vs Calculation Sheet Cross-Reference</div>
          </div>
        </div>
        <div class="ark-nav-right">
          <div class="pill">AI Assistant</div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Upload Section ────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="ark-section"><div class="ark-section-title">Upload documents</div></div>
    <div class="ark-section-rule"></div>
    """,
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)
with col1:
    report_upload = st.file_uploader(
        "M&V Report (PDF)",
        type=["pdf"],
        accept_multiple_files=False,
        help="The M&V Report PDF submitted by the ESP.",
    )
with col2:
    calc_upload = st.file_uploader(
        "M&V Calculation Sheet (Excel)",
        type=["xlsx"],
        accept_multiple_files=False,
        help="The M&V Calculation Sheet (.xlsx) submitted by the ESP.",
    )

components.html("""
<script>
(function fixUploadersFromComponent() {
    function hide(root) {
        root.querySelectorAll('[data-testid="stIconMaterial"]').forEach(el => {
            el.style.cssText = 'display:none!important;width:0!important;overflow:hidden!important;font-size:0!important;';
        });
        root.querySelectorAll('[data-testid="stFileUploaderDropzone"] button').forEach(btn => {
            btn.childNodes.forEach(node => {
                if (node.nodeType === 3 && node.textContent.trim().toLowerCase() === 'upload') {
                    node.textContent = '';
                }
            });
        });
    }
    function run() {
        try { hide(window.parent.document.body); } catch(e) {}
        try { hide(window.top.document.body); } catch(e) {}
    }
    run();
    try {
        new MutationObserver(run).observe(window.parent.document.body, { childList:true, subtree:true });
    } catch(e) {}
})();
</script>
""", height=0)

# ── Run button ────────────────────────────────────────────────────────────────
btn_left, btn_right = st.columns([7, 3])
with btn_right:
    run_btn = st.button(
        "Run Sanity Check",
        type="primary",
        disabled=(report_upload is None or calc_upload is None),
        use_container_width=True,
    )

# ── Validation & run ──────────────────────────────────────────────────────────
if run_btn:
    if report_upload is None:
        st.error("M&V Report PDF is required.")
        st.stop()
    if calc_upload is None:
        st.error("M&V Calculation Sheet (Excel) is required.")
        st.stop()

    log_box = st.empty()

    try:
        pdf_bytes  = report_upload.read()
        xlsx_bytes = calc_upload.read()

        log_box.info(
            f"Cross-referencing **{report_upload.name}** against "
            f"**{calc_upload.name}** — both documents are being read in full by the AI…"
        )

        with st.spinner(
            "Running sanity check — cross-referencing all key values between both documents… "
            "(this typically takes 30–90 seconds)"
        ):
            result = run_sanity_check(pdf_bytes, xlsx_bytes, report_upload.name)

        log_box.empty()

        st.markdown(
            """
            <div class="ark-section"><div class="ark-section-title">Results</div></div>
            <div class="ark-section-rule"></div>
            """,
            unsafe_allow_html=True,
        )
        st.success("Sanity check completed successfully.")

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="stat-card"><div class="stat-number color-blue">{result["total"]}</div><div class="stat-label">Questions Reviewed</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="stat-card"><div class="stat-number color-green">{result["approved"]}</div><div class="stat-label">Approved</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="stat-card"><div class="stat-number color-red">{result["not_approved"]}</div><div class="stat-label">Not Approved</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="stat-card"><div class="stat-number color-orange">{result["incomplete"]}</div><div class="stat-label">Incomplete</div></div>', unsafe_allow_html=True)

        if result["missing_sns"]:
            st.warning(
                f"**Warning — missing SNs:** The following questions were not returned by the AI "
                f"and have been left blank in the output: **{', '.join(result['missing_sns'])}**"
            )

        st.download_button(
            "⬇️ Download Filled Review Sheet",
            data=result["excel_bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        log_box.empty()
        st.error(f"{type(e).__name__}: {e}")
