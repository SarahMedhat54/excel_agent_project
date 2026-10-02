import asyncio
import html
import uuid

import streamlit as st
from openpyxl import load_workbook

from agent import run_agent
from server import FILES_DIR, _cell_value, _headers

st.set_page_config(page_title="Excel Assistant", page_icon="📊", layout="wide")

st.markdown(
    """
<style>
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {display:none !important;}
header[data-testid="stHeader"] {background: transparent;}
.block-container {padding-top: 1.4rem; max-width: 1250px;}

.hero {background: linear-gradient(135deg, #14543a 0%, #2e9e63 100%); color: #fff;
       border-radius: 18px; padding: 22px 28px; margin-bottom: 18px;
       box-shadow: 0 8px 24px rgba(0,0,0,.45);}
.hero h1 {margin: 0; font-size: 28px; font-weight: 700; color: #fff;}
.hero p {margin: 4px 0 0; opacity: .92; font-size: 14.5px;}

.step {display: flex; align-items: center; gap: 10px; font-weight: 700; font-size: 17px; margin-bottom: 6px; color: #e5ece7;}
.badge {display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px;
        border-radius: 50%; background: #2e9e63; color: #fff; font-size: 14px;}
.sub {color: #9bb0a2; font-size: 13px; margin: -2px 0 8px;}

div[data-testid="stVerticalBlockBorderWrapper"] {background: #16241c; border-radius: 16px !important;
        border: 1px solid #25382c !important; box-shadow: 0 4px 16px rgba(0,0,0,.35);}
.stButton > button, .stDownloadButton > button {border-radius: 10px; font-weight: 600; padding: .55rem 1rem;}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {background: #2e9e63; border: 0; color: #fff;}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {filter: brightness(1.12);}
.stButton > button[kind="secondary"] {background: #1d2f24; border: 1px solid #2c4435; color: #e5ece7;}
textarea {border-radius: 12px !important; background: #0f1a14 !important; color: #e5ece7 !important;}
section[data-testid="stFileUploaderDropzone"] {border-radius: 12px; background: #0f1a14; border: 1.5px dashed #3b6b50;}

.xl-wrap {overflow: auto; max-height: 430px; border: 1px solid #25382c; border-radius: 12px;}
.xl {border-collapse: collapse; width: 100%; font-size: 13px; color: #e5ece7;}
.xl th {position: sticky; top: 0; background: #1f7a4d; color: #fff; padding: 9px 14px; text-align: left; white-space: nowrap;}
.xl td {padding: 8px 14px; border-bottom: 1px solid #22332a; white-space: nowrap; background: #16241c;}
.xl tr:nth-child(even) td {background: #1b2c22;}
.xl td.num {text-align: right; font-variant-numeric: tabular-nums;}
.xl .rn {color: #7d9486; background: #121e17 !important; text-align: center; width: 1%;}
.xl th.rn {background: #14543a; color: #b9d8c6;}

.pill {display: inline-block; background: #1f3a2b; color: #7fd6a5; border-radius: 999px;
       padding: 3px 12px; font-size: 12.5px; font-weight: 600; margin: 4px 6px 10px 0;}
.empty {text-align: center; padding: 60px 20px; color: #9bb0a2;}
.empty .big {font-size: 46px; margin-bottom: 6px;}
</style>
""",
    unsafe_allow_html=True,
)

ss = st.session_state
ss.setdefault("history", [])       # what the agent sees
ss.setdefault("file_id", None)     # current file
ss.setdefault("reply", None)       # last agent reply
ss.setdefault("uploaded_key", None)
ss.setdefault("n", 0)              # used to reset the widgets


def clean_error(e: Exception) -> str:
    while hasattr(e, "exceptions") and e.exceptions:  # unwrap ExceptionGroup
        e = e.exceptions[0]
    msg = str(e)
    low = msg.lower()
    if "rate_limit" in low or "429" in msg:
        return "Groq rate limit reached. Wait a minute and press Run again."
    if "api key" in low or "401" in msg:
        return "Invalid Groq API key. Check GROQ_API_KEY in your .env file."
    if "model" in low and ("not found" in low or "decommissioned" in low):
        return "This model is not available on Groq. Change MODEL in .env (see console.groq.com/docs/models)."
    return msg


def preview_html(file_id: str, limit: int = 30):
    """Return (html_table, total_rows, chart_count) for the first sheet."""
    ws = load_workbook(FILES_DIR / f"{file_id}.xlsx").active
    heads = "".join(f"<th>{html.escape(str(h) if h is not None else '')}</th>" for h in _headers(ws))
    out = [f'<table class="xl"><tr><th class="rn">#</th>{heads}</tr>']
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=min(ws.max_row, limit + 1)), start=2):
        cells = [f'<td class="rn">{i}</td>']
        for c in row:
            v = _cell_value(ws, c.coordinate)
            is_num = isinstance(v, (int, float)) and not isinstance(v, bool)
            if isinstance(v, float):
                v = round(v, 2)
            cells.append(f'<td class="{"num" if is_num else ""}">{html.escape("" if v is None else str(v))}</td>')
        out.append("<tr>" + "".join(cells) + "</tr>")
    out.append("</table>")
    return f'<div class="xl-wrap">{"".join(out)}</div>', max(ws.max_row - 1, 0), len(ws._charts)


# ---------------- header ----------------
st.markdown(
    '<div class="hero"><h1>📊 Excel Assistant</h1>'
    "<p>Describe the spreadsheet you need, then edit it, analyze it, or add charts, all in plain words.</p></div>",
    unsafe_allow_html=True,
)

left, right = st.columns([1, 1.45], gap="large")

# ---------------- LEFT: upload + write ----------------
with left:
    with st.container(border=True):
        st.markdown('<div class="step"><span class="badge">1</span> Upload a file</div>'
                    '<div class="sub">Optional. Skip it if you want a brand new file.</div>', unsafe_allow_html=True)
        up = st.file_uploader("Excel file", type=["xlsx"], key=f"up{ss.n}", label_visibility="collapsed")
        if up is not None:
            key = f"{up.name}-{up.size}"
            if ss.uploaded_key != key:
                fid = uuid.uuid4().hex[:8]
                path = FILES_DIR / f"{fid}.xlsx"
                path.write_bytes(up.getvalue())
                try:
                    load_workbook(path)
                except Exception:
                    path.unlink(missing_ok=True)
                    st.error("That file could not be read as an Excel workbook.")
                else:
                    ss.file_id, ss.uploaded_key, ss.reply = fid, key, None
                    ss.history = []  # fresh conversation for the new file
                    st.success(f"Uploaded: {up.name}")

    with st.container(border=True):
        st.markdown('<div class="step"><span class="badge">2</span> Write what you want</div>'
                    '<div class="sub">Create, edit, add a chart, or analyze the data.</div>', unsafe_allow_html=True)
        request = st.text_area(
            "Request",
            key=f"req{ss.n}",
            height=150,
            placeholder="e.g. Create a sales sheet with Month, Product, Quantity, Price and 8 sample rows",
            label_visibility="collapsed",
        )
        c1, c2 = st.columns([2, 1])
        run = c1.button("▶  Run", type="primary", use_container_width=True)
        reset = c2.button("New file", use_container_width=True)

        if reset:
            ss.history, ss.file_id, ss.reply, ss.uploaded_key = [], None, None, None
            ss.n += 1
            st.rerun()

        if run:
            if not request.strip():
                st.warning("Write something first.")
            else:
                text = f"[Current file_id: {ss.file_id}]\n{request}" if ss.file_id else request
                ss.history.append({"role": "user", "content": text})
                mark = len(ss.history) - 1
                with st.spinner("Working on it..."):
                    try:
                        reply, file_ids = asyncio.run(run_agent(ss.history))
                        if file_ids:
                            ss.file_id = file_ids[-1]
                        ss.reply = reply or "Done."
                    except Exception as e:
                        del ss.history[mark:]  # drop the failed turn so the conversation stays valid
                        ss.reply = f"⚠️ {clean_error(e)}"

# ---------------- RIGHT: result ----------------
with right:
    with st.container(border=True):
        st.markdown('<div class="step"><span class="badge">3</span> Result</div>', unsafe_allow_html=True)

        if ss.reply:
            if ss.reply.startswith("⚠️"):
                st.error(ss.reply[2:].strip())
            else:
                st.markdown(ss.reply)

        path = FILES_DIR / f"{ss.file_id}.xlsx" if ss.file_id else None
        if path and path.exists():
            table, total, charts = preview_html(ss.file_id)
            pills = f'<span class="pill">{total} rows</span>'
            if charts:
                pills += f'<span class="pill">{charts} chart(s): open the file in Excel to see them</span>'
            st.markdown(pills, unsafe_allow_html=True)
            st.download_button(
                "⬇  Download Excel",
                data=path.read_bytes(),
                file_name=f"{ss.file_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            st.markdown(table, unsafe_allow_html=True)
        elif not ss.reply:
            st.markdown(
                '<div class="empty"><div class="big">📄</div>'
                "Your spreadsheet will show up here.<br>Write a request on the left and press <b>Run</b>.</div>",
                unsafe_allow_html=True,
            )