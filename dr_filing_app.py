"""
DR Filing Autopilot — Streamlit App
=====================================
Install dependencies:
    pip install streamlit anthropic openpyxl pandas

Run:
    streamlit run dr_filing_app.py

Requires environment variable:
    ANTHROPIC_API_KEY=your_key_here
  OR enter it in the sidebar at runtime.
"""

import io
import json
import re
import os
from datetime import datetime

import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DR Filing Autopilot",
    page_icon="📋",
    layout="wide",
)

# ── Exact column headers matching the template (Row 1, 0-indexed cols 0-19) ──
EXCEL_HEADERS = [
    "Run",
    "Company name",
    "Full company name",
    "Exchange name",
    "DR Ticker",
    "Units",
    "Ratio",
    "Address 1",
    "Address 2",
    "Tel",
    "Fax",
    "Company website",
    "Market name website",
    "Market website short",
    "UL Market website",
    "UL IR webiste",          # kept exactly as in template (typo intentional)
    "UL IR News",
    "ATO fee",
    "",                        # blank column
    "Period",
    "Company name (Title Case)",
    "Latest Price",
]

# ── Field definitions ─────────────────────────────────────────────────────────
# Maps UI label → session-state key → excel column index
FIELDS = [
    # key,                  label,                                              excel_col
    ("run",                 "Run (N/Y)",                                        0),
    ("companyName",         "① Company name — ALL CAPS",                        1),
    ("fullCompanyName",     "② Full company name — CAPS + (\"Nickname\")",       2),
    ("companyNameTitle",    "③ Company name — Title Case (new col)",             None),   # extra, not in original template
    ("latestPrice",         "④ Latest Stock Price",                             None),   # extra, not in original template
    ("exchangeName",        "Exchange name",                                    3),
    ("drTicker",            "DR Ticker",                                        4),
    ("units",               "Units",                                            5),
    ("ratio",               "Ratio",                                            6),
    ("address1",            "Address 1",                                        7),
    ("address2",            "Address 2",                                        8),
    ("tel",                 "Tel",                                              9),
    ("fax",                 "Fax",                                              10),
    ("companyWebsite",      "Company website",                                  11),
    ("marketNameWebsite",   "Market name website",                              12),
    ("marketWebsiteShort",  "Market website short",                             13),
    ("ulMarketWebsite",     "UL Market website",                                14),
    ("ulIrWebsite",         "UL IR webiste",                                    15),
    ("ulIrNews",            "UL IR News",                                       16),
    ("atoFee",              "ATO fee",                                          17),
    ("period",              "Period",                                           19),
]

SYSTEM_PROMPT = """Return ONLY valid JSON, no markdown, no explanation.
Keys required: companyName (ALL CAPS), fullCompanyName (ALL CAPS + ("NICKNAME")), companyNameTitle (Title Case), latestPrice (e.g. USD 205.42 as of 25 Feb 2026), exchangeName (Thai+English), exchangeNameEn, drTicker (ticker+80), ratio (100/1000/10000 int), address1, address2, tel, fax, companyWebsite, marketNameWebsite, marketWebsiteShort, ulMarketWebsite, ulIrWebsite, ulIrNews, atoFee (0.4 default; 0.5 Euronext), period (blank).
Thai exchange names: NYSE->นิวยอร์ก (NYSE) nyse.com | NASDAQ->แนสแด็ก (NASDAQ) nasdaq.com | TSE->โตเกียว (Tokyo Stock Exchange) jpx.co.jp/english | HKEX->ฮ่องกง (The Stock Exchange of Hong Kong) เขตปกครองพิเศษฮ่องกง hkex.com.hk | SSE->เซี่ยงไฮ้ (Shanghai Stock Exchange) ประเทศจีน sse.com | SZSE->เซิ้นเจิ้น (Shenzhen Stock Exchange) ประเทศจีน szse.cn | Euronext->ปารีส (Euronext Paris) euronext.com | LSE->ลอนดอน (London Stock Exchange)"""


# ── Helper: call Claude API ───────────────────────────────────────────────────
def fetch_dr_data(query: str, api_key: str) -> dict:
    import time
    client = anthropic.Anthropic(api_key=api_key)
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=800,
                tools=[{"type": "web_search_20250305", "name": "web_search"}],
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": f"DR filing data for: {query}"}],
            )
            text = ""
            for block in response.content:
                if block.type == "text":
                    text = block.text.strip()
                    break
            text = re.sub(r"```json\s*", "", text)
            text = re.sub(r"```\s*", "", text)
            return json.loads(text.strip())
        except Exception as e:
            if "rate_limit" in str(e) and attempt < 2:
                wait = 30 * (attempt + 1)
                st.warning(f"Rate limit hit — waiting {wait}s before retry {attempt+2}/3…")
                time.sleep(wait)
            else:
                raise


# ── Helper: build Excel file ──────────────────────────────────────────────────
def build_excel(stock_list: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Single Stock"

    # ── Styling helpers ──
    header_font     = Font(name="Arial", bold=True, size=10)
    header_fill     = PatternFill("solid", fgColor="1F3864")   # dark navy
    header_font_col = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    cell_font       = Font(name="Arial", size=10)
    center          = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left            = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin            = Side(style="thin", color="D0D0D0")
    border          = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Row 1: blank (matches template) ──
    ws.append([""] * len(EXCEL_HEADERS))
    ws.row_dimensions[1].height = 6

    # ── Row 2: headers ──
    ws.append(EXCEL_HEADERS)
    for col_idx, header in enumerate(EXCEL_HEADERS, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.font  = header_font_col
        cell.fill  = header_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[2].height = 30

    # ── Data rows ──
    for row_num, stock in enumerate(stock_list, start=3):
        row_data = [
            stock.get("run", "N"),
            stock.get("companyName", ""),
            stock.get("fullCompanyName", ""),
            stock.get("companyNameTitle", ""),
            stock.get("exchangeName", ""),
            stock.get("drTicker", ""),
            "•",                                    # Units — bullet as in template
            stock.get("ratio", ""),
            stock.get("address1", ""),
            stock.get("address2", ""),
            stock.get("tel", ""),
            stock.get("fax", ""),
            stock.get("companyWebsite", ""),
            stock.get("marketNameWebsite", ""),
            stock.get("marketWebsiteShort", ""),
            stock.get("ulMarketWebsite", ""),
            stock.get("ulIrWebsite", ""),
            stock.get("ulIrNews", ""),
            stock.get("atoFee", 0.4),
            "",                                     # blank col
            stock.get("period", ""),
            stock.get("companyNameTitle", ""),
            stock.get("latestPrice", ""),
        ]
        ws.append(row_data)
        for col_idx in range(1, len(EXCEL_HEADERS) + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font      = cell_font
            cell.border    = border
            cell.alignment = center if col_idx in (1, 5, 6, 7) else left
        ws.row_dimensions[row_num].height = 20

    # ── Column widths ──
    col_widths = [6, 35, 45, 32, 12, 6, 8, 35, 30, 18, 14, 30, 45, 30, 50, 45, 45, 8, 4, 10, 35, 28]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # ── Freeze pane below header ──
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Session state init ────────────────────────────────────────────────────────
if "stock_list" not in st.session_state:
    st.session_state.stock_list = []   # list of dicts, one per retrieved stock

if "editing_idx" not in st.session_state:
    st.session_state.editing_idx = None


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .stTextInput > div > div > input { font-family: monospace; }
    .badge { display:inline-block; padding:2px 7px; border-radius:4px;
             font-size:11px; font-weight:700; letter-spacing:.05em; }
    .badge-col  { background:#1f3864; color:#fff; }
    .badge-price{ background:#1a3a1a; color:#6abf6a; }
    div[data-testid="stExpander"] { border: 1px solid #e0e0e0; border-radius:8px; margin-bottom:6px; }
</style>
""", unsafe_allow_html=True)

# Header
col_logo, col_title = st.columns([1, 10])
with col_logo:
    st.markdown("## 📋")
with col_title:
    st.markdown("## DR Filing Autopilot")
    st.caption("Retrieve → Review → Build your list → Download Excel")

st.divider()

# ── Sidebar: API key + list summary ──────────────────────────────────────────
with st.sidebar:
    # Priority: Streamlit secrets → env var → manual input
    _secret_key = ""
    try:
        _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        pass
    _auto_key = _secret_key or os.environ.get("ANTHROPIC_API_KEY", "")

    st.markdown("### 🔑 API Key")
    if _auto_key:
        api_key = _auto_key
        st.success("API key loaded ✓", icon="🔒")
    else:
        api_key = st.text_input(
            "Anthropic API Key",
            type="password",
            placeholder="sk-ant-...",
            help="Or add ANTHROPIC_API_KEY to Streamlit secrets",
        )
        if not api_key:
            st.caption("💡 To avoid entering this every time, add it to **Streamlit secrets**: Settings → Secrets → `ANTHROPIC_API_KEY = 'sk-ant-...'`")

    st.divider()
    st.markdown(f"### 📂 Current List ({len(st.session_state.stock_list)} stocks)")
    if st.session_state.stock_list:
        for i, s in enumerate(st.session_state.stock_list):
            col_name, col_del = st.columns([5, 1])
            col_name.markdown(f"**{i+1}.** {s.get('companyName','—')}")
            if col_del.button("✕", key=f"del_{i}", help="Remove"):
                st.session_state.stock_list.pop(i)
                st.rerun()
    else:
        st.caption("No stocks added yet.")

    st.divider()
    if st.session_state.stock_list:
        excel_bytes = build_excel(st.session_state.stock_list)
        st.download_button(
            label="⬇️ Download Excel",
            data=excel_bytes,
            file_name=f"DR_Filing_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        if st.button("🗑️ Clear All", use_container_width=True):
            st.session_state.stock_list = []
            st.rerun()


# ── Main: Lookup panel ────────────────────────────────────────────────────────
st.markdown("### 🔍 Look Up a Stock")
col_input, col_btn = st.columns([5, 1])
with col_input:
    query = st.text_input(
        "Ticker or company name",
        placeholder="e.g. AAPL · Apple Inc · 7203.T · Samsung Electronics",
        label_visibility="collapsed",
    )
with col_btn:
    retrieve = st.button("Retrieve", type="primary", use_container_width=True)

if retrieve:
    if not api_key:
        st.error("Please enter your Anthropic API key in the sidebar.")
    elif not query.strip():
        st.warning("Please enter a ticker or company name.")
    else:
        with st.spinner(f"Searching for **{query}**…"):
            try:
                data = fetch_dr_data(query.strip(), api_key)
                # Store as pending edit
                st.session_state["pending"] = data
                st.session_state["pending_query"] = query.strip()
            except Exception as e:
                st.error(f"Error: {e}")

# ── Pending result: review & confirm ─────────────────────────────────────────
if "pending" in st.session_state:
    data = st.session_state["pending"]
    st.divider()
    st.markdown("### ✏️ Review & Edit — then Add to List")

    with st.container():
        # Price callout
        price_val = data.get("latestPrice", "—")
        exch_val  = data.get("exchangeNameEn", "")
        st.info(f"**{data.get('companyName','')}** · {exch_val} · Latest Price: **{price_val}**")

        # Name columns highlighted
        st.markdown("**Name columns:**")
        nc1, nc2, nc3 = st.columns(3)
        with nc1:
            st.markdown('<span class="badge badge-col">COL 1 — ALL CAPS</span>', unsafe_allow_html=True)
            data["companyName"] = st.text_area("Company name (ALL CAPS)", value=data.get("companyName",""), height=68, key="e_companyName")
        with nc2:
            st.markdown('<span class="badge badge-col">COL 2 — CAPS + Nickname</span>', unsafe_allow_html=True)
            data["fullCompanyName"] = st.text_area('Full company name (CAPS + "Nickname")', value=data.get("fullCompanyName",""), height=68, key="e_fullCompanyName")
        with nc3:
            st.markdown('<span class="badge badge-col">COL 3 — Title Case</span>', unsafe_allow_html=True)
            data["companyNameTitle"] = st.text_area("Company name (Title Case)", value=data.get("companyNameTitle",""), height=68, key="e_companyNameTitle")

        st.markdown('<span class="badge badge-price">PRICE</span>', unsafe_allow_html=True)
        data["latestPrice"] = st.text_input("Latest Stock Price", value=data.get("latestPrice",""), key="e_latestPrice")

        st.divider()
        st.markdown("**Filing fields:**")

        col_a, col_b = st.columns(2)
        with col_a:
            data["exchangeName"]      = st.text_input("Exchange name (Thai)",       value=data.get("exchangeName",""),      key="e_exchangeName")
            data["drTicker"]          = st.text_input("DR Ticker",                  value=data.get("drTicker",""),          key="e_drTicker")
            data["ratio"]             = st.text_input("Ratio",                      value=str(data.get("ratio","")),        key="e_ratio")
            data["address1"]          = st.text_input("Address 1",                  value=data.get("address1",""),          key="e_address1")
            data["address2"]          = st.text_input("Address 2",                  value=data.get("address2",""),          key="e_address2")
            data["tel"]               = st.text_input("Tel",                        value=data.get("tel",""),               key="e_tel")
            data["fax"]               = st.text_input("Fax",                        value=data.get("fax",""),               key="e_fax")
        with col_b:
            data["companyWebsite"]    = st.text_input("Company website",            value=data.get("companyWebsite",""),    key="e_companyWebsite")
            data["marketNameWebsite"] = st.text_area("Market name website",         value=data.get("marketNameWebsite",""), height=68, key="e_marketNameWebsite")
            data["marketWebsiteShort"]= st.text_input("Market website short",       value=data.get("marketWebsiteShort",""),key="e_marketWebsiteShort")
            data["ulMarketWebsite"]   = st.text_input("UL Market website",          value=data.get("ulMarketWebsite",""),   key="e_ulMarketWebsite")
            data["ulIrWebsite"]       = st.text_input("UL IR website",              value=data.get("ulIrWebsite",""),       key="e_ulIrWebsite")
            data["ulIrNews"]          = st.text_input("UL IR News",                 value=data.get("ulIrNews",""),          key="e_ulIrNews")

        col_fee, col_period, col_run = st.columns(3)
        with col_fee:
            data["atoFee"] = st.text_input("ATO fee", value=str(data.get("atoFee", 0.4)), key="e_atoFee")
        with col_period:
            data["period"] = st.text_input("Period", value=data.get("period",""), placeholder="e.g. Q326", key="e_period")
        with col_run:
            data["run"] = st.selectbox("Run", options=["N", "Y"], index=0 if data.get("run","N") == "N" else 1, key="e_run")

        st.divider()
        add_col, cancel_col = st.columns([2, 1])
        with add_col:
            if st.button("✅ Add to List", type="primary", use_container_width=True):
                st.session_state.stock_list.append(dict(data))
                del st.session_state["pending"]
                if "pending_query" in st.session_state:
                    del st.session_state["pending_query"]
                st.success(f"Added **{data.get('companyName','')}** to list. ({len(st.session_state.stock_list)} total)")
                st.rerun()
        with cancel_col:
            if st.button("✕ Discard", use_container_width=True):
                del st.session_state["pending"]
                if "pending_query" in st.session_state:
                    del st.session_state["pending_query"]
                st.rerun()

# ── Current list table preview ────────────────────────────────────────────────
if st.session_state.stock_list:
    st.divider()
    st.markdown(f"### 📋 Your List — {len(st.session_state.stock_list)} stock(s)")
    preview_data = []
    for i, s in enumerate(st.session_state.stock_list, 1):
        preview_data.append({
            "#": i,
            "Run": s.get("run", "N"),
            "Company name (ALL CAPS)": s.get("companyName", ""),
            "Full company name": s.get("fullCompanyName", ""),
            "Company name (Title Case)": s.get("companyNameTitle", ""),
            "Latest Price": s.get("latestPrice", ""),
            "Exchange": s.get("exchangeNameEn", s.get("exchangeName", "")),
            "DR Ticker": s.get("drTicker", ""),
            "Ratio": s.get("ratio", ""),
            "Address 1": s.get("address1", ""),
            "Address 2": s.get("address2", ""),
            "Tel": s.get("tel", ""),
            "Company website": s.get("companyWebsite", ""),
            "Market name website": s.get("marketNameWebsite", ""),
            "Market website short": s.get("marketWebsiteShort", ""),
            "UL Market website": s.get("ulMarketWebsite", ""),
            "UL IR website": s.get("ulIrWebsite", ""),
            "UL IR News": s.get("ulIrNews", ""),
            "ATO fee": s.get("atoFee", 0.4),
            "Period": s.get("period", ""),
        })
    st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

    st.markdown("👈 Use the **Download Excel** button in the sidebar to export your completed list.")

else:
    if "pending" not in st.session_state:
        st.markdown("""
        <div style="text-align:center; padding:60px 0; color:#aaa;">
            <h3>No stocks added yet</h3>
            <p>Search for a stock above, review the auto-filled data, then click <strong>Add to List</strong>.<br>
            Repeat for each stock. When done, download the Excel file from the sidebar.</p>
        </div>
        """, unsafe_allow_html=True)
