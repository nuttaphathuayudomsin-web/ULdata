"""
DR Filing Autopilot — Streamlit App
=====================================
Install:
    pip install streamlit anthropic openpyxl pandas

Run:
    streamlit run dr_filing_app.py

Set Streamlit secret: ANTHROPIC_API_KEY = 'sk-ant-...'
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
st.set_page_config(page_title="DR Filing Autopilot", page_icon="📋", layout="wide")

# ── Excel headers — exactly matching template, then 2 extra cols at end ───────
EXCEL_HEADERS = [
    "Run", "Company name", "Full company name", "Exchange name",
    "DR Ticker", "Units", "Ratio", "Address 1", "Address 2", "Tel", "Fax",
    "Company website", "Market name website", "Market website short",
    "UL Market website", "UL IR webiste", "UL IR News",
    "ATO fee", "", "Period",
    "Company name (Title Case)", "Latest Price",
]

SYSTEM_PROMPT = """You are a financial data assistant for a Thai DR (Depositary Receipt) issuer.
Given a stock ticker or company name, search the web for current data and return ONLY a valid JSON object — no markdown, no explanation.

Required keys:
{
  "companyName": "OFFICIAL NAME IN ALL CAPS e.g. TESLA INC.",
  "fullCompanyName": "ALL CAPS NAME + short nickname in double quotes e.g. TESLA INC. (\\"TESLA\\")",
  "companyNameTitle": "Title Case name e.g. Tesla Inc.",
  "latestPrice": "Current price with currency and date e.g. USD 248.50 (25 Feb 2026)",
  "exchangeName": "Thai exchange name + English — see mapping below",
  "exchangeNameEn": "English exchange name only e.g. NASDAQ",
  "drTicker": "Stock ticker + 80 e.g. TSLA80",
  "ratio": 1000,
  "address1": "Registered office address line 1",
  "address2": "City, State, Country, ZIP",
  "tel": "Phone with country code",
  "fax": "",
  "companyWebsite": "https://...",
  "marketNameWebsite": "Thai exchange name + English + (URL)",
  "marketWebsiteShort": "Exchange homepage URL",
  "ulMarketWebsite": "Direct stock quote page URL on the exchange",
  "ulIrWebsite": "Investor Relations page URL",
  "ulIrNews": "IR News / Press Releases URL",
  "atoFee": 0.4,
  "period": ""
}

ratio should be an integer: 100 if price < 100, 1000 if price 100–999, 10000 if price >= 1000.
atoFee: 0.4 for all exchanges except Euronext Paris which is 0.5.

Thai exchange name mappings (use EXACTLY):
NYSE → นิวยอร์ก (NYSE) | https://www.nyse.com/
NASDAQ → แนสแด็ก (NASDAQ) | https://www.nasdaq.com/
Tokyo Stock Exchange → โตเกียว (Tokyo Stock Exchange) | https://www.jpx.co.jp/english/
HKEX → ฮ่องกง (The Stock Exchange of Hong Kong) เขตปกครองพิเศษฮ่องกง | https://www.hkex.com.hk/
Shanghai → เซี่ยงไฮ้ (Shanghai Stock Exchange) ประเทศจีน | https://english.sse.com.cn/home/
Shenzhen → เซิ้นเจิ้น (Shenzhen Stock Exchange) ประเทศจีน | https://www.szse.cn/English/index.html
Euronext Paris → ปารีส (Euronext Paris) | https://www.euronext.com/en/markets/paris
LSE → ลอนดอน (London Stock Exchange) | https://www.londonstockexchange.com/"""


def fetch_dr_data(query: str, api_key: str) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1000,
                tools=[{"type": "web_search_20250305", "name": "web_search"}],
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": f"DR filing data for: {query}"}],
            )
            # Extract text block
            text = next((b.text for b in response.content if b.type == "text"), "")
            text = re.sub(r"```json\s*|```\s*", "", text).strip()
            m = re.search(r'\{.*\}', text, re.DOTALL)
            if not m:
                raise ValueError("No JSON found in response")
            return json.loads(m.group(0))
        except Exception as e:
            err = str(e)
            if "rate_limit" in err and attempt < 2:
                import time
                wait = 30 * (attempt + 1)
                st.warning(f"Rate limit — waiting {wait}s before retry {attempt+2}/3…")
                time.sleep(wait)
            else:
                raise


def build_excel(stock_list: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Single Stock"

    hdr_fill  = PatternFill("solid", fgColor="1F3864")
    hdr_font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    cell_font = Font(name="Arial", size=10)
    center    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="D0D0D0")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Row 1: blank
    ws.append([""] * len(EXCEL_HEADERS))
    ws.row_dimensions[1].height = 6

    # Row 2: headers
    ws.append(EXCEL_HEADERS)
    for ci in range(1, len(EXCEL_HEADERS) + 1):
        c = ws.cell(row=2, column=ci)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = border
    ws.row_dimensions[2].height = 30

    # Data rows
    for rn, s in enumerate(stock_list, start=3):
        row = [
            s.get("run", "N"),
            s.get("companyName", ""),
            s.get("fullCompanyName", ""),
            s.get("exchangeName", ""),
            s.get("drTicker", ""),
            "•",
            s.get("ratio", ""),
            s.get("address1", ""),
            s.get("address2", ""),
            s.get("tel", ""),
            s.get("fax", ""),
            s.get("companyWebsite", ""),
            s.get("marketNameWebsite", ""),
            s.get("marketWebsiteShort", ""),
            s.get("ulMarketWebsite", ""),
            s.get("ulIrWebsite", ""),
            s.get("ulIrNews", ""),
            s.get("atoFee", 0.4),
            "",
            s.get("period", ""),
            s.get("companyNameTitle", ""),
            s.get("latestPrice", ""),
        ]
        ws.append(row)
        for ci in range(1, len(EXCEL_HEADERS) + 1):
            c = ws.cell(row=rn, column=ci)
            c.font = cell_font; c.border = border
            c.alignment = center if ci in (1, 5, 6, 7) else left
        ws.row_dimensions[rn].height = 20

    col_widths = [6,35,45,32,12,6,8,35,30,18,14,30,45,30,50,45,45,8,4,10,35,28]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Session state ─────────────────────────────────────────────────────────────
if "stock_list" not in st.session_state:
    st.session_state.stock_list = []

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .badge { display:inline-block; padding:2px 8px; border-radius:4px;
             font-size:11px; font-weight:700; letter-spacing:.05em; margin-bottom:4px; }
    .badge-col   { background:#1f3864; color:#fff; }
    .badge-price { background:#1a3a1a; color:#6abf6a; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## 📋 DR Filing Autopilot")
st.caption("Claude AI + Web Search → Review & Edit → Build List → Download Excel")
st.divider()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    _secret_key = ""
    try:
        _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        pass
    api_key = _secret_key or os.environ.get("ANTHROPIC_API_KEY", "")

    st.markdown("### 🔑 Anthropic API Key")
    if api_key:
        st.success("API key loaded ✓", icon="🔒")
    else:
        api_key = st.text_input("API Key", type="password", placeholder="sk-ant-...", label_visibility="collapsed")
        if not api_key:
            st.caption("Add to Streamlit secrets: `ANTHROPIC_API_KEY = 'sk-ant-...'`")

    st.divider()
    st.markdown(f"### 📂 List ({len(st.session_state.stock_list)} stocks)")
    if st.session_state.stock_list:
        for i, s in enumerate(st.session_state.stock_list):
            c1, c2 = st.columns([5, 1])
            c1.markdown(f"**{i+1}.** {s.get('companyName','—')}")
            if c2.button("✕", key=f"del_{i}"):
                st.session_state.stock_list.pop(i)
                st.rerun()
    else:
        st.caption("No stocks added yet.")

    st.divider()
    if st.session_state.stock_list:
        st.download_button(
            "⬇️ Download Excel",
            data=build_excel(st.session_state.stock_list),
            file_name=f"DR_Filing_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )
        if st.button("🗑️ Clear All", use_container_width=True):
            st.session_state.stock_list = []
            st.rerun()

# ── Lookup ────────────────────────────────────────────────────────────────────
st.markdown("### 🔍 Look Up a Stock")

with st.expander("📖 How to enter ticker symbols", expanded=False):
    st.markdown("""
| Exchange | Format | Examples |
|---|---|---|
| 🇺🇸 NYSE / NASDAQ (USA) | `TICKER` | `AAPL` · `TSLA` · `COIN` · `CAT` |
| 🇯🇵 Tokyo Stock Exchange | `XXXXXX.T` | `7203.T` · `6857.T` · `6146.T` |
| 🇭🇰 Hong Kong (HKEX) | `XXXX.HK` | `9988.HK` · `1772.HK` · `9626.HK` |
| 🇨🇳 Shanghai (SSE) | `XXXXXX.SS` | `688041.SS` |
| 🇨🇳 Shenzhen (SZSE) | `XXXXXX.SZ` | `300476.SZ` |
| 🇫🇷 Euronext Paris | `XX.PA` | `EL.PA` · `SU.PA` |
| 🇬🇧 London (LSE) | `XX.L` | `AZN.L` · `SHEL.L` |

> 💡 You can also type the full company name e.g. **Apple Inc** or **Tesla** — Claude will find it.
""")

c1, c2 = st.columns([5, 1])
query    = c1.text_input("Ticker", placeholder="e.g. AAPL · TSLA · 7203.T · 9988.HK · SU.PA", label_visibility="collapsed")
retrieve = c2.button("Retrieve", type="primary", use_container_width=True)

if retrieve:
    if not api_key:
        st.error("Please add your Anthropic API key in the sidebar.")
    elif not query.strip():
        st.warning("Enter a ticker or company name.")
    else:
        with st.spinner(f"Searching for **{query}**…"):
            try:
                data = fetch_dr_data(query.strip(), api_key)
                data.setdefault("run", "N")
                st.session_state["pending"] = data
            except Exception as e:
                st.error(f"Error: {e}")

# ── Review form ───────────────────────────────────────────────────────────────
if "pending" in st.session_state:
    data = st.session_state["pending"]
    st.divider()
    st.markdown("### ✏️ Review & Edit — then Add to List")

    st.info(f"**{data.get('companyName','')}** · {data.get('exchangeNameEn','')} · 💰 {data.get('latestPrice','—')}")

    st.markdown("**Name columns:**")
    n1, n2, n3 = st.columns(3)
    with n1:
        st.markdown('<span class="badge badge-col">COL 1 — ALL CAPS</span>', unsafe_allow_html=True)
        data["companyName"]      = st.text_area("col1", value=data.get("companyName",""),      height=72, label_visibility="collapsed", key="e1")
    with n2:
        st.markdown('<span class="badge badge-col">COL 2 — CAPS + ("Nickname")</span>', unsafe_allow_html=True)
        data["fullCompanyName"]  = st.text_area("col2", value=data.get("fullCompanyName",""),  height=72, label_visibility="collapsed", key="e2")
    with n3:
        st.markdown('<span class="badge badge-col">COL 3 — Title Case</span>', unsafe_allow_html=True)
        data["companyNameTitle"] = st.text_area("col3", value=data.get("companyNameTitle",""), height=72, label_visibility="collapsed", key="e3")

    st.markdown('<span class="badge badge-price">PRICE</span>', unsafe_allow_html=True)
    data["latestPrice"] = st.text_input("Latest Price", value=data.get("latestPrice",""), key="ep")

    st.divider()
    st.markdown("**Filing fields:**")
    ca, cb = st.columns(2)
    with ca:
        data["exchangeName"]       = st.text_input("Exchange name (Thai)",   value=data.get("exchangeName",""),       key="ef1")
        data["drTicker"]           = st.text_input("DR Ticker",              value=data.get("drTicker",""),           key="ef2")
        data["ratio"]              = st.text_input("Ratio",                  value=str(data.get("ratio","")),         key="ef3")
        data["address1"]           = st.text_input("Address 1",              value=data.get("address1",""),           key="ef4")
        data["address2"]           = st.text_input("Address 2",              value=data.get("address2",""),           key="ef5")
        data["tel"]                = st.text_input("Tel",                    value=data.get("tel",""),                key="ef6")
        data["fax"]                = st.text_input("Fax",                    value=data.get("fax",""),                key="ef7")
    with cb:
        data["companyWebsite"]     = st.text_input("Company website",        value=data.get("companyWebsite",""),     key="ef8")
        data["marketNameWebsite"]  = st.text_area("Market name website",     value=data.get("marketNameWebsite",""),  height=72, key="ef9")
        data["marketWebsiteShort"] = st.text_input("Market website short",   value=data.get("marketWebsiteShort",""), key="ef10")
        data["ulMarketWebsite"]    = st.text_input("UL Market website",      value=data.get("ulMarketWebsite",""),    key="ef11")
        data["ulIrWebsite"]        = st.text_input("UL IR website",          value=data.get("ulIrWebsite",""),        key="ef12")
        data["ulIrNews"]           = st.text_input("UL IR News",             value=data.get("ulIrNews",""),           key="ef13")

    cf, cp, cr = st.columns(3)
    data["atoFee"] = cf.text_input("ATO fee", value=str(data.get("atoFee", 0.4)), key="ef14")
    data["period"] = cp.text_input("Period",  value=data.get("period",""), placeholder="e.g. Q326", key="ef15")
    data["run"]    = cr.selectbox("Run", ["N","Y"], index=0 if data.get("run","N")=="N" else 1, key="ef16")

    st.divider()
    ba, bb = st.columns([2, 1])
    if ba.button("✅ Add to List", type="primary", use_container_width=True):
        st.session_state.stock_list.append(dict(data))
        del st.session_state["pending"]
        st.rerun()
    if bb.button("✕ Discard", use_container_width=True):
        del st.session_state["pending"]
        st.rerun()

# ── List preview ──────────────────────────────────────────────────────────────
if st.session_state.stock_list:
    st.divider()
    st.markdown(f"### 📋 Your List — {len(st.session_state.stock_list)} stock(s)")
    rows = []
    for i, s in enumerate(st.session_state.stock_list, 1):
        rows.append({
            "#":                       i,
            "Run":                     s.get("run","N"),
            "Company name (ALL CAPS)": s.get("companyName",""),
            "Full company name":       s.get("fullCompanyName",""),
            "Title Case":              s.get("companyNameTitle",""),
            "Latest Price":            s.get("latestPrice",""),
            "Exchange":                s.get("exchangeNameEn",""),
            "DR Ticker":               s.get("drTicker",""),
            "Ratio":                   s.get("ratio",""),
            "Address 1":               s.get("address1",""),
            "Address 2":               s.get("address2",""),
            "Tel":                     s.get("tel",""),
            "Company website":         s.get("companyWebsite",""),
            "Market name website":     s.get("marketNameWebsite",""),
            "Market website short":    s.get("marketWebsiteShort",""),
            "UL Market website":       s.get("ulMarketWebsite",""),
            "UL IR website":           s.get("ulIrWebsite",""),
            "UL IR News":              s.get("ulIrNews",""),
            "ATO fee":                 s.get("atoFee",""),
            "Period":                  s.get("period",""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("👈 Click **Download Excel** in the sidebar when your list is complete.")

elif "pending" not in st.session_state:
    st.markdown("""
    <div style="text-align:center;padding:60px 0;color:#aaa">
        <h3>No stocks yet</h3>
        <p>Enter a ticker or company name above → review the data → click <b>Add to List</b><br>
        Repeat for each stock, then download your Excel.</p>
    </div>""", unsafe_allow_html=True)
