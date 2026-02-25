"""
DR Filing Autopilot — Streamlit App
=====================================
Install:
    pip install streamlit yfinance openpyxl pandas anthropic

Run:
    streamlit run dr_filing_app.py

- Yahoo Finance fetches all company data FREE, no API key needed
- Anthropic API (optional) only used to guess IR links — ~$0.001/lookup
  Set as Streamlit secret: ANTHROPIC_API_KEY = 'sk-ant-...'
"""

import io
import json
import re
import os
from datetime import datetime

import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="DR Filing Autopilot", page_icon="📋", layout="wide")

# ── Excel headers — exactly matching template then 2 extra cols at end ────────
EXCEL_HEADERS = [
    "Run", "Company name", "Full company name", "Exchange name",
    "DR Ticker", "Units", "Ratio", "Address 1", "Address 2", "Tel", "Fax",
    "Company website", "Market name website", "Market website short",
    "UL Market website", "UL IR webiste", "UL IR News",
    "ATO fee", "", "Period",
    "Company name (Title Case)", "Latest Price",
]

# ── Exchange code → (Thai name, English name, market short URL, stock URL template, ATO fee)
EXCHANGE_MAP = {
    "NMS": ("แนสแด็ก (NASDAQ)", "NASDAQ", "https://www.nasdaq.com/", "https://www.nasdaq.com/market-activity/stocks/{t}", 0.4),
    "NGM": ("แนสแด็ก (NASDAQ)", "NASDAQ", "https://www.nasdaq.com/", "https://www.nasdaq.com/market-activity/stocks/{t}", 0.4),
    "NCM": ("แนสแด็ก (NASDAQ)", "NASDAQ", "https://www.nasdaq.com/", "https://www.nasdaq.com/market-activity/stocks/{t}", 0.4),
    "NYQ": ("นิวยอร์ก (NYSE)", "NYSE", "https://www.nyse.com/", "https://www.nyse.com/quote/XNYS:{t}", 0.4),
    "PCX": ("นิวยอร์กอาร์ก้า (NYSE Arca)", "NYSE Arca", "https://www.nyse.com/markets/nyse-arca", "https://www.nyse.com/markets/nyse-arca", 0.4),
    "TSE": ("โตเกียว (Tokyo Stock Exchange)", "Tokyo Stock Exchange", "https://www.jpx.co.jp/english/", "https://www.jpx.co.jp/english/", 0.4),
    "TYO": ("โตเกียว (Tokyo Stock Exchange)", "Tokyo Stock Exchange", "https://www.jpx.co.jp/english/", "https://www.jpx.co.jp/english/", 0.4),
    "HKG": ("ฮ่องกง (The Stock Exchange of Hong Kong) เขตปกครองพิเศษฮ่องกง", "HKEX", "https://www.hkex.com.hk/", "https://www.hkex.com.hk/Market-Data/Securities-Prices/Equities/Equities-Quote?sym={t}&sc_lang=en", 0.4),
    "SHH": ("เซี่ยงไฮ้ (Shanghai Stock Exchange) ประเทศจีน", "SSE", "https://english.sse.com.cn/home/", "https://english.sse.com.cn/home/", 0.4),
    "SHZ": ("เซิ้นเจิ้น (Shenzhen Stock Exchange) ประเทศจีน", "SZSE", "https://www.szse.cn/English/index.html", "https://www.szse.cn/English/index.html", 0.4),
    "EPA": ("ปารีส (Euronext Paris)", "Euronext Paris", "https://www.euronext.com/en/markets/paris", "https://live.euronext.com/en/product/equities/{t}", 0.5),
    "LSE": ("ลอนดอน (London Stock Exchange)", "LSE", "https://www.londonstockexchange.com/", "https://www.londonstockexchange.com/", 0.4),
}

MARKET_NAME_WEBSITE_MAP = {
    "NASDAQ":               "ตลาดหลักทรัพย์แนสแด็ก (NASDAQ) (https://www.nasdaq.com/)",
    "NYSE":                 "ตลาดหลักทรัพย์นิวยอร์ก (NYSE) (https://www.nyse.com/)",
    "NYSE Arca":            "นิวยอร์กอาร์ก้า (NYSE Arca)\nhttps://www.nyse.com/markets/nyse-arca",
    "Tokyo Stock Exchange": "Tokyo Stock Exchange  (https://www.jpx.co.jp/english/)",
    "HKEX":                 "The Stock Exchange of Hong Kong (https://www.hkex.com.hk/)",
    "SSE":                  "Shanghai Stock Exchange (https://english.sse.com.cn/home/)",
    "SZSE":                 "Shenzhen Stock Exchange (https://www.szse.cn/English/index.html)",
    "Euronext Paris":       "Euronext Paris\n(https://www.euronext.com/en/markets/paris)",
    "LSE":                  "London Stock Exchange (https://www.londonstockexchange.com/)",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def to_title_case(name: str) -> str:
    minor = {"a","an","the","and","but","or","for","nor","on","at","to","by","in","of","up","as","is"}
    words = name.lower().split()
    return " ".join(w.capitalize() if i == 0 or w not in minor else w for i, w in enumerate(words))

def suggest_ratio(price: float) -> int:
    if price >= 1000: return 10000
    if price >= 100:  return 1000
    return 100

def derive_nickname(long_name: str) -> str:
    """Strip legal suffixes to get a short nickname."""
    name = long_name.upper()
    for suffix in [
        " INC.", " INC", " CORP.", " CORP", " CO., LTD.", " CO., LTD",
        " CO.,LTD", " CO. LTD", " LIMITED", " LTD.", " LTD", " PLC",
        " S.A.", " SA.", " SA", " SE", " NV", " AG", " LLC",
        " HOLDINGS CORPORATION", " HOLDINGS CORP.", " HOLDINGS",
        " GROUP CO., LTD.", " GROUP CO., LTD", " GROUP",
        " CORPORATION", " TECHNOLOGIES", " TECHNOLOGY",
    ]:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
    return name.strip().rstrip(".,")

# ── Yahoo Finance fetch ───────────────────────────────────────────────────────
def fetch_yahoo_data(ticker_sym: str) -> dict:
    tk = yf.Ticker(ticker_sym)
    info = tk.info

    if not info.get("longName") and not info.get("shortName"):
        raise ValueError(f"Ticker '{ticker_sym}' not found on Yahoo Finance. Use the exact symbol e.g. AAPL, 7203.T, 9988.HK")

    # Price
    price = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose") or 0
    currency = info.get("currency", "USD")
    today = datetime.now().strftime("%d %b %Y")
    latest_price = f"{currency} {price:,.2f} (as of {today})"

    # Names
    long_name     = info.get("longName") or info.get("shortName") or ticker_sym
    company_name  = long_name.upper()
    nickname      = derive_nickname(long_name)
    full_name     = f'{company_name} ("{nickname}")'
    title_name    = to_title_case(long_name)

    # Exchange
    exch_code = info.get("exchange", "")
    exch_data = EXCHANGE_MAP.get(exch_code)
    if exch_data:
        exch_thai, exch_en, mkt_short, ul_tmpl, ato_fee = exch_data
    else:
        full_exch  = info.get("fullExchangeName", exch_code)
        exch_thai  = full_exch
        exch_en    = full_exch
        mkt_short  = ""
        ul_tmpl    = ""
        ato_fee    = 0.4

    ul_market = ul_tmpl.replace("{t}", ticker_sym)
    mkt_name_website = MARKET_NAME_WEBSITE_MAP.get(exch_en, f"{exch_thai} ({mkt_short})")

    # Address
    addr1_parts = [info.get("address1",""), info.get("address2","")]
    address1 = ", ".join(p for p in addr1_parts if p)
    address2 = ", ".join(p for p in [
        info.get("city",""), info.get("state",""),
        info.get("zip",""), info.get("country","")
    ] if p)

    return {
        "companyName":         company_name,
        "fullCompanyName":     full_name,
        "companyNameTitle":    title_name,
        "latestPrice":         latest_price,
        "exchangeName":        exch_thai,
        "exchangeNameEn":      exch_en,
        "drTicker":            ticker_sym.split(".")[0] + "80",
        "ratio":               suggest_ratio(price),
        "address1":            address1,
        "address2":            address2,
        "tel":                 info.get("phone", ""),
        "fax":                 "",
        "companyWebsite":      info.get("website", ""),
        "marketNameWebsite":   mkt_name_website,
        "marketWebsiteShort":  mkt_short,
        "ulMarketWebsite":     ul_market,
        "ulIrWebsite":         "",
        "ulIrNews":            "",
        "atoFee":              ato_fee,
        "period":              "",
        "run":                 "N",
        "_exch_en":            exch_en,
        "_ticker":             ticker_sym,
    }

# ── Optional Claude enrichment for IR links only (~$0.001/call) ──────────────
IR_PROMPT = """Given a company name, ticker, exchange and website, return ONLY valid JSON with:
{"ulIrWebsite": "investor relations URL", "ulIrNews": "IR news/press releases URL"}
No markdown, no explanation. Make your best guess based on common IR URL patterns."""

def enrich_ir_links(data: dict, api_key: str) -> dict:
    if not api_key:
        return data
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        prompt = f"Company: {data['companyName']}\nTicker: {data['_ticker']}\nExchange: {data['exchangeNameEn']}\nWebsite: {data['companyWebsite']}"
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            system=IR_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )
        text = re.sub(r"```json\s*|```\s*", "", resp.content[0].text).strip()
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            ir = json.loads(m.group(0))
            if ir.get("ulIrWebsite"):  data["ulIrWebsite"] = ir["ulIrWebsite"]
            if ir.get("ulIrNews"):     data["ulIrNews"]    = ir["ulIrNews"]
    except Exception:
        pass
    return data

def fetch_dr_data(ticker_sym: str, api_key: str) -> dict:
    data = fetch_yahoo_data(ticker_sym.strip().upper())
    data = enrich_ir_links(data, api_key)
    data.pop("_exch_en", None)
    data.pop("_ticker", None)
    return data

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(stock_list: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Single Stock"

    hdr_fill   = PatternFill("solid", fgColor="1F3864")
    hdr_font   = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    cell_font  = Font(name="Arial", size=10)
    center     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin       = Side(style="thin", color="D0D0D0")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Row 1: blank
    ws.append([""] * len(EXCEL_HEADERS))
    ws.row_dimensions[1].height = 6

    # Row 2: headers
    ws.append(EXCEL_HEADERS)
    for ci in range(1, len(EXCEL_HEADERS)+1):
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
        for ci in range(1, len(EXCEL_HEADERS)+1):
            c = ws.cell(row=rn, column=ci)
            c.font = cell_font; c.border = border
            c.alignment = center if ci in (1,5,6,7) else left
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
    .block-container{padding-top:1.5rem}
    .badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;letter-spacing:.05em;margin-bottom:4px}
    .badge-col{background:#1f3864;color:#fff}
    .badge-price{background:#1a3a1a;color:#6abf6a}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## 📋 DR Filing Autopilot")
st.caption("Yahoo Finance → Review & Edit → Build List → Download Excel  |  Free · No token limits")
st.divider()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # API key (optional — only for IR link enrichment)
    _secret_key = ""
    try:
        _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        pass
    api_key = _secret_key or os.environ.get("ANTHROPIC_API_KEY", "")

    st.markdown("### 🔑 Anthropic API Key")
    st.caption("Optional — only used to auto-fill IR links (~$0.001/lookup)")
    if api_key:
        st.success("API key loaded ✓", icon="🔒")
    else:
        api_key = st.text_input("API Key", type="password", placeholder="sk-ant-...", label_visibility="collapsed")
        st.caption("Without key: IR links will be blank (fill manually)")

    st.divider()
    st.markdown(f"### 📂 List ({len(st.session_state.stock_list)} stocks)")
    if st.session_state.stock_list:
        for i, s in enumerate(st.session_state.stock_list):
            c1, c2 = st.columns([5,1])
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

# ── Ticker format guide ──
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

> 💡 **Tip:** Not sure of the ticker? Search on [finance.yahoo.com](https://finance.yahoo.com) first, then copy the symbol exactly as shown.
""")

c1, c2 = st.columns([5,1])
query   = c1.text_input("Ticker", placeholder="e.g. AAPL · 7203.T · 9988.HK · SU.PA", label_visibility="collapsed")
retrieve = c2.button("Retrieve", type="primary", use_container_width=True)

if retrieve:
    if not query.strip():
        st.warning("Enter a ticker symbol.")
    else:
        with st.spinner(f"Fetching **{query.upper()}** from Yahoo Finance…"):
            try:
                data = fetch_dr_data(query.strip(), api_key)
                st.session_state["pending"] = data
            except Exception as e:
                st.error(str(e))

# ── Review form ───────────────────────────────────────────────────────────────
if "pending" in st.session_state:
    data = st.session_state["pending"]
    st.divider()
    st.markdown("### ✏️ Review & Edit — then Add to List")

    st.info(f"**{data.get('companyName','')}** · {data.get('exchangeNameEn','')} · 💰 {data.get('latestPrice','—')}")

    # Name cols
    st.markdown("**Name columns:**")
    n1, n2, n3 = st.columns(3)
    with n1:
        st.markdown('<span class="badge badge-col">COL 1 — ALL CAPS</span>', unsafe_allow_html=True)
        data["companyName"]     = st.text_area("col1", value=data.get("companyName",""),     height=72, label_visibility="collapsed", key="e1")
    with n2:
        st.markdown('<span class="badge badge-col">COL 2 — CAPS + ("Nickname")</span>', unsafe_allow_html=True)
        data["fullCompanyName"] = st.text_area("col2", value=data.get("fullCompanyName",""), height=72, label_visibility="collapsed", key="e2")
    with n3:
        st.markdown('<span class="badge badge-col">COL 3 — Title Case</span>', unsafe_allow_html=True)
        data["companyNameTitle"]= st.text_area("col3", value=data.get("companyNameTitle",""),height=72, label_visibility="collapsed", key="e3")

    st.markdown('<span class="badge badge-price">PRICE</span>', unsafe_allow_html=True)
    data["latestPrice"] = st.text_input("Latest Price", value=data.get("latestPrice",""), key="ep")

    st.divider()
    st.markdown("**Filing fields:**")
    ca, cb = st.columns(2)
    with ca:
        data["exchangeName"]       = st.text_input("Exchange name (Thai)",    value=data.get("exchangeName",""),       key="ef1")
        data["drTicker"]           = st.text_input("DR Ticker",               value=data.get("drTicker",""),           key="ef2")
        data["ratio"]              = st.text_input("Ratio",                   value=str(data.get("ratio","")),         key="ef3")
        data["address1"]           = st.text_input("Address 1",               value=data.get("address1",""),           key="ef4")
        data["address2"]           = st.text_input("Address 2",               value=data.get("address2",""),           key="ef5")
        data["tel"]                = st.text_input("Tel",                     value=data.get("tel",""),                key="ef6")
        data["fax"]                = st.text_input("Fax",                     value=data.get("fax",""),                key="ef7")
    with cb:
        data["companyWebsite"]     = st.text_input("Company website",         value=data.get("companyWebsite",""),     key="ef8")
        data["marketNameWebsite"]  = st.text_area("Market name website",      value=data.get("marketNameWebsite",""),  height=72, key="ef9")
        data["marketWebsiteShort"] = st.text_input("Market website short",    value=data.get("marketWebsiteShort",""), key="ef10")
        data["ulMarketWebsite"]    = st.text_input("UL Market website",       value=data.get("ulMarketWebsite",""),    key="ef11")
        data["ulIrWebsite"]        = st.text_input("UL IR website",           value=data.get("ulIrWebsite",""),        key="ef12")
        data["ulIrNews"]           = st.text_input("UL IR News",              value=data.get("ulIrNews",""),           key="ef13")

    cf, cp, cr = st.columns(3)
    data["atoFee"] = cf.text_input("ATO fee",  value=str(data.get("atoFee",0.4)), key="ef14")
    data["period"] = cp.text_input("Period",   value=data.get("period",""), placeholder="e.g. Q326", key="ef15")
    data["run"]    = cr.selectbox("Run", ["N","Y"], index=0 if data.get("run","N")=="N" else 1, key="ef16")

    st.divider()
    ba, bb = st.columns([2,1])
    if ba.button("✅ Add to List", type="primary", use_container_width=True):
        st.session_state.stock_list.append(dict(data))
        del st.session_state["pending"]
        st.rerun()
    if bb.button("✕ Discard", use_container_width=True):
        del st.session_state["pending"]
        st.rerun()

# ── List preview table ────────────────────────────────────────────────────────
if st.session_state.stock_list:
    st.divider()
    st.markdown(f"### 📋 Your List — {len(st.session_state.stock_list)} stock(s)")
    rows = []
    for i, s in enumerate(st.session_state.stock_list, 1):
        rows.append({
            "#":                          i,
            "Run":                        s.get("run","N"),
            "Company name (ALL CAPS)":    s.get("companyName",""),
            "Full company name":          s.get("fullCompanyName",""),
            "Title Case":                 s.get("companyNameTitle",""),
            "Latest Price":               s.get("latestPrice",""),
            "Exchange":                   s.get("exchangeNameEn",""),
            "DR Ticker":                  s.get("drTicker",""),
            "Ratio":                      s.get("ratio",""),
            "Address 1":                  s.get("address1",""),
            "Address 2":                  s.get("address2",""),
            "Tel":                        s.get("tel",""),
            "Company website":            s.get("companyWebsite",""),
            "Market name website":        s.get("marketNameWebsite",""),
            "Market website short":       s.get("marketWebsiteShort",""),
            "UL Market website":          s.get("ulMarketWebsite",""),
            "UL IR website":              s.get("ulIrWebsite",""),
            "UL IR News":                 s.get("ulIrNews",""),
            "ATO fee":                    s.get("atoFee",""),
            "Period":                     s.get("period",""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("👈 Click **Download Excel** in the sidebar when your list is complete.")

elif "pending" not in st.session_state:
    st.markdown("""
    <div style="text-align:center;padding:60px 0;color:#aaa">
        <h3>No stocks yet</h3>
        <p>Enter a ticker above → review the data → click <b>Add to List</b><br>
        Repeat for each stock, then download your Excel.</p>
    </div>""", unsafe_allow_html=True)
