import { useState, useCallback } from "react";

const FIELDS = [
  { key: "companyName",      label: "① Company Name — ALL CAPS",                    example: "COINBASE GLOBAL INC.",                      badge: "COL 1" },
  { key: "fullCompanyName",  label: "② Company Name — CAPS with (\"Nickname\")",     example: 'COINBASE GLOBAL INC. ("COINBASE")',          badge: "COL 2" },
  { key: "companyNameTitle", label: "③ Company Name — Title Case",                   example: "Coinbase Global Inc.",                       badge: "COL 3" },
  { key: "latestPrice",      label: "④ Latest Stock Price",                           example: "USD 205.42 (as of Feb 25, 2025)",            badge: "PRICE" },
  { key: "exchangeName",     label: "Exchange Name (Thai + English)",                 example: "แนสแด็ก (NASDAQ)" },
  { key: "drTicker",         label: "DR Ticker",                                      example: "COIN80" },
  { key: "ratio",            label: "DR Ratio",                                       example: "1000" },
  { key: "address1",         label: "Address Line 1",                                 example: "One Madison Avenue Suite 2400 New York" },
  { key: "address2",         label: "Address Line 2 (City, State, Country, ZIP)",     example: "NEW YORK USA 10010" },
  { key: "tel",              label: "Telephone",                                      example: "1-302-636-5401" },
  { key: "companyWebsite",   label: "Company Website",                                example: "https://www.coinbase.com/" },
  { key: "marketNameWebsite",label: "Market Name + Website (Thai)",                   example: "ตลาดหลักทรัพย์แนสแด็ก (NASDAQ) (https://www.nasdaq.com/)" },
  { key: "marketWebsiteShort",label:"Market Website (short)",                         example: "https://www.nasdaq.com/" },
  { key: "ulMarketWebsite",  label: "UL Market Website (stock quote page)",           example: "https://www.nasdaq.com/market-activity/stocks/coin" },
  { key: "ulIrWebsite",      label: "UL IR Website",                                  example: "https://investor.coinbase.com/home/default.aspx" },
  { key: "ulIrNews",         label: "UL IR News Page",                                example: "https://investor.coinbase.com/news/default.aspx" },
];

const SYSTEM_PROMPT = `You are an expert financial data assistant helping fill a Thai Depositary Receipt (DR) filing template. 
When given a stock ticker or company name, use web search to find the latest stock price and company details, then return a JSON object with these exact keys:
{
  "companyName": "FULL OFFICIAL COMPANY NAME IN ALL CAPITAL LETTERS, e.g. COINBASE GLOBAL INC. — must be ALL CAPS, no lowercase",
  "fullCompanyName": "Same ALL CAPS company name followed by the commonly known short name in double quotes in parentheses, e.g. COINBASE GLOBAL INC. (\\"COINBASE\\") — Col 1 value + (\\"short nickname\\")",
  "companyNameTitle": "Same company name with only the first letter of each word capitalized (Title Case), e.g. Coinbase Global Inc. — articles like 'and', 'of', 'the' can be lowercase",
  "latestPrice": "Latest stock price with currency and date, e.g. USD 205.42 (as of Feb 25, 2026) — search for the most current price available",
  "exchangeName": "Thai exchange name + English in parentheses, e.g. แนสแด็ก (NASDAQ) or นิวยอร์ก (NYSE) etc.",
  "exchangeNameEn": "English exchange name only, e.g. NASDAQ",
  "drTicker": "Ticker symbol + 80, e.g. COIN80",
  "ratio": "Suggested DR ratio (100, 1000, or 10000 based on stock price - higher price stock = higher ratio), just the number",
  "address1": "First line of registered office address",
  "address2": "City, State/Province, Country and ZIP/postal code",
  "tel": "Main telephone number with country code",
  "companyWebsite": "Official company website URL",
  "marketNameWebsite": "Thai exchange name with English and website URL in parentheses",
  "marketWebsiteShort": "Exchange website URL only",
  "ulMarketWebsite": "Direct URL to the stock's quote page on the exchange",
  "ulIrWebsite": "Investor Relations page URL",
  "ulIrNews": "IR News/Press Releases page URL"
}
Return ONLY valid JSON, no markdown, no explanation. Be precise with Thai text for exchange names.
Thai exchange name mappings:
- NYSE / New York Stock Exchange → นิวยอร์ก (NYSE), website: https://www.nyse.com/
- NASDAQ → แนสแด็ก (NASDAQ), website: https://www.nasdaq.com/
- Tokyo Stock Exchange → โตเกียว (Tokyo Stock Exchange), website: https://www.jpx.co.jp/english/
- Hong Kong Stock Exchange → ฮ่องกง (The Stock Exchange of Hong Kong), website: https://www.hkex.com.hk/
- Shanghai Stock Exchange → เซี่ยงไฮ้ (Shanghai Stock Exchange), website: https://english.sse.com.cn/home/
- Shenzhen Stock Exchange → เซิ้นเจิ้น (Shenzhen Stock Exchange), website: https://www.szse.cn/English/index.html
- Euronext Paris → ปารีส  (Euronext Paris), website: https://www.euronext.com/en/markets/paris
- London Stock Exchange → ลอนดอน (London Stock Exchange), website: https://www.londonstockexchange.com/`;

export default function DRFilingTool() {
  const [input, setInput] = useState("");
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [editableData, setEditableData] = useState({});
  const [copyStatus, setCopyStatus] = useState({});
  const [history, setHistory] = useState([]);

  const fetchData = useCallback(async () => {
    if (!input.trim()) return;
    setLoading(true);
    setError("");
    setResults(null);

    try {
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          tools: [{ type: "web_search_20250305", name: "web_search" }],
          system: SYSTEM_PROMPT,
          messages: [{ role: "user", content: `Look up and fill the DR filing data for: ${input.trim()}` }],
        }),
      });

      const data = await response.json();
      const textBlock = data.content?.find(b => b.type === "text");
      if (!textBlock) throw new Error("No text response from API");

      let raw = textBlock.text.trim();
      raw = raw.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
      const parsed = JSON.parse(raw);
      setResults(parsed);
      setEditableData(parsed);
      setHistory(h => [{ query: input.trim(), data: parsed }, ...h.slice(0, 9)]);
    } catch (e) {
      setError("Failed to fetch data: " + e.message);
    } finally {
      setLoading(false);
    }
  }, [input]);

  const handleEdit = (key, value) => {
    setEditableData(d => ({ ...d, [key]: value }));
  };

  const copyCell = (key) => {
    navigator.clipboard.writeText(editableData[key] || "");
    setCopyStatus(s => ({ ...s, [key]: true }));
    setTimeout(() => setCopyStatus(s => ({ ...s, [key]: false })), 1500);
  };

  const copyAll = () => {
    const tsv = FIELDS.map(f => editableData[f.key] || "").join("\t");
    navigator.clipboard.writeText(tsv);
  };

  const loadHistory = (item) => {
    setInput(item.query);
    setResults(item.data);
    setEditableData(item.data);
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0a0a0f",
      color: "#e8e4d9",
      fontFamily: "'DM Mono', 'Courier New', monospace",
      display: "flex",
      flexDirection: "column",
    }}>
      {/* Header */}
      <div style={{
        borderBottom: "1px solid #1e1e2e",
        padding: "24px 32px",
        display: "flex",
        alignItems: "center",
        gap: "16px",
        background: "linear-gradient(135deg, #0f0f1a 0%, #0a0a0f 100%)",
      }}>
        <div style={{
          width: 40, height: 40,
          background: "linear-gradient(135deg, #c8a96e, #e8c87a)",
          borderRadius: 8,
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 18, fontWeight: "bold", color: "#0a0a0f",
        }}>DR</div>
        <div>
          <div style={{ fontSize: 18, fontWeight: 700, letterSpacing: "0.05em", color: "#e8c87a" }}>
            DR FILING AUTOPILOT
          </div>
          <div style={{ fontSize: 11, color: "#555570", letterSpacing: "0.1em" }}>
            DEPOSITARY RECEIPT DATA RETRIEVAL SYSTEM
          </div>
        </div>
      </div>

      <div style={{ display: "flex", flex: 1 }}>
        {/* Sidebar - History */}
        <div style={{
          width: 220, minWidth: 220,
          borderRight: "1px solid #1e1e2e",
          padding: "20px 16px",
          background: "#08080e",
        }}>
          <div style={{ fontSize: 10, letterSpacing: "0.15em", color: "#444455", marginBottom: 12 }}>
            RECENT LOOKUPS
          </div>
          {history.length === 0 && (
            <div style={{ fontSize: 11, color: "#333344", lineHeight: 1.6 }}>
              Your lookup history will appear here
            </div>
          )}
          {history.map((item, i) => (
            <div
              key={i}
              onClick={() => loadHistory(item)}
              style={{
                padding: "8px 10px",
                marginBottom: 4,
                borderRadius: 4,
                cursor: "pointer",
                fontSize: 11,
                color: "#8888aa",
                border: "1px solid #1a1a2a",
                background: "#0d0d18",
                transition: "all 0.15s",
              }}
              onMouseEnter={e => e.currentTarget.style.background = "#141425"}
              onMouseLeave={e => e.currentTarget.style.background = "#0d0d18"}
            >
              {item.query}
            </div>
          ))}
        </div>

        {/* Main content */}
        <div style={{ flex: 1, padding: "28px 32px", overflowY: "auto" }}>
          {/* Search bar */}
          <div style={{ marginBottom: 28 }}>
            <div style={{ fontSize: 11, letterSpacing: "0.12em", color: "#666680", marginBottom: 10 }}>
              STOCK TICKER OR COMPANY NAME
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <input
                value={input}
                onChange={e => setInput(e.target.value)}
                onKeyDown={e => e.key === "Enter" && fetchData()}
                placeholder="e.g. AAPL · Apple Inc · TSLA · Samsung Electronics"
                style={{
                  flex: 1,
                  background: "#0f0f1a",
                  border: "1px solid #2a2a40",
                  borderRadius: 6,
                  padding: "12px 16px",
                  fontSize: 14,
                  color: "#e8e4d9",
                  outline: "none",
                  fontFamily: "inherit",
                  letterSpacing: "0.02em",
                  transition: "border-color 0.15s",
                }}
                onFocus={e => e.target.style.borderColor = "#c8a96e"}
                onBlur={e => e.target.style.borderColor = "#2a2a40"}
              />
              <button
                onClick={fetchData}
                disabled={loading || !input.trim()}
                style={{
                  padding: "12px 24px",
                  background: loading ? "#1a1a2a" : "linear-gradient(135deg, #c8a96e, #e8c87a)",
                  border: "none",
                  borderRadius: 6,
                  cursor: loading ? "wait" : "pointer",
                  fontSize: 13,
                  fontWeight: 700,
                  color: loading ? "#444" : "#0a0a0f",
                  letterSpacing: "0.08em",
                  fontFamily: "inherit",
                  minWidth: 120,
                  transition: "all 0.15s",
                }}
              >
                {loading ? "FETCHING..." : "RETRIEVE"}
              </button>
            </div>
          </div>

          {/* Error */}
          {error && (
            <div style={{
              padding: "12px 16px",
              background: "#1a0a0a",
              border: "1px solid #442222",
              borderRadius: 6,
              fontSize: 12,
              color: "#cc6666",
              marginBottom: 20,
            }}>
              {error}
            </div>
          )}

          {/* Loading state */}
          {loading && (
            <div style={{ textAlign: "center", padding: "60px 0", color: "#444455" }}>
              <div style={{ fontSize: 12, letterSpacing: "0.15em", marginBottom: 12 }}>
                SEARCHING DATABASES
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: 6 }}>
                {[0,1,2].map(i => (
                  <div key={i} style={{
                    width: 8, height: 8, borderRadius: "50%",
                    background: "#c8a96e",
                    animation: `pulse 1.2s ${i * 0.2}s infinite`,
                  }} />
                ))}
              </div>
            </div>
          )}

          {/* Results */}
          {results && !loading && (
            <div>
              <div style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: 16,
              }}>
                <div>
                  <div style={{ fontSize: 16, color: "#e8c87a", fontWeight: 700 }}>
                    {editableData.companyName || input}
                  </div>
                  <div style={{ fontSize: 11, color: "#555570", marginTop: 2 }}>
                    {editableData.exchangeNameEn} · All fields editable before copying
                  </div>
                </div>
                <button
                  onClick={copyAll}
                  style={{
                    padding: "8px 16px",
                    background: "#1a1a2a",
                    border: "1px solid #3a3a50",
                    borderRadius: 5,
                    cursor: "pointer",
                    fontSize: 11,
                    color: "#8888aa",
                    letterSpacing: "0.08em",
                    fontFamily: "inherit",
                    transition: "all 0.15s",
                  }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = "#c8a96e"; e.currentTarget.style.color = "#c8a96e"; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = "#3a3a50"; e.currentTarget.style.color = "#8888aa"; }}
                >
                  COPY ALL AS TSV
                </button>
              </div>

              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {FIELDS.map(field => (
                  <div key={field.key} style={{
                    display: "grid",
                    gridTemplateColumns: "260px 1fr 36px",
                    gap: 8,
                    alignItems: "center",
                    padding: "8px 0",
                    borderBottom: "1px solid #12121c",
                    background: field.badge ? "rgba(200,169,110,0.03)" : "transparent",
                    borderRadius: field.badge ? 4 : 0,
                    paddingLeft: field.badge ? 6 : 0,
                  }}>
                    <div style={{ fontSize: 10, color: "#555570", letterSpacing: "0.08em", lineHeight: 1.6, display: "flex", alignItems: "center", gap: 6 }}>
                      {field.badge && (
                        <span style={{
                          fontSize: 9,
                          padding: "2px 5px",
                          borderRadius: 3,
                          background: field.badge === "PRICE" ? "#1a2a1a" : "#1a1a2e",
                          color: field.badge === "PRICE" ? "#6abf6a" : "#c8a96e",
                          border: `1px solid ${field.badge === "PRICE" ? "#2a4a2a" : "#2a2a50"}`,
                          letterSpacing: "0.05em",
                          fontWeight: 700,
                          flexShrink: 0,
                        }}>{field.badge}</span>
                      )}
                      {field.label}
                    </div>
                    <input
                      value={editableData[field.key] || ""}
                      onChange={e => handleEdit(field.key, e.target.value)}
                      style={{
                        background: "transparent",
                        border: "1px solid transparent",
                        borderRadius: 4,
                        padding: "6px 8px",
                        fontSize: 12,
                        color: field.badge ? "#e8e4d9" : "#d4d0c5",
                        outline: "none",
                        fontFamily: "inherit",
                        width: "100%",
                        boxSizing: "border-box",
                        transition: "all 0.15s",
                        fontWeight: field.badge ? 500 : 400,
                      }}
                      onFocus={e => { e.target.style.borderColor = "#2a2a40"; e.target.style.background = "#0f0f1a"; }}
                      onBlur={e => { e.target.style.borderColor = "transparent"; e.target.style.background = "transparent"; }}
                    />
                    <button
                      onClick={() => copyCell(field.key)}
                      title="Copy"
                      style={{
                        width: 28, height: 28,
                        background: "none",
                        border: "1px solid #2a2a40",
                        borderRadius: 4,
                        cursor: "pointer",
                        fontSize: 12,
                        color: copyStatus[field.key] ? "#6abf6a" : "#444455",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        flexShrink: 0,
                        transition: "all 0.15s",
                      }}
                    >
                      {copyStatus[field.key] ? "✓" : "⎘"}
                    </button>
                  </div>
                ))}
              </div>

              <div style={{
                marginTop: 20,
                padding: "12px 16px",
                background: "#0d0d18",
                border: "1px solid #1e1e2e",
                borderRadius: 6,
                fontSize: 11,
                color: "#444455",
                lineHeight: 1.7,
              }}>
                💡 <strong style={{ color: "#666680" }}>Tip:</strong> Click any field to edit before copying.
                Use <strong style={{ color: "#666680" }}>COPY ALL AS TSV</strong> to paste the entire row into Excel (paste into the first cell of a new row in your template).
                Verify URLs and confirm the DR Ticker / Ratio match your issuance plan.
              </div>
            </div>
          )}

          {/* Empty state */}
          {!results && !loading && !error && (
            <div style={{
              textAlign: "center",
              padding: "80px 0",
              color: "#222235",
            }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>⬡</div>
              <div style={{ fontSize: 13, letterSpacing: "0.1em" }}>
                ENTER A TICKER OR COMPANY NAME TO BEGIN
              </div>
              <div style={{ fontSize: 11, marginTop: 8, color: "#1e1e2e" }}>
                Works with NYSE · NASDAQ · TSE · HKEX · SSE · SZSE · Euronext
              </div>
            </div>
          )}
        </div>
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&display=swap');
        @keyframes pulse {
          0%, 100% { opacity: 0.2; transform: scale(0.8); }
          50% { opacity: 1; transform: scale(1.2); }
        }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #0a0a0f; }
        ::-webkit-scrollbar-thumb { background: #1e1e2e; border-radius: 3px; }
      `}</style>
    </div>
  );
}
