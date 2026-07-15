import { useState, useRef, useCallback, useEffect } from "react";

const API = "https://ai-for-payout-grid.onrender.com";

// ── Design system ────────────────────────────────────────────────────────────
// Subject: an operator's console for turning multi-sheet commission grids into
// certified, per-state insurance rule files. The page is built like the cover
// sheet of a regulatory filing — a sealed instrument, not a SaaS dashboard.
//
// Color    base #1A1D22 · surface #262A31 · recessed #22262D · hairline #363C46
//          ink #ECE8E0 · brass #C9A227 / #E0BB4A (hot) · ledger-green #4A8772
//          oxide #A14B4B
// Type     Fraunces (display serif, used once, large) · IBM Plex Sans (UI)
//          IBM Plex Mono (every number, filename, code — tabular-nums)
// Motif    a full wax-seal mark that draws itself in on load (stroke reveal),
//          then becomes the live status indicator — the literal certification
//          mark this tool stamps onto its output

const STATES_INDIA = [
  "ANDAMAN ISLANDS","ANDHRA PRADESH","ARUNACHAL PRADESH","ASSAM","BIHAR",
  "CHANDIGARH","CHHATTISGARH","DADRA AND NAGAR HAVELI","DAMAN AND DIU",
  "DELHI","GOA","GUJARAT","HARYANA","HIMACHAL PRADESH","JAMMU KASHMIR",
  "JHARKHAND","KARNATAKA","KERALA","LADAKH","LAKSHADWEEP","MADHYA PRADESH",
  "MAHARASHTRA","MANIPUR","MEGHALAYA","MIZORAM","NAGALAND","ODISHA",
  "PUDUCHERRY","PUNJAB","RAJASTHAN","SIKKIM","TAMIL NADU","TELANGANA",
  "TRIPURA","UTTAR PRADESH","UTTARAKHAND","WEST BENGAL",
];

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function today(offset = 0) {
  const d = new Date();
  d.setMonth(d.getMonth() + offset);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}-01`;
}

function lastDayOf(dateStr) {
  if (!dateStr) return "";
  const [y, m] = dateStr.split("-").map(Number);
  return new Date(y, m, 0).toISOString().slice(0, 10);
}

function timeNow() {
  return new Date().toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
}

// ── Seal ──────────────────────────────────────────────────────────────────────
// The signature element. An SVG ring + crest that draws itself in once on
// mount (stroke-dashoffset reveal), then drives all status color downstream.

function Seal({ status, size = 56 }) {
  const color = {
    idle: "#565D67",
    loading: "#C9A227",
    success: "#4A8772",
    error: "#A14B4B",
  }[status];

  const r = size / 2 - 4;
  const circ = 2 * Math.PI * r;

  return (
    <div style={{ width: size, height: size, position: "relative", flexShrink: 0 }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ position: "absolute", inset: 0 }}>
        <circle
          cx={size / 2} cy={size / 2} r={r}
          fill="none" stroke="#363C46" strokeWidth="1"
        />
        <circle
          cx={size / 2} cy={size / 2} r={r}
          fill="none" stroke={color} strokeWidth="1.25"
          strokeDasharray={circ}
          strokeDashoffset={0}
          className="seal-ring"
          style={{ "--circ": circ }}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        {/* tick marks like a dial */}
        {Array.from({ length: 12 }).map((_, i) => {
          const a = (i / 12) * Math.PI * 2;
          const x1 = size / 2 + Math.cos(a) * (r - 2);
          const y1 = size / 2 + Math.sin(a) * (r - 2);
          const x2 = size / 2 + Math.cos(a) * (r - 5);
          const y2 = size / 2 + Math.sin(a) * (r - 5);
          return <line key={i} x1={x1} y1={y1} x2={x2} y2={y2} stroke="#40464F" strokeWidth="1" />;
        })}
        <circle
          cx={size / 2} cy={size / 2} r={2.4}
          fill={color}
          className={status === "loading" ? "seal-pulse" : ""}
        />
      </svg>
      {status === "loading" && (
        <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ position: "absolute", inset: 0, animation: "seal-spin 4s linear infinite" }}>
          <circle
            cx={size / 2} cy={size / 2} r={r - 8}
            fill="none" stroke={color} strokeWidth="1"
            strokeDasharray={`${(2 * Math.PI * (r - 8)) / 4} ${(2 * Math.PI * (r - 8)) * 0.75 / 4}`}
            opacity="0.6"
          />
        </svg>
      )}
    </div>
  );
}

// ── Sub-components ───────────────────────────────────────────────────────────

function UploadZone({ file, onFile, disabled }) {
  const [dragging, setDragging] = useState(false);
  const inp = useRef();

  const handle = (f) => {
    if (f && f.name.endsWith(".xlsx")) onFile(f);
  };

  return (
    <div
      onClick={() => !disabled && inp.current.click()}
      onDragOver={(e) => { e.preventDefault(); if (!disabled) setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault(); setDragging(false);
        if (!disabled) handle(e.dataTransfer.files[0]);
      }}
      style={{
        border: `1px solid ${dragging ? "#C9A227" : file ? "#3D5F52" : "#363C46"}`,
        borderRadius: 4,
        padding: "26px 22px",
        cursor: disabled ? "not-allowed" : "pointer",
        background: dragging
          ? "linear-gradient(180deg, rgba(201,162,39,0.08), rgba(201,162,39,0.02))"
          : file
          ? "rgba(74,135,114,0.05)"
          : "#22262D",
        transition: "border-color 0.15s, background 0.15s",
        position: "relative",
        opacity: disabled ? 0.5 : 1,
        display: "flex", alignItems: "center", gap: 16,
        boxShadow: "inset 0 1px 2px rgba(0,0,0,0.4)",
      }}
    >
      <input ref={inp} type="file" accept=".xlsx" style={{ display: "none" }}
        onChange={(e) => handle(e.target.files[0])} disabled={disabled} />

      <div style={{
        width: 38, height: 38, borderRadius: 3, flexShrink: 0,
        border: `1px solid ${file ? "#3D5F52" : "#40464F"}`,
        display: "flex", alignItems: "center", justifyContent: "center",
        background: "#262A31",
      }}>
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
          {file ? (
            <path d="M3 8.5L6.5 12L13 4" stroke="#4A8772" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
          ) : (
            <path d="M8 2.5V11M8 2.5L4.5 6M8 2.5L11.5 6M3 13.5H13" stroke="#6B7178" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
          )}
        </svg>
      </div>

      <div style={{ flex: 1, minWidth: 0 }}>
        {file ? (
          <>
            <div style={{ color: "#ECE8E0", fontSize: 13.5, fontFamily: "'IBM Plex Mono', monospace", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {file.name}
            </div>
            <div style={{ color: "#5C6168", fontSize: 11.5, marginTop: 3, fontFamily: "'IBM Plex Mono', monospace" }}>
              {formatBytes(file.size)} received, click to replace
            </div>
          </>
        ) : (
          <>
            <div style={{ color: "#B8BCC2", fontSize: 13.5 }}>
              Drop the commission grid workbook
            </div>
            <div style={{ color: "#4A4F56", fontSize: 11.5, marginTop: 3, fontFamily: "'IBM Plex Mono', monospace" }}>
              .XLSX · TW 1+5, 1+1 / SATP, SAOD, RTO, 5+5 GRID SHEETS
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function StateSelector({ available, selected, onChange }) {
  const all = available.length > 0 ? available : STATES_INDIA;
  const allSelected = selected.length === all.length;

  const toggle = (s) => {
    if (selected.includes(s)) onChange(selected.filter(x => x !== s));
    else onChange([...selected, s]);
  };

  const toggleAll = () => {
    if (allSelected) onChange([]);
    else onChange([...all]);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <span style={{ color: "#5C6168", fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", fontVariantNumeric: "tabular-nums" }}>
          {selected.length === 0 ? `ALL ${all.length} STATES` : `${selected.length} / ${all.length} SELECTED`}
        </span>
        <button onClick={toggleAll} style={{
          background: "none", border: "1px solid #363C46", borderRadius: 3,
          color: "#9BA0A6", fontSize: 10.5, padding: "4px 10px", cursor: "pointer",
          fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.04em",
        }}>
          {allSelected ? "CLEAR" : "SELECT ALL"}
        </button>
      </div>
      <div style={{
        display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(186px, 1fr))",
        border: "1px solid #363C46", borderRadius: 4, overflow: "hidden",
        maxHeight: 322, overflowY: "auto",
        background: "#22262D",
      }}>
        {all.map((s) => {
          const on = selected.includes(s);
          return (
            <button key={s} onClick={() => toggle(s)} style={{
              background: on ? "rgba(201,162,39,0.09)" : "transparent",
              border: "none",
              borderBottom: "1px solid #2C313A",
              borderRight: "1px solid #2C313A",
              color: on ? "#D9B546" : "#7A7F86",
              fontSize: 11.5, padding: "8px 10px", cursor: "pointer",
              textAlign: "left", fontWeight: on ? 600 : 400,
              fontFamily: "'IBM Plex Mono', monospace",
              transition: "background 0.1s, color 0.1s",
              display: "flex", alignItems: "center", gap: 7,
            }}>
              <span style={{
                width: 5, height: 5, borderRadius: "50%", flexShrink: 0,
                background: on ? "#D9B546" : "#454B55",
              }} />
              {s}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function OutputModeSelector({ mode, onChange }) {
  const options = [
    { key: "per_state", label: "PER STATE", note: "one file per state" },
    { key: "combined", label: "COMBINED", note: "one file, all states" },
    { key: "both", label: "BOTH", note: "per-state + combined" },
  ];
  return (
    <div style={{
      display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8,
      marginBottom: 16,
    }}>
      {options.map(o => {
        const on = mode === o.key;
        return (
          <button key={o.key} onClick={() => onChange(o.key)} style={{
            background: on ? "rgba(201,162,39,0.09)" : "#22262D",
            border: `1px solid ${on ? "#C9A227" : "#363C46"}`,
            borderRadius: 4, padding: "10px 8px",
            cursor: "pointer", textAlign: "left",
            transition: "border-color 0.15s, background 0.15s",
          }}>
            <div style={{
              color: on ? "#D9B546" : "#B8BCC2", fontSize: 11.5, fontWeight: 700,
              fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.04em",
              display: "flex", alignItems: "center", gap: 6,
            }}>
              <span style={{
                width: 5, height: 5, borderRadius: "50%", flexShrink: 0,
                background: on ? "#D9B546" : "#454B55",
              }} />
              {o.label}
            </div>
            <div style={{
              color: "#5C6168", fontSize: 10.5, marginTop: 4,
              fontFamily: "'IBM Plex Mono', monospace",
            }}>
              {o.note}
            </div>
          </button>
        );
      })}
    </div>
  );
}

function ProgressRail({ log }) {
  if (!log || log.length === 0) return null;
  const last = log[log.length - 1];
  const pct = last.total > 0 ? Math.round((last.current / last.total) * 100) : 0;
  const done = last.message === "Complete!";

  return (
    <div style={{ marginTop: 22 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
        <span style={{
          color: done ? "#4A8772" : "#C9A227", fontSize: 12, fontWeight: 600,
          fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.02em",
        }}>
          {done ? "COMPLETE" : last.message.toUpperCase()}
        </span>
        <span style={{ color: "#5C6168", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace", fontVariantNumeric: "tabular-nums" }}>
          {pct}%
        </span>
      </div>
      <div style={{ background: "#2C313A", height: 2, borderRadius: 1, overflow: "hidden" }}>
        <div style={{
          height: "100%", width: `${pct}%`,
          background: done ? "#4A8772" : "linear-gradient(90deg, #C9A227, #E0BB4A)",
          transition: "width 0.5s cubic-bezier(.4,0,.2,1)",
        }} />
      </div>
      <div style={{ marginTop: 10 }}>
        {log.slice(-4).reverse().map((l, i) => (
          <div key={i} style={{
            color: "#454A51", fontSize: 11, marginBottom: 3,
            fontFamily: "'IBM Plex Mono', monospace", fontVariantNumeric: "tabular-nums",
            display: "flex", gap: 8,
          }}>
            <span style={{ opacity: 0.6 }}>{l.total > 0 ? `${l.current}/${l.total}` : "·"}</span>
            <span>{l.message}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function ResultManifest({ files, zipFile, sessionId }) {
  if (!files || files.length === 0) return null;
  const base = `${API}/api/download/${sessionId}`;

  return (
    <div>
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "flex-end",
        marginBottom: 20, paddingBottom: 18, borderBottom: "1px solid #363C46",
      }}>
        <div>
          <div style={{ color: "#5C6168", fontSize: 10.5, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.08em", marginBottom: 6 }}>
            RUN COMPLETE
          </div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 12 }}>
            <span style={{
              fontFamily: "'Fraunces', serif", color: "#ECE8E0", fontWeight: 600, fontSize: 28,
              fontVariantNumeric: "tabular-nums", letterSpacing: "-0.01em",
            }}>
              {files.length}
            </span>
            <span style={{ color: "#B8BCC2", fontSize: 14.5 }}>
              {files.length === 1 ? "file certified" : "files certified"}
            </span>
            <span style={{ color: "#454A51", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" }}>
              {zipFile && formatBytes(zipFile.size_bytes)}
            </span>
          </div>
        </div>
        {zipFile && (
          <a href={`${base}/${zipFile.filename}`} download style={{
            background: "#C9A227", color: "#22262D", borderRadius: 3,
            padding: "11px 20px", fontWeight: 700, fontSize: 12,
            textDecoration: "none", display: "flex", alignItems: "center", gap: 8,
            fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.03em",
            boxShadow: "0 1px 0 rgba(255,255,255,0.15) inset",
          }}>
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
              <path d="M6 1.5V8M6 8L3 5M6 8L9 5M2 10.5H10" stroke="#22262D" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            DOWNLOAD ALL
          </a>
        )}
      </div>

      <div style={{ border: "1px solid #363C46", borderRadius: 4, overflow: "hidden" }}>
        {files.map((f, i) => (
          <div key={i} style={{
            display: "flex", justifyContent: "space-between", alignItems: "center",
            padding: "11px 16px",
            borderBottom: i < files.length - 1 ? "1px solid #2C313A" : "none",
            background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.012)",
          }}
            onMouseEnter={e => e.currentTarget.style.background = "rgba(201,162,39,0.04)"}
            onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.012)"}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <span style={{ color: "#3A3E44", fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", width: 22, fontVariantNumeric: "tabular-nums" }}>
                {String(i + 1).padStart(2, "0")}
              </span>
              <span style={{ color: "#ECE8E0", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>
                {f.filename}
              </span>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
              <span style={{ color: "#454A51", fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", fontVariantNumeric: "tabular-nums" }}>
                {formatBytes(f.size_bytes)}
              </span>
              <a href={`${base}/${f.filename}`} download style={{
                color: "#C9A227", fontSize: 11.5, textDecoration: "none", fontWeight: 600,
                fontFamily: "'IBM Plex Mono', monospace",
              }}>
                ↓ GET
              </a>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────

const LC_OPTIONS = [
  { id: "digit", label: "DIGIT", sheetHint: "2W RTO's, TW 1+5, TW 1+1 & SATP, TW SAOD…" },
  { id: "tata", label: "TATA AIG", sheetHint: "TW sheet of the Standard Grid Communication workbook" },
];

export default function App() {
  const [lc, setLc] = useState("digit"); // digit | tata
  const [file, setFile] = useState(null);
  const [effStart, setEffStart] = useState(today());
  const [effEnd, setEffEnd] = useState(lastDayOf(today()));
  const [availableStates, setAvailableStates] = useState([]);
  const [selectedStates, setSelectedStates] = useState([]);
  const [statesLoading, setStatesLoading] = useState(false);
  const [sessionId, setSessionId] = useState(null);
  const [outputMode, setOutputMode] = useState("per_state"); // per_state | combined | both

  const [status, setStatus] = useState("idle"); // idle | loading | success | error
  const [progress, setProgress] = useState([]);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [mounted, setMounted] = useState(false);

  useEffect(() => { const t = setTimeout(() => setMounted(true), 60); return () => clearTimeout(t); }, []);

  const handleFile = useCallback(async (f) => {
    setFile(f);
    setAvailableStates([]);
    setSelectedStates([]);
    setStatesLoading(true);
    setResult(null);
    setError(null);
    setStatus("idle");

    try {
      const fd = new FormData();
      fd.append("file", f);
      fd.append("lc", lc);
      const res = await fetch(`${API}/api/states`, { method: "POST", body: fd });
      const data = await res.json();
      if (data.states) {
        setAvailableStates(data.states);
      }
    } catch (e) {
      // silently fall back to static list
    } finally {
      setStatesLoading(false);
    }
  }, [lc]);

  const handleLcChange = (nextLc) => {
    if (nextLc === lc || processing) return;
    setLc(nextLc);
    setFile(null);
    setAvailableStates([]);
    setSelectedStates([]);
    setResult(null);
    setError(null);
    setStatus("idle");
  };

  const handleStartDate = (v) => {
    setEffStart(v);
    setEffEnd(lastDayOf(v));
  };

  const handleSubmit = async () => {
    if (!file) { setError("Please upload an Excel file first."); return; }
    if (!effStart || !effEnd) { setError("Set both effect dates."); return; }

    setStatus("loading");
    setProgress([]);
    setResult(null);
    setError(null);

    const fd = new FormData();
    fd.append("file", file);
    fd.append("lc", lc);
    fd.append("effect_start", effStart);
    fd.append("effect_end", effEnd);
    if (selectedStates.length > 0)
      fd.append("states", JSON.stringify(selectedStates));
    fd.append("output_mode", outputMode);

    let tick = 0;
    const ticker = setInterval(() => {
      tick++;
      setProgress(p => [...p, { message: `Processing… (${tick}s)`, current: tick, total: 0 }]);
    }, 1200);

    try {
      const res = await fetch(`${API}/api/process`, { method: "POST", body: fd });
      clearInterval(ticker);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Server error");
      setProgress(data.progress || []);
      setSessionId(data.session_id);
      setResult(data);
      setStatus("success");
    } catch (e) {
      clearInterval(ticker);
      setError(e.message);
      setStatus("error");
    }
  };

  const canSubmit = !!file && !!effStart && !!effEnd && status !== "loading";
  const processing = status === "loading";
  const stateCount = (availableStates.length > 0 ? availableStates : STATES_INDIA).length;

  return (
    <div style={{
      minHeight: "100vh",
      background: "#1A1D22",
      fontFamily: "'IBM Plex Sans', system-ui, sans-serif",
      color: "#9BA0A6",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600;700&display=swap');

        @keyframes seal-spin { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
        @keyframes seal-pulse-kf { 0%,100%{opacity:1} 50%{opacity:0.35} }
        @keyframes fade-up { from{opacity:0; transform:translateY(8px)} to{opacity:1; transform:translateY(0)} }
        @keyframes draw-ring { from{ stroke-dashoffset: var(--circ); } to{ stroke-dashoffset: 0; } }

        .seal-pulse { animation: seal-pulse-kf 1.6s ease-in-out infinite; transform-origin: center; }
        .seal-ring {
          stroke-dashoffset: var(--circ);
          animation: draw-ring 1.1s cubic-bezier(.4,0,.2,1) 0.15s forwards;
        }
        .reveal { opacity: 0; animation: fade-up 0.5s cubic-bezier(.4,0,.2,1) forwards; }

        ::-webkit-scrollbar{width:4px;height:4px}
        ::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:#40464F;border-radius:2px}
        input[type=date]{color-scheme:dark}

        @media (max-width: 860px) {
          .console-grid { grid-template-columns: 1fr !important; }
          .hero-stats { display: none !important; }
        }
        @media (prefers-reduced-motion: reduce) {
          *{animation: none !important}
          .seal-ring { stroke-dashoffset: 0 !important; }
        }
        button:focus-visible, a:focus-visible, input:focus-visible {
          outline: 1.5px solid #C9A227;
          outline-offset: 2px;
        }
        input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.6); cursor: pointer; }
      `}</style>

      {/* ── Hero / masthead ───────────────────────────────────────────────── */}
      <div style={{
        borderBottom: "1px solid #363C46",
        background: "radial-gradient(ellipse 900px 400px at 15% -10%, rgba(201,162,39,0.06), transparent)",
      }}>
        <div style={{
          maxWidth: 1080, margin: "0 auto", padding: "44px 28px 36px",
          display: "flex", justifyContent: "space-between", alignItems: "flex-end", gap: 24,
        }}>
          <div className="reveal" style={{ display: "flex", alignItems: "center", gap: 22 }}>
            <Seal status={status} size={58} />
            <div>
              <div style={{
                fontSize: 10.5, color: "#C9A227", fontFamily: "'IBM Plex Mono', monospace",
                letterSpacing: "0.14em", marginBottom: 9, fontWeight: 600,
              }}>
                {lc === "tata" ? "TATA AIG" : "DIGIT"} TWO-WHEELER · INSURANCE OPERATIONS
              </div>
              <h1 style={{
                fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 34,
                color: "#ECE8E0", letterSpacing: "-0.015em", lineHeight: 1.05,
                margin: 0,
              }}>
                Rule Engine Console
              </h1>
              <div style={{
                fontSize: 13, color: "#6B7178", marginTop: 9, maxWidth: 420, lineHeight: 1.5,
              }}>
                Upload a broker commission grid. Receive certified, state-wise payout rule files, one per jurisdiction.
              </div>
            </div>
          </div>

          <div className="hero-stats reveal" style={{
            display: "flex", gap: 28, paddingBottom: 4,
            animationDelay: "0.1s",
          }}>
            <StatBlock label="States covered" value={String(stateCount)} />
            <StatBlock label="Output columns" value={lc === "tata" ? "45" : "50"} />
            <StatBlock label="Status" value={processing ? "RUNNING" : status === "success" ? "DONE" : status === "error" ? "FAILED" : "STANDBY"}
              tone={processing ? "#C9A227" : status === "success" ? "#4A8772" : status === "error" ? "#A14B4B" : "#6B7178"} />
          </div>
        </div>
      </div>

      {/* ── Console body ──────────────────────────────────────────────────── */}
      <div className="console-grid" style={{
        maxWidth: 1080, margin: "0 auto", padding: "0 28px",
        display: "grid", gridTemplateColumns: "1.1fr 0.9fr",
      }}>
        {/* LEFT spine */}
        <div style={{ borderRight: "1px solid #363C46", paddingRight: 32 }}>

          <Step n="00" label="Insurer (LC)" delay="0s">
            <div style={{ display: "flex", gap: 8 }}>
              {LC_OPTIONS.map(opt => {
                const active = lc === opt.id;
                return (
                  <button
                    key={opt.id}
                    onClick={() => handleLcChange(opt.id)}
                    disabled={processing}
                    style={{
                      flex: 1, padding: "11px 0", borderRadius: 3,
                      fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600,
                      fontSize: 12, letterSpacing: "0.04em", cursor: processing ? "not-allowed" : "pointer",
                      background: active ? "linear-gradient(180deg, #D9B546, #C9A227)" : "#22262D",
                      color: active ? "#22262D" : "#8B9096",
                      border: active ? "1px solid #E0BB4A" : "1px solid #363C46",
                    }}
                  >
                    {opt.label}
                  </button>
                );
              })}
            </div>
            <div style={{
              color: "#454A51", fontSize: 11, marginTop: 9, lineHeight: 1.5,
              fontFamily: "'IBM Plex Mono', monospace",
            }}>
              Expects: {LC_OPTIONS.find(o => o.id === lc)?.sheetHint}
            </div>
          </Step>

          <Step n="01" label="Source file" delay="0.05s">
            <UploadZone file={file} onFile={handleFile} disabled={processing} />
            {statesLoading && (
              <div style={{
                color: "#6B7178", fontSize: 11, marginTop: 10,
                fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.03em",
                display: "flex", alignItems: "center", gap: 8,
              }}>
                <span style={{ width: 5, height: 5, borderRadius: "50%", background: "#C9A227", animation: "seal-pulse-kf 1.2s infinite" }} />
                READING STATE LIST FROM WORKBOOK…
              </div>
            )}
          </Step>

          <Step n="02" label="Effect period" delay="0.1s">
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <div>
                <label style={labelStyle}>Start date</label>
                <input type="date" value={effStart}
                  onChange={e => handleStartDate(e.target.value)}
                  disabled={processing}
                  style={inputStyle} />
              </div>
              <div>
                <label style={labelStyle}>End date</label>
                <input type="date" value={effEnd}
                  onChange={e => setEffEnd(e.target.value)}
                  disabled={processing}
                  style={inputStyle} />
              </div>
            </div>
            {effStart && effEnd && (
              <div style={{
                color: "#454A51", fontSize: 11.5, marginTop: 10,
                fontFamily: "'IBM Plex Mono', monospace", fontVariantNumeric: "tabular-nums",
              }}>
                {Math.round((new Date(effEnd) - new Date(effStart)) / 86400000) + 1} days of rule coverage
              </div>
            )}
          </Step>

          <div className="reveal" style={{ padding: "26px 0 36px", animationDelay: "0.2s" }}>
            <button
              onClick={handleSubmit}
              disabled={!canSubmit}
              style={{
                background: canSubmit ? "linear-gradient(180deg, #D9B546, #C9A227)" : "#262A31",
                color: canSubmit ? "#22262D" : "#454A51",
                border: canSubmit ? "1px solid #E0BB4A" : "1px solid #363C46",
                borderRadius: 4, padding: "15px 0",
                fontWeight: 700, fontSize: 13, cursor: canSubmit ? "pointer" : "not-allowed",
                width: "100%",
                fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.05em",
                transition: "filter 0.15s",
                boxShadow: canSubmit ? "0 4px 14px rgba(201,162,39,0.18)" : "none",
              }}
              onMouseEnter={e => canSubmit && (e.currentTarget.style.filter = "brightness(1.06)")}
              onMouseLeave={e => canSubmit && (e.currentTarget.style.filter = "brightness(1)")}
            >
              {processing ? "GENERATING FILES…" : "GENERATE RULE FILES →"}
            </button>

            {error && (
              <div style={{
                background: "rgba(161,75,75,0.08)", border: "1px solid rgba(161,75,75,0.3)",
                borderRadius: 4, padding: "12px 16px", color: "#C98888", fontSize: 12.5,
                marginTop: 14, fontFamily: "'IBM Plex Mono', monospace",
              }}>
                ERROR: {error}
              </div>
            )}

            {processing && <ProgressRail log={progress} />}
            {status === "success" && <ProgressRail log={[{ message: "Complete!", current: 1, total: 1 }]} />}
          </div>
        </div>

        {/* RIGHT: State picker */}
        <div style={{ paddingLeft: 32 }}>
          <Step n="03" label="State filter" note={selectedStates.length > 0 ? `${selectedStates.length} selected` : "all states"} delay="0.15s">
            <div style={{ color: "#454A51", fontSize: 11.5, marginBottom: 14, fontFamily: "'IBM Plex Mono', monospace" }}>
              LEAVE BLANK TO PROCESS EVERY STATE IN THE FILE
            </div>
            <OutputModeSelector mode={outputMode} onChange={setOutputMode} />
            <StateSelector
              available={availableStates}
              selected={selectedStates}
              onChange={setSelectedStates}
            />
          </Step>
        </div>
      </div>

      {/* ── Results ───────────────────────────────────────────────────────── */}
      {result && (
        <div style={{ maxWidth: 1080, margin: "0 auto", padding: "36px 28px 64px" }}>
          <div style={{ borderTop: "1px solid #363C46", paddingTop: 36 }}>
            <ResultManifest
              files={result.files}
              zipFile={result.zip}
              sessionId={result.session_id}
            />
            {result.progress && result.progress.length > 0 && (
              <details style={{ marginTop: 22 }}>
                <summary style={{
                  color: "#454A51", fontSize: 11.5, cursor: "pointer",
                  fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.03em",
                }}>
                  FULL PROCESSING LOG ({result.progress.length} EVENTS)
                </summary>
                <div style={{
                  background: "#22262D", border: "1px solid #363C46", borderRadius: 4,
                  padding: 14, marginTop: 10, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11,
                  color: "#6B7178", maxHeight: 220, overflowY: "auto",
                }}>
                  {result.progress.map((l, i) => (
                    <div key={i} style={{ marginBottom: 2 }}>[{l.current}/{l.total}] {l.message}</div>
                  ))}
                </div>
              </details>
            )}
          </div>
        </div>
      )}

      {/* ── Footer ────────────────────────────────────────────────────────── */}
      <div style={{
        borderTop: "1px solid #363C46", padding: "18px 28px",
        display: "flex", justifyContent: "space-between", alignItems: "center",
        maxWidth: 1080, margin: "0 auto",
        fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#4A5159",
        letterSpacing: "0.04em",
      }}>
        <span>{lc === "tata" ? "TATA AIG" : "DIGIT"} 2W RULE ENGINE</span>
        <span>{lc === "tata" ? "45" : "50"}-COLUMN STATE GRID OUTPUT</span>
      </div>
    </div>
  );
}

// ── Shared UI ─────────────────────────────────────────────────────────────────

function StatBlock({ label, value, tone = "#ECE8E0" }) {
  return (
    <div>
      <div style={{
        fontFamily: "'Fraunces', serif", fontSize: 21, fontWeight: 600,
        color: tone, fontVariantNumeric: "tabular-nums", letterSpacing: "-0.01em",
      }}>
        {value}
      </div>
      <div style={{
        fontSize: 9.5, color: "#454A51", fontFamily: "'IBM Plex Mono', monospace",
        letterSpacing: "0.08em", marginTop: 4, textTransform: "uppercase",
      }}>
        {label}
      </div>
    </div>
  );
}

function Step({ n, label, note, delay = "0s", children }) {
  return (
    <div className="reveal" style={{ padding: "30px 0", borderBottom: "1px solid #2C313A", animationDelay: delay }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: 11, marginBottom: 18 }}>
        <span style={{
          fontFamily: "'IBM Plex Mono', monospace", color: "#C9A227", fontWeight: 600, fontSize: 11,
          letterSpacing: "0.04em",
        }}>
          {n}
        </span>
        <span style={{
          color: "#ECE8E0", fontWeight: 600, fontSize: 14, letterSpacing: "0.01em",
        }}>
          {label}
        </span>
        <span style={{ flex: 1, height: 1, background: "#2C313A" }} />
        {note && (
          <span style={{
            color: "#5C6168", fontSize: 10.5, fontFamily: "'IBM Plex Mono', monospace",
            letterSpacing: "0.04em",
          }}>
            {note.toUpperCase()}
          </span>
        )}
      </div>
      {children}
    </div>
  );
}

const labelStyle = {
  display: "block", color: "#5C6168", fontSize: 10,
  fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em",
  marginBottom: 7, fontFamily: "'IBM Plex Mono', monospace",
};

const inputStyle = {
  width: "100%", boxSizing: "border-box",
  background: "#22262D", border: "1px solid #363C46",
  borderRadius: 3, padding: "10px 11px",
  color: "#ECE8E0", fontSize: 13, outline: "none",
  fontFamily: "'IBM Plex Mono', monospace",
  boxShadow: "inset 0 1px 2px rgba(0,0,0,0.4)",
};
