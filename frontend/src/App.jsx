import { useState, useRef, useCallback } from "react";

const API = "https://ai-for-payout-grid.onrender.com";

// ── Palette ──────────────────────────────────────────────────────────────────
// Deep slate bg, warm amber accent, ghost-white type, muted borders
// Signature: pulsing "processing cells" animation in the progress bar

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

// ── Sub-components ────────────────────────────────────────────────────────────

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
        border: `2px dashed ${dragging ? "#F59E0B" : file ? "#10B981" : "#334155"}`,
        borderRadius: 12,
        padding: "40px 32px",
        textAlign: "center",
        cursor: disabled ? "not-allowed" : "pointer",
        background: dragging ? "rgba(245,158,11,0.05)" : file ? "rgba(16,185,129,0.04)" : "rgba(15,23,42,0.4)",
        transition: "all 0.2s",
        position: "relative",
        opacity: disabled ? 0.5 : 1,
      }}
    >
      <input ref={inp} type="file" accept=".xlsx" style={{ display: "none" }}
        onChange={(e) => handle(e.target.files[0])} disabled={disabled} />

      {file ? (
        <>
          <div style={{ fontSize: 36, marginBottom: 8 }}>✅</div>
          <div style={{ color: "#10B981", fontWeight: 600, fontSize: 15 }}>{file.name}</div>
          <div style={{ color: "#64748B", fontSize: 13, marginTop: 4 }}>{formatBytes(file.size)} · click to replace</div>
        </>
      ) : (
        <>
          <div style={{ fontSize: 36, marginBottom: 10 }}>📂</div>
          <div style={{ color: "#94A3B8", fontSize: 15, fontWeight: 500 }}>Drop your Excel grid here</div>
          <div style={{ color: "#475569", fontSize: 13, marginTop: 4 }}>.xlsx files only</div>
        </>
      )}
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
        <span style={{ color: "#94A3B8", fontSize: 13 }}>
          {selected.length === 0 ? "All states will be processed" : `${selected.length} of ${all.length} selected`}
        </span>
        <button onClick={toggleAll} style={{
          background: "none", border: "1px solid #334155", borderRadius: 6,
          color: "#94A3B8", fontSize: 12, padding: "3px 10px", cursor: "pointer",
        }}>
          {allSelected ? "Deselect all" : "Select all"}
        </button>
      </div>
      <div style={{
        display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
        gap: 6, maxHeight: 260, overflowY: "auto", paddingRight: 4,
      }}>
        {all.map(s => {
          const on = selected.includes(s);
          return (
            <button key={s} onClick={() => toggle(s)} style={{
              background: on ? "rgba(245,158,11,0.12)" : "rgba(30,41,59,0.6)",
              border: `1px solid ${on ? "#F59E0B" : "#1E293B"}`,
              borderRadius: 7, color: on ? "#FCD34D" : "#64748B",
              fontSize: 12, padding: "6px 10px", cursor: "pointer",
              textAlign: "left", fontWeight: on ? 600 : 400,
              transition: "all 0.15s",
            }}>
              {on ? "✓ " : ""}{s}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function ProgressBar({ log }) {
  if (!log || log.length === 0) return null;
  const last = log[log.length - 1];
  const pct = last.total > 0 ? Math.round((last.current / last.total) * 100) : 0;
  const done = last.message === "Complete!";

  return (
    <div style={{ marginTop: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
        <span style={{ color: done ? "#10B981" : "#F59E0B", fontSize: 13, fontWeight: 600 }}>
          {last.message}
        </span>
        <span style={{ color: "#64748B", fontSize: 13 }}>{pct}%</span>
      </div>
      <div style={{ background: "#1E293B", borderRadius: 6, height: 8, overflow: "hidden" }}>
        <div style={{
          height: "100%", width: `${pct}%`,
          background: done ? "#10B981" : "linear-gradient(90deg, #F59E0B, #FBBF24)",
          borderRadius: 6,
          transition: "width 0.4s ease",
          boxShadow: done ? "none" : "0 0 8px #F59E0B88",
        }} />
      </div>
      <div style={{ marginTop: 8, maxHeight: 100, overflowY: "auto" }}>
        {log.slice(-5).reverse().map((l, i) => (
          <div key={i} style={{ color: "#475569", fontSize: 11, marginBottom: 2 }}>
            {l.total > 0 && `[${l.current}/${l.total}] `}{l.message}
          </div>
        ))}
      </div>
    </div>
  );
}

function ResultTable({ files, zipFile, sessionId }) {
  if (!files || files.length === 0) return null;
  const base = `${API}/api/download/${sessionId}`;

  return (
    <div style={{ marginTop: 28 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div>
          <span style={{ color: "#F8FAFC", fontWeight: 700, fontSize: 16 }}>
            {files.length} files generated
          </span>
          <span style={{ color: "#475569", fontSize: 13, marginLeft: 10 }}>
            {zipFile && formatBytes(zipFile.size_bytes)} total
          </span>
        </div>
        {zipFile && (
          <a href={`${base}/${zipFile.filename}`} download style={{
            background: "#F59E0B", color: "#0F172A", borderRadius: 8,
            padding: "8px 18px", fontWeight: 700, fontSize: 13,
            textDecoration: "none", display: "flex", alignItems: "center", gap: 6,
          }}>
            ⬇ Download all (.zip)
          </a>
        )}
      </div>

      <div style={{
        background: "#0F172A", border: "1px solid #1E293B", borderRadius: 10,
        overflow: "hidden",
      }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ background: "#1E293B" }}>
              <th style={th}>State File</th>
              <th style={{ ...th, textAlign: "right" }}>Size</th>
              <th style={{ ...th, textAlign: "center" }}>Download</th>
            </tr>
          </thead>
          <tbody>
            {files.map((f, i) => (
              <tr key={i} style={{ borderBottom: "1px solid #1E293B" }}
                onMouseEnter={e => e.currentTarget.style.background = "#0D1A2A"}
                onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                <td style={td}>
                  <span style={{ color: "#FCD34D", fontFamily: "monospace", fontSize: 13 }}>
                    {f.filename}
                  </span>
                </td>
                <td style={{ ...td, textAlign: "right", color: "#475569", fontSize: 12 }}>
                  {formatBytes(f.size_bytes)}
                </td>
                <td style={{ ...td, textAlign: "center" }}>
                  <a href={`${base}/${f.filename}`} download style={{
                    color: "#F59E0B", fontSize: 13, textDecoration: "none", fontWeight: 500,
                  }}>⬇</a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const th = {
  padding: "10px 16px", textAlign: "left",
  color: "#64748B", fontSize: 12, fontWeight: 600,
  letterSpacing: "0.05em", textTransform: "uppercase",
};
const td = { padding: "10px 16px", color: "#CBD5E1", fontSize: 13 };

// ── Main App ──────────────────────────────────────────────────────────────────

export default function App() {
  const [file, setFile] = useState(null);
  const [effStart, setEffStart] = useState(today());
  const [effEnd, setEffEnd] = useState(lastDayOf(today()));
  const [availableStates, setAvailableStates] = useState([]);
  const [selectedStates, setSelectedStates] = useState([]);
  const [statesLoading, setStatesLoading] = useState(false);
  const [sessionId, setSessionId] = useState(null);

  const [status, setStatus] = useState("idle"); // idle | loading | success | error
  const [progress, setProgress] = useState([]);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  // When file is chosen, auto-fetch states
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
  }, []);

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
    fd.append("effect_start", effStart);
    fd.append("effect_end", effEnd);
    if (selectedStates.length > 0)
      fd.append("states", JSON.stringify(selectedStates));

    // Simulate progress ticks while waiting (real progress comes in response)
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

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0A0F1C",
      fontFamily: "'Inter', system-ui, sans-serif",
      color: "#CBD5E1",
    }}>
      {/* Top bar */}
      <div style={{
        borderBottom: "1px solid #1E293B",
        padding: "0 32px",
        display: "flex", alignItems: "center", justifyContent: "space-between",
        height: 56,
        background: "#0D1525",
        position: "sticky", top: 0, zIndex: 10,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{
            width: 28, height: 28, borderRadius: 7,
            background: "linear-gradient(135deg, #F59E0B, #D97706)",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 14,
          }}>⚡</div>
          <span style={{ color: "#F8FAFC", fontWeight: 700, fontSize: 15, letterSpacing: "-0.01em" }}>
            Digit 2W Processor
          </span>
          <span style={{
            background: "#1E293B", color: "#64748B",
            fontSize: 11, padding: "2px 8px", borderRadius: 99, fontWeight: 500,
          }}>
            Commission Grid → Rule Engine
          </span>
        </div>
        <div style={{
          width: 8, height: 8, borderRadius: "50%",
          background: processing ? "#F59E0B" : "#10B981",
          boxShadow: processing ? "0 0 8px #F59E0B" : "0 0 6px #10B981",
          animation: processing ? "pulse 1s infinite" : "none",
        }} />
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:#0F172A}
        ::-webkit-scrollbar-thumb{background:#334155;border-radius:10px}
        input[type=date]{color-scheme:dark}
      `}</style>

      {/* Main layout */}
      <div style={{
        maxWidth: 920, margin: "0 auto", padding: "36px 24px",
        display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24,
      }}>
        {/* LEFT: Config */}
        <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>

          {/* Upload */}
          <Card label="1 — Source File">
            <UploadZone file={file} onFile={handleFile} disabled={processing} />
            {statesLoading && (
              <div style={{ color: "#64748B", fontSize: 12, marginTop: 8, textAlign: "center" }}>
                Reading available states…
              </div>
            )}
          </Card>

          {/* Dates */}
          <Card label="2 — Effect Period">
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <div>
                <label style={labelStyle}>Start Date</label>
                <input type="date" value={effStart}
                  onChange={e => handleStartDate(e.target.value)}
                  disabled={processing}
                  style={inputStyle} />
              </div>
              <div>
                <label style={labelStyle}>End Date</label>
                <input type="date" value={effEnd}
                  onChange={e => setEffEnd(e.target.value)}
                  disabled={processing}
                  style={inputStyle} />
              </div>
            </div>
            {effStart && effEnd && (
              <div style={{ color: "#475569", fontSize: 12, marginTop: 8 }}>
                Rules active for {Math.round((new Date(effEnd) - new Date(effStart)) / 86400000) + 1} days
              </div>
            )}
          </Card>

          {/* Submit */}
          <button
            onClick={handleSubmit}
            disabled={!canSubmit}
            style={{
              background: canSubmit
                ? "linear-gradient(135deg, #F59E0B, #D97706)"
                : "#1E293B",
              color: canSubmit ? "#0A0F1C" : "#475569",
              border: "none", borderRadius: 10, padding: "14px 0",
              fontWeight: 700, fontSize: 15, cursor: canSubmit ? "pointer" : "not-allowed",
              width: "100%",
              transition: "all 0.2s",
              boxShadow: canSubmit ? "0 4px 16px rgba(245,158,11,0.3)" : "none",
            }}
          >
            {processing ? "⏳ Generating files…" : "▶  Generate Rule Files"}
          </button>

          {error && (
            <div style={{
              background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.3)",
              borderRadius: 8, padding: "12px 16px", color: "#FCA5A5", fontSize: 13,
            }}>
              ⚠ {error}
            </div>
          )}

          {processing && <ProgressBar log={progress} />}
          {status === "success" && <ProgressBar log={[{ message: "Complete!", current: 1, total: 1 }]} />}
        </div>

        {/* RIGHT: State picker */}
        <div>
          <Card label="3 — State Filter" badge={selectedStates.length > 0 ? `${selectedStates.length} selected` : "all"}>
            <div style={{ color: "#475569", fontSize: 12, marginBottom: 12 }}>
              Leave all deselected to process every state in the file.
            </div>
            <StateSelector
              available={availableStates}
              selected={selectedStates}
              onChange={setSelectedStates}
            />
          </Card>
        </div>
      </div>

      {/* Results */}
      {result && (
        <div style={{ maxWidth: 920, margin: "0 auto", padding: "0 24px 48px" }}>
          <ResultTable
            files={result.files}
            zipFile={result.zip}
            sessionId={result.session_id}
          />
          {result.progress && result.progress.length > 0 && (
            <details style={{ marginTop: 16 }}>
              <summary style={{ color: "#475569", fontSize: 12, cursor: "pointer" }}>
                Processing log ({result.progress.length} events)
              </summary>
              <div style={{
                background: "#0D1525", border: "1px solid #1E293B", borderRadius: 8,
                padding: 12, marginTop: 8, fontFamily: "monospace", fontSize: 11,
                color: "#475569", maxHeight: 200, overflowY: "auto",
              }}>
                {result.progress.map((l, i) => (
                  <div key={i}>[{l.current}/{l.total}] {l.message}</div>
                ))}
              </div>
            </details>
          )}
        </div>
      )}
    </div>
  );
}

// ── Shared UI ─────────────────────────────────────────────────────────────────

function Card({ label, badge, children }) {
  return (
    <div style={{
      background: "#0D1525",
      border: "1px solid #1E293B",
      borderRadius: 12,
      padding: 20,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
        <span style={{ color: "#F8FAFC", fontWeight: 700, fontSize: 13, letterSpacing: "0.02em" }}>
          {label}
        </span>
        {badge && (
          <span style={{
            background: "rgba(245,158,11,0.12)", color: "#F59E0B",
            fontSize: 10, padding: "2px 8px", borderRadius: 99, fontWeight: 600,
            textTransform: "uppercase", letterSpacing: "0.06em",
          }}>{badge}</span>
        )}
      </div>
      {children}
    </div>
  );
}

const labelStyle = {
  display: "block", color: "#64748B", fontSize: 11,
  fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em",
  marginBottom: 6,
};

const inputStyle = {
  width: "100%", boxSizing: "border-box",
  background: "#0A0F1C", border: "1px solid #1E293B",
  borderRadius: 8, padding: "9px 12px",
  color: "#F8FAFC", fontSize: 14, outline: "none",
};
