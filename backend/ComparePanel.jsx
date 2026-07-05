/**
 * ComparePanel.jsx
 * Drop this file into your src/ folder alongside App.jsx
 * Import it in App.jsx: import ComparePanel from "./ComparePanel";
 * Then render <ComparePanel /> wherever you want the section.
 */

import { useState, useRef } from "react";

const API_BASE = import.meta.env.VITE_API_BASE || "http://localhost:5050";

// ── colour tokens (matching ledger/docket identity) ─────────────
const C = {
  bg:        "#0f1117",
  surface:   "#161b27",
  border:    "#252d3d",
  borderLit: "#3a4560",
  brass:     "#c9a84c",
  brassLow:  "#8a6f2e",
  ink:       "#e8eaf0",
  inkMid:    "#8b93a8",
  inkDim:    "#555f75",
  increase:  "#3ecf8e",
  decrease:  "#f87171",
  new_:      "#60a5fa",
  removed:   "#a78bfa",
  neutral:   "#8b93a8",
};

const TAG_STYLE = {
  increase: { background: "#0d2b1e", color: C.increase,  border: `1px solid #1e5c3a` },
  decrease: { background: "#2d1010", color: C.decrease,  border: `1px solid #6b2020` },
  new:      { background: "#0d1e36", color: C.new_,      border: `1px solid #1e3a6b` },
  removed:  { background: "#1e1030", color: C.removed,   border: `1px solid #4a2a70` },
  unchanged:{ background: "#1a1d27", color: C.inkMid,    border: `1px solid ${C.border}` },
};

const ICON = {
  increase:  "↑",
  decrease:  "↓",
  new:       "+",
  removed:   "×",
  unchanged: "—",
};

// ── tiny helpers ─────────────────────────────────────────────────
function FileDrop({ label, file, onChange }) {
  const ref = useRef();
  const active = !!file;
  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={e => e.preventDefault()}
      onDrop={e => { e.preventDefault(); onChange(e.dataTransfer.files[0]); }}
      style={{
        flex: 1,
        border: `1.5px dashed ${active ? C.brass : C.border}`,
        borderRadius: 6,
        padding: "20px 16px",
        cursor: "pointer",
        background: active ? "#1a1608" : C.surface,
        textAlign: "center",
        transition: "border-color .2s, background .2s",
        minWidth: 0,
      }}
    >
      <input ref={ref} type="file" accept=".xlsx" style={{ display: "none" }}
             onChange={e => onChange(e.target.files[0])} />
      <div style={{ fontSize: 22, marginBottom: 6, color: active ? C.brass : C.inkDim }}>
        {active ? "📋" : "⬆"}
      </div>
      <div style={{ fontSize: 11, color: C.inkMid, letterSpacing: "0.08em",
                    textTransform: "uppercase", marginBottom: 4 }}>{label}</div>
      {active
        ? <div style={{ fontSize: 12, color: C.brass, fontFamily: "JetBrains Mono, monospace",
                        wordBreak: "break-all" }}>{file.name}</div>
        : <div style={{ fontSize: 11, color: C.inkDim }}>Drop .xlsx or click</div>}
    </div>
  );
}

function Pill({ type, children }) {
  const s = TAG_STYLE[type] || TAG_STYLE.unchanged;
  return (
    <span style={{
      ...s, borderRadius: 4, padding: "1px 7px",
      fontSize: 11, fontWeight: 600, letterSpacing: "0.05em",
      fontFamily: "JetBrains Mono, monospace", display: "inline-flex",
      alignItems: "center", gap: 4, whiteSpace: "nowrap",
    }}>
      {ICON[type]} {children}
    </span>
  );
}

function StateBadge({ state, type }) {
  const color = type === "increase" ? C.increase
              : type === "decrease" ? C.decrease
              : type === "new"      ? C.new_
              : type === "removed"  ? C.removed
              : C.inkMid;
  return (
    <span style={{
      display: "inline-block", background: C.surface, border: `1px solid ${C.border}`,
      borderRadius: 4, padding: "2px 8px", fontSize: 11, color,
      fontFamily: "JetBrains Mono, monospace", margin: "2px 3px",
    }}>
      {state}
    </span>
  );
}

// ── main component ────────────────────────────────────────────────
export default function ComparePanel() {
  const [filePrev, setFilePrev] = useState(null);
  const [fileCurr, setFileCurr] = useState(null);
  const [labelPrev, setLabelPrev] = useState("");
  const [labelCurr, setLabelCurr] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const [open, setOpen] = useState(true);

  async function handleCompare() {
    if (!filePrev || !fileCurr) {
      setError("Please upload both Excel files before comparing.");
      return;
    }
    setLoading(true);
    setError(null);
    setResult(null);

    const fd = new FormData();
    fd.append("file_prev", filePrev);
    fd.append("file_curr", fileCurr);
    fd.append("label_prev", labelPrev || "Previous Month");
    fd.append("label_curr", labelCurr || "Current Month");

    try {
      const res = await fetch(`${API_BASE}/api/compare`, { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Server error");
      setResult(data);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  const ins = result?.insights;

  return (
    <section style={{
      background: C.bg, border: `1px solid ${C.border}`,
      borderRadius: 8, marginBottom: 24, overflow: "hidden",
      fontFamily: "IBM Plex Sans, Inter, sans-serif",
    }}>
      {/* ── header bar ── */}
      <div
        onClick={() => setOpen(o => !o)}
        style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "12px 20px", cursor: "pointer",
          borderBottom: open ? `1px solid ${C.border}` : "none",
          background: C.surface,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{
            background: C.brass, color: "#0f1117", borderRadius: 4,
            padding: "2px 8px", fontSize: 10, fontWeight: 700,
            letterSpacing: "0.12em", textTransform: "uppercase",
            fontFamily: "JetBrains Mono, monospace",
          }}>AI</span>
          <span style={{ color: C.ink, fontWeight: 600, fontSize: 15 }}>
            Commission Grid Comparison
          </span>
          <span style={{ color: C.inkDim, fontSize: 12 }}>
            — detect rate changes across months
          </span>
        </div>
        <span style={{ color: C.inkDim, fontSize: 18 }}>{open ? "▾" : "▸"}</span>
      </div>

      {open && (
        <div style={{ padding: "20px 24px" }}>

          {/* ── file upload row ── */}
          <div style={{ display: "flex", gap: 12, marginBottom: 14 }}>
            <FileDrop label="Previous month" file={filePrev} onChange={setFilePrev} />
            <div style={{ display: "flex", alignItems: "center", color: C.inkDim, fontSize: 20 }}>→</div>
            <FileDrop label="Current month" file={fileCurr} onChange={setFileCurr} />
          </div>

          {/* ── label inputs ── */}
          <div style={{ display: "flex", gap: 12, marginBottom: 18 }}>
            {[
              { val: labelPrev, set: setLabelPrev, ph: "e.g. Jan 2026" },
              { val: labelCurr, set: setLabelCurr, ph: "e.g. Feb 2026" },
            ].map(({ val, set, ph }, i) => (
              <input key={i} value={val} onChange={e => set(e.target.value)}
                placeholder={ph}
                style={{
                  flex: 1, background: C.surface, border: `1px solid ${C.border}`,
                  borderRadius: 5, padding: "7px 12px", color: C.ink, fontSize: 12,
                  fontFamily: "JetBrains Mono, monospace", outline: "none",
                }}
              />
            ))}
          </div>

          {/* ── compare button ── */}
          <button
            onClick={handleCompare}
            disabled={loading || !filePrev || !fileCurr}
            style={{
              background: loading ? C.brassLow : C.brass,
              color: "#0f1117", border: "none", borderRadius: 5,
              padding: "9px 24px", fontWeight: 700, fontSize: 13,
              letterSpacing: "0.06em", cursor: loading ? "wait" : "pointer",
              fontFamily: "IBM Plex Sans, sans-serif",
              opacity: (!filePrev || !fileCurr) ? 0.5 : 1,
              transition: "background .2s",
            }}
          >
            {loading ? "Analysing…" : "Compare Grids"}
          </button>

          {/* ── error ── */}
          {error && (
            <div style={{
              marginTop: 14, background: "#2d1010", border: `1px solid #6b2020`,
              borderRadius: 6, padding: "10px 14px", color: C.decrease, fontSize: 13,
            }}>
              ⚠ {error}
            </div>
          )}

          {/* ── loading skeleton ── */}
          {loading && (
            <div style={{ marginTop: 20, color: C.inkMid, fontSize: 13 }}>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ animation: "spin 1s linear infinite", display: "inline-block" }}>⟳</span>
                Running LLM comparison across all sheets and states…
              </div>
              <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
            </div>
          )}

          {/* ── results ── */}
          {ins && (
            <div style={{ marginTop: 22, display: "flex", flexDirection: "column", gap: 16 }}>

              {/* divider */}
              <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 18 }}>
                <div style={{ fontSize: 10, color: C.brass, letterSpacing: "0.14em",
                              textTransform: "uppercase", fontFamily: "JetBrains Mono, monospace",
                              marginBottom: 8 }}>
                  {result.label_prev}  →  {result.label_curr}
                </div>
                <p style={{ color: C.ink, fontSize: 14, lineHeight: 1.6, margin: 0 }}>
                  {ins.summary}
                </p>
              </div>

              {/* stat strip */}
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                {[
                  { label: "Increased", count: ins.states_with_increases?.length || 0, type: "increase" },
                  { label: "Decreased", count: ins.states_with_decreases?.length || 0, type: "decrease" },
                  { label: "Added",     count: ins.states_added?.length || 0,           type: "new" },
                  { label: "Removed",   count: ins.states_removed?.length || 0,         type: "removed" },
                ].map(({ label, count, type }) => (
                  <div key={label} style={{
                    background: C.surface, border: `1px solid ${C.border}`,
                    borderRadius: 6, padding: "10px 16px", minWidth: 80, textAlign: "center",
                  }}>
                    <div style={{
                      fontSize: 22, fontWeight: 700, color:
                        type === "increase" ? C.increase :
                        type === "decrease" ? C.decrease :
                        type === "new"      ? C.new_     : C.removed,
                      fontFamily: "JetBrains Mono, monospace",
                    }}>{count}</div>
                    <div style={{ fontSize: 10, color: C.inkDim, textTransform: "uppercase",
                                  letterSpacing: "0.08em", marginTop: 2 }}>{label}</div>
                  </div>
                ))}
              </div>

              {/* highlight rows */}
              {ins.highlights?.length > 0 && (
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  <div style={{ fontSize: 10, color: C.inkDim, letterSpacing: "0.1em",
                                textTransform: "uppercase", fontFamily: "JetBrains Mono, monospace" }}>
                    Change Log
                  </div>
                  {ins.highlights.map((h, i) => (
                    <div key={i} style={{
                      display: "flex", alignItems: "flex-start", gap: 10,
                      background: C.surface, border: `1px solid ${C.border}`,
                      borderRadius: 5, padding: "8px 12px",
                    }}>
                      <Pill type={h.type}>{h.type}</Pill>
                      <div style={{ flex: 1 }}>
                        <span style={{ color: C.brass, fontWeight: 600, fontSize: 12,
                                       fontFamily: "JetBrains Mono, monospace",
                                       marginRight: 8 }}>{h.state}</span>
                        <span style={{ color: C.inkMid, fontSize: 13 }}>{h.detail}</span>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {/* state tag clouds */}
              {[
                { label: "States with rate increases", list: ins.states_with_increases, type: "increase" },
                { label: "States with rate decreases", list: ins.states_with_decreases, type: "decrease" },
                { label: "Newly added states",          list: ins.states_added,          type: "new" },
                { label: "Removed states",              list: ins.states_removed,         type: "removed" },
              ].filter(g => g.list?.length > 0).map(({ label, list, type }) => (
                <div key={label}>
                  <div style={{ fontSize: 10, color: C.inkDim, letterSpacing: "0.1em",
                                textTransform: "uppercase", marginBottom: 6,
                                fontFamily: "JetBrains Mono, monospace" }}>{label}</div>
                  <div style={{ display: "flex", flexWrap: "wrap" }}>
                    {list.map(s => <StateBadge key={s} state={s} type={type} />)}
                  </div>
                </div>
              ))}

              {/* biggest change */}
              {ins.biggest_change && (
                <div style={{
                  background: "#16120a", border: `1px solid ${C.brassLow}`,
                  borderRadius: 6, padding: "12px 16px",
                }}>
                  <div style={{ fontSize: 10, color: C.brass, letterSpacing: "0.12em",
                                textTransform: "uppercase", fontFamily: "JetBrains Mono, monospace",
                                marginBottom: 6 }}>Biggest Change</div>
                  <span style={{ color: C.brass, fontWeight: 700,
                                  fontFamily: "JetBrains Mono, monospace", marginRight: 8 }}>
                    {ins.biggest_change.state}
                  </span>
                  <span style={{ color: C.ink, fontSize: 13 }}>{ins.biggest_change.description}</span>
                </div>
              )}

              {/* recommendation */}
              {ins.recommendation && (
                <div style={{
                  background: "#0d1420", border: `1px solid #1e3a6b`,
                  borderRadius: 6, padding: "12px 16px",
                }}>
                  <div style={{ fontSize: 10, color: C.new_, letterSpacing: "0.12em",
                                textTransform: "uppercase", fontFamily: "JetBrains Mono, monospace",
                                marginBottom: 6 }}>Recommendation</div>
                  <span style={{ color: C.ink, fontSize: 13, lineHeight: 1.6 }}>
                    {ins.recommendation}
                  </span>
                </div>
              )}

            </div>
          )}
        </div>
      )}
    </section>
  );
}
