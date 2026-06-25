import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const WASTE = 0.15;
const BACKEND = import.meta.env.VITE_BACKEND_URL || "";

const MATERIAL_COLORS = {
  "ACM Panel":              "#c8a030",
  "MCM Panel":              "#c8a030",
  "Fiber Cement Panel":     "#5a8a5a",
  "Fiber Cement Plank":     "#4a7a6a",
  "Nichiha Panel":          "#7a6aaa",
  "Aluminum Wall Panel":    "#6a99aa",
  "Perforated Metal Panel": "#aa7a5a",
  "Soffit Panel":           "#5a7aaa",
  "Return/Trim":            "#aa5a7a",
  "Other":                  "#7a7a7a",
};

const buildExcel = (projectName, legend, takeoffData) => {
  const wb = XLSX.utils.book_new();

  // Tab 1: By Elevation
  const rows1 = [
    ["EXTERIOR PANEL TAKEOFF — BY ELEVATION"],
    ["Project:", projectName || ""],
    ["Date:", new Date().toLocaleDateString()],
    ["Waste Factor:", "15%"],
    [],
    ["Sheet Ref", "Elevation", "Material ID", "Material Name", "Category", "Gross SF", "Openings SF", "Net SF", "Waste (15%)", "Adj SF", "Notes"]
  ];

  const byBuilding = {};
  takeoffData.forEach(e => {
    const b = e.building || "Building";
    if (!byBuilding[b]) byBuilding[b] = [];
    byBuilding[b].push(e);
  });

  Object.entries(byBuilding).forEach(([bld, elevs]) => {
    rows1.push([bld.toUpperCase()]);
    elevs.forEach(elev => {
      const elevTotal = (elev.zones || []).reduce((s, z) => s + (z.netArea || 0), 0);
      rows1.push(["", elev.title + " — Total: " + Math.round(elevTotal) + " SF net"]);
      (elev.zones || []).forEach(z => {
        const waste = (z.netArea || 0) * WASTE;
        const adj = (z.netArea || 0) + waste;
        rows1.push([elev.sheetRef || "", elev.title || "", z.materialId || "", z.materialName || "", z.category || "", Math.round(z.grossArea || 0), Math.round(z.totalOpeningArea || 0), Math.round(z.netArea || 0), Math.round(waste), Math.round(adj), z.description || ""]);
      });
      if ((elev.flags || []).filter(Boolean).length) rows1.push(["", "⚠ " + elev.flags.filter(Boolean).join(" | ")]);
      rows1.push([]);
    });
  });

  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  ws1["!cols"] = [12, 35, 12, 30, 20, 10, 12, 10, 10, 10, 30].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, "By Elevation");

  // Tab 2: Summary with pricing boxes
  const matTotals = {};
  takeoffData.forEach(e => (e.zones || []).forEach(z => {
    const key = (z.materialId || "") + "||" + (z.materialName || "") + "||" + (z.category || "");
    if (!matTotals[key]) matTotals[key] = { materialId: z.materialId, materialName: z.materialName, category: z.category || "Other", net: 0, adj: 0 };
    matTotals[key].net += z.netArea || 0;
    matTotals[key].adj += (z.netArea || 0) * 1.15;
  }));

  const rows2 = [
    ["EXTERIOR PANEL ESTIMATE — SUMMARY"],
    ["Project:", projectName || ""],
    ["Date:", new Date().toLocaleDateString()],
    [],
    ["No.", "Material ID", "Material Name", "Category", "Net SF", "Adj SF (+15%)", "$/SF (enter rate)", "Total $", "Notes"],
  ];

  let no = 1;
  const byCategory = {};
  Object.values(matTotals).forEach(m => {
    if (!byCategory[m.category]) byCategory[m.category] = [];
    byCategory[m.category].push(m);
  });

  let dataRow = 6;
  Object.entries(byCategory).forEach(([cat, mats]) => {
    rows2.push(["", "", cat.toUpperCase()]);
    dataRow++;
    mats.forEach(m => {
      const adj = Math.round(m.adj);
      const net = Math.round(m.net);
      rows2.push([no++, m.materialId || "", m.materialName || "", m.category || "", net, adj, "", "", "Net: " + net + " SF + 15% = " + adj + " SF"]);
      dataRow++;
    });
    rows2.push([]);
    dataRow++;
  });

  rows2.push(["", "", "", "TOTAL", "", Math.round(Object.values(matTotals).reduce((s, m) => s + m.adj, 0)), "", "", ""]);

  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2["!cols"] = [6, 12, 35, 22, 10, 12, 16, 14, 35].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, "Estimate Summary");

  // Tab 3: Legend
  const rows3 = [["EXTERIOR PANEL MATERIALS LEGEND"], [], ["Material ID", "Name", "Category", "Color/Finish", "Notes"]];
  legend.forEach(m => rows3.push([m.id, m.name, m.category, m.color || "", m.notes || ""]));
  const ws3 = XLSX.utils.aoa_to_sheet(rows3);
  ws3["!cols"] = [14, 35, 22, 20, 30].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws3, "Material Legend");

  return wb;
};

export default function BPSEstimator() {
  const [file, setFile] = useState(null);
  const [phase, setPhase] = useState("idle");
  const [log, setLog] = useState([]);
  const [progress, setProgress] = useState({ label: "", pct: 0 });
  const [results, setResults] = useState(null);
  const [errMsg, setErrMsg] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [jobId, setJobId] = useState(null);
  const fileRef = useRef();
  const logRef = useRef();
  const pollRef = useRef(null);
  const seenLogs = useRef(0);

  useEffect(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, [log]);

  // Cleanup polling on unmount
  useEffect(() => () => { if (pollRef.current) clearInterval(pollRef.current); }, []);

  const handleFile = f => {
    if (f?.type === "application/pdf") {
      setFile(f); setPhase("idle"); setResults(null);
      setLog([]); setErrMsg(""); setJobId(null);
      seenLogs.current = 0;
    }
  };

  const startPolling = useCallback((id) => {
    if (pollRef.current) clearInterval(pollRef.current);
    seenLogs.current = 0;

    pollRef.current = setInterval(async () => {
      try {
        const res = await fetch(BACKEND + "/status/" + id);
        const data = await res.json();

        // Add new log entries
        if (data.log && data.log.length > seenLogs.current) {
          const newLogs = data.log.slice(seenLogs.current);
          setLog(prev => [...prev, ...newLogs]);
          seenLogs.current = data.log.length;
        }

        if (data.progress) setProgress(data.progress);
        if (data.phase) setPhase(data.phase);

        if (data.status === "done") {
          clearInterval(pollRef.current);
          setResults({
            legend: data.legend || [],
            takeoffData: data.takeoffData || [],
            projName: file?.name?.replace(".pdf", "") || "Project",
            jobId: id,
          });
          setPhase("done");
          setProgress({ label: "Complete", pct: 100 });
        } else if (data.status === "error") {
          clearInterval(pollRef.current);
          setErrMsg(data.error || "Unknown error");
          setPhase("error");
        }
      } catch (err) {
        // Network hiccup — keep polling, don't stop
        console.log("Poll hiccup:", err.message);
      }
    }, 5000); // Poll every 5 seconds
  }, [file]);

  const run = async () => {
    if (!file) return;
    setPhase("running"); setLog([]); setErrMsg(""); setResults(null);
    seenLogs.current = 0;

    try {
      setLog([{ msg: "Uploading PDF to server...", level: "info" }]);
      const formData = new FormData();
      formData.append("pdf", file);

      const res = await fetch(BACKEND + "/analyze", { method: "POST", body: formData });
      const { jobId: id } = await res.json();

      setJobId(id);
      setLog(prev => [...prev, { msg: "Analysis started — job " + id, level: "ok" }]);
      startPolling(id);
    } catch (err) {
      setErrMsg(err.message);
      setPhase("error");
    }
  };

  const exportExcel = () => {
    if (!results) return;
    const wb = buildExcel(results.projName, results.legend, results.takeoffData);
    XLSX.writeFile(wb, "BPS_Takeoff_" + (results.projName || "Project").replace(/\s+/g, "_") + ".xlsx");
  };

  const exportPDF = async () => {
    if (!results || !results.jobId) return;
    setPdfLoading(true);
    try {
      const res = await fetch(BACKEND + "/evidence-pdf/" + results.jobId);
      if (!res.ok) throw new Error("PDF not ready");
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "BPS_Takeoff_Evidence_" + (results.projName || "Project").replace(/\s+/g, "_") + ".pdf";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert("PDF export failed: " + err.message);
    } finally {
      setPdfLoading(false);
    }
  };

  const summary = results ? (() => {
    const t = {};
    results.takeoffData.forEach(e => (e.zones || []).forEach(z => {
      const k = z.category || "Other";
      if (!t[k]) t[k] = { net: 0, adj: 0, color: MATERIAL_COLORS[k] || "#7a7a7a" };
      t[k].net += z.netArea || 0;
      t[k].adj += (z.netArea || 0) * 1.15;
    }));
    return t;
  })() : null;

  const grandAdj = summary ? Object.values(summary).reduce((s, v) => s + v.adj, 0) : 0;

  const phaseStep = { idle: 0, running: 1, filtering: 1, legend: 2, analyzing: 3, done: 4, error: 0 }[phase] || 0;
  const logColor = { ok: "#6aaa50", warn: "#c8a030", error: "#cc5050", success: "#5ab870", dim: "#5a5040", info: "#8a7a60" };

  return (
    <div style={{ fontFamily: "'Courier New', monospace", background: "#080807", minHeight: "100vh", color: "#ccc4aa", display: "flex", flexDirection: "column" }}>
      <div style={{ background: "#0f0e0b", borderBottom: "2px solid #2a2618", padding: "0.9rem 1.5rem", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: "0.6rem", letterSpacing: "0.35em", color: "#6a5a30", textTransform: "uppercase" }}>Boston Panel Systems</div>
          <div style={{ fontSize: "1.25rem", fontWeight: 700, color: "#e0cc80" }}>AI Panel Estimator</div>
        </div>
        <div style={{ textAlign: "right", fontSize: "0.65rem", color: "#4a4030", lineHeight: 1.8 }}>
          <div>Panels · Soffits · Returns</div>
          <div>Waste: 15% · Scale-based measurement</div>
        </div>
      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        <div style={{ width: 280, borderRight: "1px solid #1e1c14", padding: "1.25rem", display: "flex", flexDirection: "column", gap: "1rem", overflowY: "auto", background: "#0c0b08" }}>
          {/* Upload */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.6rem" }}>Blueprint Set</div>
            <div onClick={() => fileRef.current?.click()} onDrop={e => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }} onDragOver={e => e.preventDefault()}
              style={{ border: "1px dashed " + (file ? "#5a8030" : "#2a2618"), borderRadius: 3, padding: "1.25rem 0.75rem", textAlign: "center", cursor: "pointer", background: file ? "#0b100a" : "transparent" }}>
              <div style={{ fontSize: "1.6rem" }}>{file ? "📋" : "📁"}</div>
              {file ? (<><div style={{ fontSize: "0.72rem", color: "#8ab060", marginTop: "0.3rem", wordBreak: "break-all" }}>{file.name}</div><div style={{ fontSize: "0.62rem", color: "#4a6030" }}>{(file.size / 1e6).toFixed(1)} MB</div></>) : (<div style={{ fontSize: "0.72rem", color: "#3a3020", marginTop: "0.3rem" }}>Drop full blueprint PDF<br />or click to browse</div>)}
            </div>
            <input ref={fileRef} type="file" accept=".pdf" onChange={e => handleFile(e.target.files[0])} style={{ display: "none" }} />
          </div>

          {/* Steps */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.6rem" }}>Process</div>
            {[{ n: 1, label: "Read Sheet Index" }, { n: 2, label: "Read Material Legend" }, { n: 3, label: "Analyze Elevations + Soffits + Returns" }, { n: 4, label: "Export Excel + PDF" }].map(({ n, label }) => {
              const done = phaseStep > n; const active = phaseStep === n;
              return (
                <div key={n} style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.45rem" }}>
                  <div style={{ width: 20, height: 20, borderRadius: "50%", background: done ? "#4a7a30" : active ? "#b89020" : "#181710", border: "1px solid " + (done ? "#4a7a30" : active ? "#b89020" : "#252318"), display: "flex", alignItems: "center", justifyContent: "center", fontSize: "0.58rem", color: done ? "#d0f0a0" : active ? "#080807" : "#2a2818", fontWeight: 700, flexShrink: 0 }}>
                    {done ? "✓" : n}
                  </div>
                  <span style={{ fontSize: "0.68rem", color: done ? "#7aaa50" : active ? "#d0b040" : "#3a3020" }}>{label}</span>
                </div>
              );
            })}
          </div>

          {/* Progress */}
          {progress.pct > 0 && progress.pct < 100 && (
            <div>
              <div style={{ fontSize: "0.62rem", color: "#5a4a20", marginBottom: "0.3rem" }}>{progress.label}</div>
              <div style={{ background: "#181710", borderRadius: 6, height: 5, overflow: "hidden" }}>
                <div style={{ width: progress.pct + "%", height: "100%", background: "#b89020", borderRadius: 6, transition: "width 0.4s" }} />
              </div>
            </div>
          )}

          {/* Buttons */}
          <button onClick={run} disabled={!file || (phase !== "idle" && phase !== "done" && phase !== "error")}
            style={{ padding: "0.8rem", background: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "#151410" : "#b89020", color: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "#2a2418" : "#080807", border: "none", borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: "pointer" }}>
            {phase === "idle" || phase === "error" ? "▶  Run Analysis" : phase === "done" ? "↺  Run Again" : "⏳  Analyzing..."}
          </button>

          {phase === "done" && (
            <>
              <button onClick={exportExcel} style={{ padding: "0.8rem", background: "transparent", color: "#5aaa40", border: "1px solid #3a7a20", borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: "pointer" }}>
                ↓  Export Excel
              </button>
              <button onClick={exportPDF} disabled={pdfLoading} style={{ padding: "0.8rem", background: "transparent", color: pdfLoading ? "#3a5a8a" : "#5a8aaa", border: "1px solid " + (pdfLoading ? "#2a3a5a" : "#3a6a8a"), borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: pdfLoading ? "not-allowed" : "pointer" }}>
                {pdfLoading ? "⏳  Generating..." : "↓  Export Evidence PDF"}
              </button>
            </>
          )}

          {phase === "error" && errMsg && (<div style={{ padding: "0.6rem", background: "#140808", border: "1px solid #4a1818", borderRadius: 3, fontSize: "0.65rem", color: "#cc6060" }}>⚠ {errMsg}</div>)}

          {/* Materials legend */}
          {results && results.legend.length > 0 && (
            <div>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.5rem" }}>Materials</div>
              {results.legend.map((m, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.3rem" }}>
                  <div style={{ width: 10, height: 10, borderRadius: 2, background: MATERIAL_COLORS[m.category] || "#7a7a7a", flexShrink: 0 }} />
                  <span style={{ fontSize: "0.65rem", color: "#8a7a50" }}>{m.id}: {m.name}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Right Panel */}
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {/* Summary cards */}
          {phase === "done" && summary && (
            <div style={{ padding: "1rem 1.25rem", borderBottom: "1px solid #1e1c14", display: "flex", gap: "0.75rem", flexWrap: "wrap", background: "#0a0908" }}>
              {Object.entries(summary).map(([cat, { net, adj, color }]) => (
                <div key={cat} style={{ background: "#111008", border: "1px solid #2a2618", borderLeft: "3px solid " + color, borderRadius: 3, padding: "0.6rem 0.9rem", minWidth: 150 }}>
                  <div style={{ fontSize: "0.62rem", color: "#6a5a30", marginBottom: "0.2rem" }}>{cat}</div>
                  <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#e0cc80" }}>{Math.round(adj).toLocaleString()}</div>
                  <div style={{ fontSize: "0.6rem", color: "#4a4020" }}>SF adj · {Math.round(net).toLocaleString()} SF net</div>
                </div>
              ))}
              <div style={{ background: "#111a08", border: "1px solid #3a5018", borderLeft: "3px solid #5aaa40", borderRadius: 3, padding: "0.6rem 0.9rem", minWidth: 150 }}>
                <div style={{ fontSize: "0.62rem", color: "#5a8030", marginBottom: "0.2rem" }}>GRAND TOTAL</div>
                <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#90e060" }}>{Math.round(grandAdj).toLocaleString()}</div>
                <div style={{ fontSize: "0.6rem", color: "#3a5020" }}>SF adjusted all panels</div>
              </div>
            </div>
          )}

          {/* Elevation breakdown */}
          {phase === "done" && results && (
            <div style={{ flex: 1, overflowY: "auto", padding: "1rem 1.25rem" }}>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.75rem" }}>Breakdown by Elevation</div>
              {results.takeoffData.map((elev, i) => {
                const elevTotal = (elev.zones || []).reduce((s, z) => s + (z.netArea || 0), 0);
                if (elevTotal === 0) return null;
                return (
                  <div key={i} style={{ background: "#0e0d0a", border: "1px solid #1e1c14", borderRadius: 3, marginBottom: "0.6rem", overflow: "hidden" }}>
                    <div style={{ background: "#141208", padding: "0.45rem 0.75rem", display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: "0.78rem", color: "#d0b840", fontWeight: 700 }}>{elev.title}</span>
                      <span style={{ fontSize: "0.6rem", color: "#6a5a30" }}>{elev.sheetRef}{elev.scale ? " · " + elev.scale : ""} · {Math.round(elevTotal)} SF</span>
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.68rem" }}>
                      <thead><tr>{["ID", "Material", "Category", "Gross", "Openings", "Net SF", "Adj +15%"].map(h => (
                        <th key={h} style={{ padding: "0.3rem 0.5rem", textAlign: ["Gross", "Openings", "Net SF", "Adj +15%"].includes(h) ? "right" : "left", color: "#4a3a18", fontWeight: 400, borderBottom: "1px solid #1a1810" }}>{h}</th>
                      ))}</tr></thead>
                      <tbody>
                        {(elev.zones || []).map((z, zi) => (
                          <tr key={zi} style={{ borderBottom: "1px solid #141208" }}>
                            <td style={{ padding: "0.3rem 0.5rem" }}><span style={{ display: "inline-block", width: 8, height: 8, borderRadius: 2, background: MATERIAL_COLORS[z.category] || "#7a7a7a", marginRight: 4 }} /><span style={{ color: "#b89020" }}>{z.materialId || "—"}</span></td>
                            <td style={{ padding: "0.3rem 0.5rem", color: "#a09060" }}>{z.materialName}</td>
                            <td style={{ padding: "0.3rem 0.5rem", color: "#6a5a30" }}>{z.category}</td>
                            <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#907830" }}>{Math.round(z.grossArea || 0)}</td>
                            <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#5a4820" }}>({Math.round(z.totalOpeningArea || 0)})</td>
                            <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#c0a840", fontWeight: 600 }}>{Math.round(z.netArea || 0)}</td>
                            <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#e0cc80", fontWeight: 700 }}>{Math.round((z.netArea || 0) * 1.15)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {(elev.flags || []).filter(Boolean).length > 0 && (<div style={{ padding: "0.3rem 0.5rem", fontSize: "0.62rem", color: "#8a7020", borderTop: "1px solid #1a1810", background: "#110f07" }}>⚠ {elev.flags.filter(Boolean).join(" · ")}</div>)}
                  </div>
                );
              })}
            </div>
          )}

          {/* Idle */}
          {phase === "idle" && (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: "#2a2518", padding: "2rem", textAlign: "center" }}>
              <div style={{ fontSize: "4rem", marginBottom: "1rem", opacity: 0.4 }}>🏗</div>
              <div style={{ fontSize: "0.85rem", color: "#3a3020", marginBottom: "0.5rem" }}>Upload the full blueprint PDF and click Run Analysis</div>
              <div style={{ fontSize: "0.72rem", lineHeight: 1.9 }}>
                Reads sheet index → Material legend → All elevations<br />
                Soffits · Returns · Per-elevation SF breakdown<br />
                Panels only · Evidence PDF + Excel export
              </div>
            </div>
          )}

          {/* Log */}
          <div style={{ borderTop: "1px solid #1e1c14", padding: "0.75rem 1.25rem", background: "#0a0908", flexShrink: 0 }}>
            <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#4a3a18", textTransform: "uppercase", marginBottom: "0.4rem" }}>
              Activity Log {phase !== "idle" && phase !== "done" && phase !== "error" ? "— checking every 5 seconds..." : ""}
            </div>
            <div ref={logRef} style={{ fontFamily: "monospace", fontSize: "0.68rem", maxHeight: 150, overflowY: "auto", lineHeight: 1.7 }}>
              {log.length === 0 ? <span style={{ color: "#2a2418" }}>Waiting...</span> : log.map((l, i) => <div key={i} style={{ color: logColor[l.level] || "#6a5a30" }}>{l.msg}</div>)}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
