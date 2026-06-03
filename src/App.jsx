import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const WASTE = 0.15;
const MODEL = "claude-sonnet-4-20250514";

// ─── Load PDF.js ──────────────────────────────────────────────────────────────
const loadPDFJS = () =>
  new Promise((resolve, reject) => {
    if (window.pdfjsLib) { resolve(window.pdfjsLib); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    s.onerror = reject;
    document.head.appendChild(s);
  });

// ─── Render one PDF page to base64 JPEG ──────────────────────────────────────
const renderPage = async (pdf, pageNum, scale = 1.0) => {
  const page = await pdf.getPage(pageNum);
  const vp = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = vp.width;
  canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
  return canvas.toDataURL("image/jpeg", 0.75).split(",")[1];
};

// ─── Call Anthropic API directly from browser ─────────────────────────────────
const claude = async (content, system) => {
  const apiKey = import.meta.env.VITE_ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("API key not found. Make sure VITE_ANTHROPIC_API_KEY is set in Vercel.");

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 4000,
      system,
      messages: [{ role: "user", content }],
    }),
  });

  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.find((b) => b.type === "text")?.text || "";
};

// ─── Parse JSON from Claude response ─────────────────────────────────────────
const parseJSON = (text) => {
  try {
    const m = text.match(/```json\s*([\s\S]*?)```/);
    return JSON.parse(m ? m[1] : text);
  } catch {
    const s = text.indexOf("{"), e = text.lastIndexOf("}");
    if (s !== -1 && e !== -1) try { return JSON.parse(text.slice(s, e + 1)); } catch {}
    return null;
  }
};

// ─── Build Excel ──────────────────────────────────────────────────────────────
const buildExcel = (projectName, legend, takeoffData) => {
  const wb = XLSX.utils.book_new();

  // Tab 1: JU Estimating detail format
  const rows1 = [
    ["", "", "", "", "PROJECT:", "", projectName || "Commercial Project"],
    ["", "", "", "", "DATE:", "", new Date().toLocaleDateString()],
    ["", "", "", "", "WASTE FACTOR:", "", "15%"],
    [],
    ["SR #", "CSI SECT", "DESCRIPTION", "QUANTITY", "WASTAGE (15%)", "QTY WITH WASTAGE", "UNIT", "UNIT COST", "TOTAL ITEM COST"],
  ];

  let sr = 1;
  const byBuilding = {};
  takeoffData.forEach((e) => {
    const b = e.building || "Building";
    if (!byBuilding[b]) byBuilding[b] = [];
    byBuilding[b].push(e);
  });

  Object.entries(byBuilding).forEach(([bld, elevs]) => {
    rows1.push(["", "DIV. 09", bld.toUpperCase(), "", "", "", "", "", ""]);
    elevs.forEach((elev) => {
      rows1.push(["", "", `Siding (${elev.sheetRef || elev.title})`, "", "", "", "", "", ""]);
      elev.zones?.forEach((z) => {
        const adj = +((z.netArea || 0) * 1.15).toFixed(1);
        rows1.push([sr++, "", `${z.materialId ? z.materialId + ": " : ""}${z.materialName}${z.description ? " - " + z.description : ""}`, +(z.netArea || 0).toFixed(1), 0.15, adj, "SF", "", ""]);
      });
      if (elev.flags?.length) rows1.push(["", "", `NOTE: ${elev.flags.join(" | ")}`, "", "", "", "", "", ""]);
      rows1.push([]);
    });
  });

  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  ws1["!cols"] = [6, 10, 50, 12, 14, 16, 8, 12, 14].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, "JU Estimating");

  // Tab 2: Estimate summary format (128 Spring St style)
  const matTotals = {};
  takeoffData.forEach((e) =>
    e.zones?.forEach((z) => {
      const key = `${z.materialId || ""}||${z.materialName}||${z.category || ""}`;
      if (!matTotals[key]) matTotals[key] = { materialId: z.materialId, materialName: z.materialName, category: z.category || "Other", net: 0, adj: 0 };
      matTotals[key].net += z.netArea || 0;
      matTotals[key].adj += (z.netArea || 0) * 1.15;
    })
  );

  const rows2 = [
    [projectName || "Commercial Project"],
    ["PANELS"],
    [],
    ["No.", "MATERIAL DESCRIPTION", "Quantity", "Unit", "Conv", "Rate", "Amount", "", "Notes"],
  ];

  let lineNo = 1;
  const byCategory = {};
  Object.values(matTotals).forEach((m) => {
    if (!byCategory[m.category]) byCategory[m.category] = [];
    byCategory[m.category].push(m);
  });

  Object.entries(byCategory).forEach(([cat, mats]) => {
    rows2.push(["", cat.toUpperCase()]);
    mats.forEach((m) => {
      const adj = Math.round(m.adj);
      rows2.push([lineNo++, `${m.materialId ? m.materialId + " - " : ""}${m.materialName}`, adj, "SF", 1, "", "", "", `Net: ${Math.round(m.net)} SF + 15% = ${adj} SF`]);
      rows2.push(["", "BACKUP / SUB-FRAMING", adj, "SF", 1, "", "", "", ""]);
      rows2.push([]);
    });
  });

  rows2.push(["TOTALS"]);
  Object.entries(byCategory).forEach(([cat, mats]) => {
    rows2.push(["", cat, Math.round(mats.reduce((s, m) => s + m.adj, 0)), "SF adj."]);
  });

  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2["!cols"] = [6, 45, 12, 6, 6, 12, 14, 6, 40].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, "Estimate");

  // Tab 3: Legend reference
  const rows3 = [["EXTERIOR BUILDING MATERIALS LEGEND"], [], ["Material ID", "Name", "Category", "Color/Finish", "Notes"]];
  legend.forEach((m) => rows3.push([m.id, m.name, m.category, m.color || "", m.notes || ""]));
  const ws3 = XLSX.utils.aoa_to_sheet(rows3);
  ws3["!cols"] = [14, 35, 20, 20, 30].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws3, "Material Legend");

  return wb;
};

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function BPSEstimator() {
  const [file, setFile] = useState(null);
  const [phase, setPhase] = useState("idle");
  const [log, setLog] = useState([]);
  const [progress, setProgress] = useState({ label: "", pct: 0 });
  const [results, setResults] = useState(null);
  const [errMsg, setErrMsg] = useState("");
  const fileRef = useRef();
  const logRef = useRef();

  const addLog = (msg, type = "info") => setLog((p) => [...p, { msg, type }]);
  const setProg = (label, pct) => setProgress({ label, pct });

  useEffect(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, [log]);

  const handleFile = (f) => {
    if (f?.type === "application/pdf") { setFile(f); setPhase("idle"); setResults(null); setLog([]); setErrMsg(""); }
  };

  const run = async () => {
    if (!file) return;
    setPhase("running"); setLog([]); setErrMsg(""); setResults(null);

    try {
      // Load PDF.js
      addLog("Loading PDF...");
      setProg("Loading PDF", 2);
      const pdfjs = await loadPDFJS();
      const buf = await file.arrayBuffer();
      const pdf = await pdfjs.getDocument({ data: buf }).promise;
      const total = pdf.numPages;
      addLog(`✓ ${total} pages loaded — ${file.name}`, "ok");

      // ── STEP 1: Filter pages — send 1 page at a time at very low res ──
      setPhase("filtering");
      addLog("Scanning pages to find elevations, floor plans, and legend...", "info");

      const relevant = { floorPlans: [], exteriorElevations: [], returnElevations: [], materialLegend: [], views3d: [], enlargedDetails: [] };

      for (let p = 1; p <= total; p++) {
        setProg(`Scanning page ${p} of ${total}`, Math.round((p / total) * 30));
        const b64 = await renderPage(pdf, p, 0.1); // very small thumbnail

        const res = parseJSON(await claude(
          [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } },
            { type: "text", text: `This is page ${p} of a commercial architectural blueprint. Classify it. Return ONLY JSON: {"types":["FLOOR_PLAN"|"EXTERIOR_ELEVATION"|"RETURN_ELEVATION"|"MATERIAL_LEGEND"|"3D_VIEW"|"ENLARGED_DETAIL"|"IRRELEVANT"]}` },
          ],
          "Classify architectural blueprint pages. Return ONLY valid JSON."
        ));

        if (res?.types) {
          const t = res.types;
          if (t.includes("FLOOR_PLAN")) relevant.floorPlans.push(p);
          if (t.includes("EXTERIOR_ELEVATION")) relevant.exteriorElevations.push(p);
          if (t.includes("RETURN_ELEVATION")) relevant.returnElevations.push(p);
          if (t.includes("MATERIAL_LEGEND")) relevant.materialLegend.push(p);
          if (t.includes("3D_VIEW")) relevant.views3d.push(p);
          if (t.includes("ENLARGED_DETAIL")) relevant.enlargedDetails.push(p);
          if (!t.includes("IRRELEVANT") && !t.includes("MATERIAL_LEGEND")) {
            addLog(`Page ${p}: ${t.join(", ")}`, "dim");
          }
        }
      }

      addLog(`✓ Found: ${relevant.materialLegend.length} legend | ${relevant.exteriorElevations.length} elevations | ${relevant.returnElevations.length} returns | ${relevant.views3d.length} 3D views`, "ok");

      if (!relevant.exteriorElevations.length && !relevant.returnElevations.length) {
        throw new Error("No exterior elevation pages found in this PDF.");
      }

      // ── STEP 2: Read Material Legend ──
      setPhase("legend");
      addLog("Reading material legend...", "info");
      setProg("Reading legend", 35);

      let legend = [];
      const legendPages = relevant.materialLegend.length ? relevant.materialLegend : relevant.exteriorElevations.slice(0, 2);

      for (const p of legendPages.slice(0, 3)) {
        const b64 = await renderPage(pdf, p, 0.7);
        const raw = await claude(
          [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } },
            { type: "text", text: `Find the EXTERIOR BUILDING MATERIALS LEGEND on this page. Extract every material. Return ONLY JSON: {"projectName":"if visible","materials":[{"id":"e.g. ACM-1","name":"full name","category":"ACM Panel|Fiber Cement Panel|Fiber Cement Plank|Soffit|Return/Trim|Aluminum Panel|Nichiha|Other","color":"if noted","notes":"spec notes"}]}` },
          ],
          "Extract exterior material legends from architectural drawings. Return ONLY valid JSON."
        );
        const parsed = parseJSON(raw);
        if (parsed?.materials?.length) {
          legend = parsed.materials;
          addLog(`✓ Found ${legend.length} materials: ${legend.map((m) => m.id).join(", ")}`, "ok");
          break;
        }
      }

      if (!legend.length) addLog("⚠ No legend found — will identify materials from drawing labels", "warn");

      // ── STEP 3: Analyze each elevation ──
      setPhase("analyzing");
      const elevPages = [
        ...relevant.exteriorElevations.map((p) => ({ p, type: "elevation" })),
        ...relevant.returnElevations.map((p) => ({ p, type: "return" })),
        ...relevant.enlargedDetails.map((p) => ({ p, type: "detail" })),
      ];

      addLog(`Analyzing ${elevPages.length} elevation pages...`, "info");
      const legendCtx = legend.length ? `MATERIAL LEGEND: ${JSON.stringify(legend)}` : "Identify materials from callouts on the drawing.";
      const takeoffData = [];

      for (let i = 0; i < elevPages.length; i++) {
        const { p, type } = elevPages[i];
        setProg(`Analyzing elevation ${i + 1} of ${elevPages.length}`, 40 + Math.round((i / elevPages.length) * 50));

        const b64 = await renderPage(pdf, p, 0.7);
        const prompt = `${legendCtx}

Page ${p} — ${type}. For EVERY elevation drawing on this page:
1. Read the title (e.g. "Building 1 South Elevation")
2. Read the sheet reference (e.g. "1/A-201")
3. Read the SCALE (e.g. 1/8"=1'-0")
4. For each material zone: identify material from legend, calculate gross area using scale, subtract openings (windows/doors/louvers), get net area in SF
5. Flag soffits (underside of overhangs) and returns (corner wraps) separately
6. If no dimensions shown, use scale bar to measure

Return ONLY JSON:
{"pageNumber":${p},"elevations":[{"title":"","sheetRef":"","scale":"","building":"","direction":"","zones":[{"materialId":"","materialName":"","category":"","description":"","grossArea":0,"totalOpeningArea":0,"netArea":0}],"flags":[]}]}`;

        const raw = await claude(
          [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } },
            { type: "text", text: prompt },
          ],
          "You are a senior commercial siding estimator doing material takeoffs. Be precise. Return ONLY valid JSON."
        );

        const parsed = parseJSON(raw);
        if (parsed?.elevations?.length) {
          takeoffData.push(...parsed.elevations);
          parsed.elevations.forEach((e) => {
            const sf = e.zones?.reduce((s, z) => s + (z.netArea || 0), 0) || 0;
            addLog(`✓ ${e.title} — ${e.zones?.length || 0} materials, ${Math.round(sf)} SF`, "ok");
            e.flags?.forEach((f) => addLog(`  ⚠ ${f}`, "warn"));
          });
        } else {
          addLog(`⚠ Page ${p}: could not read — manual review needed`, "warn");
        }
      }

      // ── STEP 4: 3D cross-reference ──
      if (relevant.views3d.length) {
        setProg("Cross-referencing 3D views", 92);
        addLog("Cross-checking 3D views for soffits and returns...", "info");
        const b64 = await renderPage(pdf, relevant.views3d[0], 1.0);
        const cr = parseJSON(await claude(
          [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } },
            { type: "text", text: `${legendCtx}\n\nThis is a 3D exterior view. Are there soffits or returns visible that flat elevations might miss? Return ONLY JSON: {"warnings":["items to double check"],"notes":"what you see"}` },
          ],
          "Review 3D renderings to catch missed soffits and returns."
        ));
        if (cr?.warnings?.length) cr.warnings.forEach((w) => addLog(`⚠ 3D: ${w}`, "warn"));
        else addLog("✓ 3D cross-reference complete", "ok");
      }

      setResults({ legend, takeoffData, projName: file.name.replace(".pdf", "") });
      setPhase("done");
      setProg("Complete", 100);
      addLog(`✅ Done — ${takeoffData.length} elevations analyzed`, "success");

    } catch (err) {
      setErrMsg(err.message);
      setPhase("error");
      addLog(`❌ ${err.message}`, "error");
    }
  };

  const exportExcel = () => {
    if (!results) return;
    const wb = buildExcel(results.projName, results.legend, results.takeoffData);
    XLSX.writeFile(wb, `BPS_Takeoff_${(results.projName || "Project").replace(/\s+/g, "_")}.xlsx`);
  };

  const summary = results ? (() => {
    const t = {};
    results.takeoffData.forEach((e) => e.zones?.forEach((z) => {
      const k = z.category || "Other";
      if (!t[k]) t[k] = { net: 0, adj: 0 };
      t[k].net += z.netArea || 0;
      t[k].adj += (z.netArea || 0) * 1.15;
    }));
    return t;
  })() : null;

  const grandAdj = summary ? Object.values(summary).reduce((s, v) => s + v.adj, 0) : 0;
  const phaseStep = { idle: 0, running: 1, filtering: 1, legend: 2, analyzing: 3, crossref: 4, done: 5, error: 0 }[phase] || 0;
  const logColor = { ok: "#6aaa50", warn: "#c8a030", error: "#cc5050", success: "#5ab870", dim: "#5a5040", info: "#8a7a60" };

  return (
    <div style={{ fontFamily: "'Courier New', monospace", background: "#080807", minHeight: "100vh", color: "#ccc4aa", display: "flex", flexDirection: "column" }}>
      {/* Header */}
      <div style={{ background: "#0f0e0b", borderBottom: "2px solid #2a2618", padding: "0.9rem 1.5rem", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: "0.6rem", letterSpacing: "0.35em", color: "#6a5a30", textTransform: "uppercase" }}>Boston Panel Systems</div>
          <div style={{ fontSize: "1.25rem", fontWeight: 700, color: "#e0cc80" }}>AI Siding Estimator</div>
        </div>
        <div style={{ textAlign: "right", fontSize: "0.65rem", color: "#4a4030", lineHeight: 1.8 }}>
          <div>Waste Factor: 15%</div>
          <div>ACM · Fiber Cement · Soffit · Returns</div>
        </div>
      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* Left Panel */}
        <div style={{ width: 280, borderRight: "1px solid #1e1c14", padding: "1.25rem", display: "flex", flexDirection: "column", gap: "1.25rem", overflowY: "auto", background: "#0c0b08" }}>
          {/* Upload */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.6rem" }}>Blueprint Set</div>
            <div onClick={() => fileRef.current?.click()} onDrop={(e) => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }} onDragOver={(e) => e.preventDefault()}
              style={{ border: `1px dashed ${file ? "#5a8030" : "#2a2618"}`, borderRadius: 3, padding: "1.25rem 0.75rem", textAlign: "center", cursor: "pointer", background: file ? "#0b100a" : "transparent" }}>
              <div style={{ fontSize: "1.6rem" }}>{file ? "📋" : "📁"}</div>
              {file ? (
                <>
                  <div style={{ fontSize: "0.72rem", color: "#8ab060", marginTop: "0.3rem", wordBreak: "break-all" }}>{file.name}</div>
                  <div style={{ fontSize: "0.62rem", color: "#4a6030" }}>{(file.size / 1e6).toFixed(1)} MB</div>
                </>
              ) : (
                <div style={{ fontSize: "0.72rem", color: "#3a3020", marginTop: "0.3rem" }}>Drop full blueprint PDF<br />or click to browse</div>
              )}
            </div>
            <input ref={fileRef} type="file" accept=".pdf" onChange={(e) => handleFile(e.target.files[0])} style={{ display: "none" }} />
          </div>

          {/* Steps */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.6rem" }}>Process</div>
            {[{ n: 1, label: "Filter Pages" }, { n: 2, label: "Read Material Legend" }, { n: 3, label: "Analyze All Elevations" }, { n: 4, label: "Cross-Reference 3D Views" }, { n: 5, label: "Export Excel" }].map(({ n, label }) => {
              const done = phaseStep > n;
              const active = phaseStep === n;
              return (
                <div key={n} style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.45rem" }}>
                  <div style={{ width: 20, height: 20, borderRadius: "50%", background: done ? "#4a7a30" : active ? "#b89020" : "#181710", border: `1px solid ${done ? "#4a7a30" : active ? "#b89020" : "#252318"}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "0.58rem", color: done ? "#d0f0a0" : active ? "#080807" : "#2a2818", fontWeight: 700, flexShrink: 0 }}>
                    {done ? "✓" : n}
                  </div>
                  <span style={{ fontSize: "0.72rem", color: done ? "#7aaa50" : active ? "#d0b040" : "#3a3020" }}>{label}</span>
                </div>
              );
            })}
          </div>

          {/* Progress */}
          {progress.pct > 0 && progress.pct < 100 && (
            <div>
              <div style={{ fontSize: "0.62rem", color: "#5a4a20", marginBottom: "0.3rem" }}>{progress.label}</div>
              <div style={{ background: "#181710", borderRadius: 6, height: 5, overflow: "hidden" }}>
                <div style={{ width: `${progress.pct}%`, height: "100%", background: "#b89020", borderRadius: 6, transition: "width 0.4s" }} />
              </div>
            </div>
          )}

          {/* Run */}
          <button onClick={run} disabled={!file || (phase !== "idle" && phase !== "done" && phase !== "error")}
            style={{ padding: "0.8rem", background: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "#151410" : "#b89020", color: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "#2a2418" : "#080807", border: "none", borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: "pointer" }}>
            {phase === "idle" || phase === "error" ? "▶  Run Analysis" : phase === "done" ? "↺  Run Again" : "⏳  Analyzing..."}
          </button>

          {/* Export */}
          {phase === "done" && (
            <button onClick={exportExcel} style={{ padding: "0.8rem", background: "transparent", color: "#5aaa40", border: "1px solid #3a7a20", borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: "pointer" }}>
              ↓  Export to Excel
            </button>
          )}

          {phase === "error" && errMsg && (
            <div style={{ padding: "0.6rem", background: "#140808", border: "1px solid #4a1818", borderRadius: 3, fontSize: "0.65rem", color: "#cc6060" }}>⚠ {errMsg}</div>
          )}
        </div>

        {/* Right Panel */}
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {/* Summary cards */}
          {phase === "done" && summary && (
            <div style={{ padding: "1rem 1.25rem", borderBottom: "1px solid #1e1c14", display: "flex", gap: "0.75rem", flexWrap: "wrap", background: "#0a0908" }}>
              {Object.entries(summary).map(([cat, { net, adj }]) => (
                <div key={cat} style={{ background: "#111008", border: "1px solid #2a2618", borderLeft: "3px solid #b89020", borderRadius: 3, padding: "0.6rem 0.9rem", minWidth: 150 }}>
                  <div style={{ fontSize: "0.62rem", color: "#6a5a30", marginBottom: "0.2rem" }}>{cat}</div>
                  <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#e0cc80" }}>{Math.round(adj).toLocaleString()}</div>
                  <div style={{ fontSize: "0.6rem", color: "#4a4020" }}>SF adj · {Math.round(net).toLocaleString()} net</div>
                </div>
              ))}
              <div style={{ background: "#111a08", border: "1px solid #3a5018", borderLeft: "3px solid #5aaa40", borderRadius: 3, padding: "0.6rem 0.9rem", minWidth: 150 }}>
                <div style={{ fontSize: "0.62rem", color: "#5a8030", marginBottom: "0.2rem" }}>GRAND TOTAL</div>
                <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#90e060" }}>{Math.round(grandAdj).toLocaleString()}</div>
                <div style={{ fontSize: "0.6rem", color: "#3a5020" }}>SF adjusted total</div>
              </div>
            </div>
          )}

          {/* Elevation table */}
          {phase === "done" && results && (
            <div style={{ flex: 1, overflowY: "auto", padding: "1rem 1.25rem" }}>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.75rem" }}>Elevation Breakdown</div>
              {results.takeoffData.map((elev, i) => (
                <div key={i} style={{ background: "#0e0d0a", border: "1px solid #1e1c14", borderRadius: 3, marginBottom: "0.6rem", overflow: "hidden" }}>
                  <div style={{ background: "#141208", padding: "0.45rem 0.75rem", display: "flex", justifyContent: "space-between" }}>
                    <span style={{ fontSize: "0.78rem", color: "#d0b840", fontWeight: 700 }}>{elev.title}</span>
                    <span style={{ fontSize: "0.6rem", color: "#4a3a18" }}>{elev.sheetRef}{elev.scale ? ` · ${elev.scale}` : ""}</span>
                  </div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.68rem" }}>
                    <thead>
                      <tr>{["Material ID", "Name", "Category", "Gross SF", "Openings", "Net SF", "Adj SF (+15%)"].map((h) => (
                        <th key={h} style={{ padding: "0.3rem 0.5rem", textAlign: h.includes("SF") || h === "Openings" ? "right" : "left", color: "#4a3a18", fontWeight: 400, borderBottom: "1px solid #1a1810" }}>{h}</th>
                      ))}</tr>
                    </thead>
                    <tbody>
                      {elev.zones?.map((z, zi) => (
                        <tr key={zi} style={{ borderBottom: "1px solid #141208" }}>
                          <td style={{ padding: "0.3rem 0.5rem", color: "#b89020" }}>{z.materialId || "—"}</td>
                          <td style={{ padding: "0.3rem 0.5rem", color: "#a09060" }}>{z.materialName}</td>
                          <td style={{ padding: "0.3rem 0.5rem", color: "#6a5a30" }}>{z.category}</td>
                          <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#907830" }}>{Math.round(z.grossArea || 0)}</td>
                          <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#5a4820" }}>({Math.round(z.totalOpeningArea || 0)})</td>
                          <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#c0a840" }}>{Math.round(z.netArea || 0)}</td>
                          <td style={{ padding: "0.3rem 0.5rem", textAlign: "right", color: "#e0cc80", fontWeight: 700 }}>{Math.round((z.netArea || 0) * 1.15)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {elev.flags?.filter(Boolean).length > 0 && (
                    <div style={{ padding: "0.3rem 0.5rem", fontSize: "0.62rem", color: "#8a7020", borderTop: "1px solid #1a1810", background: "#110f07" }}>
                      ⚠ {elev.flags.join(" · ")}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}

          {/* Idle */}
          {phase === "idle" && (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: "#2a2518", padding: "2rem", textAlign: "center" }}>
              <div style={{ fontSize: "4rem", marginBottom: "1rem", opacity: 0.4 }}>🏗</div>
              <div style={{ fontSize: "0.85rem", color: "#3a3020", marginBottom: "0.5rem" }}>Upload the full blueprint PDF and click Run Analysis</div>
              <div style={{ fontSize: "0.72rem", lineHeight: 1.8 }}>Filter pages → Read legend → Measure elevations<br />Cross-check 3D views → Export to Excel</div>
            </div>
          )}

          {/* Log */}
          <div style={{ borderTop: "1px solid #1e1c14", padding: "0.75rem 1.25rem", background: "#0a0908", flexShrink: 0 }}>
            <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#4a3a18", textTransform: "uppercase", marginBottom: "0.4rem" }}>Activity Log</div>
            <div ref={logRef} style={{ fontFamily: "monospace", fontSize: "0.68rem", maxHeight: 130, overflowY: "auto", lineHeight: 1.7 }}>
              {log.length === 0 ? <span style={{ color: "#2a2418" }}>Waiting...</span> : log.map((l, i) => <div key={i} style={{ color: logColor[l.type] || "#6a5a30" }}>{l.msg}</div>)}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
