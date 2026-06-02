import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const WASTE = 0.15;
const API_URL = "/api/analyze";
const MODEL = "claude-sonnet-4-20250514";

// ─── Load PDF.js from CDN ─────────────────────────────────────────────────────
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

// ─── Render PDF page → base64 JPEG ───────────────────────────────────────────
const renderPage = async (pdf, pageNum, scale = 1.5) => {
  const page = await pdf.getPage(pageNum);
  const vp = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = vp.width;
  canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
  return canvas.toDataURL("image/jpeg", 0.82).split(",")[1];
};

// ─── Claude API call ──────────────────────────────────────────────────────────
const claude = async (content, system) => {
  const res = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", "anthropic-version": "2023-06-01" },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 4000,
      system,
      messages: [{ role: "user", content }],
    }),
  });
  const d = await res.json();
  if (d.error) throw new Error(d.error.message);
  return d.content?.find((b) => b.type === "text")?.text || "";
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

// ─── Build Excel matching BPS format ─────────────────────────────────────────
const buildExcel = (projectName, legend, takeoffData) => {
  const wb = XLSX.utils.book_new();

  // ── TAB 1: JU Estimating style detail ──
  const rows1 = [
    ["", "", "", "", "PROJECT NAME:", "", projectName || "Commercial Project"],
    ["", "", "", "", "DATE:", "", new Date().toLocaleDateString()],
    ["", "", "", "", "WASTE FACTOR:", "", "15%"],
    [],
    ["SR #", "CSI SECT", "DESCRIPTION", "QUANTITY", `WASTAGE\n(15%)`, `QTY WITH\nWASTAGE`, `UNIT OF\nMEASURMENT`, "UNIT COST", `TOTAL ITEM\nCOST`],
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
        const wastePct = WASTE;
        const adjQty = +(z.netArea * (1 + wastePct)).toFixed(1);
        rows1.push([
          sr++,
          "",
          `${z.materialId ? z.materialId + ": " : ""}${z.materialName}${z.description ? " - " + z.description : ""}`,
          +(z.netArea || 0).toFixed(1),
          wastePct,
          adjQty,
          "SF",
          "",
          "",
        ]);
      });
      if (elev.flags?.length) {
        rows1.push(["", "", `⚠ NOTE: ${elev.flags.join(" | ")}`, "", "", "", "", "", ""]);
      }
      rows1.push([]);
    });
  });

  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  ws1["!cols"] = [6, 10, 50, 12, 10, 14, 14, 12, 14].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, "JU Estimating");

  // ── TAB 2: Estimate style (128 Spring St format) ──
  const matTotals = {};
  takeoffData.forEach((e) =>
    e.zones?.forEach((z) => {
      const key = `${z.materialId || ""}||${z.materialName}||${z.category || ""}`;
      if (!matTotals[key]) matTotals[key] = { materialId: z.materialId, materialName: z.materialName, category: z.category || "Other", net: 0, adj: 0 };
      matTotals[key].net += z.netArea || 0;
      matTotals[key].adj += (z.netArea || 0) * (1 + WASTE);
    })
  );

  const rows2 = [
    [projectName || "Commercial Project"],
    ["PANELS"],
    ["No.", "MATERIAL DESCRIPTION", "Quantity", null, "Conv", "Rate", "Amount", null, null, "Notes / Sheet Ref"],
    [],
  ];

  let lineNo = 1;
  const byCategory = {};
  Object.values(matTotals).forEach((m) => {
    if (!byCategory[m.category]) byCategory[m.category] = [];
    byCategory[m.category].push(m);
  });

  Object.entries(byCategory).forEach(([cat, mats]) => {
    rows2.push(["", cat.toUpperCase(), "", "", "", "", "", "", "", ""]);
    mats.forEach((m) => {
      const adjRounded = Math.round(m.adj);
      rows2.push([lineNo++, `${m.materialId ? m.materialId + " - " : ""}${m.materialName}`, adjRounded, "SF", 1, "", ``, null, null, `Net: ${Math.round(m.net)} SF + 15% waste = ${adjRounded} SF`]);
      rows2.push(["", "BACKUP / SUB-FRAMING", adjRounded, "SF", 1, "", ``, null, null, ""]);
      rows2.push([]);
    });
  });

  // Totals summary block
  rows2.push(["MATERIAL TOTALS SUMMARY", "", "", "", "", "", ""]);
  Object.entries(byCategory).forEach(([cat, mats]) => {
    const total = mats.reduce((s, m) => s + m.adj, 0);
    rows2.push(["", cat, Math.round(total), "SF adj.", "", "", ""]);
  });

  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2["!cols"] = [6, 45, 12, 6, 8, 12, 14, 6, 6, 40].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, "Estimate");

  // ── TAB 3: Material Legend reference ──
  const rows3 = [
    ["EXTERIOR BUILDING MATERIALS LEGEND"],
    [],
    ["Material ID", "Material Name", "Category", "Color / Finish", "Notes"],
  ];
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

  const addLog = (msg, type = "info") => setLog((p) => [...p, { msg, type, ts: Date.now() }]);
  const setProg = (label, pct) => setProgress({ label, pct });

  useEffect(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, [log]);

  const handleFile = (f) => {
    if (f?.type === "application/pdf") { setFile(f); setPhase("idle"); setResults(null); setLog([]); setErrMsg(""); }
  };

  const run = async () => {
    if (!file) return;
    setPhase("running"); setLog([]); setErrMsg(""); setResults(null);

    try {
      // ── STEP 1: Load PDF ──
      addLog("Loading PDF...", "info");
      setProg("Loading PDF", 2);
      const pdfjs = await loadPDFJS();
      const buf = await file.arrayBuffer();
      const pdf = await pdfjs.getDocument({ data: buf }).promise;
      const total = pdf.numPages;
      addLog(`✓ ${total} pages loaded — ${file.name}`, "ok");

      // ── STEP 2: Filter pages ──
      addLog("Scanning pages to identify relevant drawings...", "info");
      setPhase("filtering");
      const BATCH = 6;
      const relevant = { floorPlans: [], exteriorElevations: [], returnElevations: [], materialLegend: [], views3d: [], enlargedDetails: [] };

      for (let start = 1; start <= total; start += BATCH) {
        const end = Math.min(start + BATCH - 1, total);
        setProg(`Filtering pages ${start}–${end} of ${total}`, Math.round((end / total) * 30));

        const imgs = [];
        for (let p = start; p <= end; p++) {
          const b64 = await renderPage(pdf, p, 0.25);
          imgs.push({ type: "text", text: `PAGE ${p}:` });
          imgs.push({ type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } });
        }

        const prompt = `Pages ${start}–${end} from an architectural blueprint set. Identify ONLY pages relevant to exterior cladding/siding estimation. Return ONLY JSON (no markdown): {"floorPlans":[page numbers],"exteriorElevations":[page numbers],"returnElevations":[page numbers],"materialLegend":[page numbers],"views3d":[page numbers],"enlargedDetails":[page numbers]}. Return empty arrays for categories not found in this batch.`;
        imgs.push({ type: "text", text: prompt });

        const res = parseJSON(await claude(imgs, "You scan architectural blueprint pages and identify which ones are relevant for exterior cladding estimation. Return ONLY valid JSON."));
        if (res) Object.keys(relevant).forEach((k) => { if (res[k]?.length) relevant[k].push(...res[k]); });
        addLog(`Scanned pages ${start}–${end}`, "dim");
      }

      const totalRelevant = [...new Set(Object.values(relevant).flat())].length;
      addLog(`✓ Filtered: ${relevant.materialLegend.length} legend | ${relevant.exteriorElevations.length} elevation | ${relevant.returnElevations.length} return | ${relevant.views3d.length} 3D view pages`, "ok");
      addLog(`Kept ${totalRelevant} of ${total} pages`, "dim");

      if (relevant.exteriorElevations.length === 0 && relevant.returnElevations.length === 0) {
        throw new Error("No exterior elevation pages found. Make sure the PDF contains architectural elevation drawings.");
      }

      // ── STEP 3: Read Material Legend ──
      setPhase("legend");
      addLog("Reading exterior material legend/key...", "info");
      setProg("Reading material legend", 35);

      let legend = [];
      const legendPages = relevant.materialLegend.length ? relevant.materialLegend : relevant.exteriorElevations.slice(0, 2);

      const legendImgs = [];
      for (const p of legendPages.slice(0, 3)) {
        const b64 = await renderPage(pdf, p, 1.0);
        legendImgs.push({ type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } });
      }
      legendImgs.push({
        type: "text",
        text: `Find the EXTERIOR BUILDING MATERIALS LEGEND or finish schedule on these pages. Extract EVERY material listed. Return ONLY JSON: {"projectName":"project name if visible","materials":[{"id":"material code e.g. ACM-1 or 1 or FCP-01","name":"full material name","category":"ACM Panel|Fiber Cement Panel|Fiber Cement Plank|Soffit|Return/Trim|Aluminum Panel|Nichiha|Other","color":"color or finish if noted","notes":"spec notes"}]}`,
      });

      const legendRaw = await claude(legendImgs, "You read architectural material legends from blueprint drawings. Extract every exterior material entry. Return ONLY valid JSON.");
      const legendParsed = parseJSON(legendRaw);
      legend = legendParsed?.materials || [];
      const projName = legendParsed?.projectName || file.name.replace(".pdf", "");

      if (legend.length > 0) {
        addLog(`✓ Found ${legend.length} materials: ${legend.map((m) => m.id).filter(Boolean).join(", ")}`, "ok");
      } else {
        addLog("⚠ No legend found — will identify materials from drawing callouts directly", "warn");
      }

      // ── STEP 4: Analyze elevations ──
      setPhase("analyzing");
      const elevPages = [
        ...relevant.exteriorElevations.map((p) => ({ p, type: "elevation" })),
        ...relevant.returnElevations.map((p) => ({ p, type: "return" })),
        ...relevant.enlargedDetails.map((p) => ({ p, type: "detail" })),
      ];

      addLog(`Analyzing ${elevPages.length} elevation pages...`, "info");
      const takeoffData = [];
      const legendCtx = legend.length > 0 ? `MATERIAL LEGEND: ${JSON.stringify(legend)}` : "Identify materials from labels and callouts on the drawings.";

      for (let i = 0; i < elevPages.length; i++) {
        const { p, type } = elevPages[i];
        const pct = 40 + Math.round((i / elevPages.length) * 50);
        setProg(`Analyzing page ${p} (${i + 1}/${elevPages.length})`, pct);

        const b64 = await renderPage(pdf, p, 1.2);
        const prompt = `${legendCtx}

This is page ${p} — type: ${type} elevation.

For EVERY elevation drawing visible on this page:
1. Read the drawing title (e.g. "Building 1 South Elevation", "Return North 01", "Enlarged Elevation Building 1 South")
2. Read the sheet reference (e.g. "1/A-201", "3/A-220")  
3. Read the SCALE printed on the drawing (e.g. 1/8"=1'-0", 1/4"=1'-0")
4. Identify EVERY material zone using the legend. Each zone has a hatch pattern, color fill, or callout label.
5. Using the scale, calculate GROSS dimensions (width × height) of each zone in real-world feet → convert to SF
6. List ALL openings within each zone (windows, doors, louvers, curtainwall) with their dimensions → calculate opening area
7. Net Area = Gross Area − Total Opening Area
8. For SOFFITS: these are the underside of overhangs — measure width × depth of overhang
9. For RETURNS: these are corner wraps — measure height × return depth
10. If dimensions are not explicitly shown, use the scale bar and proportional measurement

Return ONLY valid JSON — no markdown:
{
  "pageNumber": ${p},
  "elevations": [
    {
      "title": "Building 1 South Elevation",
      "sheetRef": "1/A-201",
      "scale": "1/8\\"=1'-0\\"",
      "building": "Building 1",
      "direction": "South",
      "zones": [
        {
          "materialId": "ACM-1",
          "materialName": "Champagne ACM Panel",
          "category": "ACM Panel",
          "description": "Main wall field",
          "grossWidth": 0,
          "grossHeight": 0,
          "grossArea": 0,
          "openings": [{"label":"Window","width":0,"height":0,"qty":0,"area":0}],
          "totalOpeningArea": 0,
          "netArea": 0
        }
      ],
      "flags": ["list any dimensions that were unclear or estimated"]
    }
  ]
}`;

        const raw = await claude(
          [{ type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } }, { type: "text", text: prompt }],
          "You are a senior commercial siding estimator performing material takeoffs from architectural drawings. Be precise, thorough, and systematic. Use the scale to calculate real-world dimensions. Return ONLY valid JSON."
        );

        const parsed = parseJSON(raw);
        if (parsed?.elevations?.length) {
          takeoffData.push(...parsed.elevations);
          parsed.elevations.forEach((e) => {
            const totalSF = e.zones?.reduce((s, z) => s + (z.netArea || 0), 0) || 0;
            addLog(`✓ Page ${p}: ${e.title} — ${e.zones?.length || 0} materials, ${Math.round(totalSF)} SF net`, "ok");
            if (e.flags?.length) e.flags.forEach((f) => addLog(`  ⚠ ${e.title}: ${f}`, "warn"));
          });
        } else {
          addLog(`⚠ Page ${p}: could not parse elevation data — may need manual review`, "warn");
        }
      }

      // ── STEP 5: Cross-reference 3D views ──
      if (relevant.views3d.length > 0) {
        setPhase("crossref");
        setProg("Cross-referencing 3D views for soffits/returns", 92);
        addLog("Cross-referencing 3D views for missed soffits and returns...", "info");

        const v3d = await renderPage(pdf, relevant.views3d[0], 1.8);
        const cr = parseJSON(await claude(
          [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: v3d } },
            { type: "text", text: `${legendCtx}\n\nThis is a 3D exterior rendering. Identify any SOFFITS (underside of overhangs, canopies) or RETURNS (corner wraps) visible that may not be fully captured in the flat elevation drawings. Return ONLY JSON: {"soffitsVisible":true/false,"returnsVisible":true/false,"warnings":["list areas to double check"],"notes":"description"}` },
          ],
          "You review 3D exterior renderings to catch soffits and returns that flat elevations might miss."
        ));

        if (cr?.warnings?.length) cr.warnings.forEach((w) => addLog(`⚠ 3D CHECK: ${w}`, "warn"));
        if (cr?.notes) addLog(`3D View: ${cr.notes}`, "dim");
        if (!cr?.warnings?.length) addLog("✓ 3D cross-reference complete — no additional items flagged", "ok");
      }

      // ── STEP 6: Done ──
      setProgress({ label: "Complete", pct: 100 });
      setResults({ legend, takeoffData, projName });
      setPhase("done");
      addLog(`✅ Analysis complete — ${takeoffData.length} elevations processed`, "success");

    } catch (err) {
      setErrMsg(err.message);
      setPhase("error");
      addLog(`❌ ${err.message}`, "error");
    }
  };

  const exportExcel = () => {
    if (!results) return;
    const { legend, takeoffData, projName } = results;
    const wb = buildExcel(projName, legend, takeoffData);
    XLSX.writeFile(wb, `BPS_Takeoff_${(projName || "Project").replace(/\s+/g, "_")}_${new Date().toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" }).replace(/\//g, "-")}.xlsx`);
  };

  // Compute summary totals
  const summary = results
    ? (() => {
        const t = {};
        results.takeoffData.forEach((e) =>
          e.zones?.forEach((z) => {
            const k = z.category || "Other";
            if (!t[k]) t[k] = { net: 0, adj: 0 };
            t[k].net += z.netArea || 0;
            t[k].adj += (z.netArea || 0) * 1.15;
          })
        );
        return t;
      })()
    : null;

  const grandAdj = summary ? Object.values(summary).reduce((s, v) => s + v.adj, 0) : 0;

  const phaseStep = { idle: 0, running: 1, filtering: 1, legend: 2, analyzing: 3, crossref: 4, done: 5, error: 0 }[phase] || 0;

  const logColor = { ok: "#6aaa50", warn: "#c8a030", error: "#cc5050", success: "#5ab870", dim: "#5a5040", info: "#8a7a60" };

  return (
    <div style={{ fontFamily: "'Courier New', Courier, monospace", background: "#080807", minHeight: "100vh", color: "#ccc4aa", display: "flex", flexDirection: "column" }}>

      {/* ── Header ── */}
      <div style={{ background: "#0f0e0b", borderBottom: "2px solid #2a2618", padding: "0.9rem 1.5rem", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
        <div>
          <div style={{ fontSize: "0.6rem", letterSpacing: "0.35em", color: "#6a5a30", textTransform: "uppercase" }}>Boston Panel Systems</div>
          <div style={{ fontSize: "1.25rem", fontWeight: 700, color: "#e0cc80", letterSpacing: "0.02em" }}>AI Siding Estimator</div>
        </div>
        <div style={{ textAlign: "right", fontSize: "0.65rem", color: "#4a4030", lineHeight: 1.8 }}>
          <div>Waste Factor: 15%</div>
          <div>ACM · Fiber Cement · Soffit · Returns</div>
          <div>Scale-based measurement</div>
        </div>
      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>

        {/* ── Left Panel ── */}
        <div style={{ width: 280, borderRight: "1px solid #1e1c14", padding: "1.25rem", display: "flex", flexDirection: "column", gap: "1.25rem", overflowY: "auto", flexShrink: 0, background: "#0c0b08" }}>

          {/* Upload */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.6rem" }}>Blueprint Set</div>
            <div
              onClick={() => fileRef.current?.click()}
              onDrop={(e) => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }}
              onDragOver={(e) => e.preventDefault()}
              style={{ border: `1px dashed ${file ? "#5a8030" : "#2a2618"}`, borderRadius: 3, padding: "1.25rem 0.75rem", textAlign: "center", cursor: "pointer", background: file ? "#0b100a" : "transparent", transition: "all 0.2s" }}
            >
              <div style={{ fontSize: "1.6rem" }}>{file ? "📋" : "📁"}</div>
              {file ? (
                <>
                  <div style={{ fontSize: "0.72rem", color: "#8ab060", marginTop: "0.3rem", wordBreak: "break-all" }}>{file.name}</div>
                  <div style={{ fontSize: "0.62rem", color: "#4a6030", marginTop: "0.2rem" }}>{(file.size / 1e6).toFixed(1)} MB</div>
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
            {[
              { n: 1, label: "Filter Pages", key: "filtering" },
              { n: 2, label: "Read Material Legend", key: "legend" },
              { n: 3, label: "Analyze All Elevations", key: "analyzing" },
              { n: 4, label: "Cross-Reference 3D Views", key: "crossref" },
              { n: 5, label: "Export Excel", key: "done" },
            ].map(({ n, label, key }) => {
              const done = phaseStep > n;
              const active = phase === key || (phase === "running" && n === 1);
              return (
                <div key={key} style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.45rem" }}>
                  <div style={{ width: 20, height: 20, borderRadius: "50%", background: done ? "#4a7a30" : active ? "#b89020" : "#181710", border: `1px solid ${done ? "#4a7a30" : active ? "#b89020" : "#252318"}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "0.58rem", color: done ? "#d0f0a0" : active ? "#080807" : "#2a2818", fontWeight: 700, flexShrink: 0, transition: "all 0.3s" }}>
                    {done ? "✓" : n}
                  </div>
                  <span style={{ fontSize: "0.72rem", color: done ? "#7aaa50" : active ? "#d0b040" : "#3a3020", transition: "color 0.3s" }}>{label}</span>
                  {active && ["running", "filtering", "legend", "analyzing", "crossref"].includes(phase) && (
                    <span style={{ fontSize: "0.6rem", color: "#6a5020", animation: "pulse 1s infinite" }}>•••</span>
                  )}
                </div>
              );
            })}
          </div>

          {/* Progress bar */}
          {progress.pct > 0 && progress.pct < 100 && (
            <div>
              <div style={{ fontSize: "0.62rem", color: "#5a4a20", marginBottom: "0.3rem" }}>{progress.label}</div>
              <div style={{ background: "#181710", borderRadius: 6, height: 5, overflow: "hidden" }}>
                <div style={{ width: `${progress.pct}%`, height: "100%", background: "#b89020", borderRadius: 6, transition: "width 0.4s" }} />
              </div>
            </div>
          )}

          {/* Run button */}
          <button
            onClick={run}
            disabled={!file || (phase !== "idle" && phase !== "done" && phase !== "error")}
            style={{ padding: "0.8rem", background: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "#151410" : "#b89020", color: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "#2a2418" : "#080807", border: "none", borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: !file || (phase !== "idle" && phase !== "done" && phase !== "error") ? "not-allowed" : "pointer", transition: "all 0.2s" }}
          >
            {phase === "idle" || phase === "error" ? "▶  Run Analysis" : phase === "done" ? "↺  Run Again" : "⏳  Analyzing..."}
          </button>

          {/* Export */}
          {phase === "done" && (
            <button
              onClick={exportExcel}
              style={{ padding: "0.8rem", background: "transparent", color: "#5aaa40", border: "1px solid #3a7a20", borderRadius: 3, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", cursor: "pointer" }}
            >
              ↓  Export to Excel
            </button>
          )}

          {/* Error */}
          {phase === "error" && errMsg && (
            <div style={{ padding: "0.6rem", background: "#140808", border: "1px solid #4a1818", borderRadius: 3, fontSize: "0.65rem", color: "#cc6060" }}>⚠ {errMsg}</div>
          )}
        </div>

        {/* ── Right Panel ── */}
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>

          {/* Summary cards */}
          {phase === "done" && summary && (
            <div style={{ padding: "1rem 1.25rem", borderBottom: "1px solid #1e1c14", display: "flex", gap: "0.75rem", flexWrap: "wrap", background: "#0a0908" }}>
              {Object.entries(summary).map(([cat, { net, adj }]) => (
                <div key={cat} style={{ background: "#111008", border: "1px solid #2a2618", borderLeft: "3px solid #b89020", borderRadius: 3, padding: "0.6rem 0.9rem", minWidth: 160 }}>
                  <div style={{ fontSize: "0.62rem", color: "#6a5a30", marginBottom: "0.2rem" }}>{cat}</div>
                  <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#e0cc80", lineHeight: 1 }}>{Math.round(adj).toLocaleString()}</div>
                  <div style={{ fontSize: "0.6rem", color: "#4a4020", marginTop: "0.15rem" }}>SF adj. · {Math.round(net).toLocaleString()} SF net</div>
                </div>
              ))}
              <div style={{ background: "#111a08", border: "1px solid #3a5018", borderLeft: "3px solid #5aaa40", borderRadius: 3, padding: "0.6rem 0.9rem", minWidth: 160 }}>
                <div style={{ fontSize: "0.62rem", color: "#5a8030", marginBottom: "0.2rem" }}>GRAND TOTAL</div>
                <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#90e060", lineHeight: 1 }}>{Math.round(grandAdj).toLocaleString()}</div>
                <div style={{ fontSize: "0.6rem", color: "#3a5020", marginTop: "0.15rem" }}>SF adjusted (all materials)</div>
              </div>
            </div>
          )}

          {/* Elevation table */}
          {phase === "done" && results && (
            <div style={{ flex: 1, overflowY: "auto", padding: "1rem 1.25rem" }}>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.75rem" }}>Elevation Breakdown</div>
              {results.takeoffData.map((elev, i) => (
                <div key={i} style={{ background: "#0e0d0a", border: "1px solid #1e1c14", borderRadius: 3, marginBottom: "0.6rem", overflow: "hidden" }}>
                  <div style={{ background: "#141208", padding: "0.45rem 0.75rem", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontSize: "0.78rem", color: "#d0b840", fontWeight: 700 }}>{elev.title}</span>
                    <span style={{ fontSize: "0.6rem", color: "#4a3a18" }}>{elev.sheetRef}{elev.scale ? ` · ${elev.scale}` : ""}</span>
                  </div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.68rem" }}>
                    <thead>
                      <tr>
                        {["Material ID", "Name", "Category", "Gross SF", "Openings", "Net SF", "Adj SF (+15%)"].map((h) => (
                          <th key={h} style={{ padding: "0.3rem 0.5rem", textAlign: h.includes("SF") || h.includes("Open") ? "right" : "left", color: "#4a3a18", fontWeight: 400, borderBottom: "1px solid #1a1810" }}>{h}</th>
                        ))}
                      </tr>
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

          {/* Idle state */}
          {(phase === "idle") && (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: "#2a2518", padding: "2rem" }}>
              <div style={{ fontSize: "4rem", marginBottom: "1rem", opacity: 0.4 }}>🏗</div>
              <div style={{ fontSize: "0.85rem", marginBottom: "0.5rem", color: "#3a3020" }}>Upload the full blueprint PDF set and click Run Analysis</div>
              <div style={{ fontSize: "0.72rem", textAlign: "center", lineHeight: 1.8, color: "#2a2418" }}>
                The AI will filter pages → read the material legend<br />
                → measure every elevation using the drawing scale<br />
                → cross-reference 3D views for soffits & returns<br />
                → export to Excel in BPS estimating format
              </div>
            </div>
          )}

          {/* Activity log */}
          <div style={{ borderTop: "1px solid #1e1c14", padding: "0.75rem 1.25rem", background: "#0a0908", flexShrink: 0 }}>
            <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#4a3a18", textTransform: "uppercase", marginBottom: "0.4rem" }}>Activity Log</div>
            <div ref={logRef} style={{ fontFamily: "monospace", fontSize: "0.68rem", maxHeight: 130, overflowY: "auto", lineHeight: 1.7 }}>
              {log.length === 0
                ? <span style={{ color: "#2a2418" }}>Waiting for analysis...</span>
                : log.map((l, i) => <div key={i} style={{ color: logColor[l.type] || "#6a5a30" }}>{l.msg}</div>)}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
