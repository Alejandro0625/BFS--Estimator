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

// ── Interactive Takeoff View ──────────────────────────────────────────────────
function InteractiveView({ results, BACKEND }) {
  const [elevIdx, setElevIdx] = useState(0);
  const [pageImage, setPageImage] = useState(null);
  const [imgLoaded, setImgLoaded] = useState(false);
  const [pagePolygons, setPagePolygons] = useState([]);
  const [pageDims, setPageDims] = useState({ width: 612, height: 792 });
  const [assignments, setAssignments] = useState({}); // "elevIdx:polyId" → { category, materialId, materialName, area_sf }
  const [activeZone, setActiveZone] = useState(null);
  const [imgNaturalSize, setImgNaturalSize] = useState({ w: 1, h: 1 });
  const imgRef = useRef();

  const elevations = results.takeoffData.filter(e => e.pageNumber);
  const elev = elevations[elevIdx];
  const pageNum = elev?.pageNumber;

  useEffect(() => {
    if (!pageNum || !results.jobId) return;
    setPageImage(null);
    setImgLoaded(false);
    setPagePolygons([]);
    setActiveZone(null);

    fetch(BACKEND + "/polygons/" + results.jobId + "/" + pageNum)
      .then(r => r.ok ? r.json() : { polygons: [], width: 612, height: 792 })
      .then(d => {
        setPagePolygons(d.polygons || []);
        setPageDims({ width: d.width || 612, height: d.height || 792 });
      })
      .catch(() => {});

    setPageImage(BACKEND + "/page-image/" + results.jobId + "/" + pageNum);
  }, [elevIdx, pageNum, results.jobId, BACKEND]);

  // Bluebeam annotations → vector polygons → Claude bbox fallback
  const polyMethod = pagePolygons[0]?.source || (pagePolygons.length > 0 ? "vector" : "box");
  const useVectorMode = pagePolygons.length > 0;
  const displayZones = useVectorMode
    ? pagePolygons
    : (elev?.zones || []).map((z, i) => ({
        id: i,
        points: [
          [z.x0pct / 100, z.y0pct / 100],
          [z.x1pct / 100, z.y0pct / 100],
          [z.x1pct / 100, z.y1pct / 100],
          [z.x0pct / 100, z.y1pct / 100],
        ],
        area_sf: z.netArea || 0,
        cx: (z.x0pct + z.x1pct) / 200,
        cy: (z.y0pct + z.y1pct) / 200,
        source: "box",
        suggestedCategory: z.category,
        suggestedId: z.materialId,
        suggestedName: z.materialName,
      }));

  // Group zones by fill_color (Bluebeam uses color to distinguish materials)
  const colorGroups = {};
  displayZones.forEach(z => {
    const key = z.fill_color ? z.fill_color.join(",") : "none";
    if (!colorGroups[key]) colorGroups[key] = [];
    colorGroups[key].push(z.id);
  });

  const assignKey = id => elevIdx + ":" + id;
  const getAssignment = id => assignments[assignKey(id)];

  const assignZone = (zoneId, mat) => {
    const zone = displayZones.find(z => z.id === zoneId);
    setAssignments(prev => ({
      ...prev,
      [assignKey(zoneId)]: { ...mat, area_sf: zone?.area_sf || 0 },
    }));
    setActiveZone(null);
  };

  const removeAssignment = zoneId => {
    const k = assignKey(zoneId);
    setAssignments(prev => { const n = { ...prev }; delete n[k]; return n; });
  };

  // Totals across all elevations
  const totals = {};
  Object.values(assignments).forEach(a => {
    if (!totals[a.category]) totals[a.category] = 0;
    totals[a.category] += a.area_sf || 0;
  });
  const grandTotal = Object.values(totals).reduce((s, v) => s + v, 0);

  // SVG is sized to match the actual rendered image element
  const svgW = imgRef.current?.offsetWidth || imgNaturalSize.w;
  const svgH = imgRef.current?.offsetHeight || imgNaturalSize.h;

  // Convert normalized 0-1 coords → SVG point string (viewBox is pageDims)
  const toSVGPoints = pts =>
    pts.map(([nx, ny]) => `${(nx * pageDims.width).toFixed(1)},${(ny * pageDims.height).toFixed(1)}`).join(" ");

  const matList = results.legend.length > 0
    ? results.legend
    : Object.keys(MATERIAL_COLORS).map(cat => ({ id: cat.substring(0, 3).toUpperCase() + "-1", name: cat, category: cat }));

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden", background: "#0a0908" }}>

      {/* ── Left: elevation list ── */}
      <div style={{ width: 175, borderRight: "1px solid #1e1c14", overflowY: "auto", flexShrink: 0, background: "#0c0b08" }}>
        <div style={{ padding: "0.6rem 0.75rem", fontSize: "0.58rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase" }}>
          Elevations
        </div>
        {elevations.map((e, i) => {
          const assigned = Object.keys(assignments).filter(k => k.startsWith(i + ":")).length;
          return (
            <div key={i} onClick={() => setElevIdx(i)}
              style={{ padding: "0.5rem 0.75rem", cursor: "pointer", borderBottom: "1px solid #141208",
                background: i === elevIdx ? "#141208" : "transparent",
                borderLeft: i === elevIdx ? "2px solid #b89020" : "2px solid transparent" }}>
              <div style={{ fontSize: "0.68rem", color: i === elevIdx ? "#d0b840" : "#7a6a40", lineHeight: 1.3 }}>
                {e.title || "Page " + e.pageNumber}
              </div>
              <div style={{ fontSize: "0.58rem", color: assigned > 0 ? "#5a8030" : "#4a3a18", marginTop: 2 }}>
                {assigned > 0 ? "✓ " + assigned + " assigned" : (e.zones || []).length + " zones · p." + e.pageNumber}
              </div>
            </div>
          );
        })}
        {elevations.length === 0 && (
          <div style={{ padding: "0.75rem", fontSize: "0.65rem", color: "#3a3020" }}>
            No elevations with page numbers. Re-run analysis.
          </div>
        )}
      </div>

      {/* ── Center: image + SVG overlay ── */}
      <div style={{ flex: 1, overflow: "auto", display: "flex", flexDirection: "column", alignItems: "center", padding: "1rem", gap: "0.5rem" }}>
        <div style={{ fontSize: "0.65rem", color: "#5a4a20", alignSelf: "flex-start" }}>
          {polyMethod === "bluebeam"
            ? "📐 Bluebeam polygons — " + displayZones.length + " surfaces from estimator markup · SF values exact"
            : polyMethod === "vector"
            ? "📏 Vector mode — " + displayZones.length + " surfaces from PDF geometry"
            : "AI box mode — upload a Bluebeam-marked PDF for exact shapes"}
        </div>

        {!pageImage ? (
          <div style={{ color: "#3a3020", fontSize: "0.75rem", marginTop: "3rem" }}>Loading elevation image...</div>
        ) : (
          <div style={{ position: "relative", display: "inline-block", maxWidth: "100%" }}>
            <img
              ref={imgRef}
              src={pageImage}
              alt={elev?.title || "Elevation"}
              onLoad={e => {
                setImgNaturalSize({ w: e.target.naturalWidth, h: e.target.naturalHeight });
                setImgLoaded(true);
              }}
              style={{ display: "block", maxWidth: "100%", maxHeight: "calc(100vh - 180px)", objectFit: "contain", border: "1px solid #1e1c14" }}
            />
            {imgLoaded && (
              <svg
                viewBox={`0 0 ${pageDims.width} ${pageDims.height}`}
                style={{ position: "absolute", top: 0, left: 0, width: svgW, height: svgH, overflow: "visible" }}
              >
                {displayZones.map(zone => {
                  const a = getAssignment(zone.id);
                  // Color priority: assigned material → Bluebeam fill color → default
                  let color = "#8888aa";
                  if (a) {
                    color = MATERIAL_COLORS[a.category] || "#7a7a7a";
                  } else if (zone.fill_color && zone.fill_color.length === 3) {
                    const [r2, g2, b2] = zone.fill_color;
                    color = "rgb(" + Math.round(r2*255) + "," + Math.round(g2*255) + "," + Math.round(b2*255) + ")";
                  }
                  const isActive = activeZone === zone.id;
                  const pts = toSVGPoints(zone.points);
                  const labelX = zone.cx * pageDims.width;
                  const labelY = zone.cy * pageDims.height;
                  const showLabel = a || isActive || zone.source === "bluebeam";
                  return (
                    <g key={zone.id} style={{ cursor: "pointer" }}
                      onClick={e => { e.stopPropagation(); setActiveZone(isActive ? null : zone.id); }}>
                      <polygon
                        points={pts}
                        fill={color}
                        fillOpacity={isActive ? 0.65 : zone.source === "bluebeam" ? 0.45 : a ? 0.38 : 0.18}
                        stroke={isActive ? "#ffffff" : color}
                        strokeWidth={isActive ? 3 : zone.source === "bluebeam" ? 2 : a ? 2 : 1}
                        strokeOpacity={isActive ? 1 : 0.9}
                      />
                      {showLabel && (
                        <>
                          <rect
                            x={labelX - 30} y={labelY - 9}
                            width={60} height={17}
                            fill="rgba(0,0,0,0.7)" rx={2}
                          />
                          <text
                            x={labelX} y={labelY + 2}
                            textAnchor="middle" dominantBaseline="middle"
                            fill="white" fontSize={pageDims.width / 75}
                            fontFamily="Arial" fontWeight="bold"
                          >
                            {Math.round(zone.area_sf)} SF
                          </text>
                        </>
                      )}
                    </g>
                  );
                })}
              </svg>
            )}
          </div>
        )}
      </div>

      {/* ── Right: material palette + totals ── */}
      <div style={{ width: 210, borderLeft: "1px solid #1e1c14", padding: "0.75rem", overflowY: "auto", flexShrink: 0, background: "#0c0b08" }}>

        {activeZone !== null ? (
          /* Material picker when a zone is selected */
          <>
            <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.5rem" }}>
              Assign Surface
            </div>
            <div style={{ padding: "0.4rem 0.5rem", marginBottom: "0.6rem", background: "#141208", borderRadius: 3, fontSize: "0.65rem", color: "#8a7a50" }}>
              {Math.round(displayZones.find(z => z.id === activeZone)?.area_sf || 0)} SF selected
            </div>

            {matList.map((mat, i) => {
              // Find all same-color zones (Bluebeam color group)
              const activeZoneData = displayZones.find(z => z.id === activeZone);
              const colorKey = activeZoneData?.fill_color ? activeZoneData.fill_color.join(",") : null;
              const sameColorZones = colorKey ? (colorGroups[colorKey] || []) : [];
              const canBulkAssign = sameColorZones.length > 1;

              return (
                <div key={i} style={{ marginBottom: "0.35rem" }}>
                  <div onClick={() => assignZone(activeZone, mat)}
                    style={{ display: "flex", alignItems: "center", gap: "0.5rem", padding: "0.45rem 0.5rem",
                      cursor: "pointer", background: "#141208", borderRadius: canBulkAssign ? "3px 3px 0 0" : 3,
                      border: "1px solid #2a2618", borderBottom: canBulkAssign ? "none" : "1px solid #2a2618" }}
                    onMouseEnter={e => e.currentTarget.style.borderColor = "#5a4a20"}
                    onMouseLeave={e => e.currentTarget.style.borderColor = "#2a2618"}>
                    <div style={{ width: 12, height: 12, borderRadius: 2, background: MATERIAL_COLORS[mat.category] || "#7a7a7a", flexShrink: 0 }} />
                    <div style={{ fontSize: "0.62rem", color: "#a09060", lineHeight: 1.3, flex: 1 }}>
                      <span style={{ color: "#b89020" }}>{mat.id}</span>
                      {" "}<span>{mat.name || mat.category}</span>
                    </div>
                  </div>
                  {canBulkAssign && (
                    <div onClick={() => {
                      sameColorZones.forEach(zid => {
                        const z = displayZones.find(dz => dz.id === zid);
                        setAssignments(prev => ({ ...prev, [elevIdx + ":" + zid]: { ...mat, area_sf: z?.area_sf || 0 } }));
                      });
                      setActiveZone(null);
                    }}
                      style={{ padding: "0.3rem 0.5rem", background: "#0e0d0a", cursor: "pointer",
                        fontSize: "0.58rem", color: "#7a9a40", border: "1px solid #2a2618",
                        borderTop: "1px solid #1a1810", borderRadius: "0 0 3px 3px" }}
                      onMouseEnter={e => e.currentTarget.style.color = "#a0c060"}
                      onMouseLeave={e => e.currentTarget.style.color = "#7a9a40"}>
                      ↳ Assign all {sameColorZones.length} same-color zones
                    </div>
                  )}
                </div>
              );
            })}

            {getAssignment(activeZone) && (
              <div onClick={() => removeAssignment(activeZone)}
                style={{ padding: "0.4rem 0.5rem", marginTop: "0.4rem", textAlign: "center", fontSize: "0.62rem", color: "#cc5050", cursor: "pointer", border: "1px solid #3a1818", borderRadius: 3 }}>
                Remove assignment
              </div>
            )}
            <div onClick={() => setActiveZone(null)}
              style={{ padding: "0.4rem 0.5rem", marginTop: "0.25rem", textAlign: "center", fontSize: "0.62rem", color: "#5a4a20", cursor: "pointer", border: "1px solid #2a2618", borderRadius: 3 }}>
              Cancel
            </div>
          </>
        ) : (
          /* SF totals when nothing selected */
          <>
            <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.6rem" }}>
              SF Totals
            </div>

            {Object.keys(totals).length === 0 ? (
              <div style={{ fontSize: "0.65rem", color: "#3a3020", lineHeight: 1.7 }}>
                Click any highlighted<br />zone to assign material.<br /><br />
                <span style={{ color: "#4a3a18" }}>
                  {useVectorMode
                    ? "Surfaces are from PDF vector data — pixel-accurate."
                    : "Surfaces from AI zone detection."}
                </span>
              </div>
            ) : (
              <>
                {Object.entries(totals).map(([cat, sf]) => (
                  <div key={cat} style={{ marginBottom: "0.6rem" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
                      <div style={{ width: 8, height: 8, borderRadius: 2, background: MATERIAL_COLORS[cat] || "#7a7a7a" }} />
                      <span style={{ fontSize: "0.6rem", color: "#7a6a40" }}>{cat}</span>
                    </div>
                    <div style={{ fontSize: "1rem", fontWeight: 700, color: "#e0cc80", paddingLeft: "1rem" }}>
                      {Math.round(sf).toLocaleString()} SF
                    </div>
                    <div style={{ fontSize: "0.6rem", color: "#5a4a20", paddingLeft: "1rem" }}>
                      {Math.round(sf * 1.15).toLocaleString()} adj +15%
                    </div>
                  </div>
                ))}
                <div style={{ marginTop: "0.75rem", paddingTop: "0.6rem", borderTop: "1px solid #1e1c14" }}>
                  <div style={{ fontSize: "0.6rem", color: "#5a8030" }}>Grand Total</div>
                  <div style={{ fontSize: "1.2rem", fontWeight: 700, color: "#90e060" }}>
                    {Math.round(grandTotal * 1.15).toLocaleString()} SF
                  </div>
                  <div style={{ fontSize: "0.58rem", color: "#3a5020" }}>adjusted +15%</div>
                </div>
              </>
            )}

            <div style={{ marginTop: "1rem", fontSize: "0.58rem", color: "#4a3a18", lineHeight: 1.6 }}>
              {Object.keys(assignments).length} zones assigned<br />
              across {elevations.length} elevations
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function BPSEstimator() {
  const [file, setFile] = useState(null);
  const [phase, setPhase] = useState("idle");
  const [log, setLog] = useState([]);
  const [progress, setProgress] = useState({ label: "", pct: 0 });
  const [results, setResults] = useState(null);
  const [errMsg, setErrMsg] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [jobId, setJobId] = useState(null);
  const [viewMode, setViewMode] = useState("table"); // "table" | "interactive"
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
          {/* View mode toggle */}
          {phase === "done" && results && (
            <div style={{ padding: "0.5rem 1.25rem", borderBottom: "1px solid #1e1c14", display: "flex", gap: "0.5rem", background: "#0a0908", alignItems: "center" }}>
              <span style={{ fontSize: "0.58rem", color: "#4a3a18", letterSpacing: "0.2em", textTransform: "uppercase", marginRight: "0.25rem" }}>View:</span>
              {[["table", "📊 Table"], ["interactive", "🎨 Interactive Takeoff"]].map(([mode, label]) => (
                <button key={mode} onClick={() => setViewMode(mode)}
                  style={{ padding: "0.3rem 0.75rem", background: viewMode === mode ? "#2a2618" : "transparent",
                    color: viewMode === mode ? "#d0b840" : "#5a4a20",
                    border: "1px solid " + (viewMode === mode ? "#5a4a20" : "#1e1c14"),
                    borderRadius: 3, fontSize: "0.65rem", fontFamily: "inherit", cursor: "pointer" }}>
                  {label}
                </button>
              ))}
              {viewMode === "interactive" && (
                <span style={{ fontSize: "0.6rem", color: "#4a3a18", marginLeft: "0.5rem" }}>
                  Click any zone on the elevation → assign material
                </span>
              )}
            </div>
          )}

          {/* Summary cards */}
          {phase === "done" && summary && viewMode === "table" && (
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

          {/* Interactive Takeoff View */}
          {phase === "done" && results && viewMode === "interactive" && (
            <div style={{ flex: 1, overflow: "hidden" }}>
              <InteractiveView results={results} BACKEND={BACKEND} />
            </div>
          )}

          {/* Elevation breakdown */}
          {phase === "done" && results && viewMode === "table" && (
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

