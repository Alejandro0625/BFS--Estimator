import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const BACKEND = import.meta.env.VITE_BACKEND_URL || "";

// BFS Blue from logo: #4A86C8
const BFS_BLUE = "#4A86C8";
const BFS_DARK = "#0F1C2E";
const BFS_BLUE_DARK = "#3A6FA8";
const BFS_BLUE_PALE = "#EBF2FA";

const MATERIAL_COLORS = {
  "ACM Panel":              BFS_BLUE,
  "MCM Panel":              BFS_BLUE,
  "Fiber Cement Panel":     "#22C55E",
  "Fiber Cement Plank":     "#0D9488",
  "Nichiha Panel":          "#8B5CF6",
  "Aluminum Wall Panel":    "#06B6D4",
  "Perforated Metal Panel": "#F97316",
  "Soffit Panel":           "#3B82F6",
  "Return/Trim":            "#EC4899",
  "Other":                  "#9CA3AF",
};

const buildExcel = (projectName, materials) => {
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString("en-US");
  const n11 = () => Array(11).fill(null);
  const n10 = () => Array(10).fill(null);
  const eRows = [];
  let er = n10(); er[0] = projectName || ""; eRows.push(er);
  eRows.push(n10()); eRows.push(n10());
  er = n10(); er[0] = "PANELS"; eRows.push(er);
  er = n10(); er[0] = "No."; er[1] = "ACM/ACP"; er[2] = "Quantity"; er[4] = "Conv"; er[5] = "Rate"; er[6] = "Amount"; eRows.push(er);
  eRows.push(n10());
  const amtCells = [];
  materials.forEach((mat, idx) => {
    er = n10(); er[1] = mat.name; eRows.push(er);
    const itemExcelRow = eRows.length + 1;
    er = n10(); er[0] = idx + 1; er[1] = mat.name; er[2] = Math.round(mat.sf); er[4] = 1; er[5] = ""; er[6] = `=C${itemExcelRow}*F${itemExcelRow}`; amtCells.push(`G${itemExcelRow}`); eRows.push(er);
  });
  while (eRows.length < 24) eRows.push(n10());
  const totalExcelRow = eRows.length + 1;
  er = n10(); er[1] = "Total "; er[2] = materials.reduce((s, m) => s + Math.round(m.sf), 0); er[6] = amtCells.length ? `=${amtCells.join("+")}` : 0; eRows.push(er);
  eRows.push(n10());
  er = n10(); er[0] = "PANEL BACK-UP SYSTEM- Z-Girts, Hat Channel, Insulation"; eRows.push(er);
  eRows.push(n10());
  er = n10(); er[1] = "Furnish and install the quantity of new metal panels required\nAny exterior caulking required\nLifts / tie-off required per site policy\nAny break metal/flashing required\nStructural calculations and PE stamp\nShop drawings\nTaxes"; eRows.push(er);
  while (eRows.length < 34) eRows.push(n10());
  er = n10(); er[0] = "SPECIFICATIONS"; eRows.push(er);
  eRows.push(n10());
  ["GC", "Location", "Profit/Non-Profit", "Taxable/Non-Taxable", "Prevailing Wage", "Drawing Set", "Building Height"].forEach(f => { er = n10(); er[0] = f; eRows.push(er); });
  const wsE = XLSX.utils.aoa_to_sheet(eRows);
  wsE["!cols"] = [5.14, 57.43, 11.86, 4.14, 5.43, 12.86, 12.43].map(w => ({ wch: w }));
  wsE["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 7 } }, { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 0 }, e: { r: 3, c: 6 } }, { s: { r: totalExcelRow - 1, c: 0 }, e: { r: totalExcelRow - 1, c: 9 } }, { s: { r: eRows.length - 8, c: 0 }, e: { r: eRows.length - 8, c: 6 } }];
  XLSX.utils.book_append_sheet(wb, wsE, "Estimate");
  const pRows = [];
  let pr = n11(); pr[5] = "PROPOSAL"; pRows.push(pr);
  pr = n11(); pr[5] = "DATE:"; pr[6] = today; pRows.push(pr);
  pr = n11(); pr[0] = "ACM. Trespa. Terracotta  &  Specialty Metal Panels"; pr[5] = "This proposal may be withdrawn by us if not accepted within 30 days."; pRows.push(pr);
  pr = n11(); pr[0] = "15 Erie Drive"; pr[5] = "E-mail:"; pRows.push(pr);
  pr = n11(); pr[0] = "Natick, MA 01760"; pr[5] = ""; pRows.push(pr);
  pr = n11(); pr[0] = "PH: 617-458-2000  "; pr[5] = "Phone:"; pRows.push(pr);
  pr = n11(); pr[0] = "To:"; pr[1] = ""; pr[5] = ""; pRows.push(pr);
  pr = n11(); pr[1] = ""; pr[5] = "Job Name / location:"; pRows.push(pr);
  pr = n11(); pr[1] = ""; pr[5] = projectName || ""; pRows.push(pr);
  pr = n11(); pr[1] = ""; pr[5] = "Job number: "; pRows.push(pr);
  pr = n11(); pr[5] = ""; pRows.push(pr);
  pr = n11(); pr[0] = "We hereby submit specifications and estimates for:"; pRows.push(pr);
  const mainMat = materials[0]?.name || "[Material]";
  const totalSF = materials.reduce((s, m) => s + Math.round(m.sf), 0);
  pr = n11(); pr[1] = `Install ${totalSF.toLocaleString()}sf of ${mainMat}.`; pRows.push(pr);
  const scopeItems = ["Include all OSHA and fall protection compliance for the installation of panels", "Include all staging and lifts for the performance of work.", `F&I ${materials.map(m => m.name).join(", ")} as specified.`, "F&I all metal trim and accessories with panels as specified.", "Remove and dispose of all job related debris to the general contractor's dumpster.", "MA Sales Tax Included on all materials if applicable."];
  scopeItems.forEach((item, i) => { pr = n11(); pr[0] = i + 1; pr[1] = item; pRows.push(pr); });
  pr = n11(); pr[1] = "ADD/ALT: ENGINEERING DESIGN AND CALCULATIONS"; pr[7] = ": $4,500"; pRows.push(pr);
  pr = n11(); pr[1] = "NOTE: Air Vapor barrier behind all exterior panel system not included"; pRows.push(pr);
  pr = n11(); pr[0] = "NOTE: THIS IS A BUDGETARY NUMBER ONLY PENDING FINAL SCOPE REVIEW & ENGINEERING CRITERIA"; pRows.push(pr);
  pr = n11(); pr[2] = "PRICING GOOD FOR 30 DAYS DUE TO INDUSTRY-WIDE PRICE ESCALATION"; pRows.push(pr);
  pr = n11(); pr[1] = "NIC: blocking, framing, plywood substrate, police details & street permits,"; pRows.push(pr);
  pr = n11(); pr[1] = "thru-wall flashings, flashings not associated with the panel installations,"; pRows.push(pr);
  pr = n11(); pr[1] = " custom colors* (except where noted), winter conditions"; pRows.push(pr);
  pr = n11(); pr[1] = '*** all contracts to have "BPS conditions for Metal Panels/Siding" attached.'; pRows.push(pr);
  pRows.push(n11());
  pr = n11(); pr[0] = "We propose hereby to furnish materials and labor - complete in accordance with above specifications for the sum of:"; pRows.push(pr);
  pr = n11(); pr[7] = "TOTAL:"; pr[8] = `=Estimate!G${totalExcelRow}`; pRows.push(pr);
  pr = n11(); pr[0] = "Payment to be made as follows:"; pr[4] = "AIA Format"; pRows.push(pr);
  pr = n11(); pr[6] = "Akshita Patel"; pRows.push(pr);
  pr = n11(); pr[7] = "Authorized Signature"; pRows.push(pr);
  pr = n11(); pr[0] = "All material to be as specified. All work to be performed in a professional manner according to standard practices. Any alteration or deviation from above specifications involving additional costs will be executed only upon written orders and will be an extra charge. Payment is due in full within thirty (30) days of the date of the invoice. Interest will be charged on outstanding balances at twelve percent (12%) per annum."; pRows.push(pr);
  pr = n11(); pr[0] = "Acceptance of Proposal — The above prices, specifications and conditions are satisfactory and hereby accepted. You are authorized to do the work as specified. Payment as outlined above."; pRows.push(pr);
  pr = n11(); pr[3] = "Date:"; pr[7] = "Signature"; pRows.push(pr);
  const wsP = XLSX.utils.aoa_to_sheet(pRows);
  wsP["!cols"] = [7.14, 9.14, 12, 12, 16.71, 9.57, 9.14, 12, 12, 12, 12].map(w => ({ wch: w }));
  wsP["!merges"] = [{ s: { r: 0, c: 5 }, e: { r: 0, c: 10 } }, { s: { r: 1, c: 6 }, e: { r: 1, c: 7 } }, { s: { r: 2, c: 5 }, e: { r: 2, c: 10 } }, { s: { r: 3, c: 5 }, e: { r: 3, c: 10 } }, { s: { r: 4, c: 5 }, e: { r: 4, c: 10 } }, { s: { r: 5, c: 5 }, e: { r: 5, c: 10 } }, { s: { r: 6, c: 5 }, e: { r: 6, c: 10 } }, { s: { r: 7, c: 5 }, e: { r: 7, c: 10 } }, { s: { r: 8, c: 5 }, e: { r: 8, c: 10 } }, { s: { r: 9, c: 5 }, e: { r: 9, c: 10 } }, { s: { r: 10, c: 5 }, e: { r: 10, c: 10 } }, { s: { r: pRows.length - 4, c: 0 }, e: { r: pRows.length - 3, c: 5 } }];
  XLSX.utils.book_append_sheet(wb, wsP, "Proposal");
  return wb;
};

const CLUSTER_COLORS = ["#3B82F6","#F97316","#22C55E","#EC4899","#EAB308","#8B5CF6","#14B8A6","#EF4444"];

// ─── Shared style helpers ────────────────────────────────────────────────────
const card = (extra = {}) => ({
  background: "#fff",
  borderRadius: 10,
  boxShadow: "0 1px 4px rgba(0,0,0,0.08), 0 0 0 1px rgba(0,0,0,0.04)",
  ...extra,
});

const pill = (active, extra = {}) => ({
  padding: "0.35rem 0.9rem",
  borderRadius: 6,
  fontSize: "0.72rem",
  fontWeight: active ? 600 : 400,
  cursor: "pointer",
  border: "none",
  fontFamily: "inherit",
  background: active ? BFS_BLUE : "#F1F5F9",
  color: active ? "#fff" : "#64748B",
  transition: "all 0.15s",
  ...extra,
});

const btn = (variant = "primary", disabled = false, extra = {}) => {
  const base = { padding: "0.65rem 1.1rem", borderRadius: 8, fontSize: "0.78rem", fontWeight: 600, fontFamily: "inherit", cursor: disabled ? "not-allowed" : "pointer", border: "none", display: "flex", alignItems: "center", justifyContent: "center", gap: "0.4rem", transition: "all 0.15s", width: "100%", letterSpacing: "0.01em", ...extra };
  if (disabled) return { ...base, background: "#E2E8F0", color: "#94A3B8" };
  if (variant === "primary") return { ...base, background: BFS_BLUE, color: "#fff", boxShadow: "0 2px 8px rgba(74,134,200,0.35)" };
  if (variant === "outline") return { ...base, background: "#fff", color: BFS_BLUE, border: `1.5px solid ${BFS_BLUE}` };
  if (variant === "ghost") return { ...base, background: "#F8FAFC", color: "#475569", border: "1.5px solid #E2E8F0" };
  return base;
};

// ─── Interactive Takeoff ─────────────────────────────────────────────────────
function InteractiveView({ results, BACKEND }) {
  const [elevIdx, setElevIdx] = useState(0);
  const [pageImage, setPageImage] = useState(null);
  const [imgLoaded, setImgLoaded] = useState(false);
  const [pagePolygons, setPagePolygons] = useState([]);
  const [pageDims, setPageDims] = useState({ width: 612, height: 792 });
  const [assignments, setAssignments] = useState({});
  const [activeZone, setActiveZone] = useState(null);
  const [imgNaturalSize, setImgNaturalSize] = useState({ w: 1, h: 1 });
  const imgRef = useRef();

  const elevations = results.takeoffData.filter(e => e.pageNumber);
  const elev = elevations[elevIdx];
  const pageNum = elev?.pageNumber;

  useEffect(() => {
    if (!pageNum || !results.jobId) return;
    setPageImage(null); setImgLoaded(false); setPagePolygons([]); setActiveZone(null);
    fetch(BACKEND + "/polygons/" + results.jobId + "/" + pageNum)
      .then(r => r.ok ? r.json() : { polygons: [], width: 612, height: 792 })
      .then(d => { setPagePolygons(d.polygons || []); setPageDims({ width: d.width || 612, height: d.height || 792 }); })
      .catch(() => {});
    setPageImage(BACKEND + "/page-image/" + results.jobId + "/" + pageNum);
  }, [elevIdx, pageNum, results.jobId, BACKEND]);

  const polyMethod = pagePolygons[0]?.source || (pagePolygons.length > 0 ? "vector" : "box");
  const useVectorMode = pagePolygons.length > 0;
  const displayZones = useVectorMode ? pagePolygons : (elev?.zones || []).map((z, i) => ({
    id: i, points: [[z.x0pct/100,z.y0pct/100],[z.x1pct/100,z.y0pct/100],[z.x1pct/100,z.y1pct/100],[z.x0pct/100,z.y1pct/100]],
    area_sf: z.netArea || 0, cx: (z.x0pct+z.x1pct)/200, cy: (z.y0pct+z.y1pct)/200, source: "box",
    suggestedCategory: z.category, suggestedId: z.materialId, suggestedName: z.materialName,
  }));

  const colorGroups = {};
  displayZones.forEach(z => {
    const key = z.cluster_id !== undefined ? "c_"+z.cluster_id : z.fill_color ? "f_"+z.fill_color.join(",") : "none";
    if (!colorGroups[key]) colorGroups[key] = [];
    colorGroups[key].push(z.id);
  });

  const clusterSummary = {};
  displayZones.forEach(z => {
    const k = z.cluster_id !== undefined ? z.cluster_id : -1;
    if (!clusterSummary[k]) clusterSummary[k] = { total_sf: 0, count: 0, color: "#94A3B8" };
    clusterSummary[k].total_sf += z.area_sf || 0;
    clusterSummary[k].count += 1;
    clusterSummary[k].color = z.cluster_id !== undefined ? CLUSTER_COLORS[z.cluster_id % CLUSTER_COLORS.length] : "#94A3B8";
  });

  const assignKey = id => elevIdx + ":" + id;
  const getAssignment = id => assignments[assignKey(id)];
  const assignZone = (zoneId, mat) => {
    const zone = displayZones.find(z => z.id === zoneId);
    setAssignments(prev => ({ ...prev, [assignKey(zoneId)]: { ...mat, area_sf: zone?.area_sf || 0 } }));
    setActiveZone(null);
  };
  const removeAssignment = zoneId => {
    const k = assignKey(zoneId);
    setAssignments(prev => { const n = { ...prev }; delete n[k]; return n; });
  };
  const exportInteractiveExcel = () => {
    const matTotals = {};
    Object.values(assignments).forEach(a => {
      const key = a.materialName || a.category || "Panel";
      if (!matTotals[key]) matTotals[key] = { name: key, sf: 0 };
      matTotals[key].sf += a.area_sf || 0;
    });
    const wb = buildExcel(results.projName || "Project", Object.values(matTotals));
    XLSX.writeFile(wb, "BFS_Takeoff_" + (results.projName||"Project").replace(/\s+/g,"_") + ".xlsx");
  };

  const totals = {};
  Object.values(assignments).forEach(a => { if (!totals[a.category]) totals[a.category] = 0; totals[a.category] += a.area_sf || 0; });
  const grandTotal = Object.values(totals).reduce((s,v) => s+v, 0);
  const svgW = imgRef.current?.offsetWidth || imgNaturalSize.w;
  const svgH = imgRef.current?.offsetHeight || imgNaturalSize.h;
  const toSVGPoints = pts => pts.map(([nx,ny]) => `${(nx*pageDims.width).toFixed(1)},${(ny*pageDims.height).toFixed(1)}`).join(" ");
  const matList = results.legend.length > 0 ? results.legend : Object.keys(MATERIAL_COLORS).map(cat => ({ id: cat.substring(0,3).toUpperCase()+"-1", name: cat, category: cat }));

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden", background: "#0D1B2E", fontFamily: "'Inter','Segoe UI',sans-serif" }}>

      {/* Left: elevation list */}
      <div style={{ width: 190, borderRight: "1px solid #1E3A5F", overflowY: "auto", flexShrink: 0, background: "#0A1628" }}>
        <div style={{ padding: "0.875rem", fontSize: "0.6rem", letterSpacing: "0.12em", color: "#64748B", textTransform: "uppercase", fontWeight: 700, borderBottom: "1px solid #1E3A5F" }}>
          Elevations
        </div>
        {elevations.map((e, i) => {
          const assigned = Object.keys(assignments).filter(k => k.startsWith(i+":")).length;
          return (
            <div key={i} onClick={() => setElevIdx(i)}
              style={{ padding: "0.65rem 0.875rem", cursor: "pointer", borderBottom: "1px solid #1E3A5F",
                background: i === elevIdx ? "#1E3A5F" : "transparent",
                borderLeft: i === elevIdx ? `3px solid ${BFS_BLUE}` : "3px solid transparent" }}>
              <div style={{ fontSize: "0.72rem", color: i === elevIdx ? "#E2E8F0" : "#94A3B8", fontWeight: i === elevIdx ? 600 : 400, lineHeight: 1.3 }}>
                {e.title || "Page " + e.pageNumber}
              </div>
              <div style={{ fontSize: "0.62rem", color: assigned > 0 ? "#4ADE80" : "#475569", marginTop: 3 }}>
                {assigned > 0 ? `✓ ${assigned} assigned` : `${(e.zones||[]).length} zones · p.${e.pageNumber}`}
              </div>
            </div>
          );
        })}
        {elevations.length === 0 && <div style={{ padding: "1rem", fontSize: "0.65rem", color: "#475569" }}>No elevations found.</div>}
      </div>

      {/* Center: drawing */}
      <div style={{ flex: 1, overflow: "auto", display: "flex", flexDirection: "column", alignItems: "center", padding: "1rem", gap: "0.75rem" }}>
        <div style={{ alignSelf: "flex-start", fontSize: "0.65rem", color: "#64748B", background: "#1E3A5F", padding: "0.3rem 0.75rem", borderRadius: 20, border: "1px solid #2D5280" }}>
          {polyMethod === "bluebeam" ? `📐 Bluebeam — ${displayZones.length} surfaces`
            : polyMethod === "vector_cluster" || polyMethod === "vector" ? `📏 Vector — ${displayZones.length} surfaces`
            : polyMethod === "claude_vision" ? `🧠 AI Vision — ${displayZones.length} surfaces detected`
            : "No surfaces on this page"}
        </div>
        {!pageImage
          ? <div style={{ color: "#475569", fontSize: "0.8rem", marginTop: "4rem" }}>Loading elevation...</div>
          : <div style={{ position: "relative", display: "inline-block", maxWidth: "100%" }}>
              <img ref={imgRef} src={pageImage} alt={elev?.title}
                onLoad={e => { setImgNaturalSize({ w: e.target.naturalWidth, h: e.target.naturalHeight }); setImgLoaded(true); }}
                style={{ display: "block", maxWidth: "100%", maxHeight: "calc(100vh - 180px)", objectFit: "contain", borderRadius: 6, border: "1px solid #1E3A5F" }}
              />
              {imgLoaded && (
                <svg viewBox={`0 0 ${pageDims.width} ${pageDims.height}`}
                  style={{ position: "absolute", top: 0, left: 0, width: svgW, height: svgH, overflow: "visible" }}>
                  {displayZones.map(zone => {
                    const a = getAssignment(zone.id);
                    let color = "#94A3B8";
                    if (a) color = MATERIAL_COLORS[a.category] || "#9CA3AF";
                    else if (zone.cluster_id !== undefined) color = CLUSTER_COLORS[zone.cluster_id % CLUSTER_COLORS.length];
                    else if (zone.fill_color?.length === 3) { const [r,g,b] = zone.fill_color; color = `rgb(${Math.round(r*255)},${Math.round(g*255)},${Math.round(b*255)})`; }
                    const isActive = activeZone === zone.id;
                    const pts = toSVGPoints(zone.points);
                    const lx = zone.cx * pageDims.width, ly = zone.cy * pageDims.height;
                    const showLabel = a || isActive || zone.source === "bluebeam" || zone.source === "claude_vision";
                    return (
                      <g key={zone.id} style={{ cursor: "pointer" }} onClick={e => { e.stopPropagation(); setActiveZone(isActive ? null : zone.id); }}>
                        <polygon points={pts} fill={color} fillOpacity={isActive ? 0.65 : zone.source==="bluebeam" ? 0.45 : a ? 0.38 : 0.2}
                          stroke={isActive ? "#fff" : color} strokeWidth={isActive ? 2.5 : 1.5} strokeOpacity={0.9} />
                        {showLabel && (<>
                          <rect x={lx-32} y={ly-9} width={64} height={18} fill="rgba(0,0,0,0.8)" rx={4} />
                          <text x={lx} y={ly+2} textAnchor="middle" dominantBaseline="middle" fill="white" fontSize={pageDims.width/75} fontFamily="Inter,Arial" fontWeight="bold">
                            {zone.source==="claude_vision"&&!a ? zone.material_type : Math.round(zone.area_sf)+" SF"}
                          </text>
                        </>)}
                      </g>
                    );
                  })}
                </svg>
              )}
            </div>
        }
      </div>

      {/* Right: palette + totals */}
      <div style={{ width: 230, borderLeft: "1px solid #1E3A5F", padding: "1rem", overflowY: "auto", flexShrink: 0, background: "#0A1628" }}>
        {activeZone !== null ? (
          <>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: "#64748B", textTransform: "uppercase", fontWeight: 700, marginBottom: "0.5rem" }}>Assign Surface</div>
            <div style={{ padding: "0.5rem 0.65rem", marginBottom: "0.75rem", background: "#1E3A5F", borderRadius: 6, fontSize: "0.7rem", color: "#CBD5E1", border: "1px solid #2D5280" }}>
              {Math.round(displayZones.find(z => z.id === activeZone)?.area_sf || 0).toLocaleString()} SF selected
            </div>
            {matList.map((mat, i) => {
              const azd = displayZones.find(z => z.id === activeZone);
              const ck = azd?.cluster_id !== undefined ? "c_"+azd.cluster_id : azd?.fill_color ? "f_"+azd.fill_color.join(",") : null;
              const scz = ck ? (colorGroups[ck]||[]) : [];
              const bulk = scz.length > 1;
              return (
                <div key={i} style={{ marginBottom: "0.4rem" }}>
                  <div onClick={() => assignZone(activeZone, mat)}
                    style={{ display: "flex", alignItems: "center", gap: "0.5rem", padding: "0.5rem 0.65rem", cursor: "pointer", background: "#1E3A5F",
                      borderRadius: bulk ? "6px 6px 0 0" : 6, border: "1px solid #2D5280", borderBottom: bulk ? "none" : undefined }}
                    onMouseEnter={e => e.currentTarget.style.borderColor = BFS_BLUE}
                    onMouseLeave={e => e.currentTarget.style.borderColor = "#2D5280"}>
                    <div style={{ width: 10, height: 10, borderRadius: 3, background: MATERIAL_COLORS[mat.category]||"#9CA3AF", flexShrink: 0 }} />
                    <div style={{ fontSize: "0.65rem", color: "#CBD5E1", lineHeight: 1.3, flex: 1 }}>
                      <span style={{ color: BFS_BLUE, fontWeight: 700 }}>{mat.id}</span> {mat.name||mat.category}
                    </div>
                  </div>
                  {bulk && (
                    <div onClick={() => { scz.forEach(zid => { const z=displayZones.find(dz=>dz.id===zid); setAssignments(prev=>({...prev,[elevIdx+":"+zid]:{...mat,area_sf:z?.area_sf||0}})); }); setActiveZone(null); }}
                      style={{ padding: "0.3rem 0.65rem", background: "#122035", cursor: "pointer", fontSize: "0.62rem", color: "#4ADE80", border: "1px solid #2D5280", borderTop: "none", borderRadius: "0 0 6px 6px" }}
                      onMouseEnter={e => e.currentTarget.style.color="#86EFAC"} onMouseLeave={e => e.currentTarget.style.color="#4ADE80"}>
                      ↳ Assign all {scz.length} same-color zones
                    </div>
                  )}
                </div>
              );
            })}
            {getAssignment(activeZone) && (
              <div onClick={() => removeAssignment(activeZone)}
                style={{ padding: "0.45rem", marginTop: "0.5rem", textAlign: "center", fontSize: "0.65rem", color: "#F87171", cursor: "pointer", border: "1px solid #7F1D1D", borderRadius: 6 }}>
                Remove assignment
              </div>
            )}
            <div onClick={() => setActiveZone(null)}
              style={{ padding: "0.45rem", marginTop: "0.3rem", textAlign: "center", fontSize: "0.65rem", color: "#64748B", cursor: "pointer", border: "1px solid #2D5280", borderRadius: 6 }}>
              Cancel
            </div>
          </>
        ) : (
          <>
            {Object.keys(clusterSummary).length > 0 && (<>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: "#64748B", textTransform: "uppercase", fontWeight: 700, marginBottom: "0.4rem" }}>Detected Textures</div>
              <div style={{ fontSize: "0.62rem", color: "#475569", marginBottom: "0.65rem" }}>Click a zone → assign material</div>
              {Object.entries(clusterSummary).map(([cid, info]) => {
                const zids = colorGroups["c_"+cid]||[];
                const assigned = zids.filter(id => getAssignment(id)).length;
                return (
                  <div key={cid} style={{ marginBottom: "0.4rem", padding: "0.5rem 0.65rem", background: "#1E3A5F", borderRadius: 6, borderLeft: `3px solid ${info.color}` }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
                        <div style={{ width: 8, height: 8, borderRadius: 2, background: info.color }} />
                        <span style={{ fontSize: "0.65rem", color: "#94A3B8" }}>Texture {parseInt(cid)+1}</span>
                      </div>
                      <span style={{ fontSize: "0.8rem", fontWeight: 700, color: "#E2E8F0" }}>{Math.round(info.total_sf).toLocaleString()}</span>
                    </div>
                    <div style={{ fontSize: "0.6rem", color: assigned>0?"#4ADE80":"#475569", marginTop: 3 }}>
                      {assigned>0 ? `✓ ${assigned}/${info.count} assigned` : `${info.count} zone${info.count!==1?"s":""}`}
                    </div>
                  </div>
                );
              })}
              <div style={{ height: 1, background: "#1E3A5F", margin: "0.75rem 0" }} />
            </>)}

            {Object.keys(totals).length > 0 && (<>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: "#64748B", textTransform: "uppercase", fontWeight: 700, marginBottom: "0.5rem" }}>Your Takeoff</div>
              {Object.entries(totals).map(([cat, sf]) => (
                <div key={cat} style={{ marginBottom: "0.5rem", padding: "0.5rem 0.65rem", background: "#1E3A5F", borderRadius: 6, borderLeft: `3px solid ${MATERIAL_COLORS[cat]||"#9CA3AF"}` }}>
                  <div style={{ fontSize: "0.62rem", color: "#94A3B8" }}>{cat}</div>
                  <div style={{ fontSize: "1.1rem", fontWeight: 700, color: "#E2E8F0" }}>{Math.round(sf).toLocaleString()} <span style={{ fontSize: "0.65rem", fontWeight: 400 }}>SF net</span></div>
                  <div style={{ fontSize: "0.6rem", color: "#64748B" }}>{Math.round(sf*1.15).toLocaleString()} SF +15%</div>
                </div>
              ))}
              <div style={{ padding: "0.65rem", background: "#1E3A5F", borderRadius: 6, border: `1px solid ${BFS_BLUE}40`, marginBottom: "0.75rem" }}>
                <div style={{ fontSize: "0.6rem", color: "#4ADE80", fontWeight: 700 }}>GRAND TOTAL</div>
                <div style={{ fontSize: "1.4rem", fontWeight: 700, color: "#4ADE80" }}>{Math.round(grandTotal*1.15).toLocaleString()} <span style={{ fontSize: "0.65rem", fontWeight: 400 }}>SF</span></div>
                <div style={{ fontSize: "0.6rem", color: "#475569" }}>+15% waste factor</div>
              </div>
              <button onClick={exportInteractiveExcel}
                style={{ width: "100%", padding: "0.65rem", background: BFS_BLUE, color: "#fff", border: "none", borderRadius: 7, fontSize: "0.72rem", fontWeight: 700, fontFamily: "inherit", cursor: "pointer" }}>
                ↓ Export Excel
              </button>
            </>)}

            {Object.keys(clusterSummary).length===0 && Object.keys(totals).length===0 && (
              <div style={{ fontSize: "0.7rem", color: "#475569", lineHeight: 1.8, paddingTop: "0.5rem" }}>
                Click any colored zone on the elevation to assign a material type.
              </div>
            )}
            <div style={{ marginTop: "1rem", fontSize: "0.6rem", color: "#334155" }}>
              {Object.keys(assignments).length} zones · {elevations.length} elevations
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── Main App ────────────────────────────────────────────────────────────────
export default function BFSEstimator() {
  const [file, setFile] = useState(null);
  const [phase, setPhase] = useState("idle");
  const [log, setLog] = useState([]);
  const [progress, setProgress] = useState({ label: "", pct: 0 });
  const [results, setResults] = useState(null);
  const [errMsg, setErrMsg] = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [viewMode, setViewMode] = useState("table");
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef();
  const logRef = useRef();
  const pollRef = useRef(null);
  const seenLogs = useRef(0);

  useEffect(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, [log]);
  useEffect(() => () => { if (pollRef.current) clearInterval(pollRef.current); }, []);

  const handleFile = f => {
    if (f?.type === "application/pdf") {
      setFile(f); setPhase("idle"); setResults(null); setLog([]); setErrMsg(""); seenLogs.current = 0;
    }
  };

  const startPolling = useCallback((id) => {
    if (pollRef.current) clearInterval(pollRef.current);
    seenLogs.current = 0;
    pollRef.current = setInterval(async () => {
      try {
        const res = await fetch(BACKEND + "/status/" + id);
        const data = await res.json();
        if (data.log?.length > seenLogs.current) { setLog(prev => [...prev, ...data.log.slice(seenLogs.current)]); seenLogs.current = data.log.length; }
        if (data.progress) setProgress(data.progress);
        if (data.phase) setPhase(data.phase);
        if (data.status === "done") {
          clearInterval(pollRef.current);
          setResults({ legend: data.legend||[], takeoffData: data.takeoffData||[], projName: file?.name?.replace(".pdf","")||"Project", jobId: id });
          setPhase("done"); setProgress({ label: "Complete", pct: 100 });
        } else if (data.status === "error") { clearInterval(pollRef.current); setErrMsg(data.error||"Unknown error"); setPhase("error"); }
      } catch (e) { console.log("poll hiccup", e.message); }
    }, 5000);
  }, [file]);

  const run = async () => {
    if (!file) return;
    setPhase("running"); setLog([]); setErrMsg(""); setResults(null); seenLogs.current = 0;
    try {
      setLog([{ msg: "Uploading PDF...", level: "info" }]);
      const fd = new FormData(); fd.append("pdf", file);
      const res = await fetch(BACKEND + "/analyze", { method: "POST", body: fd });
      const { jobId: id } = await res.json();
      setLog(prev => [...prev, { msg: "Analysis started — job " + id, level: "ok" }]);
      startPolling(id);
    } catch (err) { setErrMsg(err.message); setPhase("error"); }
  };

  const exportExcel = () => {
    if (!results) return;
    const mt = {};
    results.takeoffData.forEach(e => (e.zones||[]).forEach(z => { const k=z.materialName||z.category||"Panel"; if(!mt[k])mt[k]={name:k,sf:0}; mt[k].sf+=z.netArea||0; }));
    const wb = buildExcel(results.projName||"Project", Object.values(mt));
    XLSX.writeFile(wb, "BFS_Takeoff_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
  };

  const exportPDF = async () => {
    if (!results?.jobId) return;
    setPdfLoading(true);
    try {
      const res = await fetch(BACKEND + "/evidence-pdf/" + results.jobId);
      if (!res.ok) throw new Error("PDF not ready");
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href=url; a.download="BFS_Evidence_"+(results.projName||"Project").replace(/\s+/g,"_")+".pdf"; a.click();
      URL.revokeObjectURL(url);
    } catch (err) { alert("PDF export failed: "+err.message); }
    finally { setPdfLoading(false); }
  };

  const summary = results ? (() => {
    const t = {};
    results.takeoffData.forEach(e => (e.zones||[]).forEach(z => {
      const k = z.category||"Other";
      if (!t[k]) t[k] = { net:0, adj:0, color: MATERIAL_COLORS[k]||"#9CA3AF" };
      t[k].net += z.netArea||0; t[k].adj += (z.netArea||0)*1.15;
    }));
    return t;
  })() : null;

  const grandAdj = summary ? Object.values(summary).reduce((s,v) => s+v.adj, 0) : 0;
  const phaseStep = { idle:0, running:1, filtering:1, legend:2, analyzing:3, done:4, error:0 }[phase]||0;
  const isRunning = !["idle","done","error"].includes(phase);
  const logColor = { ok:"#22C55E", warn:"#F59E0B", error:"#EF4444", success:"#22C55E", dim:"#94A3B8", info:"#64748B" };

  const steps = [
    { n:1, label:"Read Sheet Index" },
    { n:2, label:"Read Material Legend" },
    { n:3, label:"Analyze Elevations" },
    { n:4, label:"Export Results" },
  ];

  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',-apple-system,sans-serif", background:"#F0F4F8", minHeight:"100vh", display:"flex", flexDirection:"column", color:"#1E293B" }}>

      {/* ── Header ── */}
      <header style={{ background: BFS_DARK, height: 64, padding: "0 1.5rem", display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, boxShadow:"0 2px 12px rgba(0,0,0,0.3)" }}>
        <div style={{ display:"flex", alignItems:"center", gap:"0.875rem" }}>
          <img src="/logo-bfs.png" alt="BFS" style={{ height: 42, width:"auto" }} />
          <div style={{ width:1, height:36, background:"rgba(255,255,255,0.12)" }} />
          <div>
            <div style={{ fontSize:"0.6rem", letterSpacing:"0.18em", color:"rgba(255,255,255,0.5)", textTransform:"uppercase", fontWeight:500, lineHeight:1 }}>Boston Facade Systems</div>
            <div style={{ fontSize:"1.05rem", fontWeight:700, color:"#fff", lineHeight:1.3, letterSpacing:"-0.01em" }}>AI Panel Estimator</div>
          </div>
        </div>
        <div style={{ fontSize:"0.65rem", color:"rgba(255,255,255,0.35)", textAlign:"right", lineHeight:1.8 }}>
          Panels · Soffits · Returns<br/>Waste factor: 15%
        </div>
      </header>

      <div style={{ display:"flex", flex:1, overflow:"hidden" }}>

        {/* ── Sidebar ── */}
        <aside style={{ width:295, background:"#fff", borderRight:"1px solid #E2E8F0", padding:"1.5rem 1.25rem", display:"flex", flexDirection:"column", gap:"1.5rem", overflowY:"auto", flexShrink:0 }}>

          {/* Upload */}
          <div>
            <div style={{ fontSize:"0.65rem", letterSpacing:"0.1em", color:BFS_BLUE, textTransform:"uppercase", fontWeight:700, marginBottom:"0.65rem" }}>Blueprint PDF</div>
            <div
              onClick={() => fileRef.current?.click()}
              onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              style={{ border: `2px dashed ${dragOver ? BFS_BLUE : file ? "#22C55E" : "#CBD5E1"}`, borderRadius:10, padding:"1.5rem 1rem", textAlign:"center", cursor:"pointer",
                background: dragOver ? BFS_BLUE_PALE : file ? "#F0FDF4" : "#FAFBFC", transition:"all 0.2s" }}>
              <div style={{ fontSize:"2.2rem", marginBottom:"0.4rem" }}>{file ? "📋" : "📂"}</div>
              {file ? (<>
                <div style={{ fontSize:"0.75rem", fontWeight:600, color:"#16A34A", wordBreak:"break-all" }}>{file.name}</div>
                <div style={{ fontSize:"0.65rem", color:"#86EFAC", marginTop:3 }}>{(file.size/1e6).toFixed(1)} MB · Ready</div>
              </>) : (<>
                <div style={{ fontSize:"0.78rem", fontWeight:500, color:"#64748B" }}>Drop PDF here</div>
                <div style={{ fontSize:"0.68rem", color:"#94A3B8", marginTop:3 }}>or click to browse</div>
              </>)}
            </div>
            <input ref={fileRef} type="file" accept=".pdf" onChange={e => handleFile(e.target.files[0])} style={{ display:"none" }} />
          </div>

          {/* Steps */}
          <div>
            <div style={{ fontSize:"0.65rem", letterSpacing:"0.1em", color:BFS_BLUE, textTransform:"uppercase", fontWeight:700, marginBottom:"0.75rem" }}>Process Steps</div>
            <div style={{ display:"flex", flexDirection:"column", gap:"0.5rem" }}>
              {steps.map(({ n, label }) => {
                const done = phaseStep > n, active = phaseStep === n;
                return (
                  <div key={n} style={{ display:"flex", alignItems:"center", gap:"0.65rem" }}>
                    <div style={{ width:26, height:26, borderRadius:"50%", flexShrink:0, display:"flex", alignItems:"center", justifyContent:"center",
                      fontSize:"0.65rem", fontWeight:700, transition:"all 0.2s",
                      background: done ? "#22C55E" : active ? BFS_BLUE : "#F1F5F9",
                      color: done||active ? "#fff" : "#94A3B8",
                      boxShadow: active ? `0 0 0 3px ${BFS_BLUE}25` : "none" }}>
                      {done ? "✓" : n}
                    </div>
                    <span style={{ fontSize:"0.72rem", color: done ? "#16A34A" : active ? BFS_BLUE : "#94A3B8", fontWeight: done||active ? 600 : 400 }}>
                      {label}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Progress */}
          {progress.pct > 0 && progress.pct < 100 && (
            <div>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:"0.4rem" }}>
                <span style={{ fontSize:"0.68rem", color:"#64748B" }}>{progress.label}</span>
                <span style={{ fontSize:"0.68rem", fontWeight:600, color:BFS_BLUE }}>{Math.round(progress.pct)}%</span>
              </div>
              <div style={{ background:"#E2E8F0", borderRadius:8, height:7, overflow:"hidden" }}>
                <div style={{ width:progress.pct+"%", height:"100%", background:BFS_BLUE, borderRadius:8, transition:"width 0.4s" }} />
              </div>
            </div>
          )}

          {/* Buttons */}
          <div style={{ display:"flex", flexDirection:"column", gap:"0.5rem" }}>
            <button onClick={run} disabled={!file||isRunning} style={btn("primary", !file||isRunning)}>
              {phase==="idle"||phase==="error" ? <><span>▶</span> Run Analysis</> : phase==="done" ? <><span>↺</span> Run Again</> : <><span>⏳</span> Analyzing...</>}
            </button>
            {phase==="done" && (<>
              <button onClick={exportExcel} style={btn("outline")}>
                <span>↓</span> Export Excel
              </button>
              <button onClick={exportPDF} disabled={pdfLoading} style={btn("ghost", pdfLoading)}>
                <span>{pdfLoading?"⏳":"↓"}</span> {pdfLoading?"Generating PDF...":"Export Evidence PDF"}
              </button>
            </>)}
          </div>

          {phase==="error" && errMsg && (
            <div style={{ padding:"0.75rem", background:"#FEF2F2", border:"1px solid #FECACA", borderRadius:8, fontSize:"0.68rem", color:"#DC2626" }}>
              ⚠ {errMsg}
            </div>
          )}

          {/* Materials */}
          {results?.legend?.length > 0 && (
            <div>
              <div style={{ fontSize:"0.65rem", letterSpacing:"0.1em", color:BFS_BLUE, textTransform:"uppercase", fontWeight:700, marginBottom:"0.65rem" }}>Materials</div>
              <div style={{ display:"flex", flexDirection:"column", gap:"0.35rem" }}>
                {results.legend.map((m, i) => (
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:"0.5rem" }}>
                    <div style={{ width:10, height:10, borderRadius:3, background:MATERIAL_COLORS[m.category]||"#9CA3AF", flexShrink:0 }} />
                    <span style={{ fontSize:"0.7rem", color:"#475569" }}><span style={{ fontWeight:600, color:"#1E293B" }}>{m.id}</span>: {m.name}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* BFS full logo at bottom */}
          <div style={{ marginTop:"auto", paddingTop:"1.5rem", display:"flex", flexDirection:"column", alignItems:"center", gap:"0.5rem", opacity:0.25 }}>
            <img src="/logo-full.png" alt="BFS" style={{ height:90, width:"auto" }} />
          </div>
        </aside>

        {/* ── Main content ── */}
        <main style={{ flex:1, overflow:"hidden", display:"flex", flexDirection:"column" }}>

          {/* View toggle */}
          {phase==="done" && results && (
            <div style={{ padding:"0.75rem 1.5rem", borderBottom:"1px solid #E2E8F0", display:"flex", gap:"0.5rem", background:"#fff", alignItems:"center" }}>
              <span style={{ fontSize:"0.62rem", color:"#94A3B8", letterSpacing:"0.1em", textTransform:"uppercase", fontWeight:600, marginRight:"0.25rem" }}>View</span>
              {[["table","📊  Table"],["interactive","🎨  Interactive Takeoff"]].map(([mode,label]) => (
                <button key={mode} onClick={() => setViewMode(mode)} style={pill(viewMode===mode)}>{label}</button>
              ))}
              {viewMode==="interactive" && <span style={{ fontSize:"0.65rem", color:"#94A3B8", marginLeft:"0.5rem" }}>Click any zone → assign material</span>}
            </div>
          )}

          {/* Summary strip */}
          {phase==="done" && summary && viewMode==="table" && (
            <div style={{ padding:"1rem 1.5rem", borderBottom:"1px solid #E2E8F0", display:"flex", gap:"0.75rem", flexWrap:"wrap", background:"#FAFBFC" }}>
              {Object.entries(summary).map(([cat, { net, adj, color }]) => (
                <div key={cat} style={{ ...card({ padding:"0.75rem 1rem", minWidth:155, borderLeft:`4px solid ${color}` }) }}>
                  <div style={{ fontSize:"0.62rem", color:"#94A3B8", fontWeight:500, marginBottom:"0.2rem" }}>{cat}</div>
                  <div style={{ fontSize:"1.5rem", fontWeight:700, color:"#0F172A", lineHeight:1 }}>{Math.round(adj).toLocaleString()}</div>
                  <div style={{ fontSize:"0.62rem", color:"#94A3B8", marginTop:"0.2rem" }}>SF adj · {Math.round(net).toLocaleString()} net</div>
                </div>
              ))}
              <div style={{ ...card({ padding:"0.75rem 1rem", minWidth:155, borderLeft:`4px solid ${BFS_BLUE}`, background: BFS_BLUE_PALE }) }}>
                <div style={{ fontSize:"0.62rem", color:BFS_BLUE, fontWeight:700, marginBottom:"0.2rem", textTransform:"uppercase", letterSpacing:"0.06em" }}>Grand Total</div>
                <div style={{ fontSize:"1.5rem", fontWeight:700, color:BFS_BLUE, lineHeight:1 }}>{Math.round(grandAdj).toLocaleString()}</div>
                <div style={{ fontSize:"0.62rem", color:BFS_BLUE_DARK+"99", marginTop:"0.2rem" }}>SF adjusted · all panels</div>
              </div>
            </div>
          )}

          {/* Interactive view */}
          {phase==="done" && results && viewMode==="interactive" && (
            <div style={{ flex:1, overflow:"hidden" }}>
              <InteractiveView results={results} BACKEND={BACKEND} />
            </div>
          )}

          {/* Table view */}
          {phase==="done" && results && viewMode==="table" && (
            <div style={{ flex:1, overflowY:"auto", padding:"1.25rem 1.5rem" }}>
              <div style={{ fontSize:"0.65rem", letterSpacing:"0.1em", color:BFS_BLUE, textTransform:"uppercase", fontWeight:700, marginBottom:"1rem" }}>Breakdown by Elevation</div>
              {results.takeoffData.map((elev, i) => {
                const total = (elev.zones||[]).reduce((s,z) => s+(z.netArea||0), 0);
                if (total===0) return null;
                return (
                  <div key={i} style={{ ...card({ marginBottom:"0.875rem", overflow:"hidden" }) }}>
                    <div style={{ padding:"0.65rem 1rem", display:"flex", justifyContent:"space-between", alignItems:"center", borderBottom:"1px solid #F1F5F9", background:"#FAFBFC" }}>
                      <span style={{ fontSize:"0.85rem", fontWeight:700, color:"#0F172A" }}>{elev.title}</span>
                      <span style={{ fontSize:"0.65rem", color:"#94A3B8" }}>{elev.sheetRef}{elev.scale?" · "+elev.scale:""} · {Math.round(total).toLocaleString()} SF</span>
                    </div>
                    <table style={{ width:"100%", borderCollapse:"collapse" }}>
                      <thead>
                        <tr style={{ background:"#F8FAFC" }}>
                          {["ID","Material","Category","Gross","Openings","Net SF","Adj +15%"].map(h => (
                            <th key={h} style={{ padding:"0.4rem 0.75rem", textAlign:["Gross","Openings","Net SF","Adj +15%"].includes(h)?"right":"left",
                              fontSize:"0.6rem", fontWeight:600, color:"#94A3B8", textTransform:"uppercase", letterSpacing:"0.05em",
                              borderBottom:"1px solid #F1F5F9" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(elev.zones||[]).map((z, zi) => (
                          <tr key={zi} style={{ borderBottom:"1px solid #F8FAFC" }}
                            onMouseEnter={e => e.currentTarget.style.background="#FAFBFC"}
                            onMouseLeave={e => e.currentTarget.style.background="transparent"}>
                            <td style={{ padding:"0.45rem 0.75rem" }}>
                              <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                                <div style={{ width:8, height:8, borderRadius:2, background:MATERIAL_COLORS[z.category]||"#9CA3AF" }} />
                                <span style={{ fontWeight:700, color:BFS_BLUE, fontSize:"0.72rem" }}>{z.materialId||"—"}</span>
                              </div>
                            </td>
                            <td style={{ padding:"0.45rem 0.75rem", fontSize:"0.72rem", color:"#374151", fontWeight:500 }}>{z.materialName}</td>
                            <td style={{ padding:"0.45rem 0.75rem", fontSize:"0.7rem", color:"#9CA3AF" }}>{z.category}</td>
                            <td style={{ padding:"0.45rem 0.75rem", textAlign:"right", fontSize:"0.72rem", color:"#64748B" }}>{Math.round(z.grossArea||0).toLocaleString()}</td>
                            <td style={{ padding:"0.45rem 0.75rem", textAlign:"right", fontSize:"0.72rem", color:"#9CA3AF" }}>({Math.round(z.totalOpeningArea||0).toLocaleString()})</td>
                            <td style={{ padding:"0.45rem 0.75rem", textAlign:"right", fontSize:"0.75rem", fontWeight:600, color:"#374151" }}>{Math.round(z.netArea||0).toLocaleString()}</td>
                            <td style={{ padding:"0.45rem 0.75rem", textAlign:"right", fontSize:"0.78rem", fontWeight:700, color:BFS_BLUE }}>{Math.round((z.netArea||0)*1.15).toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {(elev.flags||[]).filter(Boolean).length>0 && (
                      <div style={{ padding:"0.4rem 0.75rem", fontSize:"0.65rem", color:"#92400E", background:"#FFFBEB", borderTop:"1px solid #FEF3C7" }}>
                        ⚠ {elev.flags.filter(Boolean).join(" · ")}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* Idle state */}
          {phase==="idle" && (
            <div style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"2rem", textAlign:"center" }}>
              <div style={{ ...card({ padding:"3rem 2.5rem", maxWidth:420 }) }}>
                <div style={{ fontSize:"4rem", marginBottom:"1rem" }}>🏗</div>
                <div style={{ fontSize:"1.1rem", fontWeight:700, color:"#0F172A", marginBottom:"0.5rem" }}>Ready to Estimate</div>
                <div style={{ fontSize:"0.8rem", color:"#94A3B8", lineHeight:1.9 }}>
                  Upload a blueprint PDF and click <strong style={{ color:BFS_BLUE }}>Run Analysis</strong><br/>
                  Reads sheet index → Material legend → All elevations<br/>
                  Soffits · Returns · Per-elevation SF breakdown
                </div>
              </div>
            </div>
          )}

          {/* Activity log */}
          <div style={{ borderTop:"1px solid #E2E8F0", padding:"0.75rem 1.5rem", background:"#fff", flexShrink:0 }}>
            <div style={{ fontSize:"0.6rem", letterSpacing:"0.1em", color:"#CBD5E1", textTransform:"uppercase", fontWeight:600, marginBottom:"0.35rem" }}>
              Activity Log {isRunning ? "· polling every 5s" : ""}
            </div>
            <div ref={logRef} style={{ fontFamily:"'Courier New',monospace", fontSize:"0.7rem", maxHeight:120, overflowY:"auto", lineHeight:1.8 }}>
              {log.length===0
                ? <span style={{ color:"#E2E8F0" }}>Waiting...</span>
                : log.map((l,i) => <div key={i} style={{ color:logColor[l.level]||"#94A3B8" }}>{l.msg}</div>)}
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}
