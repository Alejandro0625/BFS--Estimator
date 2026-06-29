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

// Builds Excel matching the exact BPS proposal + estimate template
// materials = [{ name, sf }]  — one row per assigned material type
const buildExcel = (projectName, materials) => {
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString("en-US");
  const n11 = () => Array(11).fill(null);
  const n10 = () => Array(10).fill(null);

  // ── ESTIMATE SHEET ──────────────────────────────────────────────
  // Cols: A(0) B(1) C(2) D(3) E(4) F(5) G(6) H(7) I(8) J(9)
  const eRows = [];

  // Rows 1-2: project name (A1:H2 merged)
  let er = n10(); er[0] = projectName || ""; eRows.push(er);
  eRows.push(n10());

  // Row 3: blank section spacer (A3:H3 merged)
  eRows.push(n10());

  // Row 4: PANELS header (A4:G4 merged)
  er = n10(); er[0] = "PANELS"; eRows.push(er);

  // Row 5: column headers
  er = n10();
  er[0] = "No."; er[1] = "ACM/ACP"; er[2] = "Quantity"; er[4] = "Conv"; er[5] = "Rate"; er[6] = "Amount";
  eRows.push(er);

  // Row 6: blank
  eRows.push(n10());

  // Rows 7+: one section per material (header row + item row)
  const amtCells = [];
  materials.forEach((mat, idx) => {
    // material name subheader
    er = n10(); er[1] = mat.name; eRows.push(er);
    // item row
    const itemExcelRow = eRows.length + 1; // 1-indexed
    er = n10();
    er[0] = idx + 1;
    er[1] = mat.name;
    er[2] = Math.round(mat.sf);
    er[4] = 1;
    er[5] = "";          // $/SF — estimator fills this in
    er[6] = `=C${itemExcelRow}*F${itemExcelRow}`;
    amtCells.push(`G${itemExcelRow}`);
    eRows.push(er);
  });

  // Pad to row 25 area for Total
  while (eRows.length < 24) eRows.push(n10());
  const totalExcelRow = eRows.length + 1;
  er = n10();
  er[1] = "Total ";
  er[2] = materials.reduce((s, m) => s + Math.round(m.sf), 0);
  er[6] = amtCells.length ? `=${amtCells.join("+")}` : 0;
  eRows.push(er);

  // Row after total: blank
  eRows.push(n10());

  // PANEL BACK-UP SYSTEM section
  er = n10(); er[0] = "PANEL BACK-UP SYSTEM- Z-Girts, Hat Channel, Insulation"; eRows.push(er);
  eRows.push(n10());
  er = n10();
  er[1] = "Furnish and install the quantity of new metal panels required\nAny exterior caulking required\nLifts / tie-off required per site policy\nAny break metal/flashing required\nStructural calculations and PE stamp\nShop drawings\nTaxes";
  eRows.push(er);

  // Pad to row 35
  while (eRows.length < 34) eRows.push(n10());

  // SPECIFICATIONS section
  er = n10(); er[0] = "SPECIFICATIONS"; eRows.push(er);
  eRows.push(n10());
  ["GC", "Location", "Profit/Non-Profit", "Taxable/Non-Taxable", "Prevailing Wage", "Drawing Set", "Building Height"].forEach(f => {
    er = n10(); er[0] = f; eRows.push(er);
  });

  const wsE = XLSX.utils.aoa_to_sheet(eRows);
  wsE["!cols"] = [5.14, 57.43, 11.86, 4.14, 5.43, 12.86, 12.43].map(w => ({ wch: w }));
  wsE["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 7 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 6 } },
    { s: { r: totalExcelRow - 1, c: 0 }, e: { r: totalExcelRow - 1, c: 9 } },
    { s: { r: eRows.length - 8, c: 0 }, e: { r: eRows.length - 8, c: 6 } },
  ];
  XLSX.utils.book_append_sheet(wb, wsE, "Estimate");

  // ── PROPOSAL SHEET ──────────────────────────────────────────────
  // Cols A–K (0–10), 11 wide. Right side (F=5 onward) = GC/job info
  const pRows = [];

  let pr = n11(); pr[5] = "PROPOSAL"; pRows.push(pr);
  pr = n11(); pr[5] = "DATE:"; pr[6] = today; pRows.push(pr);
  pr = n11(); pr[0] = "ACM. Trespa. Terracotta  &  Specialty Metal Panels"; pr[5] = "This proposal may be withdrawn by us if not accepted within 30 days."; pRows.push(pr);
  pr = n11(); pr[0] = "15 Erie Drive"; pr[5] = "E-mail:"; pRows.push(pr);
  pr = n11(); pr[0] = "Natick, MA 01760"; pr[5] = ""; pRows.push(pr);          // GC email — leave blank
  pr = n11(); pr[0] = "PH: 617-458-2000  "; pr[5] = "Phone:"; pRows.push(pr);
  pr = n11(); pr[0] = "To:"; pr[1] = ""; pr[5] = ""; pRows.push(pr);           // GC PM name + phone
  pr = n11(); pr[1] = ""; pr[5] = "Job Name / location:"; pRows.push(pr);       // GC PM title
  pr = n11(); pr[1] = ""; pr[5] = projectName || ""; pRows.push(pr);            // GC company + job name
  pr = n11(); pr[1] = ""; pr[5] = "Job number: "; pRows.push(pr);
  pr = n11(); pr[5] = ""; pRows.push(pr);                                        // job number
  pr = n11(); pr[0] = "We hereby submit specifications and estimates for:"; pRows.push(pr);

  const mainMat = materials[0]?.name || "[Material]";
  const totalSF = materials.reduce((s, m) => s + Math.round(m.sf), 0);
  pr = n11(); pr[1] = `Install ${totalSF.toLocaleString()}sf of ${mainMat}.`; pRows.push(pr);

  const scopeItems = [
    "Include all OSHA and fall protection compliance for the installation of panels",
    "Include all staging and lifts for the performance of work.",
    `F&I ${materials.map(m => m.name).join(", ")} as specified.`,
    "F&I all metal trim and accessories with panels as specified.",
    "Remove and dispose of all job related debris to the general contractor's dumpster.",
    "MA Sales Tax Included on all materials if applicable.",
  ];
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

  // TOTAL: references Estimate sheet total row
  pr = n11(); pr[7] = "TOTAL:"; pr[8] = `=Estimate!G${totalExcelRow}`; pRows.push(pr);

  pr = n11(); pr[0] = "Payment to be made as follows:"; pr[4] = "AIA Format"; pRows.push(pr);
  pr = n11(); pr[6] = "Akshita Patel"; pRows.push(pr);
  pr = n11(); pr[7] = "Authorized Signature"; pRows.push(pr);
  pr = n11(); pr[0] = "All material to be as specified. All work to be performed in a professional manner according to standard practices. Any alteration or deviation from above specifications involving additional costs will be executed only upon written orders and will be an extra charge. Payment is due in full within thirty (30) days of the date of the invoice. Interest will be charged on outstanding balances at twelve percent (12%) per annum."; pRows.push(pr);
  pr = n11(); pr[0] = "Acceptance of Proposal — The above prices, specifications and conditions are satisfactory and hereby accepted. You are authorized to do the work as specified. Payment as outlined above."; pRows.push(pr);
  pr = n11(); pr[3] = "Date:"; pr[7] = "Signature"; pRows.push(pr);

  const wsP = XLSX.utils.aoa_to_sheet(pRows);
  wsP["!cols"] = [7.14, 9.14, 12, 12, 16.71, 9.57, 9.14, 12, 12, 12, 12].map(w => ({ wch: w }));
  wsP["!merges"] = [
    { s: { r: 0, c: 5 }, e: { r: 0, c: 10 } },
    { s: { r: 1, c: 6 }, e: { r: 1, c: 7 } },
    { s: { r: 2, c: 5 }, e: { r: 2, c: 10 } },
    { s: { r: 3, c: 5 }, e: { r: 3, c: 10 } },
    { s: { r: 4, c: 5 }, e: { r: 4, c: 10 } },
    { s: { r: 5, c: 5 }, e: { r: 5, c: 10 } },
    { s: { r: 6, c: 5 }, e: { r: 6, c: 10 } },
    { s: { r: 7, c: 5 }, e: { r: 7, c: 10 } },
    { s: { r: 8, c: 5 }, e: { r: 8, c: 10 } },
    { s: { r: 9, c: 5 }, e: { r: 9, c: 10 } },
    { s: { r: 10, c: 5 }, e: { r: 10, c: 10 } },
    { s: { r: pRows.length - 4, c: 0 }, e: { r: pRows.length - 3, c: 5 } },
  ];
  XLSX.utils.book_append_sheet(wb, wsP, "Proposal");

  return wb;
};

// Distinct colors for auto-detected texture clusters
// These are what the estimator SEES on screen — each texture gets one color
const CLUSTER_UI_COLORS = [
  "#4a9de0",  // blue
  "#e08c30",  // orange
  "#4dc47a",  // green
  "#d44a9d",  // pink
  "#c8c040",  // yellow
  "#9a4ae0",  // purple
  "#4ac8c0",  // teal
  "#e04a4a",  // red
];

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

  // Group zones by texture cluster (cluster_id for CAD, fill_color for Bluebeam)
  const colorGroups = {};
  displayZones.forEach(z => {
    const key = z.cluster_id !== undefined
      ? "c_" + z.cluster_id
      : z.fill_color ? "f_" + z.fill_color.join(",") : "none";
    if (!colorGroups[key]) colorGroups[key] = [];
    colorGroups[key].push(z.id);
  });

  // Cluster summary: total SF per cluster (unassigned)
  const clusterSummary = {};
  displayZones.forEach(z => {
    const key = z.cluster_id !== undefined ? z.cluster_id : -1;
    if (!clusterSummary[key]) clusterSummary[key] = { total_sf: 0, count: 0, color: "#8888aa" };
    clusterSummary[key].total_sf += z.area_sf || 0;
    clusterSummary[key].count += 1;
    clusterSummary[key].color = z.cluster_id !== undefined
      ? CLUSTER_UI_COLORS[z.cluster_id % CLUSTER_UI_COLORS.length]
      : "#8888aa";
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

  const exportInteractiveExcel = () => {
    // Aggregate SF by material name across all assignments
    const matTotals = {};
    Object.values(assignments).forEach(a => {
      const key = a.materialName || a.category || "Panel";
      if (!matTotals[key]) matTotals[key] = { name: key, sf: 0 };
      matTotals[key].sf += a.area_sf || 0;
    });
    const materials = Object.values(matTotals);
    const projName = results.projName || "Project";
    const wb = buildExcel(projName, materials);
    XLSX.writeFile(wb, "BPS_Takeoff_" + projName.replace(/\s+/g, "_") + ".xlsx");
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
            ? "📐 Bluebeam polygons — " + displayZones.length + " surfaces · SF exact from markup"
            : polyMethod === "vector_cluster" || polyMethod === "vector"
            ? "📏 Vector mode — " + displayZones.length + " surfaces from CAD geometry"
            : polyMethod === "claude_vision"
            ? "🧠 AI Vision — " + displayZones.length + " surfaces detected from drawing patterns · click to assign"
            : "No surfaces detected on this page"}
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
                  // Color priority:
                  //   1. Assigned material color (confirmed by estimator)
                  //   2. Cluster UI color (auto-detected texture group)
                  //   3. Bluebeam fill color (if annotations present)
                  //   4. Default gray
                  let color = "#8888aa";
                  if (a) {
                    color = MATERIAL_COLORS[a.category] || "#7a7a7a";
                  } else if (zone.cluster_id !== undefined) {
                    color = CLUSTER_UI_COLORS[zone.cluster_id % CLUSTER_UI_COLORS.length];
                  } else if (zone.fill_color && zone.fill_color.length === 3) {
                    const [r2, g2, b2] = zone.fill_color;
                    color = "rgb(" + Math.round(r2*255) + "," + Math.round(g2*255) + "," + Math.round(b2*255) + ")";
                  }
                  const isActive = activeZone === zone.id;
                  const pts = toSVGPoints(zone.points);
                  const labelX = zone.cx * pageDims.width;
                  const labelY = zone.cy * pageDims.height;
                  const showLabel = a || isActive || zone.source === "bluebeam" || zone.source === "claude_vision";
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
                            {zone.source === "claude_vision" && !a
                              ? zone.material_type
                              : Math.round(zone.area_sf) + " SF"}
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
              const colorKey = activeZoneData?.cluster_id !== undefined
                ? "c_" + activeZoneData.cluster_id
                : activeZoneData?.fill_color ? "f_" + activeZoneData.fill_color.join(",") : null;
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
          /* Right panel — nothing selected: show detected textures + assigned SF */
          <>
            {/* Detected texture clusters */}
            {Object.keys(clusterSummary).length > 0 && (
              <>
                <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.4rem" }}>
                  Detected Textures
                </div>
                <div style={{ fontSize: "0.6rem", color: "#4a3a18", marginBottom: "0.5rem" }}>
                  Click a colored zone → assign material
                </div>
                {Object.entries(clusterSummary).map(([cid, info]) => {
                  const groupKey = "c_" + cid;
                  const zoneIds = colorGroups[groupKey] || [];
                  const assigned = zoneIds.filter(id => getAssignment(id)).length;
                  return (
                    <div key={cid} style={{ marginBottom: "0.4rem", padding: "0.4rem 0.5rem", background: "#0e0d0a", borderRadius: 3, borderLeft: "3px solid " + info.color }}>
                      <div style={{ display: "flex", justifyContent: "space-between" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
                          <div style={{ width: 10, height: 10, borderRadius: 2, background: info.color }} />
                          <span style={{ fontSize: "0.62rem", color: "#8a7a50" }}>Texture {parseInt(cid)+1}</span>
                        </div>
                        <span style={{ fontSize: "0.75rem", fontWeight: 700, color: "#c0a840" }}>
                          {Math.round(info.total_sf).toLocaleString()} SF
                        </span>
                      </div>
                      <div style={{ fontSize: "0.58rem", color: assigned > 0 ? "#5a8030" : "#4a3a18", marginTop: 2 }}>
                        {assigned > 0 ? "✓ " + assigned + "/" + info.count + " assigned" : info.count + " zone" + (info.count !== 1 ? "s" : "") + " · click to assign"}
                      </div>
                    </div>
                  );
                })}
                <div style={{ height: "1px", background: "#1e1c14", margin: "0.6rem 0" }} />
              </>
            )}

            {/* Confirmed SF totals */}
            {Object.keys(totals).length > 0 && (
              <>
                <div style={{ fontSize: "0.58rem", letterSpacing: "0.3em", color: "#6a5a30", textTransform: "uppercase", marginBottom: "0.5rem" }}>
                  Your Takeoff
                </div>
                {Object.entries(totals).map(([cat, sf]) => (
                  <div key={cat} style={{ marginBottom: "0.5rem" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
                      <div style={{ width: 8, height: 8, borderRadius: 2, background: MATERIAL_COLORS[cat] || "#7a7a7a" }} />
                      <span style={{ fontSize: "0.6rem", color: "#7a6a40" }}>{cat}</span>
                    </div>
                    <div style={{ fontSize: "1rem", fontWeight: 700, color: "#e0cc80", paddingLeft: "1rem" }}>
                      {Math.round(sf).toLocaleString()} SF net
                    </div>
                    <div style={{ fontSize: "0.6rem", color: "#5a4a20", paddingLeft: "1rem" }}>
                      {Math.round(sf * 1.15).toLocaleString()} SF +15%
                    </div>
                  </div>
                ))}
                <div style={{ marginTop: "0.6rem", paddingTop: "0.5rem", borderTop: "1px solid #1e1c14" }}>
                  <div style={{ fontSize: "0.6rem", color: "#5a8030" }}>Grand Total</div>
                  <div style={{ fontSize: "1.2rem", fontWeight: 700, color: "#90e060" }}>
                    {Math.round(grandTotal * 1.15).toLocaleString()} SF
                  </div>
                  <div style={{ fontSize: "0.58rem", color: "#3a5020" }}>adjusted +15% waste</div>
                </div>

                <button onClick={exportInteractiveExcel}
                  style={{ marginTop: "0.75rem", width: "100%", padding: "0.6rem", background: "transparent",
                    color: "#5aaa40", border: "1px solid #3a7a20", borderRadius: 3, fontSize: "0.65rem",
                    fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", cursor: "pointer" }}>
                  ↓ Export Excel
                </button>
              </>
            )}

            {Object.keys(clusterSummary).length === 0 && Object.keys(totals).length === 0 && (
              <div style={{ fontSize: "0.65rem", color: "#3a3020", lineHeight: 1.7 }}>
                Click any colored zone on the elevation → assign material
              </div>
            )}

            <div style={{ marginTop: "0.75rem", fontSize: "0.58rem", color: "#3a3018" }}>
              {Object.keys(assignments).length} zones confirmed · {elevations.length} elevations
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
    // Build material list from Claude AI takeoff data
    const matTotals = {};
    results.takeoffData.forEach(e => (e.zones || []).forEach(z => {
      const key = z.materialName || z.category || "Panel";
      if (!matTotals[key]) matTotals[key] = { name: key, sf: 0 };
      matTotals[key].sf += z.netArea || 0;
    }));
    const wb = buildExcel(results.projName || "Project", Object.values(matTotals));
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

