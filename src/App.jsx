import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const WASTE = 0.15;
const BACKEND = import.meta.env.VITE_BACKEND_URL || "";

// BPS brand colors
const C = {
  bpsBlue:    "#1D4E89",
  bpsBlueLt:  "#2563AB",
  bpsBluePale:"#EBF2FA",
  bpsBlueMid: "#3B7DD8",
  white:      "#FFFFFF",
  gray50:     "#F9FAFB",
  gray100:    "#F3F4F6",
  gray200:    "#E5E7EB",
  gray300:    "#D1D5DB",
  gray400:    "#9CA3AF",
  gray500:    "#6B7280",
  gray600:    "#4B5563",
  gray700:    "#374151",
  gray800:    "#1F2937",
  gray900:    "#111827",
  green:      "#16A34A",
  greenLt:    "#DCFCE7",
  red:        "#DC2626",
  redLt:      "#FEE2E2",
  // Interactive takeoff (dark drawing canvas)
  navy:       "#0D1B2E",
  navyMid:    "#112240",
  navyLight:  "#1E3A5F",
  navyBorder: "#243B55",
  slateText:  "#94A3B8",
  slateLight: "#CBD5E1",
  slateActive:"#E2E8F0",
};

const MATERIAL_COLORS = {
  "ACM Panel":              "#2563AB",
  "MCM Panel":              "#2563AB",
  "Fiber Cement Panel":     "#16A34A",
  "Fiber Cement Plank":     "#0D9488",
  "Nichiha Panel":          "#7C3AED",
  "Aluminum Wall Panel":    "#0891B2",
  "Perforated Metal Panel": "#EA580C",
  "Soffit Panel":           "#2563AB",
  "Return/Trim":            "#DB2777",
  "Other":                  "#6B7280",
};

const buildExcel = (projectName, materials) => {
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString("en-US");
  const n11 = () => Array(11).fill(null);
  const n10 = () => Array(10).fill(null);

  const eRows = [];
  let er = n10(); er[0] = projectName || ""; eRows.push(er);
  eRows.push(n10());
  eRows.push(n10());
  er = n10(); er[0] = "PANELS"; eRows.push(er);
  er = n10();
  er[0] = "No."; er[1] = "ACM/ACP"; er[2] = "Quantity"; er[4] = "Conv"; er[5] = "Rate"; er[6] = "Amount";
  eRows.push(er);
  eRows.push(n10());

  const amtCells = [];
  materials.forEach((mat, idx) => {
    er = n10(); er[1] = mat.name; eRows.push(er);
    const itemExcelRow = eRows.length + 1;
    er = n10();
    er[0] = idx + 1;
    er[1] = mat.name;
    er[2] = Math.round(mat.sf);
    er[4] = 1;
    er[5] = "";
    er[6] = `=C${itemExcelRow}*F${itemExcelRow}`;
    amtCells.push(`G${itemExcelRow}`);
    eRows.push(er);
  });

  while (eRows.length < 24) eRows.push(n10());
  const totalExcelRow = eRows.length + 1;
  er = n10();
  er[1] = "Total ";
  er[2] = materials.reduce((s, m) => s + Math.round(m.sf), 0);
  er[6] = amtCells.length ? `=${amtCells.join("+")}` : 0;
  eRows.push(er);
  eRows.push(n10());

  er = n10(); er[0] = "PANEL BACK-UP SYSTEM- Z-Girts, Hat Channel, Insulation"; eRows.push(er);
  eRows.push(n10());
  er = n10();
  er[1] = "Furnish and install the quantity of new metal panels required\nAny exterior caulking required\nLifts / tie-off required per site policy\nAny break metal/flashing required\nStructural calculations and PE stamp\nShop drawings\nTaxes";
  eRows.push(er);

  while (eRows.length < 34) eRows.push(n10());
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

const CLUSTER_UI_COLORS = [
  "#3B82F6", "#F97316", "#22C55E", "#EC4899",
  "#EAB308", "#8B5CF6", "#14B8A6", "#EF4444",
];


// ── Interactive Takeoff View ──────────────────────────────────────────────────
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

  const colorGroups = {};
  displayZones.forEach(z => {
    const key = z.cluster_id !== undefined
      ? "c_" + z.cluster_id
      : z.fill_color ? "f_" + z.fill_color.join(",") : "none";
    if (!colorGroups[key]) colorGroups[key] = [];
    colorGroups[key].push(z.id);
  });

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

  const totals = {};
  Object.values(assignments).forEach(a => {
    if (!totals[a.category]) totals[a.category] = 0;
    totals[a.category] += a.area_sf || 0;
  });
  const grandTotal = Object.values(totals).reduce((s, v) => s + v, 0);

  const svgW = imgRef.current?.offsetWidth || imgNaturalSize.w;
  const svgH = imgRef.current?.offsetHeight || imgNaturalSize.h;

  const toSVGPoints = pts =>
    pts.map(([nx, ny]) => `${(nx * pageDims.width).toFixed(1)},${(ny * pageDims.height).toFixed(1)}`).join(" ");

  const matList = results.legend.length > 0
    ? results.legend
    : Object.keys(MATERIAL_COLORS).map(cat => ({ id: cat.substring(0, 3).toUpperCase() + "-1", name: cat, category: cat }));

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden", background: C.navy, fontFamily: "'Inter', 'Segoe UI', sans-serif" }}>

      {/* ── Left: elevation list ── */}
      <div style={{ width: 185, borderRight: "1px solid " + C.navyBorder, overflowY: "auto", flexShrink: 0, background: C.navyMid }}>
        <div style={{ padding: "0.75rem 0.875rem", fontSize: "0.6rem", letterSpacing: "0.15em", color: C.slateText, textTransform: "uppercase", fontWeight: 600, borderBottom: "1px solid " + C.navyBorder }}>
          Elevations
        </div>
        {elevations.map((e, i) => {
          const assigned = Object.keys(assignments).filter(k => k.startsWith(i + ":")).length;
          return (
            <div key={i} onClick={() => setElevIdx(i)}
              style={{ padding: "0.6rem 0.875rem", cursor: "pointer", borderBottom: "1px solid " + C.navyBorder,
                background: i === elevIdx ? C.navyLight : "transparent",
                borderLeft: i === elevIdx ? "3px solid " + C.bpsBlueMid : "3px solid transparent" }}>
              <div style={{ fontSize: "0.7rem", color: i === elevIdx ? C.slateActive : C.slateText, lineHeight: 1.3, fontWeight: i === elevIdx ? 600 : 400 }}>
                {e.title || "Page " + e.pageNumber}
              </div>
              <div style={{ fontSize: "0.6rem", color: assigned > 0 ? "#4ADE80" : C.slateText + "80", marginTop: 2 }}>
                {assigned > 0 ? "✓ " + assigned + " assigned" : (e.zones || []).length + " zones · p." + e.pageNumber}
              </div>
            </div>
          );
        })}
        {elevations.length === 0 && (
          <div style={{ padding: "0.875rem", fontSize: "0.65rem", color: C.slateText + "60" }}>
            No elevations with page numbers. Re-run analysis.
          </div>
        )}
      </div>

      {/* ── Center: image + SVG overlay ── */}
      <div style={{ flex: 1, overflow: "auto", display: "flex", flexDirection: "column", alignItems: "center", padding: "1rem", gap: "0.5rem", background: C.navy }}>
        <div style={{ fontSize: "0.65rem", color: C.slateText, alignSelf: "flex-start", background: C.navyLight, padding: "0.3rem 0.6rem", borderRadius: 4, border: "1px solid " + C.navyBorder }}>
          {polyMethod === "bluebeam"
            ? "📐 Bluebeam polygons — " + displayZones.length + " surfaces · SF exact from markup"
            : polyMethod === "vector_cluster" || polyMethod === "vector"
            ? "📏 Vector mode — " + displayZones.length + " surfaces from CAD geometry"
            : polyMethod === "claude_vision"
            ? "🧠 AI Vision — " + displayZones.length + " surfaces detected from drawing patterns · click to assign"
            : "No surfaces detected on this page"}
        </div>

        {!pageImage ? (
          <div style={{ color: C.slateText, fontSize: "0.8rem", marginTop: "3rem" }}>Loading elevation image...</div>
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
              style={{ display: "block", maxWidth: "100%", maxHeight: "calc(100vh - 180px)", objectFit: "contain", border: "1px solid " + C.navyBorder, borderRadius: 4 }}
            />
            {imgLoaded && (
              <svg
                viewBox={`0 0 ${pageDims.width} ${pageDims.height}`}
                style={{ position: "absolute", top: 0, left: 0, width: svgW, height: svgH, overflow: "visible" }}
              >
                {displayZones.map(zone => {
                  const a = getAssignment(zone.id);
                  let color = "#8888aa";
                  if (a) {
                    color = MATERIAL_COLORS[a.category] || "#6B7280";
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
                          <rect x={labelX - 30} y={labelY - 9} width={60} height={17} fill="rgba(0,0,0,0.75)" rx={3} />
                          <text
                            x={labelX} y={labelY + 2}
                            textAnchor="middle" dominantBaseline="middle"
                            fill="white" fontSize={pageDims.width / 75}
                            fontFamily="Inter, Arial" fontWeight="bold"
                          >
                            {zone.source === "claude_vision" && !a ? zone.material_type : Math.round(zone.area_sf) + " SF"}
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
      <div style={{ width: 220, borderLeft: "1px solid " + C.navyBorder, padding: "0.875rem", overflowY: "auto", flexShrink: 0, background: C.navyMid }}>

        {activeZone !== null ? (
          <>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.15em", color: C.slateText, textTransform: "uppercase", marginBottom: "0.5rem", fontWeight: 600 }}>
              Assign Surface
            </div>
            <div style={{ padding: "0.4rem 0.6rem", marginBottom: "0.75rem", background: C.navy, borderRadius: 4, fontSize: "0.68rem", color: C.slateLight, border: "1px solid " + C.navyBorder }}>
              {Math.round(displayZones.find(z => z.id === activeZone)?.area_sf || 0)} SF selected
            </div>

            {matList.map((mat, i) => {
              const activeZoneData = displayZones.find(z => z.id === activeZone);
              const colorKey = activeZoneData?.cluster_id !== undefined
                ? "c_" + activeZoneData.cluster_id
                : activeZoneData?.fill_color ? "f_" + activeZoneData.fill_color.join(",") : null;
              const sameColorZones = colorKey ? (colorGroups[colorKey] || []) : [];
              const canBulkAssign = sameColorZones.length > 1;

              return (
                <div key={i} style={{ marginBottom: "0.4rem" }}>
                  <div onClick={() => assignZone(activeZone, mat)}
                    style={{ display: "flex", alignItems: "center", gap: "0.5rem", padding: "0.45rem 0.6rem",
                      cursor: "pointer", background: C.navy, borderRadius: canBulkAssign ? "4px 4px 0 0" : 4,
                      border: "1px solid " + C.navyBorder, borderBottom: canBulkAssign ? "none" : "1px solid " + C.navyBorder,
                      transition: "border-color 0.15s" }}
                    onMouseEnter={e => e.currentTarget.style.borderColor = C.bpsBlueMid}
                    onMouseLeave={e => e.currentTarget.style.borderColor = C.navyBorder}>
                    <div style={{ width: 10, height: 10, borderRadius: 2, background: MATERIAL_COLORS[mat.category] || "#6B7280", flexShrink: 0 }} />
                    <div style={{ fontSize: "0.63rem", color: C.slateLight, lineHeight: 1.3, flex: 1 }}>
                      <span style={{ color: C.bpsBlueMid, fontWeight: 600 }}>{mat.id}</span>
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
                      style={{ padding: "0.3rem 0.6rem", background: C.navyLight, cursor: "pointer",
                        fontSize: "0.6rem", color: "#4ADE80", border: "1px solid " + C.navyBorder,
                        borderTop: "none", borderRadius: "0 0 4px 4px" }}
                      onMouseEnter={e => e.currentTarget.style.color = "#86EFAC"}
                      onMouseLeave={e => e.currentTarget.style.color = "#4ADE80"}>
                      ↳ Assign all {sameColorZones.length} same-color zones
                    </div>
                  )}
                </div>
              );
            })}

            {getAssignment(activeZone) && (
              <div onClick={() => removeAssignment(activeZone)}
                style={{ padding: "0.4rem 0.6rem", marginTop: "0.5rem", textAlign: "center", fontSize: "0.63rem", color: "#F87171", cursor: "pointer", border: "1px solid #7F1D1D", borderRadius: 4 }}>
                Remove assignment
              </div>
            )}
            <div onClick={() => setActiveZone(null)}
              style={{ padding: "0.4rem 0.6rem", marginTop: "0.3rem", textAlign: "center", fontSize: "0.63rem", color: C.slateText, cursor: "pointer", border: "1px solid " + C.navyBorder, borderRadius: 4 }}>
              Cancel
            </div>
          </>
        ) : (
          <>
            {Object.keys(clusterSummary).length > 0 && (
              <>
                <div style={{ fontSize: "0.6rem", letterSpacing: "0.15em", color: C.slateText, textTransform: "uppercase", marginBottom: "0.4rem", fontWeight: 600 }}>
                  Detected Textures
                </div>
                <div style={{ fontSize: "0.62rem", color: C.slateText + "80", marginBottom: "0.6rem" }}>
                  Click a colored zone → assign material
                </div>
                {Object.entries(clusterSummary).map(([cid, info]) => {
                  const groupKey = "c_" + cid;
                  const zoneIds = colorGroups[groupKey] || [];
                  const assigned = zoneIds.filter(id => getAssignment(id)).length;
                  return (
                    <div key={cid} style={{ marginBottom: "0.4rem", padding: "0.5rem 0.6rem", background: C.navy, borderRadius: 4, borderLeft: "3px solid " + info.color, border: "1px solid " + C.navyBorder, borderLeftWidth: 3 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
                          <div style={{ width: 8, height: 8, borderRadius: 2, background: info.color }} />
                          <span style={{ fontSize: "0.63rem", color: C.slateText }}>Texture {parseInt(cid)+1}</span>
                        </div>
                        <span style={{ fontSize: "0.78rem", fontWeight: 700, color: C.slateActive }}>
                          {Math.round(info.total_sf).toLocaleString()} SF
                        </span>
                      </div>
                      <div style={{ fontSize: "0.6rem", color: assigned > 0 ? "#4ADE80" : C.slateText + "60", marginTop: 3 }}>
                        {assigned > 0 ? "✓ " + assigned + "/" + info.count + " assigned" : info.count + " zone" + (info.count !== 1 ? "s" : "") + " · click to assign"}
                      </div>
                    </div>
                  );
                })}
                <div style={{ height: "1px", background: C.navyBorder, margin: "0.75rem 0" }} />
              </>
            )}

            {Object.keys(totals).length > 0 && (
              <>
                <div style={{ fontSize: "0.6rem", letterSpacing: "0.15em", color: C.slateText, textTransform: "uppercase", marginBottom: "0.6rem", fontWeight: 600 }}>
                  Your Takeoff
                </div>
                {Object.entries(totals).map(([cat, sf]) => (
                  <div key={cat} style={{ marginBottom: "0.6rem", padding: "0.5rem 0.6rem", background: C.navy, borderRadius: 4, border: "1px solid " + C.navyBorder }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "0.4rem", marginBottom: "0.2rem" }}>
                      <div style={{ width: 8, height: 8, borderRadius: 2, background: MATERIAL_COLORS[cat] || "#6B7280" }} />
                      <span style={{ fontSize: "0.63rem", color: C.slateText }}>{cat}</span>
                    </div>
                    <div style={{ fontSize: "1.1rem", fontWeight: 700, color: C.slateActive }}>
                      {Math.round(sf).toLocaleString()} SF net
                    </div>
                    <div style={{ fontSize: "0.6rem", color: C.slateText + "80" }}>
                      {Math.round(sf * 1.15).toLocaleString()} SF +15%
                    </div>
                  </div>
                ))}
                <div style={{ marginTop: "0.5rem", padding: "0.6rem 0.6rem", borderRadius: 4, background: C.navyLight, border: "1px solid " + C.bpsBlueMid + "60" }}>
                  <div style={{ fontSize: "0.6rem", color: "#4ADE80", fontWeight: 600, marginBottom: "0.2rem" }}>Grand Total</div>
                  <div style={{ fontSize: "1.3rem", fontWeight: 700, color: "#4ADE80" }}>
                    {Math.round(grandTotal * 1.15).toLocaleString()} SF
                  </div>
                  <div style={{ fontSize: "0.6rem", color: C.slateText + "80" }}>adjusted +15% waste</div>
                </div>

                <button onClick={exportInteractiveExcel}
                  style={{ marginTop: "0.75rem", width: "100%", padding: "0.6rem", background: C.bpsBlue,
                    color: C.white, border: "none", borderRadius: 4, fontSize: "0.65rem",
                    fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.05em", cursor: "pointer" }}>
                  ↓ Export Excel
                </button>
              </>
            )}

            {Object.keys(clusterSummary).length === 0 && Object.keys(totals).length === 0 && (
              <div style={{ fontSize: "0.65rem", color: C.slateText + "80", lineHeight: 1.7 }}>
                Click any colored zone on the elevation → assign material
              </div>
            )}

            <div style={{ marginTop: "0.875rem", fontSize: "0.6rem", color: C.slateText + "60" }}>
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
  const [viewMode, setViewMode] = useState("table");
  const fileRef = useRef();
  const logRef = useRef();
  const pollRef = useRef(null);
  const seenLogs = useRef(0);

  useEffect(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, [log]);
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
        console.log("Poll hiccup:", err.message);
      }
    }, 5000);
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
      if (!t[k]) t[k] = { net: 0, adj: 0, color: MATERIAL_COLORS[k] || "#6B7280" };
      t[k].net += z.netArea || 0;
      t[k].adj += (z.netArea || 0) * 1.15;
    }));
    return t;
  })() : null;

  const grandAdj = summary ? Object.values(summary).reduce((s, v) => s + v.adj, 0) : 0;
  const phaseStep = { idle: 0, running: 1, filtering: 1, legend: 2, analyzing: 3, done: 4, error: 0 }[phase] || 0;
  const logColor = { ok: C.green, warn: "#D97706", error: C.red, success: C.green, dim: C.gray400, info: C.gray500 };

  const isRunning = phase !== "idle" && phase !== "done" && phase !== "error";

  return (
    <div style={{ fontFamily: "'Inter', 'Segoe UI', -apple-system, sans-serif", background: C.gray100, minHeight: "100vh", color: C.gray800, display: "flex", flexDirection: "column" }}>

      {/* ── Header ── */}
      <div style={{ background: C.bpsBlue, padding: "0 1.5rem", height: 60, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0, boxShadow: "0 2px 8px rgba(0,0,0,0.25)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "1rem" }}>
          <img src="/logo.png" alt="BFS Logo" style={{ height: 44, width: "auto" }} />
          <div>
            <div style={{ fontSize: "0.55rem", letterSpacing: "0.2em", color: "rgba(255,255,255,0.7)", textTransform: "uppercase", fontWeight: 500 }}>Boston Facade Systems</div>
            <div style={{ fontSize: "1rem", fontWeight: 700, color: C.white, letterSpacing: "-0.01em" }}>AI Panel Estimator</div>
          </div>
        </div>
        <div style={{ textAlign: "right", fontSize: "0.65rem", color: "rgba(255,255,255,0.55)", lineHeight: 1.8 }}>
          <div>Panels · Soffits · Returns</div>
          <div>Waste: 15% · Scale-based measurement</div>
        </div>
      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>

        {/* ── Left Sidebar ── */}
        <div style={{ width: 280, borderRight: "1px solid " + C.gray200, padding: "1.25rem", display: "flex", flexDirection: "column", gap: "1.25rem", overflowY: "auto", background: C.white, flexShrink: 0 }}>

          {/* Upload */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: C.bpsBlue, textTransform: "uppercase", marginBottom: "0.6rem", fontWeight: 700 }}>Blueprint Set</div>
            <div onClick={() => fileRef.current?.click()} onDrop={e => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }} onDragOver={e => e.preventDefault()}
              style={{ border: "2px dashed " + (file ? C.green : C.gray300), borderRadius: 6, padding: "1.25rem 0.75rem", textAlign: "center", cursor: "pointer", background: file ? "#F0FDF4" : C.gray50, transition: "all 0.2s" }}>
              <div style={{ fontSize: "1.8rem" }}>{file ? "📋" : "📁"}</div>
              {file ? (
                <>
                  <div style={{ fontSize: "0.72rem", color: C.green, marginTop: "0.4rem", wordBreak: "break-all", fontWeight: 600 }}>{file.name}</div>
                  <div style={{ fontSize: "0.62rem", color: C.gray400, marginTop: 2 }}>{(file.size / 1e6).toFixed(1)} MB</div>
                </>
              ) : (
                <div style={{ fontSize: "0.72rem", color: C.gray400, marginTop: "0.4rem", lineHeight: 1.5 }}>Drop full blueprint PDF<br />or click to browse</div>
              )}
            </div>
            <input ref={fileRef} type="file" accept=".pdf" onChange={e => handleFile(e.target.files[0])} style={{ display: "none" }} />
          </div>

          {/* Steps */}
          <div>
            <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: C.bpsBlue, textTransform: "uppercase", marginBottom: "0.75rem", fontWeight: 700 }}>Process</div>
            {[
              { n: 1, label: "Read Sheet Index" },
              { n: 2, label: "Read Material Legend" },
              { n: 3, label: "Analyze Elevations + Soffits + Returns" },
              { n: 4, label: "Export Excel + PDF" },
            ].map(({ n, label }) => {
              const done = phaseStep > n;
              const active = phaseStep === n;
              return (
                <div key={n} style={{ display: "flex", alignItems: "center", gap: "0.65rem", marginBottom: "0.55rem" }}>
                  <div style={{
                    width: 22, height: 22, borderRadius: "50%",
                    background: done ? C.green : active ? C.bpsBlue : C.gray100,
                    border: "2px solid " + (done ? C.green : active ? C.bpsBlue : C.gray300),
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: "0.6rem", color: done || active ? C.white : C.gray400,
                    fontWeight: 700, flexShrink: 0, transition: "all 0.2s"
                  }}>
                    {done ? "✓" : n}
                  </div>
                  <span style={{ fontSize: "0.7rem", color: done ? C.green : active ? C.bpsBlue : C.gray400, fontWeight: done || active ? 600 : 400 }}>
                    {label}
                  </span>
                </div>
              );
            })}
          </div>

          {/* Progress bar */}
          {progress.pct > 0 && progress.pct < 100 && (
            <div>
              <div style={{ fontSize: "0.62rem", color: C.gray500, marginBottom: "0.4rem" }}>{progress.label}</div>
              <div style={{ background: C.gray200, borderRadius: 6, height: 6, overflow: "hidden" }}>
                <div style={{ width: progress.pct + "%", height: "100%", background: C.bpsBlue, borderRadius: 6, transition: "width 0.4s" }} />
              </div>
            </div>
          )}

          {/* Buttons */}
          <button onClick={run} disabled={!file || isRunning}
            style={{
              padding: "0.75rem", borderRadius: 6, fontSize: "0.75rem", fontFamily: "inherit",
              fontWeight: 700, letterSpacing: "0.05em", cursor: !file || isRunning ? "not-allowed" : "pointer",
              border: "none", transition: "all 0.2s",
              background: !file || isRunning ? C.gray200 : C.bpsBlue,
              color: !file || isRunning ? C.gray400 : C.white,
              boxShadow: !file || isRunning ? "none" : "0 2px 6px rgba(29,78,137,0.3)"
            }}>
            {phase === "idle" || phase === "error" ? "▶  Run Analysis" : phase === "done" ? "↺  Run Again" : "⏳  Analyzing..."}
          </button>

          {phase === "done" && (
            <>
              <button onClick={exportExcel}
                style={{ padding: "0.75rem", background: C.white, color: C.bpsBlue, border: "2px solid " + C.bpsBlue, borderRadius: 6, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, cursor: "pointer" }}>
                ↓  Export Excel
              </button>
              <button onClick={exportPDF} disabled={pdfLoading}
                style={{ padding: "0.75rem", background: C.white, color: pdfLoading ? C.gray400 : C.gray600, border: "2px solid " + (pdfLoading ? C.gray200 : C.gray300), borderRadius: 6, fontSize: "0.72rem", fontFamily: "inherit", fontWeight: 700, cursor: pdfLoading ? "not-allowed" : "pointer" }}>
                {pdfLoading ? "⏳  Generating..." : "↓  Export Evidence PDF"}
              </button>
            </>
          )}

          {phase === "error" && errMsg && (
            <div style={{ padding: "0.75rem", background: C.redLt, border: "1px solid #FECACA", borderRadius: 6, fontSize: "0.65rem", color: C.red }}>
              ⚠ {errMsg}
            </div>
          )}

          {/* Materials legend */}
          {results && results.legend.length > 0 && (
            <div>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: C.bpsBlue, textTransform: "uppercase", marginBottom: "0.6rem", fontWeight: 700 }}>Materials</div>
              {results.legend.map((m, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.35rem" }}>
                  <div style={{ width: 10, height: 10, borderRadius: 2, background: MATERIAL_COLORS[m.category] || "#6B7280", flexShrink: 0 }} />
                  <span style={{ fontSize: "0.67rem", color: C.gray600 }}>{m.id}: {m.name}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* ── Right Panel ── */}
        <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>

          {/* View toggle */}
          {phase === "done" && results && (
            <div style={{ padding: "0.75rem 1.25rem", borderBottom: "1px solid " + C.gray200, display: "flex", gap: "0.5rem", background: C.white, alignItems: "center" }}>
              <span style={{ fontSize: "0.6rem", color: C.gray400, letterSpacing: "0.1em", textTransform: "uppercase", marginRight: "0.25rem", fontWeight: 600 }}>View:</span>
              {[["table", "📊 Table"], ["interactive", "🎨 Interactive Takeoff"]].map(([mode, label]) => (
                <button key={mode} onClick={() => setViewMode(mode)}
                  style={{ padding: "0.35rem 0.875rem",
                    background: viewMode === mode ? C.bpsBlue : C.white,
                    color: viewMode === mode ? C.white : C.gray500,
                    border: "1.5px solid " + (viewMode === mode ? C.bpsBlue : C.gray300),
                    borderRadius: 5, fontSize: "0.67rem", fontFamily: "inherit", fontWeight: viewMode === mode ? 700 : 400, cursor: "pointer", transition: "all 0.15s" }}>
                  {label}
                </button>
              ))}
              {viewMode === "interactive" && (
                <span style={{ fontSize: "0.62rem", color: C.gray400, marginLeft: "0.5rem" }}>
                  Click any zone on the elevation → assign material
                </span>
              )}
            </div>
          )}

          {/* Summary cards */}
          {phase === "done" && summary && viewMode === "table" && (
            <div style={{ padding: "1rem 1.25rem", borderBottom: "1px solid " + C.gray200, display: "flex", gap: "0.75rem", flexWrap: "wrap", background: C.gray50 }}>
              {Object.entries(summary).map(([cat, { net, adj, color }]) => (
                <div key={cat} style={{ background: C.white, border: "1px solid " + C.gray200, borderLeft: "4px solid " + color, borderRadius: 6, padding: "0.75rem 1rem", minWidth: 155, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
                  <div style={{ fontSize: "0.62rem", color: C.gray400, marginBottom: "0.25rem", fontWeight: 500 }}>{cat}</div>
                  <div style={{ fontSize: "1.4rem", fontWeight: 700, color: C.gray900, lineHeight: 1 }}>{Math.round(adj).toLocaleString()}</div>
                  <div style={{ fontSize: "0.6rem", color: C.gray400, marginTop: "0.2rem" }}>SF adj · {Math.round(net).toLocaleString()} SF net</div>
                </div>
              ))}
              <div style={{ background: C.bpsBluePale, border: "1px solid " + C.bpsBlueMid + "40", borderLeft: "4px solid " + C.bpsBlue, borderRadius: 6, padding: "0.75rem 1rem", minWidth: 155, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
                <div style={{ fontSize: "0.62rem", color: C.bpsBlue, marginBottom: "0.25rem", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>Grand Total</div>
                <div style={{ fontSize: "1.4rem", fontWeight: 700, color: C.bpsBlue, lineHeight: 1 }}>{Math.round(grandAdj).toLocaleString()}</div>
                <div style={{ fontSize: "0.6rem", color: C.bpsBlue + "99", marginTop: "0.2rem" }}>SF adjusted all panels</div>
              </div>
            </div>
          )}

          {/* Interactive Takeoff */}
          {phase === "done" && results && viewMode === "interactive" && (
            <div style={{ flex: 1, overflow: "hidden" }}>
              <InteractiveView results={results} BACKEND={BACKEND} />
            </div>
          )}

          {/* Elevation table */}
          {phase === "done" && results && viewMode === "table" && (
            <div style={{ flex: 1, overflowY: "auto", padding: "1.25rem" }}>
              <div style={{ fontSize: "0.6rem", letterSpacing: "0.12em", color: C.bpsBlue, textTransform: "uppercase", marginBottom: "0.875rem", fontWeight: 700 }}>Breakdown by Elevation</div>
              {results.takeoffData.map((elev, i) => {
                const elevTotal = (elev.zones || []).reduce((s, z) => s + (z.netArea || 0), 0);
                if (elevTotal === 0) return null;
                return (
                  <div key={i} style={{ background: C.white, border: "1px solid " + C.gray200, borderRadius: 6, marginBottom: "0.75rem", overflow: "hidden", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
                    <div style={{ background: C.gray50, padding: "0.55rem 0.875rem", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid " + C.gray200 }}>
                      <span style={{ fontSize: "0.82rem", color: C.gray900, fontWeight: 700 }}>{elev.title}</span>
                      <span style={{ fontSize: "0.62rem", color: C.gray400 }}>{elev.sheetRef}{elev.scale ? " · " + elev.scale : ""} · {Math.round(elevTotal)} SF</span>
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.69rem" }}>
                      <thead>
                        <tr style={{ background: C.gray50 }}>
                          {["ID", "Material", "Category", "Gross", "Openings", "Net SF", "Adj +15%"].map(h => (
                            <th key={h} style={{ padding: "0.4rem 0.625rem", textAlign: ["Gross", "Openings", "Net SF", "Adj +15%"].includes(h) ? "right" : "left", color: C.gray400, fontWeight: 600, borderBottom: "1px solid " + C.gray200, fontSize: "0.6rem", letterSpacing: "0.05em", textTransform: "uppercase" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(elev.zones || []).map((z, zi) => (
                          <tr key={zi} style={{ borderBottom: "1px solid " + C.gray100 }}>
                            <td style={{ padding: "0.4rem 0.625rem" }}>
                              <span style={{ display: "inline-block", width: 8, height: 8, borderRadius: 2, background: MATERIAL_COLORS[z.category] || "#6B7280", marginRight: 5, verticalAlign: "middle" }} />
                              <span style={{ color: C.bpsBlue, fontWeight: 600 }}>{z.materialId || "—"}</span>
                            </td>
                            <td style={{ padding: "0.4rem 0.625rem", color: C.gray700 }}>{z.materialName}</td>
                            <td style={{ padding: "0.4rem 0.625rem", color: C.gray400 }}>{z.category}</td>
                            <td style={{ padding: "0.4rem 0.625rem", textAlign: "right", color: C.gray500 }}>{Math.round(z.grossArea || 0)}</td>
                            <td style={{ padding: "0.4rem 0.625rem", textAlign: "right", color: C.gray400 }}>({Math.round(z.totalOpeningArea || 0)})</td>
                            <td style={{ padding: "0.4rem 0.625rem", textAlign: "right", color: C.gray700, fontWeight: 600 }}>{Math.round(z.netArea || 0)}</td>
                            <td style={{ padding: "0.4rem 0.625rem", textAlign: "right", color: C.bpsBlue, fontWeight: 700 }}>{Math.round((z.netArea || 0) * 1.15)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {(elev.flags || []).filter(Boolean).length > 0 && (
                      <div style={{ padding: "0.35rem 0.625rem", fontSize: "0.62rem", color: "#B45309", borderTop: "1px solid " + C.gray200, background: "#FFFBEB" }}>
                        ⚠ {elev.flags.filter(Boolean).join(" · ")}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* Idle state */}
          {phase === "idle" && (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: C.gray300, padding: "2rem", textAlign: "center" }}>
              <div style={{ fontSize: "5rem", marginBottom: "1.5rem", opacity: 0.3 }}>🏗</div>
              <div style={{ fontSize: "1rem", color: C.gray500, marginBottom: "0.625rem", fontWeight: 600 }}>Upload a blueprint PDF and click Run Analysis</div>
              <div style={{ fontSize: "0.75rem", color: C.gray400, lineHeight: 2 }}>
                Reads sheet index → Material legend → All elevations<br />
                Soffits · Returns · Per-elevation SF breakdown<br />
                Panels only · Evidence PDF + Excel export
              </div>
            </div>
          )}

          {/* Activity Log */}
          <div style={{ borderTop: "1px solid " + C.gray200, padding: "0.75rem 1.25rem", background: C.white, flexShrink: 0 }}>
            <div style={{ fontSize: "0.58rem", letterSpacing: "0.12em", color: C.gray400, textTransform: "uppercase", marginBottom: "0.4rem", fontWeight: 600 }}>
              Activity Log {isRunning ? "— checking every 5 seconds..." : ""}
            </div>
            <div ref={logRef} style={{ fontFamily: "'Courier New', monospace", fontSize: "0.67rem", maxHeight: 130, overflowY: "auto", lineHeight: 1.7 }}>
              {log.length === 0
                ? <span style={{ color: C.gray300 }}>Waiting...</span>
                : log.map((l, i) => <div key={i} style={{ color: logColor[l.level] || C.gray500 }}>{l.msg}</div>)}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
