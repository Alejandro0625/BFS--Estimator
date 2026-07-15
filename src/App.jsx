// BFS AI Estimator — UI build v2 (bigger header, cross-tab gradient design)
import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { Stage, Layer, Line, Circle, Image as KImage } from "react-konva";

const BACKEND = import.meta.env.VITE_BACKEND_URL || "";
const BLUE      = "#4A86C8";
const BLUE_DARK = "#2C5F9A";
const BLUE_PALE = "#EBF2FA";
const NAVY      = "#0C1B2E";
const NAVY_MID  = "#122035";
const NAVY_LT   = "#1E3A5F";

const MAT_COLORS = {
  "ACM Panel": BLUE, "MCM Panel": BLUE,
  "Fiber Cement Panel": "#22C55E", "Fiber Cement Plank": "#0D9488",
  "Nichiha Panel": "#8B5CF6", "Aluminum Wall Panel": "#06B6D4",
  "Perforated Metal Panel": "#F97316", "Soffit Panel": "#3B82F6",
  "Return/Trim": "#EC4899", "Other": "#9CA3AF",
};
const CLUSTER_COLORS = ["#3B82F6","#F97316","#22C55E","#EC4899","#EAB308","#8B5CF6","#14B8A6","#EF4444"];
// stable distinct color for materials not in MAT_COLORS (vector/callout names vary per job)
const hashColor = (s) => { let h=0; for(const ch of String(s||"")) h=(h*31+ch.charCodeAt(0))>>>0; return CLUSTER_COLORS[h%CLUSTER_COLORS.length]; };

/* Material-family fingerprint: lets the Scope tab's materials list pre-check the Takeoff's
   detected materials by MEANING, not exact spelling ("insulated metal wall panels" ↔
   "Vertical panel - 3.0' seams" both = metal-panel family). Verify-first hint, never a gate. */
const MAT_FAMILIES = [
  ["metal-panel", /metl[\s-]?span|\bacm\b|\bmcm\b|composite panel|aluminum|insulated metal|\bimp\b|rib panel|metal (wall )?panel|vertical panel|horizontal panel|corrugated/i],
  ["fiber-cement", /fiber|cement(?!itious mortar)|hardie|artisan|nichiha|swisspearl|cembrit/i],
  ["lap-siding", /\blap\b|clapboard|\bsiding\b|courses|board\s*(&|and)\s*batten|shake|shingle|longboard|woodtone/i],
  ["masonry", /brick|masonry|veneer|stone|terra[\s-]?cotta|cast stone|\bcmu\b|block|granite|precast/i],
  ["trim-linear", /\btrim\b|soffit|fascia|coping|flashing|\bpvc\b|azek|watertable|j[\s-]?channel/i],
  ["panel-generic", /\bpanel\b|\bpnl\b/i],
];
const matFamilies = (name) => MAT_FAMILIES.filter(([,re])=>re.test(String(name||""))).map(([f])=>f);
const matFamilyMatch = (a, b) => { const fa=matFamilies(a); return fa.length>0 && matFamilies(b).some(f=>fa.includes(f)); };

/* While the engine works, tell the story of what it's doing — in estimator language.
   Every line is a real step of the pipeline, not marketing. */
const NARRATION = [
  "📄 Finding the elevation sheets in the set…",
  "📏 Reading the scale three ways — title block, dimension strings, level markers…",
  "🧵 Tracing panel seams and lap courses in the drawing's own geometry…",
  "🎨 Reading color-coded siding areas and matching them to the legend…",
  "🏷️ Picking up material tags — MT-5, EIFS-3, SFC — straight off the walls…",
  "✂️ Splitting walls at structural joints, the way an estimator would…",
  "🪟 Cutting out windows and doors where the pattern breaks…",
  "🧮 Measuring every wall from the drawing's coordinates — never guessing…",
  "🔍 Checking soffits, returns and canopies so nothing gets forgotten…",
  "✅ Double-checking totals against the drawing's own schedule…",
];
function LiveNarration(){
  const [i,setI]=useState(0);
  useEffect(()=>{ const t=setInterval(()=>setI(v=>(v+1)%NARRATION.length),2300); return ()=>clearInterval(t); },[]);
  return (
    <div style={{minHeight:22,marginBottom:"0.6rem"}}>
      <div key={i} style={{fontSize:"0.82rem",fontWeight:600,color:"#AFCDEE",animation:"bfsFadeUp 0.45s cubic-bezier(.2,.8,.2,1)"}}>
        {NARRATION[i]}
      </div>
    </div>
  );
}

/* Default installed rates ($/SF, material+labor). Editable — these are ballpark starting points. */
const DEFAULT_RATES = {
  "ACM Panel":38,"MCM Panel":40,"Fiber Cement Panel":22,"Fiber Cement Plank":18,
  "Nichiha Panel":24,"Aluminum Wall Panel":32,"Perforated Metal Panel":45,
  "Soffit Panel":28,"Return/Trim":30,"Other":25,
};

/* Cross-run memory (per browser): remember hatch→material so repeat drawings get easier each time */
const LEARN_KEY="bfs_learn_hatch_v1";
const loadLearned=()=>{ try{ return JSON.parse(localStorage.getItem(LEARN_KEY))||{}; }catch{ return {}; } };
const saveLearned=m=>{ try{ localStorage.setItem(LEARN_KEY, JSON.stringify(m)); }catch{} };
const hatchSig=z=> z&&z.fill_color&&z.fill_color.length ? "fc:"+z.fill_color.map(c=>Math.round(c*255)).join(",") : (z&&z.material_type?"mt:"+z.material_type:null);

/* Shoelace area of a normalized polygon → square feet, given ft-per-paper-inch */
const polyAreaSF = (points, ftPerInch, W, H) => {
  if (!points || points.length < 3) return 0;
  let a = 0;
  for (let i = 0; i < points.length; i++) {
    const [x1, y1] = points[i], [x2, y2] = points[(i + 1) % points.length];
    a += (x1 * W) * (y2 * H) - (x2 * W) * (y1 * H);
  }
  const ftPerPt = ftPerInch / 72;
  return (Math.abs(a) / 2) * ftPerPt * ftPerPt;
};

/* Linear footage by edge class: top=coping/parapet, bottom=base/starter, vertical=corners/trim */
const linearFt = (zones, ftPerInch, W, H) => {
  const ftPerPt = ftPerInch / 72;
  let top=0, bottom=0, vert=0;
  (zones||[]).forEach(z=>{
    const pts=z.points; if(!pts||pts.length<3) return;
    const ys=pts.map(p=>p[1]); const ymin=Math.min(...ys), ymax=Math.max(...ys); const yr=(ymax-ymin)||1;
    for(let i=0;i<pts.length;i++){
      const [x1,y1]=pts[i], [x2,y2]=pts[(i+1)%pts.length];
      const dx=(x2-x1)*W, dy=(y2-y1)*H;
      const lenFt=Math.sqrt(dx*dx+dy*dy)*ftPerPt;
      if(Math.abs(dx)>=Math.abs(dy)){ if(((y1+y2)/2-ymin)/yr < 0.4) top+=lenFt; else bottom+=lenFt; }
      else vert+=lenFt;
    }
  });
  return { top:Math.round(top), bottom:Math.round(bottom), vert:Math.round(vert), total:Math.round(top+bottom+vert) };
};

/* ── Excel builder ── */
const buildExcel = (projectName, materials, pricing) => {
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString("en-US");
  const waste = (pricing?.wastePct || 0) / 100;
  const margin = (pricing?.marginPct || 0) / 100;
  const n11 = () => Array(11).fill(null);
  const n10 = () => Array(10).fill(null);
  const eRows = [];
  let er = n10(); er[0] = projectName||""; eRows.push(er);
  eRows.push(n10()); eRows.push(n10());
  er = n10(); er[0] = "PANELS"; eRows.push(er);
  er = n10(); er[0]="No."; er[1]="ACM/ACP"; er[2]="Quantity"; er[4]="Conv"; er[5]="Rate"; er[6]="Amount"; eRows.push(er);
  eRows.push(n10());
  const amtCells = [];
  materials.forEach((mat, idx) => {
    er = n10(); er[1] = mat.name; eRows.push(er);
    const row = eRows.length + 1;
    er = n10(); er[0]=idx+1; er[1]=mat.name; er[2]=Math.round(mat.sf*(1+waste)); er[4]=1; er[5]=mat.rate!=null?mat.rate:""; er[6]=`=C${row}*F${row}`; amtCells.push(`G${row}`); eRows.push(er);
  });
  while (eRows.length < 24) eRows.push(n10());
  const totalRow = eRows.length + 1;
  er = n10(); er[1]="Total "; er[2]=materials.reduce((s,m)=>s+Math.round(m.sf*(1+waste)),0); er[6]=amtCells.length?`=${amtCells.join("+")}`:0; eRows.push(er);
  eRows.push(n10());
  er = n10(); er[0]="PANEL BACK-UP SYSTEM- Z-Girts, Hat Channel, Insulation"; eRows.push(er);
  eRows.push(n10());
  er = n10(); er[1]="Furnish and install the quantity of new metal panels required\nAny exterior caulking required\nLifts / tie-off required per site policy\nAny break metal/flashing required\nStructural calculations and PE stamp\nShop drawings\nTaxes"; eRows.push(er);
  while (eRows.length < 34) eRows.push(n10());
  er = n10(); er[0]="SPECIFICATIONS"; eRows.push(er);
  eRows.push(n10());
  ["GC","Location","Profit/Non-Profit","Taxable/Non-Taxable","Prevailing Wage","Drawing Set","Building Height"].forEach(f=>{er=n10();er[0]=f;eRows.push(er);});
  const wsE = XLSX.utils.aoa_to_sheet(eRows);
  wsE["!cols"]=[5.14,57.43,11.86,4.14,5.43,12.86,12.43].map(w=>({wch:w}));
  wsE["!merges"]=[{s:{r:0,c:0},e:{r:1,c:7}},{s:{r:2,c:0},e:{r:2,c:7}},{s:{r:3,c:0},e:{r:3,c:6}},{s:{r:totalRow-1,c:0},e:{r:totalRow-1,c:9}},{s:{r:eRows.length-8,c:0},e:{r:eRows.length-8,c:6}}];
  XLSX.utils.book_append_sheet(wb, wsE, "Estimate");
  const pRows = [];
  let pr=n11(); pr[5]="PROPOSAL"; pRows.push(pr);
  pr=n11(); pr[5]="DATE:"; pr[6]=today; pRows.push(pr);
  pr=n11(); pr[0]="ACM. Trespa. Terracotta  &  Specialty Metal Panels"; pr[5]="This proposal may be withdrawn by us if not accepted within 30 days."; pRows.push(pr);
  pr=n11(); pr[0]="15 Erie Drive"; pr[5]="E-mail:"; pRows.push(pr);
  pr=n11(); pr[0]="Natick, MA 01760"; pr[5]=""; pRows.push(pr);
  pr=n11(); pr[0]="PH: 617-458-2000  "; pr[5]="Phone:"; pRows.push(pr);
  pr=n11(); pr[0]="To:"; pr[1]=""; pr[5]=""; pRows.push(pr);
  pr=n11(); pr[1]=""; pr[5]="Job Name / location:"; pRows.push(pr);
  pr=n11(); pr[1]=""; pr[5]=projectName||""; pRows.push(pr);
  pr=n11(); pr[1]=""; pr[5]="Job number: "; pRows.push(pr);
  pr=n11(); pr[5]=""; pRows.push(pr);
  pr=n11(); pr[0]="We hereby submit specifications and estimates for:"; pRows.push(pr);
  const mainMat=materials[0]?.name||"[Material]";
  const totalSF=materials.reduce((s,m)=>s+Math.round(m.sf),0);
  pr=n11(); pr[1]=`Install ${totalSF.toLocaleString()}sf of ${mainMat}.`; pRows.push(pr);
  ["Include all OSHA and fall protection compliance for the installation of panels","Include all staging and lifts for the performance of work.",`F&I ${materials.map(m=>m.name).join(", ")} as specified.`,"F&I all metal trim and accessories with panels as specified.","Remove and dispose of all job related debris to the general contractor's dumpster.","MA Sales Tax Included on all materials if applicable."].forEach((item,i)=>{pr=n11();pr[0]=i+1;pr[1]=item;pRows.push(pr);});
  pr=n11(); pr[1]="ADD/ALT: ENGINEERING DESIGN AND CALCULATIONS"; pr[7]=": $4,500"; pRows.push(pr);
  pr=n11(); pr[1]="NOTE: Air Vapor barrier behind all exterior panel system not included"; pRows.push(pr);
  pr=n11(); pr[0]="NOTE: THIS IS A BUDGETARY NUMBER ONLY PENDING FINAL SCOPE REVIEW & ENGINEERING CRITERIA"; pRows.push(pr);
  pr=n11(); pr[2]="PRICING GOOD FOR 30 DAYS DUE TO INDUSTRY-WIDE PRICE ESCALATION"; pRows.push(pr);
  pr=n11(); pr[1]="NIC: blocking, framing, plywood substrate, police details & street permits,"; pRows.push(pr);
  pr=n11(); pr[1]="thru-wall flashings, flashings not associated with the panel installations,"; pRows.push(pr);
  pr=n11(); pr[1]=" custom colors* (except where noted), winter conditions"; pRows.push(pr);
  pr=n11(); pr[1]='*** all contracts to have "BPS conditions for Metal Panels/Siding" attached.'; pRows.push(pr);
  pRows.push(n11());
  pr=n11(); pr[0]="We propose hereby to furnish materials and labor - complete in accordance with above specifications for the sum of:"; pRows.push(pr);
  pr=n11(); pr[7]="TOTAL:"; pr[8]=margin?`=Estimate!G${totalRow}*${(1+margin).toFixed(4)}`:`=Estimate!G${totalRow}`; pRows.push(pr);
  pr=n11(); pr[0]="Payment to be made as follows:"; pr[4]="AIA Format"; pRows.push(pr);
  pr=n11(); pr[6]="Akshita Patel"; pRows.push(pr);
  pr=n11(); pr[7]="Authorized Signature"; pRows.push(pr);
  pr=n11(); pr[0]="All material to be as specified. All work to be performed in a professional manner according to standard practices. Any alteration or deviation from above specifications involving additional costs will be executed only upon written orders and will be an extra charge."; pRows.push(pr);
  pr=n11(); pr[0]="Acceptance of Proposal — The above prices, specifications and conditions are satisfactory and hereby accepted."; pRows.push(pr);
  pr=n11(); pr[3]="Date:"; pr[7]="Signature"; pRows.push(pr);
  const wsP = XLSX.utils.aoa_to_sheet(pRows);
  wsP["!cols"]=[7.14,9.14,12,12,16.71,9.57,9.14,12,12,12,12].map(w=>({wch:w}));
  wsP["!merges"]=[{s:{r:0,c:5},e:{r:0,c:10}},{s:{r:1,c:6},e:{r:1,c:7}},{s:{r:2,c:5},e:{r:2,c:10}},{s:{r:3,c:5},e:{r:3,c:10}},{s:{r:4,c:5},e:{r:4,c:10}},{s:{r:5,c:5},e:{r:5,c:10}},{s:{r:6,c:5},e:{r:6,c:10}},{s:{r:7,c:5},e:{r:7,c:10}},{s:{r:8,c:5},e:{r:8,c:10}},{s:{r:9,c:5},e:{r:9,c:10}},{s:{r:10,c:5},e:{r:10,c:10}},{s:{r:pRows.length-4,c:0},e:{r:pRows.length-3,c:5}}];
  XLSX.utils.book_append_sheet(wb, wsP, "Proposal");
  return wb;
};

/* ── Interactive Takeoff ── */
/* ── Bluebeam-style deep zoom: wheel zooms to the cursor, shift/middle-drag pans; past 2.2×
     it fetches a crisp HIGH-RES render of just the viewport so textures stay sharp. ── */
function DeepZoom({ BACKEND, jobId, pageNum, children }) {
  const boxRef = useRef(); const innerRef = useRef();
  const [k, setK] = useState(1);
  const [tx, setTx] = useState(0); const [ty, setTy] = useState(0);
  const [crop, setCrop] = useState(null);
  const drag = useRef(null); const timer = useRef(); const justDragged = useRef(false);
  const stateRef = useRef({ k: 1, tx: 0, ty: 0 }); stateRef.current = { k, tx, ty };
  useEffect(() => () => clearTimeout(timer.current), []);
  useEffect(() => { setK(1); setTx(0); setTy(0); setCrop(null); }, [jobId, pageNum]);
  const schedule = (nk, nx, ny) => {
    clearTimeout(timer.current);
    timer.current = setTimeout(() => {
      if (nk < 2.2 || !jobId) { setCrop(null); return; }
      const box = boxRef.current, inner = innerRef.current; if (!box || !inner) return;
      const w = inner.offsetWidth, h = inner.offsetHeight; if (!w || !h) return;
      const x0 = Math.max(0, -nx / (nk * w)), y0 = Math.max(0, -ny / (nk * h));
      const x1 = Math.min(1, x0 + box.clientWidth / (nk * w)), y1 = Math.min(1, y0 + box.clientHeight / (nk * h));
      const url = `${BACKEND}/page-crop/${jobId}/${pageNum}?x0=${x0.toFixed(4)}&y0=${y0.toFixed(4)}&x1=${x1.toFixed(4)}&y1=${y1.toFixed(4)}&px=2000`;
      const im = new window.Image();
      im.onload = () => setCrop({ url, x0, y0, x1, y1 });
      im.src = url;
    }, 350);
  };
  useEffect(() => {
    const box = boxRef.current; if (!box) return;
    const onWheel = e => {
      e.preventDefault();
      const { k, tx, ty } = stateRef.current;
      const r = box.getBoundingClientRect();
      const mx = e.clientX - r.left, my = e.clientY - r.top;
      let nk = Math.max(1, Math.min(14, k * (e.deltaY < 0 ? 1.2 : 1 / 1.2)));
      let nx = mx - ((mx - tx) / k) * nk, ny = my - ((my - ty) / k) * nk;
      if (nk === 1) { nx = 0; ny = 0; }
      setK(nk); setTx(nx); setTy(ny); schedule(nk, nx, ny);
    };
    box.addEventListener("wheel", onWheel, { passive: false });
    return () => box.removeEventListener("wheel", onWheel);
  }, [jobId, pageNum]);
  const onMouseDown = e => {
    const { k } = stateRef.current;
    if (e.button === 1 || (e.button === 0 && e.shiftKey && k > 1)) {
      e.preventDefault();
      drag.current = { sx: e.clientX, sy: e.clientY, tx: stateRef.current.tx, ty: stateRef.current.ty };
    }
  };
  const onMouseMove = e => { if (!drag.current) return; const d = drag.current; setTx(d.tx + e.clientX - d.sx); setTy(d.ty + e.clientY - d.sy); };
  const onMouseUp = e => {
    if (drag.current) {
      if (Math.abs(e.clientX - drag.current.sx) + Math.abs(e.clientY - drag.current.sy) > 4) justDragged.current = true;
      drag.current = null;
      schedule(stateRef.current.k, stateRef.current.tx, stateRef.current.ty);
    }
  };
  const onClickCapture = e => { if (justDragged.current) { justDragged.current = false; e.stopPropagation(); e.preventDefault(); } };
  const innerW = innerRef.current?.offsetWidth || 0;
  const innerH = innerRef.current?.offsetHeight || 0;
  return (
    <div ref={boxRef} onMouseDown={onMouseDown} onMouseMove={onMouseMove} onMouseUp={onMouseUp} onMouseLeave={onMouseUp} onClickCapture={onClickCapture}
      style={{ overflow: "hidden", position: "relative", maxWidth: "100%" }}>
      <div ref={innerRef} style={{ transform: `translate(${tx}px,${ty}px) scale(${k})`, transformOrigin: "0 0", position: "relative", display: "inline-block" }}>
        {children}
        {crop && k >= 2.2 && innerW > 0 && (
          <img src={crop.url} alt="" style={{ position: "absolute", left: crop.x0 * innerW, top: crop.y0 * innerH,
            width: (crop.x1 - crop.x0) * innerW, height: (crop.y1 - crop.y0) * innerH, pointerEvents: "none", zIndex: 1 }}/>
        )}
      </div>
      {k > 1 && <div style={{ position: "absolute", top: 8, right: 8, fontSize: "0.6rem", padding: "0.2rem 0.55rem", borderRadius: 12, background: "rgba(12,27,46,0.78)", color: "#9FC3EA", pointerEvents: "none", zIndex: 3 }}>{k.toFixed(1)}× · wheel = zoom · shift-drag = pan</div>}
    </div>
  );
}

function InteractiveView({ results, BACKEND, assignments, setAssignments, groupRename={}, setGroupRename=()=>{}, setResults, hiddenIds={}, setHiddenIds=()=>{}, deletedStack=[], setDeletedStack=()=>{}, bucketShapes=[], setBucketShapes=()=>{}, bucketColorNames={}, setBucketColorNames=()=>{} }) {
  const [elevIdx, setElevIdx] = useState(0);
  const [pageImage, setPageImage] = useState(null);
  const [imgLoaded, setImgLoaded] = useState(false);
  const [pagePolygons, setPagePolygons] = useState([]);
  const [pageDims, setPageDims] = useState({ width:612, height:792 });
  const [activeGroup, setActiveGroup] = useState(null);
  const [groupNameDraft, setGroupNameDraft] = useState("");
  const [imgNaturalSize, setImgNaturalSize] = useState({ w:1, h:1 });
  const [calibMode, setCalibMode] = useState(false);
  const [calibPts, setCalibPts] = useState([]);
  const [pageScales, setPageScales] = useState({});   // scale is PER PAGE (sheets mix scales); calibFt derives from it below
  const [realDist, setRealDist] = useState("");
  // Bucket-fill (coloring-book) assist — click a wall → exact SF from vector geometry. Additive/opt-in.
  const [bucketMode, setBucketMode] = useState(false);
  const [splitMode, setSplitMode] = useState(false);   // ✂ estimator's knife: click a face's boundary → it splits there
  const BUCKET_COLS = ["#22D3EE","#E85DA0","#3FB36B","#F0A23C","#9B6FD4","#F4D03F"]; // cyan first — his Bluebeam PNL color
  const [curBucketColor, setCurBucketColor] = useState("#22D3EE");
  const [cornerMode, setCornerMode] = useState(false);   // fallback when a bucket click "leaks"
  const [cornerPts, setCornerPts] = useState([]);
  // bucketShapes is LIFTED to the app (props) so bucket-added walls flow into the bid & persist
  const [snapMsg, setSnapMsg] = useState("");
  const [snapBusy, setSnapBusy] = useState(false);
  // Material-group PREVIEW — AI suggests groups (may be imperfect); estimator SELECTS the ones she's bidding.
  const [groupMode, setGroupMode] = useState(false);
  const [previewGroups, setPreviewGroups] = useState([]);
  const [selGroups, setSelGroups] = useState({});   // { group_id: materialName }
  const [groupBusy, setGroupBusy] = useState(false);
  // hiddenIds / deletedStack are LIFTED to the parent so deletes save + restore with the bid
  const imgRef = useRef();
  const svgRef = useRef();
  const elevations = results.takeoffData.filter(e => e.pageNumber);
  const elev = elevations[elevIdx];
  const pageNum = elev?.pageNumber;
  // this page's scale: one-click chip or 2-point calibrate sets it; same recompute path as before
  const calibFt = pageScales[pageNum] ?? null;
  const setCalibFt = v => setPageScales(prev=>{const n={...prev}; if(v==null) delete n[pageNum]; else n[pageNum]=v; return n;});

  useEffect(() => {
    if (!pageNum || !results.jobId) return;
    setPageImage(null); setImgLoaded(false); setPagePolygons([]); setActiveGroup(null); setCalibMode(false); setCalibPts([]);
    setGroupMode(false); setPreviewGroups([]); setSelGroups({}); setBucketMode(false); setCornerMode(false);
    fetch(BACKEND+"/polygons/"+results.jobId+"/"+pageNum)
      .then(r=>r.ok?r.json():{polygons:[],width:612,height:792})
      .then(d=>{setPagePolygons(d.polygons||[]);setPageDims({width:d.width||612,height:d.height||792});})
      .catch(()=>{});
    setPageImage(BACKEND+"/page-image/"+results.jobId+"/"+pageNum);
  }, [elevIdx, pageNum, results.jobId, BACKEND]);

  // NEW JOB → clear all per-job marks (bucket fills, group picks). Without this, shapes from the
  // previous drawing would silently flow into the next job's totals/Excel — a money bug.
  useEffect(()=>{ setSelGroups({}); setPreviewGroups([]); setSnapMsg(""); setPageScales({}); }, [results.jobId]);  // bucketShapes reset at app level (survives tab switches)
  // Pull shared learning from the server into local memory (so repeats are pre-identified)
  useEffect(()=>{
    fetch(BACKEND+"/recall").then(r=>r.ok?r.json():null).then(d=>{
      if(d&&d.hatches){ const m=loadLearned(); let ch=false; for(const k in d.hatches){ if(!m[k]){ m[k]=d.hatches[k]; ch=true; } } if(ch) saveLearned(m); }
    }).catch(()=>{});
  },[BACKEND]);

  const polyMethod = pagePolygons[0]?.source||(pagePolygons.length>0?"vector":"box");
  const rawZonesAll = pagePolygons.length>0 ? pagePolygons : (elev?.zones||[]).map((z,i)=>({
    id:i,points:[[z.x0pct/100,z.y0pct/100],[z.x1pct/100,z.y0pct/100],[z.x1pct/100,z.y1pct/100],[z.x0pct/100,z.y1pct/100]],
    area_sf:z.netArea||0,cx:(z.x0pct+z.x1pct)/200,cy:(z.y0pct+z.y1pct)/200,source:"box",
  }));
  const rawZones = rawZonesAll.filter(z=>!hiddenIds[pageNum+":"+z.id]);   // deleted highlights stay gone
  // When the user calibrates the scale, recompute SF from polygon geometry — but NEVER for
  // zones whose SF is the estimator's own measured label (sf_exact) or texture net SF. Those are ground truth.
  const displayZones = calibFt
    ? rawZones.map(z => (z.source && z.source!=="box" && z.source!=="texture" && !z.sf_exact)
        ? {...z, area_sf: polyAreaSF(z.points, calibFt, pageDims.width, pageDims.height)}
        : z)   // digitize-markup (exact label) + texture groups keep their trusted backend SF
    : rawZones;
  // Effective scale: calibrated value, else back it out from a zone's known SF + geometry
  const effFtPerInch = calibFt || (()=>{
    for(const z of rawZones){
      if(z.area_sf>0 && z.points?.length>=3 && z.source!=="texture"){
        const shoePts = polyAreaSF(z.points, 72, pageDims.width, pageDims.height);
        if(shoePts>0) return 72*Math.sqrt(z.area_sf/shoePts);
      }
    }
    return 8;
  })();
  const lf = linearFt(displayZones, effFtPerInch, pageDims.width, pageDims.height);
  // Group key: same cluster OR same fill color = same hatch/pattern
  const gkey = z => z.cluster_id!==undefined ? "c_"+z.cluster_id : (z.fill_color&&z.fill_color.length ? "f_"+z.fill_color.join(",") : "z_"+z.id);
  const colorGroups = {};
  displayZones.forEach(z=>{const k=gkey(z);if(!colorGroups[k])colorGroups[k]=[];colorGroups[k].push(z.id);});
  const clusterSummary = {};
  displayZones.forEach(z=>{const k=z.cluster_id!==undefined?z.cluster_id:-1;if(!clusterSummary[k])clusterSummary[k]={total_sf:0,count:0,color:"#94A3B8"};clusterSummary[k].total_sf+=z.area_sf||0;clusterSummary[k].count+=1;clusterSummary[k].color=z.cluster_id!==undefined?CLUSTER_COLORS[z.cluster_id%CLUSTER_COLORS.length]:"#94A3B8";});
  const assignKey = id => elevIdx+":"+id;
  const getAssignment = id => assignments[assignKey(id)];
  // The currently-selected hatch group (all areas sharing the clicked pattern)
  const selectedIds = activeGroup ? (colorGroups[activeGroup]||[]) : [];
  const selectedZones = displayZones.filter(z=>selectedIds.includes(z.id));
  const selectedSF = selectedZones.reduce((s,z)=>s+(z.area_sf||0),0);
  const groupSig = hatchSig(selectedZones.find(z=>hatchSig(z)));
  // Pre-fill the rename box with the group's current name when a group is selected
  useEffect(()=>{
    if(!activeGroup){ setGroupNameDraft(""); return; }
    const z0=selectedZones[0];
    const base=z0?.material || z0?.category || "";
    const cur=groupRename[base] || selectedZones.map(z=>getAssignment(z.id)).find(a=>a)?.name || base || "";
    setGroupNameDraft(cur);
  // eslint-disable-next-line
  },[activeGroup]);
  // Renaming a group applies it to that SAME texture group on EVERY page of the drawing,
  // remembers it for next time, and silently saves the correction as training data — no button.
  const renameGroup = (nm)=>{
    const v=(nm||"").trim(); if(!v) return;
    const z0=selectedZones[0]; const base=z0?.material || z0?.category || "group";
    setGroupRename(prev=>({...prev,[base]:v}));
    if(groupSig){ const m=loadLearned(); m[groupSig]={category:v,materialName:v,id:v,at:Date.now()}; saveLearned(m); }
    if(results?.jobId) fetch(BACKEND+"/learn",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum,source:"rename",group:base,name:v,shapes:selectedZones.map(z=>({points:z.points,name:v,color:z.fill_color,type:"add"}))})}).catch(()=>{});
  };
  const learnedMat = groupSig ? loadLearned()[groupSig] : null;
  const assignGroup = mat => {
    setAssignments(prev=>{const n={...prev};selectedZones.forEach(z=>{n[assignKey(z.id)]={...mat,area_sf:z.area_sf||0};});return n;});
    const sig=hatchSig(selectedZones.find(z=>hatchSig(z)));
    if(sig){ const m=loadLearned(); m[sig]={category:mat.category,materialName:mat.name||mat.category,id:mat.id,at:Date.now()}; saveLearned(m);
      fetch(BACKEND+"/learn",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({hatches:[{signature:sig,category:mat.category,materialName:mat.name||mat.category,materialId:mat.id}]})}).catch(()=>{});
    }
    setActiveGroup(null);
  };
  const removeGroup = () => { setAssignments(prev=>{const n={...prev};selectedIds.forEach(id=>delete n[assignKey(id)]);return n;}); };
  // DELETE a bad highlight: hides the shapes AND subtracts their SF from the page's real
  // takeoff zones (summary/pricing/Excel follow) — junk detections can't pollute the bid.
  const deleteGroup = () => {
    if(!selectedZones.length) return;
    const byMat = {};
    selectedZones.forEach(z=>{ const k=z.material||z.category||""; byMat[k]=(byMat[k]||0)+(z.area_sf||0); });
    // undo snapshot: the exact zones array of this page before the subtraction
    const pageBefore = (results.takeoffData||[]).find(e=>e.pageNumber===pageNum);
    setDeletedStack(prev=>[...prev,{ page:pageNum, ids:selectedZones.map(z=>z.id), zonesBefore:(pageBefore?.zones||[]).map(z=>({...z})) }]);
    setHiddenIds(prev=>{ const n={...prev}; selectedZones.forEach(z=>{ n[pageNum+":"+z.id]=true; }); return n; });
    setAssignments(prev=>{ const n={...prev}; selectedIds.forEach(id=>delete n[assignKey(id)]); return n; });
    if(setResults) setResults(prev=>({ ...prev, takeoffData: prev.takeoffData.map(e=>{
      if(e.pageNumber!==pageNum) return e;
      const zones=(e.zones||[]).map(z=>({...z}));
      Object.entries(byMat).forEach(([mat,sf])=>{
        let z = zones.find(zz=>zz.materialName===mat||zz.category===mat);
        if(!z&&zones.length===1) z=zones[0];
        if(z){ z.netArea=Math.max(0,(z.netArea||0)-sf); z.grossArea=Math.max(0,(z.grossArea||0)-sf); }
      });
      return {...e, zones: zones.filter(z=>(z.netArea||0)>1)};
    })}));
    // teach the flywheel: this pattern was NOT cladding here
    if(results?.jobId) fetch(BACKEND+"/learn",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum,source:"delete",shapes:selectedZones.map(z=>({points:z.points,name:"NOT-CLADDING",color:z.fill_color,type:"delete"}))})}).catch(()=>{});
    setActiveGroup(null);
  };
  // Undo the last delete: restore the exact zone numbers and un-hide the shapes
  const undoDelete = () => {
    setDeletedStack(prev=>{
      if(!prev.length) return prev;
      const last = prev[prev.length-1];
      setHiddenIds(h=>{ const n={...h}; last.ids.forEach(id=>delete n[last.page+":"+id]); return n; });
      if(setResults) setResults(r=>({ ...r, takeoffData: r.takeoffData.map(e=>e.pageNumber===last.page?{...e, zones:last.zonesBefore.map(z=>({...z}))}:e) }));
      return prev.slice(0,-1);
    });
  };
  const exportInteractiveExcel = () => {
    const badScale=(results.takeoffData||[]).filter(e=>(e.zones||[]).some(z=>(z.netArea||0)>0)&&(e.scaleSource==="default"||(!e.verifiedScale&&!e.scale))).length;
    if(badScale>0 && !window.confirm("⚠ "+badScale+" page(s) used a DEFAULT scale (couldn't read it) — SF could be far off. Calibrate first.\n\nExport anyway?")) return;
    const mt={};
    Object.values(assignments).forEach(a=>{const k=a.materialName||a.category||"Panel";if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=a.area_sf||0;});
    bucketShapes.forEach(s=>{const k=s.material||"Cladding (bucket)";if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=s.area_sf||0;});
    selGroupList.forEach(g=>{const k=(selGroups[g.group]||"Cladding (preview)")+" (preview)";if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=g.approx_sf||0;});
    const wb=buildExcel(results.projName||"Project",Object.values(mt));
    XLSX.writeFile(wb,"BFS_Takeoff_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
  };
  const totals={};
  Object.values(assignments).forEach(a=>{if(!totals[a.category])totals[a.category]=0;totals[a.category]+=a.area_sf||0;});
  const grandTotal=Object.values(totals).reduce((s,v)=>s+v,0);
  const svgW=imgRef.current?.offsetWidth||imgNaturalSize.w;
  const svgH=imgRef.current?.offsetHeight||imgNaturalSize.h;
  const toSVGPoints=pts=>pts.map(([nx,ny])=>`${(nx*pageDims.width).toFixed(1)},${(ny*pageDims.height).toFixed(1)}`).join(" ");
  const getSvgPoint=evt=>{
    const svg=svgRef.current; if(!svg) return null;
    const pt=svg.createSVGPoint(); pt.x=evt.clientX; pt.y=evt.clientY;
    const m=svg.getScreenCTM(); if(!m) return null;
    const p=pt.matrixTransform(m.inverse());
    return { x:p.x/pageDims.width, y:p.y/pageDims.height };
  };
  // ✂ SPLIT: the estimator clicks the material boundary on a face — the backend cuts it
  // along the nearest structural line; both halves keep their share of the net SF. Every
  // split is saved as boundary training data (how the system learns HIS material lines).
  const doSplit=async p=>{
    const inPoly=(pt,poly)=>{let ins=false;for(let i=0,j=poly.length-1;i<poly.length;j=i++){const[x1,y1]=poly[i],[x2,y2]=poly[j];if((y1>pt.y)!==(y2>pt.y)&&pt.x<(x2-x1)*(pt.y-y1)/(y2-y1+1e-12)+x1)ins=!ins;}return ins;};
    const cands=displayZones.filter(z=>(z.points||[]).length>=3&&inPoly(p,z.points));
    if(!cands.length){ setSnapMsg("Click inside the face you want to split"); return; }
    const target=cands.reduce((a,b)=>(a.area_sf||0)<=(b.area_sf||0)?a:b);   // smallest face under the click
    setSnapBusy(true); setSnapMsg("Splitting…");
    try{
      const r=await fetch(BACKEND+"/split-shape",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({jobId:results.jobId,page:pageNum,points:target.points,area_sf:target.area_sf||0,holes:target.holes||[],click:[p.x,p.y]})}).then(r=>r.json());
      if(r.status==="ok"&&r.shapes?.length===2){
        setPagePolygons(prev=>{
          const nid=Math.max(0,...prev.map(z=>z.id||0));
          return prev.flatMap(z=>z.id===target.id
            ?r.shapes.map((s,i)=>({...z,id:nid+1+i,points:s.points,area_sf:s.area_sf,holes:s.holes||[],cx:s.points.reduce((a,q)=>a+q[0],0)/s.points.length,cy:s.points.reduce((a,q)=>a+q[1],0)/s.points.length}))
            :[z]);
        });
        setSnapMsg(`✂ split into ${Math.round(r.shapes[0].area_sf).toLocaleString()} + ${Math.round(r.shapes[1].area_sf).toLocaleString()} SF — click each to tag its material`);
        fetch(BACKEND+"/learn",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum,source:"split",axis:r.axis,cut_at:r.cut_at,face:target.points})}).catch(()=>{});
      } else setSnapMsg("Couldn't split there — click nearer the boundary line");
    }catch{ setSnapMsg("Split failed — check connection"); }
    setSnapBusy(false);
  };
  const handleSvgClick=evt=>{
    if(suppressClickRef.current){ suppressClickRef.current=false; return; }   // that click was a vertex drag
    if(calibMode){ const p=getSvgPoint(evt); if(p) setCalibPts(prev=>prev.length>=2?[p]:[...prev,p]); return; }
    if(splitMode&&!snapBusy){ const p=getSvgPoint(evt); if(p) doSplit(p); return; }
    if(!bucketMode||snapBusy) return;
    const p=getSvgPoint(evt); if(!p) return;
    if(cornerMode){ setCornerPts(prev=>[...prev,p]); return; }
    doBucket(p);
  };
  const addBucketShape=(points,area_sf,extra={})=>{
    // sfScale = the shape's own implied ft/in (from the backend's exact SF), so vertex edits recompute
    // with the SAME scale the number was born with — Bluebeam-style: drag a corner, SF follows exactly.
    let sfScale=null;
    if(extra.gross_sf){ const a72=polyAreaSF(points,72,pageDims.width,pageDims.height); if(a72>0) sfScale=72*Math.sqrt(extra.gross_sf/a72); }
    setBucketShapes(prev=>[...prev,{id:Date.now()+Math.random(),page:pageNum,points,area_sf,material:"",sfScale,...extra}]);
  };
  const dragRef=useRef(null);           // {sid, vi} while a vertex is being dragged
  const suppressClickRef=useRef(false); // swallow the click that follows a handle drag
  const [editShape,setEditShape]=useState(null);
  const handleSvgMove=e=>{
    const d=dragRef.current; if(!d) return;
    const p=getSvgPoint(e); if(!p) return;
    setBucketShapes(prev=>prev.map(sh=>{
      if(sh.id!==d.sid) return sh;
      const pts=sh.points.map((pt,ix)=>ix===d.vi?[Math.min(Math.max(p.x,0),1),Math.min(Math.max(p.y,0),1)]:pt);
      if(!sh.sfScale) return {...sh, points:pts};
      const gross=polyAreaSF(pts,sh.sfScale,pageDims.width,pageDims.height);
      return {...sh, points:pts, gross_sf:Math.round(gross*10)/10, area_sf:Math.round(Math.max(0,gross-(sh.opening_sf||0))*10)/10};
    }));
  };
  const endDrag=()=>{ if(dragRef.current){ dragRef.current=null; suppressClickRef.current=true; } };
  // Veto a deducted opening: add its SF back and drop its outline (never a silent deduction)
  const vetoOpening=(shapeId,idx)=>setBucketShapes(prev=>prev.map(s=>{
    if(s.id!==shapeId||!s.openings) return s;
    const op=s.openings[idx]; if(!op) return s;
    const per=(s.opening_sf||0)/Math.max(s.openings.length,1);
    return {...s, openings:s.openings.filter((_,i)=>i!==idx), opening_sf:Math.max(0,(s.opening_sf||0)-per), area_sf:(s.area_sf||0)+per};
  }));
  const doBucket=async p=>{
    setSnapBusy(true); setSnapMsg("Filling…");
    try{
      const r=await fetch(BACKEND+"/snap-fill",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum,point:[p.x,p.y]})}).then(r=>r.json());
      if(r.status==="ok"){
        // HIS BLUEBEAM HABIT: the selected COLOR is the material — fills paint in it, and the
        // color's name (set once) flows to every shape of that color, the Budget and the Excel
        const mat=bucketColorNames[curBucketColor]||r.material||"";
        addBucketShape(r.points,r.area_sf,{holes:r.holes||[],pattern_sig:r.pattern_sig||"",material:mat,color:curBucketColor});
        (r.siblings||[]).forEach(s=>addBucketShape(s.points,s.area_sf,{holes:s.holes||[],material:mat,color:curBucketColor}));
        const n=1+(r.siblings||[]).length;
        const tot=r.pattern_total_sf||r.area_sf;
        setSnapMsg(`✓ ${n>1?`${n} areas with this pattern · `:""}${Math.round(tot).toLocaleString()} SF${(r.holes||[]).length?" net":""}${mat?` · ${mat}`:""}`);
      }
      else if(r.status==="leak"){ setCornerMode(true); setCornerPts([]); setSnapMsg("Open field — click each corner (peaks too), then Finish"); }
      else setSnapMsg("Couldn't fill there — try clicking the corners");
    }catch(e){ setSnapMsg("Snap failed — check connection"); }
    setSnapBusy(false);
  };
  const finishCorners=async()=>{
    if(cornerPts.length<3){ setSnapMsg("Click at least 3 corners"); return; }
    setSnapBusy(true);
    try{
      const r=await fetch(BACKEND+"/snap-fill",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum,corners:cornerPts.map(c=>[c.x,c.y])})}).then(r=>r.json());
      if(r.status==="ok"){
        addBucketShape(r.points,r.area_sf,{gross_sf:r.gross_sf,opening_sf:r.opening_sf||0,openings:r.openings||[],openings_review:!!r.openings_review});
        setSnapMsg((r.opening_sf>0?`✓ ${Math.round(r.area_sf).toLocaleString()} SF net (−${Math.round(r.opening_sf).toLocaleString()} openings — tap ✕ on any to undo)`:`✓ ${Math.round(r.area_sf).toLocaleString()} SF`)+(r.openings_review?" ⚠ check deductions":""));
      }
      else setSnapMsg("Couldn't snap those corners");
    }catch(e){ setSnapMsg("Snap failed"); }
    setCornerMode(false); setCornerPts([]); setSnapBusy(false);
  };
  const cancelCorners=()=>{ setCornerMode(false); setCornerPts([]); setSnapMsg(""); };
  const removeBucketShape=id=>setBucketShapes(prev=>prev.filter(s=>s.id!==id));
  const setBucketMaterial=(id,mat)=>setBucketShapes(prev=>prev.map(s=>s.id===id?{...s,material:mat}:s));
  // on blur/Enter: naming one shape names every UNNAMED shape with the SAME PATTERN
  // (estimator's rule: regions sharing a pattern are the same material, job-wide)
  const propagateMaterial=(id)=>setBucketShapes(prev=>{
    const src=prev.find(s=>s.id===id);
    const sig=src?.pattern_sig; const mat=(src?.material||"").trim();
    if(!sig||sig==="plain"||!mat) return prev;
    return prev.map(s=>(s.id!==id&&s.pattern_sig===sig&&!(s.material||"").trim())?{...s,material:mat}:s);
  });
  const pageBucketShapes=bucketShapes.filter(s=>s.page===pageNum);
  const bucketTotalSF=pageBucketShapes.reduce((s,x)=>s+(x.area_sf||0),0);
  const loadGroups=async()=>{
    setGroupBusy(true); setPreviewGroups([]); setSelGroups({});
    try{
      const r=await fetch(BACKEND+"/material-groups",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum})}).then(r=>r.json());
      setPreviewGroups(r.groups||[]);
    }catch(e){ setPreviewGroups([]); }
    setGroupBusy(false);
  };
  const toggleGroup=g=>setSelGroups(prev=>{const n={...prev}; if(n[g.group]!==undefined) delete n[g.group]; else n[g.group]=""; return n;});
  // ⚡ EXACT-ON-SELECT: turn a blocky preview group into corner-snapped exact shapes (bid-grade SF)
  const refineGroup=async g=>{
    setGroupBusy(true);
    try{
      const r=await fetch(BACKEND+"/refine-group",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jobId:results.jobId,page:pageNum,patches:g.patches})}).then(r=>r.json());
      if(r.status==="ok"&&r.shapes.length){
        r.shapes.forEach(s=>addBucketShape(s.points,s.area_sf,{gross_sf:s.gross_sf,opening_sf:s.opening_sf,openings:s.openings,openings_review:s.openings_review,material:selGroups[g.group]||""}));
        toggleGroup(g);
        setSnapMsg(`⚡ made exact: ${r.shapes.length} shapes · ${Math.round(r.total_sf).toLocaleString()} SF`);
      } else setSnapMsg("Couldn't refine that group");
    }catch(e){ setSnapMsg("Refine failed — check connection"); }
    setGroupBusy(false);
  };
  const setGroupMat=(gid,mat)=>setSelGroups(prev=>({...prev,[gid]:mat}));
  const selGroupList=previewGroups.filter(g=>selGroups[g.group]!==undefined);
  const selGroupSF=selGroupList.reduce((s,g)=>s+(g.approx_sf||0),0);
  const applyCalibration=()=>{
    if(calibPts.length<2) return;
    const [a,b]=calibPts;
    const dx=(b.x-a.x)*pageDims.width, dy=(b.y-a.y)*pageDims.height;
    const distPts=Math.sqrt(dx*dx+dy*dy);
    const feet=parseFloat(realDist);
    if(!feet||!distPts) return;
    setCalibFt(feet/(distPts/72));
    setCalibMode(false); setCalibPts([]); setRealDist("");
  };
  const matList=results.legend.length>0?results.legend:Object.keys(MAT_COLORS).map(cat=>({id:cat.substring(0,3).toUpperCase()+"-1",name:cat,category:cat}));

  return (
    <div style={{display:"flex",height:"100%",overflow:"hidden",background:NAVY,fontFamily:"'Inter','Segoe UI',sans-serif"}}>
      {/* Elevation list */}
      <div style={{width:185,borderRight:"1px solid "+NAVY_LT,overflowY:"auto",flexShrink:0,background:NAVY_MID}}>
        <div style={{padding:"0.875rem",fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,borderBottom:"1px solid "+NAVY_LT}}>Elevations</div>
        {elevations.map((e,i)=>{
          const assigned=Object.keys(assignments).filter(k=>k.startsWith(i+":")).length;
          return <div key={i} onClick={()=>setElevIdx(i)} style={{padding:"0.65rem 0.875rem",cursor:"pointer",borderBottom:"1px solid "+NAVY_LT,background:i===elevIdx?NAVY_LT:"transparent",borderLeft:i===elevIdx?"3px solid "+BLUE:"3px solid transparent"}}>
            <div style={{fontSize:"0.72rem",color:i===elevIdx?"#E2E8F0":"#94A3B8",fontWeight:i===elevIdx?600:400,lineHeight:1.3}}>{e.title||"Page "+e.pageNumber}</div>
            <div style={{fontSize:"0.6rem",color:assigned>0?"#4ADE80":"#475569",marginTop:3}}>{assigned>0?`✓ ${assigned} assigned`:`${(e.zones||[]).length} zones · p.${e.pageNumber}`}</div>
          </div>;
        })}
      </div>
      {/* Drawing */}
      <div style={{flex:1,overflow:"auto",display:"flex",flexDirection:"column",alignItems:"center",padding:"1rem",gap:"0.75rem"}}>
        <div style={{alignSelf:"stretch",display:"flex",alignItems:"center",gap:"0.5rem",flexWrap:"wrap"}}>
          <div style={{fontSize:"0.65rem",color:"#64748B",background:NAVY_LT,padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid #2D5280"}}>
            {displayZones.length===0?"No surfaces on this page"
              :polyMethod==="bluebeam"?`📐 Bluebeam markup — ${displayZones.length} surfaces (exact)`
              :polyMethod==="vector"||polyMethod==="vector_cluster"?`📐 Read from drawing geometry — ${displayZones.length} surfaces`
              :polyMethod==="model"?`🧠 AI model — ${displayZones.length} surfaces (verify)`
              :polyMethod==="texture"?`🎨 AI texture — ${displayZones.length} surfaces (verify)`
              :`${displayZones.length} surfaces`}
          </div>
          <div style={{fontSize:"0.62rem",color:"#5E7BA0",padding:"0.3rem 0.6rem",borderRadius:20,border:"1px dashed #2D5280"}}>🔍 wheel = zoom · shift-drag = pan</div>
          {deletedStack.length>0&&<button onClick={undoDelete} style={{fontSize:"0.65rem",padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid #B45309",background:"#451A03",color:"#FCD34D",cursor:"pointer",fontFamily:"inherit"}}>↩ Undo delete ({deletedStack.length})</button>}
          <button onClick={()=>{setCalibMode(m=>!m);setCalibPts([]);}} style={{fontSize:"0.65rem",padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid "+(calibMode?"#EF4444":"#2D5280"),background:calibMode?"#7F1D1D":NAVY_LT,color:calibMode?"#FCA5A5":"#94A3B8",cursor:"pointer",fontFamily:"inherit"}}>📏 {calibMode?(calibPts.length<2?`Click point ${calibPts.length+1} of 2`:"2 points set"):"Calibrate scale"}</button>
          {calibFt&&!calibMode&&<div style={{fontSize:"0.62rem",padding:"0.3rem 0.6rem",borderRadius:20,background:"#064E3B",color:"#6EE7B7",border:"1px solid #065F46"}}>✓ Calibrated · {calibFt.toFixed(2)} ft/in<span onClick={()=>setCalibFt(null)} style={{cursor:"pointer",textDecoration:"underline",marginLeft:6}}>reset</span></div>}
          {calibMode&&calibPts.length===2&&<div style={{display:"flex",alignItems:"center",gap:"0.35rem",fontSize:"0.65rem",color:"#CBD5E1"}}>
            <span>Real distance (ft):</span>
            <input value={realDist} onChange={e=>setRealDist(e.target.value)} onKeyDown={e=>e.key==="Enter"&&applyCalibration()} placeholder="e.g. 20" style={{width:64,padding:"0.25rem 0.4rem",borderRadius:5,border:"1px solid #2D5280",background:NAVY,color:"#E2E8F0",fontSize:"0.65rem",fontFamily:"inherit"}}/>
            <button onClick={applyCalibration} style={{fontSize:"0.65rem",padding:"0.25rem 0.6rem",borderRadius:5,border:"none",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",color:"#fff",cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>Apply</button>
          </div>}
          {!calibMode&&<div style={{display:"flex",alignItems:"center",gap:"0.2rem"}}>
            <span style={{fontSize:"0.6rem",fontWeight:700,color:(elev?.scaleSource==="default"||(!elev?.verifiedScale&&!elev?.scale))&&!calibFt?"#FBBF24":"#64748B"}}>{(elev?.scaleSource==="default"||(!elev?.verifiedScale&&!elev?.scale))&&!calibFt?"⚠ set scale:":"scale:"}</span>
            {[["1/16",16],["3/32",32/3],["1/8",8],["3/16",16/3],["1/4",4],["3/8",8/3],["1/2",2],["1in",1]].map(([lab,v])=>{
              const on=Math.abs((effFtPerInch||0)-v)<0.02;
              return <button key={lab} onClick={()=>setCalibFt(on?null:v)} title={lab.replace("in","\"")+'"=1\'-0" — sets this page\'s scale'} style={{fontSize:"0.6rem",padding:"0.22rem 0.4rem",borderRadius:5,border:"1px solid "+(on?BLUE:"#2D5280"),background:on?BLUE:NAVY_LT,color:on?"#fff":"#94A3B8",cursor:"pointer",fontFamily:"inherit",fontWeight:on?700:500}}>{lab}</button>;
            })}
          </div>}
          <button onClick={()=>{setBucketMode(m=>!m);setSplitMode(false);setCornerMode(false);setCornerPts([]);setSnapMsg("");setCalibMode(false);}} style={{fontSize:"0.65rem",padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid "+(bucketMode?"#10B981":"#2D5280"),background:bucketMode?"#064E3B":NAVY_LT,color:bucketMode?"#6EE7B7":"#94A3B8",cursor:"pointer",fontFamily:"inherit"}}>🪣 {bucketMode?"Bucket ON — click a wall":"Bucket fill"}</button>
          <button onClick={()=>{setSplitMode(m=>!m);setBucketMode(false);setCornerMode(false);setCalibMode(false);setSnapMsg("");}} title="Click a face where the material changes — it splits along the structural line there" style={{fontSize:"0.65rem",padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid "+(splitMode?"#F59E0B":"#2D5280"),background:splitMode?"#451A03":NAVY_LT,color:splitMode?"#FCD34D":"#94A3B8",cursor:"pointer",fontFamily:"inherit"}}>✂ {splitMode?"Split ON — click the boundary":"Split face"}</button>
          {bucketMode&&<div style={{display:"flex",alignItems:"center",gap:"0.3rem",padding:"0.2rem 0.5rem",borderRadius:20,background:NAVY_LT,border:"1px solid #2D5280"}}>
            {BUCKET_COLS.map(c=>(
              <div key={c} onClick={()=>setCurBucketColor(c)} title={bucketColorNames[c]||"click to paint with this color"} style={{width:18,height:18,borderRadius:5,background:c,cursor:"pointer",border:curBucketColor===c?"2px solid #fff":"2px solid transparent",boxShadow:curBucketColor===c?"0 0 6px "+c:"none"}}/>
            ))}
            <input value={bucketColorNames[curBucketColor]||""} onChange={e=>setBucketColorNames({...bucketColorNames,[curBucketColor]:e.target.value})} placeholder="name this color (e.g. PNL-1)" style={{width:130,marginLeft:4,padding:"0.22rem 0.45rem",borderRadius:5,border:"1px solid #2D5280",background:NAVY,color:curBucketColor,fontWeight:700,fontSize:"0.64rem",fontFamily:"inherit"}}/>
          </div>}
          {bucketMode&&cornerMode&&<div style={{display:"flex",alignItems:"center",gap:"0.35rem"}}>
            <span style={{fontSize:"0.63rem",color:"#FCD34D"}}>Corners: {cornerPts.length}</span>
            <button onClick={finishCorners} disabled={cornerPts.length<3} style={{fontSize:"0.63rem",padding:"0.25rem 0.6rem",borderRadius:5,border:"none",background:cornerPts.length<3?"#334155":"linear-gradient(180deg,#34D399,#10B981)",color:"#fff",cursor:cornerPts.length<3?"default":"pointer",fontFamily:"inherit",fontWeight:700}}>Finish</button>
            <button onClick={cancelCorners} style={{fontSize:"0.63rem",padding:"0.25rem 0.5rem",borderRadius:5,border:"1px solid #2D5280",background:NAVY_LT,color:"#94A3B8",cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
          </div>}
          {(bucketMode||groupMode)&&snapMsg&&<div style={{fontSize:"0.63rem",padding:"0.3rem 0.6rem",borderRadius:20,background:NAVY_LT,color:(snapMsg[0]==="✓"||snapMsg[0]==="⚡")?"#6EE7B7":"#CBD5E1",border:"1px solid #2D5280"}}>{snapMsg}</div>}
          <button onClick={()=>{const nm=!groupMode;setGroupMode(nm);setBucketMode(false);setCalibMode(false);setCornerMode(false);if(nm&&previewGroups.length===0)loadGroups();}} style={{fontSize:"0.65rem",padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid "+(groupMode?"#A78BFA":"#2D5280"),background:groupMode?"#3730A3":NAVY_LT,color:groupMode?"#DDD6FE":"#94A3B8",cursor:"pointer",fontFamily:"inherit"}}>🎨 {groupMode?(groupBusy?"Finding groups…":"Auto-groups — click to pick"):"Auto-groups (preview)"}</button>
          {groupMode&&!groupBusy&&<div style={{fontSize:"0.62rem",padding:"0.3rem 0.6rem",borderRadius:20,background:NAVY_LT,color:"#C4B5FD",border:"1px solid #4C1D95"}}>Preview — pick the groups you're bidding {previewGroups.length>0?`(${previewGroups.length} found)`:""}</div>}
        </div>
        {!pageImage?<div style={{color:"#475569",fontSize:"0.8rem",marginTop:"4rem"}}>Loading elevation...</div>:
          <DeepZoom BACKEND={BACKEND} jobId={results.jobId} pageNum={pageNum}>
          <div style={{position:"relative",display:"inline-block",maxWidth:"100%"}}>
            <img ref={imgRef} src={pageImage} alt={elev?.title} onLoad={e=>{setImgNaturalSize({w:e.target.naturalWidth,h:e.target.naturalHeight});setImgLoaded(true);}} style={{display:"block",maxWidth:"100%",maxHeight:"calc(100vh - 180px)",objectFit:"contain",borderRadius:6,border:"1px solid "+NAVY_LT}}/>
            {imgLoaded&&<svg ref={svgRef} onClick={handleSvgClick} onMouseMove={handleSvgMove} onMouseUp={endDrag} onMouseLeave={endDrag} viewBox={`0 0 ${pageDims.width} ${pageDims.height}`} style={{position:"absolute",top:0,left:0,width:svgW,height:svgH,overflow:"visible",zIndex:2,cursor:(calibMode||bucketMode||splitMode)?"crosshair":"default"}}>
              {displayZones.map(zone=>{
                const a=getAssignment(zone.id);
                let color="#94A3B8";
                if(a)color=MAT_COLORS[a.category]||"#9CA3AF";
                else if(zone.cluster_id!==undefined)color=CLUSTER_COLORS[zone.cluster_id%CLUSTER_COLORS.length];
                else if(zone.fill_color?.length===3){const[r,g,b]=zone.fill_color;color=`rgb(${Math.round(r*255)},${Math.round(g*255)},${Math.round(b*255)})`;}
                const isSel=activeGroup&&gkey(zone)===activeGroup;
                const dimmed=activeGroup&&!isSel;
                const pts=toSVGPoints(zone.points);
                const lx=zone.cx*pageDims.width,ly=zone.cy*pageDims.height;
                const showLabel=!dimmed&&(zone.area_sf||0)>0;   // every surface carries its SF — same black-pill style as bucket fills
                const labelTxt=zone.source==="claude_vision"&&!a?zone.material_type:Math.round(zone.area_sf).toLocaleString()+" SF";
                const pillW=Math.max(64,(labelTxt||"").length*(pageDims.width/130));
                return <g key={zone.id} style={{cursor:(calibMode||bucketMode||splitMode)?"crosshair":"pointer",pointerEvents:groupMode?"none":"auto"}} onClick={e=>{if(calibMode||bucketMode||groupMode||splitMode)return;e.stopPropagation();const k=gkey(zone);setActiveGroup(activeGroup===k?null:k);}}>
                  {/* FLAT Bluebeam-style fill (the estimator's gold-standard look): solid color,
                      whisper outline — imperfect borders stop screaming, windows show as cutouts */}
                  <polygon points={pts} fill={color} fillOpacity={dimmed?0.07:isSel?0.66:0.5} stroke={isSel?"#fff":color} strokeWidth={isSel?2.5:1} strokeOpacity={dimmed?0.2:isSel?1:0.45}/>
                  {!dimmed&&(zone.holes||[]).map((hp,hi)=>(
                    <polygon key={"zh"+hi} points={toSVGPoints(hp)} fill="#FFFFFF" fillOpacity={0.85} stroke={color} strokeWidth={1} strokeOpacity={0.6} style={{pointerEvents:"none"}}/>
                  ))}
                  {showLabel&&<><rect x={lx-pillW/2} y={ly-10} width={pillW} height={20} fill="rgba(0,0,0,0.82)" rx={5}/><text x={lx} y={ly+2} textAnchor="middle" dominantBaseline="middle" fill={isSel?"#6EE7B7":"#FFFFFF"} fontSize={pageDims.width/75} fontFamily="Inter,Arial" fontWeight="bold">{labelTxt}</text></>}
                </g>;
              })}
              {calibPts.length===2&&<line x1={calibPts[0].x*pageDims.width} y1={calibPts[0].y*pageDims.height} x2={calibPts[1].x*pageDims.width} y2={calibPts[1].y*pageDims.height} stroke="#EF4444" strokeWidth={pageDims.width/350} strokeDasharray={pageDims.width/90}/>}
              {calibPts.map((p,i)=><circle key={"cp"+i} cx={p.x*pageDims.width} cy={p.y*pageDims.height} r={pageDims.width/110} fill="#EF4444" stroke="#fff" strokeWidth={pageDims.width/600}/>)}
              {pageBucketShapes.map(s=>{const cx=s.points.reduce((a,p)=>a+p[0],0)/s.points.length*pageDims.width,cy=s.points.reduce((a,p)=>a+p[1],0)/s.points.length*pageDims.height;const warn=s.openings_review;const col=s.color||(warn?"#F59E0B":"#10B981");return <g key={"bk"+s.id}>
                <polygon points={toSVGPoints(s.points)} fill={col} fillOpacity={editShape===s.id?0.6:0.5} stroke={editShape===s.id?"#fff":col} strokeWidth={editShape===s.id?2.6:1.4} onClick={s.sfScale?(e=>{e.stopPropagation();setEditShape(editShape===s.id?null:s.id);}):undefined} style={s.sfScale?{cursor:"pointer"}:undefined}/>
                {editShape===s.id&&s.points.map((pt,i)=>{const q=s.points[(i+1)%s.points.length];const mx=(pt[0]+q[0])/2*pageDims.width,my=(pt[1]+q[1])/2*pageDims.height,hs=pageDims.width/300;return <rect key={"m"+i} x={mx-hs} y={my-hs} width={hs*2} height={hs*2} fill="#93C5FD" stroke="#1D4ED8" strokeWidth={pageDims.width/900} style={{cursor:"copy"}} onMouseDown={e=>{e.stopPropagation();e.preventDefault();const mpt=[(pt[0]+q[0])/2,(pt[1]+q[1])/2];setBucketShapes(prev=>prev.map(sh=>sh.id!==s.id?sh:{...sh,points:[...sh.points.slice(0,i+1),mpt,...sh.points.slice(i+1)]}));dragRef.current={sid:s.id,vi:i+1};}}/>;})}
                {editShape===s.id&&s.points.map((pt,i)=><circle key={"v"+i} cx={pt[0]*pageDims.width} cy={pt[1]*pageDims.height} r={pageDims.width/150} fill="#fff" stroke="#10B981" strokeWidth={pageDims.width/500} style={{cursor:"grab"}} onMouseDown={e=>{e.stopPropagation();e.preventDefault();dragRef.current={sid:s.id,vi:i};}} onDoubleClick={e=>{e.stopPropagation();if(s.points.length<=3)return;setBucketShapes(prev=>prev.map(sh=>{if(sh.id!==s.id)return sh;const pts=sh.points.filter((_,ix)=>ix!==i);if(!sh.sfScale)return{...sh,points:pts};const gross=polyAreaSF(pts,sh.sfScale,pageDims.width,pageDims.height);return{...sh,points:pts,gross_sf:Math.round(gross*10)/10,area_sf:Math.round(Math.max(0,gross-(sh.opening_sf||0))*10)/10};}));}}/>)}
                {(s.holes||[]).map((hp,i)=><polygon key={"h"+i} points={toSVGPoints(hp)} fill="#FFFFFF" fillOpacity={0.75} stroke="#64748B" strokeWidth={1.2} strokeDasharray={pageDims.width/260}/>)}
                {(s.openings||[]).map((op,i)=>{const ox=op.reduce((a,p)=>a+p[0],0)/op.length*pageDims.width,oy=op.reduce((a,p)=>a+p[1],0)/op.length*pageDims.height;return <g key={i} onClick={e=>{e.stopPropagation();vetoOpening(s.id,i);}} style={{cursor:"pointer"}}>
                  <polygon points={toSVGPoints(op)} fill="#EF4444" fillOpacity={0.18} stroke="#EF4444" strokeWidth={1.6} strokeDasharray={pageDims.width/220}/>
                  <circle cx={ox} cy={oy} r={pageDims.width/140} fill="#EF4444"/><text x={ox} y={oy+pageDims.width/400} textAnchor="middle" dominantBaseline="middle" fill="#fff" fontSize={pageDims.width/110} fontFamily="Inter,Arial" fontWeight="bold">✕</text>
                </g>;})}
                <rect x={cx-52} y={cy-9} width={104} height={18} fill="rgba(0,0,0,0.82)" rx={4}/><text x={cx} y={cy+2} textAnchor="middle" dominantBaseline="middle" fill={warn?"#FCD34D":"#6EE7B7"} fontSize={pageDims.width/75} fontFamily="Inter,Arial" fontWeight="bold">{Math.round(s.area_sf).toLocaleString()} SF{s.opening_sf>0?" net":""}</text>
              </g>;})}
              {cornerMode&&cornerPts.length>0&&<polyline points={toSVGPoints(cornerPts.map(p=>[p.x,p.y]))} fill="none" stroke="#FCD34D" strokeWidth={pageDims.width/400} strokeDasharray={pageDims.width/120}/>}
              {cornerMode&&cornerPts.map((p,i)=><circle key={"kp"+i} cx={p.x*pageDims.width} cy={p.y*pageDims.height} r={pageDims.width/130} fill="#FCD34D" stroke="#fff" strokeWidth={pageDims.width/700}/>)}
              {groupMode&&previewGroups.map(g=>{const sel=selGroups[g.group]!==undefined;const col=`rgb(${g.color.map(c=>Math.round(c*255)).join(",")})`;return <g key={"pg"+g.group} onClick={e=>{e.stopPropagation();toggleGroup(g);}} style={{cursor:"pointer"}}>
                {g.patches.map((p,i)=><rect key={i} x={p[0]*pageDims.width} y={p[1]*pageDims.height} width={p[2]*pageDims.width} height={p[3]*pageDims.height} fill={col} fillOpacity={sel?0.6:0.25} stroke={sel?"#fff":"none"} strokeWidth={sel?pageDims.width/1500:0}/>)}
              </g>;})}
            </svg>}
          </div>
          </DeepZoom>
        }
      </div>
      {/* Right panel */}
      <div style={{width:230,borderLeft:"1px solid "+NAVY_LT,padding:"1rem",overflowY:"auto",flexShrink:0,background:NAVY_MID}}>
        {activeGroup?(
          <>
            <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>Selected Hatch</div>
            <div style={{padding:"0.65rem",marginBottom:"0.75rem",background:NAVY_LT,borderRadius:8,border:"1px solid "+BLUE+"55",textAlign:"center"}}>
              <div style={{fontSize:"1.6rem",fontWeight:800,color:"#fff",lineHeight:1.1}}>{Math.round(selectedSF).toLocaleString()} <span style={{fontSize:"0.7rem",fontWeight:400,color:"#94A3B8"}}>SF</span></div>
              <div style={{fontSize:"0.6rem",color:"#94A3B8",marginTop:3}}>{selectedZones.length} area{selectedZones.length!==1?"s":""} with this pattern</div>
              {(()=>{ /* TRUST = show the arithmetic: gross − openings = net, verifiable in seconds */
                const cs=selectedZones.map(z=>z.sf_calc).filter(Boolean);
                if(!cs.length) return null;
                const g=cs.reduce((s,c)=>s+(c.gross_sf||0),0), o=cs.reduce((s,c)=>s+(c.openings_sf||0),0), n=cs.reduce((s,c)=>s+(c.n_openings||0),0);
                return <div style={{fontSize:"0.6rem",color:"#7FB0E0",marginTop:5,padding:"0.3rem 0.4rem",background:"rgba(0,0,0,0.25)",borderRadius:5,fontVariantNumeric:"tabular-nums"}}>
                  {Math.round(g).toLocaleString()} gross − {Math.round(o).toLocaleString()} openings{n?` (${n})`:""} = <b style={{color:"#fff"}}>{Math.round(g-o).toLocaleString()} SF</b>
                  <div style={{color:"#64748B",marginTop:1}}>{cs[0].basis} · measured from the drawing's own coordinates</div>
                </div>;
              })()}
              {(()=>{ /* TRUST: which reader measured this — provenance, not mystery */
                const readerOf=(z)=>{const m=((z.group||z.material)||"")+"";
                  if(z.source==="bluebeam"||z.source==="bluebeam-linear") return "Your Bluebeam markup (exact)";
                  if(m.startsWith("Hatched area")) return "Hatch-pattern reader";
                  if(m.startsWith("Color fill")) return "Drawn color fills";
                  if(m.startsWith("Rendered")) return "Rendered-elevation reader";
                  if(m.startsWith("Wall area (AI boundary")) return "AI boundary model (v13)";
                  if(m.startsWith("Wall band")) return "Story-band reader";
                  if(m.startsWith("Wall area")) return "Structural flood fill";
                  if(m.startsWith("Panel wall")) return "Drawn wall fills";
                  if(/vertical panel|horizontal panel|lap \/|courses|seams/i.test(m)) return "Seam-pattern reader";
                  return z.source==="model"?"Trained extent model":"Drawing geometry";
                };
                const rs=[...new Set(selectedZones.map(readerOf))];
                if(!rs.length) return null;
                return <div style={{fontSize:"0.58rem",color:"#94A3B8",marginTop:4}}>
                  📖 Read by: <b style={{color:"#CBD5E1"}}>{rs.join(" + ")}</b>
                </div>;
              })()}
            </div>
            <div style={{marginBottom:"0.7rem"}}>
              <div style={{fontSize:"0.6rem",color:"#64748B",marginBottom:"0.3rem"}}>Name this group → applies to <b style={{color:"#94A3B8"}}>every page</b> + the Excel:</div>
              <div style={{display:"flex",gap:"0.35rem"}}>
                <input value={groupNameDraft} onChange={e=>setGroupNameDraft(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")renameGroup(groupNameDraft);}} placeholder="e.g. Metal Panel" style={{flex:1,minWidth:0,background:NAVY,border:"1px solid #2D5280",borderRadius:6,color:"#fff",fontSize:"0.7rem",padding:"0.4rem 0.5rem",fontFamily:"inherit"}}/>
                <button onClick={()=>renameGroup(groupNameDraft)} style={{background:"linear-gradient(180deg,#5A92D2,#3F79BC)",border:"none",borderRadius:6,color:"#fff",fontSize:"0.66rem",fontWeight:700,padding:"0 0.65rem",cursor:"pointer",fontFamily:"inherit"}}>Set</button>
              </div>
            </div>
            {learnedMat&&<div onClick={()=>assignGroup({id:learnedMat.id,name:learnedMat.materialName,category:learnedMat.category})} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.5rem 0.65rem",cursor:"pointer",background:"#064E3B",borderRadius:6,border:"1px solid #10B981",marginBottom:"0.6rem"}}>
              <span style={{fontSize:"0.85rem"}}>✨</span>
              <div style={{fontSize:"0.63rem",color:"#A7F3D0",flex:1}}>From memory: this hatch was <b>{learnedMat.materialName}</b> — tap to apply</div>
            </div>}
            <div style={{fontSize:"0.6rem",color:"#64748B",marginBottom:"0.4rem"}}>Tag this hatch as (optional):</div>
            {matList.map((mat,i)=>(
              <div key={i} onClick={()=>assignGroup(mat)} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.5rem 0.65rem",cursor:"pointer",background:NAVY,borderRadius:6,border:"1px solid #2D5280",marginBottom:"0.35rem"}} onMouseEnter={e=>e.currentTarget.style.borderColor=BLUE} onMouseLeave={e=>e.currentTarget.style.borderColor="#2D5280"}>
                <div style={{width:10,height:10,borderRadius:3,background:MAT_COLORS[mat.category]||"#9CA3AF",flexShrink:0}}/>
                <div style={{fontSize:"0.65rem",color:"#CBD5E1",flex:1}}><span style={{color:BLUE,fontWeight:700}}>{mat.id}</span> {mat.name||mat.category}</div>
              </div>
            ))}
            {selectedIds.some(id=>getAssignment(id))&&<div onClick={removeGroup} style={{padding:"0.45rem",marginTop:"0.4rem",textAlign:"center",fontSize:"0.65rem",color:"#F87171",cursor:"pointer",border:"1px solid #7F1D1D",borderRadius:6}}>Remove tag from this hatch</div>}
            <div onClick={deleteGroup} style={{padding:"0.5rem",marginTop:"0.4rem",textAlign:"center",fontSize:"0.68rem",fontWeight:700,color:"#fff",cursor:"pointer",background:"linear-gradient(180deg,#EF4444,#C0392B)",borderRadius:7}}>🗑 Delete this highlight ({selectedZones.length} area{selectedZones.length!==1?"s":""} · −{Math.round(selectedSF).toLocaleString()} SF)</div>
            <div style={{fontSize:"0.56rem",color:"#64748B",marginTop:"0.3rem",lineHeight:1.4,textAlign:"center"}}>Wrong detection? Deleting removes it from the takeoff, the totals and the Excel.</div>
            <div onClick={()=>setActiveGroup(null)} style={{padding:"0.45rem",marginTop:"0.3rem",textAlign:"center",fontSize:"0.65rem",color:"#64748B",cursor:"pointer",border:"1px solid #2D5280",borderRadius:6}}>Deselect</div>
          </>
        ):(
          <>
            {Object.keys(clusterSummary).length>0&&<>
              <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.4rem"}}>Hatches Detected</div>
              <div style={{fontSize:"0.62rem",color:"#475569",marginBottom:"0.65rem"}}>Click a hatch (or any area on the drawing) → all matching areas select &amp; show their SF. Skip brick.</div>
              {Object.entries(clusterSummary).map(([cid,info])=>{
                const zids=colorGroups["c_"+cid]||[];
                const assigned=zids.filter(id=>getAssignment(id)).length;
                return <div key={cid} onClick={()=>setActiveGroup("c_"+cid)} style={{marginBottom:"0.4rem",padding:"0.5rem 0.65rem",background:NAVY,borderRadius:6,borderLeft:"3px solid "+info.color,cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background=NAVY_LT} onMouseLeave={e=>e.currentTarget.style.background=NAVY}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"0.4rem"}}><div style={{width:8,height:8,borderRadius:2,background:info.color}}/><span style={{fontSize:"0.65rem",color:"#94A3B8"}}>Hatch {parseInt(cid)+1}</span></div>
                    <span style={{fontSize:"0.8rem",fontWeight:700,color:"#E2E8F0"}}>{Math.round(info.total_sf).toLocaleString()} SF</span>
                  </div>
                  <div style={{fontSize:"0.6rem",color:assigned>0?"#4ADE80":"#475569",marginTop:3}}>{assigned>0?`✓ ${assigned}/${info.count} tagged`:`${info.count} area${info.count!==1?"s":""} · click to select`}</div>
                </div>;
              })}
              <div style={{height:1,background:NAVY_LT,margin:"0.75rem 0"}}/>
            </>}
            {displayZones.length>0&&lf.total>0&&<>
              <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.4rem"}}>Linear (estimated)</div>
              <div style={{fontSize:"0.58rem",color:"#475569",marginBottom:"0.5rem"}}>From surface edges{calibFt?" · calibrated scale":""} — refine on site</div>
              {[["Coping / top",lf.top],["Base / starter",lf.bottom],["Corners / vert. trim",lf.vert]].map(([label,val])=>(
                <div key={label} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"0.35rem 0.6rem",background:NAVY,borderRadius:6,marginBottom:"0.3rem"}}>
                  <span style={{fontSize:"0.62rem",color:"#94A3B8"}}>{label}</span>
                  <span style={{fontSize:"0.72rem",fontWeight:700,color:"#E2E8F0"}}>{val.toLocaleString()} LF</span>
                </div>
              ))}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"0.4rem 0.6rem",background:NAVY_LT,borderRadius:6}}>
                <span style={{fontSize:"0.62rem",color:"#CBD5E1",fontWeight:700}}>Total perimeter</span>
                <span style={{fontSize:"0.75rem",fontWeight:800,color:BLUE}}>{lf.total.toLocaleString()} LF</span>
              </div>
              <div style={{height:1,background:NAVY_LT,margin:"0.75rem 0"}}/>
            </>}
            {Object.keys(totals).length>0&&<>
              <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>Your Takeoff</div>
              {Object.entries(totals).map(([cat,sf])=><div key={cat} style={{marginBottom:"0.5rem",padding:"0.5rem 0.65rem",background:NAVY,borderRadius:6,borderLeft:"3px solid "+(MAT_COLORS[cat]||"#9CA3AF")}}>
                <div style={{fontSize:"0.62rem",color:"#94A3B8"}}>{cat}</div>
                <div style={{fontSize:"1.1rem",fontWeight:700,color:"#E2E8F0"}}>{Math.round(sf).toLocaleString()} <span style={{fontSize:"0.65rem",fontWeight:400}}>SF net</span></div>
                <div style={{fontSize:"0.6rem",color:"#64748B"}}>{Math.round(sf*1.15).toLocaleString()} SF +15%</div>
              </div>)}
              <div style={{padding:"0.65rem",background:NAVY_LT,borderRadius:6,border:"1px solid "+BLUE+"40",marginBottom:"0.75rem"}}>
                <div style={{fontSize:"0.6rem",color:"#4ADE80",fontWeight:700}}>GRAND TOTAL</div>
                <div style={{fontSize:"1.4rem",fontWeight:700,color:"#4ADE80"}}>{Math.round(grandTotal*1.15).toLocaleString()} <span style={{fontSize:"0.65rem",fontWeight:400}}>SF</span></div>
                <div style={{fontSize:"0.6rem",color:"#475569"}}>+15% waste factor</div>
              </div>
              <button onClick={exportInteractiveExcel} style={{width:"100%",padding:"0.65rem",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",color:"#fff",border:"none",borderRadius:7,fontSize:"0.72rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓ Export Excel</button>
            </>}
            {selGroupList.length>0&&<div style={{marginTop:"0.5rem",padding:"0.6rem",background:NAVY,borderRadius:7,border:"1px solid #8B5CF655"}}>
              <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#C4B5FD",textTransform:"uppercase",fontWeight:700,marginBottom:"0.4rem"}}>🎨 Selected groups</div>
              {selGroupList.map(g=><div key={g.group} style={{display:"flex",alignItems:"center",gap:"0.35rem",marginBottom:"0.3rem"}}>
                <span style={{width:12,height:12,borderRadius:3,background:`rgb(${g.color.map(c=>Math.round(c*255)).join(",")})`,flexShrink:0}}/>
                <span style={{fontSize:"0.68rem",fontWeight:700,color:"#E2E8F0",minWidth:52}}>~{Math.round(g.approx_sf).toLocaleString()}</span>
                <input value={selGroups[g.group]} onChange={e=>setGroupMat(g.group,e.target.value)} placeholder="material…" style={{flex:1,minWidth:0,padding:"0.2rem 0.4rem",borderRadius:5,border:"1px solid #2D5280",background:NAVY_LT,color:"#E2E8F0",fontSize:"0.6rem",fontFamily:"inherit"}}/>
                <button onClick={()=>refineGroup(g)} disabled={groupBusy} title="Snap this group to the drawing's lines → exact SF" style={{padding:"0.18rem 0.42rem",borderRadius:5,border:"none",background:groupBusy?"#334155":"linear-gradient(180deg,#FBBF24,#F59E0B)",color:"#1F2937",fontSize:"0.62rem",fontWeight:800,cursor:groupBusy?"default":"pointer",fontFamily:"inherit"}}>⚡</button>
              </div>)}
              <div style={{display:"flex",justifyContent:"space-between",padding:"0.3rem 0",fontSize:"0.68rem",color:"#C4B5FD",fontWeight:700}}><span>Preview total</span><span>~{Math.round(selGroupSF).toLocaleString()} SF</span></div>
              <div style={{fontSize:"0.56rem",color:"#F59E0B",marginBottom:"0.35rem",lineHeight:1.4}}>⚠ Preview estimate — hit ⚡ on a group to snap it to the drawing for exact SF</div>
              <button onClick={exportInteractiveExcel} style={{width:"100%",padding:"0.45rem",background:"linear-gradient(180deg,#A78BFA,#8B5CF6)",color:"#fff",border:"none",borderRadius:6,fontSize:"0.66rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓ Export (incl. preview)</button>
            </div>}
            {bucketShapes.length>0&&<div style={{marginTop:"0.5rem",padding:"0.6rem",background:NAVY,borderRadius:7,border:"1px solid #10B98155"}}>
              <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#6EE7B7",textTransform:"uppercase",fontWeight:700,marginBottom:"0.15rem"}}>🪣 Measured shapes</div>
              <div style={{fontSize:"0.55rem",color:"#64748B",marginBottom:"0.5rem"}}>click a shape → drag corners (SF follows) · grab a blue midpoint to add a corner · double-click a corner to remove it</div>
              {bucketShapes.map(s=><div key={s.id} style={{display:"flex",alignItems:"center",gap:"0.35rem",marginBottom:"0.3rem"}}>
                <span style={{fontSize:"0.7rem",fontWeight:700,color:s.openings_review?"#FCD34D":"#E2E8F0",minWidth:60}} title={s.opening_sf>0?`net of ${Math.round(s.opening_sf).toLocaleString()} SF openings — click ✕ on the drawing to undo any`:""}>{Math.round(s.area_sf).toLocaleString()} SF{s.opening_sf>0?"*":""}{s.openings_review?" ⚠":""}</span>
                <input value={s.material} onChange={e=>setBucketMaterial(s.id,e.target.value)} onBlur={()=>propagateMaterial(s.id)} onKeyDown={e=>{if(e.key==="Enter")propagateMaterial(s.id);}} placeholder="material…" title="naming this also names unnamed shapes with the same pattern" style={{flex:1,minWidth:0,padding:"0.2rem 0.4rem",borderRadius:5,border:"1px solid #2D5280",background:NAVY_LT,color:"#E2E8F0",fontSize:"0.62rem",fontFamily:"inherit"}}/>
                {s.page!==pageNum&&<span style={{fontSize:"0.55rem",color:"#475569"}}>p{s.page}</span>}
                <span onClick={()=>removeBucketShape(s.id)} title="remove" style={{cursor:"pointer",color:"#F87171",fontSize:"0.85rem",fontWeight:700}}>×</span>
              </div>)}
              <div style={{display:"flex",justifyContent:"space-between",padding:"0.35rem 0",fontSize:"0.7rem",color:"#6EE7B7",fontWeight:700}}><span>Bucket total</span><span>{Math.round(bucketShapes.reduce((a,s)=>a+(s.area_sf||0),0)).toLocaleString()} SF</span></div>
              <button onClick={exportInteractiveExcel} style={{width:"100%",padding:"0.5rem",marginTop:"0.3rem",background:"linear-gradient(180deg,#34D399,#10B981)",color:"#fff",border:"none",borderRadius:6,fontSize:"0.68rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓ Export (incl. bucket)</button>
            </div>}
            {Object.keys(clusterSummary).length===0&&Object.keys(totals).length===0&&bucketShapes.length===0&&<div style={{fontSize:"0.7rem",color:"#475569",lineHeight:1.8}}>Click any colored area → every area with that same hatch selects and shows its total SF.<br/><br/>Or hit <b style={{color:"#6EE7B7"}}>🪣 Bucket fill</b> and click a wall to measure it straight from the drawing.</div>}
            <div style={{marginTop:"1rem",fontSize:"0.6rem",color:"#334155"}}>{Object.keys(assignments).length} areas tagged · {elevations.length} elevations</div>
          </>
        )}
      </div>
    </div>
  );
}

/* ── Konva deep-zoom hook: wheel zooms to cursor (1–14×), drag pans when zoomed, past 2.2×
     swaps in a hi-res render of the viewport. localPos() gives zoom-corrected coordinates. ── */
function useKonvaZoom(BACKEND, jobId, pageNum, stageW, stageH) {
  const [view, setView] = useState({ k: 1, x: 0, y: 0 });
  const [hi, setHi] = useState(null);
  const timer = useRef();
  useEffect(() => { setView({ k: 1, x: 0, y: 0 }); setHi(null); }, [jobId, pageNum]);
  useEffect(() => () => clearTimeout(timer.current), []);
  const sched = (k, x, y) => {
    clearTimeout(timer.current);
    timer.current = setTimeout(() => {
      if (k < 2.2 || !jobId || !stageW || !stageH) { setHi(null); return; }
      const nx0 = Math.max(0, (-x) / k / stageW), ny0 = Math.max(0, (-y) / k / stageH);
      const nx1 = Math.min(1, ((-x) / k + stageW / k) / stageW), ny1 = Math.min(1, ((-y) / k + stageH / k) / stageH);
      const url = `${BACKEND}/page-crop/${jobId}/${pageNum}?x0=${nx0.toFixed(4)}&y0=${ny0.toFixed(4)}&x1=${nx1.toFixed(4)}&y1=${ny1.toFixed(4)}&px=2000`;
      const im = new window.Image(); im.crossOrigin = "anonymous";
      im.onload = () => setHi({ im, x: nx0 * stageW, y: ny0 * stageH, w: (nx1 - nx0) * stageW, h: (ny1 - ny0) * stageH });
      im.src = url;
    }, 350);
  };
  const onWheel = e => {
    e.evt.preventDefault();
    const st = e.target.getStage(); const p = st.getPointerPosition(); if (!p) return;
    const ok = view.k;
    const nk = Math.max(1, Math.min(14, ok * (e.evt.deltaY < 0 ? 1.2 : 1 / 1.2)));
    const mx = (p.x - view.x) / ok, my = (p.y - view.y) / ok;
    const nv = nk === 1 ? { k: 1, x: 0, y: 0 } : { k: nk, x: p.x - mx * nk, y: p.y - my * nk };
    setView(nv); sched(nv.k, nv.x, nv.y);
  };
  const onDragEnd = e => {
    const st = e.target;
    if (st.getStage && st === st.getStage()) {
      setView(v => { const nv = { ...v, x: st.x(), y: st.y() }; sched(nv.k, nv.x, nv.y); return nv; });
    }
  };
  const stageProps = { scaleX: view.k, scaleY: view.k, x: view.x, y: view.y, draggable: view.k > 1, dragDistance: 4, onWheel, onDragEnd };
  const hiResNode = hi ? <KImage image={hi.im} x={hi.x} y={hi.y} width={hi.w} height={hi.h} listening={false} /> : null;
  const localPos = st => { const p = st.getPointerPosition(); if (!p) return null; return { x: (p.x - view.x) / view.k, y: (p.y - view.y) / view.k }; };
  return { view, stageProps, hiResNode, localPos };
}

/* ── Konva polygon editor: drag vertices to correct the AI, live SF ── */
function EditorView({ results, BACKEND, setResults }) {
  const elevations = results.takeoffData.filter(e=>e.pageNumber);
  const [elevIdx,setElevIdx]=useState(0);
  const elev=elevations[elevIdx];
  const pageNum=elev?.pageNumber;
  const [polys,setPolys]=useState([]);
  const [pageDims,setPageDims]=useState({width:612,height:792});
  const [img,setImg]=useState(null);
  const [selId,setSelId]=useState(null);
  const [stageW,setStageW]=useState(800);
  const [mode,setMode]=useState("select");   // "select" | "cut" — cut = draw a window/door the AI missed, gets subtracted
  const [draft,setDraft]=useState([]);
  const wrapRef=useRef();

  useEffect(()=>{
    if(!pageNum||!results.jobId) return;
    setPolys([]); setSelId(null); setImg(null); setDraft([]); setMode("select");
    fetch(BACKEND+"/polygons/"+results.jobId+"/"+pageNum).then(r=>r.ok?r.json():{polygons:[]}).then(d=>{
      if(d.width) setPageDims({width:d.width,height:d.height||792});
      setPolys((d.polygons||[]).map((p,i)=>({id:i,points:p.points||[],category:p.material_type||p.category||"Other",area_sf:p.area_sf||0})));
    }).catch(()=>{});
    const im=new window.Image();
    im.src=BACKEND+"/page-image/"+results.jobId+"/"+pageNum;
    im.onload=()=>setImg(im);
  },[elevIdx,pageNum,results.jobId,BACKEND]);

  useEffect(()=>{
    const update=()=>{ if(wrapRef.current) setStageW(Math.max(320, wrapRef.current.offsetWidth-32)); };
    update(); window.addEventListener("resize",update); return ()=>window.removeEventListener("resize",update);
  },[]);

  const sc = stageW/pageDims.width;
  const stageH = pageDims.height*sc;
  const zoom = useKonvaZoom(BACKEND, results.jobId, pageNum, stageW, stageH);
  const effFt = (()=>{ for(const p of polys){ if(p.area_sf>0&&p.points.length>=3){ const sh=polyAreaSF(p.points,72,pageDims.width,pageDims.height); if(sh>0) return 72*Math.sqrt(p.area_sf/sh);}} return 8; })();
  const areaOf=p=>polyAreaSF(p.points,effFt,pageDims.width,pageDims.height);
  const totalSF=polys.reduce((s,p)=>s+(p.type==="cut"?-areaOf(p):areaOf(p)),0);
  const updateVertex=(pid,vi,nx,ny)=>setPolys(prev=>prev.map(p=>p.id!==pid?p:{...p,points:p.points.map((pt,i)=>i===vi?[nx,ny]:pt)}));
  const deletePoly=pid=>{
    const p=polys.find(x=>x.id===pid);
    if(p&&p.type==="cut"&&p.applied&&setResults){   // undoing an applied cut-out puts its SF back
      const sf=areaOf(p);
      const host=polys.find(q=>q.type!=="cut"&&p.points[0]&&pip(p.points[0],q.points));
      const hostCat=host?host.category:null;
      setResults(prev=>({...prev,takeoffData:prev.takeoffData.map(e=>{
        if(e.pageNumber!==pageNum) return e;
        const zones=(e.zones||[]).map(z=>({...z}));
        let z=hostCat?zones.find(zz=>zz.category===hostCat||zz.materialName===hostCat):null;
        if(!z&&zones.length) z=zones.reduce((a,b)=>(a.netArea||0)>=(b.netArea||0)?a:b);
        if(z){ z.netArea=(z.netArea||0)+sf; z.totalOpeningArea=Math.max(0,(z.totalOpeningArea||0)-sf); }
        return {...e,zones};
      })}));
    }
    setPolys(prev=>prev.filter(x=>x.id!==pid)); setSelId(null);
  };
  const finishCut=()=>{
    if(draft.length<3) return;
    setPolys(prev=>[...prev,{id:Date.now(),points:draft,category:"Cut-out (opening)",area_sf:0,type:"cut"}]);
    setDraft([]);
  };
  // ray-cast point-in-polygon on normalized coords — finds which surface a cut-out sits inside
  const pip=(pt,poly)=>{
    let inside=false;
    for(let i=0,j=poly.length-1;i<poly.length;j=i++){
      const [xi,yi]=poly[i],[xj,yj]=poly[j];
      if(((yi>pt[1])!==(yj>pt[1]))&&(pt[0]<(xj-xi)*(pt[1]-yi)/(yj-yi)+xi)) inside=true;
    }
    return inside;
  };
  const pendingCuts=polys.filter(p=>p.type==="cut"&&!p.applied);
  // Subtract the drawn cut-outs from this elevation's REAL takeoff numbers (summary, pricing, Excel all follow)
  const applyCuts=()=>{
    if(!pendingCuts.length||!setResults) return;
    const cuts=pendingCuts.map(c=>{
      const host=polys.find(p=>p.type!=="cut"&&c.points[0]&&pip(c.points[0],p.points));
      return {sf:areaOf(c), hostCat:host?host.category:null};
    });
    setResults(prev=>({...prev,takeoffData:prev.takeoffData.map(e=>{
      if(e.pageNumber!==pageNum) return e;
      const zones=(e.zones||[]).map(z=>({...z}));
      cuts.forEach(({sf,hostCat})=>{
        let z=hostCat?zones.find(zz=>zz.category===hostCat||zz.materialName===hostCat):null;
        if(!z&&zones.length) z=zones.reduce((a,b)=>(a.netArea||0)>=(b.netArea||0)?a:b);   // fall back to the biggest surface on the sheet
        if(z){ z.netArea=Math.max(0,(z.netArea||0)-sf); z.totalOpeningArea=(z.totalOpeningArea||0)+sf; }
      });
      return {...e,zones};
    })}));
    setPolys(prev=>prev.map(p=>p.type==="cut"&&!p.applied?{...p,applied:true}:p));
  };

  return (
    <div style={{display:"flex",height:"100%",overflow:"hidden",background:NAVY,fontFamily:"'Inter','Segoe UI',sans-serif"}}>
      <div style={{width:170,borderRight:"1px solid "+NAVY_LT,overflowY:"auto",flexShrink:0,background:NAVY_MID}}>
        <div style={{padding:"0.8rem",fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,borderBottom:"1px solid "+NAVY_LT}}>Elevations</div>
        {elevations.map((e,i)=><div key={i} onClick={()=>setElevIdx(i)} style={{padding:"0.6rem 0.8rem",cursor:"pointer",borderBottom:"1px solid "+NAVY_LT,background:i===elevIdx?NAVY_LT:"transparent",fontSize:"0.7rem",color:i===elevIdx?"#E2E8F0":"#94A3B8",borderLeft:i===elevIdx?"3px solid "+BLUE:"3px solid transparent"}}>{e.title||"Page "+e.pageNumber}</div>)}
      </div>
      <div ref={wrapRef} style={{flex:1,overflow:"auto",padding:"1rem"}}>
        <div style={{fontSize:"0.62rem",color:"#94A3B8",marginBottom:"0.5rem"}}>{mode==="cut"?"Cut-out mode: click around the window/door the AI missed, then Finish — it subtracts from the SF.":"Click a shape to select · drag its dots to correct it · SF updates live"}<span style={{color:"#64748B"}}> · wheel = zoom · drag = pan when zoomed</span></div>
        {img?<Stage width={stageW} height={stageH} {...zoom.stageProps} onClick={e=>{
            if(mode==="cut"){ const pos=zoom.localPos(e.target.getStage()); if(pos) setDraft(prev=>[...prev,[pos.x/(pageDims.width*sc),pos.y/(pageDims.height*sc)]]); return; }
            if(e.target===e.target.getStage()) setSelId(null);
          }}>
          <Layer>
            <KImage image={img} width={stageW} height={stageH}/>
            {zoom.hiResNode}
            {polys.map(p=>{
              const flat=p.points.flatMap(([nx,ny])=>[nx*pageDims.width*sc, ny*pageDims.height*sc]);
              const isCut=p.type==="cut";
              const col=isCut?"#EF4444":CLUSTER_COLORS[p.id%CLUSTER_COLORS.length];
              const sel=selId===p.id;
              return <Line key={p.id} points={flat} closed fill={col+(isCut?(p.applied?"22":"44"):sel?"66":"33")} stroke={sel?"#ffffff":col} strokeWidth={sel?2:isCut?1.5:1} dash={isCut?[7,4]:undefined} onClick={()=>setSelId(p.id)} onTap={()=>setSelId(p.id)}/>;
            })}
            {draft.length>0&&<Line points={draft.flatMap(([nx,ny])=>[nx*pageDims.width*sc,ny*pageDims.height*sc])} stroke="#F59E0B" strokeWidth={2} dash={[6,4]}/>}
            {draft.map(([nx,ny],i)=><Circle key={"d"+i} x={nx*pageDims.width*sc} y={ny*pageDims.height*sc} radius={4} fill="#F59E0B"/>)}
            {selId!==null&&(polys.find(p=>p.id===selId)?.points||[]).map(([nx,ny],vi)=>(
              <Circle key={vi} x={nx*pageDims.width*sc} y={ny*pageDims.height*sc} radius={5} fill="#ffffff" stroke="#3B82F6" strokeWidth={2} draggable
                onDragMove={e=>updateVertex(selId, vi, e.target.x()/(pageDims.width*sc), e.target.y()/(pageDims.height*sc))}/>
            ))}
          </Layer>
        </Stage>:<div style={{color:"#475569",fontSize:"0.8rem",marginTop:"3rem"}}>{elevations.length?"Loading elevation…":"This takeoff found no measurable pages — nothing to edit here. Re-run the drawing in the Takeoff tab."}</div>}
      </div>
      <div style={{width:220,borderLeft:"1px solid "+NAVY_LT,padding:"1rem",background:NAVY_MID,overflowY:"auto",flexShrink:0}}>
        <div style={{display:"flex",gap:"0.35rem",marginBottom:"0.65rem"}}>
          {[["select","Select / edit"],["cut","– Cut-out"]].map(([m,label])=>(
            <button key={m} onClick={()=>{setMode(m);setDraft([]);setSelId(null);}} style={{flex:1,padding:"0.5rem",borderRadius:7,border:"1px solid "+(mode===m?"transparent":NAVY_LT),background:mode===m?(m==="cut"?"linear-gradient(180deg,#EF4444,#C0392B)":"linear-gradient(180deg,#5A92D2,#3F79BC)"):"transparent",color:mode===m?"#fff":"#94A3B8",fontSize:"0.64rem",fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{label}</button>
          ))}
        </div>
        {mode==="cut"&&draft.length>0&&<div style={{display:"flex",gap:"0.35rem",marginBottom:"0.65rem"}}>
          <button onClick={finishCut} style={{flex:1,padding:"0.55rem",background:"linear-gradient(180deg,#22A860,#16874B)",border:"none",borderRadius:7,color:"#fff",fontSize:"0.66rem",fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✓ Finish cut-out ({draft.length} pts)</button>
          <button onClick={()=>setDraft([])} style={{padding:"0.55rem 0.7rem",background:"transparent",border:"1px solid "+NAVY_LT,borderRadius:7,color:"#94A3B8",fontSize:"0.66rem",fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✕</button>
        </div>}
        {pendingCuts.length>0&&<button onClick={applyCuts} style={{width:"100%",padding:"0.6rem",marginBottom:"0.65rem",background:"linear-gradient(180deg,#F59E0B,#D97706)",border:"none",borderRadius:8,color:"#fff",fontSize:"0.68rem",fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Apply {pendingCuts.length} cut-out{pendingCuts.length!==1?"s":""} to the takeoff (−{Math.round(pendingCuts.reduce((s,c)=>s+areaOf(c),0)).toLocaleString()} SF)</button>}
        <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>Surfaces ({polys.length})</div>
        {polys.map(p=>{
          const sel=selId===p.id;
          const isCut=p.type==="cut";
          return <div key={p.id} onClick={()=>setSelId(p.id)} style={{padding:"0.5rem 0.65rem",marginBottom:"0.35rem",background:sel?NAVY_LT:NAVY,borderRadius:6,cursor:"pointer",borderLeft:"3px solid "+(isCut?"#EF4444":CLUSTER_COLORS[p.id%CLUSTER_COLORS.length])}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontSize:"0.65rem",color:isCut?"#FCA5A5":"#CBD5E1"}}>{p.category}{isCut&&p.applied?" ✓":""}</span>
              <span style={{fontSize:"0.72rem",fontWeight:700,color:isCut?"#F87171":"#E2E8F0"}}>{isCut?"−":""}{Math.round(areaOf(p)).toLocaleString()} SF</span>
            </div>
            {sel&&<div onClick={ev=>{ev.stopPropagation();deletePoly(p.id);}} style={{marginTop:"0.4rem",fontSize:"0.6rem",color:"#F87171",textAlign:"center",border:"1px solid #7F1D1D",borderRadius:5,padding:"0.3rem",cursor:"pointer"}}>Delete surface</div>}
          </div>;
        })}
        <div style={{marginTop:"0.75rem",padding:"0.6rem 0.75rem",background:NAVY_LT,borderRadius:8,border:"1px solid "+BLUE+"40"}}>
          <div style={{fontSize:"0.6rem",color:"#4ADE80",fontWeight:700}}>EDITED TOTAL</div>
          <div style={{fontSize:"1.3rem",fontWeight:800,color:"#4ADE80"}}>{Math.round(totalSF).toLocaleString()} <span style={{fontSize:"0.65rem",fontWeight:400}}>SF</span></div>
        </div>
      </div>
    </div>
  );
}

/* ── Drawing-intelligence chip (verified scale, dims, schedule openings) ── */
const fmtDims = bd => {
  if(!bd) return null;
  const w=bd.overall_width_ft, h=bd.overall_height_ft;
  if(w&&h) return `${w}′ W × ${h}′ H`;
  if(w) return `${w}′ wide`;
  if(h) return `${h}′ tall`;
  return null;
};
function InfoChip({ label, value, sub, warn }) {
  return <div style={{display:"flex",flexDirection:"column",padding:"0.3rem 0.6rem",background:"#fff",borderRadius:6,border:"1px solid "+(warn?"#FCD34D":"#DBEAFE")}}>
    <span style={{fontSize:"0.52rem",letterSpacing:"0.07em",color:warn?"#B45309":"#64748B",textTransform:"uppercase",fontWeight:700}}>{label}</span>
    <span style={{fontSize:"0.72rem",fontWeight:700,color:warn?"#92400E":"#0F172A",lineHeight:1.3}}>{value}</span>
    {sub&&<span style={{fontSize:"0.55rem",color:warn?"#B45309":"#94A3B8"}}>{sub}</span>}
  </div>;
}

/* Per-elevation trust score from signals we already capture → triage "ready" vs "review" */
const elevConfidence = (elev) => {
  const total = (elev.zones||[]).reduce((s,z)=>s+(z.netArea||0),0);
  let score = 70; const reasons = [];
  if (elev.verifiedScale || elev.scaleSource==="claude_vision") score += 15;
  else if (elev.scaleSource==="easyocr") score += 5;
  else if (elev.scaleSource==="default" || (!elev.verifiedScale && !elev.scale)) { score -= 30; reasons.push("scale not confirmed — calibrate to be sure"); }
  if (elev.expectedFacadeSF) {
    const ratio = total / elev.expectedFacadeSF;
    if (ratio > 1.4) { score -= 35; reasons.push("measured area exceeds the building face — likely scale error"); }
    else if (ratio >= 0.12 && ratio <= 1.05) score += 10;
  } else { score -= 5; reasons.push("no printed dimensions to cross-check against"); }
  if (elev.scheduleOpeningSF > 0) score += 5; else reasons.push("openings estimated (no schedule found)");
  score = Math.max(5, Math.min(99, score));
  const status = score >= 85 ? "ready" : score >= 65 ? "review" : "attention";
  return { score, status, reasons };
};
const STATUS_STYLE = {
  ready:     { bg:"#DCFCE7", fg:"#15803D", label:"✓ Ready" },
  review:    { bg:"#FEF3C7", fg:"#B45309", label:"Review" },
  attention: { bg:"#FEE2E2", fg:"#B91C1C", label:"⚠ Check" },
};

function ScopeSection({ title, tone, children }) {
  return <div style={{background:"#fff",borderRadius:10,border:"1px solid "+(tone==="amber"?"#FDE68A":"#EEF2F7"),padding:"0.9rem 1.1rem"}}>
    <div style={{fontSize:"0.6rem",letterSpacing:"0.1em",color:tone==="amber"?"#B45309":BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.6rem"}}>{title}</div>
    {children}
  </div>;
}
const OWNER = { BFS:{bg:"#DBEAFE",fg:"#1D4ED8",t:"YOU"}, others:{bg:"#F1F5F9",fg:"#64748B",t:"OTHERS"}, unclear:{bg:"#FEF3C7",fg:"#B45309",t:"ASK"} };

/* ── Scope tab: read the project scope document, cross-check the bid ── */
function ScopeView({ result=null, setResult=()=>{} }) {
  const [scopeFile, setScopeFile] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef();
  const extractText = async (f) => {
    const name = (f.name||"").toLowerCase();
    if (name.endsWith(".xlsx")||name.endsWith(".xls")||name.endsWith(".csv")) {
      const wb = XLSX.read(await f.arrayBuffer(), {type:"array"});
      return wb.SheetNames.map(n=>`# Sheet: ${n}\n`+XLSX.utils.sheet_to_csv(wb.Sheets[n])).join("\n\n");
    }
    if (name.endsWith(".txt")) return await f.text();
    if (name.endsWith(".pdf")) {
      const fd = new FormData(); fd.append("pdf", f);
      const r = await fetch(BACKEND + "/scope-read", { method:"POST", body: fd });
      if (!r.ok) throw new Error("PDF read failed ("+r.status+")");
      const d = await r.json();
      if (!d.text || d.text.replace(/\s/g,"").length < 60) throw new Error("This PDF looks scanned (image-only) — upload the Excel, or a text-based PDF.");
      return d.text;
    }
    return null;
  };
  const analyze = async (text) => {
    const prompt = `The GC has sent this SCOPE OF WORK sheet for the siding / metal-panel bid package. You are the BFS estimator. Go DOWN the list and for EVERY scope line decide our response:
- "Y" = we include it / we agree (it's our scope)
- "N" = NOT our scope or we exclude it (e.g. air & weather barrier "by others", permits marked NO)
- "?" = ambiguous — we must ASK the GC before we can answer
Then compile every question we need to send the GC.

Return ONLY JSON (no markdown):
{
 "project":"name if present",
 "base_bid":"$ amount if present",
 "items":[
   {"line":"1","section":"General items","text":"concise scope item text under 140 chars","answer":"Y|N|?","reason":"short why — especially trade boundary: ours vs roofer vs window installer"}
 ],
 "questions":["the clarifications / RFIs to send the GC before bidding"],
 "materials":[
   {"name":"the cladding/facade material as named in the scope (e.g. insulated metal wall panels, fiber cement siding, brick veneer)","ours":true,"note":"short why in/out of our scope"}
 ]
}
Rules:
- Include every real scope line; skip blank or header-only rows. Keep the line number and the section heading it falls under.
- Trade-boundary lines (flashing, air & weather barrier, thru-wall flashing, counterflashing, sealant, soffit, roof edge): decide ours (Y) vs by-others (N) and say why; if truly unclear use "?".
- "questions" = everything marked "?" plus any spec-vs-detail conflicts, pending prices, or missing info.
- "materials" = every FACADE/CLADDING material the scope mentions (metal panel, fiber cement, lap siding, ACM, brick/masonry, stone, trim systems...), with "ours": true if this package installs it, false if it's by others (e.g. brick usually = mason's). This list pre-checks the takeoff.

SCOPE DOCUMENT:
${text.slice(0,30000)}`;
    const res = await fetch("/api/analyze",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-opus-4-8",max_tokens:12000,messages:[{role:"user",content:[{type:"text",text:prompt}]}]})});
    if(!res.ok) throw new Error("analysis service error ("+res.status+")");
    const data = await res.json();
    if(data?.error) throw new Error(data.error.message||"analysis error");
    let t = data?.content?.find(b=>b.type==="text")?.text || "";
    const m = t.match(/\{[\s\S]+\}/); if(m) t = m[0];
    return JSON.parse(t);
  };
  const handle = async (f) => {
    if(!f) return;
    setScopeFile(f); setResult(null); setError("");
    try {
      const text = await extractText(f);
      if(text===null){ setError("Upload the scope as Excel (.xlsx), PDF, or a text file."); return; }
      setBusy(true);
      setResult(await analyze(text));
    } catch(e){ setError("Couldn't read that scope sheet: "+e.message); }
    finally{ setBusy(false); }
  };
  return (
    <div style={{flex:1,overflowY:"auto",padding:"2.5rem 2rem"}}>
      <div style={{maxWidth:780,margin:"0 auto"}}>
        <div style={{fontSize:"0.7rem",letterSpacing:"0.18em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.4rem"}}>Scope Review</div>
        <h2 style={{fontSize:"1.5rem",fontWeight:800,color:"#0F172A",margin:"0 0 0.5rem",letterSpacing:"-0.02em"}}>Read the scope, know what's yours</h2>
        <p style={{fontSize:"0.85rem",color:"#64748B",lineHeight:1.6,margin:"0 0 1.5rem"}}>Upload the project scope document. The system simplifies it into plain English, flags every trade boundary — is that flashing <i>yours</i>, the roofer's, or the window installer's? — and preps the questions to bring to the scope meeting, so you never bid someone else's work or get stuck eating a cost you didn't price.</p>
        <div
          onClick={()=>fileRef.current?.click()}
          onDrop={e=>{e.preventDefault();setDragOver(false);handle(e.dataTransfer.files[0]);}}
          onDragOver={e=>{e.preventDefault();setDragOver(true);}}
          onDragLeave={()=>setDragOver(false)}
          style={{border:`2px dashed ${dragOver?BLUE:scopeFile?"#22C55E":"#CBD5E1"}`,borderRadius:14,padding:"2rem",textAlign:"center",cursor:"pointer",background:dragOver?BLUE_PALE:scopeFile?"#F0FDF4":"#fff",transition:"all 0.2s"}}>
          <input ref={fileRef} type="file" accept=".pdf,.doc,.docx,.txt,.xlsx,.csv" style={{display:"none"}} onChange={e=>handle(e.target.files[0])}/>
          {!scopeFile?(<>
            <div style={{fontSize:"2.2rem",marginBottom:"0.4rem",opacity:0.5}}>📄</div>
            <div style={{fontSize:"0.95rem",fontWeight:600,color:"#334155"}}>Drop the scope document here</div>
            <div style={{fontSize:"0.75rem",color:"#94A3B8",marginTop:"0.25rem"}}>PDF, Word, Excel, or text · or click to browse</div>
          </>):(<>
            <div style={{fontSize:"2rem",marginBottom:"0.35rem"}}>✅</div>
            <div style={{fontSize:"0.95rem",fontWeight:700,color:"#15803D",wordBreak:"break-all"}}>{scopeFile.name}</div>
            <div style={{fontSize:"0.72rem",color:"#16A34A",marginTop:"0.2rem"}}>Loaded · {(scopeFile.size/1024).toFixed(0)} KB</div>
          </>)}
        </div>
        {busy&&<div style={{marginTop:"1.5rem",textAlign:"center",color:"#64748B",fontSize:"0.85rem"}}>📖 Reading the scope…</div>}
        {error&&<div style={{marginTop:"1.25rem",padding:"0.85rem 1rem",background:"#FEF2F2",borderRadius:10,border:"1px solid #FECACA",fontSize:"0.78rem",color:"#B91C1C"}}>{error}</div>}

        {!result&&!busy&&!error&&(
          <div style={{marginTop:"1.5rem",display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0.75rem"}}>
            {[
              ["🧾","Plain-English summary","Everything the scope covers, simplified — no wading through pages of legalese"],
              ["🧱","Trade-boundary check","Every flashing, coping & sealant joint flagged: yours, the roofer's, or the window installer's"],
              ["❓","Questions for the meeting","Auto-prepped RFIs on anything doubtful, so you walk into the scope meeting ready"],
              ["📐","Quantities + in/exclusions","Required material quantities pulled out, plus proposal inclusions/exclusions"],
            ].map(([ic,t,d])=>(
              <div key={t} style={{padding:"0.85rem 1rem",background:"#fff",borderRadius:10,border:"1px solid #EEF2F7"}}>
                <div style={{fontSize:"1.1rem",marginBottom:"0.3rem"}}>{ic}</div>
                <div style={{fontSize:"0.78rem",fontWeight:700,color:"#0F172A"}}>{t}</div>
                <div style={{fontSize:"0.68rem",color:"#94A3B8",marginTop:"0.15rem",lineHeight:1.4}}>{d}</div>
              </div>
            ))}
          </div>
        )}

        {result&&(()=>{
          const items = result.items||[];
          const yN=items.filter(i=>i.answer==="Y").length, nN=items.filter(i=>i.answer==="N").length, qN=items.filter(i=>i.answer==="?").length;
          const setAnswer=(idx,a)=>setResult(r=>({...r,items:r.items.map((it,i)=>i===idx?{...it,answer:a}:it)}));
          const aStyle={Y:{bg:"#DCFCE7",fg:"#15803D"},N:{bg:"#FEE2E2",fg:"#B91C1C"},"?":{bg:"#FEF3C7",fg:"#B45309"}};
          return <div style={{marginTop:"1.5rem",display:"flex",flexDirection:"column",gap:"1rem"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:"0.5rem"}}>
              <div><div style={{fontSize:"0.95rem",fontWeight:800,color:"#0F172A"}}>{result.project||"Scope of Work"}</div>{result.base_bid&&<div style={{fontSize:"0.7rem",color:"#64748B"}}>Base bid: {result.base_bid}</div>}</div>
              <div style={{display:"flex",gap:"0.4rem"}}>
                <span style={{fontSize:"0.7rem",fontWeight:700,padding:"0.25rem 0.6rem",borderRadius:20,background:"#DCFCE7",color:"#15803D"}}>{yN} Yes</span>
                <span style={{fontSize:"0.7rem",fontWeight:700,padding:"0.25rem 0.6rem",borderRadius:20,background:"#FEE2E2",color:"#B91C1C"}}>{nN} No</span>
                <span style={{fontSize:"0.7rem",fontWeight:700,padding:"0.25rem 0.6rem",borderRadius:20,background:"#FEF3C7",color:"#B45309"}}>{qN} to ask</span>
              </div>
            </div>
            {result.questions?.length>0&&<ScopeSection title="❓ Questions to send the GC" tone="amber">
              <ul style={{margin:0,paddingLeft:"1.1rem"}}>{result.questions.map((q,i)=><li key={i} style={{fontSize:"0.74rem",color:"#92400E",marginBottom:"0.35rem",lineHeight:1.45}}>{q}</li>)}</ul>
            </ScopeSection>}
            {result.materials?.length>0&&(
              <div style={{padding:"0.6rem 0.85rem",background:"#F0FDF4",border:"1px solid #BBF7D0",borderRadius:10,fontSize:"0.7rem",color:"#15803D"}}>
                🔗 <b>{result.materials.length} material{result.materials.length>1?"s":""} read from the scope</b> — {result.materials.filter(m=>m.ours).map(m=>m.name).join(", ")||"none ours"}{result.materials.some(m=>!m.ours)?` · by others: ${result.materials.filter(m=>!m.ours).map(m=>m.name).join(", ")}`:""}. The Takeoff tab now pre-checks detected surfaces against this list.
              </div>
            )}
            <div style={{fontSize:"0.6rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700}}>Scope checklist — review each Yes / No / Ask</div>
            <div style={{background:"#fff",borderRadius:10,border:"1px solid #EEF2F7",overflow:"hidden"}}>
              {items.map((it,idx)=>{
                const showH = idx===0||items[idx-1].section!==it.section;
                return <div key={idx}>
                  {showH&&it.section&&<div style={{padding:"0.5rem 0.85rem",background:"#F8FAFC",fontSize:"0.6rem",letterSpacing:"0.08em",textTransform:"uppercase",fontWeight:700,color:"#64748B",borderTop:idx>0?"1px solid #F1F5F9":"none"}}>{it.section}</div>}
                  <div style={{display:"flex",gap:"0.6rem",alignItems:"flex-start",padding:"0.5rem 0.85rem",borderTop:"1px solid #F6F8FA"}}>
                    <span style={{flexShrink:0,fontSize:"0.6rem",color:"#CBD5E1",fontWeight:700,width:18,textAlign:"right"}}>{it.line}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:"0.74rem",color:"#0F172A"}}>{it.text}</div>
                      {it.reason&&<div style={{fontSize:"0.64rem",color:"#94A3B8",marginTop:1}}>{it.reason}</div>}
                    </div>
                    <div style={{display:"flex",gap:2,flexShrink:0}}>
                      {["Y","N","?"].map(a=>{const on=it.answer===a;const c=aStyle[a];return <button key={a} onClick={()=>setAnswer(idx,a)} style={{padding:"0.2rem 0.5rem",borderRadius:5,border:"1px solid "+(on?c.fg:"#E2E8F0"),background:on?c.bg:"#fff",color:on?c.fg:"#94A3B8",fontSize:"0.62rem",fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>{a}</button>;})}
                    </div>
                  </div>
                </div>;
              })}
            </div>
            <button onClick={()=>{setScopeFile(null);setResult(null);setError("");}} style={{alignSelf:"flex-start",padding:"0.5rem 1rem",background:"#fff",color:"#64748B",border:"1px solid #E2E8F0",borderRadius:7,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↺ New scope</button>
          </div>;
        })()}
      </div>
    </div>
  );
}

/* ── Model & Data tab: the moat, made visible ── */
const MOAT_MATERIALS = [
  ["Fiber Cement Panel", 9241],
  ["Metal Panel", 1048],
  ["Lap Siding", 880],
  ["ACM / Composite Panel", 643],
  ["Soffit / Trim", 244],
  ["PVC", 208],
  ["Shingle / Shake", 106],
  ["Board & Batten", 104],
];
/* ── Compare Lab: upload a MARKED takeoff → AI reads the stripped blank sheet →
      side-by-side "Estimator vs AI" on the same drawing, graded wall by wall.
      The killer demo: proof the system converges on the estimator's own work. ── */
function CompareLab() {
  const [st, setSt] = useState({ phase: "idle" });
  const inputRef = useRef();
  const pollRef = useRef(null);
  useEffect(() => () => { if (pollRef.current) clearInterval(pollRef.current); }, []);
  const run = async (file) => {
    if (!file) return;
    setSt({ phase: "running", label: "Uploading the marked takeoff…" });
    try {
      const fd = new FormData(); fd.append("pdf", file);
      const res = await (await fetch(BACKEND + "/compare", { method: "POST", body: fd })).json();
      const jobId = res.jobId;
      pollRef.current = setInterval(async () => {
        try {
          const r = await (await fetch(BACKEND + "/compare-result/" + jobId)).json();
          if (r.status === "done" && r.compare) {
            clearInterval(pollRef.current); pollRef.current = null;
            setSt({ phase: "done", jobId, data: r.compare });
          } else if (r.status === "error") {
            clearInterval(pollRef.current); pollRef.current = null;
            setSt({ phase: "error", error: r.error || "Comparison failed" });
          } else {
            setSt(s => ({ ...s, label: (r.progress || {}).label || "AI reading the blank sheet…" }));
          }
        } catch {}
      }, 4000);
    } catch (e) { setSt({ phase: "error", error: String(e) }); }
  };
  const P = (pts) => pts.map(p => `${(p[0] * 1000).toFixed(1)},${(p[1] * 1000).toFixed(1)}`).join(" ");
  const Pane = ({ jobId, page, polys, holes, color, title, total }) => (
    <div style={{ flex: 1, minWidth: 0 }}>
      <div style={{ fontSize: "0.62rem", fontWeight: 800, color: "#334155", marginBottom: 4 }}>{title}</div>
      <div style={{ position: "relative", border: "1px solid #E2E8F0", borderRadius: 8, overflow: "hidden", background: "#fff" }}>
        <img src={`${BACKEND}/page-image/${jobId}/${page}`} alt={title} style={{ width: "100%", display: "block" }} />
        <svg viewBox="0 0 1000 1000" preserveAspectRatio="none"
             style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}>
          {polys.map((p, i) => (
            <polygon key={i} points={P(p.points)} fill={color} fillOpacity="0.42" stroke={color} strokeOpacity="0.85" strokeWidth="1.2" />
          ))}
          {(holes || []).map((h, i) => (
            <polygon key={"h" + i} points={P(h)} fill="#fff" fillOpacity="0.9" />
          ))}
        </svg>
      </div>
      <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#0F172A", marginTop: 4 }}>{Math.round(total).toLocaleString()} SF</div>
    </div>
  );
  return (
    <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #EEF2F7", padding: "1.1rem 1.25rem", marginBottom: "1.5rem" }}>
      <div style={{ fontSize: "0.6rem", letterSpacing: "0.1em", color: BLUE, textTransform: "uppercase", fontWeight: 700, marginBottom: "0.35rem" }}>⚔️ Compare Lab — estimator vs AI, same drawing</div>
      <div style={{ fontSize: "0.7rem", color: "#64748B", lineHeight: 1.55, marginBottom: "0.75rem" }}>
        Drop any <b>marked-up</b> Bluebeam set. The system reads the takeoff exactly, <b>strips it to a blank sheet</b>, does its own takeoff from scratch, and grades itself against every wall you measured. Nothing hidden — this is the benchmark, live.
      </div>
      {st.phase !== "done" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
          <button onClick={() => inputRef.current && inputRef.current.click()} disabled={st.phase === "running"}
                  style={{ padding: "0.55rem 1.1rem", borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 800, fontSize: "0.75rem", background: st.phase === "running" ? "#CBD5E1" : BLUE, color: "#fff" }}>
            {st.phase === "running" ? "Running…" : "Upload a marked takeoff"}
          </button>
          {st.phase === "running" && <span style={{ fontSize: "0.68rem", color: "#64748B" }}>{st.label}</span>}
          {st.phase === "error" && <span style={{ fontSize: "0.68rem", color: "#B91C1C" }}>{st.error}</span>}
          <input ref={inputRef} type="file" accept="application/pdf" style={{ display: "none" }}
                 onChange={e => run(e.target.files && e.target.files[0])} />
        </div>
      )}
      {st.phase === "done" && st.data && (() => {
        const s = st.data.summary || {};
        return (
          <div>
            <div style={{ display: "flex", gap: "0.6rem", flexWrap: "wrap", marginBottom: "0.9rem" }}>
              {[[`${s.money} of ${s.walls}`, "walls MONEY-RIGHT, pre-human", s.money / Math.max(1, s.walls) >= 0.6 ? "#15803D" : "#B45309"],
                [`${s.found} of ${s.walls}`, "walls found", "#334155"],
                [`${Math.round(s.hisTotal).toLocaleString()}`, "estimator total SF", "#334155"],
                [`${Math.round(s.ourTotal).toLocaleString()}`, "AI total SF (full sheet)", "#334155"]].map(([n, l, c]) => (
                <div key={l} style={{ padding: "0.55rem 0.8rem", background: "#F8FAFC", borderRadius: 8, border: "1px solid #EEF2F7" }}>
                  <div style={{ fontSize: "1.05rem", fontWeight: 800, color: c }}>{n}</div>
                  <div style={{ fontSize: "0.56rem", color: "#94A3B8" }}>{l}</div>
                </div>
              ))}
              <button onClick={() => setSt({ phase: "idle" })} style={{ marginLeft: "auto", alignSelf: "center", padding: "0.4rem 0.8rem", borderRadius: 8, border: "1px solid #E2E8F0", background: "#fff", cursor: "pointer", fontSize: "0.66rem", fontWeight: 700, color: "#475569" }}>Run another</button>
            </div>
            {st.data.pages.map(pg => (
              <div key={pg.page} style={{ marginBottom: "1.1rem" }}>
                <div style={{ fontSize: "0.66rem", fontWeight: 800, color: "#475569", margin: "0.4rem 0" }}>Page {pg.page}{!pg.scale_confirmed && <span style={{ color: "#B45309" }}> · scale unconfirmed</span>}</div>
                <div style={{ display: "flex", gap: "0.75rem" }}>
                  <Pane jobId={st.jobId} page={pg.page} polys={pg.his} color="#15803D" title="ESTIMATOR (hand-marked)" total={pg.hisTotal} />
                  <Pane jobId={st.jobId} page={pg.page} polys={pg.ours} holes={pg.ours.flatMap(o => o.holes || [])} color="#3F79BC" title="AI (from the blank sheet)" total={pg.ourTotal} />
                </div>
                <div style={{ marginTop: "0.5rem" }}>
                  {pg.walls.map((w, i) => (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: "0.6rem", padding: "0.25rem 0", borderBottom: "1px solid #F1F5F9", fontSize: "0.68rem" }}>
                      <div style={{ width: 190, color: "#374151", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{w.mat}</div>
                      <div style={{ width: 90, textAlign: "right", fontWeight: 700, color: "#0F172A" }}>{Math.round(w.sf).toLocaleString()} SF</div>
                      <div style={{ width: 20, textAlign: "center", color: "#94A3B8" }}>→</div>
                      <div style={{ width: 90, textAlign: "right", fontWeight: 700, color: "#334155" }}>{Math.round(w.got).toLocaleString()} SF</div>
                      <div style={{ width: 70, textAlign: "right", color: "#64748B" }}>cov {Math.round(w.cov * 100)}%</div>
                      {w.money
                        ? <span style={{ fontSize: "0.6rem", fontWeight: 800, color: "#15803D", background: "#DCFCE7", borderRadius: 6, padding: "0.1rem 0.45rem" }}>MONEY ✓</span>
                        : <span style={{ fontSize: "0.6rem", fontWeight: 700, color: "#B45309", background: "#FEF3C7", borderRadius: 6, padding: "0.1rem 0.45rem" }}>{w.cov >= 0.3 ? "close" : "miss"}</span>}
                    </div>
                  ))}
                </div>
              </div>
            ))}
            <div style={{ fontSize: "0.6rem", color: "#94A3B8", marginTop: "0.4rem" }}>MONEY ✓ = the AI's assembled SF for that wall lands within ±15% with ≥70% shape coverage — the same bar the internal benchmark uses.</div>
          </div>
        );
      })()}
    </div>
  );
}

function ModelView() {
  const maxN = Math.max(...MOAT_MATERIALS.map(m=>m[1]));
  const [auto, setAuto] = useState(null);   // live autonomy meter: auto output vs human-confirmed finals
  useEffect(()=>{ fetch(BACKEND+"/autonomy-status").then(r=>r.ok?r.json():null).then(setAuto).catch(()=>{}); },[]);
  return (
    <div style={{flex:1,overflowY:"auto",padding:"2.5rem 2rem"}}>
      <div style={{maxWidth:860,margin:"0 auto"}}>
        <div style={{fontSize:"0.7rem",letterSpacing:"0.18em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.4rem"}}>The Model</div>
        <h2 style={{fontSize:"1.6rem",fontWeight:800,color:"#0F172A",margin:"0 0 0.4rem",letterSpacing:"-0.02em"}}>Trained on Boston Facade Systems' own work</h2>
        <p style={{fontSize:"0.85rem",color:"#9FB3CC",lineHeight:1.6,margin:"0 0 1.75rem",maxWidth:660}}>This isn't generic AI. It learns to highlight materials and trace shapes the way <i>your</i> estimators do — from every takeoff your team has ever marked up. The more bids you run, the sharper it gets. No competitor can copy this — it's your data.</p>
        <CompareLab/>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"0.75rem",marginBottom:"1.75rem"}}>
          {[["657","marked takeoffs"],["16,368","labeled regions"],["2024–26","3 years"],["12","material types"]].map(([n,l])=>(
            <div key={l} style={{padding:"1rem",background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",textAlign:"center"}}>
              <div style={{fontSize:"1.55rem",fontWeight:800,color:BLUE}}>{n}</div>
              <div style={{fontSize:"0.6rem",color:"#94A3B8",marginTop:2}}>{l}</div>
            </div>
          ))}
        </div>
        <div style={{background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",padding:"1.1rem 1.25rem",marginBottom:"1.5rem"}}>
          <div style={{fontSize:"0.6rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.85rem"}}>Materials it's learning (training examples)</div>
          {MOAT_MATERIALS.map(([m,n])=>(
            <div key={m} style={{display:"flex",alignItems:"center",gap:"0.6rem",marginBottom:"0.5rem"}}>
              <div style={{width:160,fontSize:"0.7rem",color:"#374151",flexShrink:0}}>{m}</div>
              <div style={{flex:1,height:14,background:"#F1F5F9",borderRadius:7,overflow:"hidden"}}>
                <div style={{width:(n/maxN*100)+"%",height:"100%",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",borderRadius:7}}/>
              </div>
              <div style={{width:56,textAlign:"right",fontSize:"0.66rem",color:"#64748B",fontWeight:700}}>{n.toLocaleString()}</div>
            </div>
          ))}
        </div>
        {auto&&auto.n>0&&(
          <div style={{background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",padding:"1.1rem 1.25rem",marginBottom:"1.5rem"}}>
            <div style={{fontSize:"0.6rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.35rem"}}>Autonomy meter — live from your bids</div>
            <div style={{display:"flex",alignItems:"baseline",gap:"0.75rem",marginBottom:"0.75rem"}}>
              <span style={{fontSize:"2.2rem",fontWeight:800,color:auto.avg_agreement>=85?"#15803D":auto.avg_agreement>=60?"#B45309":"#B91C1C"}}>{auto.avg_agreement}%</span>
              <span style={{fontSize:"0.72rem",color:"#64748B"}}>of the final SF, the system got right <b>before any human touched it</b> · {auto.n} exported bid{auto.n!==1?"s":""}</span>
            </div>
            {auto.jobs.slice(0,6).map(j=>(
              <div key={j.jobId} style={{display:"flex",alignItems:"center",gap:"0.6rem",marginBottom:"0.45rem"}}>
                <div style={{width:180,fontSize:"0.68rem",color:"#374151",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flexShrink:0}}>{j.projName}</div>
                <div style={{flex:1,height:12,background:"#F1F5F9",borderRadius:6,overflow:"hidden"}}>
                  <div style={{width:j.agreement+"%",height:"100%",background:j.agreement>=85?"linear-gradient(180deg,#34D399,#10B981)":"linear-gradient(180deg,#FBBF24,#D97706)",borderRadius:6}}/>
                </div>
                <div style={{width:100,textAlign:"right",fontSize:"0.64rem",color:"#64748B"}}><b>{j.agreement}%</b> · {j.auto_sf.toLocaleString()}→{j.final_sf.toLocaleString()}</div>
              </div>
            ))}
            <div style={{fontSize:"0.6rem",color:"#94A3B8",marginTop:"0.5rem"}}>Auto SF → confirmed SF per bid. When this meter lives near 100%, estimators stop reviewing and only price.</div>
          </div>
        )}
        <div style={{background:NAVY,borderRadius:12,padding:"1.1rem 1.25rem",color:"#fff",marginBottom:"1.5rem"}}>
          <div style={{fontSize:"0.6rem",letterSpacing:"0.1em",color:"rgba(255,255,255,0.5)",textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>Model status — measured honestly</div>
          <div style={{fontSize:"0.95rem",fontWeight:700}}>🟢 v11 in production — fair-eval winner across 3 trained models</div>
          <div style={{fontSize:"0.68rem",color:"rgba(255,255,255,0.55)",marginTop:"0.4rem",lineHeight:1.5}}>Every number below is measured on jobs the model has <i>never seen</i> (split by job — no leakage). The honest read, not a demo score.</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"0.6rem",marginTop:"0.9rem"}}>
            {[["0.664","cladding-detection IoU","v11 beat v10's 0.607 on the same held-out exam — ~10% more cladding found"],["Per-wall","first paint splits at joints","walls arrive as separate pieces, like a marked takeoff"],["Exact","your bid SF","from geometry + your confirm — never a guess"]].map(([n,t,d])=>(
              <div key={t} style={{background:"rgba(255,255,255,0.06)",borderRadius:8,padding:"0.7rem 0.75rem"}}>
                <div style={{fontSize:"1.25rem",fontWeight:800,color:"#7FB0E0"}}>{n}</div>
                <div style={{fontSize:"0.64rem",fontWeight:700,color:"rgba(255,255,255,0.85)",marginTop:2}}>{t}</div>
                <div style={{fontSize:"0.58rem",color:"rgba(255,255,255,0.45)",marginTop:2,lineHeight:1.35}}>{d}</div>
              </div>
            ))}
          </div>
          <div style={{fontSize:"0.64rem",color:"rgba(255,255,255,0.6)",marginTop:"0.85rem",lineHeight:1.5,borderTop:"1px solid rgba(255,255,255,0.08)",paddingTop:"0.7rem"}}>
            <b style={{color:"#93C5FD"}}>How to read this:</b> the model's job is to <i>find + pre-mark</i> the cladding so the estimator selects instead of tracing. The <b>bid SF stays exact</b> because she confirms each group and the number comes from the drawing's own geometry. Naming metal-vs-lap across different architects isn't reliable from the drawing (it lives in the legend) — so she names the material in one click, and the model gets sharper on <i>your</i> recurring jobs every cycle.
          </div>
        </div>
        <div style={{background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",padding:"1.1rem 1.25rem",marginBottom:"1.5rem"}}>
          <div style={{fontSize:"0.6rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>📏 Scored against your own takeoffs — wall by wall</div>
          <div style={{fontSize:"0.7rem",color:"#64748B",lineHeight:1.55,marginBottom:"0.85rem"}}>Every past marked-up bid is an answer key: the system strips your markups, reads the bare drawing, and gets graded against what you measured — <b>per wall, not just totals</b> (totals can lie; two wrong walls can sum right).</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"0.6rem"}}>
            {[["228 → 228","reference job, wall-exact","clicked wall returns the SF you marked, to the foot"],["−0.7%","reference job, all 10 walls","one click per wall vs your hand takeoff"],["13%","across 12 other jobs","the dial being driven up — each session is graded on more of your past bids"]].map(([n,t,d])=>(
              <div key={t} style={{background:"#F8FAFC",borderRadius:8,padding:"0.7rem 0.75rem",border:"1px solid #EEF2F7"}}>
                <div style={{fontSize:"1.15rem",fontWeight:800,color:BLUE}}>{n}</div>
                <div style={{fontSize:"0.64rem",fontWeight:700,color:"#334155",marginTop:2}}>{t}</div>
                <div style={{fontSize:"0.58rem",color:"#94A3B8",marginTop:2,lineHeight:1.35}}>{d}</div>
              </div>
            ))}
          </div>
          <div style={{fontSize:"0.62rem",color:"#94A3B8",marginTop:"0.7rem",lineHeight:1.5}}>Next reader in training: <b style={{color:"#475569"}}>diagonal-hatch materials</b> (masonry/EIFS drawn as 45° strokes) — found by this benchmark, ships only when the reference job's numbers hold exactly.</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"0.75rem"}}>
          {[["📐","Run a bid","Estimator marks the takeoff like always"],["🧠","It learns","Every shape becomes a new training example"],["⚡","Gets sharper","Next similar drawing, it does more for you"]].map(([ic,t,d])=>(
            <div key={t} style={{padding:"0.9rem 1rem",background:"#fff",borderRadius:10,border:"1px solid #EEF2F7"}}>
              <div style={{fontSize:"1.2rem"}}>{ic}</div>
              <div style={{fontSize:"0.78rem",fontWeight:700,color:"#0F172A",marginTop:"0.2rem"}}>{t}</div>
              <div style={{fontSize:"0.66rem",color:"#94A3B8",marginTop:"0.1rem",lineHeight:1.4}}>{d}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ── Queue tab: batch-process the day's drawings ── */
function QueueView({ onOpen }) {
  const [items, setItems] = useState(() => { try { return JSON.parse(localStorage.getItem("bfs_queue_v1")) || []; } catch { return []; } });
  const filesRef = useRef({});      // id -> File (not persisted across reload)
  const pollingRef = useRef({});    // jobId -> interval
  const busyRef = useRef(false);
  const [drag, setDrag] = useState(false);
  const inputRef = useRef();

  useEffect(() => { try { localStorage.setItem("bfs_queue_v1", JSON.stringify(items)); } catch {} }, [items]);
  const patch = (id, p) => setItems(prev => prev.map(it => it.id === id ? { ...it, ...p } : it));
  const sfOf = (r) => Math.round((r?.takeoffData || []).reduce((s, e) => s + (e.zones || []).reduce((a, z) => a + (z.netArea || 0), 0), 0));

  const pollJob = useCallback((id, jobId) => {
    if (pollingRef.current[jobId]) return;
    const iv = setInterval(async () => {
      try {
        const s = await (await fetch(BACKEND + "/status/" + jobId)).json();
        if (s.progress) patch(id, { progress: s.progress.pct || 0, stage: s.progress.label || s.phase || "" });
        if (s.status === "done") {
          clearInterval(iv); delete pollingRef.current[jobId]; busyRef.current = false;
          patch(id, { status: "done", progress: 100, stage: "Complete",
            results: { legend: s.legend || [], takeoffData: s.takeoffData || [], scheduleData: s.scheduleData || null, drawingSchedule: s.drawingSchedule || null, ocrMaterials: s.ocrMaterials || null, projName: (s.projName) || "", jobId } });
        } else if (s.status === "error") {
          clearInterval(iv); delete pollingRef.current[jobId]; busyRef.current = false;
          patch(id, { status: "error", error: s.error || "Analysis failed" });
        }
      } catch {}
    }, 5000);
    pollingRef.current[jobId] = iv;
  }, []);

  // resume any in-flight jobs across a reload (jobId survives even if the File doesn't)
  useEffect(() => {
    items.forEach(it => { if (it.jobId && it.status === "running") pollJob(it.id, it.jobId); });
    return () => { Object.values(pollingRef.current).forEach(clearInterval); pollingRef.current = {}; };
  }, []); // eslint-disable-line

  // sequential processor: start the next queued file whenever nothing is running
  useEffect(() => {
    if (busyRef.current) return;
    if (items.some(it => it.status === "running")) { busyRef.current = true; return; }
    const next = items.find(it => it.status === "queued" && filesRef.current[it.id]);
    if (!next) return;
    busyRef.current = true;
    (async () => {
      try {
        patch(next.id, { status: "running", progress: 4, stage: "Uploading" });
        const fd = new FormData(); fd.append("pdf", filesRef.current[next.id]);
        const r = await fetch(BACKEND + "/analyze", { method: "POST", body: fd });
        const { jobId } = await r.json();
        patch(next.id, { jobId }); pollJob(next.id, jobId);
      } catch (e) { busyRef.current = false; patch(next.id, { status: "error", error: e.message }); }
    })();
  }, [items, pollJob]);

  const addFiles = (list) => {
    const add = [...list].filter(f => f.type === "application/pdf").map(f => {
      const id = "q" + Date.now() + Math.random().toString(36).slice(2, 6); filesRef.current[id] = f;
      return { id, name: f.name, status: "queued", progress: 0, addedAt: Date.now() };
    });
    if (add.length) setItems(prev => [...prev, ...add]);
  };
  const removeItem = (id) => {
    const it = items.find(x => x.id === id);
    if (it?.jobId && pollingRef.current[it.jobId]) { clearInterval(pollingRef.current[it.jobId]); delete pollingRef.current[it.jobId]; }
    if (it?.status === "running") busyRef.current = false;
    setItems(prev => prev.filter(x => x.id !== id));
  };
  const clearDone = () => setItems(prev => prev.filter(it => it.status !== "done"));

  const BADGE = { queued: ["#64748B", "Queued"], running: ["#B45309", "Working"], done: ["#15803D", "Done"], error: ["#B91C1C", "Error"] };
  return (
    <div style={{ flex: 1, overflowY: "auto", padding: "2rem" }}>
      <div style={{ maxWidth: 820, margin: "0 auto" }}>
        <h2 style={{ fontSize: "1.4rem", fontWeight: 800, color: "#0F172A", margin: "0 0 0.3rem" }}>Daily Queue</h2>
        <p style={{ fontSize: "0.82rem", color: "#64748B", margin: "0 0 1.25rem" }}>Drop the whole day's drawings here in the morning. They process one after another — open each result as it finishes.</p>
        <div onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); addFiles(e.dataTransfer.files); }}
          onClick={() => inputRef.current?.click()}
          style={{ border: "2px dashed " + (drag ? BLUE : "#CBD5E1"), background: drag ? "#EFF6FF" : "#fff", borderRadius: 14, padding: "2rem", textAlign: "center", cursor: "pointer", marginBottom: "1.25rem" }}>
          <input ref={inputRef} type="file" accept="application/pdf" multiple style={{ display: "none" }} onChange={e => addFiles(e.target.files)} />
          <div style={{ fontSize: "1.6rem" }}>📥</div>
          <div style={{ fontSize: "0.85rem", fontWeight: 700, color: "#334155", marginTop: "0.4rem" }}>Drop PDFs here, or click to choose</div>
          <div style={{ fontSize: "0.7rem", color: "#94A3B8", marginTop: "0.2rem" }}>Add as many as you want — they line up and process automatically</div>
        </div>
        {items.length > 0 && (
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.6rem" }}>
            <div style={{ fontSize: "0.72rem", color: "#64748B", fontWeight: 600 }}>{items.filter(i => i.status === "done").length} of {items.length} done</div>
            {items.some(i => i.status === "done") && <button onClick={clearDone} style={{ fontSize: "0.68rem", color: "#64748B", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Clear finished</button>}
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem" }}>
          {items.map(it => {
            const [bc, bl] = BADGE[it.status] || BADGE.queued; const sf = it.results ? sfOf(it.results) : 0;
            return (
              <div key={it.id} style={{ background: "#fff", border: "1px solid #EEF2F7", borderRadius: 10, padding: "0.8rem 1rem", display: "flex", alignItems: "center", gap: "0.85rem" }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: "0.8rem", fontWeight: 700, color: "#1E293B", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{it.name}</div>
                  <div style={{ fontSize: "0.66rem", color: "#94A3B8", marginTop: 2 }}>
                    {it.status === "done" ? (sf > 0 ? sf.toLocaleString() + " SF measured" : "Complete") : it.status === "error" ? (it.error || "Failed") : (it.stage || "Waiting") + (it.status === "running" ? " · " + (it.progress || 0) + "%" : "")}
                  </div>
                  {it.status === "running" && <div style={{ height: 4, background: "#F1F5F9", borderRadius: 3, marginTop: 6, overflow: "hidden" }}><div style={{ width: (it.progress || 0) + "%", height: "100%", background: BLUE, transition: "width 0.4s" }} /></div>}
                </div>
                <span style={{ fontSize: "0.58rem", fontWeight: 700, color: "#fff", background: bc, padding: "0.2rem 0.5rem", borderRadius: 20, whiteSpace: "nowrap" }}>{bl}</span>
                {it.status === "done" && <button onClick={() => onOpen(it.results)} style={{ fontSize: "0.7rem", fontWeight: 700, color: "#fff", background: BLUE, border: "none", borderRadius: 7, padding: "0.35rem 0.8rem", cursor: "pointer" }}>Open</button>}
                <button onClick={() => removeItem(it.id)} title="Remove" style={{ fontSize: "0.85rem", color: "#CBD5E1", background: "none", border: "none", cursor: "pointer" }}>✕</button>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ── Manual takeoff: draw the facade, subtract cut-outs, get the area (Bluebeam-style) ── */
function ManualView({ results, BACKEND }) {
  const elevations = (results?.takeoffData || []).filter(e => e.pageNumber);
  const [elevIdx, setElevIdx] = useState(0);
  const elev = elevations[elevIdx]; const pageNum = elev?.pageNumber;
  const [img, setImg] = useState(null);
  const [pageDims, setPageDims] = useState({ width: 612, height: 792 });
  const [stageW, setStageW] = useState(760);
  const wrapRef = useRef();
  const [shapes, setShapes] = useState([]);     // {id, points[[nx,ny]], type:"add"|"cut", name, color}
  const [draft, setDraft] = useState([]);        // points being placed
  const [mode, setMode] = useState("add");       // "add" | "cut" | "calib"
  const [calib, setCalib] = useState(null);      // { ftPerPx }
  const [calibPts, setCalibPts] = useState([]);
  const [realFt, setRealFt] = useState("");
  const [selId, setSelId] = useState(null);
  const [curColor, setCurColor] = useState("#4A86C8");
  const [colorNames, setColorNames] = useState({});   // color -> material name
  const [taught, setTaught] = useState(null);
  const [snapPts, setSnapPts] = useState([]);          // drawing's real CAD corners (Bluebeam-style snap)
  const edited = useRef(false);                        // true once she actually edits (so flywheel learns only real corrections)
  const undoRef = useRef([]);                          // history of shapes snapshots (Ctrl+Z)
  const redoRef = useRef([]);
  const shapesRef = useRef(shapes); shapesRef.current = shapes;
  const draftRef = useRef(draft); draftRef.current = draft;
  const pushHist = () => { undoRef.current = [...undoRef.current.slice(-49), shapesRef.current]; redoRef.current = []; };

  const rgbToHex = c => (Array.isArray(c) && c.length >= 3)
    ? "#" + c.slice(0, 3).map(v => Math.max(0, Math.min(255, Math.round(v * 255))).toString(16).padStart(2, "0")).join("")
    : curColor;

  useEffect(() => {
    if (!pageNum || !results?.jobId) { setImg(null); return; }
    setShapes([]); setDraft([]); setSelId(null); setImg(null); setCalib(null); setCalibPts([]); setSnapPts([]);
    edited.current = false; undoRef.current = []; redoRef.current = [];
    const im = new window.Image();
    im.crossOrigin = "anonymous";
    im.src = BACKEND + "/page-image/" + results.jobId + "/" + pageNum;
    im.onload = () => { setImg(im); setPageDims({ width: im.naturalWidth || 612, height: im.naturalHeight || 792 }); };
    fetch(BACKEND + "/snap-points/" + results.jobId + "/" + pageNum)
      .then(r => r.ok ? r.json() : { points: [] }).then(d => setSnapPts(d.points || [])).catch(() => setSnapPts([]));
    // EDIT-THE-AI: pre-load the AI's detected shapes so she edits the AI's work, not redraw from scratch
    fetch(BACKEND + "/polygons/" + results.jobId + "/" + pageNum)
      .then(r => r.ok ? r.json() : { polygons: [] }).then(d => {
        const W = d.width || 612, H = d.height || 792;
        const shoe = pts => { let a = 0; for (let i = 0; i < pts.length; i++) { const x1 = pts[i][0] * W, y1 = pts[i][1] * H, x2 = pts[(i + 1) % pts.length][0] * W, y2 = pts[(i + 1) % pts.length][1] * H; a += x1 * y2 - x2 * y1; } return Math.abs(a) / 2; };
        const ref = (d.polygons || []).find(p => p.area_sf > 0 && p.points && p.points.length >= 3);
        if (ref) { const sp = shoe(ref.points); if (sp > 0) setCalib({ ftPerPx: Math.sqrt(ref.area_sf / sp) }); }  // auto-calibrate from AI's exact SF
        setShapes((d.polygons || []).filter(p => p.points && p.points.length >= 3).map((p, i) => ({
          id: Date.now() + i, points: p.points, type: "add",
          name: p.material || p.category || ("Area " + (i + 1)), color: rgbToHex(p.fill_color)
        })));
      }).catch(() => {});
  }, [elevIdx, pageNum, results?.jobId, BACKEND]);

  useEffect(() => {
    const u = () => { if (wrapRef.current) setStageW(Math.max(320, wrapRef.current.offsetWidth - 32)); };
    u(); window.addEventListener("resize", u); return () => window.removeEventListener("resize", u);
  }, []);

  const sc = stageW / pageDims.width;
  const stageH = pageDims.height * sc;
  const zoom = useKonvaZoom(BACKEND, results?.jobId, pageNum, stageW, stageH);
  const COLS = ["#4A86C8", "#E85DA0", "#3FB36B", "#F0A23C", "#9B6FD4", "#46C5C5"];

  const shoelacePx = pts => {
    let a = 0; const n = pts.length;
    for (let i = 0; i < n; i++) {
      const x1 = pts[i][0] * pageDims.width, y1 = pts[i][1] * pageDims.height;
      const x2 = pts[(i + 1) % n][0] * pageDims.width, y2 = pts[(i + 1) % n][1] * pageDims.height;
      a += x1 * y2 - x2 * y1;
    }
    return Math.abs(a) / 2;
  };
  const areaSF = s => calib ? shoelacePx(s.points) * calib.ftPerPx * calib.ftPerPx : 0;
  const total = shapes.reduce((t, s) => t + (s.type === "cut" ? -areaSF(s) : areaSF(s)), 0);
  const byColor = {};
  shapes.forEach(s => { if (!byColor[s.color]) byColor[s.color] = { sf: 0, n: 0 }; byColor[s.color].sf += (s.type === "cut" ? -1 : 1) * areaSF(s); byColor[s.color].n++; });

  // Bluebeam-style object snap: corners first, then ortho-lock to the previous point.
  // Radius shrinks as you zoom in (constant SCREEN feel, precise when zoomed).
  const snapEval = (nx, ny) => {
    const tol = 14 / (zoom.view?.k || 1);
    let best = null, bd = tol;
    for (const c of snapPts) {
      const d = Math.hypot((c[0] - nx) * stageW, (c[1] - ny) * stageH);
      if (d < bd) { bd = d; best = c; }
    }
    if (best) return { pt: [best[0], best[1]], kind: "corner" };
    const prev = draft.length ? draft[draft.length - 1] : null;
    if (prev) {  // 90° ortho lock like Bluebeam's axis cue
      const dx = Math.abs(nx - prev[0]) * stageW, dy = Math.abs(ny - prev[1]) * stageH;
      if (dx < tol && dy > tol) return { pt: [prev[0], ny], kind: "ortho" };
      if (dy < tol && dx > tol) return { pt: [nx, prev[1]], kind: "ortho" };
    }
    return { pt: [nx, ny], kind: null };
  };
  const [hoverSnap, setHoverSnap] = useState(null);   // live indicator: what the next click will lock to
  const onHover = e => {
    const st = e.target.getStage(); const pos = zoom.localPos(st); if (!pos) return;
    const s = snapEval(pos.x / stageW, pos.y / stageH);
    setHoverSnap(s.kind ? s : null);
  };
  const onDown = e => {
    const st = e.target.getStage(); const pos = zoom.localPos(st); if (!pos) return;
    const [nx, ny] = snapEval(pos.x / stageW, pos.y / stageH).pt;  // lock to the exact corner like Bluebeam
    if (mode === "calib") { setCalibPts(prev => [...prev, [nx, ny]].slice(-2)); return; }
    setDraft(prev => [...prev, [nx, ny]]);
  };
  const finish = () => {
    if (draft.length < 3) return;
    edited.current = true;
    pushHist();
    setShapes(prev => [...prev, { id: Date.now(), points: draft, type: mode === "cut" ? "cut" : "add",
      name: mode === "cut" ? "Cutout" : "Area " + (prev.filter(s => s.type !== "cut").length + 1),
      color: curColor }]);
    setDraft([]);
  };
  const applyCalib = () => {
    const ft = parseFloat(realFt);
    if (calibPts.length < 2 || !ft) return;
    const [p1, p2] = calibPts;
    const dpx = Math.hypot((p2[0] - p1[0]) * pageDims.width, (p2[1] - p1[1]) * pageDims.height);
    if (dpx > 0) setCalib({ ftPerPx: ft / dpx });
    setCalibPts([]); setMode("add");
  };
  const rename = (id, nm) => { edited.current = true; setShapes(prev => prev.map(s => s.id === id ? { ...s, name: nm } : s)); };
  const del = id => { edited.current = true; pushHist(); setShapes(prev => prev.filter(s => s.id !== id)); setSelId(null); };
  // Ctrl+Z / Ctrl+Y — undo an accidental click or shape (drew by accident = must be reversible)
  useEffect(() => {
    const onKey = e => {
      const tag = (e.target && e.target.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea") return;
      const k = (e.key || "").toLowerCase();
      if (k === "escape" && draftRef.current.length) { setDraft([]); return; }
      if (!(e.ctrlKey || e.metaKey)) return;
      if (k === "z" && !e.shiftKey) {
        e.preventDefault();
        if (draftRef.current.length) { setDraft(d => d.slice(0, -1)); return; }  // step back one click first
        if (!undoRef.current.length) return;
        const prev = undoRef.current.pop();
        redoRef.current.push(shapesRef.current);
        edited.current = true; setShapes(prev); setSelId(null);
      } else if (k === "y" || (k === "z" && e.shiftKey)) {
        e.preventDefault();
        if (!redoRef.current.length) return;
        const next = redoRef.current.pop();
        undoRef.current.push(shapesRef.current);
        edited.current = true; setShapes(next); setSelId(null);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);
  // Silent auto-learning: only save as a training example once SHE actually edits (not the AI's own pre-loaded shapes).
  useEffect(() => {
    if (!results?.jobId || !shapes.length || !edited.current) return;
    const t = setTimeout(() => {
      fetch(BACKEND + "/learn", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ jobId: results.jobId, page: pageNum, source: "draw",
          shapes: shapes.map(s => ({ points: s.points, name: colorNames[s.color] || s.name, color: s.color, type: s.type })) }) })
        .then(() => setTaught("saved")).catch(() => {});
    }, 2500);
    return () => clearTimeout(t);
  }, [shapes, colorNames, results?.jobId, pageNum, BACKEND]);

  const btn = (active) => ({ flex: 1, padding: "0.5rem", borderRadius: 8, border: "1px solid " + (active ? "transparent" : "#2D5280"), background: active ? "linear-gradient(180deg,#5A92D2,#3F79BC)" : "transparent", color: active ? "#fff" : "#94A3B8", fontSize: "0.66rem", fontWeight: 600, cursor: "pointer", fontFamily: "inherit", boxShadow: active ? "0 4px 12px -3px rgba(74,134,200,0.5)" : "none", letterSpacing: "-0.01em" });

  return (
    <div style={{ display: "flex", flex: 1, minHeight: 0, overflow: "hidden", background: "linear-gradient(180deg,#F5F8FC,#E9F0F8)", fontFamily: "'Inter',system-ui,sans-serif" }}>
      <div style={{ width: 176, borderRight: "1px solid #E4EAF1", overflowY: "auto", flexShrink: 0, background: "#fff" }}>
        <div style={{ padding: "0.95rem", fontSize: "0.6rem", letterSpacing: "0.13em", color: "#94A3B8", textTransform: "uppercase", fontWeight: 700, borderBottom: "1px solid #EEF2F7" }}>Pages</div>
        {elevations.map((e, i) => <div key={i} onClick={() => setElevIdx(i)} style={{ padding: "0.68rem 0.95rem", cursor: "pointer", borderBottom: "1px solid #F1F5F9", background: i === elevIdx ? "#EFF5FC" : "transparent", fontSize: "0.72rem", color: i === elevIdx ? "#0F2138" : "#64748B", fontWeight: i === elevIdx ? 600 : 500, borderLeft: i === elevIdx ? "3px solid #3F79BC" : "3px solid transparent" }}>{e.title || "Page " + e.pageNumber}</div>)}
      </div>
      <div ref={wrapRef} style={{ flex: 1, overflow: "auto", padding: "1.25rem" }}>
        <div style={{ fontSize: "0.64rem", color: "#64748B", marginBottom: "0.65rem" }}>
          {!calib ? "① Set scale first: click 'Scale', click two points a known distance apart, type the feet." : mode === "cut" ? "Cut-out mode: click around a window/door, then Finish — it subtracts." : "Click around the wall, then Finish. SF uses your scale."}<span style={{ color: "#94A3B8" }}> · Ctrl+Z undo · Ctrl+Y redo · Esc cancels the shape · wheel = zoom, drag = pan when zoomed</span>
        </div>
        {img ? <div style={{ display:"inline-block", borderRadius:10, overflow:"hidden", boxShadow:"0 12px 40px -14px rgba(15,23,42,0.30), 0 0 0 1px #E1E8F0" }}><Stage width={stageW} height={stageH} {...zoom.stageProps} onClick={onDown} onTap={onDown} onMouseMove={onHover} onMouseLeave={()=>setHoverSnap(null)}>
          <Layer>
            <KImage image={img} width={stageW} height={stageH} />
            {zoom.hiResNode}
            {shapes.map(s => {
              const flat = s.points.flatMap(([nx, ny]) => [nx * stageW, ny * stageH]);
              const sel = selId === s.id;
              return <Line key={s.id} points={flat} closed fill={s.color + (s.type === "cut" ? "55" : sel ? "66" : "40")}
                stroke={sel ? "#1E293B" : s.color} strokeWidth={sel ? 2.5 : 1.5} dash={s.type === "cut" ? [8, 5] : undefined}
                onClick={() => setSelId(s.id)} onTap={() => setSelId(s.id)} />;
            })}
            {draft.length > 0 && <Line points={draft.flatMap(([nx, ny]) => [nx * stageW, ny * stageH])} stroke="#F59E0B" strokeWidth={2} dash={[6, 4]} />}
            {draft.map(([nx, ny], i) => <Circle key={i} x={nx * stageW} y={ny * stageH} radius={4} fill="#F59E0B" />)}
            {calibPts.map(([nx, ny], i) => <Circle key={"c" + i} x={nx * stageW} y={ny * stageH} radius={5} fill="#10B981" stroke="#fff" strokeWidth={1} />)}
            {calibPts.length === 2 && <Line points={calibPts.flatMap(([nx, ny]) => [nx * stageW, ny * stageH])} stroke="#10B981" strokeWidth={2} />}
            {/* Bluebeam-style snap cue: shows the exact point the next click will lock to */}
            {hoverSnap && (() => { const k = zoom.view?.k || 1; const x = hoverSnap.pt[0] * stageW, y = hoverSnap.pt[1] * stageH, r = 7 / k;
              return <>
                <Circle x={x} y={y} radius={r} stroke={hoverSnap.kind === "corner" ? "#22C55E" : "#F59E0B"} strokeWidth={2 / k} listening={false} />
                <Line points={[x - r * 1.8, y, x + r * 1.8, y]} stroke={hoverSnap.kind === "corner" ? "#22C55E" : "#F59E0B"} strokeWidth={1 / k} listening={false} />
                <Line points={[x, y - r * 1.8, x, y + r * 1.8]} stroke={hoverSnap.kind === "corner" ? "#22C55E" : "#F59E0B"} strokeWidth={1 / k} listening={false} />
              </>; })()}
          </Layer>
        </Stage></div> : <div style={{ color: "#94A3B8", fontSize: "0.82rem", marginTop: "3rem" }}>{elevations.length ? "Loading drawing…" : "Run a drawing in the Takeoff tab first, then edit it here."}</div>}
      </div>
      <div style={{ width: 250, borderLeft: "1px solid #E4EAF1", padding: "1.1rem", background: "#fff", overflowY: "auto", flexShrink: 0 }}>
        <div style={{ display: "flex", gap: "0.35rem", marginBottom: "0.65rem" }}>
          <button style={btn(mode === "add")} onClick={() => setMode("add")}>+ Area</button>
          <button style={btn(mode === "cut")} onClick={() => setMode("cut")}>– Cut-out</button>
          <button style={btn(mode === "calib")} onClick={() => { setMode("calib"); setCalibPts([]); }}>Scale</button>
        </div>
        {mode === "calib" && <div style={{ background: "#F8FAFC", borderRadius: 9, padding: "0.7rem", marginBottom: "0.65rem", border: "1px solid #EEF2F7" }}>
          <div style={{ fontSize: "0.62rem", color: "#64748B", marginBottom: "0.4rem" }}>Click 2 points on a known dimension ({calibPts.length}/2), then enter feet:</div>
          <div style={{ display: "flex", gap: "0.35rem" }}>
            <input value={realFt} onChange={e => setRealFt(e.target.value)} placeholder="feet" style={{ flex: 1, minWidth: 0, background: "#fff", border: "1px solid #D8E2EE", borderRadius: 7, color: "#0F172A", fontSize: "0.72rem", padding: "0.45rem", fontFamily: "inherit" }} />
            <button onClick={applyCalib} style={{ background: "linear-gradient(180deg,#5A92D2,#3F79BC)", border: "none", borderRadius: 7, color: "#fff", fontSize: "0.66rem", fontWeight: 600, padding: "0 0.7rem", cursor: "pointer", fontFamily: "inherit" }}>Set</button>
          </div>
        </div>}
        {draft.length > 0 && <button onClick={finish} style={{ width: "100%", padding: "0.6rem", marginBottom: "0.65rem", background: "linear-gradient(180deg,#22A860,#16874B)", border: "none", borderRadius: 9, color: "#fff", fontSize: "0.73rem", fontWeight: 600, cursor: "pointer", fontFamily: "inherit", boxShadow: "0 6px 16px -5px rgba(21,128,61,0.5)" }}>✓ Finish shape ({draft.length} pts)</button>}
        <div style={{ fontSize: "0.6rem", color: calib ? "#15803D" : "#DC2626", marginBottom: "0.7rem", fontWeight: 600 }}>{calib ? "✓ Scale set" : "⚠ Set scale to get SF"}</div>
        <div style={{ marginBottom: "0.7rem" }}>
          <div style={{ fontSize: "0.6rem", color: "#94A3B8", marginBottom: "0.35rem", textTransform: "uppercase", letterSpacing: "0.08em", fontWeight: 600 }}>Color = material</div>
          <div style={{ display: "flex", gap: "0.35rem", flexWrap: "wrap" }}>
            {COLS.map(c => <div key={c} onClick={() => setCurColor(c)} style={{ width: 24, height: 24, borderRadius: 6, background: c, cursor: "pointer", border: curColor === c ? "2px solid #1E293B" : "2px solid #E4EAF1" }} />)}
          </div>
        </div>
        <div style={{ fontSize: "0.6rem", letterSpacing: "0.08em", color: "#94A3B8", textTransform: "uppercase", fontWeight: 600, marginBottom: "0.45rem" }}>Shapes ({shapes.length})</div>
        {shapes.map(s => {
          const sel = selId === s.id;
          return <div key={s.id} onClick={() => setSelId(s.id)} style={{ padding: "0.55rem 0.65rem", marginBottom: "0.4rem", background: sel ? "#EFF5FC" : "#F8FAFC", borderRadius: 8, cursor: "pointer", borderLeft: "3px solid " + s.color, border: sel ? "1px solid #CFE0F2" : "1px solid #EEF2F7" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "0.4rem" }}>
              <input value={s.name} onClick={e => e.stopPropagation()} onChange={e => rename(s.id, e.target.value)} style={{ flex: 1, minWidth: 0, background: "transparent", border: "none", color: "#334155", fontSize: "0.67rem", fontFamily: "inherit", fontWeight: 500 }} />
              <span style={{ fontSize: "0.73rem", fontWeight: 700, color: s.type === "cut" ? "#DC2626" : "#0F172A" }}>{s.type === "cut" ? "–" : ""}{Math.round(areaSF(s)).toLocaleString()} SF</span>
            </div>
            {sel && <div onClick={e => { e.stopPropagation(); del(s.id); }} style={{ marginTop: "0.4rem", fontSize: "0.6rem", color: "#DC2626", textAlign: "center", border: "1px solid #FECACA", borderRadius: 6, padding: "0.28rem", cursor: "pointer" }}>Delete</div>}
          </div>;
        })}
        {Object.keys(byColor).length > 0 && <>
          <div style={{ fontSize: "0.6rem", letterSpacing: "0.08em", color: "#94A3B8", textTransform: "uppercase", fontWeight: 600, margin: "0.9rem 0 0.45rem" }}>Totals by material</div>
          {Object.entries(byColor).map(([c, d]) => (
            <div key={c} style={{ display: "flex", alignItems: "center", gap: "0.4rem", padding: "0.5rem 0.6rem", marginBottom: "0.35rem", background: "#F8FAFC", borderRadius: 8, borderLeft: "3px solid " + c, border: "1px solid #EEF2F7" }}>
              <div style={{ width: 12, height: 12, borderRadius: 3, background: c, flexShrink: 0 }} />
              <input value={colorNames[c] || ""} onChange={e => setColorNames(p => ({ ...p, [c]: e.target.value }))} placeholder="name this material…" style={{ flex: 1, minWidth: 0, background: "transparent", border: "none", color: "#334155", fontSize: "0.65rem", fontFamily: "inherit" }} />
              <span style={{ fontSize: "0.73rem", fontWeight: 700, color: "#0F172A" }}>{Math.round(d.sf).toLocaleString()} SF</span>
            </div>
          ))}
        </>}
        <div style={{ marginTop: "0.85rem", padding: "0.75rem 0.85rem", background: "linear-gradient(180deg,#F0F9F4,#E3F5EA)", borderRadius: 10, border: "1px solid #BBF0CE" }}>
          <div style={{ fontSize: "0.58rem", color: "#15803D", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em" }}>Net total</div>
          <div style={{ fontSize: "1.4rem", fontWeight: 800, color: "#15803D", letterSpacing: "-0.02em" }}>{Math.round(total).toLocaleString()} <span style={{ fontSize: "0.65rem", fontWeight: 500 }}>SF</span></div>
        </div>
        <div style={{ fontSize: "0.58rem", color: "#94A3B8", marginTop: "0.65rem", lineHeight: 1.5 }}>{taught === "saved" ? "Learned from this drawing." : "The AI learns from your edits automatically — what you change here teaches it for next time."}</div>
      </div>
    </div>
  );
}

/* ── Main App ── */
export default function BFSEstimator() {
  const [file, setFile]       = useState(null);
  const [phase, setPhase]     = useState("idle");
  const [log, setLog]         = useState([]);
  const [progress, setProgress] = useState({ label:"", pct:0 });
  const [results, setResults] = useState(null);
  const [errMsg, setErrMsg]   = useState("");
  const [pdfLoading, setPdfLoading] = useState(false);
  const [viewMode, setViewMode] = useState("table");
  const [dragOver, setDragOver] = useState(false);
  const [appTab, setAppTab] = useState("takeoff");
  const [pricing, setPricing] = useState({ rates:{}, wastePct:15, marginPct:20 });
  // Saved rate cards (per-firm/per-job pricing you reuse) — persisted so you never re-enter rates
  const [rateCards, setRateCards] = useState(()=>{ try{ return JSON.parse(localStorage.getItem("bfs_rate_cards"))||{}; }catch{ return {}; } });
  const [rateCardName, setRateCardName] = useState("");
  const [assignments, setAssignments] = useState({});
  const [groupRename, setGroupRename] = useState({});   // {backendGroupName: estimator name} — propagates across all pages of a job
  const [hiddenIds, setHiddenIds] = useState({});       // deleted highlights — lifted here so they SAVE with the bid
  const [deletedStack, setDeletedStack] = useState([]); // undo history for deletes (persists with the bid too)
  const [bucketShapes, setBucketShapes] = useState([]); // walls the estimator bucket-added (fix AI misses) — lifted so they COUNT in the total, Budget & exports
  const [bucketColorNames, setBucketColorNames] = useState({}); // color -> material name (his Bluebeam habit: cyan IS PNL-1) — lifted, saves with the bid
  const [scopeResult, setScopeResult] = useState(null); // Scope tab's analysis — lifted so it survives tab switches AND pre-checks the Takeoff's materials (specs decide what's ours)

  // ── UI polish: load real fonts + global interactions (the app referenced 'Inter' but never loaded it) ──
  useEffect(() => {
    if (document.getElementById("bfs-ui-polish")) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap";
    document.head.appendChild(link);
    const s = document.createElement("style");
    s.id = "bfs-ui-polish";
    s.textContent = `
      html { font-size: 17.5px; }  /* whole app is rem-based — one dial scales every screen up ~9% */
      html, body, #root { font-family: 'Inter', system-ui, -apple-system, sans-serif; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; text-rendering: optimizeLegibility; letter-spacing: -0.012em; font-variant-numeric: tabular-nums; }
      h1,h2,h3 { letter-spacing: -0.03em; }
      /* the money numbers glow: any big bold figure reads like a dashboard, not a form */
      h2 { font-size: 1.9rem !important; }
      div[style*="font-size: 2.1rem"], div[style*="font-size: 2.2rem"], span[style*="font-size: 2.2rem"] {
        background: linear-gradient(135deg, #5A92D2 0%, #7FB0E0 55%, #AFCDEE 100%);
        -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
        filter: drop-shadow(0 1px 8px rgba(90,146,210,.25));
      }
      /* signature: a slow beam of light traveling the facade line under the header */
      header { position: relative; }
      header::after { content:""; position:absolute; left:0; right:0; bottom:-2px; height:2px; pointer-events:none;
        background: linear-gradient(90deg, transparent 0%, rgba(90,146,210,.9) 30%, #AFCDEE 50%, rgba(90,146,210,.9) 70%, transparent 100%);
        background-size: 220% 100%; animation: bfsShimmer 7s linear infinite; opacity:.85; }
      header img { transition: filter .3s ease, transform .3s ease; }
      header img:hover { filter: drop-shadow(0 0 14px rgba(127,176,224,.55)); transform: scale(1.03); }
      button, [role=button], a { transition: transform .13s cubic-bezier(.2,.8,.2,1), box-shadow .2s ease, background .2s ease, border-color .2s ease, color .2s ease, opacity .2s ease, filter .2s ease; }
      button:not(:disabled):hover, [role=button]:hover { transform: translateY(-1px); box-shadow: 0 4px 14px -4px rgba(15,33,56,.35); }
      button:not(:disabled):active, [role=button]:active { transform: translateY(0) scale(.98); box-shadow: none; }
      /* the drawing view feels alive: selections/hatches ease in, handles breathe */
      svg polygon, svg rect, svg circle { transition: fill-opacity .18s ease, stroke-width .15s ease, stroke .18s ease; }
      svg circle[style*="grab"]:hover { stroke-width: 3px; }
      #root > div { animation: bfsFadeUp .45s cubic-bezier(.2,.8,.2,1); }
      /* interactive cards: anything card-shaped lifts + zooms a touch under the cursor */
      div[style*="border-radius: 12px"], div[style*="border-radius: 10px"] { transition: transform .18s cubic-bezier(.2,.8,.2,1), box-shadow .22s ease; }
      div[style*="border-radius: 12px"]:hover, div[style*="border-radius: 10px"]:hover { transform: translateY(-2px) scale(1.012); box-shadow: 0 10px 30px -12px rgba(15,33,56,.28); }
      /* titles zoom gently when you sweep over them; page titles read light on the navy backdrop */
      h1, h2 { color: #F1F5F9 !important; }
      h1, h2, h3 { transition: transform .2s cubic-bezier(.2,.8,.2,1); transform-origin: left center; }
      h1:hover, h2:hover, h3:hover { transform: scale(1.022); }
      @media (prefers-reduced-motion: reduce) { *, ::after { animation: none !important; transition: none !important; } }
      input, textarea, select { transition: border-color .2s ease, box-shadow .2s ease, background .2s ease; }
      input:focus-visible, textarea:focus-visible, select:focus-visible, button:focus-visible { outline: none; box-shadow: 0 0 0 3px rgba(74,134,200,.35); }
      ::selection { background: rgba(74,134,200,.30); color: inherit; }
      ::-webkit-scrollbar { width: 10px; height: 10px; }
      ::-webkit-scrollbar-track { background: transparent; }
      ::-webkit-scrollbar-thumb { background: rgba(120,150,190,.32); border-radius: 999px; border: 2px solid transparent; background-clip: padding-box; }
      ::-webkit-scrollbar-thumb:hover { background: rgba(120,150,190,.55); background-clip: padding-box; }
      * { scrollbar-width: thin; scrollbar-color: rgba(120,150,190,.4) transparent; }
      @keyframes bfsFadeUp { from { opacity:0; transform: translateY(8px) } to { opacity:1; transform:none } }
      @keyframes bfsShimmer { 100% { background-position: 200% 0 } }
    `;
    document.head.appendChild(s);
  }, []);
  useEffect(()=>{ setGroupRename({}); }, [results?.jobId]);
  const dispName = n => groupRename[n] || n;
  const [savedBids, setSavedBids] = useState([]);
  const fileRef  = useRef();
  const logRef   = useRef();
  const pollRef  = useRef(null);
  const seenLogs = useRef(0);

  useEffect(()=>{ if(logRef.current) logRef.current.scrollTop=logRef.current.scrollHeight; },[log]);
  useEffect(()=>()=>{ if(pollRef.current) clearInterval(pollRef.current); },[]);

  const refreshSaved = useCallback(()=>{
    const out=[];
    try{ for(let i=0;i<localStorage.length;i++){ const k=localStorage.key(i); if(k&&k.startsWith("bfs_bid_")){ try{ out.push(JSON.parse(localStorage.getItem(k))); }catch{} } } }catch{}
    out.sort((a,b)=>(b.savedAt||0)-(a.savedAt||0));
    setSavedBids(out);
  },[]);
  useEffect(()=>{ refreshSaved(); },[refreshSaved]);
  useEffect(()=>{ try{ const p=JSON.parse(localStorage.getItem("bfs_pricing_v1")); if(p&&p.rates) setPricing(p); }catch{} },[]);
  useEffect(()=>{ try{ localStorage.setItem("bfs_pricing_v1", JSON.stringify(pricing)); }catch{} },[pricing]);

  const handleFile = f => {
    if(f?.type==="application/pdf"){
      setFile(f); setPhase("idle"); setResults(null); setLog([]); setErrMsg(""); seenLogs.current=0;
    }
  };

  // Open a finished Queue result in the Takeoff view
  const openResult = (r) => {
    if(!r) return;
    setResults(r); setAssignments(r.assignments||{}); setErrMsg("");
    setHiddenIds(r.hiddenIds||{}); setDeletedStack(r.deletedStack||[]); setBucketShapes(r.bucketShapes||[]);
    setPricing(p=>({...p, sfOverride:{}, customLines:[]}));  // fresh job → fresh per-job numbers
    setPhase("done"); setProgress({ label:"Complete", pct:100 });
    setAppTab("takeoff");
  };

  const startPolling = useCallback((id)=>{
    if(pollRef.current) clearInterval(pollRef.current);
    seenLogs.current=0;
    pollRef.current=setInterval(async()=>{
      try{
        const res=await fetch(BACKEND+"/status/"+id);
        const data=await res.json();
        if(data.log?.length>seenLogs.current){setLog(prev=>[...prev,...data.log.slice(seenLogs.current)]);seenLogs.current=data.log.length;}
        if(data.progress)setProgress(data.progress);
        if(data.phase)setPhase(data.phase);
        if(data.status==="done"){
          clearInterval(pollRef.current);
          setResults({legend:data.legend||[],takeoffData:data.takeoffData||[],scheduleData:data.scheduleData||null,drawingSchedule:data.drawingSchedule||null,ocrMaterials:data.ocrMaterials||null,projName:file?.name?.replace(".pdf","")||"Project",jobId:id});
          setPhase("done");setProgress({label:"Complete",pct:100});
        }else if(data.status==="error"){clearInterval(pollRef.current);setErrMsg(data.error||"Unknown error");setPhase("error");}
      }catch(e){console.log("poll",e.message);}
    },5000);
  },[file]);

  const run = async()=>{
    if(!file)return;
    setPhase("running");setLog([]);setErrMsg("");setResults(null);setAssignments({});setHiddenIds({});setDeletedStack([]);setBucketShapes([]);seenLogs.current=0;
    setPricing(p=>({...p, sfOverride:{}, customLines:[]}));  // per-JOB numbers must never carry into the next bid (rates/waste/margin persist as the working rate card)
    try{
      setLog([{msg:"Uploading PDF...",level:"info"}]);
      const fd=new FormData();fd.append("pdf",file);
      const res=await fetch(BACKEND+"/analyze",{method:"POST",body:fd});
      const{jobId:id}=await res.json();
      setLog(prev=>[...prev,{msg:"Analysis started — job "+id,level:"ok"}]);
      startPolling(id);
    }catch(err){setErrMsg(err.message);setPhase("error");}
  };

  // MONEY GATE: never let a silently-wrong SF into a bid file. Unread scale = up to 4x SF error (it squares).
  const moneyGuard=()=>{
    const warns=[];
    if(defaultScaleN>0) warns.push("• "+defaultScaleN+" page(s) used a DEFAULT scale (couldn't read it) — SF could be far off. Open the elevation and Calibrate first.");
    if(labelWarnN>0) warns.push("• "+labelWarnN+" page(s) have a markup label that looks like a typo (flagged in review).");
    if(missWarnN>0) warns.push("• "+missWarnN+" elevation(s) look UNDER-marked vs the building face — a wall may be missing from the takeoff.");
    return warns.length===0 || window.confirm("⚠ CHECK BEFORE BIDDING\n\n"+warns.join("\n")+"\n\nExport anyway?");
  };
  // THE ANSWER KEY: every export = a finished, human-confirmed takeoff. Save the complete
  // final state as training gold — this is what eventually makes the auto-takeoff fully
  // autonomous (the estimator's job converges to pricing only).
  const captureFinal=(kind)=>{
    if(!results?.jobId) return;
    try{
      fetch(BACKEND+"/learn",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({
        jobId:results.jobId, source:"final-"+kind, projName:results.projName,
        takeoffData:(results.takeoffData||[]).map(e=>({pageNumber:e.pageNumber,source:e.source,scale:e.scale,
          zones:(e.zones||[]).map(z=>({materialName:z.materialName,category:z.category,netArea:z.netArea,grossArea:z.grossArea,totalOpeningArea:z.totalOpeningArea})),
          linearItems:e.linearItems||[]})),
        assignments, groupRename, bucketShapes:(bucketShapes||[]).map(s=>({area_sf:s.area_sf,material:s.material,page:s.page})),
      })}).catch(()=>{});
    }catch{}
  };
  const exportExcel=()=>{
    if(!results)return;
    if(!moneyGuard())return;
    const mt={};
    results.takeoffData.forEach(e=>(e.zones||[]).forEach(z=>{const k=dispName(z.materialName||z.category||"Panel");if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=z.netArea||0;}));
    Object.entries(bucketByMat).forEach(([k,sf])=>{if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=sf;});  // walls the estimator bucket-added
    const wb=buildExcel(results.projName||"Project",Object.values(mt));
    XLSX.writeFile(wb,"BFS_Takeoff_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
    captureFinal("takeoff");
  };

  const exportPDF=async()=>{
    if(!results?.jobId)return;
    setPdfLoading(true);
    try{
      // if she assigned/selected specific materials, mark only those on the PDF; else mark all
      const picked=[...new Set(Object.values(assignments).flatMap(a=>[a.category,a.materialName].filter(Boolean)))];
      const q=picked.length?("?materials="+encodeURIComponent(picked.join(","))):"";
      const res=await fetch(BACKEND+"/evidence-pdf/"+results.jobId+q);
      if(!res.ok)throw new Error("PDF not ready");
      const blob=await res.blob();
      const url=URL.createObjectURL(blob);
      const a=document.createElement("a");a.href=url;a.download="BFS_Evidence_"+(results.projName||"Project").replace(/\s+/g,"_")+".pdf";a.click();
      URL.revokeObjectURL(url);
    }catch(err){alert("PDF export failed: "+err.message);}
    finally{setPdfLoading(false);}
  };

  const saveBid=()=>{
    if(!results)return;
    const id=results.jobId||String(Date.now());
    const rec={ id, projName:results.projName, savedAt:Date.now(),
      data:{ legend:results.legend, takeoffData:results.takeoffData, scheduleData:results.scheduleData||null, drawingSchedule:results.drawingSchedule||null, ocrMaterials:results.ocrMaterials||null, projName:results.projName, jobId:results.jobId },
      assignments, pricing, hiddenIds, deletedStack, groupRename, bucketShapes, bucketColorNames, scopeResult };
    try{ localStorage.setItem("bfs_bid_"+id, JSON.stringify(rec)); refreshSaved(); }
    catch(e){ alert("Could not save bid: "+e.message); }
  };
  const loadBid=rec=>{
    setResults(rec.data); setAssignments(rec.assignments||{});
    setPricing(rec.pricing||{rates:{},wastePct:15,marginPct:20});
    setHiddenIds(rec.hiddenIds||{}); setDeletedStack(rec.deletedStack||[]); setGroupRename(rec.groupRename||{}); setBucketShapes(rec.bucketShapes||[]); setBucketColorNames(rec.bucketColorNames||{});
    if(rec.scopeResult) setScopeResult(rec.scopeResult);
    setPhase("done"); setViewMode("table"); setFile(null);
  };
  const deleteBid=(id,ev)=>{ if(ev)ev.stopPropagation(); try{ localStorage.removeItem("bfs_bid_"+id); }catch{} refreshSaved(); };

  // Walls the estimator bucket-added to fix AI misses — roll up by material so they COUNT
  // in the total, the Budget and the exports (additive: these walls aren't in zones/assignments).
  // name resolution: explicit material > the color's name (set once, applies to ALL shapes
  // of that color — even retroactively) > generic
  const bucketByMat = (bucketShapes||[]).reduce((acc,s)=>{const k=dispName(s.material||bucketColorNames[s.color]||"Cladding (added)");acc[k]=(acc[k]||0)+(s.area_sf||0);return acc;},{});
  const summary=results?()=>{
    const t={};
    results.takeoffData.forEach(e=>(e.zones||[]).forEach(z=>{const k=dispName(z.category||"Other");if(!t[k])t[k]={net:0,adj:0,color:MAT_COLORS[k]||hashColor(k)};t[k].net+=z.netArea||0;t[k].adj+=(z.netArea||0)*1.15;}));
    Object.entries(bucketByMat).forEach(([k,sf])=>{if(!t[k])t[k]={net:0,adj:0,color:MAT_COLORS[k]||hashColor(k)};t[k].net+=sf;t[k].adj+=sf*1.15;});
    return t;
  }:null;
  const summaryData = summary ? summary() : null;
  const grandAdj = summaryData ? Object.values(summaryData).reduce((s,v)=>s+v.adj,0) : 0;
  // SCOPE→TAKEOFF pre-check: match each detected material against the Scope tab's materials
  // list (family match, not spelling). Result = hint badges + "in scope but not detected"
  // reconciliation. Verify-first: informs the estimator, never auto-drops a single SF.
  const scopeCheck = (()=>{
    const mats = scopeResult?.materials; if(!mats||!mats.length||!summaryData) return null;
    const rows = Object.keys(summaryData).map(cat=>{
      const hits = mats.filter(m=>matFamilyMatch(cat, m.name));
      const best = hits.find(m=>m.ours===true) || hits.find(m=>m.ours===false) || hits[0] || null;
      return { cat, verdict: best ? (best.ours ? "ours" : "others") : "unknown", scopeName: best?.name, note: best?.note };
    });
    const missing = mats.filter(m=>m.ours===true && !Object.keys(summaryData).some(cat=>matFamilyMatch(cat, m.name)));
    return { rows, missing };
  })();
  // Reviewed takeoff = what the user assigned in Interactive; drives the bid when present
  const reviewedSummary = Object.values(assignments).reduce((acc,a)=>{
    const cat=dispName(a.category||a.materialName||"Panel"); if(!acc[cat])acc[cat]={net:0}; acc[cat].net+=a.area_sf||0; return acc;
  },{});
  // hasReviewed keys off HER ASSIGNMENTS only (before folding bucket) — so bucket-added walls
  // never flip the pricing source and drop the auto-detected zones from the bid.
  const hasReviewed = Object.keys(reviewedSummary).length>0;
  Object.entries(bucketByMat).forEach(([k,sf])=>{if(!reviewedSummary[k])reviewedSummary[k]={net:0};reviewedSummary[k].net+=sf;});
  const pricingSource = hasReviewed ? reviewedSummary : (summaryData||{});
  const wasteOf = cat => (pricing.wastePerMat && pricing.wastePerMat[cat]!=null) ? pricing.wastePerMat[cat] : pricing.wastePct;  // waste differs by material (lap ~10, shake ~15+, cut panels less)
  const setWasteMat = (cat,v)=>setPricing(p=>({...p,wastePerMat:{...(p.wastePerMat||{}),[cat]:v}}));
  const priceRows = Object.entries(pricingSource).map(([cat,{net:net0}])=>{
    const net = (pricing.sfOverride && pricing.sfOverride[cat]!=null) ? pricing.sfOverride[cat] : net0;  // editable per job
    const rate = pricing.rates[cat]!=null ? pricing.rates[cat] : (DEFAULT_RATES[cat]??DEFAULT_RATES.Other);
    const wastePct = wasteOf(cat);
    const adjSF = net*(1+wastePct/100);
    return { cat, net, adjSF, rate, wastePct, cost:adjSF*rate };
  });
  const costSubtotal = priceRows.reduce((s,r)=>s+r.cost,0);
  const bidTotal = costSubtotal*(1+pricing.marginPct/100);
  const exportPricedExcel=()=>{
    if(!priceRows.length)return;
    if(!moneyGuard())return;
    const mats=priceRows.map(r=>({name:r.cat,sf:r.net*(1+r.wastePct/100)/(1+pricing.wastePct/100),rate:r.rate}));  // per-material waste
    const wb=buildExcel(results.projName||"Project",mats,{wastePct:pricing.wastePct,marginPct:pricing.marginPct});
    XLSX.writeFile(wb,"BFS_Bid_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
    captureFinal("bid");
  };
  // ── Pre-bid reliability checks ──
  const reviewElevs = (results?.takeoffData||[]).filter(e=>(e.zones||[]).some(z=>(z.netArea||0)>0));
  const defaultScaleN = reviewElevs.filter(e=>e.scaleSource==="default" || (!e.verifiedScale && !e.scale)).length;
  const scaleWarnN = reviewElevs.filter(e=>e.expectedFacadeSF && (e.zones||[]).reduce((s,z)=>s+(z.netArea||0),0) > e.expectedFacadeSF*1.4).length;
  const labelWarnN = reviewElevs.filter(e=>(e.flags||[]).some(f=>/implies ≈/.test(f))).length;  // typo'd markup labels caught by the backend
  // UNDER-marking guard — the catastrophic direction (missed wall = missing SF = under-bid)
  const missWarnN = reviewElevs.filter(e=>e.expectedFacadeSF && (e.zones||[]).reduce((s,z)=>s+(z.netArea||0),0) < e.expectedFacadeSF*0.5).length;
  const hasSchedule = !!(results?.scheduleData?.total_opening_sf>0);
  const reviewOk = reviewElevs.length>0 && defaultScaleN===0 && scaleWarnN===0 && labelWarnN===0;
  // Estimator's LINEAR (LF) measurements — trim/soffit/fascia captured from polyline markups
  const linearRollup = (()=>{ const m={}; (results?.takeoffData||[]).forEach(e=>(e.linearItems||[]).forEach(it=>{const k=dispName(it.material||"Linear");if(!m[k])m[k]=0;m[k]+=it.lf||0;})); return Object.entries(m).map(([material,lf])=>({material,lf})).sort((a,b)=>b.lf-a.lf); })();
  // SUGGESTED trim derived from face geometry (backend autoTrim) — verify-first, kept SEPARATE
  // from the estimator's confirmed linear measurements so an unverified LF never enters a bid.
  const autoTrimRollup = (()=>{ const m={}; (results?.takeoffData||[]).forEach(e=>(e.autoTrim||[]).forEach(it=>{const k=it.material||"Trim (auto)";if(!m[k])m[k]=0;m[k]+=it.lf||0;})); return Object.entries(m).map(([material,lf])=>({material,lf})).sort((a,b)=>b.lf-a.lf); })();
  const autoTrimTotalLF = autoTrimRollup.reduce((s,r)=>s+r.lf,0);
  // ── Budget tab: SF (+ LF trim) × the rates the estimator sets per job → live bid + one-click Excel ──
  const setRate = (cat,v)=>setPricing(p=>({...p,rates:{...p.rates,[cat]:v}}));
  const setLfRate = (m,v)=>setPricing(p=>({...p,lfRates:{...(p.lfRates||{}),[m]:v}}));
  const lfRateOf = m => (pricing.lfRates && pricing.lfRates[m]!=null) ? pricing.lfRates[m] : 12;
  const lfRows = linearRollup.map(it=>({...it, rate:lfRateOf(it.material), cost: (it.lf||0)*lfRateOf(it.material)}));
  const lfSubtotal = lfRows.reduce((s,r)=>s+r.cost,0);
  const setSf=(cat,v)=>setPricing(p=>({...p,sfOverride:{...(p.sfOverride||{}),[cat]:v}}));   // override a material's SF for THIS job
  const customLines = pricing.customLines || [];
  const addCustomLine=()=>setPricing(p=>({...p,customLines:[...(p.customLines||[]),{id:Date.now(),name:"",qty:1,rate:0}]}));
  const updCustomLine=(id,k,v)=>setPricing(p=>({...p,customLines:(p.customLines||[]).map(l=>l.id===id?{...l,[k]:v}:l)}));
  const delCustomLine=id=>setPricing(p=>({...p,customLines:(p.customLines||[]).filter(l=>l.id!==id)}));
  const customSubtotal = customLines.reduce((s,l)=>s+(l.qty||0)*(l.rate||0),0);
  const budgetSubtotal = costSubtotal + lfSubtotal + customSubtotal;
  const budgetTotal = budgetSubtotal*(1+pricing.marginPct/100);
  const exportBudgetExcel=()=>{
    if(!results) return;
    if(!moneyGuard())return;
    const w=1+(pricing.wastePct||0)/100;   // buildExcel applies waste to every qty; pre-divide LF/custom so only material SF gets waste
    const mats=[...priceRows.map(r=>({name:r.cat,sf:r.net*(1+r.wastePct/100)/w,rate:r.rate})),   // per-material waste (pre-scaled: buildExcel re-applies the global %)
                ...lfRows.map(r=>({name:r.material+" (per LF)",sf:(r.lf||0)/w,rate:r.rate})),
                ...customLines.filter(l=>(l.name||l.rate)).map(l=>({name:l.name||"Line item",sf:(l.qty||0)/w,rate:l.rate||0}))];
    const wb=buildExcel(results.projName||"Project",mats,{wastePct:pricing.wastePct,marginPct:pricing.marginPct});
    XLSX.writeFile(wb,"BFS_Budget_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
    captureFinal("budget");
  };
  // ── Rate cards: save the rates you set, reuse them on the next job (per firm/job) ──
  const persistRateCards=rc=>{ setRateCards(rc); try{ localStorage.setItem("bfs_rate_cards",JSON.stringify(rc)); }catch{} };
  const saveRateCard=()=>{
    const nm=((rateCardName||results?.projName||"My rates").trim()).slice(0,32)||"My rates";
    persistRateCards({...rateCards,[nm]:{ rates:{...pricing.rates}, wastePct:pricing.wastePct, marginPct:pricing.marginPct, lfRates:{...(pricing.lfRates||{})}, wastePerMat:{...(pricing.wastePerMat||{})}, savedAt:Date.now() }});
    setRateCardName(nm);
  };
  const loadRateCard=nm=>{ const c=rateCards[nm]; if(!c)return; setPricing(p=>({...p, rates:{...c.rates}, wastePct:c.wastePct, marginPct:c.marginPct, lfRates:{...(c.lfRates||{})}, wastePerMat:{...(c.wastePerMat||{})} })); setRateCardName(nm); };
  const deleteRateCard=nm=>{ const rc={...rateCards}; delete rc[nm]; persistRateCards(rc); if(rateCardName===nm)setRateCardName(""); };
  const linearTotalLF = linearRollup.reduce((s,r)=>s+r.lf,0);
  const triage = reviewElevs.map(e=>({ ...elevConfidence(e) }));
  const readyN = triage.filter(t=>t.status==="ready").length;
  const reviewN = triage.filter(t=>t.status==="review").length;
  const attnN = triage.filter(t=>t.status==="attention").length;
  const jobConfidence = triage.length ? Math.round(triage.reduce((s,t)=>s+t.score,0)/triage.length) : 0;
  const phaseStep={idle:0,running:1,filtering:1,legend:2,analyzing:3,done:4,error:0}[phase]||0;
  const isRunning=!["idle","done","error"].includes(phase);
  const logColor={ok:"#22C55E",warn:"#F59E0B",error:"#EF4444",success:"#22C55E",dim:"#94A3B8",info:"#64748B"};

  const showResults = phase==="done" && results;
  const showUploadScreen = !showResults;

  return (
    <div style={{fontFamily:"'Inter',system-ui,-apple-system,sans-serif",background:"radial-gradient(1100px 520px at 15% -10%, rgba(74,134,200,0.16), transparent 60%), radial-gradient(900px 460px at 85% -8%, rgba(90,146,210,0.13), transparent 58%), radial-gradient(1500px 800px at 50% 118%, rgba(63,121,188,0.10), transparent 55%), #0C1B2E",minHeight:"100vh",display:"flex",flexDirection:"column",color:"#1E293B"}}>

      {/* ── Header ── */}
      <header style={{background:"linear-gradient(180deg,#0F2138,#0B1728)",height:96,padding:"0 2.4rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0,boxShadow:"0 1px 0 rgba(255,255,255,0.06), 0 6px 28px -8px rgba(0,0,0,0.55)",zIndex:10}}>
        <div style={{flex:1,minWidth:0,display:"flex",alignItems:"center",gap:"1.25rem"}}>
          <img src="/logo-bfs.png" alt="BFS" style={{height:66,width:"auto"}}/>
          <div style={{width:1,height:54,background:"rgba(255,255,255,0.13)"}}/>
          <div>
            <div style={{fontSize:"0.64rem",letterSpacing:"0.27em",color:"rgba(255,255,255,0.46)",textTransform:"uppercase",fontWeight:600,marginBottom:2}}>Boston Facade Systems</div>
            <div style={{fontSize:"1.4rem",fontWeight:600,color:"#fff",letterSpacing:"-0.025em",fontFamily:"'Space Grotesk',sans-serif"}}>AI Estimator</div>
          </div>
        </div>
        {/* Top-level nav tabs — front and center */}
        <div style={{display:"flex",gap:"0.25rem",background:"rgba(255,255,255,0.07)",borderRadius:11,padding:"0.28rem",border:"1px solid rgba(255,255,255,0.09)",boxShadow:"inset 0 1px 4px rgba(0,0,0,0.35)"}}>
          {[["takeoff","Takeoff"],["queue","Queue"],["manual","Draw"],["scope","Scope"],["budget","Budget"],["model","Model"]].map(([t,label])=>(
            <button key={t} onClick={()=>setAppTab(t)} style={{padding:"0.5rem 1.15rem",borderRadius:8,border:"none",fontSize:"0.78rem",fontWeight:appTab===t?700:600,fontFamily:"inherit",cursor:"pointer",background:appTab===t?"linear-gradient(180deg,#5A92D2,#3F79BC)":"transparent",color:appTab===t?"#fff":"rgba(255,255,255,0.55)",boxShadow:appTab===t?"0 2px 12px rgba(74,134,200,0.5)":"none",letterSpacing:"-0.01em",transition:"all 0.15s"}}>{label}</button>
          ))}
        </div>
        <div style={{flex:1,minWidth:0,display:"flex",alignItems:"center",justifyContent:"flex-end",gap:"1rem"}}>
          {appTab==="takeoff"&&showResults&&(
            <div style={{display:"flex",gap:"0.5rem"}}>
              <button onClick={exportExcel} style={{padding:"0.45rem 1rem",background:"transparent",color:"rgba(255,255,255,0.7)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↓ Excel</button>
              <button onClick={exportPDF} disabled={pdfLoading} style={{padding:"0.45rem 1rem",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",color:"#fff",border:"none",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↓ {pdfLoading?"Generating...":"Evidence PDF"}</button>
              <button onClick={saveBid} style={{padding:"0.45rem 1rem",background:"transparent",color:"rgba(255,255,255,0.7)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>💾 Save</button>
              <button onClick={()=>{setFile(null);setPhase("idle");setResults(null);setLog([]);setAssignments({});setHiddenIds({});setDeletedStack([]);setBucketShapes([]);}} style={{padding:"0.45rem 1rem",background:"rgba(255,255,255,0.08)",color:"rgba(255,255,255,0.6)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↺ New</button>
            </div>
          )}
        </div>
      </header>

      {appTab==="scope"&&<ScopeView result={scopeResult} setResult={setScopeResult}/>}
      {appTab==="model"&&<ModelView/>}
      {appTab==="budget"&&(
        <div style={{flex:1,overflowY:"auto",padding:"2rem"}}>
          <div style={{maxWidth:840,margin:"0 auto"}}>
            <div style={{fontSize:"0.7rem",letterSpacing:"0.18em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"0.3rem"}}>Budget</div>
            <h2 style={{fontSize:"1.5rem",fontWeight:800,color:"#0F172A",margin:"0 0 0.3rem",letterSpacing:"-0.02em"}}>Price the bid — no Excel needed</h2>
            <p style={{fontSize:"0.82rem",color:"#9FB3CC",margin:"0 0 1.5rem",lineHeight:1.6}}>Your takeoff SF × the rates you set for <i>this</i> job (you don't always charge the same). Edit a rate and the total updates live. Export writes your BFS estimate sheet automatically.</p>
            {(!results||!priceRows.length)?(
              <div style={{padding:"2.5rem",textAlign:"center",color:"#94A3B8",background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",fontSize:"0.85rem"}}>Run a takeoff and tag your materials first — then set your prices here.</div>
            ):(<>
              <div style={{display:"flex",gap:"0.4rem",alignItems:"center",flexWrap:"wrap",marginBottom:"1rem",padding:"0.6rem 0.8rem",background:"#fff",borderRadius:10,border:"1px solid #EEF2F7"}}>
                <span style={{fontSize:"0.64rem",color:"#64748B",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.05em"}}>Rate cards</span>
                {Object.keys(rateCards).length===0&&<span style={{fontSize:"0.66rem",color:"#94A3B8"}}>none yet — set your rates below and save</span>}
                {Object.keys(rateCards).map(nm=>(
                  <span key={nm} style={{display:"inline-flex",alignItems:"center",gap:"0.35rem",padding:"0.25rem 0.55rem",borderRadius:16,fontSize:"0.66rem",fontWeight:600,background:rateCardName===nm?BLUE:"#F1F5F9",color:rateCardName===nm?"#fff":"#475569",border:"1px solid "+(rateCardName===nm?BLUE:"#E2E8F0")}}>
                    <span onClick={()=>loadRateCard(nm)} style={{cursor:"pointer"}}>{nm}</span>
                    <span onClick={()=>deleteRateCard(nm)} style={{cursor:"pointer",opacity:0.55}} title="delete">×</span>
                  </span>
                ))}
                <div style={{flex:1,minWidth:8}}/>
                <input value={rateCardName} onChange={e=>setRateCardName(e.target.value)} placeholder="name (e.g. Windover)" style={{width:132,padding:"0.28rem 0.5rem",borderRadius:5,border:"1px solid #CBD5E1",fontSize:"0.68rem",fontFamily:"inherit"}}/>
                <button onClick={saveRateCard} style={{padding:"0.3rem 0.7rem",borderRadius:6,border:"none",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",color:"#fff",fontSize:"0.68rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>💾 Save rates</button>
              </div>
              <div style={{background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",overflow:"hidden",marginBottom:"1rem"}}>
                <div style={{display:"grid",gridTemplateColumns:"1.7fr 0.9fr 0.9fr 1fr 1fr",padding:"0.6rem 1rem",background:"#F8FAFC",fontSize:"0.58rem",fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:"0.06em"}}>
                  <div>Material</div><div style={{textAlign:"right"}}>Net SF</div><div style={{textAlign:"right"}}>+Waste</div><div style={{textAlign:"right"}}>Rate $/SF</div><div style={{textAlign:"right"}}>Cost</div>
                </div>
                {priceRows.map(r=>(
                  <div key={r.cat} style={{display:"grid",gridTemplateColumns:"1.7fr 0.9fr 0.9fr 1fr 1fr",padding:"0.5rem 1rem",borderTop:"1px solid #F1F5F9",alignItems:"center",fontSize:"0.76rem"}}>
                    <div style={{fontWeight:600,color:"#0F172A"}}>{r.cat}</div>
                    <div style={{textAlign:"right"}}><input type="number" value={Math.round(r.net)} onChange={e=>setSf(r.cat,parseFloat(e.target.value)||0)} style={{width:64,textAlign:"right",padding:"0.22rem 0.35rem",borderRadius:5,border:"1px solid #E2E8F0",fontSize:"0.74rem",fontFamily:"inherit",color:"#334155"}}/></div>
                    <div style={{textAlign:"right",whiteSpace:"nowrap"}}><input type="number" value={r.wastePct} onChange={e=>setWasteMat(r.cat,parseFloat(e.target.value)||0)} title="waste % for this material" style={{width:36,textAlign:"right",padding:"0.22rem 0.25rem",borderRadius:5,border:"1px solid #E2E8F0",fontSize:"0.7rem",fontFamily:"inherit",color:"#64748B"}}/><span style={{fontSize:"0.6rem",color:"#94A3B8"}}>% → </span><span style={{color:"#94A3B8"}}>{Math.round(r.adjSF).toLocaleString()}</span></div>
                    <div style={{textAlign:"right",whiteSpace:"nowrap"}}><span style={{color:"#94A3B8"}}>$</span><input type="number" value={r.rate} onChange={e=>setRate(r.cat,parseFloat(e.target.value)||0)} style={{width:60,textAlign:"right",padding:"0.22rem 0.35rem",borderRadius:5,border:"1px solid #CBD5E1",fontSize:"0.74rem",fontFamily:"inherit"}}/></div>
                    <div style={{textAlign:"right",fontWeight:700,color:"#0F172A"}}>${Math.round(r.cost).toLocaleString()}</div>
                  </div>
                ))}
              </div>
              {lfRows.length>0&&<div style={{background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",overflow:"hidden",marginBottom:"1rem"}}>
                <div style={{display:"grid",gridTemplateColumns:"1.7fr 0.9fr 0.9fr 1fr 1fr",padding:"0.6rem 1rem",background:"#F8FAFC",fontSize:"0.58rem",fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:"0.06em"}}>
                  <div>Trim / Linear</div><div style={{textAlign:"right"}}>LF</div><div/><div style={{textAlign:"right"}}>Rate $/LF</div><div style={{textAlign:"right"}}>Cost</div>
                </div>
                {lfRows.map(r=>(
                  <div key={r.material} style={{display:"grid",gridTemplateColumns:"1.7fr 0.9fr 0.9fr 1fr 1fr",padding:"0.5rem 1rem",borderTop:"1px solid #F1F5F9",alignItems:"center",fontSize:"0.76rem"}}>
                    <div style={{fontWeight:600,color:"#0F172A"}}>{r.material}</div>
                    <div style={{textAlign:"right",color:"#64748B"}}>{Math.round(r.lf).toLocaleString()}</div>
                    <div/>
                    <div style={{textAlign:"right",whiteSpace:"nowrap"}}><span style={{color:"#94A3B8"}}>$</span><input type="number" value={r.rate} onChange={e=>setLfRate(r.material,parseFloat(e.target.value)||0)} style={{width:60,textAlign:"right",padding:"0.22rem 0.35rem",borderRadius:5,border:"1px solid #CBD5E1",fontSize:"0.74rem",fontFamily:"inherit"}}/></div>
                    <div style={{textAlign:"right",fontWeight:700,color:"#0F172A"}}>${Math.round(r.cost).toLocaleString()}</div>
                  </div>
                ))}
              </div>}
              <div style={{background:"#fff",borderRadius:12,border:"1px solid #EEF2F7",overflow:"hidden",marginBottom:"1rem"}}>
                <div style={{display:"grid",gridTemplateColumns:"1.7fr 0.9fr 0.9fr 1fr 1fr",padding:"0.6rem 1rem",background:"#F8FAFC",fontSize:"0.58rem",fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:"0.06em"}}>
                  <div>Custom / adders</div><div style={{textAlign:"right"}}>Qty</div><div/><div style={{textAlign:"right"}}>$ / unit</div><div style={{textAlign:"right"}}>Cost</div>
                </div>
                {customLines.map(l=>(
                  <div key={l.id} style={{display:"grid",gridTemplateColumns:"1.7fr 0.9fr 0.9fr 1fr 1fr",padding:"0.45rem 1rem",borderTop:"1px solid #F1F5F9",alignItems:"center",fontSize:"0.76rem",gap:"0.3rem"}}>
                    <input value={l.name} onChange={e=>updCustomLine(l.id,"name",e.target.value)} placeholder="e.g. Scaffolding, mobilization…" style={{padding:"0.22rem 0.4rem",borderRadius:5,border:"1px solid #E2E8F0",fontSize:"0.73rem",fontFamily:"inherit",minWidth:0}}/>
                    <input type="number" value={l.qty} onChange={e=>updCustomLine(l.id,"qty",parseFloat(e.target.value)||0)} style={{textAlign:"right",padding:"0.22rem 0.35rem",borderRadius:5,border:"1px solid #E2E8F0",fontSize:"0.74rem",fontFamily:"inherit",minWidth:0}}/>
                    <div/>
                    <div style={{textAlign:"right",whiteSpace:"nowrap"}}><span style={{color:"#94A3B8"}}>$</span><input type="number" value={l.rate} onChange={e=>updCustomLine(l.id,"rate",parseFloat(e.target.value)||0)} style={{width:58,textAlign:"right",padding:"0.22rem 0.35rem",borderRadius:5,border:"1px solid #E2E8F0",fontSize:"0.74rem",fontFamily:"inherit"}}/></div>
                    <div style={{textAlign:"right",fontWeight:700,color:"#0F172A",display:"flex",justifyContent:"flex-end",alignItems:"center",gap:"0.4rem"}}>${Math.round((l.qty||0)*(l.rate||0)).toLocaleString()}<span onClick={()=>delCustomLine(l.id)} title="remove" style={{cursor:"pointer",color:"#F87171",fontWeight:700}}>×</span></div>
                  </div>
                ))}
                <div onClick={addCustomLine} style={{padding:"0.5rem 1rem",borderTop:"1px solid #F1F5F9",fontSize:"0.73rem",color:BLUE,fontWeight:700,cursor:"pointer"}}>+ Add line item (lump sum, adder, extra scope…)</div>
              </div>
              <div style={{display:"flex",gap:"0.75rem",marginBottom:"1rem"}}>
                {[["Waste %","wastePct"],["Margin %","marginPct"]].map(([lab,key])=>(
                  <div key={key} style={{flex:1,background:"#fff",borderRadius:10,border:"1px solid #EEF2F7",padding:"0.6rem 0.85rem",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <span style={{fontSize:"0.72rem",color:"#64748B",fontWeight:600}}>{lab}</span>
                    <input type="number" value={pricing[key]} onChange={e=>setPricing(p=>({...p,[key]:parseFloat(e.target.value)||0}))} style={{width:56,textAlign:"right",padding:"0.25rem 0.4rem",borderRadius:5,border:"1px solid #CBD5E1",fontSize:"0.78rem",fontFamily:"inherit"}}/>
                  </div>
                ))}
              </div>
              <div style={{background:NAVY,borderRadius:12,padding:"1.15rem 1.35rem",color:"#fff"}}>
                {[["Materials",costSubtotal],...(lfSubtotal>0?[["Trim / linear",lfSubtotal]]:[]),...(customSubtotal>0?[["Custom / adders",customSubtotal]]:[]),["Subtotal",budgetSubtotal],["Margin ("+pricing.marginPct+"%)",budgetSubtotal*pricing.marginPct/100]].map(([l,v])=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",fontSize:"0.76rem",color:"rgba(255,255,255,0.7)",marginBottom:"0.35rem"}}><span>{l}</span><span>${Math.round(v).toLocaleString()}</span></div>
                ))}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",borderTop:"1px solid rgba(255,255,255,0.12)",paddingTop:"0.6rem",marginTop:"0.4rem"}}>
                  <span style={{fontSize:"0.7rem",fontWeight:700,color:"#4ADE80",letterSpacing:"0.08em"}}>BID TOTAL</span>
                  <span style={{fontSize:"1.7rem",fontWeight:800,color:"#4ADE80"}}>${Math.round(budgetTotal).toLocaleString()}</span>
                </div>
                <button onClick={exportBudgetExcel} style={{width:"100%",marginTop:"0.9rem",padding:"0.7rem",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",color:"#fff",border:"none",borderRadius:8,fontSize:"0.78rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓ Export / update the Excel</button>
              </div>
            </>)}
          </div>
        </div>
      )}
      {appTab==="queue"&&<QueueView onOpen={openResult}/>}
      {appTab==="manual"&&<ManualView results={results} BACKEND={BACKEND}/>}

      {appTab==="takeoff"&&(<>

      {/* ══ UPLOAD SCREEN ══ */}
      {showUploadScreen&&(
        <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"2rem",position:"relative",overflow:"hidden"}}>
          {/* Background grid */}
          <div style={{position:"absolute",inset:0,backgroundImage:`linear-gradient(rgba(74,134,200,0.05) 1px, transparent 1px), linear-gradient(90deg, rgba(74,134,200,0.05) 1px, transparent 1px)`,backgroundSize:"44px 44px",pointerEvents:"none",maskImage:"radial-gradient(ellipse 80% 60% at 50% 40%, #000 40%, transparent 100%)",WebkitMaskImage:"radial-gradient(ellipse 80% 60% at 50% 40%, #000 40%, transparent 100%)"}}/>
          {/* Soft blue aura */}
          <div style={{position:"absolute",top:"18%",left:"50%",transform:"translateX(-50%)",width:760,height:520,background:"radial-gradient(ellipse at center, rgba(74,134,200,0.28), rgba(74,134,200,0.06) 45%, transparent 70%)",filter:"blur(30px)",pointerEvents:"none"}}/>

          {/* Center card */}
          <div style={{position:"relative",width:"100%",maxWidth:540,display:"flex",flexDirection:"column",gap:"2rem"}}>

            {/* Hero text */}
            <div style={{textAlign:"center"}}>
              <div style={{fontSize:"0.68rem",letterSpacing:"0.28em",color:"#7FB0E0",textTransform:"uppercase",fontWeight:700,marginBottom:"0.9rem"}}>Boston Facade Systems</div>
              <h1 style={{fontSize:"2.7rem",fontWeight:700,margin:0,letterSpacing:"-0.035em",lineHeight:1.04,fontFamily:"'Space Grotesk',sans-serif",background:"linear-gradient(180deg,#ffffff 30%,#AFCDEE)",WebkitBackgroundClip:"text",backgroundClip:"text",WebkitTextFillColor:"transparent"}}>AI Estimator</h1>
              <p style={{fontSize:"0.95rem",color:"rgba(255,255,255,0.6)",marginTop:"0.9rem",lineHeight:1.6}}>Upload your blueprint PDF and get a full material takeoff<br/>with SF breakdown by elevation in seconds.</p>
              <div style={{display:"flex",justifyContent:"center",gap:"1.1rem",marginTop:"1rem",flexWrap:"wrap"}}>
                {["Measured from drawing geometry","Verified scale","Bid-ready Excel"].map(t=>(
                  <div key={t} style={{display:"flex",alignItems:"center",gap:"0.35rem",fontSize:"0.66rem",color:"rgba(255,255,255,0.5)"}}>
                    <span style={{color:"#4ADE80",fontWeight:700}}>✓</span>{t}
                  </div>
                ))}
              </div>
            </div>

            {/* Drop zone */}
            <div
              onClick={()=>!isRunning&&fileRef.current?.click()}
              onDrop={e=>{e.preventDefault();setDragOver(false);if(!isRunning)handleFile(e.dataTransfer.files[0]);}}
              onDragOver={e=>{e.preventDefault();setDragOver(true);}}
              onDragLeave={()=>setDragOver(false)}
              style={{
                border:`1.5px dashed ${dragOver?BLUE:file?"#22C55E":"rgba(127,176,224,0.45)"}`,
                borderRadius:18,
                padding:"2.75rem 2rem",
                textAlign:"center",
                cursor:isRunning?"default":"pointer",
                background:dragOver?"rgba(74,134,200,0.12)":file?"rgba(34,197,94,0.06)":"rgba(255,255,255,0.04)",
                backdropFilter:"blur(10px)",
                transition:"all 0.22s cubic-bezier(.2,.8,.2,1)",
                position:"relative",
                boxShadow:dragOver?"0 20px 60px -12px rgba(74,134,200,0.55), inset 0 1px 0 rgba(255,255,255,0.08)":"0 12px 40px -16px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.06)",
                transform:dragOver?"scale(1.015)":"scale(1)",
              }}>
              {!file&&!isRunning&&(
                <>
                  <div style={{fontSize:"3rem",marginBottom:"0.75rem",opacity:0.5}}>📂</div>
                  <div style={{fontSize:"1rem",fontWeight:600,color:"rgba(255,255,255,0.75)",marginBottom:"0.35rem"}}>Drop your blueprint PDF here</div>
                  <div style={{fontSize:"0.78rem",color:"rgba(255,255,255,0.3)"}}>or click to browse your files</div>
                  <div style={{fontSize:"0.68rem",color:"rgba(127,176,224,0.55)",marginTop:"0.6rem"}}>Upload the full drawing set — the system finds the elevations. Marked-up Bluebeam sets read exact.</div>
                </>
              )}
              {file&&!isRunning&&(
                <>
                  <div style={{fontSize:"2.5rem",marginBottom:"0.5rem"}}>📋</div>
                  <div style={{fontSize:"1rem",fontWeight:700,color:"#4ADE80",marginBottom:"0.25rem",wordBreak:"break-all"}}>{file.name}</div>
                  <div style={{fontSize:"0.78rem",color:"rgba(74,222,128,0.6)"}}>{(file.size/1e6).toFixed(1)} MB · Ready to analyze</div>
                  <div style={{fontSize:"0.72rem",color:"rgba(255,255,255,0.2)",marginTop:"0.5rem"}}>Click to change file</div>
                </>
              )}
              {isRunning&&(
                <div style={{padding:"0.5rem 0"}}>
                  <LiveNarration/>
                  <div style={{fontSize:"0.72rem",color:"rgba(255,255,255,0.4)",marginBottom:"0.75rem",letterSpacing:"0.05em"}}>{progress.label||"Analyzing..."}</div>
                  <div style={{background:"rgba(255,255,255,0.08)",borderRadius:8,height:6,overflow:"hidden",marginBottom:"0.75rem"}}>
                    <div style={{width:(progress.pct||0)+"%",height:"100%",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",borderRadius:8,transition:"width 0.4s"}}/>
                  </div>
                  {/* Steps */}
                  <div style={{display:"flex",justifyContent:"center",gap:"1.5rem"}}>
                    {[{n:1,label:"Index"},{n:2,label:"Legend"},{n:3,label:"Elevations"},{n:4,label:"Export"}].map(({n,label})=>{
                      const done=phaseStep>n,active=phaseStep===n;
                      return <div key={n} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"0.3rem"}}>
                        <div style={{width:28,height:28,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.65rem",fontWeight:700,background:done?"#22C55E":active?BLUE:"rgba(255,255,255,0.08)",color:done||active?"#fff":"rgba(255,255,255,0.2)",boxShadow:active?`0 0 0 3px ${BLUE}30`:"none",transition:"all 0.3s"}}>{done?"✓":n}</div>
                        <div style={{fontSize:"0.58rem",color:done?"#4ADE80":active?"rgba(255,255,255,0.7)":"rgba(255,255,255,0.2)",fontWeight:active?600:400}}>{label}</div>
                      </div>;
                    })}
                  </div>
                </div>
              )}
            </div>

            {/* CTA Button */}
            {!isRunning&&(
              <button onClick={run} disabled={!file}
                style={{padding:"0.95rem",borderRadius:12,fontSize:"0.85rem",fontWeight:600,fontFamily:"inherit",border:"none",cursor:file?"pointer":"not-allowed",letterSpacing:"-0.01em",background:file?"linear-gradient(180deg,#5A92D2,#3F79BC)":"rgba(255,255,255,0.06)",color:file?"#fff":"rgba(255,255,255,0.2)",boxShadow:file?"0 10px 30px -6px rgba(74,134,200,0.6)":"none"}}>
                {phase==="error"?"↺  Try Again":"▶  Run Analysis"}
              </button>
            )}

            {phase==="error"&&errMsg&&(
              <div style={{padding:"0.75rem 1rem",background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.3)",borderRadius:8,fontSize:"0.72rem",color:"#FCA5A5",textAlign:"center"}}>⚠ {errMsg}</div>
            )}

            {/* Log */}
            {log.length>0&&(
              <div style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:10,padding:"0.75rem 1rem"}}>
                <div style={{fontSize:"0.58rem",letterSpacing:"0.12em",color:"rgba(255,255,255,0.2)",textTransform:"uppercase",fontWeight:600,marginBottom:"0.4rem"}}>Activity Log</div>
                <div ref={logRef} style={{fontFamily:"'Courier New',monospace",fontSize:"0.68rem",maxHeight:100,overflowY:"auto",lineHeight:1.8}}>
                  {log.map((l,i)=><div key={i} style={{color:logColor[l.level]||"#64748B"}}>{l.msg}</div>)}
                </div>
              </div>
            )}

            {/* What happens to your file — the pipeline, in the estimator's language */}
            {!file&&!isRunning&&(
              <div style={{background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"1.1rem 1.25rem"}}>
                <div style={{fontSize:"0.6rem",letterSpacing:"0.14em",color:"rgba(127,176,224,0.8)",textTransform:"uppercase",fontWeight:700,marginBottom:"0.8rem"}}>Drop the full set — here's what happens</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"0.7rem 1rem"}}>
                  {[
                    ["1","Finds the elevations","Scans every sheet of the full set — plans, details and schedules are passed over, elevation views are pulled out"],
                    ["2","Reads the drawing itself","Panel seams, lap courses, hatch patterns, wall fills, structural joints and material tags (MT-5, EIFS-3…) — measured from the drawing's own geometry, never guessed"],
                    ["3","Marks it up like an estimator","Each wall arrives as its own highlighted piece with windows cut out and its SF on a label — the same way a hand takeoff looks"],
                    ["4","Checks its own numbers","Scale verified three ways: title block, dimension strings, elevation markers. If they disagree, it tells you to calibrate — it never ships a confident wrong number"],
                    ["5","You select scope & price","Bucket-click the materials that are yours, set rates in Budget — exact SF flows to the BFS Excel and the evidence PDF"],
                    ["6","Every correction teaches it","Renames, deletes, splits and your final export all train the system on your standards — it reads the next set better than the last"],
                  ].map(([n,t,d])=>(
                    <div key={n} style={{display:"flex",gap:"0.6rem",alignItems:"flex-start"}}>
                      <div style={{width:22,height:22,borderRadius:"50%",background:"linear-gradient(135deg,#3F79BC,#5A92D2)",color:"#fff",fontSize:"0.68rem",fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,marginTop:1}}>{n}</div>
                      <div>
                        <div style={{fontSize:"0.74rem",fontWeight:700,color:"rgba(255,255,255,0.88)"}}>{t}</div>
                        <div style={{fontSize:"0.62rem",color:"rgba(255,255,255,0.42)",lineHeight:1.5,marginTop:2}}>{d}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Saved bids */}
            {savedBids.length>0&&!isRunning&&(
              <div style={{background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:10,padding:"0.75rem 1rem"}}>
                <div style={{fontSize:"0.58rem",letterSpacing:"0.12em",color:"rgba(255,255,255,0.3)",textTransform:"uppercase",fontWeight:600,marginBottom:"0.5rem"}}>Saved bids</div>
                <div style={{display:"flex",flexDirection:"column",gap:"0.3rem"}}>
                  {savedBids.slice(0,6).map(b=>(
                    <div key={b.id} onClick={()=>loadBid(b)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0.4rem 0.6rem",borderRadius:6,background:"rgba(255,255,255,0.04)",cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background="rgba(74,134,200,0.12)"} onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.04)"}>
                      <span style={{fontSize:"0.72rem",color:"rgba(255,255,255,0.8)",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{b.projName||"Project"}</span>
                      <span style={{display:"flex",alignItems:"center",gap:"0.6rem",flexShrink:0}}>
                        <span style={{fontSize:"0.62rem",color:"rgba(255,255,255,0.3)"}}>{new Date(b.savedAt).toLocaleDateString()}</span>
                        <span onClick={e=>deleteBid(b.id,e)} style={{fontSize:"0.85rem",color:"rgba(255,255,255,0.3)",cursor:"pointer"}}>×</span>
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ══ RESULTS SCREEN ══ */}
      {showResults&&(
        <div style={{display:"flex",flex:1,overflow:"hidden"}}>

          {/* Sidebar */}
          <aside style={{width:260,background:"#fff",borderRight:"1px solid #E2E8F0",padding:"1.25rem",display:"flex",flexDirection:"column",gap:"1.25rem",overflowY:"auto",flexShrink:0}}>

            {/* File info */}
            <div style={{padding:"0.75rem",background:BLUE_PALE,borderRadius:8,border:"1px solid "+BLUE+"25"}}>
              <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.3rem"}}>Project</div>
              <div style={{fontSize:"0.8rem",fontWeight:600,color:"#0F172A",wordBreak:"break-all"}}>{results.projName}</div>
            </div>

            {/* Pre-bid reliability check */}
            {reviewElevs.length>0&&(
              <div style={{padding:"0.75rem",background:reviewOk?"#F0FDF4":"#FFFBEB",borderRadius:8,border:"1px solid "+(reviewOk?"#BBF7D0":"#FDE68A")}}>
                <div style={{fontSize:"0.6rem",color:reviewOk?"#15803D":"#B45309",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.45rem"}}>{reviewOk?"✓ Checks passed":"⚠ Check before bidding"}</div>
                <div style={{display:"flex",flexDirection:"column",gap:"0.32rem",fontSize:"0.64rem",color:"#475569",lineHeight:1.4}}>
                  <div>{reviewElevs.length} elevation{reviewElevs.length!==1?"s":""} measured</div>
                  {defaultScaleN>0
                    ? <div style={{color:"#B45309"}}>⚠ {defaultScaleN} used a default scale — open Interactive → Calibrate to confirm</div>
                    : <div style={{color:"#15803D"}}>✓ Scale read on every elevation</div>}
                  {scaleWarnN>0&&<div style={{color:"#B45309"}}>⚠ {scaleWarnN} measured bigger than the building face — likely a scale error</div>}
                  {labelWarnN>0
                    ? <div style={{color:"#B45309"}}>⚠ {labelWarnN} elevation{labelWarnN!==1?"s have":" has"} a region whose SF label doesn't match the drawing — likely a typo, verify below</div>
                    : <div style={{color:"#15803D"}}>✓ Every SF label matches its region on the sheet</div>}
                  {hasSchedule
                    ? <div style={{color:"#15803D"}}>✓ Window/door openings exact (from schedule)</div>
                    : <div style={{color:"#92855B"}}>• Openings estimated (no schedule found)</div>}
                  <div style={{color:"#64748B"}}>• Review shapes in Edit Surfaces before exporting</div>
                </div>
              </div>
            )}

            {/* Summary stats */}
            {summaryData&&Object.keys(summaryData).length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.65rem"}}>Materials</div>
                <div style={{display:"flex",flexDirection:"column",gap:"0.4rem"}}>
                  {Object.entries(summaryData).map(([cat,{adj,color}])=>(
                    <div key={cat} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0.45rem 0.6rem",borderRadius:6,background:"#F8FAFC",borderLeft:"3px solid "+color}}>
                      <span style={{fontSize:"0.68rem",color:"#475569",fontWeight:500}}>{cat}</span>
                      <span style={{fontSize:"0.75rem",fontWeight:700,color:"#0F172A"}}>{Math.round(adj).toLocaleString()}</span>
                    </div>
                  ))}
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0.5rem 0.6rem",borderRadius:6,background:BLUE_PALE,borderLeft:"3px solid "+BLUE,marginTop:"0.25rem"}}>
                    <span style={{fontSize:"0.68rem",color:BLUE_DARK,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.05em"}}>Total</span>
                    <span style={{fontSize:"0.85rem",fontWeight:800,color:BLUE}}>{Math.round(grandAdj).toLocaleString()} SF</span>
                  </div>
                </div>
              </div>
            )}

            {/* SCOPE CHECK — the Scope tab's conclusions pre-check the detected materials */}
            {scopeCheck&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>🔗 Scope check</div>
                <div style={{border:"1px solid #DDE7F2",borderRadius:8,overflow:"hidden"}}>
                  {scopeCheck.rows.map((r,i)=>{
                    const badge = r.verdict==="ours" ? {t:"IN SCOPE",bg:"#DCFCE7",fg:"#15803D"} : r.verdict==="others" ? {t:"BY OTHERS",bg:"#FEE2E2",fg:"#B91C1C"} : {t:"NOT IN SCOPE DOC",bg:"#F1F5F9",fg:"#64748B"};
                    return (
                      <div key={i} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.42rem 0.6rem",background:i%2?"#F8FAFC":"#fff",borderTop:i?"1px solid #F1F5F9":"none"}} title={r.scopeName?`Scope: ${r.scopeName}${r.note?" — "+r.note:""}`:""}>
                        <span style={{fontSize:"0.64rem",color:"#334155",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.cat}</span>
                        <span style={{fontSize:"0.52rem",fontWeight:800,padding:"0.14rem 0.4rem",borderRadius:4,background:badge.bg,color:badge.fg,flexShrink:0}}>{badge.t}</span>
                      </div>
                    );
                  })}
                </div>
                {scopeCheck.missing.length>0&&(
                  <div style={{marginTop:"0.5rem",padding:"0.5rem 0.65rem",background:"#FFFBEB",border:"1px dashed #F59E0B",borderRadius:8}}>
                    <div style={{fontSize:"0.58rem",fontWeight:800,color:"#B45309",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:"0.25rem"}}>⚠ In scope but not detected</div>
                    {scopeCheck.missing.map((m,i)=>(
                      <div key={i} style={{fontSize:"0.62rem",color:"#92400E",lineHeight:1.45}}>• {m.name}{m.note?<span style={{color:"#B45309"}}> — {m.note}</span>:""}</div>
                    ))}
                    <div style={{fontSize:"0.55rem",color:"#B45309",marginTop:"0.3rem"}}>The scope says these are yours — check the drawings for faces the auto-detect missed.</div>
                  </div>
                )}
                <div style={{fontSize:"0.56rem",color:"#94A3B8",marginTop:"0.35rem",lineHeight:1.4}}>Matched against the Scope tab's material list — a hint, not a gate. Nothing is dropped automatically.</div>
              </div>
            )}

            {/* Materials READ OFF a flattened drawing with OCR (sets with no extractable text) */}
            {results.ocrMaterials&&results.ocrMaterials.length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>🔎 Materials read off the drawing</div>
                <div style={{border:"1px solid #DDE7F2",borderRadius:8,overflow:"hidden"}}>
                  {results.ocrMaterials.map((m,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.42rem 0.6rem",background:i%2?"#F8FAFC":"#fff",borderTop:i?"1px solid #F1F5F9":"none"}}>
                      <div style={{width:8,height:8,borderRadius:2,background:hashColor(m.text),flexShrink:0}}/>
                      <span style={{fontSize:"0.64rem",color:"#334155",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={m.text}>{m.text}</span>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:"0.56rem",color:"#94A3B8",marginTop:"0.35rem",lineHeight:1.4}}>This set has no digital text, so the AI read the material spec straight off the drawing image — tag your regions with these.</div>
              </div>
            )}

            {/* Architect's own material schedule read off the drawing — a sanity-check target */}
            {results.drawingSchedule&&results.drawingSchedule.items&&results.drawingSchedule.items.length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>📐 Drawing's own schedule</div>
                <div style={{border:"1px solid #DDE7F2",borderRadius:8,overflow:"hidden"}}>
                  {results.drawingSchedule.items.map((it,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",gap:"0.4rem",padding:"0.4rem 0.6rem",background:i%2?"#F8FAFC":"#fff",borderTop:i?"1px solid #F1F5F9":"none"}}>
                      <span style={{fontSize:"0.62rem",color:"#475569",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={it.material}><b style={{color:BLUE}}>{it.key}</b> {it.material}</span>
                      <span style={{fontSize:"0.66rem",fontWeight:700,color:"#0F172A",whiteSpace:"nowrap"}}>{it.sf.toLocaleString()}</span>
                    </div>
                  ))}
                  <div style={{display:"flex",justifyContent:"space-between",padding:"0.45rem 0.6rem",background:BLUE_PALE,borderTop:"1px solid "+BLUE+"30"}}>
                    <span style={{fontSize:"0.62rem",color:BLUE_DARK,fontWeight:700,textTransform:"uppercase"}}>Stated total</span>
                    <span style={{fontSize:"0.72rem",fontWeight:800,color:BLUE}}>{results.drawingSchedule.total.toLocaleString()} SF</span>
                  </div>
                </div>
                <div style={{fontSize:"0.56rem",color:"#94A3B8",marginTop:"0.35rem",lineHeight:1.4}}>The architect's stated quantities, read straight off the drawing — check your takeoff against them.</div>
              </div>
            )}

            {/* Window/door COUNT surface — openings the readers detected + cut out of the SF */}
            {(()=>{const oc=(results.takeoffData||[]).reduce((s,e)=>s+(e.openingsCount||0),0);return oc>0&&(
              <div style={{border:"1px solid #DDE7F2",borderRadius:8,padding:"0.55rem 0.7rem",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:"0.62rem",color:"#475569"}}>🪟 <b>Openings detected</b> (windows/doors cut from SF)</span>
                <span style={{fontSize:"0.8rem",fontWeight:800,color:BLUE}}>{oc}</span>
              </div>
            );})()}

            {/* Openings from schedule */}
            {results.scheduleData&&results.scheduleData.total_opening_sf>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>Openings · Schedule</div>
                <div style={{padding:"0.6rem 0.75rem",background:"#FFFBEB",borderRadius:8,border:"1px solid #FDE68A"}}>
                  <div style={{fontSize:"1.05rem",fontWeight:800,color:"#92400E"}}>{Math.round(results.scheduleData.total_opening_sf).toLocaleString()} <span style={{fontSize:"0.62rem",fontWeight:400}}>SF</span></div>
                  <div style={{fontSize:"0.6rem",color:"#B45309",marginTop:2}}>{results.scheduleData.windows.length} window type{results.scheduleData.windows.length!==1?"s":""} · {results.scheduleData.doors.length} door type{results.scheduleData.doors.length!==1?"s":""} · exact from schedule</div>
                </div>
              </div>
            )}

            {/* Linear measurements (trim / soffit / fascia) captured from the estimator's polyline markups */}
            {linearRollup.length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>Trim &amp; Linear · {Math.round(linearTotalLF).toLocaleString()} LF</div>
                <div style={{border:"1px solid #E2E8F0",borderRadius:8,overflow:"hidden"}}>
                  {linearRollup.map((r,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"0.4rem 0.7rem",fontSize:"0.7rem",background:i%2?"#F8FAFC":"#fff",borderTop:i?"1px solid #F1F5F9":"none"}}>
                      <span style={{color:"#334155"}}>{r.material}</span>
                      <span style={{fontWeight:700,color:"#0F172A"}}>{Math.round(r.lf).toLocaleString()} LF</span>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:"0.58rem",color:"#64748B",marginTop:"0.4rem"}}>Captured from linear markups — priced per LF, separate from panel SF.</div>
              </div>
            )}

            {autoTrimRollup.length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:"#B45309",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>Suggested trim (verify) · {Math.round(autoTrimTotalLF).toLocaleString()} LF</div>
                <div style={{border:"1px dashed #F59E0B",borderRadius:8,overflow:"hidden",background:"#FFFBEB"}}>
                  {autoTrimRollup.map((r,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"0.4rem 0.7rem",fontSize:"0.7rem",borderTop:i?"1px solid #FDE68A":"none"}}>
                      <span style={{color:"#334155"}}>{r.material}</span>
                      <span style={{fontWeight:700,color:"#92400E"}}>{Math.round(r.lf).toLocaleString()} LF</span>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:"0.58rem",color:"#92400E",marginTop:"0.4rem"}}>Auto-derived from the detected faces (corners, base/top, openings). Estimates to verify — NOT priced automatically.</div>
              </div>
            )}

            {/* Legend */}
            {results.legend.length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>Legend</div>
                <div style={{display:"flex",flexDirection:"column",gap:"0.3rem"}}>
                  {results.legend.map((m,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:"0.5rem"}}>
                      <div style={{width:10,height:10,borderRadius:3,background:MAT_COLORS[m.category]||hashColor(m.name||m.id),flexShrink:0}}/>
                      <span style={{fontSize:"0.68rem",color:"#475569"}}>{m.id===m.name?<span style={{fontWeight:600,color:"#0F172A"}}>{m.name}</span>:<><span style={{fontWeight:600,color:"#0F172A"}}>{m.id}</span>: {m.name}</>}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* View toggle */}
            <div>
              <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>View</div>
              <div style={{display:"flex",flexDirection:"column",gap:"0.4rem"}}>
                {[["table","Table view"],["interactive","Interactive takeoff"],["edit","Edit surfaces"],["pricing","Pricing & bid"]].map(([mode,label])=>(
                  <button key={mode} onClick={()=>setViewMode(mode)} style={{padding:"0.6rem 0.85rem",borderRadius:9,fontSize:"0.73rem",fontWeight:viewMode===mode?600:500,fontFamily:"inherit",cursor:"pointer",border:"1px solid "+(viewMode===mode?"transparent":"#E4EAF1"),textAlign:"left",background:viewMode===mode?"linear-gradient(180deg,#5A92D2,#3F79BC)":"#fff",color:viewMode===mode?"#fff":"#475569",boxShadow:viewMode===mode?"0 6px 16px -4px rgba(74,134,200,0.5)":"0 1px 2px rgba(15,23,42,0.04)",letterSpacing:"-0.01em"}}>{label}</button>
                ))}
              </div>
            </div>

            {/* Export */}
            <div style={{display:"flex",flexDirection:"column",gap:"0.4rem"}}>
              <button onClick={exportExcel} style={{padding:"0.65rem",background:"linear-gradient(180deg,#5A92D2,#3F79BC)",color:"#fff",border:"none",borderRadius:9,fontSize:"0.73rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer",boxShadow:"0 6px 16px -5px rgba(74,134,200,0.55)",letterSpacing:"-0.01em"}}>↓  Export Excel</button>
              <button onClick={exportPDF} disabled={pdfLoading} style={{padding:"0.65rem",background:"#fff",color:pdfLoading?"#9CA3AF":"#3F79BC",border:"1px solid "+(pdfLoading?"#E2E8F0":"#CFE0F2"),borderRadius:9,fontSize:"0.73rem",fontWeight:600,fontFamily:"inherit",cursor:pdfLoading?"not-allowed":"pointer",letterSpacing:"-0.01em"}}>↓  {pdfLoading?"Generating…":"Evidence PDF"}</button>
            </div>

            {/* Measurement basis — builds confidence in the SF */}
            <div style={{padding:"0.6rem 0.7rem",background:"#F8FAFC",borderRadius:8,border:"1px solid #EEF2F7"}}>
              <div style={{fontSize:"0.58rem",color:"#94A3B8",lineHeight:1.55}}><span style={{color:"#475569",fontWeight:700}}>How SF is measured:</span> computed from the drawing's vector geometry at the verified drawing scale — not visual guesses. Use <b>Calibrate</b> on any sheet for exact, defensible numbers.</div>
            </div>

            {/* Log */}
            <div style={{marginTop:"auto"}}>
              <div style={{fontSize:"0.58rem",letterSpacing:"0.1em",color:"#CBD5E1",textTransform:"uppercase",fontWeight:600,marginBottom:"0.35rem"}}>Log</div>
              <div ref={logRef} style={{fontFamily:"'Courier New',monospace",fontSize:"0.65rem",maxHeight:100,overflowY:"auto",lineHeight:1.8,color:"#94A3B8"}}>
                {log.map((l,i)=><div key={i} style={{color:logColor[l.level]||"#94A3B8"}}>{l.msg}</div>)}
              </div>
            </div>
          </aside>

          {/* Main */}
          <main style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
            {viewMode==="interactive"&&(
              <div style={{flex:1,overflow:"hidden"}}><InteractiveView results={results} BACKEND={BACKEND} assignments={assignments} setAssignments={setAssignments} groupRename={groupRename} setGroupRename={setGroupRename} setResults={setResults} hiddenIds={hiddenIds} setHiddenIds={setHiddenIds} deletedStack={deletedStack} setDeletedStack={setDeletedStack} bucketShapes={bucketShapes} setBucketShapes={setBucketShapes} bucketColorNames={bucketColorNames} setBucketColorNames={setBucketColorNames}/></div>
            )}
            {viewMode==="edit"&&(
              <div style={{flex:1,overflow:"hidden"}}><EditorView results={results} BACKEND={BACKEND} setResults={setResults}/></div>
            )}
            {viewMode==="table"&&(
              <div style={{flex:1,overflowY:"auto",padding:"2rem"}}>
                <div style={{maxWidth:1020,margin:"0 auto"}}>

                {/* THE ANSWER, big and first: one card per material + the job total */}
                {summaryData&&Object.keys(summaryData).length>0&&(
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(215px,1fr))",gap:"1rem",marginBottom:"1.5rem"}}>
                    {Object.entries(summaryData).map(([cat,{net,adj,color}])=>(
                      <div key={cat} style={{background:"#fff",borderRadius:14,border:"1px solid #EEF2F7",borderTop:"4px solid "+color,padding:"1.1rem 1.25rem",boxShadow:"0 2px 10px rgba(15,23,42,0.05)"}}>
                        <div style={{fontSize:"0.85rem",fontWeight:700,color:"#334155",marginBottom:"0.4rem"}}>{cat}</div>
                        <div style={{fontSize:"2.1rem",fontWeight:800,color:"#0F172A",letterSpacing:"-0.02em",lineHeight:1}}>{Math.round(net).toLocaleString()}<span style={{fontSize:"0.85rem",fontWeight:600,color:"#94A3B8",marginLeft:6}}>SF</span></div>
                        <div style={{fontSize:"0.74rem",color:"#94A3B8",marginTop:"0.5rem"}}>with 15% waste → <b style={{color:"#475569"}}>{Math.round(adj).toLocaleString()} SF</b></div>
                      </div>
                    ))}
                    <div style={{background:NAVY,borderRadius:14,padding:"1.1rem 1.25rem",boxShadow:"0 8px 24px -8px rgba(15,33,56,0.5)"}}>
                      <div style={{fontSize:"0.85rem",fontWeight:700,color:"#7FB0E0",marginBottom:"0.4rem"}}>Job total (+15%)</div>
                      <div style={{fontSize:"2.1rem",fontWeight:800,color:"#fff",letterSpacing:"-0.02em",lineHeight:1}}>{Math.round(grandAdj).toLocaleString()}<span style={{fontSize:"0.85rem",fontWeight:600,color:"rgba(255,255,255,0.5)",marginLeft:6}}>SF</span></div>
                      <div style={{fontSize:"0.74rem",color:"rgba(255,255,255,0.55)",marginTop:"0.5rem"}}>{linearTotalLF>0?`+ ${Math.round(linearTotalLF).toLocaleString()} LF trim & linear`:"all materials, all pages"}</div>
                    </div>
                  </div>
                )}

                {/* Triage summary — tells the estimator where to look */}
                {triage.length>0&&(
                  <div style={{display:"flex",alignItems:"center",gap:"0.9rem",flexWrap:"wrap",padding:"0.85rem 1.1rem",marginBottom:"1.5rem",background:"#fff",borderRadius:12,border:"1px solid #F1F5F9",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
                    <div style={{display:"flex",flexDirection:"column"}}>
                      <span style={{fontSize:"0.6rem",letterSpacing:"0.08em",color:"#94A3B8",textTransform:"uppercase",fontWeight:700}}>Job confidence</span>
                      <span style={{fontSize:"1.5rem",fontWeight:800,color:jobConfidence>=85?"#15803D":jobConfidence>=65?"#B45309":"#B91C1C"}}>{jobConfidence}%</span>
                    </div>
                    <div style={{width:1,height:38,background:"#E2E8F0"}}/>
                    <div style={{display:"flex",gap:"0.5rem",flexWrap:"wrap",flex:1}}>
                      {readyN>0&&<span style={{fontSize:"0.78rem",fontWeight:700,padding:"0.35rem 0.8rem",borderRadius:20,background:"#DCFCE7",color:"#15803D"}}>✓ {readyN} ready to confirm</span>}
                      {reviewN>0&&<span style={{fontSize:"0.78rem",fontWeight:700,padding:"0.35rem 0.8rem",borderRadius:20,background:"#FEF3C7",color:"#B45309"}}>{reviewN} to review</span>}
                      {attnN>0&&<span style={{fontSize:"0.78rem",fontWeight:700,padding:"0.35rem 0.8rem",borderRadius:20,background:"#FEE2E2",color:"#B91C1C"}}>⚠ {attnN} need attention</span>}
                    </div>
                    <span style={{fontSize:"0.7rem",color:"#94A3B8"}}>{attnN+reviewN===0?"All clear — spot-check & send":"Start with the flagged ones below"}</span>
                  </div>
                )}

                {reviewElevs.length===0&&(
                  <div style={{background:"#fff",borderRadius:14,border:"1px solid #FDE68A",padding:"2.25rem 2rem",textAlign:"center",boxShadow:"0 2px 10px rgba(15,23,42,0.05)"}}>
                    <div style={{fontSize:"2.2rem",marginBottom:"0.5rem"}}>🤔</div>
                    <div style={{fontSize:"1.05rem",fontWeight:800,color:"#0F172A",marginBottom:"0.4rem"}}>No measurable areas found on this drawing</div>
                    <div style={{fontSize:"0.82rem",color:"#64748B",lineHeight:1.7,maxWidth:520,margin:"0 auto"}}>The AI couldn't detect cladding on these pages and found no Bluebeam measurements to read. You can still measure it yourself: open the <b>Draw</b> tab, set the scale, and trace the walls.</div>
                  </div>
                )}
                {reviewElevs.length>0&&<div style={{fontSize:"0.72rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"1rem"}}>By elevation{reviewElevs.length>6?" · biggest first":""}</div>}
                {(reviewElevs.length>6
                  ? [...results.takeoffData].sort((a,b)=>(b.zones||[]).reduce((s,z)=>s+(z.netArea||0),0)-(a.zones||[]).reduce((s,z)=>s+(z.netArea||0),0))
                  : results.takeoffData).map((elev,i)=>{
                  const total=(elev.zones||[]).reduce((s,z)=>s+(z.netArea||0),0);
                  if(total===0)return null;
                  const conf=elevConfidence(elev);
                  const cs=STATUS_STYLE[conf.status];
                  const dispScale=elev.verifiedScale||elev.scale;
                  const overshoot=elev.expectedFacadeSF&&total>elev.expectedFacadeSF*1.4;
                  const warnLines=[
                    ...(conf.status!=="ready"&&conf.reasons.length?["Check: "+conf.reasons.join(" · ")]:[]),
                    ...(elev.scaleSource==="default"?["Default scale used — calibrate before trusting SF"]:[]),
                    ...(overshoot?["Panel SF is bigger than the building face — possible scale error"]:[]),
                    ...(elev.flags||[]).filter(Boolean),
                  ];
                  return <div key={i} style={{background:"#fff",borderRadius:14,boxShadow:"0 2px 10px rgba(15,23,42,0.05)",border:"1px solid #F1F5F9",marginBottom:"1.25rem",overflow:"hidden"}}>
                    <div style={{padding:"0.9rem 1.25rem",display:"flex",justifyContent:"space-between",alignItems:"center",gap:"0.75rem",borderBottom:"1px solid #F1F5F9",background:"#FAFBFC"}}>
                      <div style={{display:"flex",alignItems:"center",gap:"0.6rem",minWidth:0}}>
                        <span style={{fontSize:"1.05rem",fontWeight:800,color:"#0F172A",letterSpacing:"-0.01em"}}>{elev.title}</span>
                        <span title={conf.reasons.join(" · ")} style={{fontSize:"0.62rem",fontWeight:700,padding:"0.18rem 0.55rem",borderRadius:20,background:cs.bg,color:cs.fg,whiteSpace:"nowrap"}}>{cs.label}</span>
                        {dispScale&&<span style={{fontSize:"0.66rem",color:"#94A3B8",whiteSpace:"nowrap"}}>{elev.sheetRef?elev.sheetRef+" · ":""}scale {dispScale}</span>}
                      </div>
                      <span style={{fontSize:"1.15rem",fontWeight:800,color:BLUE,whiteSpace:"nowrap"}}>{Math.round(total).toLocaleString()} <span style={{fontSize:"0.7rem",fontWeight:600,color:"#94A3B8"}}>SF</span></span>
                    </div>
                    {warnLines.length>0&&<div style={{padding:"0.55rem 1.25rem",fontSize:"0.74rem",color:"#92400E",background:"#FFFBEB",borderBottom:"1px solid #FEF3C7",lineHeight:1.5}}>⚠ {warnLines.join(" · ")}</div>}
                    {(elev.zones||[]).map((z,zi)=>(
                      <div key={zi} style={{display:"flex",alignItems:"center",gap:"0.8rem",padding:"0.85rem 1.25rem",borderTop:zi?"1px solid #F5F8FB":"none"}}>
                        <div style={{width:13,height:13,borderRadius:4,background:MAT_COLORS[z.category]||hashColor(z.materialName||z.category),flexShrink:0}}/>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontSize:"0.95rem",fontWeight:600,color:"#1E293B"}}>{z.materialName}{z.materialId?<span style={{fontSize:"0.7rem",fontWeight:700,color:BLUE,marginLeft:8}}>{z.materialId}</span>:null}</div>
                          {(z.totalOpeningArea||0)>0&&<div style={{fontSize:"0.7rem",color:"#94A3B8",marginTop:2}}>{Math.round(z.grossArea||0).toLocaleString()} gross − {Math.round(z.totalOpeningArea||0).toLocaleString()} openings</div>}
                        </div>
                        <div style={{textAlign:"right",whiteSpace:"nowrap"}}>
                          <div style={{fontSize:"1.15rem",fontWeight:800,color:"#0F172A",letterSpacing:"-0.01em"}}>{Math.round(z.netArea||0).toLocaleString()} <span style={{fontSize:"0.68rem",fontWeight:600,color:"#94A3B8"}}>SF</span></div>
                          <div style={{fontSize:"0.68rem",color:"#94A3B8"}}>+15% → {Math.round((z.netArea||0)*1.15).toLocaleString()}</div>
                        </div>
                      </div>
                    ))}
                  </div>;
                })}
                </div>
              </div>
            )}
            {viewMode==="pricing"&&(
              <div style={{flex:1,overflowY:"auto",padding:"1.5rem"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
                  <div style={{fontSize:"0.65rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700}}>Pricing &amp; Bid</div>
                  <button onClick={exportPricedExcel} disabled={!priceRows.length} style={{padding:"0.5rem 1rem",background:priceRows.length?BLUE:"#E2E8F0",color:priceRows.length?"#fff":"#9CA3AF",border:"none",borderRadius:7,fontSize:"0.72rem",fontWeight:700,fontFamily:"inherit",cursor:priceRows.length?"pointer":"not-allowed"}}>↓ Export Priced Bid (Excel)</button>
                </div>
                {!priceRows.length?(
                  <div style={{fontSize:"0.8rem",color:"#94A3B8"}}>No materials detected yet — run an analysis first.</div>
                ):(
                  <>
                    <div style={{fontSize:"0.64rem",padding:"0.5rem 0.75rem",borderRadius:7,marginBottom:"1rem",background:hasReviewed?"#F0FDF4":"#F8FAFC",border:"1px solid "+(hasReviewed?"#BBF7D0":"#E2E8F0"),color:hasReviewed?"#15803D":"#64748B"}}>{hasReviewed?`✓ Pricing your reviewed takeoff — ${Object.keys(reviewedSummary).length} material${Object.keys(reviewedSummary).length!==1?"s":""} you assigned in Interactive Takeoff`:"Pricing the AI's auto-detected materials. Assign surfaces in Interactive Takeoff to price your reviewed numbers instead."}</div>
                    <div style={{display:"flex",gap:"1rem",marginBottom:"1rem",flexWrap:"wrap"}}>
                      <div style={{display:"flex",flexDirection:"column",gap:"0.25rem"}}>
                        <label style={{fontSize:"0.6rem",color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.05em",fontWeight:600}}>Waste %</label>
                        <input type="number" value={pricing.wastePct} onChange={e=>setPricing(p=>({...p,wastePct:parseFloat(e.target.value)||0}))} style={{width:80,padding:"0.4rem 0.5rem",borderRadius:6,border:"1px solid #E2E8F0",fontSize:"0.8rem",fontFamily:"inherit"}}/>
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:"0.25rem"}}>
                        <label style={{fontSize:"0.6rem",color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.05em",fontWeight:600}}>Margin %</label>
                        <input type="number" value={pricing.marginPct} onChange={e=>setPricing(p=>({...p,marginPct:parseFloat(e.target.value)||0}))} style={{width:80,padding:"0.4rem 0.5rem",borderRadius:6,border:"1px solid #E2E8F0",fontSize:"0.8rem",fontFamily:"inherit"}}/>
                      </div>
                    </div>
                    <div style={{background:"#fff",borderRadius:10,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:"1px solid #F1F5F9",overflow:"hidden",marginBottom:"1rem"}}>
                      <table style={{width:"100%",borderCollapse:"collapse"}}>
                        <thead><tr style={{background:"#F8FAFC"}}>{["Material","Net SF","+Waste SF","Rate $/SF","Extended $"].map(h=>(
                          <th key={h} style={{padding:"0.5rem 0.9rem",textAlign:h==="Material"?"left":"right",fontSize:"0.6rem",fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.05em",borderBottom:"1px solid #F1F5F9"}}>{h}</th>
                        ))}</tr></thead>
                        <tbody>{priceRows.map(r=>(
                          <tr key={r.cat} style={{borderBottom:"1px solid #F8FAFC"}}>
                            <td style={{padding:"0.5rem 0.9rem"}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:2,background:MAT_COLORS[r.cat]||"#9CA3AF"}}/><span style={{fontSize:"0.75rem",color:"#374151",fontWeight:500}}>{r.cat}</span></div></td>
                            <td style={{padding:"0.5rem 0.9rem",textAlign:"right",fontSize:"0.72rem",color:"#64748B"}}>{Math.round(r.net).toLocaleString()}</td>
                            <td style={{padding:"0.5rem 0.9rem",textAlign:"right",fontSize:"0.72rem",color:"#64748B"}}>{Math.round(r.adjSF).toLocaleString()}</td>
                            <td style={{padding:"0.5rem 0.9rem",textAlign:"right"}}>
                              <span style={{color:"#94A3B8",fontSize:"0.72rem"}}>$</span>
                              <input type="number" value={r.rate} onChange={e=>setPricing(p=>({...p,rates:{...p.rates,[r.cat]:parseFloat(e.target.value)||0}}))} style={{width:64,padding:"0.3rem 0.4rem",borderRadius:5,border:"1px solid #E2E8F0",fontSize:"0.75rem",textAlign:"right",fontFamily:"inherit"}}/>
                            </td>
                            <td style={{padding:"0.5rem 0.9rem",textAlign:"right",fontSize:"0.8rem",fontWeight:700,color:"#0F172A"}}>${Math.round(r.cost).toLocaleString()}</td>
                          </tr>
                        ))}</tbody>
                      </table>
                    </div>
                    <div style={{maxWidth:360,marginLeft:"auto",display:"flex",flexDirection:"column",gap:"0.4rem"}}>
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:"0.78rem",color:"#475569"}}><span>Cost subtotal (incl. {pricing.wastePct}% waste)</span><span style={{fontWeight:600}}>${Math.round(costSubtotal).toLocaleString()}</span></div>
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:"0.78rem",color:"#475569"}}><span>Margin ({pricing.marginPct}%)</span><span style={{fontWeight:600}}>${Math.round(bidTotal-costSubtotal).toLocaleString()}</span></div>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"0.65rem 0.85rem",background:BLUE_PALE,borderRadius:8,border:"1px solid "+BLUE+"40",marginTop:"0.3rem"}}>
                        <span style={{fontSize:"0.7rem",color:BLUE_DARK,fontWeight:800,textTransform:"uppercase",letterSpacing:"0.05em"}}>Bid Total</span>
                        <span style={{fontSize:"1.2rem",fontWeight:800,color:BLUE}}>${Math.round(bidTotal).toLocaleString()}</span>
                      </div>
                    </div>
                    <div style={{fontSize:"0.62rem",color:"#94A3B8",marginTop:"1rem",maxWidth:620,lineHeight:1.6}}>Rates are editable ballpark defaults ($/SF installed). The exported Estimate sheet is at cost; the Proposal total applies your margin. Set rates to your real numbers before bidding.</div>
                  </>
                )}
              </div>
            )}
          </main>
        </div>
      )}
      </>)}
    </div>
  );
}
