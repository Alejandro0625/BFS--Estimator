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
function InteractiveView({ results, BACKEND, assignments, setAssignments }) {
  const [elevIdx, setElevIdx] = useState(0);
  const [pageImage, setPageImage] = useState(null);
  const [imgLoaded, setImgLoaded] = useState(false);
  const [pagePolygons, setPagePolygons] = useState([]);
  const [pageDims, setPageDims] = useState({ width:612, height:792 });
  const [activeGroup, setActiveGroup] = useState(null);
  const [imgNaturalSize, setImgNaturalSize] = useState({ w:1, h:1 });
  const [calibMode, setCalibMode] = useState(false);
  const [calibPts, setCalibPts] = useState([]);
  const [calibFt, setCalibFt] = useState(null);
  const [realDist, setRealDist] = useState("");
  const imgRef = useRef();
  const svgRef = useRef();
  const elevations = results.takeoffData.filter(e => e.pageNumber);
  const elev = elevations[elevIdx];
  const pageNum = elev?.pageNumber;

  useEffect(() => {
    if (!pageNum || !results.jobId) return;
    setPageImage(null); setImgLoaded(false); setPagePolygons([]); setActiveGroup(null); setCalibMode(false); setCalibPts([]);
    fetch(BACKEND+"/polygons/"+results.jobId+"/"+pageNum)
      .then(r=>r.ok?r.json():{polygons:[],width:612,height:792})
      .then(d=>{setPagePolygons(d.polygons||[]);setPageDims({width:d.width||612,height:d.height||792});})
      .catch(()=>{});
    setPageImage(BACKEND+"/page-image/"+results.jobId+"/"+pageNum);
  }, [elevIdx, pageNum, results.jobId, BACKEND]);

  // Pull shared learning from the server into local memory (so repeats are pre-identified)
  useEffect(()=>{
    fetch(BACKEND+"/recall").then(r=>r.ok?r.json():null).then(d=>{
      if(d&&d.hatches){ const m=loadLearned(); let ch=false; for(const k in d.hatches){ if(!m[k]){ m[k]=d.hatches[k]; ch=true; } } if(ch) saveLearned(m); }
    }).catch(()=>{});
  },[BACKEND]);

  const polyMethod = pagePolygons[0]?.source||(pagePolygons.length>0?"vector":"box");
  const rawZones = pagePolygons.length>0 ? pagePolygons : (elev?.zones||[]).map((z,i)=>({
    id:i,points:[[z.x0pct/100,z.y0pct/100],[z.x1pct/100,z.y0pct/100],[z.x1pct/100,z.y1pct/100],[z.x0pct/100,z.y1pct/100]],
    area_sf:z.netArea||0,cx:(z.x0pct+z.x1pct)/200,cy:(z.y0pct+z.y1pct)/200,source:"box",
  }));
  // When the user calibrates the scale, recompute SF from polygon geometry (real polygons only)
  const displayZones = calibFt
    ? rawZones.map(z => (z.source && z.source!=="box")
        ? {...z, area_sf: polyAreaSF(z.points, calibFt, pageDims.width, pageDims.height)}
        : z)
    : rawZones;
  // Effective scale: calibrated value, else back it out from a zone's known SF + geometry
  const effFtPerInch = calibFt || (()=>{
    for(const z of rawZones){
      if(z.area_sf>0 && z.points?.length>=3){
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
  const exportInteractiveExcel = () => {
    const mt={};
    Object.values(assignments).forEach(a=>{const k=a.materialName||a.category||"Panel";if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=a.area_sf||0;});
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
  const handleSvgClick=evt=>{
    if(!calibMode) return;
    const p=getSvgPoint(evt); if(!p) return;
    setCalibPts(prev=>prev.length>=2?[p]:[...prev,p]);
  };
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
            {polyMethod==="bluebeam"?`📐 Bluebeam — ${displayZones.length} surfaces`:polyMethod==="vector_cluster"||polyMethod==="vector"?`📏 Vector — ${displayZones.length} surfaces`:polyMethod==="claude_vision"?`🧠 AI Vision — ${displayZones.length} surfaces`:"No surfaces on this page"}
          </div>
          <button onClick={()=>{setCalibMode(m=>!m);setCalibPts([]);}} style={{fontSize:"0.65rem",padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid "+(calibMode?"#EF4444":"#2D5280"),background:calibMode?"#7F1D1D":NAVY_LT,color:calibMode?"#FCA5A5":"#94A3B8",cursor:"pointer",fontFamily:"inherit"}}>📏 {calibMode?(calibPts.length<2?`Click point ${calibPts.length+1} of 2`:"2 points set"):"Calibrate scale"}</button>
          {calibFt&&!calibMode&&<div style={{fontSize:"0.62rem",padding:"0.3rem 0.6rem",borderRadius:20,background:"#064E3B",color:"#6EE7B7",border:"1px solid #065F46"}}>✓ Calibrated · {calibFt.toFixed(2)} ft/in<span onClick={()=>setCalibFt(null)} style={{cursor:"pointer",textDecoration:"underline",marginLeft:6}}>reset</span></div>}
          {calibMode&&calibPts.length===2&&<div style={{display:"flex",alignItems:"center",gap:"0.35rem",fontSize:"0.65rem",color:"#CBD5E1"}}>
            <span>Real distance (ft):</span>
            <input value={realDist} onChange={e=>setRealDist(e.target.value)} onKeyDown={e=>e.key==="Enter"&&applyCalibration()} placeholder="e.g. 20" style={{width:64,padding:"0.25rem 0.4rem",borderRadius:5,border:"1px solid #2D5280",background:NAVY,color:"#E2E8F0",fontSize:"0.65rem",fontFamily:"inherit"}}/>
            <button onClick={applyCalibration} style={{fontSize:"0.65rem",padding:"0.25rem 0.6rem",borderRadius:5,border:"none",background:BLUE,color:"#fff",cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>Apply</button>
          </div>}
        </div>
        {!pageImage?<div style={{color:"#475569",fontSize:"0.8rem",marginTop:"4rem"}}>Loading elevation...</div>:
          <div style={{position:"relative",display:"inline-block",maxWidth:"100%"}}>
            <img ref={imgRef} src={pageImage} alt={elev?.title} onLoad={e=>{setImgNaturalSize({w:e.target.naturalWidth,h:e.target.naturalHeight});setImgLoaded(true);}} style={{display:"block",maxWidth:"100%",maxHeight:"calc(100vh - 180px)",objectFit:"contain",borderRadius:6,border:"1px solid "+NAVY_LT}}/>
            {imgLoaded&&<svg ref={svgRef} onClick={handleSvgClick} viewBox={`0 0 ${pageDims.width} ${pageDims.height}`} style={{position:"absolute",top:0,left:0,width:svgW,height:svgH,overflow:"visible",cursor:calibMode?"crosshair":"default"}}>
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
                const showLabel=!dimmed&&(a||isSel||zone.source==="bluebeam"||zone.source==="claude_vision");
                return <g key={zone.id} style={{cursor:calibMode?"crosshair":"pointer"}} onClick={e=>{if(calibMode)return;e.stopPropagation();const k=gkey(zone);setActiveGroup(activeGroup===k?null:k);}}>
                  <polygon points={pts} fill={color} fillOpacity={dimmed?0.06:isSel?0.6:zone.source==="bluebeam"?0.45:a?0.38:0.22} stroke={isSel?"#fff":color} strokeWidth={isSel?2.5:1.5} strokeOpacity={dimmed?0.25:0.9}/>
                  {showLabel&&<><rect x={lx-32} y={ly-9} width={64} height={18} fill="rgba(0,0,0,0.8)" rx={4}/><text x={lx} y={ly+2} textAnchor="middle" dominantBaseline="middle" fill="white" fontSize={pageDims.width/75} fontFamily="Inter,Arial" fontWeight="bold">{zone.source==="claude_vision"&&!a?zone.material_type:Math.round(zone.area_sf)+" SF"}</text></>}
                </g>;
              })}
              {calibPts.length===2&&<line x1={calibPts[0].x*pageDims.width} y1={calibPts[0].y*pageDims.height} x2={calibPts[1].x*pageDims.width} y2={calibPts[1].y*pageDims.height} stroke="#EF4444" strokeWidth={pageDims.width/350} strokeDasharray={pageDims.width/90}/>}
              {calibPts.map((p,i)=><circle key={"cp"+i} cx={p.x*pageDims.width} cy={p.y*pageDims.height} r={pageDims.width/110} fill="#EF4444" stroke="#fff" strokeWidth={pageDims.width/600}/>)}
            </svg>}
          </div>
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
              <button onClick={exportInteractiveExcel} style={{width:"100%",padding:"0.65rem",background:BLUE,color:"#fff",border:"none",borderRadius:7,fontSize:"0.72rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓ Export Excel</button>
            </>}
            {Object.keys(clusterSummary).length===0&&Object.keys(totals).length===0&&<div style={{fontSize:"0.7rem",color:"#475569",lineHeight:1.8}}>Click any colored area → every area with that same hatch selects and shows its total SF.</div>}
            <div style={{marginTop:"1rem",fontSize:"0.6rem",color:"#334155"}}>{Object.keys(assignments).length} areas tagged · {elevations.length} elevations</div>
          </>
        )}
      </div>
    </div>
  );
}

/* ── Konva polygon editor: drag vertices to correct the AI, live SF ── */
function EditorView({ results, BACKEND }) {
  const elevations = results.takeoffData.filter(e=>e.pageNumber);
  const [elevIdx,setElevIdx]=useState(0);
  const elev=elevations[elevIdx];
  const pageNum=elev?.pageNumber;
  const [polys,setPolys]=useState([]);
  const [pageDims,setPageDims]=useState({width:612,height:792});
  const [img,setImg]=useState(null);
  const [selId,setSelId]=useState(null);
  const [stageW,setStageW]=useState(800);
  const wrapRef=useRef();

  useEffect(()=>{
    if(!pageNum||!results.jobId) return;
    setPolys([]); setSelId(null); setImg(null);
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
  const effFt = (()=>{ for(const p of polys){ if(p.area_sf>0&&p.points.length>=3){ const sh=polyAreaSF(p.points,72,pageDims.width,pageDims.height); if(sh>0) return 72*Math.sqrt(p.area_sf/sh);}} return 8; })();
  const areaOf=p=>polyAreaSF(p.points,effFt,pageDims.width,pageDims.height);
  const totalSF=polys.reduce((s,p)=>s+areaOf(p),0);
  const updateVertex=(pid,vi,nx,ny)=>setPolys(prev=>prev.map(p=>p.id!==pid?p:{...p,points:p.points.map((pt,i)=>i===vi?[nx,ny]:pt)}));
  const deletePoly=pid=>{ setPolys(prev=>prev.filter(p=>p.id!==pid)); setSelId(null); };

  return (
    <div style={{display:"flex",height:"100%",overflow:"hidden",background:NAVY,fontFamily:"'Inter','Segoe UI',sans-serif"}}>
      <div style={{width:170,borderRight:"1px solid "+NAVY_LT,overflowY:"auto",flexShrink:0,background:NAVY_MID}}>
        <div style={{padding:"0.8rem",fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,borderBottom:"1px solid "+NAVY_LT}}>Elevations</div>
        {elevations.map((e,i)=><div key={i} onClick={()=>setElevIdx(i)} style={{padding:"0.6rem 0.8rem",cursor:"pointer",borderBottom:"1px solid "+NAVY_LT,background:i===elevIdx?NAVY_LT:"transparent",fontSize:"0.7rem",color:i===elevIdx?"#E2E8F0":"#94A3B8",borderLeft:i===elevIdx?"3px solid "+BLUE:"3px solid transparent"}}>{e.title||"Page "+e.pageNumber}</div>)}
      </div>
      <div ref={wrapRef} style={{flex:1,overflow:"auto",padding:"1rem"}}>
        <div style={{fontSize:"0.62rem",color:"#94A3B8",marginBottom:"0.5rem"}}>Click a shape to select · drag its dots to correct it · SF updates live</div>
        {img?<Stage width={stageW} height={stageH} onMouseDown={e=>{ if(e.target===e.target.getStage()) setSelId(null); }}>
          <Layer>
            <KImage image={img} width={stageW} height={stageH}/>
            {polys.map(p=>{
              const flat=p.points.flatMap(([nx,ny])=>[nx*pageDims.width*sc, ny*pageDims.height*sc]);
              const col=CLUSTER_COLORS[p.id%CLUSTER_COLORS.length];
              const sel=selId===p.id;
              return <Line key={p.id} points={flat} closed fill={col+(sel?"66":"33")} stroke={sel?"#ffffff":col} strokeWidth={sel?2:1} onClick={()=>setSelId(p.id)} onTap={()=>setSelId(p.id)}/>;
            })}
            {selId!==null&&(polys.find(p=>p.id===selId)?.points||[]).map(([nx,ny],vi)=>(
              <Circle key={vi} x={nx*pageDims.width*sc} y={ny*pageDims.height*sc} radius={5} fill="#ffffff" stroke="#3B82F6" strokeWidth={2} draggable
                onDragMove={e=>updateVertex(selId, vi, e.target.x()/(pageDims.width*sc), e.target.y()/(pageDims.height*sc))}/>
            ))}
          </Layer>
        </Stage>:<div style={{color:"#475569",fontSize:"0.8rem",marginTop:"3rem"}}>Loading elevation…</div>}
      </div>
      <div style={{width:220,borderLeft:"1px solid "+NAVY_LT,padding:"1rem",background:NAVY_MID,overflowY:"auto",flexShrink:0}}>
        <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>Surfaces ({polys.length})</div>
        {polys.map(p=>{
          const sel=selId===p.id;
          return <div key={p.id} onClick={()=>setSelId(p.id)} style={{padding:"0.5rem 0.65rem",marginBottom:"0.35rem",background:sel?NAVY_LT:NAVY,borderRadius:6,cursor:"pointer",borderLeft:"3px solid "+CLUSTER_COLORS[p.id%CLUSTER_COLORS.length]}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontSize:"0.65rem",color:"#CBD5E1"}}>{p.category}</span>
              <span style={{fontSize:"0.72rem",fontWeight:700,color:"#E2E8F0"}}>{Math.round(areaOf(p)).toLocaleString()} SF</span>
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
  const [pricing, setPricing] = useState({ rates:{}, wastePct:15, marginPct:20 });
  const [assignments, setAssignments] = useState({});
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
          setResults({legend:data.legend||[],takeoffData:data.takeoffData||[],scheduleData:data.scheduleData||null,projName:file?.name?.replace(".pdf","")||"Project",jobId:id});
          setPhase("done");setProgress({label:"Complete",pct:100});
        }else if(data.status==="error"){clearInterval(pollRef.current);setErrMsg(data.error||"Unknown error");setPhase("error");}
      }catch(e){console.log("poll",e.message);}
    },5000);
  },[file]);

  const run = async()=>{
    if(!file)return;
    setPhase("running");setLog([]);setErrMsg("");setResults(null);setAssignments({});seenLogs.current=0;
    try{
      setLog([{msg:"Uploading PDF...",level:"info"}]);
      const fd=new FormData();fd.append("pdf",file);
      const res=await fetch(BACKEND+"/analyze",{method:"POST",body:fd});
      const{jobId:id}=await res.json();
      setLog(prev=>[...prev,{msg:"Analysis started — job "+id,level:"ok"}]);
      startPolling(id);
    }catch(err){setErrMsg(err.message);setPhase("error");}
  };

  const exportExcel=()=>{
    if(!results)return;
    const mt={};
    results.takeoffData.forEach(e=>(e.zones||[]).forEach(z=>{const k=z.materialName||z.category||"Panel";if(!mt[k])mt[k]={name:k,sf:0};mt[k].sf+=z.netArea||0;}));
    const wb=buildExcel(results.projName||"Project",Object.values(mt));
    XLSX.writeFile(wb,"BFS_Takeoff_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
  };

  const exportPDF=async()=>{
    if(!results?.jobId)return;
    setPdfLoading(true);
    try{
      const res=await fetch(BACKEND+"/evidence-pdf/"+results.jobId);
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
      data:{ legend:results.legend, takeoffData:results.takeoffData, scheduleData:results.scheduleData||null, projName:results.projName, jobId:results.jobId },
      assignments, pricing };
    try{ localStorage.setItem("bfs_bid_"+id, JSON.stringify(rec)); refreshSaved(); }
    catch(e){ alert("Could not save bid: "+e.message); }
  };
  const loadBid=rec=>{
    setResults(rec.data); setAssignments(rec.assignments||{});
    setPricing(rec.pricing||{rates:{},wastePct:15,marginPct:20});
    setPhase("done"); setViewMode("table"); setFile(null);
  };
  const deleteBid=(id,ev)=>{ if(ev)ev.stopPropagation(); try{ localStorage.removeItem("bfs_bid_"+id); }catch{} refreshSaved(); };

  const summary=results?()=>{
    const t={};
    results.takeoffData.forEach(e=>(e.zones||[]).forEach(z=>{const k=z.category||"Other";if(!t[k])t[k]={net:0,adj:0,color:MAT_COLORS[k]||"#9CA3AF"};t[k].net+=z.netArea||0;t[k].adj+=(z.netArea||0)*1.15;}));
    return t;
  }:null;
  const summaryData = summary ? summary() : null;
  const grandAdj = summaryData ? Object.values(summaryData).reduce((s,v)=>s+v.adj,0) : 0;
  // Reviewed takeoff = what the user assigned in Interactive; drives the bid when present
  const reviewedSummary = Object.values(assignments).reduce((acc,a)=>{
    const cat=a.category||a.materialName||"Panel"; if(!acc[cat])acc[cat]={net:0}; acc[cat].net+=a.area_sf||0; return acc;
  },{});
  const hasReviewed = Object.keys(reviewedSummary).length>0;
  const pricingSource = hasReviewed ? reviewedSummary : (summaryData||{});
  const priceRows = Object.entries(pricingSource).map(([cat,{net}])=>{
    const rate = pricing.rates[cat]!=null ? pricing.rates[cat] : (DEFAULT_RATES[cat]??DEFAULT_RATES.Other);
    const adjSF = net*(1+pricing.wastePct/100);
    return { cat, net, adjSF, rate, cost:adjSF*rate };
  });
  const costSubtotal = priceRows.reduce((s,r)=>s+r.cost,0);
  const bidTotal = costSubtotal*(1+pricing.marginPct/100);
  const exportPricedExcel=()=>{
    if(!priceRows.length)return;
    const mats=priceRows.map(r=>({name:r.cat,sf:r.net,rate:r.rate}));
    const wb=buildExcel(results.projName||"Project",mats,{wastePct:pricing.wastePct,marginPct:pricing.marginPct});
    XLSX.writeFile(wb,"BFS_Bid_"+(results.projName||"Project").replace(/\s+/g,"_")+".xlsx");
  };
  // ── Pre-bid reliability checks ──
  const reviewElevs = (results?.takeoffData||[]).filter(e=>(e.zones||[]).some(z=>(z.netArea||0)>0));
  const defaultScaleN = reviewElevs.filter(e=>e.scaleSource==="default" || (!e.verifiedScale && !e.scale)).length;
  const scaleWarnN = reviewElevs.filter(e=>e.expectedFacadeSF && (e.zones||[]).reduce((s,z)=>s+(z.netArea||0),0) > e.expectedFacadeSF*1.4).length;
  const hasSchedule = !!(results?.scheduleData?.total_opening_sf>0);
  const reviewOk = reviewElevs.length>0 && defaultScaleN===0 && scaleWarnN===0;
  const phaseStep={idle:0,running:1,filtering:1,legend:2,analyzing:3,done:4,error:0}[phase]||0;
  const isRunning=!["idle","done","error"].includes(phase);
  const logColor={ok:"#22C55E",warn:"#F59E0B",error:"#EF4444",success:"#22C55E",dim:"#94A3B8",info:"#64748B"};

  const showResults = phase==="done" && results;
  const showUploadScreen = !showResults;

  return (
    <div style={{fontFamily:"'Inter','Segoe UI',-apple-system,sans-serif",background:showUploadScreen?"#0C1B2E":"#F0F4F8",minHeight:"100vh",display:"flex",flexDirection:"column",color:"#1E293B"}}>

      {/* ── Header ── */}
      <header style={{background:NAVY,height:60,padding:"0 1.75rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0,boxShadow:"0 1px 0 rgba(255,255,255,0.06)",zIndex:10}}>
        <div style={{display:"flex",alignItems:"center",gap:"0.875rem"}}>
          <img src="/logo-bfs.png" alt="BFS" style={{height:40,width:"auto"}}/>
          <div style={{width:1,height:32,background:"rgba(255,255,255,0.1)"}}/>
          <div>
            <div style={{fontSize:"0.55rem",letterSpacing:"0.2em",color:"rgba(255,255,255,0.4)",textTransform:"uppercase",fontWeight:500}}>Boston Facade Systems</div>
            <div style={{fontSize:"0.95rem",fontWeight:700,color:"#fff",letterSpacing:"-0.01em"}}>AI Panel Estimator</div>
          </div>
        </div>
        {showResults&&(
          <div style={{display:"flex",gap:"0.5rem"}}>
            <button onClick={exportExcel} style={{padding:"0.45rem 1rem",background:"transparent",color:"rgba(255,255,255,0.7)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↓ Excel</button>
            <button onClick={exportPDF} disabled={pdfLoading} style={{padding:"0.45rem 1rem",background:BLUE,color:"#fff",border:"none",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↓ {pdfLoading?"Generating...":"Evidence PDF"}</button>
            <button onClick={saveBid} style={{padding:"0.45rem 1rem",background:"transparent",color:"rgba(255,255,255,0.7)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>💾 Save</button>
            <button onClick={()=>{setFile(null);setPhase("idle");setResults(null);setLog([]);setAssignments({});}} style={{padding:"0.45rem 1rem",background:"rgba(255,255,255,0.08)",color:"rgba(255,255,255,0.6)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↺ New</button>
          </div>
        )}
      </header>

      {/* ══ UPLOAD SCREEN ══ */}
      {showUploadScreen&&(
        <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"2rem",position:"relative",overflow:"hidden"}}>
          {/* Background grid */}
          <div style={{position:"absolute",inset:0,backgroundImage:`linear-gradient(rgba(74,134,200,0.04) 1px, transparent 1px), linear-gradient(90deg, rgba(74,134,200,0.04) 1px, transparent 1px)`,backgroundSize:"40px 40px",pointerEvents:"none"}}/>

          {/* Center card */}
          <div style={{position:"relative",width:"100%",maxWidth:540,display:"flex",flexDirection:"column",gap:"2rem"}}>

            {/* Hero text */}
            <div style={{textAlign:"center"}}>
              <div style={{fontSize:"0.7rem",letterSpacing:"0.2em",color:BLUE,textTransform:"uppercase",fontWeight:600,marginBottom:"0.75rem"}}>Boston Facade Systems</div>
              <h1 style={{fontSize:"2.25rem",fontWeight:800,color:"#fff",margin:0,letterSpacing:"-0.03em",lineHeight:1.1}}>AI Panel Estimator</h1>
              <p style={{fontSize:"0.9rem",color:"rgba(255,255,255,0.35)",marginTop:"0.75rem",lineHeight:1.6}}>Upload your blueprint PDF and get a full material takeoff<br/>with SF breakdown by elevation in seconds.</p>
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
                border:`2px dashed ${dragOver?BLUE:file?"#22C55E":"rgba(74,134,200,0.35)"}`,
                borderRadius:16,
                padding:"2.5rem 2rem",
                textAlign:"center",
                cursor:isRunning?"default":"pointer",
                background:dragOver?"rgba(74,134,200,0.08)":file?"rgba(34,197,94,0.05)":"rgba(255,255,255,0.03)",
                backdropFilter:"blur(8px)",
                transition:"all 0.2s",
                position:"relative",
              }}>
              {!file&&!isRunning&&(
                <>
                  <div style={{fontSize:"3rem",marginBottom:"0.75rem",opacity:0.5}}>📂</div>
                  <div style={{fontSize:"1rem",fontWeight:600,color:"rgba(255,255,255,0.75)",marginBottom:"0.35rem"}}>Drop your blueprint PDF here</div>
                  <div style={{fontSize:"0.78rem",color:"rgba(255,255,255,0.3)"}}>or click to browse your files</div>
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
                  <div style={{fontSize:"0.72rem",color:"rgba(255,255,255,0.4)",marginBottom:"0.75rem",letterSpacing:"0.05em"}}>{progress.label||"Analyzing..."}</div>
                  <div style={{background:"rgba(255,255,255,0.08)",borderRadius:8,height:6,overflow:"hidden",marginBottom:"0.75rem"}}>
                    <div style={{width:(progress.pct||0)+"%",height:"100%",background:BLUE,borderRadius:8,transition:"width 0.4s"}}/>
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
                style={{padding:"0.9rem",borderRadius:10,fontSize:"0.85rem",fontWeight:700,fontFamily:"inherit",border:"none",cursor:file?"pointer":"not-allowed",letterSpacing:"0.02em",transition:"all 0.2s",background:file?BLUE:"rgba(255,255,255,0.06)",color:file?"#fff":"rgba(255,255,255,0.2)",boxShadow:file?"0 4px 20px rgba(74,134,200,0.4)":"none"}}>
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

            {/* Feature pills */}
            {!file&&!isRunning&&(
              <div style={{display:"flex",justifyContent:"center",gap:"0.5rem",flexWrap:"wrap"}}>
                {["Elevations","Soffits","Returns","Excel Export","Evidence PDF"].map(tag=>(
                  <div key={tag} style={{padding:"0.3rem 0.75rem",borderRadius:20,background:"rgba(74,134,200,0.1)",border:"1px solid rgba(74,134,200,0.2)",fontSize:"0.68rem",color:"rgba(74,134,200,0.8)",fontWeight:500}}>{tag}</div>
                ))}
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

            {/* Legend */}
            {results.legend.length>0&&(
              <div>
                <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>Legend</div>
                <div style={{display:"flex",flexDirection:"column",gap:"0.3rem"}}>
                  {results.legend.map((m,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:"0.5rem"}}>
                      <div style={{width:10,height:10,borderRadius:3,background:MAT_COLORS[m.category]||"#9CA3AF",flexShrink:0}}/>
                      <span style={{fontSize:"0.68rem",color:"#475569"}}><span style={{fontWeight:600,color:"#0F172A"}}>{m.id}</span>: {m.name}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* View toggle */}
            <div>
              <div style={{fontSize:"0.6rem",color:BLUE,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"0.6rem"}}>View</div>
              <div style={{display:"flex",flexDirection:"column",gap:"0.4rem"}}>
                {[["table","📊  Table View"],["interactive","🎨  Interactive Takeoff"],["edit","✏️  Edit Surfaces"],["pricing","💵  Pricing & Bid"]].map(([mode,label])=>(
                  <button key={mode} onClick={()=>setViewMode(mode)} style={{padding:"0.55rem 0.75rem",borderRadius:7,fontSize:"0.72rem",fontWeight:viewMode===mode?700:400,fontFamily:"inherit",cursor:"pointer",border:"none",textAlign:"left",background:viewMode===mode?BLUE:"#F1F5F9",color:viewMode===mode?"#fff":"#64748B",transition:"all 0.15s"}}>{label}</button>
                ))}
              </div>
            </div>

            {/* Export */}
            <div style={{display:"flex",flexDirection:"column",gap:"0.4rem"}}>
              <button onClick={exportExcel} style={{padding:"0.6rem",background:"#fff",color:BLUE,border:"1.5px solid "+BLUE,borderRadius:7,fontSize:"0.72rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓  Export Excel</button>
              <button onClick={exportPDF} disabled={pdfLoading} style={{padding:"0.6rem",background:"#F8FAFC",color:pdfLoading?"#9CA3AF":"#64748B",border:"1.5px solid #E2E8F0",borderRadius:7,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:pdfLoading?"not-allowed":"pointer"}}>↓  {pdfLoading?"Generating...":"Evidence PDF"}</button>
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
              <div style={{flex:1,overflow:"hidden"}}><InteractiveView results={results} BACKEND={BACKEND} assignments={assignments} setAssignments={setAssignments}/></div>
            )}
            {viewMode==="edit"&&(
              <div style={{flex:1,overflow:"hidden"}}><EditorView results={results} BACKEND={BACKEND}/></div>
            )}
            {viewMode==="table"&&(
              <div style={{flex:1,overflowY:"auto",padding:"1.5rem"}}>
                <div style={{fontSize:"0.65rem",letterSpacing:"0.1em",color:BLUE,textTransform:"uppercase",fontWeight:700,marginBottom:"1rem"}}>Breakdown by Elevation</div>
                {results.takeoffData.map((elev,i)=>{
                  const total=(elev.zones||[]).reduce((s,z)=>s+(z.netArea||0),0);
                  if(total===0)return null;
                  const dispScale=elev.verifiedScale||elev.scale;
                  const scaleSub=elev.scaleSource==="claude_vision"?"read from drawing":elev.scaleSource==="easyocr"?"OCR · title block":elev.scaleSource==="default"?"default — verify!":null;
                  const dims=fmtDims(elev.buildingDimensions);
                  const overshoot=elev.expectedFacadeSF&&total>elev.expectedFacadeSF*1.4;
                  const hasIntel=dispScale||dims||elev.expectedFacadeSF||elev.scheduleOpeningSF>0||overshoot;
                  return <div key={i} style={{background:"#fff",borderRadius:10,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:"1px solid #F1F5F9",marginBottom:"1rem",overflow:"hidden"}}>
                    <div style={{padding:"0.65rem 1rem",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #F1F5F9",background:"#FAFBFC"}}>
                      <span style={{fontSize:"0.85rem",fontWeight:700,color:"#0F172A"}}>{elev.title}</span>
                      <span style={{fontSize:"0.65rem",color:"#94A3B8"}}>{elev.sheetRef} · {Math.round(total).toLocaleString()} SF</span>
                    </div>
                    {hasIntel&&<div style={{display:"flex",flexWrap:"wrap",gap:"0.5rem",padding:"0.6rem 1rem",background:"#F8FAFC",borderBottom:"1px solid #F1F5F9"}}>
                      {dispScale&&<InfoChip label="Scale" value={dispScale} sub={scaleSub} warn={elev.scaleSource==="default"}/>}
                      {dims&&<InfoChip label="Building" value={dims}/>}
                      {elev.expectedFacadeSF&&<InfoChip label="Gross face" value={Math.round(elev.expectedFacadeSF).toLocaleString()+" SF"}/>}
                      {elev.scheduleOpeningSF>0&&<InfoChip label="Openings · schedule" value={Math.round(elev.scheduleOpeningSF).toLocaleString()+" SF"}/>}
                      {overshoot&&<InfoChip label="⚠ Check scale" value="Panel SF > building face" sub="possible scale error" warn={true}/>}
                    </div>}
                    <table style={{width:"100%",borderCollapse:"collapse"}}>
                      <thead><tr style={{background:"#F8FAFC"}}>{["ID","Material","Category","Gross","Openings","Net SF","Adj +15%"].map(h=>(
                        <th key={h} style={{padding:"0.4rem 0.75rem",textAlign:["Gross","Openings","Net SF","Adj +15%"].includes(h)?"right":"left",fontSize:"0.6rem",fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.05em",borderBottom:"1px solid #F1F5F9"}}>{h}</th>
                      ))}</tr></thead>
                      <tbody>{(elev.zones||[]).map((z,zi)=>(
                        <tr key={zi} style={{borderBottom:"1px solid #F8FAFC"}} onMouseEnter={e=>e.currentTarget.style.background="#FAFBFC"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                          <td style={{padding:"0.45rem 0.75rem"}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:2,background:MAT_COLORS[z.category]||"#9CA3AF"}}/><span style={{fontWeight:700,color:BLUE,fontSize:"0.72rem"}}>{z.materialId||"—"}</span></div></td>
                          <td style={{padding:"0.45rem 0.75rem",fontSize:"0.72rem",color:"#374151",fontWeight:500}}>{z.materialName}</td>
                          <td style={{padding:"0.45rem 0.75rem",fontSize:"0.7rem",color:"#9CA3AF"}}>{z.category}</td>
                          <td style={{padding:"0.45rem 0.75rem",textAlign:"right",fontSize:"0.72rem",color:"#64748B"}}>{Math.round(z.grossArea||0).toLocaleString()}</td>
                          <td style={{padding:"0.45rem 0.75rem",textAlign:"right",fontSize:"0.72rem",color:"#9CA3AF"}}>({Math.round(z.totalOpeningArea||0).toLocaleString()})</td>
                          <td style={{padding:"0.45rem 0.75rem",textAlign:"right",fontSize:"0.75rem",fontWeight:600,color:"#374151"}}>{Math.round(z.netArea||0).toLocaleString()}</td>
                          <td style={{padding:"0.45rem 0.75rem",textAlign:"right",fontSize:"0.78rem",fontWeight:700,color:BLUE}}>{Math.round((z.netArea||0)*1.15).toLocaleString()}</td>
                        </tr>
                      ))}</tbody>
                    </table>
                    {(elev.flags||[]).filter(Boolean).length>0&&<div style={{padding:"0.4rem 0.75rem",fontSize:"0.65rem",color:"#92400E",background:"#FFFBEB",borderTop:"1px solid #FEF3C7"}}>⚠ {elev.flags.filter(Boolean).join(" · ")}</div>}
                  </div>;
                })}
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
    </div>
  );
}
