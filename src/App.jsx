import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

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

/* ── Excel builder ── */
const buildExcel = (projectName, materials) => {
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString("en-US");
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
    er = n10(); er[0]=idx+1; er[1]=mat.name; er[2]=Math.round(mat.sf); er[4]=1; er[5]=""; er[6]=`=C${row}*F${row}`; amtCells.push(`G${row}`); eRows.push(er);
  });
  while (eRows.length < 24) eRows.push(n10());
  const totalRow = eRows.length + 1;
  er = n10(); er[1]="Total "; er[2]=materials.reduce((s,m)=>s+Math.round(m.sf),0); er[6]=amtCells.length?`=${amtCells.join("+")}`:0; eRows.push(er);
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
  pr=n11(); pr[7]="TOTAL:"; pr[8]=`=Estimate!G${totalRow}`; pRows.push(pr);
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
function InteractiveView({ results, BACKEND }) {
  const [elevIdx, setElevIdx] = useState(0);
  const [pageImage, setPageImage] = useState(null);
  const [imgLoaded, setImgLoaded] = useState(false);
  const [pagePolygons, setPagePolygons] = useState([]);
  const [pageDims, setPageDims] = useState({ width:612, height:792 });
  const [assignments, setAssignments] = useState({});
  const [activeZone, setActiveZone] = useState(null);
  const [imgNaturalSize, setImgNaturalSize] = useState({ w:1, h:1 });
  const imgRef = useRef();
  const elevations = results.takeoffData.filter(e => e.pageNumber);
  const elev = elevations[elevIdx];
  const pageNum = elev?.pageNumber;

  useEffect(() => {
    if (!pageNum || !results.jobId) return;
    setPageImage(null); setImgLoaded(false); setPagePolygons([]); setActiveZone(null);
    fetch(BACKEND+"/polygons/"+results.jobId+"/"+pageNum)
      .then(r=>r.ok?r.json():{polygons:[],width:612,height:792})
      .then(d=>{setPagePolygons(d.polygons||[]);setPageDims({width:d.width||612,height:d.height||792});})
      .catch(()=>{});
    setPageImage(BACKEND+"/page-image/"+results.jobId+"/"+pageNum);
  }, [elevIdx, pageNum, results.jobId, BACKEND]);

  const polyMethod = pagePolygons[0]?.source||(pagePolygons.length>0?"vector":"box");
  const displayZones = pagePolygons.length>0 ? pagePolygons : (elev?.zones||[]).map((z,i)=>({
    id:i,points:[[z.x0pct/100,z.y0pct/100],[z.x1pct/100,z.y0pct/100],[z.x1pct/100,z.y1pct/100],[z.x0pct/100,z.y1pct/100]],
    area_sf:z.netArea||0,cx:(z.x0pct+z.x1pct)/200,cy:(z.y0pct+z.y1pct)/200,source:"box",
  }));
  const colorGroups = {};
  displayZones.forEach(z=>{const k=z.cluster_id!==undefined?"c_"+z.cluster_id:z.fill_color?"f_"+z.fill_color.join(","):"none";if(!colorGroups[k])colorGroups[k]=[];colorGroups[k].push(z.id);});
  const clusterSummary = {};
  displayZones.forEach(z=>{const k=z.cluster_id!==undefined?z.cluster_id:-1;if(!clusterSummary[k])clusterSummary[k]={total_sf:0,count:0,color:"#94A3B8"};clusterSummary[k].total_sf+=z.area_sf||0;clusterSummary[k].count+=1;clusterSummary[k].color=z.cluster_id!==undefined?CLUSTER_COLORS[z.cluster_id%CLUSTER_COLORS.length]:"#94A3B8";});
  const assignKey = id => elevIdx+":"+id;
  const getAssignment = id => assignments[assignKey(id)];
  const assignZone = (zoneId, mat) => {
    const zone=displayZones.find(z=>z.id===zoneId);
    setAssignments(prev=>({...prev,[assignKey(zoneId)]:{...mat,area_sf:zone?.area_sf||0}}));
    setActiveZone(null);
  };
  const removeAssignment = zoneId => { const k=assignKey(zoneId); setAssignments(prev=>{const n={...prev};delete n[k];return n;}); };
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
        <div style={{alignSelf:"flex-start",fontSize:"0.65rem",color:"#64748B",background:NAVY_LT,padding:"0.3rem 0.75rem",borderRadius:20,border:"1px solid #2D5280"}}>
          {polyMethod==="bluebeam"?`📐 Bluebeam — ${displayZones.length} surfaces`:polyMethod==="vector_cluster"||polyMethod==="vector"?`📏 Vector — ${displayZones.length} surfaces`:polyMethod==="claude_vision"?`🧠 AI Vision — ${displayZones.length} surfaces`:"No surfaces on this page"}
        </div>
        {!pageImage?<div style={{color:"#475569",fontSize:"0.8rem",marginTop:"4rem"}}>Loading elevation...</div>:
          <div style={{position:"relative",display:"inline-block",maxWidth:"100%"}}>
            <img ref={imgRef} src={pageImage} alt={elev?.title} onLoad={e=>{setImgNaturalSize({w:e.target.naturalWidth,h:e.target.naturalHeight});setImgLoaded(true);}} style={{display:"block",maxWidth:"100%",maxHeight:"calc(100vh - 180px)",objectFit:"contain",borderRadius:6,border:"1px solid "+NAVY_LT}}/>
            {imgLoaded&&<svg viewBox={`0 0 ${pageDims.width} ${pageDims.height}`} style={{position:"absolute",top:0,left:0,width:svgW,height:svgH,overflow:"visible"}}>
              {displayZones.map(zone=>{
                const a=getAssignment(zone.id);
                let color="#94A3B8";
                if(a)color=MAT_COLORS[a.category]||"#9CA3AF";
                else if(zone.cluster_id!==undefined)color=CLUSTER_COLORS[zone.cluster_id%CLUSTER_COLORS.length];
                else if(zone.fill_color?.length===3){const[r,g,b]=zone.fill_color;color=`rgb(${Math.round(r*255)},${Math.round(g*255)},${Math.round(b*255)})`;}
                const isActive=activeZone===zone.id;
                const pts=toSVGPoints(zone.points);
                const lx=zone.cx*pageDims.width,ly=zone.cy*pageDims.height;
                const showLabel=a||isActive||zone.source==="bluebeam"||zone.source==="claude_vision";
                return <g key={zone.id} style={{cursor:"pointer"}} onClick={e=>{e.stopPropagation();setActiveZone(isActive?null:zone.id);}}>
                  <polygon points={pts} fill={color} fillOpacity={isActive?0.65:zone.source==="bluebeam"?0.45:a?0.38:0.2} stroke={isActive?"#fff":color} strokeWidth={isActive?2.5:1.5} strokeOpacity={0.9}/>
                  {showLabel&&<><rect x={lx-32} y={ly-9} width={64} height={18} fill="rgba(0,0,0,0.8)" rx={4}/><text x={lx} y={ly+2} textAnchor="middle" dominantBaseline="middle" fill="white" fontSize={pageDims.width/75} fontFamily="Inter,Arial" fontWeight="bold">{zone.source==="claude_vision"&&!a?zone.material_type:Math.round(zone.area_sf)+" SF"}</text></>}
                </g>;
              })}
            </svg>}
          </div>
        }
      </div>
      {/* Right panel */}
      <div style={{width:230,borderLeft:"1px solid "+NAVY_LT,padding:"1rem",overflowY:"auto",flexShrink:0,background:NAVY_MID}}>
        {activeZone!==null?(
          <>
            <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.5rem"}}>Assign Surface</div>
            <div style={{padding:"0.5rem 0.65rem",marginBottom:"0.75rem",background:NAVY,borderRadius:6,fontSize:"0.7rem",color:"#CBD5E1",border:"1px solid #2D5280"}}>{Math.round(displayZones.find(z=>z.id===activeZone)?.area_sf||0).toLocaleString()} SF selected</div>
            {matList.map((mat,i)=>{
              const azd=displayZones.find(z=>z.id===activeZone);
              const ck=azd?.cluster_id!==undefined?"c_"+azd.cluster_id:azd?.fill_color?"f_"+azd.fill_color.join(","):null;
              const scz=ck?(colorGroups[ck]||[]):[];
              const bulk=scz.length>1;
              return <div key={i} style={{marginBottom:"0.4rem"}}>
                <div onClick={()=>assignZone(activeZone,mat)} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.5rem 0.65rem",cursor:"pointer",background:NAVY,borderRadius:bulk?"6px 6px 0 0":6,border:"1px solid #2D5280",borderBottom:bulk?"none":undefined}} onMouseEnter={e=>e.currentTarget.style.borderColor=BLUE} onMouseLeave={e=>e.currentTarget.style.borderColor="#2D5280"}>
                  <div style={{width:10,height:10,borderRadius:3,background:MAT_COLORS[mat.category]||"#9CA3AF",flexShrink:0}}/>
                  <div style={{fontSize:"0.65rem",color:"#CBD5E1",flex:1}}><span style={{color:BLUE,fontWeight:700}}>{mat.id}</span> {mat.name||mat.category}</div>
                </div>
                {bulk&&<div onClick={()=>{scz.forEach(zid=>{const z=displayZones.find(dz=>dz.id===zid);setAssignments(prev=>({...prev,[elevIdx+":"+zid]:{...mat,area_sf:z?.area_sf||0}}));});setActiveZone(null);}} style={{padding:"0.3rem 0.65rem",background:"#122035",cursor:"pointer",fontSize:"0.62rem",color:"#4ADE80",border:"1px solid #2D5280",borderTop:"none",borderRadius:"0 0 6px 6px"}} onMouseEnter={e=>e.currentTarget.style.color="#86EFAC"} onMouseLeave={e=>e.currentTarget.style.color="#4ADE80"}>↳ Assign all {scz.length} same-color zones</div>}
              </div>;
            })}
            {getAssignment(activeZone)&&<div onClick={()=>removeAssignment(activeZone)} style={{padding:"0.45rem",marginTop:"0.5rem",textAlign:"center",fontSize:"0.65rem",color:"#F87171",cursor:"pointer",border:"1px solid #7F1D1D",borderRadius:6}}>Remove assignment</div>}
            <div onClick={()=>setActiveZone(null)} style={{padding:"0.45rem",marginTop:"0.3rem",textAlign:"center",fontSize:"0.65rem",color:"#64748B",cursor:"pointer",border:"1px solid #2D5280",borderRadius:6}}>Cancel</div>
          </>
        ):(
          <>
            {Object.keys(clusterSummary).length>0&&<>
              <div style={{fontSize:"0.6rem",letterSpacing:"0.12em",color:"#64748B",textTransform:"uppercase",fontWeight:700,marginBottom:"0.4rem"}}>Detected Textures</div>
              <div style={{fontSize:"0.62rem",color:"#475569",marginBottom:"0.65rem"}}>Click a zone → assign material</div>
              {Object.entries(clusterSummary).map(([cid,info])=>{
                const zids=colorGroups["c_"+cid]||[];
                const assigned=zids.filter(id=>getAssignment(id)).length;
                return <div key={cid} style={{marginBottom:"0.4rem",padding:"0.5rem 0.65rem",background:NAVY,borderRadius:6,borderLeft:"3px solid "+info.color}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"0.4rem"}}><div style={{width:8,height:8,borderRadius:2,background:info.color}}/><span style={{fontSize:"0.65rem",color:"#94A3B8"}}>Texture {parseInt(cid)+1}</span></div>
                    <span style={{fontSize:"0.8rem",fontWeight:700,color:"#E2E8F0"}}>{Math.round(info.total_sf).toLocaleString()}</span>
                  </div>
                  <div style={{fontSize:"0.6rem",color:assigned>0?"#4ADE80":"#475569",marginTop:3}}>{assigned>0?`✓ ${assigned}/${info.count} assigned`:`${info.count} zone${info.count!==1?"s":""}`}</div>
                </div>;
              })}
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
            {Object.keys(clusterSummary).length===0&&Object.keys(totals).length===0&&<div style={{fontSize:"0.7rem",color:"#475569",lineHeight:1.8}}>Click any colored zone to assign a material type.</div>}
            <div style={{marginTop:"1rem",fontSize:"0.6rem",color:"#334155"}}>{Object.keys(assignments).length} zones · {elevations.length} elevations</div>
          </>
        )}
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
  const fileRef  = useRef();
  const logRef   = useRef();
  const pollRef  = useRef(null);
  const seenLogs = useRef(0);

  useEffect(()=>{ if(logRef.current) logRef.current.scrollTop=logRef.current.scrollHeight; },[log]);
  useEffect(()=>()=>{ if(pollRef.current) clearInterval(pollRef.current); },[]);

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
          setResults({legend:data.legend||[],takeoffData:data.takeoffData||[],projName:file?.name?.replace(".pdf","")||"Project",jobId:id});
          setPhase("done");setProgress({label:"Complete",pct:100});
        }else if(data.status==="error"){clearInterval(pollRef.current);setErrMsg(data.error||"Unknown error");setPhase("error");}
      }catch(e){console.log("poll",e.message);}
    },5000);
  },[file]);

  const run = async()=>{
    if(!file)return;
    setPhase("running");setLog([]);setErrMsg("");setResults(null);seenLogs.current=0;
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

  const summary=results?()=>{
    const t={};
    results.takeoffData.forEach(e=>(e.zones||[]).forEach(z=>{const k=z.category||"Other";if(!t[k])t[k]={net:0,adj:0,color:MAT_COLORS[k]||"#9CA3AF"};t[k].net+=z.netArea||0;t[k].adj+=(z.netArea||0)*1.15;}));
    return t;
  }:null;
  const summaryData = summary ? summary() : null;
  const grandAdj = summaryData ? Object.values(summaryData).reduce((s,v)=>s+v.adj,0) : 0;
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
            <button onClick={()=>{setFile(null);setPhase("idle");setResults(null);setLog([]);}} style={{padding:"0.45rem 1rem",background:"rgba(255,255,255,0.08)",color:"rgba(255,255,255,0.6)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>↺ New</button>
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
                {[["table","📊  Table View"],["interactive","🎨  Interactive Takeoff"]].map(([mode,label])=>(
                  <button key={mode} onClick={()=>setViewMode(mode)} style={{padding:"0.55rem 0.75rem",borderRadius:7,fontSize:"0.72rem",fontWeight:viewMode===mode?700:400,fontFamily:"inherit",cursor:"pointer",border:"none",textAlign:"left",background:viewMode===mode?BLUE:"#F1F5F9",color:viewMode===mode?"#fff":"#64748B",transition:"all 0.15s"}}>{label}</button>
                ))}
              </div>
            </div>

            {/* Export */}
            <div style={{display:"flex",flexDirection:"column",gap:"0.4rem"}}>
              <button onClick={exportExcel} style={{padding:"0.6rem",background:"#fff",color:BLUE,border:"1.5px solid "+BLUE,borderRadius:7,fontSize:"0.72rem",fontWeight:700,fontFamily:"inherit",cursor:"pointer"}}>↓  Export Excel</button>
              <button onClick={exportPDF} disabled={pdfLoading} style={{padding:"0.6rem",background:"#F8FAFC",color:pdfLoading?"#9CA3AF":"#64748B",border:"1.5px solid #E2E8F0",borderRadius:7,fontSize:"0.72rem",fontWeight:600,fontFamily:"inherit",cursor:pdfLoading?"not-allowed":"pointer"}}>↓  {pdfLoading?"Generating...":"Evidence PDF"}</button>
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
              <div style={{flex:1,overflow:"hidden"}}><InteractiveView results={results} BACKEND={BACKEND}/></div>
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
          </main>
        </div>
      )}
    </div>
  );
}
