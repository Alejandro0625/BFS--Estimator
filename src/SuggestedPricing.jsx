// BFS Suggested Pricing — learns the $/SF we WON at from past winning-bid Excels and
// suggests a price range per material for new takeoffs. Fully self-contained module:
// App.jsx only imports and renders it — nothing here writes into Estimator state, and a
// crash inside this tab is caught by the Guard below (the rest of the app keeps working).
import { useState, useRef, Component } from "react";
import * as XLSX from "xlsx";

const BLUE = "#1B4F8A";
const LIB_KEY = "bfs_pricing_library";

/* Canonical pricing families (order matters: most specific first, so
   "James Hardie Siding with AZEK Trims" lands in Fiber Cement, not Trim). */
const FAMILIES = [
  ["Nichiha", /nichiha/i],
  ["Perforated Metal", /perforated|\bperf\b/i],
  ["Soffit", /soffit/i],
  ["Fiber Cement", /fiber|cement|hardie|artisan|swisspearl|cembrit/i],
  ["ACM / Metal Panel", /\bacm\b|\bmcm\b|\bacp\b|composite|metl[\s-]?span|insulated metal|\bimp\b|metal\s*(wall\s*)?panel|rib panel|corrugated/i],
  ["Aluminum Wall Panel", /aluminum|aluminium/i],
  ["Returns / Trim", /return|\btrim\b|fascia|coping|azek|\bpvc\b|flashing|watertable/i],
  ["Lap Siding", /\blap\b|clapboard|siding|shake|shingle|board\s*(&|and)\s*batten|longboard|woodtone/i],
];
const FAMILY_NAMES = [...FAMILIES.map(f => f[0]), "Other"];
const famOf = (name) => (FAMILIES.find(([, re]) => re.test(String(name || ""))) || ["Other"])[0];

const loadLib = () => { try { const j = JSON.parse(localStorage.getItem(LIB_KEY) || "null"); return Array.isArray(j?.bids) ? j.bids : []; } catch { return []; } };
const persistLib = (bids) => { try { localStorage.setItem(LIB_KEY, JSON.stringify({ v: 1, bids })); } catch {} };

/* Newer wins count more: half-life 2 years (a 2-year-old win weighs half a fresh one). */
const recencyW = (b) => {
  const t = b.date ? Date.parse(b.date) : (b.addedAt || Date.now());
  const yrs = Math.max(0, (Date.now() - (isNaN(t) ? Date.now() : t)) / (365.25 * 24 * 3600 * 1000));
  return Math.pow(0.5, yrs / 2);
};
const wQuantile = (pairs, q) => {   // pairs [{v,w}] sorted by v ascending
  const tot = pairs.reduce((s, p) => s + p.w, 0); if (!tot) return null;
  let acc = 0; for (const p of pairs) { acc += p.w; if (acc / tot >= q) return p.v; }
  return pairs[pairs.length - 1].v;
};
/* Suggestion = recency- and size-weighted quartiles of the won rates in this family.
   Size weight: a past job the same size as this takeoff counts full; 10× bigger/smaller counts half. */
const suggestFor = (bids, family, unit, targetQty) => {
  const hits = bids.filter(b => b.family === family && (b.unit || "sf") === unit && b.rate > 0);
  if (!hits.length) return null;
  const pairs = hits.map(b => {
    let w = recencyW(b);
    if (targetQty > 0 && b.qty > 0) w *= 1 / (1 + Math.abs(Math.log10(b.qty / targetQty)));
    return { v: b.rate, w, b };
  }).sort((a, c) => a.v - c.v);
  return {
    n: hits.length,
    low: wQuantile(pairs, 0.25), mid: wQuantile(pairs, 0.5), high: wQuantile(pairs, 0.75),
    latest: hits.reduce((s, b) => Math.max(s, b.date ? new Date(b.date).getFullYear() : 0), 0) || null,
    evidence: pairs.slice().sort((a, c) => c.w - a.w).map(p => p.b),
  };
};

const money = (v) => "$" + Math.round(v || 0).toLocaleString();
const rate2 = (v) => (v == null ? "—" : "$" + (+v).toFixed(2));
const dedupeKey = (b) => [b.job, b.material, b.unit, b.qty, b.rate].join("|").toLowerCase();

/* ── Excel → learned entries (same /api/analyze proxy the Scope tab uses; prompt
      validated on a real winning bid: derives rate from total ÷ SF even in prose notes) ── */
const readSheetText = async (f) => {
  const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
  return wb.SheetNames.map(n => `# Sheet: ${n}\n` + XLSX.utils.sheet_to_csv(wb.Sheets[n])).join("\n\n");
};
const extractWins = async (f) => {
  const text = await readSheetText(f);
  const prompt = `You are reading a WINNING BID spreadsheet from Boston Facade Systems (a facade/siding subcontractor). This bid WON — extract the price per unit we won at, per cladding material, so future bids can be priced from history.

Return ONLY JSON (no markdown):
{
 "job":"project name (from the sheet or the file name)",
 "date":"YYYY-MM-DD bid date if findable in the sheet or file name, else null",
 "entries":[
   {"material":"material as named in the sheet (e.g. James Hardie siding, ACM panel)",
    "unit":"sf|lf",
    "qty":12345,
    "total":250000,
    "rate":20.25,
    "confidence":"high|low",
    "note":"one short line: where in the sheet this came from"}
 ]
}
Rules:
- entries = CLADDING/FACADE materials only (metal panel, ACM/MCM, fiber cement, Nichiha, lap siding, aluminum wall panel, perforated panel, soffit, trim/returns, siding). Skip permits, engineering, tax, bonds.
- rate = the WON price per unit ($/SF or $/LF). If the sheet gives a total price and a quantity, rate = total/qty. If the only usable numbers are a project total and one total SF (even inside a prose note), return ONE entry from those.
- NEVER invent a number that is not in the sheet. If a material has no price+quantity pair, skip it or mark confidence "low".
- qty is in the entry's unit; total is dollars.

FILE NAME: ${f.name}

SPREADSHEET:
${text.slice(0, 30000)}`;
  const res = await fetch("/api/analyze", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ model: "claude-opus-4-8", max_tokens: 4000, messages: [{ role: "user", content: [{ type: "text", text: prompt }] }] }) });
  if (!res.ok) throw new Error("analysis service error (" + res.status + ")");
  const data = await res.json();
  if (data?.error) throw new Error(data.error.message || "analysis error");
  let t = data?.content?.find(b => b.type === "text")?.text || "";
  const m = t.match(/\{[\s\S]+\}/); if (m) t = m[0];
  const j = JSON.parse(t);
  return (j.entries || []).map((e, i) => ({
    id: Date.now() + "-" + i + "-" + Math.random().toString(36).slice(2, 7),
    job: j.job || f.name, date: j.date || null,
    material: e.material || "Material", family: famOf(e.material),
    unit: e.unit === "lf" ? "lf" : "sf",
    qty: +e.qty || 0, total: +e.total || 0,
    rate: +e.rate > 0 ? +(+e.rate).toFixed(2) : (+e.total > 0 && +e.qty > 0 ? +(e.total / e.qty).toFixed(2) : 0),
    confidence: e.confidence === "low" ? "low" : "high", note: e.note || "", file: f.name, addedAt: Date.now(),
  }));
};

/* ── shared small styles ── */
const card = { background: "#fff", borderRadius: 12, border: "1px solid #EEF2F7", padding: "1.2rem 1.4rem", marginBottom: "1.2rem" };
const secLabel = { fontSize: "0.6rem", letterSpacing: "0.12em", color: BLUE, textTransform: "uppercase", fontWeight: 700, marginBottom: "0.6rem" };
const inp = { padding: "0.4rem 0.55rem", border: "1px solid #E2E8F0", borderRadius: 7, fontSize: "0.78rem", fontFamily: "inherit", width: "100%", boxSizing: "border-box" };
const btn = (primary) => ({ padding: "0.5rem 1rem", background: primary ? "linear-gradient(180deg,#5A92D2,#3F79BC)" : "#fff", color: primary ? "#fff" : "#64748B", border: primary ? "none" : "1px solid #E2E8F0", borderRadius: 7, fontSize: "0.74rem", fontWeight: 700, fontFamily: "inherit", cursor: "pointer" });
const th = { textAlign: "left", padding: "0.45rem 0.6rem", fontSize: "0.62rem", color: "#64748B", textTransform: "uppercase", letterSpacing: "0.06em", fontWeight: 700, borderBottom: "1px solid #EEF2F7" };
const td = { padding: "0.5rem 0.6rem", fontSize: "0.8rem", color: "#0F172A", borderBottom: "1px solid #F1F5F9", verticalAlign: "middle" };

/* Crash shield: an error anywhere in this tab renders HERE, never white-screens the app. */
class Guard extends Component {
  constructor(p) { super(p); this.state = { err: null }; }
  static getDerivedStateFromError(e) { return { err: e }; }
  render() {
    if (!this.state.err) return this.props.children;
    return (
      <div style={{ flex: 1, padding: "3rem", textAlign: "center" }}>
        <div style={{ fontSize: "1rem", fontWeight: 800, color: "#0F172A", marginBottom: 8 }}>Suggested Pricing hit an error</div>
        <div style={{ fontSize: "0.8rem", color: "#94A3B8", marginBottom: 16 }}>{String(this.state.err?.message || this.state.err)} — the rest of the app is unaffected.</div>
        <button onClick={() => this.setState({ err: null })} style={btn(true)}>Try again</button>
      </div>
    );
  }
}

function Inner({ results = null, priceRows = [], linearRollup = [], buildExcel = null }) {
  const [lib, setLibState] = useState(loadLib);
  const [pending, setPending] = useState([]);      // parsed wins awaiting the estimator's confirm
  const [busyFile, setBusyFile] = useState("");
  const [errs, setErrs] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [rateEdit, setRateEdit] = useState({});     // estimator's overrides of suggested rates
  const [marginPct, setMarginPct] = useState(15);
  const [wastePct, setWastePct] = useState(10);
  const [openFam, setOpenFam] = useState({});       // expanded evidence per family
  const [manual, setManual] = useState({ job: "", material: "", unit: "sf", qty: "", total: "", rate: "", date: "" });
  const fileRef = useRef(); const importRef = useRef();
  const setLib = (bids) => { setLibState(bids); persistLib(bids); };

  /* ingest winning-bid files (sequential; each error reported, others still land) */
  const onFiles = async (files) => {
    const list = Array.from(files || []).filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
    if (!list.length) { setErrs(e => [...e, "Drop Excel files (.xlsx / .xls / .csv)"]); return; }
    setErrs([]);
    for (const f of list) {
      setBusyFile(f.name);
      try {
        const entries = await extractWins(f);
        if (!entries.length) setErrs(e => [...e, `${f.name}: no priced cladding materials found`]);
        setPending(p => [...p, ...entries]);
      } catch (ex) { setErrs(e => [...e, `${f.name}: ${ex.message}`]); }
    }
    setBusyFile("");
  };
  const updPending = (id, k, v) => setPending(p => p.map(r => {
    if (r.id !== id) return r;
    const n = { ...r, [k]: v };
    if (k === "material") n.family = famOf(v);
    if ((k === "qty" || k === "total") && +n.total > 0 && +n.qty > 0) n.rate = +(+n.total / +n.qty).toFixed(2);
    return n;
  }));
  const confirmPending = () => {
    const seen = new Set(lib.map(dedupeKey));
    const fresh = pending.filter(b => +b.rate > 0 && !seen.has(dedupeKey(b)))
      .map(b => ({ ...b, qty: +b.qty || 0, total: +b.total || 0, rate: +b.rate }));
    setLib([...lib, ...fresh]);
    setPending([]);
    if (fresh.length < pending.length) setErrs(e => [...e, `${pending.length - fresh.length} entr${pending.length - fresh.length > 1 ? "ies" : "y"} skipped (duplicate or no $/unit)`]);
  };
  const addManual = () => {
    const qty = +manual.qty || 0, total = +manual.total || 0;
    const rate = +manual.rate > 0 ? +manual.rate : (qty > 0 && total > 0 ? +(total / qty).toFixed(2) : 0);
    if (!(rate > 0)) { setErrs(e => [...e, "Manual entry needs a $/unit rate (or total + quantity)"]); return; }
    setLib([...lib, {
      id: Date.now() + "-m-" + Math.random().toString(36).slice(2, 7),
      job: manual.job || "Manual entry", date: manual.date || null,
      material: manual.material || "Material", family: famOf(manual.material),
      unit: manual.unit, qty, total, rate, confidence: "high", note: "entered by hand", file: "", addedAt: Date.now(),
    }]);
    setManual({ job: "", material: "", unit: "sf", qty: "", total: "", rate: "", date: "" });
  };
  const exportJson = () => {
    const blob = new Blob([JSON.stringify({ v: 1, bids: lib }, null, 1)], { type: "application/json" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
    a.download = "BFS_pricing_library.json"; a.click(); URL.revokeObjectURL(a.href);
  };
  const importJson = async (f) => {
    try {
      const j = JSON.parse(await f.text());
      const bids = Array.isArray(j?.bids) ? j.bids : (Array.isArray(j) ? j : []);
      const seen = new Set(lib.map(dedupeKey));
      const fresh = bids.filter(b => b && +b.rate > 0 && !seen.has(dedupeKey(b)));
      setLib([...lib, ...fresh]);
      setErrs([`Imported ${fresh.length} win${fresh.length === 1 ? "" : "s"} (${bids.length - fresh.length} already known)`]);
    } catch { setErrs(e => [...e, "Couldn't read that library file"]); }
  };

  /* suggestions for the CURRENT takeoff (read-only view of Estimator data) */
  const targetSF = priceRows.reduce((s, r) => s + (r.net || 0), 0);
  const sfRows = priceRows.map(r => {
    const family = famOf(r.cat);
    const sug = suggestFor(lib, family, "sf", r.net);
    const key = "sf:" + r.cat;
    const rate = rateEdit[key] !== undefined ? rateEdit[key] : (sug ? sug.mid.toFixed(2) : "");
    return { key, name: r.cat, family, qty: r.net, unit: "sf", sug, rate, amount: (r.net || 0) * (1 + wastePct / 100) * (+rate || 0) };
  });
  const lfRows = linearRollup.map(it => {
    const family = famOf(it.material);
    const sug = suggestFor(lib, family, "lf", it.lf);
    const key = "lf:" + it.material;
    const rate = rateEdit[key] !== undefined ? rateEdit[key] : (sug ? sug.mid.toFixed(2) : "");
    return { key, name: it.material, family, qty: it.lf, unit: "lf", sug, rate, amount: (it.lf || 0) * (+rate || 0) };
  });
  const allRows = [...sfRows, ...lfRows];
  const subtotal = allRows.reduce((s, r) => s + r.amount, 0);
  const total = subtotal * (1 + marginPct / 100);
  const missingRate = allRows.some(r => !(+r.rate > 0));

  const exportBid = () => {
    if (!buildExcel || !results) return;
    const w = 1 + wastePct / 100;   // buildExcel re-applies waste to every qty; pre-divide LF so only SF carries it
    const mats = [...sfRows.map(r => ({ name: r.name, sf: r.qty, rate: +r.rate || 0 })),
                  ...lfRows.map(r => ({ name: r.name + " (per LF)", sf: (r.qty || 0) / w, rate: +r.rate || 0 }))];
    const wb = buildExcel(results.projName || "Project", mats, { wastePct, marginPct });
    XLSX.writeFile(wb, "BFS_SuggestedBid_" + (results.projName || "Project").replace(/\s+/g, "_") + ".xlsx");
  };

  /* library grouped by family for display */
  const byFam = FAMILY_NAMES.map(f => ({ fam: f, bids: lib.filter(b => b.family === f).sort((a, c) => (c.date || "").localeCompare(a.date || "")) })).filter(g => g.bids.length);

  return (
    <div style={{ flex: 1, overflowY: "auto", padding: "2rem", background: "#F6F8FB" }}>
      <div style={{ maxWidth: 1020, margin: "0 auto" }}>
        <div style={{ fontSize: "0.7rem", letterSpacing: "0.18em", color: BLUE, textTransform: "uppercase", fontWeight: 700, marginBottom: "0.3rem" }}>Suggested Pricing</div>
        <h2 style={{ fontSize: "1.5rem", fontWeight: 800, color: "#0F172A", margin: "0 0 0.3rem", letterSpacing: "-0.02em" }}>Price from the jobs you already won</h2>
        <p style={{ fontSize: "0.82rem", color: "#64748B", margin: "0 0 1.5rem", lineHeight: 1.6 }}>
          Upload past <b>winning bid</b> Excels once — the system learns the $/SF you won at for each material.
          New takeoffs get a suggested range from similar wins (newer wins count more). You always confirm what it learns and can edit every price before the bid sheet is written.
        </p>

        {errs.map((e, i) => <div key={i} style={{ padding: "0.5rem 0.8rem", background: "#FFFBEB", border: "1px solid #FDE68A", color: "#92400E", borderRadius: 8, fontSize: "0.75rem", marginBottom: "0.5rem" }}>{e}</div>)}

        {/* ── 1 · Suggested prices for the current takeoff ── */}
        <div style={card}>
          <div style={secLabel}>1 · This takeoff — suggested prices</div>
          {(!results || !allRows.length) ? (
            <div style={{ padding: "1.6rem", textAlign: "center", color: "#94A3B8", fontSize: "0.82rem" }}>
              Run a takeoff in the Takeoff tab — its materials will appear here with a suggested $/SF from your winning history.
            </div>
          ) : (<>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={th}>Material</th><th style={th}>Takeoff qty</th><th style={th}>Won history</th><th style={{ ...th, width: 110 }}>Your price</th><th style={{ ...th, textAlign: "right" }}>Amount</th></tr></thead>
              <tbody>
                {allRows.map(r => (
                  <tr key={r.key}>
                    <td style={td}><b>{r.name}</b><div style={{ fontSize: "0.62rem", color: "#94A3B8" }}>{r.family}</div></td>
                    <td style={td}>{Math.round(r.qty).toLocaleString()} {r.unit.toUpperCase()}</td>
                    <td style={td}>{r.sug ? (
                      <span title={"suggested " + rate2(r.sug.mid)}>
                        <b style={{ color: BLUE }}>{rate2(r.sug.low)} – {rate2(r.sug.high)}</b>
                        <span style={{ fontSize: "0.66rem", color: "#94A3B8" }}> /{r.unit} · {r.sug.n} win{r.sug.n > 1 ? "s" : ""}{r.sug.latest ? " · newest " + r.sug.latest : ""}</span>
                      </span>
                    ) : <span style={{ color: "#CBD5E1", fontSize: "0.72rem" }}>no history yet</span>}</td>
                    <td style={td}><input value={r.rate} onChange={e => setRateEdit(o => ({ ...o, [r.key]: e.target.value }))} placeholder="$/unit" style={{ ...inp, fontWeight: 700, color: rateEdit[r.key] !== undefined ? "#0F172A" : BLUE }} /></td>
                    <td style={{ ...td, textAlign: "right", fontWeight: 700 }}>{money(r.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ display: "flex", alignItems: "center", gap: "1.2rem", marginTop: "1rem", flexWrap: "wrap" }}>
              <label style={{ fontSize: "0.72rem", color: "#64748B", fontWeight: 600 }}>Waste % <input value={wastePct} onChange={e => setWastePct(+e.target.value || 0)} style={{ ...inp, width: 64, marginLeft: 6 }} /></label>
              <label style={{ fontSize: "0.72rem", color: "#64748B", fontWeight: 600 }}>Margin % <input value={marginPct} onChange={e => setMarginPct(+e.target.value || 0)} style={{ ...inp, width: 64, marginLeft: 6 }} /></label>
              <div style={{ marginLeft: "auto", textAlign: "right" }}>
                <div style={{ fontSize: "0.7rem", color: "#64748B" }}>Subtotal {money(subtotal)} · Margin {money(total - subtotal)}</div>
                <div style={{ fontSize: "1.35rem", fontWeight: 800, color: "#0F172A" }}>TOTAL {money(total)}</div>
              </div>
              <button onClick={exportBid} disabled={missingRate} title={missingRate ? "Every line needs a price first" : ""} style={{ ...btn(true), opacity: missingRate ? 0.45 : 1, cursor: missingRate ? "not-allowed" : "pointer" }}>↓ Priced Bid Sheet (Excel)</button>
            </div>
            <div style={{ fontSize: "0.62rem", color: "#94A3B8", marginTop: "0.5rem" }}>Blue prices are suggestions from your winning history — edit any of them. The Excel uses the same BFS estimate template as the Budget tab.</div>
          </>)}
        </div>

        {/* ── 2 · Learn new wins ── */}
        <div style={card}>
          <div style={secLabel}>2 · Teach it your winning bids</div>
          <div
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={e => { e.preventDefault(); setDragOver(false); onFiles(e.dataTransfer.files); }}
            onClick={() => fileRef.current?.click()}
            style={{ border: "2px dashed " + (dragOver ? BLUE : "#CBD5E1"), background: dragOver ? "#EBF2FA" : "#FAFCFE", borderRadius: 10, padding: "1.4rem", textAlign: "center", cursor: "pointer" }}>
            <div style={{ fontSize: "0.85rem", fontWeight: 700, color: "#0F172A" }}>{busyFile ? "🧠 Reading " + busyFile + "…" : "Drop winning bid Excel files here (or click)"}</div>
            <div style={{ fontSize: "0.68rem", color: "#94A3B8", marginTop: 4 }}>.xlsx · .xls · .csv — several at once is fine. Nothing is learned until you confirm below.</div>
          </div>
          <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => { onFiles(e.target.files); e.target.value = ""; }} />

          {pending.length > 0 && (<>
            <div style={{ ...secLabel, marginTop: "1.1rem" }}>Check what it read — fix anything, then confirm</div>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={th}>Job</th><th style={th}>Material</th><th style={th}>Family</th><th style={th}>Unit</th><th style={th}>Qty</th><th style={th}>Total $</th><th style={th}>$ / unit</th><th style={th}>Date</th><th style={th}></th></tr></thead>
              <tbody>
                {pending.map(b => (
                  <tr key={b.id} style={{ background: b.confidence === "low" ? "#FFFBEB" : "transparent" }}>
                    <td style={td}><input value={b.job} onChange={e => updPending(b.id, "job", e.target.value)} style={inp} /></td>
                    <td style={td}><input value={b.material} onChange={e => updPending(b.id, "material", e.target.value)} style={inp} title={b.note} /></td>
                    <td style={td}><select value={b.family} onChange={e => updPending(b.id, "family", e.target.value)} style={inp}>{FAMILY_NAMES.map(f => <option key={f}>{f}</option>)}</select></td>
                    <td style={td}><select value={b.unit} onChange={e => updPending(b.id, "unit", e.target.value)} style={inp}><option value="sf">SF</option><option value="lf">LF</option></select></td>
                    <td style={td}><input value={b.qty} onChange={e => updPending(b.id, "qty", e.target.value)} style={{ ...inp, width: 84 }} /></td>
                    <td style={td}><input value={b.total} onChange={e => updPending(b.id, "total", e.target.value)} style={{ ...inp, width: 96 }} /></td>
                    <td style={td}><input value={b.rate} onChange={e => updPending(b.id, "rate", e.target.value)} style={{ ...inp, width: 70, fontWeight: 700 }} /></td>
                    <td style={td}><input value={b.date || ""} onChange={e => updPending(b.id, "date", e.target.value)} placeholder="YYYY-MM-DD" style={{ ...inp, width: 100 }} /></td>
                    <td style={td}><button onClick={() => setPending(p => p.filter(x => x.id !== b.id))} style={{ ...btn(false), padding: "0.3rem 0.55rem" }}>✕</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ display: "flex", gap: "0.6rem", marginTop: "0.8rem" }}>
              <button onClick={confirmPending} style={btn(true)}>✓ Add {pending.length} to the library</button>
              <button onClick={() => setPending([])} style={btn(false)}>Discard</button>
            </div>
          </>)}

          {/* manual entry — for wins that only exist on paper */}
          <div style={{ ...secLabel, marginTop: "1.1rem" }}>Or add a win by hand</div>
          <div style={{ display: "flex", gap: "0.45rem", flexWrap: "wrap", alignItems: "center" }}>
            <input value={manual.job} onChange={e => setManual(m => ({ ...m, job: e.target.value }))} placeholder="Job name" style={{ ...inp, width: 160 }} />
            <input value={manual.material} onChange={e => setManual(m => ({ ...m, material: e.target.value }))} placeholder="Material (e.g. ACM panel)" style={{ ...inp, width: 180 }} />
            <select value={manual.unit} onChange={e => setManual(m => ({ ...m, unit: e.target.value }))} style={{ ...inp, width: 64 }}><option value="sf">SF</option><option value="lf">LF</option></select>
            <input value={manual.qty} onChange={e => setManual(m => ({ ...m, qty: e.target.value }))} placeholder="Qty" style={{ ...inp, width: 80 }} />
            <input value={manual.total} onChange={e => setManual(m => ({ ...m, total: e.target.value }))} placeholder="Total $" style={{ ...inp, width: 96 }} />
            <input value={manual.rate} onChange={e => setManual(m => ({ ...m, rate: e.target.value }))} placeholder="$/unit" style={{ ...inp, width: 76 }} />
            <input value={manual.date} onChange={e => setManual(m => ({ ...m, date: e.target.value }))} placeholder="YYYY-MM-DD" style={{ ...inp, width: 104 }} />
            <button onClick={addManual} style={btn(false)}>+ Add</button>
          </div>
        </div>

        {/* ── 3 · The library ── */}
        <div style={card}>
          <div style={{ display: "flex", alignItems: "center", gap: "0.6rem" }}>
            <div style={{ ...secLabel, marginBottom: 0, flex: 1 }}>3 · Winning-bid library — {lib.length} win{lib.length === 1 ? "" : "s"}</div>
            <button onClick={exportJson} disabled={!lib.length} style={{ ...btn(false), opacity: lib.length ? 1 : 0.4 }}>↓ Share library</button>
            <button onClick={() => importRef.current?.click()} style={btn(false)}>↑ Import library</button>
            <input ref={importRef} type="file" accept=".json" style={{ display: "none" }} onChange={e => { if (e.target.files?.[0]) importJson(e.target.files[0]); e.target.value = ""; }} />
          </div>
          {!lib.length ? (
            <div style={{ padding: "1.2rem 0 0.4rem", color: "#94A3B8", fontSize: "0.8rem" }}>Nothing learned yet — upload winning bids above. The library lives in this browser; use Share / Import to move it between the two estimators' computers.</div>
          ) : byFam.map(g => {
            const s = suggestFor(lib, g.fam, "sf", 0) || suggestFor(lib, g.fam, "lf", 0);
            return (
              <div key={g.fam} style={{ marginTop: "0.9rem" }}>
                <div onClick={() => setOpenFam(o => ({ ...o, [g.fam]: !o[g.fam] }))} style={{ display: "flex", alignItems: "center", gap: "0.6rem", cursor: "pointer", padding: "0.55rem 0.7rem", background: "#F8FAFC", borderRadius: 8, border: "1px solid #EEF2F7" }}>
                  <span style={{ fontWeight: 800, fontSize: "0.82rem", color: "#0F172A" }}>{g.fam}</span>
                  <span style={{ fontSize: "0.68rem", color: "#64748B" }}>{g.bids.length} win{g.bids.length > 1 ? "s" : ""}</span>
                  {s && <span style={{ marginLeft: "auto", fontSize: "0.74rem", fontWeight: 700, color: BLUE }}>{rate2(s.low)} – {rate2(s.high)}</span>}
                  <span style={{ color: "#94A3B8", fontSize: "0.7rem" }}>{openFam[g.fam] ? "▲" : "▼"}</span>
                </div>
                {openFam[g.fam] && (
                  <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 4 }}>
                    <tbody>
                      {g.bids.map(b => (
                        <tr key={b.id}>
                          <td style={td}><b>{b.job}</b><div style={{ fontSize: "0.62rem", color: "#94A3B8" }}>{b.material}{b.file ? " · " + b.file : ""}</div></td>
                          <td style={td}>{b.date || "no date"}</td>
                          <td style={td}>{b.qty ? Math.round(b.qty).toLocaleString() + " " + (b.unit || "sf").toUpperCase() : "—"}</td>
                          <td style={td}>{b.total ? money(b.total) : "—"}</td>
                          <td style={{ ...td, fontWeight: 800, color: BLUE }}>{rate2(b.rate)}/{b.unit || "sf"}</td>
                          <td style={{ ...td, textAlign: "right" }}><button onClick={() => setLib(lib.filter(x => x.id !== b.id))} style={{ ...btn(false), padding: "0.28rem 0.5rem" }}>✕</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

export default function SuggestedPricingView(props) {
  return <Guard><Inner {...props} /></Guard>;
}
