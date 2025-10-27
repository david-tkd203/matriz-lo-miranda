/* inicial.js (relaciones con hojas + filtro por factor + detalle enriquecido + sincronía panel derecho)
   - Lee HOJA INICIAL (fila 3, columnas B..Q)
   - Guarda workbook (WB) para consultar hojas relacionadas al abrir el detalle
   - Filtros en cascada + por factor/estado
   - Offcanvas: busca filas matching (Área/Puesto/Tareas) en hojas por factor y colorea estado
   - Expone getMatrizFilters() y dispara evento 'matrizFiltersChanged' para que el panel derecho se sincronice
*/

const COLS = {
  B: "Área",
  C: "Puesto de trabajo",
  D: "Tareas del puesto de trabajo",
  E: "Horario de funcionamiento",
  F: "Horas extras POR DIA",
  G: "Horas extras POR SEMANA",
  H: "N° Trabajadores Expuestos HOMBRE",
  I: "N° Trabajadores Expuestos MUJER",
  J: "Trabajo repetitivo de miembros superiores.",
  K: "Postura de trabajo estática",
  L: "MMC Levantamiento/Descenso",
  M: "MMC Empuje/Arrastre",
  N: "Manejo manual de pacientes / personas",
  O: "Vibración de cuerpo completo",
  P: "Vibración segmento mano – brazo",
  Q: "Resultado identificación inicial",
};

const RISKS = [
  { key:'J', label:COLS.J, hoja:'Mov. Rep.' },
  { key:'K', label:COLS.K, hoja:'Postura estatica' },
  { key:'L', label:COLS.L, hoja:'MMC Levantamiento-Descenso' },
  { key:'M', label:COLS.M, hoja:'MMC Empuje-Arrastre' },
  { key:'N', label:COLS.N, hoja:'Manejo manual PCTS' },
  { key:'O', label:COLS.O, hoja:'Vibración Cuerpo completo' },
  { key:'P', label:COLS.P, hoja:'Vibración Mano-Brazo' },
];

let RAW_ROWS = [];
let FILTERS = { area:"", puesto:"", tarea:"", factorKey:"", factorState:"" };

// Guardamos workbook y un mapa de hojas para detalle
let WB = null;
let WS_MAP = {}; // nombreNormalizado -> worksheet

const el = (id) => document.getElementById(id);
const norm = (s) => String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\s+/g,'').trim();

window.getMatrizFilters = () => ({...FILTERS}); // para sincronía con panel derecho

document.addEventListener("DOMContentLoaded", () => {
  attemptFetchDefault();
  wireUI();
});

function wireUI(){
  el("fileInput").addEventListener("change", handleFile);

  el("filterArea").addEventListener("change", () => {
    FILTERS.area = el("filterArea").value || "";
    populatePuesto();
    FILTERS.puesto = ""; el("filterPuesto").value = "";
    populateTarea();
    FILTERS.tarea = ""; el("filterTarea").value = "";
    render();
  });

  el("filterPuesto").addEventListener("change", () => {
    FILTERS.puesto = el("filterPuesto").value || "";
    populateTarea();
    FILTERS.tarea = ""; el("filterTarea").value = "";
    render();
  });

  el("filterTarea").addEventListener("change", () => {
    FILTERS.tarea = el("filterTarea").value || "";
    render();
  });

  // Filtro por factor
  populateFactor();
  el("filterFactor").addEventListener("change", () => {
    FILTERS.factorKey = el("filterFactor").value || "";
    render();
  });
  el("filterFactorState").addEventListener("change", () => {
    FILTERS.factorState = el("filterFactorState").value || "";
    render();
  });

  // Botones
  el("btnReset")?.addEventListener("click", (e) => {
    e.preventDefault();
    FILTERS = { area:"", puesto:"", tarea:"", factorKey:"", factorState:"" };
    ["filterArea","filterPuesto","filterTarea","filterFactor","filterFactorState"].forEach(id => { const s=el(id); if(s) s.value=""; });
    populatePuesto(true); populateTarea(true);
    render();
  });
  el("btnReload")?.addEventListener("click", attemptFetchDefault);

  // Abrir offcanvas al clicar card
  el("cardsWrap").addEventListener("click", (ev) => {
    const card = ev.target.closest("[data-idx]");
    if(!card) return;
    const idx = Number(card.dataset.idx);
    if(Number.isFinite(idx) && RAW_ROWS[idx]) openDetail(RAW_ROWS[idx]);
  });
}

async function attemptFetchDefault(){
  if(!window.DEFAULT_XLSX_PATH) return;
  try{
    const res = await fetch(window.DEFAULT_XLSX_PATH, { cache:"no-store" });
    if(!res.ok) throw new Error("Fetch failed");
    const buf = await res.arrayBuffer();
    processWorkbook(buf);
  }catch(e){
    console.warn("No se pudo cargar automáticamente el Excel. Seleccione el archivo manualmente.", e);
  }
}

function handleFile(evt){
  const file = evt.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = (e) => processWorkbook(e.target.result);
  reader.readAsArrayBuffer(file);
}

function pickInicialSheet(wb){
  const target = (wb.SheetNames || []).find(n => /inicial|inicio/i.test(String(n||"")));
  return target || wb.SheetNames[0];
}

function processWorkbook(arrayBuffer){
  WB = XLSX.read(arrayBuffer, { type:"array" });

  // Mapa de hojas normalizado
  WS_MAP = {};
  (WB.SheetNames||[]).forEach(name => { WS_MAP[norm(name)] = WB.Sheets[name]; });

  const initialSheetName = pickInicialSheet(WB);
  const ws = WB.Sheets[initialSheetName];

  if(!ws || !ws['!ref']){
    RAW_ROWS = []; render(); return;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];

  for(let r = 2; r <= range.e.r; r++){ // fila 3 (idx 2)
    const vals = {};
    const getCell = (c) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      const v = cell ? (cell.w ?? cell.v) : "";
      return v == null ? "" : String(v).trim();
    };
    const colMap = { B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16 };
    for(const [k, idx] of Object.entries(colMap)) vals[k] = getCell(idx);

    if(!(vals.B || vals.C || vals.D)) continue;

    // Normalización SI/NO
    ['J','K','L','M','N','O','P'].forEach(k => {
      const up = String(vals[k]||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
      if(["SI","YES","Y","S"].includes(up)) vals[k] = "SI";
      else if(["NO","N"].includes(up)) vals[k] = "NO";
    });

    // Q calculado si falta
    const allNo = ['J','K','L','M','N','O','P'].every(k => (vals[k]||"").toUpperCase()==="NO");
    const anyPresent = ['J','K','L','M','N','O','P'].some(k => (vals[k]||"")!=="");
    vals.Q = vals.Q || (allNo && anyPresent
      ? "Ausencia total del riesgo, reevaluar cada 3 años con nueva identificación inicial"
      : "Aplicar identificación avanzada-condición aceptable para cada tipo de factor de riesgo identificado");

    rows.push(vals);
  }

  RAW_ROWS = rows;
  populateArea(); populatePuesto(true); populateTarea(true);
  render();
}

/* ===== Filtros ===== */
function uniqueSorted(arr){
  return [...new Set(arr.filter(v => v && String(v).trim() !== ""))].sort((a,b)=> String(a).localeCompare(String(b)));
}
function populateArea(){
  const sel = el("filterArea");
  const opts = uniqueSorted(RAW_ROWS.map(r => r.B));
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = false;
}
function populatePuesto(){
  const sel = el("filterPuesto");
  let list = RAW_ROWS; if(FILTERS.area) list = list.filter(r => r.B===FILTERS.area);
  const opts = uniqueSorted(list.map(r => r.C));
  sel.innerHTML = `<option value="">(Todos)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length===0;
}
function populateTarea(){
  const sel = el("filterTarea");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B===FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C===FILTERS.puesto);
  const opts = uniqueSorted(list.map(r => r.D));
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length===0;
}
function populateFactor(){
  const sel = el("filterFactor");
  sel.innerHTML = `<option value="">(Todos)</option>` + RISKS.map(r => `<option value="${r.key}">${escapeHtml(r.label)}</option>`).join("");
}

function filteredRows(){
  let list = RAW_ROWS.slice();
  if(FILTERS.area)   list = list.filter(r => r.B===FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C===FILTERS.puesto);
  if(FILTERS.tarea)  list = list.filter(r => r.D===FILTERS.tarea);

  if(FILTERS.factorKey){
    const key = FILTERS.factorKey, state = (FILTERS.factorState||"").toUpperCase();
    list = list.filter(r => {
      const v = (r[key]||"").toUpperCase();
      if(state==="SI") return v==="SI";
      if(state==="NO") return v==="NO";
      return v==="SI"||v==="NO"; // si hay factor, excluye N/A
    });
  }
  return list;
}

/* ===== Render ===== */
function render(){
  const target = el("cardsWrap");
  const data = filteredRows();
  el("countRows").textContent = data.length;
  if(data.length === 0){
    target.innerHTML = `<div class="col-12"><div class="alert alert-warning mb-0"><i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.</div></div>`;
  }else{
    target.innerHTML = data.map((r,idx) => cardHtml(r, RAW_ROWS.indexOf(r))).join("");
  }

  // sincronizar panel derecho
  document.dispatchEvent(new CustomEvent("matrizFiltersChanged", { detail: { filters: getMatrizFilters() }}));
}

function cardHtml(r, idx){
  const badges = RISKS.map(x => {
    const v = (r[x.key]||"").toString().toUpperCase();
    const cls = v==="SI" ? "badge-yes" : (v==="NO" ? "badge-no" : "badge-na");
    return `<span class="badge ${cls}" title="${escapeHtml(x.hoja)}">${escapeHtml(x.label)}: ${v||"N/A"}</span>`;
  }).join(" ");

  const needsAdvanced = needsAdvancedEval(r);
  const resultClass = needsAdvanced ? "alert-danger" : ((r.Q||"").startsWith("Ausencia total") ? "alert-success" : "alert-primary");

  return `
    <div class="col-12 col-md-6 col-lg-4">
      <div class="card card-ficha h-100 shadow-sm" role="button" tabindex="0" data-idx="${idx}">
        <div class="card-body">
          <div class="d-flex align-items-start justify-content-between">
            <div>
              <div class="small text-muted mb-1">Área</div>
              <h5 class="title mb-2">${escapeHtml(r.B||"-")}</h5>
            </div>
            <div class="text-end">
              <span class="chip"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
            </div>
          </div>
          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C||"-")}</div>
          <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D||"-")}</div>

          <div class="row g-2 small">
            <div class="col-6"><i class="bi bi-clock"></i> <strong>Horario:</strong> ${escapeHtml(r.E||"-")}</div>
            <div class="col-6"><i class="bi bi-plus-circle"></i> <strong>HE/Día:</strong> ${escapeHtml(r.F||"0")}</div>
            <div class="col-6"><i class="bi bi-plus-circle-dotted"></i> <strong>HE/Semana:</strong> ${escapeHtml(r.G||"0")}</div>
          </div>

          <hr>
          <div class="small"><strong>Factores (J–P):</strong></div>
          <div class="factors-wrap">${badges}</div>

          <div class="alert ${resultClass} mt-2 mb-0" role="alert">
            <i class="bi bi-clipboard-check"></i> <strong>Resultado:</strong> ${escapeHtml(r.Q||"-")}
          </div>
        </div>
      </div>
    </div>
  `;
}

/* ===== Detalle con relación a hojas ===== */

function openDetail(r){
  const body = el("detailBody");
  const adv = needsAdvancedEval(r);
  const mmcWarn = (r.L==="SI" || r.M==="SI");

  const tilesHtml = RISKS.map(x => {
    // estado básico (SI/NO/N/A)
    const flag = (r[x.key]||"").toString().toUpperCase();
    // buscar en hoja relacionada
    const rel = lookupRelated(x.hoja, r); // {state:'ok|warn|risk', comment, metrics:[{k,v}]}
    const cClass = rel.state==="risk" ? "hl-risk" : rel.state==="warn" ? "hl-warn" : "hl-ok";
    const dot = rel.state==="risk" ? "m-risk" : rel.state==="warn" ? "m-warn" : "m-ok";
    const line = (rel.comment||"").trim() || (flag==="SI" ? "Se detecta factor. Evaluar en profundidad." : (flag==="NO" ? "Sin hallazgo aparente." : "Sin información."));

    const metrics = (rel.metrics||[]).map(m => `<div class="h-metric"><span class="m-dot ${dot}"></span> ${escapeHtml(m.k)}: <strong>${escapeHtml(m.v)}</strong></div>`).join("");

    return `
      <div class="adv-tile ${cClass}">
        <div class="h-name">${escapeHtml(x.label)}</div>
        <div class="h-sheet">${escapeHtml(x.hoja)}</div>
        <div class="meter"><span class="m-dot ${dot}"></span> ${escapeHtml(line)}</div>
        ${metrics ? `<div class="mt-1">${metrics}</div>` : ""}
      </div>
    `;
  }).join("");

  const advAlert = adv
    ? `<div class="alert alert-danger"><i class="bi bi-x-octagon"></i> <strong>Corresponde identificación avanzada.</strong> Uno o más factores presentan condición no aceptable.</div>`
    : `<div class="alert alert-success"><i class="bi bi-check2-circle"></i> No se identifican condiciones que requieran avanzada.</div>`;

  const mmcAlert = mmcWarn
    ? `<div class="alert alert-warning mmc-alert"><i class="bi bi-exclamation-triangle"></i> Atención MMC:
        ${ r.L==="SI" ? "<span class='pill warn'>Levantamiento/Descenso: SI</span> " : "" }
        ${ r.M==="SI" ? "<span class='pill warn'>Empuje/Arrastre: SI</span>" : "" }
      </div>` : "";

  body.innerHTML = `
    <div class="d-flex justify-content-between align-items-start">
      <div>
        <div class="small text-muted">Área</div>
        <h5 class="mb-1">${escapeHtml(r.B || "-")}</h5>
        <div class="mb-2"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
        <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>
        <div class="d-flex gap-2 flex-wrap">
          <span class="pill"><i class="bi bi-clock"></i> ${escapeHtml(r.E || "-")}</span>
          <span class="pill"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
        </div>
      </div>
    </div>

    <hr>
    ${advAlert}
    ${mmcAlert}

    <h6 class="mt-3">Resultados detallados por factor</h6>
    <div class="d-grid" style="grid-template-columns:1fr; gap:.55rem;">
      ${tilesHtml}
    </div>

    <hr>
    <div class="small text-muted">Resultado hoja inicial</div>
    <div>${escapeHtml(r.Q||"-")}</div>
  `;

  el("detailTitle").textContent = `Detalle · ${r.B || "-"}`;
  bootstrap.Offcanvas.getOrCreateInstance('#detailPanel').show();
}

/* Heurística: si cualquier factor está "SI" o Q sugiere avanzada */
function needsAdvancedEval(r){
  const q = (r.Q||"").toLowerCase();
  const qSaysAdv = q.includes("aplicar identificación avanzada") || q.includes("aplicar identificacion avanzada");
  const anyYes = ['J','K','L','M','N','O','P'].some(k => (r[k]||"").toUpperCase()==="SI");
  return qSaysAdv || anyYes;
}

/* ===== Relación a hojas: búsqueda flexible por cabeceras ===== */

function sheetToObjectsFlexible(ws){
  // Lee como 2D, detecta fila de headers (buscando columnas tipo 'Area','Puesto','Tareas')
  const rows2D = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
  if(!rows2D.length) return [];

  // normalizador
  const n = (s)=>String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();

  let headerRowIdx = 0, headers = rows2D[0].map(String);
  const score = (arr) => {
    const set = new Set(arr.map(n));
    let s=0;
    ["area","puesto","tareas","tarea"].forEach(k=>{ if(set.has(k)) s++; });
    return s;
  };
  let best = score(headers);
  for(let i=1;i<Math.min(rows2D.length,8);i++){
    const sc = score(rows2D[i].map(String));
    if(sc>best){ best=sc; headerRowIdx=i; headers = rows2D[i].map(String); }
  }
  const normHeaders = headers.map(h => n(h).replace(/\s+/g,'_').replace(/[^\w]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,''));

  const objs = [];
  for(let r=headerRowIdx+1; r<rows2D.length; r++){
    const row = rows2D[r];
    if(!row || row.every(v => (v==null || String(v).trim()===""))) continue;
    const obj = {};
    normHeaders.forEach((k, idx)=> obj[k] = row[idx]==null ? "" : row[idx]);
    objs.push(obj);
  }
  return objs;
}

function lookupRelated(sheetDisplayName, r){
  try{
    const ws = WS_MAP[norm(sheetDisplayName)] ||
               // tolerancia por diferencias leves
               Object.entries(WS_MAP).find(([k]) => k.includes(norm(sheetDisplayName)))?.[1];
    if(!ws) return { state: baseFromFlag((r||{})), comment:"Hoja no encontrada" };

    const rows = sheetToObjectsFlexible(ws); // objetos con claves normalizadas
    if(!rows.length) return { state:"ok", comment:"Sin registros" };

    const targetArea  = norm(r.B);
    const targetPuesto= norm(r.C);
    const targetTarea = norm(r.D);

    // detectar campos clave en la hoja
    const pickKey = (obj, aliases)=> {
      for(const a of aliases){
        if(a in obj) return a;
      }
      return null;
    };

    // Usa la primera fila para saber campos existentes
    const sample = rows[0];
    const kArea   = pickKey(sample, ["area"]);
    const kPuesto = pickKey(sample, ["puesto","puesto_de_trabajo","cargo"]);
    const kTarea  = pickKey(sample, ["tareas","tarea","tareas_del_puesto"]);

    // filtrar matching (tolerante: si falta una clave la ignora)
    let matches = rows.filter(o => {
      const okA = !kArea   || norm(o[kArea])   === targetArea;
      const okP = !kPuesto || norm(o[kPuesto]) === targetPuesto;
      const okT = !kTarea  || norm(o[kTarea])  === targetTarea;
      return okA && okP && okT;
    });

    if(!matches.length){
      // fallback: solo por área
      matches = rows.filter(o => kArea && norm(o[kArea])===targetArea);
      if(!matches.length) return baseFromFlag(r);
    }

    // extraer campos "resultado"/"nivel"/"riesgo" (heurística)
    const toPairs = (o)=> Object.entries(o).map(([k,v])=>({k, v:String(v)}));
    const m = matches[0]; // primera coincidencia
    const pairs = toPairs(m);

    const resPair = pairs.find(p => /resultado|aceptable|no_?aceptable|avanzad/i.test(p.k)) || 
                    pairs.find(p => /nivel|riesgo|severidad/i.test(p.k));
    let state = "ok", comment = "", metrics = [];

    if(resPair){
      const val = (resPair.v||"").toLowerCase();
      if(/no\s*aceptable|avanzad|alto|rojo|riesgo\s*alto/.test(val)) { state="risk"; }
      else if(/moderado|precauc|medio|amarill/.test(val))           { state="warn"; }
      else if(/aceptable|bajo|verde|sin\s*riesgo/.test(val))         { state="ok"; }
      comment = `${labelize(resPair.k)}: ${resPair.v}`;
    }else{
      // si no hay columna resultado, miramos algún umbral u observación
      const obs = pairs.find(p => /observaci|coment/i.test(p.k));
      if(obs){ comment = `${labelize(obs.k)}: ${obs.v}`; }
      state = baseFromFlag(r).state;
    }

    // tomar algunas métricas relevantes si existen
    const keepKeys = ["resultado","nivel","riesgo","puntuacion","puntos","score","indice","umbral","categoria","clasificacion"];
    pairs.forEach(p => {
      if(keepKeys.some(k => p.k.includes(k)) && p.k !== (resPair?.k||"")){
        metrics.push({ k: labelize(p.k), v: p.v });
      }
    });

    return { state, comment, metrics };
  }catch(_){
    return baseFromFlag(r);
  }
}

function baseFromFlag(r){
  // Si el campo del factor está SI -> warn por defecto, si NO -> ok
  return { state: needsAdvancedEval(r) ? "risk" : "ok", comment:"" };
}

function labelize(k){
  return String(k||"").replace(/_/g,' ').replace(/\b\w/g, m => m.toUpperCase());
}

/* ===== Utils ===== */
function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}
