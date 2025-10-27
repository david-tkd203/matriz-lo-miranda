/* inicial.js (paginación “slide”, 1 tarjeta por fila, modal fullscreen)
   - Lee HOJA INICIAL (B..Q) y oculta “0” en Área/Puesto
   - Une con hoja “Movimiento repetitivo” y muestra estado basado en columnas P y W
   - Card: muestra estado (badge). Modal fullscreen: detalles P y W + resto de preguntas/respuestas disponibles.
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
  { key: 'J', label: COLS.J, hoja: 'Mov. Rep.' },
  { key: 'K', label: COLS.K, hoja: 'Postura estatica' },
  { key: 'L', label: COLS.L, hoja: 'MMC Levantamiento-Descenso' },
  { key: 'M', label: COLS.M, hoja: 'MMC Empuje-Arrastre' },
  { key: 'N', label: COLS.N, hoja: 'Manejo manual PCTS' },
  { key: 'O', label: COLS.O, hoja: 'Vibración Cuerpo completo' },
  { key: 'P', label: COLS.P, hoja: 'Vibración Mano-Brazo' },
];

let RAW_ROWS = [];
let MOVREP_MAP = Object.create(null); // key -> {P, W, rowObj, headers}
let MOVREP_HEADERS = [];

let FILTERS = { area: "", puesto: "", tarea: "", factorKey: "", factorState: "" };
let STATE = { page:1, perPage:10, pageMax:1 };

const el = (id) => document.getElementById(id);

/* ======= Helpers ======= */
function escapeHtml(str){
  return String(str ?? "").replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
function toLowerNoAccents(s){
  return String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}
function isZeroish(v){
  if(v===0) return true;
  const s=String(v??"").trim();
  return !!s && /^0(\.0+)?$/.test(s);
}
function keyTriple(area, puesto, tarea){
  return `${toLowerNoAccents(area)}|${toLowerNoAccents(puesto)}|${toLowerNoAccents(tarea)}`;
}

/* ======= Bootstrap ======= */
document.addEventListener("DOMContentLoaded", () => {
  attemptFetchDefault();
  wireUI();
});

/* ======= UI ======= */
function wireUI(){
  el("fileInput").addEventListener("change", handleFile);

  el("filterArea").addEventListener("change", () => {
    FILTERS.area = el("filterArea").value || "";
    populatePuesto();
    FILTERS.puesto = "";
    el("filterPuesto").value = "";
    populateTarea();
    FILTERS.tarea = "";
    el("filterTarea").value = "";
    STATE.page = 1;
    render();
  });
  el("filterPuesto").addEventListener("change", () => {
    FILTERS.puesto = el("filterPuesto").value || "";
    populateTarea();
    FILTERS.tarea = "";
    el("filterTarea").value = "";
    STATE.page = 1;
    render();
  });
  el("filterTarea").addEventListener("change", () => {
    FILTERS.tarea = el("filterTarea").value || "";
    STATE.page = 1;
    render();
  });

  populateFactor();
  el("filterFactor").addEventListener("change", () => {
    FILTERS.factorKey = el("filterFactor").value || "";
    STATE.page = 1;
    render();
  });
  el("filterFactorState").addEventListener("change", () => {
    FILTERS.factorState = el("filterFactorState").value || "";
    STATE.page = 1;
    render();
  });

  el("btnReset").addEventListener("click", (e) => {
    e.preventDefault();
    FILTERS = { area: "", puesto: "", tarea: "", factorKey:"", factorState:"" };
    el("filterArea").value = "";
    el("filterPuesto").value = "";
    el("filterTarea").value = "";
    el("filterFactor").value = "";
    el("filterFactorState").value = "";
    populatePuesto(true);
    populateTarea(true);
    STATE.page = 1;
    render();
  });
  el("btnReload").addEventListener("click", attemptFetchDefault);

  el("btnPrev").addEventListener("click", ()=>{ if(STATE.page>1){ STATE.page--; render(); window.scrollTo({top:0,behavior:'smooth'});} });
  el("btnNext").addEventListener("click", ()=>{ if(STATE.page<STATE.pageMax){ STATE.page++; render(); window.scrollTo({top:0,behavior:'smooth'});} });
  el("perPage").addEventListener("change", ()=>{ STATE.perPage = parseInt(el("perPage").value,10)||10; STATE.page=1; render(); });
  el("btnTop").addEventListener("click", ()=> window.scrollTo({top:0,behavior:'smooth'}));

  // Click en tarjeta → modal fullscreen
  el("cardsWrap").addEventListener("click", (ev) => {
    const open = ev.target.closest("[data-open]");
    const card = ev.target.closest("[data-idx]");
    if(open && card){
      const idx = Number(card.dataset.idx);
      if(Number.isFinite(idx) && RAW_ROWS[idx]) openDetail(RAW_ROWS[idx]);
    }
  });
}

/* ======= Carga de archivo por defecto ======= */
async function attemptFetchDefault(){
  if(!window.DEFAULT_XLSX_PATH) return;
  try{
    const res = await fetch(window.DEFAULT_XLSX_PATH + `?v=${Date.now()}`, {cache:"no-store"});
    if(!res.ok){ throw new Error("Fetch failed"); }
    const buf = await res.arrayBuffer();
    processWorkbook(buf);
  }catch(e){
    console.warn("No se pudo cargar el Excel por defecto. Seleccione manualmente.", e);
  }
}
function handleFile(evt){
  const file = evt.target.files?.[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = (e) => processWorkbook(e.target.result);
  reader.readAsArrayBuffer(file);
}

/* ======= Parse libro ======= */
function pickInicialSheet(wb){
  const target = (wb.SheetNames || []).find(n => /inicial|inicio/i.test(String(n||"")));
  return target || wb.SheetNames[0];
}
function pickMovRepSheet(wb){
  const cand = wb.SheetNames.find(n => /mov|repet/i.test(n.toLowerCase()));
  return cand || null;
}

function processWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  /* Hoja INICIAL */
  const initialSheetName = pickInicialSheet(wb);
  const ws = wb.Sheets[initialSheetName];

  RAW_ROWS = [];
  if(ws && ws['!ref']){
    const range = XLSX.utils.decode_range(ws['!ref']);
    for(let r = 2; r <= range.e.r; r++){ // fila 3 visible
      const vals = {};
      function getCell(c){
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        const value = cell ? (cell.w ?? cell.v) : "";
        return value == null ? "" : String(value).trim();
      }
      const colMap = { B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16 };
      for(const [k, idx] of Object.entries(colMap)){ vals[k] = getCell(idx); }

      if(!(vals.B || vals.C || vals.D)) continue;

      // Normaliza SI/NO
      ['J','K','L','M','N','O','P'].forEach(k => {
        if(vals[k]){
          const up = vals[k].toString().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
          if(["SI","YES","Y","S"].includes(up)) vals[k] = "SI";
          else if(["NO","N"].includes(up)) vals[k] = "NO";
        }
      });

      // Resultado por defecto si viene vacío
      const allNo = ['J','K','L','M','N','O','P'].every(k => (vals[k] || "").toUpperCase() === "NO");
      const anyPresent = ['J','K','L','M','N','O','P'].some(k => (vals[k] || "") !== "");
      const Qcalc = (allNo && anyPresent)
        ? "Ausencia total del riesgo, reevaluar cada 3 años con nueva identificación inicial"
        : "Aplicar identificación avanzada-condición aceptable para cada tipo de factor de riesgo identificado";
      vals.Q = vals.Q || Qcalc;

      // Filtra falsos positivos "0"
      if(isZeroish(vals.B) || isZeroish(vals.C) || isZeroish(vals.D)) continue;

      RAW_ROWS.push(vals);
    }
  }

  /* Hoja Movimiento repetitivo */
  MOVREP_MAP = Object.create(null);
  MOVREP_HEADERS = [];
  const movSheetName = pickMovRepSheet(wb);
  if(movSheetName){
    const wsMov = wb.Sheets[movSheetName];
    if(wsMov){
      const rows2d = XLSX.utils.sheet_to_json(wsMov, { header:1, defval:"" });
      if(rows2d.length){
        const headers = rows2d[0].map(h => String(h||""));
        MOVREP_HEADERS = headers;

        // Intentamos ubicar índices de Área/Puesto/Tarea por nombre; si no, asumimos B,C,D
        const idxArea  = findIndexInsensitive(headers, ["area","área"]) ?? 1;
        const idxPuesto= findIndexInsensitive(headers, ["puesto"]) ?? 2;
        const idxTarea = findIndexInsensitive(headers, ["tarea","tareas"]) ?? 3;

        // Columnas P (letra 16) y W (letra 22) por requerimiento
        const idxP = 16; // A=0
        const idxW = 22;

        for(let i=1;i<rows2d.length;i++){
          const r = rows2d[i] || [];
          const area  = r[idxArea] ?? "";
          const puesto= r[idxPuesto] ?? "";
          const tarea = r[idxTarea] ?? "";
          if(!(area||puesto||tarea)) continue;
          const k = keyTriple(area, puesto, tarea);

          const rec = objectFromRow(headers, r);
          MOVREP_MAP[k] = {
            P: r[idxP] ?? "",
            W: r[idxW] ?? "",
            rowObj: rec
          };
        }
      }
    }
  }

  populateArea();
  populatePuesto(true);
  populateTarea(true);
  render();
}

function findIndexInsensitive(headers, keys){
  const norm = headers.map(h => toLowerNoAccents(String(h||"")));
  for(const k of keys){
    const idx = norm.indexOf(toLowerNoAccents(k));
    if(idx>=0) return idx;
  }
  return null;
}
function objectFromRow(headers, row){
  const o={};
  headers.forEach((h,i)=>{ o[h||`Col${i+1}`] = row[i]??""; });
  return o;
}

/* ======= Filtros ======= */
function uniqueSorted(arr){
  return [...new Set(arr.filter(v => v && String(v).trim() !== ""))].sort((a,b)=> String(a).localeCompare(String(b)));
}
function populateArea(){
  const opts = uniqueSorted(RAW_ROWS.map(r => r.B));
  const sel = el("filterArea");
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = false;
}
function populatePuesto(){
  const sel = el("filterPuesto");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B === FILTERS.area);
  const opts = uniqueSorted(list.map(r => r.C));
  sel.innerHTML = `<option value="">(Todos)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length === 0;
}
function populateTarea(){
  const sel = el("filterTarea");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B === FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C === FILTERS.puesto);
  const opts = uniqueSorted(list.map(r => r.D));
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length === 0;
}
function populateFactor(){
  const sel = el("filterFactor");
  sel.innerHTML = `<option value="">(Todos)</option>` +
    RISKS.map(r => `<option value="${r.key}">${escapeHtml(r.label)}</option>`).join("");
}

function filteredRows(){
  let list = RAW_ROWS.slice();
  if(FILTERS.area)   list = list.filter(r => r.B === FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C === FILTERS.puesto);
  if(FILTERS.tarea)  list = list.filter(r => r.D === FILTERS.tarea);

  if(FILTERS.factorKey){
    const key = FILTERS.factorKey;
    const state = (FILTERS.factorState || "").toUpperCase();
    list = list.filter(r => {
      const v = (r[key] || "").toUpperCase();
      if(state === "SI") return v === "SI";
      if(state === "NO") return v === "NO";
      return v === "SI" || v === "NO";
    });
  }
  return list;
}

/* ======= Render + paginación ======= */
function render(){
  const target = el("cardsWrap");
  const all = filteredRows();

  // Paginación “slide”
  const per = STATE.perPage = parseInt(el("perPage").value,10) || 10;
  STATE.pageMax = Math.max(1, Math.ceil(all.length / per));
  if(STATE.page > STATE.pageMax) STATE.page = STATE.pageMax;
  const start = (STATE.page - 1) * per;
  const pageData = all.slice(start, start + per);

  el("countRows").textContent = pageData.length;
  el("countRowsTotal").textContent = all.length;
  el("pageCur").textContent = STATE.page;
  el("pageMax").textContent = STATE.pageMax;

  if(pageData.length === 0){
    target.innerHTML = `<div class="col"><div class="alert alert-warning mb-0">
      <i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.
    </div></div>`;
    return;
  }

  target.innerHTML = pageData.map((r) => {
    const idx = RAW_ROWS.indexOf(r);
    return cardHtml(r, idx);
  }).join("");
}

/* ======= Estado Movimiento repetitivo (P/W) ======= */
function getMovRepFor(r){
  const k = keyTriple(r.B, r.C, r.D);
  return MOVREP_MAP[k] || null;
}
function classifyMovRep(p, w){
  const s = `${String(p||"")} ${String(w||"")}`;
  const t = toLowerNoAccents(s);
  if(!t.trim()) return { cls:"status-unk", label:"Sin dato" };

  // Heurística: acepta/precaución/no aceptable
  if(t.includes("no acept") || t.includes("alto") || t.includes("riesgo alto") || t.includes("critico") || t.includes("crítico"))
    return { cls:"status-bad", label:"No aceptable" };
  if(t.includes("moderad") || t.includes("medio") || t.includes("precauc") || t.includes("mejorable"))
    return { cls:"status-warn", label:"Precaución" };
  if(t.includes("acept") || t.includes("bajo") || t.includes("sin riesgo"))
    return { cls:"status-ok", label:"Aceptable" };
  return { cls:"status-unk", label:"Revisar" };
}

/* ======= HTML Tarjeta + Modal ======= */
function cardHtml(r, idx){
  const mov = getMovRepFor(r);
  const status = classifyMovRep(mov?.P, mov?.W);

  return `
    <div class="col" data-idx="${idx}">
      <div class="card card-ficha h-100 shadow-sm">
        <div class="card-body">
          <div class="d-flex align-items-start justify-content-between">
            <div>
              <div class="small text-muted mb-1">Área</div>
              <h5 class="title mb-2">${escapeHtml(r.B || "-")}</h5>
            </div>
            <div class="text-end">
              <span class="chip"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
            </div>
          </div>

          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
          <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>

          <div class="d-flex flex-wrap align-items-center gap-2 mt-2">
            <span class="status-pill ${status.cls}" title="Estado según hoja Movimiento repetitivo (P/W)">
              <i class="bi bi-activity"></i> ${status.label}
            </span>
            ${mov ? `<span class="pill"><strong>P:</strong> ${escapeHtml(mov.P??"")}</span>
                     <span class="pill"><strong>W:</strong> ${escapeHtml(mov.W??"")}</span>` 
                 : `<span class="pill">Hoja Mov. repetitivo: sin coincidencia</span>`}
          </div>

          <hr>
          <div class="row g-2 small">
            <div class="col-6"><i class="bi bi-clock"></i> <strong>Horario:</strong> ${escapeHtml(r.E || "-")}</div>
            <div class="col-6"><i class="bi bi-plus-circle"></i> <strong>HE/Día:</strong> ${escapeHtml(r.F || "0")}</div>
            <div class="col-6"><i class="bi bi-plus-circle-dotted"></i> <strong>HE/Semana:</strong> ${escapeHtml(r.G || "0")}</div>
          </div>

          <div class="d-flex justify-content-end mt-3">
            <button type="button" class="btn btn-primary btn-sm btn-open" data-open>
              <i class="bi bi-arrows-fullscreen"></i> Ver detalles
            </button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function openDetail(r){
  const mov = getMovRepFor(r);
  const status = classifyMovRep(mov?.P, mov?.W);

  const header = `
    <div class="detail-card mb-3">
      <div class="d-flex flex-wrap justify-content-between align-items-start gap-2">
        <div>
          <div class="small text-muted">Área</div>
          <h5 class="mb-1">${escapeHtml(r.B || "-")}</h5>
          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
          <div class="mb-1"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>
        </div>
        <div class="d-flex flex-column align-items-end gap-2">
          <span class="status-pill ${status.cls}" style="font-size:1rem;">
            <i class="bi bi-activity"></i> ${status.label}
          </span>
          ${mov ? `<div class="d-flex gap-2">
                    <span class="pill"><strong>P:</strong> ${escapeHtml(mov.P??"")}</span>
                    <span class="pill"><strong>W:</strong> ${escapeHtml(mov.W??"")}</span>
                   </div>` : ``}
        </div>
      </div>
    </div>
  `;

  // Tabla de preguntas/respuestas de la fila completa de “Mov. Repetitivo”
  let qa = "";
  if(mov && mov.rowObj){
    const entries = Object.entries(mov.rowObj)
      .filter(([k,v]) => String(v).trim() !== "");
    const headHtml = `<thead><tr><th style="min-width:220px">Pregunta</th><th>Respuesta</th></tr></thead>`;
    const bodyHtml = `<tbody>${
      entries.map(([k,v]) => `<tr><th>${escapeHtml(k)}</th><td>${escapeHtml(String(v))}</td></tr>`).join("")
    }</tbody>`;

    qa = `
      <div class="table-like">
        <table>${headHtml}${bodyHtml}</table>
      </div>
    `;
  }else{
    qa = `<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> No se encontraron detalles coincidentes en la hoja “Movimiento repetitivo”.</div>`;
  }

  el("detailBody").innerHTML = `
    ${header}
    <div class="detail-grid">
      ${qa}
    </div>
  `;
  el("detailTitle").textContent = `Detalle · Movimiento repetitivo`;
  const modal = bootstrap.Modal.getOrCreateInstance('#detailModal');
  modal.show();
}

/* ======= Utils ======= */
function escapeCSV(str){ return `"${String(str??"").replace(/"/g,'""')}"`; }
