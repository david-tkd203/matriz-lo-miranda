/* inicial.js (ampliado con filtro por factor + panel detalle)
   - Lee HOJA INICIAL (fila 3, columnas B..Q)
   - Filtros en cascada y por Factor (con estado SI/NO/Todos)
   - Cards clicables → Offcanvas con detalle, mostrando si aplica evaluación avanzada (rojo)
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
let FILTERS = { area: "", puesto: "", tarea: "", factorKey: "", factorState: "" };

const el = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", () => {
  attemptFetchDefault();
  wireUI();
});

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
    render();
  });
  el("filterPuesto").addEventListener("change", () => {
    FILTERS.puesto = el("filterPuesto").value || "";
    populateTarea();
    FILTERS.tarea = "";
    el("filterTarea").value = "";
    render();
  });
  el("filterTarea").addEventListener("change", () => {
    FILTERS.tarea = el("filterTarea").value || "";
    render();
  });

  // NUEVO: Filtro por factor y estado
  populateFactor();
  el("filterFactor").addEventListener("change", () => {
    FILTERS.factorKey = el("filterFactor").value || "";
    render();
  });
  el("filterFactorState").addEventListener("change", () => {
    FILTERS.factorState = el("filterFactorState").value || "";
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
    render();
  });
  el("btnReload").addEventListener("click", attemptFetchDefault);

  // Delegación: abrir detalle al click en card
  el("cardsWrap").addEventListener("click", (ev) => {
    const card = ev.target.closest("[data-idx]");
    if(!card) return;
    const idx = Number(card.dataset.idx);
    if(Number.isFinite(idx) && RAW_ROWS[idx]){
      openDetail(RAW_ROWS[idx]);
    }
  });
}

async function attemptFetchDefault(){
  if(!window.DEFAULT_XLSX_PATH) return;
  try{
    const res = await fetch(window.DEFAULT_XLSX_PATH);
    if(!res.ok){ throw new Error("Fetch failed"); }
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
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const initialSheetName = pickInicialSheet(wb);
  const ws = wb.Sheets[initialSheetName];

  if(!ws || !ws['!ref']){
    RAW_ROWS = [];
    render();
    return;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];

  for(let r = 2; r <= range.e.r; r++){ // fila 3 (idx 2)
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

    ['J','K','L','M','N','O','P'].forEach(k => {
      if(vals[k]){
        const up = vals[k].toString().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
        if(["SI","YES","Y","S"].includes(up)) vals[k] = "SI";
        else if(["NO","N"].includes(up)) vals[k] = "NO";
      }
    });

    const allNo = ['J','K','L','M','N','O','P'].every(k => (vals[k] || "").toUpperCase() === "NO");
    const anyPresent = ['J','K','L','M','N','O','P'].some(k => (vals[k] || "") !== "");
    const Qcalc = (allNo && anyPresent)
      ? "Ausencia total del riesgo, reevaluar cada 3 años con nueva identificación inicial"
      : "Aplicar identificación avanzada-condición aceptable para cada tipo de factor de riesgo identificado";
    vals.Q = vals.Q || Qcalc;

    rows.push(vals);
  }

  RAW_ROWS = rows;
  populateArea();
  populatePuesto(true);
  populateTarea(true);
  render();
}

/* ====== Poblar filtros ====== */
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

/* ====== Filtro compuesto ====== */
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
      return v === "SI" || v === "NO"; // si elige factor pero estado = (Todos) excluimos N/A
    });
  }
  return list;
}

/* ====== Render ====== */
function render(){
  const target = el("cardsWrap");
  const data = filteredRows();
  el("countRows").textContent = data.length;
  if(data.length === 0){
    target.innerHTML = `<div class="col-12"><div class="alert alert-warning mb-0"><i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.</div></div>`;
    return;
  }
  target.innerHTML = data.map((r,idx) => cardHtml(r, RAW_ROWS.indexOf(r))).join("");
}

function cardHtml(r, idx){
  const badges = RISKS.map(x => {
    const v = (r[x.key]||"").toString().toUpperCase();
    const cls = v === "SI" ? "badge-yes" : (v === "NO" ? "badge-no" : "badge-na");
    const label = v || "N/A";
    return `<span class="badge ${cls}" title="${escapeHtml(x.hoja)}">${escapeHtml(x.label)}: ${label}</span>`;
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
              <h5 class="title mb-2">${escapeHtml(r.B || "-")}</h5>
            </div>
            <div class="text-end">
              <span class="chip"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
            </div>
          </div>
          <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
          <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>

          <div class="row g-2 small">
            <div class="col-6"><i class="bi bi-clock"></i> <strong>Horario:</strong> ${escapeHtml(r.E || "-")}</div>
            <div class="col-6"><i class="bi bi-plus-circle"></i> <strong>HE/Día:</strong> ${escapeHtml(r.F || "0")}</div>
            <div class="col-6"><i class="bi bi-plus-circle-dotted"></i> <strong>HE/Semana:</strong> ${escapeHtml(r.G || "0")}</div>
          </div>

          <hr>
          <div class="small"><strong>Factores (J–P):</strong></div>
          <div class="d-flex flex-wrap gap-1 my-1">${badges}</div>

          <div class="alert ${resultClass} mt-2 mb-0" role="alert">
            <i class="bi bi-clipboard-check"></i> <strong>Resultado:</strong> ${escapeHtml(r.Q || "-")}
          </div>
        </div>
      </div>
    </div>
  `;
}

/* ====== Panel Detalle ====== */
/* ====== Panel Detalle (enriquecido) ====== */
function openDetail(r){
  const body = el("detailBody");
  const adv = needsAdvancedEval(r);
  const mmcWarn = (r.L === "SI" || r.M === "SI");

  // Factores con su “cumple/no” y hoja relacionada
  const factorsHtml = RISKS.map(x => {
    const v = (r[x.key]||"").toString().toUpperCase() || "N/A";
    const cls = v === "SI" ? "text-bg-danger" : (v === "NO" ? "text-bg-success" : "text-bg-secondary");
    const note = v === "SI" ? "Revisar criterios – podría requerir avanzada." :
                v === "NO" ? "Sin hallazgo para este factor." :
                             "Dato no disponible.";
    return `
      <div class="d-flex align-items-start justify-content-between py-1 border-bottom">
        <div class="me-2">
          <div class="fw-semibold">${escapeHtml(x.label)}</div>
          <div class="small text-muted">Hoja: ${escapeHtml(x.hoja)} · ${escapeHtml(note)}</div>
        </div>
        <span class="badge ${cls}">${v}</span>
      </div>`;
  }).join("");

  // Q y banderas
  const advAlert = adv
    ? `<div class="alert alert-danger"><i class="bi bi-x-octagon"></i> <strong>Corresponde identificación avanzada.</strong></div>`
    : `<div class="alert alert-success"><i class="bi bi-check2"></i> Sin hallazgos que requieran avanzada.</div>`;

  const mmcAlert = mmcWarn
    ? `<div class="alert alert-warning"><i class="bi bi-exclamation-triangle"></i> Atención MMC:
        ${ r.L==="SI" ? `<span class="badge rounded-pill text-bg-warning ms-1">Levantamiento/Descenso</span>` : "" }
        ${ r.M==="SI" ? `<span class="badge rounded-pill text-bg-warning ms-1">Empuje/Arrastre</span>` : "" }
      </div>` : "";

  // Resumen de cuestionarios del Área (toma de cache si existe)
  const qx = getAreaSurveySummary(r.B);
  const qxHtml = qx
    ? `<div class="mt-3">
        <div class="small text-muted mb-1">Cuestionarios del área (últimos 12 meses)</div>
        <div class="d-flex flex-wrap gap-2 mb-2">
          <span class="badge text-bg-success">Vigentes: ${qx.vigentes}</span>
          <span class="badge text-bg-danger">Vencidos: ${qx.vencidos}</span>
          <span class="badge text-bg-secondary">Total: ${qx.total}</span>
        </div>
        ${qx.sample.length ? `<ul class="list-group list-group-flush">
          ${qx.sample.map(p => `
            <li class="list-group-item px-0 d-flex justify-content-between">
              <div>${escapeHtml(p.nombre)} <span class="small text-muted">· ${escapeHtml(p.fecha)}</span></div>
              <span class="badge rounded-pill ${p.vigente ? 'text-bg-success' : 'text-bg-danger'}">${p.vigente?'Vigente':'Vencido'}</span>
            </li>`).join("")}
        </ul>` : `<div class="text-muted small">No hay respuestas registradas para esta área.</div>`}
        <div class="mt-2">
          <a class="btn btn-sm btn-outline-primary" href="cuestionarios_respondidos.html?area=${encodeURIComponent(r.B||'')}">
            <i class="bi bi-list-ul"></i> Ver cuestionarios del área
          </a>
        </div>
      </div>`
    : "";

  body.innerHTML = `
    <div>
      <div class="small text-muted">Área</div>
      <h5 class="mb-1">${escapeHtml(r.B || "-")}</h5>
      <div class="mb-2"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
      <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>
      <div class="d-flex flex-wrap gap-2 mb-2">
        <span class="badge text-bg-light"><i class="bi bi-clock"></i> ${escapeHtml(r.E || "-")}</span>
        <span class="badge text-bg-light"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
        <span class="badge ${ (r.Q||"").toLowerCase().includes('avanzada') ? 'text-bg-danger' : 'text-bg-primary' }">
          ${escapeHtml(r.Q||"-")}
        </span>
      </div>
      ${advAlert}
      ${mmcAlert}
      <hr>
      <h6 class="mb-2">Factores evaluados</h6>
      ${factorsHtml}
      ${qxHtml}
    </div>
  `;

  el("detailTitle").textContent = `Detalle · ${r.B || "-"}`;
  const oc = bootstrap.Offcanvas.getOrCreateInstance('#detailPanel');
  oc.show();
}

// Lee cache y resume cuestionarios del área (vigentes/vencidos + muestra corta)
function getAreaSurveySummary(area){
  try{
    const cache = JSON.parse(localStorage.getItem("RESPUESTAS_CACHE_V1") || "null");
    if(!cache || !Array.isArray(cache.rows)) return null;
    const rows = cache.rows.filter(r => String(r.Area||"").trim() === String(area||"").trim());
    if(!rows.length) return {total:0, vigentes:0, vencidos:0, sample:[]};

    // quedarnos con última respuesta por persona
    const last = new Map();
    rows.forEach(r => {
      const k = (String(r.Puesto_de_Trabajo||"") + "||" + String(r.Nombre_Trabajador||"")).toLowerCase();
      const prev = last.get(k);
      if(!prev){ last.set(k,r); return; }
      const d1 = parseFlex(prev.Fecha), d2 = parseFlex(r.Fecha);
      if(d2.getTime() > d1.getTime()) last.set(k,r);
    });

    let vig=0, ven=0;
    const items = [];
    last.forEach(r => {
      const d = parseFlex(r.Fecha);
      const vigente = (Date.now() - d.getTime()) / 86400000 <= 365;
      if(vigente) vig++; else ven++;
      items.push({ nombre: String(r.Nombre_Trabajador||""), fecha: String(r.Fecha||""), vigente });
    });

    // muestra corta (hasta 5)
    items.sort((a,b)=> a.nombre.localeCompare(b.nombre));
    return { total: items.length, vigentes: vig, vencidos: ven, sample: items.slice(0,5) };
  }catch(_){ return null; }

  function parseFlex(s){
    const t = String(s||""); 
    if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)){ const [dd,mm,yy]=t.split("/").map(n=>+n); return new Date(yy,mm-1,dd); }
    if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(t)){ const [yy,mm,dd]=t.split("-").map(n=>+n); return new Date(yy,mm-1,dd); }
    const d = new Date(t); return isNaN(d)? new Date(0):d;
  }
}



function needsAdvancedEval(r){
  // Avanzada si Q lo indica o si algún factor en SI
  const q = (r.Q||"").toLowerCase();
  const qSaysAdv = q.includes("aplicar identificación avanzada") || q.includes("aplicar identificacion avanzada");
  const anyYes = ['J','K','L','M','N','O','P'].some(k => (r[k]||"").toUpperCase() === "SI");
  return qSaysAdv || anyYes;
}

/* ====== Utils ====== */
function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
