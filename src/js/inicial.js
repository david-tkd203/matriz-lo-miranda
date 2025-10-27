/* inicial.js (1 por fila + sticky bottom header + paginación + omitir H+M=0)
   - Lee HOJA INICIAL (fila 3, columnas B..Q)
   - Filtros en cascada, por Factor (SI/NO/Todos)
   - Cards clicables → Offcanvas detalle (con colores)
   - Paginación con barra sticky inferior
   - Omite filas con (H+I)==0 o Área/Puesto "0"
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

// Paginación
let PAGE = 1;
let PER_PAGE = 10;

const el = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", () => {
  attemptFetchDefault();
  wireUI();
});

function wireUI(){
  // Carga manual de Excel
  el("fileInput").addEventListener("change", handleFile);

  // Filtros cascada
  el("filterArea").addEventListener("change", () => {
    FILTERS.area = el("filterArea").value || "";
    PAGE = 1;
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
    PAGE = 1;
    populateTarea();
    FILTERS.tarea = "";
    el("filterTarea").value = "";
    render();
  });

  el("filterTarea").addEventListener("change", () => {
    FILTERS.tarea = el("filterTarea").value || "";
    PAGE = 1;
    render();
  });

  // Filtro por factor
  populateFactor();
  el("filterFactor").addEventListener("change", () => {
    FILTERS.factorKey = el("filterFactor").value || "";
    PAGE = 1;
    render();
  });
  el("filterFactorState").addEventListener("change", () => {
    FILTERS.factorState = el("filterFactorState").value || "";
    PAGE = 1;
    render();
  });

  // Reset & Reload
  el("btnReset").addEventListener("click", (e) => {
    e.preventDefault();
    FILTERS = { area: "", puesto: "", tarea: "", factorKey:"", factorState:"" };
    PAGE = 1;
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

  // Delegación: abrir detalle
  el("cardsWrap").addEventListener("click", (ev) => {
    const card = ev.target.closest("[data-idx]");
    if(!card) return;
    const idx = Number(card.dataset.idx);
    if(Number.isFinite(idx) && RAW_ROWS[idx]){
      openDetail(RAW_ROWS[idx]);
    }
  });

  // Paginación
  el("perPage").addEventListener("change", () => {
    PER_PAGE = Math.max(1, parseInt(el("perPage").value, 10) || 10);
    PAGE = 1;
    render();
  });
  el("btnPrev").addEventListener("click", () => {
    PAGE = Math.max(1, PAGE - 1);
    render();
  });
  el("btnNext").addEventListener("click", () => {
    const total = filteredRows().length;
    const maxPage = Math.max(1, Math.ceil(total / PER_PAGE));
    PAGE = Math.min(maxPage, PAGE + 1);
    render();
  });
  el("btnTop").addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
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

function toNum(v){
  if(v == null) return 0;
  const n = Number(String(v).replace(/[, ]/g,"").trim());
  return isNaN(n) ? 0 : n;
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
    const getCell = (c) => {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const value = cell ? (cell.w ?? cell.v) : "";
      return value == null ? "" : String(value).trim();
    };
    const colMap = { B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16 };
    for(const [k, idx] of Object.entries(colMap)){ vals[k] = getCell(idx); }

    // Omitir "falsos positivos":
    // - Área o Puesto con "0" literal
    // - Dotación total (H + M) == 0
    const areaStr = (vals.B || "").trim();
    const puestoStr = (vals.C || "").trim();
    const hNum = toNum(vals.H);
    const iNum = toNum(vals.I);
    if(areaStr === "0" || puestoStr === "0") continue;
    if((hNum + iNum) <= 0) continue;

    // Normalizar SI/NO en factores
    ['J','K','L','M','N','O','P'].forEach(k => {
      if(vals[k]){
        const up = vals[k].toString().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
        if(["SI","YES","Y","S"].includes(up)) vals[k] = "SI";
        else if(["NO","N"].includes(up)) vals[k] = "NO";
      }
    });

    // Calcular Q
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
      return v === "SI" || v === "NO"; // si elige factor pero estado=(Todos), excluye N/A
    });
  }
  return list;
}

/* ====== Render con paginación ====== */
function render(){
  const all = filteredRows();
  const total = all.length;
  const maxPage = Math.max(1, Math.ceil(total / PER_PAGE));
  if(PAGE > maxPage) PAGE = maxPage;

  const start = (PAGE - 1) * PER_PAGE;
  const data = all.slice(start, start + PER_PAGE);

  el("countRows").textContent = data.length;
  el("countRowsTotal").textContent = total;
  el("pageCur").textContent = PAGE;
  el("pageMax").textContent = maxPage;

  const target = el("cardsWrap");
  if(data.length === 0){
    target.innerHTML = `<div class="col-12"><div class="alert alert-warning mb-0"><i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.</div></div>`;
    return;
  }
  target.innerHTML = data.map(r => cardHtml(r, RAW_ROWS.indexOf(r))).join("");
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
    <div class="col-12">
      <div class="card card-ficha h-100 shadow-sm" role="button" tabindex="0" data-idx="${idx}">
        <div class="card-body">
          <div class="d-flex align-items-start justify-content-between">
            <div>
              <div class="small text-muted mb-1">Área</div>
              <h5 class="title">${escapeHtml(r.B || "-")}</h5>
              <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
              <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>
            </div>
            <div class="text-end">
              <span class="chip"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
              <div class="small mt-1"><i class="bi bi-clock"></i> ${escapeHtml(r.E || "-")}</div>
              <div class="small mt-1"><i class="bi bi-plus-circle"></i> HE/Día: ${escapeHtml(r.F || "0")}</div>
              <div class="small"><i class="bi bi-plus-circle-dotted"></i> HE/Semana: ${escapeHtml(r.G || "0")}</div>
            </div>
          </div>

          <hr>
          <div class="small"><strong>Factores (J–P):</strong></div>
          <div class="badges-wrap my-1">${badges}</div>

          <div class="alert ${resultClass} mt-2 mb-0" role="alert">
            <i class="bi bi-clipboard-check"></i> <strong>Resultado:</strong> ${escapeHtml(r.Q || "-")}
          </div>
        </div>
      </div>
    </div>
  `;
}

/* ====== Panel Detalle ====== */
function openDetail(r){
  const body = el("detailBody");
  const adv = needsAdvancedEval(r);
  const mmcWarn = (r.L === "SI" || r.M === "SI");

  const factorsHtml = RISKS.map(x => {
    const v = (r[x.key]||"").toString().toUpperCase() || "N/A";
    const cls = v === "SI" ? "badge-yes" : (v === "NO" ? "badge-no" : "badge-na");
    return `<div class="factor-row">
      <div><strong>${escapeHtml(x.label)}</strong> <span class="text-muted">(${escapeHtml(x.hoja)})</span></div>
      <div><span class="badge ${cls}">${v}</span></div>
    </div>`;
  }).join("");

  const advAlert = adv
    ? `<div class="alert alert-danger"><i class="bi bi-x-octagon"></i> <strong>Corresponde identificación avanzada.</strong></div>`
    : `<div class="alert alert-success"><i class="bi bi-check2-circle"></i> Sin hallazgos que requieran avanzada.</div>`;

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
        <div class="d-flex flex-wrap gap-2">
          <span class="pill"><i class="bi bi-clock"></i> ${escapeHtml(r.E || "-")}</span>
          <span class="pill"><i class="bi bi-people"></i> H ${escapeHtml(r.H||"0")} · M ${escapeHtml(r.I||"0")}</span>
        </div>
      </div>
    </div>

    <hr>

    ${advAlert}
    ${mmcAlert}

    <h6 class="mt-3">Factores evaluados</h6>
    ${factorsHtml}

    <hr>
    <div class="small text-muted">Resultado hoja inicial</div>
    <div>${escapeHtml(r.Q||"-")}</div>
  `;

  el("detailTitle").textContent = `Detalle · ${r.B || "-"}`;
  const oc = bootstrap.Offcanvas.getOrCreateInstance('#detailPanel');
  oc.show();
}

function needsAdvancedEval(r){
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
