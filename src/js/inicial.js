
/* inicial.js (reparado)
   - Busca la hoja por nombre que contenga "inicial" (case-insensitive). Si no existe, usa la primera.
   - Lee desde fila 3 (índice 2) y columnas B..Q mapeadas explícitamente.
   - Filtros en cascada: B=Área, C=Puesto, D=Tareas.
   - Render de fichas: B–I + factores J–P (mostrando el NOMBRE del factor) + Q (calculado).
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

let RAW_ROWS = []; // rows parsed
let FILTERS = { area: "", puesto: "", tarea: "" };

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
  el("btnReset").addEventListener("click", () => {
    FILTERS = { area: "", puesto: "", tarea: "" };
    el("filterArea").value = "";
    populatePuesto(true);
    populateTarea(true);
    render();
  });
  el("btnReload").addEventListener("click", attemptFetchDefault);
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
    console.warn("La hoja inicial no tiene datos visibles.");
    RAW_ROWS = [];
    render();
    return;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];

  for(let r = 2; r <= range.e.r; r++){ // fila 3 (índice 2)
    const vals = {};
    function getCell(c){
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if(!cell) return "";
      // preferir .w (texto formateado) si existe; si no, usar .v
      const value = (cell.w != null ? cell.w : cell.v);
      if(value == null) return "";
      return String(value).trim();
    }
    const colMap = { B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16 };
    for(const [k, idx] of Object.entries(colMap)){
      vals[k] = getCell(idx);
    }

    // Ignorar si B, C y D están todos vacíos
    if(!(vals.B || vals.C || vals.D)) continue;

    // Normalización de Sí/No en factores J..P
    ['J','K','L','M','N','O','P'].forEach(k => {
      if(vals[k]){
        const up = vals[k].toString().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
        if(["SI","YES","Y","S"].includes(up)) vals[k] = "SI";
        else if(["NO","N"].includes(up)) vals[k] = "NO";
        else vals[k] = vals[k]; // valor no reconocido -> badge N/A
      }
    });

    // Calcular Q según condición
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

function uniqueSorted(arr){
  return [...new Set(arr.filter(v => v && String(v).trim() !== ""))].sort((a,b)=> String(a).localeCompare(String(b)));
}

function populateArea(){
  const opts = uniqueSorted(RAW_ROWS.map(r => r.B));
  const sel = el("filterArea");
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = false;
}

function populatePuesto(reset=false){
  const sel = el("filterPuesto");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B === FILTERS.area);
  const opts = uniqueSorted(list.map(r => r.C));
  sel.innerHTML = `<option value="">(Todos)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length === 0;
}

function populateTarea(reset=false){
  const sel = el("filterTarea");
  let list = RAW_ROWS;
  if(FILTERS.area) list = list.filter(r => r.B === FILTERS.area);
  if(FILTERS.puesto) list = list.filter(r => r.C === FILTERS.puesto);
  const opts = uniqueSorted(list.map(r => r.D));
  sel.innerHTML = `<option value="">(Todas)</option>` + opts.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sel.disabled = opts.length === 0;
}

function filteredRows(){
  return RAW_ROWS.filter(r => {
    if(FILTERS.area && r.B !== FILTERS.area) return false;
    if(FILTERS.puesto && r.C !== FILTERS.puesto) return false;
    if(FILTERS.tarea && r.D !== FILTERS.tarea) return false;
    return true;
  });
}

function render(){
  const target = el("cardsWrap");
  const data = filteredRows();
  el("countRows").textContent = data.length;
  if(data.length === 0){
    target.innerHTML = `<div class="col-12"><div class="alert alert-warning mb-0"><i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.</div></div>`;
    return;
  }
  target.innerHTML = data.map(r => cardHtml(r)).join("");
}

function cardHtml(r){
  const badges = RISKS.map(x => {
    const v = (r[x.key]||"").toString().toUpperCase();
    const cls = v === "SI" ? "badge-yes" : (v === "NO" ? "badge-no" : "badge-na");
    const label = v || "N/A";
    // Mostrar el NOMBRE del factor (no la letra)
    return `<span class="badge ${cls}" title="${escapeHtml(x.hoja)}">${escapeHtml(x.label)}: ${label}</span>`;
  }).join(" ");

  return `
    <div class="col-12 col-md-6 col-lg-4">
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

          <div class="row g-2 small">
            <div class="col-6"><i class="bi bi-clock"></i> <strong>Horario:</strong> ${escapeHtml(r.E || "-")}</div>
            <div class="col-6"><i class="bi bi-plus-circle"></i> <strong>HE/Día:</strong> ${escapeHtml(r.F || "0")}</div>
            <div class="col-6"><i class="bi bi-plus-circle-dotted"></i> <strong>HE/Semana:</strong> ${escapeHtml(r.G || "0")}</div>
          </div>

          <hr>
          <div class="small"><strong>Factores (J–P):</strong></div>
          <div class="d-flex flex-wrap gap-1 my-1">${badges}</div>

          <div class="alert ${ (r.Q||"").startsWith("Ausencia total") ? "alert-success" : "alert-primary"} mt-2 mb-0" role="alert">
            <i class="bi bi-clipboard-check"></i> <strong>Resultado:</strong> ${escapeHtml(r.Q || "-")}
          </div>
        </div>
      </div>
    </div>
  `;
}

function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
