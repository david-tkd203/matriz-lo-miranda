/* inicial.js (v. “Mov. Rep. con modal fullscreen + slide por página + 1 tarjeta/fila”)
   - Lee hoja INICIAL (B..Q) y hoja "Mov. Rep."
   - Une por (Área|Puesto|Tareas)
   - Tarjeta muestra estado derivado de columnas P y W de Mov. Rep.
   - Modal fullscreen con P/W y preguntas/respuestas (todas las columnas con texto)
   - Paginación con animación “slide”
   - Omite filas con Área o Puesto = "0"
*/

const COLS = {
  B: "Área", C: "Puesto de trabajo", D: "Tareas del puesto de trabajo",
  E: "Horario de funcionamiento", F: "Horas extras POR DIA", G: "Horas extras POR SEMANA",
  H: "N° Trabajadores Expuestos HOMBRE", I: "N° Trabajadores Expuestos MUJER",
  J: "Trabajo repetitivo de miembros superiores.", K: "Postura de trabajo estática",
  L: "MMC Levantamiento/Descenso", M: "MMC Empuje/Arrastre", N: "Manejo manual de pacientes / personas",
  O: "Vibración de cuerpo completo", P: "Vibración segmento mano – brazo", Q: "Resultado identificación inicial",
};

// Estado de la UI / datos
let RAW_ROWS = [];        // Hoja INICIAL
let MOVREP_INDEX = new Map(); // key -> { rowObj, headers, valuesByCol, P, W }
let FILTERS = { area:"", puesto:"", tarea:"" };
let STATE = { page:1, perPage:10, pageMax:1, lastPage:1 };

const el = (id)=>document.getElementById(id);

/* ===== Helpers ===== */
function isZeroish(v){
  if(v===0) return true;
  const s = String(v??"").trim();
  return /^0(\.0+)?$/.test(s);
}
function escapeHtml(str){ return String(str).replace(/[&<>"']/g, s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[s]); }
function norm(s){ return String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim(); }
function kAreaPuestoTarea(a,p,t){ return `${norm(a)}|${norm(p)}|${norm(t)}`; }
function uniqueSorted(arr){ return [...new Set(arr.filter(v => v && String(v).trim() !== ""))].sort((a,b)=> String(a).localeCompare(String(b))); }

/* ===== Carga ===== */
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
    STATE.page = 1;
    render();
  });
  el("filterPuesto").addEventListener("change", () => {
    FILTERS.puesto = el("filterPuesto").value || "";
    populateTarea(); FILTERS.tarea = ""; el("filterTarea").value = "";
    STATE.page = 1;
    render();
  });
  el("filterTarea").addEventListener("change", () => {
    FILTERS.tarea = el("filterTarea").value || "";
    STATE.page = 1;
    render();
  });

  el("btnReset").addEventListener("click", (e)=>{ e.preventDefault();
    FILTERS = { area:"", puesto:"", tarea:"" };
    el("filterArea").value=""; el("filterPuesto").value=""; el("filterTarea").value="";
    STATE.page=1; render();
  });
  el("btnReload").addEventListener("click", attemptFetchDefault);

  // paginación / slide
  el("perPage").addEventListener("change", () => { STATE.perPage = parseInt(el("perPage").value,10)||10; STATE.page=1; render(); });
  el("btnPrev").addEventListener("click", ()=> { if(STATE.page>1){ STATE.lastPage=STATE.page; STATE.page--; render(true); }});
  el("btnNext").addEventListener("click", ()=> { if(STATE.page<STATE.pageMax){ STATE.lastPage=STATE.page; STATE.page++; render(true); }});
  el("btnTop").addEventListener("click", ()=> window.scrollTo({top:0, behavior:"smooth"}));

  // abrir modal desde botón de tarjeta (delegación)
  el("cardsWrap").addEventListener("click", (ev)=>{
    const btn = ev.target.closest("[data-open-detail]");
    if(!btn) return;
    const idx = parseInt(btn.dataset.openDetail,10);
    const row = CURRENT_ROWS[idx];
    if(row) openDetail(row);
  });
}

async function attemptFetchDefault(){
  if(!window.DEFAULT_XLSX_PATH) return;
  try{
    const res = await fetch(window.DEFAULT_XLSX_PATH);
    if(!res.ok) throw new Error("Fetch failed");
    const buf = await res.arrayBuffer();
    processWorkbook(buf);
  }catch(e){
    console.warn("No se pudo cargar automáticamente el Excel. Seleccione el archivo manualmente.", e);
  }
}

function handleFile(evt){
  const file = evt.target.files?.[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = (e) => processWorkbook(e.target.result);
  reader.readAsArrayBuffer(file);
}

function pickInicialSheet(wb){
  const target = (wb.SheetNames || []).find(n => /inicial|inicio/i.test(String(n||"")));
  return target || wb.SheetNames[0];
}
function pickMovRepSheet(wb){
  // nombres frecuentes: "Mov. Rep.", "Movimiento repetitivo", "Mov Rep", etc.
  const target = (wb.SheetNames || []).find(n => /mov|\brep/i.test(String(n||"")));
  return target || null;
}

function processWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type:"array" });

  // ===== Hoja INICIAL (B..Q)
  const initialSheetName = pickInicialSheet(wb);
  const ws = wb.Sheets[initialSheetName];
  RAW_ROWS = [];
  if(ws && ws['!ref']){
    const range = XLSX.utils.decode_range(ws['!ref']);
    for(let r=2; r<=range.e.r; r++){ // desde fila 3
      const vals = {};
      const colMap = { B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16 };
      for(const [k,c] of Object.entries(colMap)){
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        vals[k] = cell ? (cell.w ?? cell.v) : "";
        vals[k] = vals[k]==null ? "" : String(vals[k]).trim();
      }
      if(!(vals.B || vals.C || vals.D)) continue;
      if(isZeroish(vals.B) || isZeroish(vals.C)) continue; // omitir falsos positivos cero
      RAW_ROWS.push(vals);
    }
  }

  // ===== Hoja Mov. Rep. (tomamos columnas P y W + Q&A)
  MOVREP_INDEX = buildMovRepIndex(wb);

  populateArea();
  populatePuesto(true);
  populateTarea(true);
  STATE.page = 1;
  render();
}

function buildMovRepIndex(wb){
  const name = pickMovRepSheet(wb);
  const out = new Map();
  if(!name) return out;
  const ws = wb.Sheets[name];
  if(!ws || !ws['!ref']) return out;

  // 2D para detectar encabezados
  const rows2D = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
  if(!rows2D.length) return out;

  // Elegimos la fila de encabezados: primera que contenga "Área" y "Puesto" y "Tareas", si no la 1ª
  let hRow = 0;
  for(let i=0;i<Math.min(5,rows2D.length);i++){
    const line = rows2D[i].map(v => String(v).toLowerCase());
    if(line.some(v=>v.includes("área")||v.includes("area")) &&
       line.some(v=>v.includes("puesto")) &&
       line.some(v=>v.includes("tarea"))) { hRow = i; break; }
  }
  const headers = rows2D[hRow].map(h => String(h||"").trim());
  const startRow = hRow + 1;

  // utilidad: obtener por índice de columna o por letra fija
  const COL = { A:0,B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18,T:19,U:20,V:21,W:22,X:23,Y:24,Z:25 };

  for(let r = startRow; r < rows2D.length; r++){
    const row = rows2D[r] || [];
    const area   = String(row[COL.B] ?? "").trim();
    const puesto = String(row[COL.C] ?? "").trim();
    const tarea  = String(row[COL.D] ?? "").trim();
    if(!(area || puesto || tarea)) continue;
    if(isZeroish(area) || isZeroish(puesto)) continue;

    const P = String(row[COL.P] ?? "").trim(); // Col P
    const W = String(row[COL.W] ?? "").trim(); // Col W

    // Construir listado de Q&A (omitimos B,C,D,P,W y celdas vacías)
    const qa = [];
    for(let c=0;c<row.length;c++){
      if([COL.B, COL.C, COL.D, COL.P, COL.W].includes(c)) continue;
      const val = row[c];
      const label = headers[c] || `Col ${XLSX.utils.encode_col(c)}`;
      if(val !== "" && label) qa.push({ label: String(label).trim(), value: String(val).trim() });
    }

    out.set(kAreaPuestoTarea(area, puesto, tarea), {
      area, puesto, tarea, P, W, qa, headers
    });
  }
  return out;
}

/* ===== Poblar filtros ===== */
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

/* ===== Filtro compuesto ===== */
function filteredRows(){
  return RAW_ROWS.filter(r => {
    if(FILTERS.area && r.B !== FILTERS.area) return false;
    if(FILTERS.puesto && r.C !== FILTERS.puesto) return false;
    if(FILTERS.tarea && r.D !== FILTERS.tarea) return false;
    return true;
  });
}

/* ===== Estado por P/W ===== */
function pwStatus(p, w){
  const s = (w || p || "").toString().toLowerCase();
  // heurística robusta
  if(/no acept|alto|riesg|crit|aplicar.*avanz|evaluaci[oó]n avanzada|requiere/.test(s))
    return {cls:"bad",  icon:"bi-x-octagon", label:"Requiere avanzada"};
  if(/aceptable|sin riesgo|no aplica|apto|acept/.test(s))
    return {cls:"ok",   icon:"bi-check-circle", label:"Aceptable"};
  return {cls:"warn", icon:"bi-exclamation-triangle", label:"Revisar"};
}

/* ===== Render con slide ===== */
let CURRENT_ROWS = []; // página actual (para abrir modal por índice)
function render(withSlide=false){
  const data = filteredRows();
  el("countRowsTotal").textContent = data.length;

  // paginación
  STATE.pageMax = Math.max(1, Math.ceil(data.length / STATE.perPage));
  if(STATE.page > STATE.pageMax) STATE.page = STATE.pageMax;

  const start = (STATE.page - 1) * STATE.perPage;
  CURRENT_ROWS = data.slice(start, start + STATE.perPage);

  const wrap = el("cardsWrap");

  const direction = STATE.page > STATE.lastPage ? "right" : "left";
  if(withSlide){
    wrap.classList.remove("slide-enter-left","slide-enter-right");
    void wrap.offsetWidth; // reflow
    wrap.classList.add(direction==="right" ? "slide-enter-right" : "slide-enter-left");
  }

  if(CURRENT_ROWS.length === 0){
    wrap.innerHTML = `<div class="col-12"><div class="alert alert-warning mb-0">
      <i class="bi bi-exclamation-triangle"></i> No hay resultados con los filtros aplicados.
    </div></div>`;
  }else{
    wrap.innerHTML = CURRENT_ROWS.map((r, idx) => cardHtml(r, idx)).join("");
  }

  el("countRows").textContent = CURRENT_ROWS.length;
  el("pageCur").textContent = STATE.page;
  el("pageMax").textContent = STATE.pageMax;
}

function cardHtml(r, idx){
  const key = kAreaPuestoTarea(r.B, r.C, r.D);
  const mr = MOVREP_INDEX.get(key);
  const P = mr?.P || "";
  const W = mr?.W || "";
  const st = pwStatus(P, W);

  return `
    <div class="col-12">
      <div class="card card-ficha h-100 shadow-sm">
        <div class="card-body">
          <div class="d-flex align-items-start justify-content-between mb-2">
            <div>
              <div class="small text-muted">Área</div>
              <h5 class="title mb-1">${escapeHtml(r.B || "-")}</h5>
              <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
              <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>
            </div>
            <span class="status ${st.cls}" title="Mov. Repetitivo · P/W">
              <i class="bi ${st.icon}"></i> ${st.label}
            </span>
          </div>

          <div class="row g-2 small">
            <div class="col-6"><i class="bi bi-clock"></i> <strong>Horario:</strong> ${escapeHtml(r.E || "-")}</div>
            <div class="col-6"><i class="bi bi-people"></i> <strong>H:</strong> ${escapeHtml(r.H||"0")} · <strong>M:</strong> ${escapeHtml(r.I||"0")}</div>
          </div>

          <hr class="my-2">
          <div class="d-flex gap-2 flex-wrap">
            <span class="pill"><i class="bi bi-upc-scan"></i> P: ${escapeHtml(P || "N/A")}</span>
            <span class="pill"><i class="bi bi-clipboard2-check"></i> W: ${escapeHtml(W || "N/A")}</span>
          </div>

          <div class="text-end mt-3">
            <button class="btn btn-sm btn-primary" data-open-detail="${idx}">
              <i class="bi bi-eye"></i> Ver detalles
            </button>
          </div>
        </div>
      </div>
    </div>
  `;
}

/* ===== Modal Detalle (QA + P/W) ===== */
function openDetail(r){
  const key = kAreaPuestoTarea(r.B, r.C, r.D);
  const mr = MOVREP_INDEX.get(key);

  const P = mr?.P || "";
  const W = mr?.W || "";
  const st = pwStatus(P, W);

  const qaHtml = (mr?.qa || [])
    .map(q => `
      <div class="qrow">
        <div class="qlabel">${escapeHtml(q.label)}</div>
        <div class="qval">${escapeHtml(q.value)}</div>
      </div>
    `).join("") || `<div class="alert alert-secondary">No hay preguntas/respuestas disponibles para esta fila.</div>`;

  const body = `
    <div class="mb-2">
      <div class="small text-muted">Área</div>
      <h4 class="mb-1">${escapeHtml(r.B || "-")}</h4>
      <div class="mb-1"><i class="bi bi-person-badge"></i> <strong>Puesto:</strong> ${escapeHtml(r.C || "-")}</div>
      <div class="mb-2"><i class="bi bi-list-check"></i> <strong>Tareas:</strong> ${escapeHtml(r.D || "-")}</div>
    </div>

    <div class="d-flex flex-wrap gap-2 mb-3">
      <span class="status ${st.cls}"><i class="bi ${st.icon}"></i> ${st.label}</span>
      <span class="pill"><i class="bi bi-upc-scan"></i> Columna P: <strong>${escapeHtml(P || "N/A")}</strong></span>
      <span class="pill"><i class="bi bi-clipboard2-check"></i> Columna W: <strong>${escapeHtml(W || "N/A")}</strong></span>
    </div>

    <h6 class="mb-2">Preguntas y respuestas</h6>
    <div class="detail-grid">
      ${qaHtml}
    </div>
  `;

  el("detailTitle").textContent = "Detalle · Movimiento Repetitivo";
  el("detailBody").innerHTML = body;

  const modal = bootstrap.Modal.getOrCreateInstance('#detailModal');
  modal.show();
}
