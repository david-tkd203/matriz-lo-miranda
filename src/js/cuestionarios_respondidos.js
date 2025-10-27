// cuestionarios_respondidos.js (pro)
// - Filtros: Área, Sexo, Nombre, Vigencia
// - Stats: total / vigentes / vencidos
// - Ordenamiento por columnas
// - Paginación con barra inferior sticky
// - Exportar vista (CSV)
// - Carga robusta del XLSX + fallback a archivo local
// - Hidrata desde cache local (RESPUESTAS_CACHE_V1) si existe

document.addEventListener("DOMContentLoaded", () => {
  ensureUiHooks();
  tryLoadFromCacheFirst();  // pinta algo si hay cache
  loadExcelAuto();          // luego intenta el archivo real

  // Filtros
  byId("filterArea").addEventListener("change", resetPageAndRender);
  byId("filterSexo").addEventListener("change", resetPageAndRender);
  byId("filterVigencia").addEventListener("change", resetPageAndRender);
  byId("filterName").addEventListener("input", resetPageAndRender);

  // Acciones
  byId("btnReset").addEventListener("click", resetFilters);
  byId("btnExport").addEventListener("click", exportCsv);
  byId("fileXlsx").addEventListener("change", handleLocalXlsx);

  // Ordenamiento por th
  document.querySelectorAll("thead th[data-sort]").forEach(th => {
    th.addEventListener("click", () => {
      const k = th.getAttribute("data-sort");
      if(SORT.key === k){
        SORT.dir = SORT.dir === "asc" ? "desc" : "asc";
      }else{
        SORT.key = k; SORT.dir = "asc";
      }
      render();
    });
  });

  // Paginación
  byId("perPage").addEventListener("change", () => {
    PER_PAGE = Math.max(1, parseInt(byId("perPage").value, 10) || 25);
    PAGE = 1; render();
  });
  byId("btnPrev").addEventListener("click", () => { PAGE = Math.max(1, PAGE-1); render(); });
  byId("btnNext").addEventListener("click", () => {
    const total = filteredRows().length;
    const maxPage = Math.max(1, Math.ceil(total / PER_PAGE));
    PAGE = Math.min(maxPage, PAGE+1); render();
  });
  byId("btnTop").addEventListener("click", () => window.scrollTo({top:0, behavior:"smooth"}));
});

let ROWS = [];
let AREAS = [];
let SEXOS = []; // ['H','M']
let SORT = { key: "Fecha", dir: "desc" };
let PAGE = 1;
let PER_PAGE = 25;

// ---------- Utils ----------
function byId(id){ return document.getElementById(id); }
function ensureUiHooks(){
  if(!byId("loadError")){
    const container = document.querySelector(".container");
    if(container){
      const d = document.createElement("div");
      d.id = "loadError";
      d.className = "d-none";
      container.prepend(d);
    }
  }
}
function showError(msg){
  const el = byId("loadError");
  el.classList.remove("d-none");
  el.innerHTML = `
    <div class="alert alert-warning d-flex align-items-start gap-2 mb-3">
      <i class="bi bi-exclamation-triangle mt-1"></i>
      <div>
        <div class="fw-semibold">No se pudo leer el Excel de cuestionarios.</div>
        <div class="small">${escapeHtml(msg)}</div>
      </div>
    </div>`;
}
function escapeHtml(str){
  return String(str ?? "").replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
function toLowerNoAccents(s){
  return String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}
function normalizeHeader(h){
  if(h == null) return "";
  let s = String(h).replace(/\uFEFF/g,"").trim();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  s = s.replace(/\s+/g,'_').replace(/[^\w]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');
  return s.toLowerCase();
}
function rowToObj(headersNorm, rowArr){
  const o = {};
  headersNorm.forEach((keyNorm, idx) => { o[keyNorm] = rowArr[idx] == null ? "" : rowArr[idx]; });
  return o;
}
function mapToLogical(objNorm){
  const pick = (cands) => { for(const k of cands){ if(k in objNorm) return objNorm[k]; } return ""; };
  const out = {
    ID: pick(["id"]),
    Fecha: pick(["fecha"]),
    Area: pick(["area"]),
    Puesto_de_Trabajo: pick(["puesto_de_trabajo","puesto","cargo"]),
    Nombre_Trabajador: pick(["nombre_trabajador","nombre","trabajador"]),
    Sexo: pick(["sexo","genero"]),
    Edad: pick(["edad"]),
    Diestro_Zurdo: pick(["diestro_zurdo","lateralidad"]),
  };
  for(const [k,v] of Object.entries(objNorm)){
    if(["id","fecha","area","puesto_de_trabajo","puesto","cargo","nombre_trabajador","nombre","trabajador","sexo","genero","edad","diestro_zurdo","lateralidad"].includes(k)) continue;
    out[k] = v;
  }
  // normalizar tipos
  out.ID = String(out.ID || "").trim();
  out.Fecha = String(out.Fecha || "").trim();
  out.Area = String(out.Area || "").trim();
  out.Puesto_de_Trabajo = String(out.Puesto_de_Trabajo || "").trim();
  out.Nombre_Trabajador = String(out.Nombre_Trabajador || "").trim();
  out.Sexo = String(out.Sexo || "").trim();

  // Derivados de fecha (vigencia)
  const meta = computeDateMeta(out.Fecha);
  out.__date = meta.date;     // Date | null
  out.__dias = meta.days;     // number | null
  out.__vig = meta.status;    // 'vigente' | 'vencido' | 'desconocido'
  return out;
}
function computeDateMeta(fechaStr){
  // Soporta 'dd/mm/yyyy', 'yyyy-mm-dd', 'dd-mm-yyyy', etc.
  const s = String(fechaStr || "").trim();
  if(!s) return { date:null, days:null, status:"desconocido" };
  let d = null;

  // dd/mm/yyyy
  const m1 = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/.exec(s);
  if(m1){
    const dd = parseInt(m1[1],10), mm = parseInt(m1[2],10)-1, yy = parseInt(m1[3],10);
    const yyyy = yy < 100 ? 2000 + yy : yy;
    d = new Date(yyyy, mm, dd);
  }
  // yyyy-mm-dd
  if(!d){
    const m2 = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if(m2){
      d = new Date(parseInt(m2[1],10), parseInt(m2[2],10)-1, parseInt(m2[3],10));
    }
  }
  if(!d || isNaN(d.getTime())) return { date:null, days:null, status:"desconocido" };

  const now = new Date();
  const days = Math.floor((now - d) / (1000*60*60*24));
  const status = days <= 365 ? "vigente" : "vencido";
  return { date:d, days, status };
}
function normalizeSexoValue(s){
  const v = toLowerNoAccents(s);
  if(!v) return {code:"", label:""};
  if(v.startsWith("h")) return {code:"H", label:"Hombre"};
  if(v.startsWith("m")) return {code:"M", label:"Mujer"};
  if(v.startsWith("masc")) return {code:"H", label:"Hombre"};
  if(v.startsWith("fem"))  return {code:"M", label:"Mujer"};
  return {code:"", label:""};
}
function addCacheBuster(path){ return `${path}${path.includes("?") ? "&" : "?"}v=${Date.now()}`; }

// ---------- Carga principal ----------
async function loadExcelAuto(){
  try{
    if(typeof XLSX === "undefined"){ showError("SheetJS (XLSX) no está cargado en la página."); return; }

    const basePaths = [];
    if (window.RESP_XLSX) basePaths.push(window.RESP_XLSX);
    basePaths.push("../source/respuestas_cuestionario.xlsx");
    basePaths.push("/source/respuestas_cuestionario.xlsx");
    basePaths.push("./respuestas_cuestionario.xlsx");

    let parsed = null, lastErr = null;

    for(const p of basePaths){
      try{
        const url = addCacheBuster(p);
        const res = await fetch(url, { cache: "no-store" });
        if(!res.ok){ lastErr = `HTTP ${res.status} ${res.statusText} en ${url}`; continue; }
        const buf = await res.arrayBuffer();
        if(!buf || buf.byteLength < 50){ lastErr = `Archivo vacío o muy pequeño en ${url}`; continue; }
        parsed = parseWorkbook(buf);
        break;
      }catch(e){
        lastErr = e.message || String(e);
      }
    }
    if(!parsed){ showError(`No se pudo descargar/parsing del Excel. Último error: ${escapeHtml(lastErr||"desconocido")}`); return; }

    ROWS = parsed.rows;

    // Áreas
    const seenAreas = new Map();
    for(const r of ROWS){
      const key = toLowerNoAccents(r.Area);
      if(r.Area && !seenAreas.has(key)) seenAreas.set(key, r.Area);
    }
    AREAS = Array.from(seenAreas.values()).sort((a,b)=> a.localeCompare(b));

    // Sexos
    const seenSex = new Set();
    for(const r of ROWS){
      const {code} = normalizeSexoValue(r.Sexo);
      if(code) seenSex.add(code);
    }
    SEXOS = Array.from(seenSex.values());

    paintFilters();
    render();

  }catch(e){
    showError(e.message || "Error desconocido al cargar el Excel.");
  }
}

function parseWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  let ws = wb.Sheets["Respuestas"];
  if(!ws){
    const goal = toLowerNoAccents("Respuestas").replace(/\s+/g,"");
    for(const name of wb.SheetNames){
      const norm = toLowerNoAccents(name).replace(/\s+/g,"");
      if(norm === goal || norm.includes("respuesta")){ ws = wb.Sheets[name]; break; }
    }
  }
  ws = ws || wb.Sheets[wb.SheetNames[0]];
  if(!ws) throw new Error("El libro no contiene hojas.");

  const rows2D = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if(!rows2D.length) throw new Error("La hoja seleccionada está vacía.");

  const rawHeaders = rows2D[0] || [];
  const headersNorm = rawHeaders.map(normalizeHeader);
  const objs = rows2D.slice(1).map(r => rowToObj(headersNorm, r));
  const logical = objs.map(mapToLogical);

  return { rows: logical };
}

function handleLocalXlsx(evt){
  const file = evt.target.files && evt.target.files[0];
  if(!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    try{
      const buf = reader.result;
      const parsed = parseWorkbook(buf);
      ROWS = parsed.rows;

      const seenAreas = new Map();
      for(const r of ROWS){
        const key = toLowerNoAccents(r.Area);
        if(r.Area && !seenAreas.has(key)) seenAreas.set(key, r.Area);
      }
      AREAS = Array.from(seenAreas.values()).sort((a,b)=> a.localeCompare(b));

      const seenSex = new Set();
      for(const r of ROWS){
        const {code} = normalizeSexoValue(r.Sexo);
        if(code) seenSex.add(code);
      }
      SEXOS = Array.from(seenSex.values());

      paintFilters();
      byId("loadError").classList.add("d-none");
      PAGE = 1;
      render();
    }catch(e){
      showError("Error leyendo el archivo local: " + escapeHtml(e.message||"desconocido"));
    }
  };
  reader.onerror = () => showError("No se pudo leer el archivo local.");
  reader.readAsArrayBuffer(file);
}

// ---------- Cache local (hidratar rápido) ----------
const CACHE_KEY_RESP = "RESPUESTAS_CACHE_V1";
function tryLoadFromCacheFirst(){
  try{
    const cache = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "null");
    if(cache && Array.isArray(cache.rows)){
      ROWS = cache.rows.map(mapToLogical); // reprocesa para meta-fecha
      const seenAreas = new Map();
      for(const r of ROWS){
        const key = toLowerNoAccents(r.Area);
        if(r.Area && !seenAreas.has(key)) seenAreas.set(key, r.Area);
      }
      AREAS = Array.from(seenAreas.values()).sort((a,b)=> a.localeCompare(b));
      const seenSex = new Set();
      for(const r of ROWS){
        const {code} = normalizeSexoValue(r.Sexo);
        if(code) seenSex.add(code);
      }
      SEXOS = Array.from(seenSex.values());
      paintFilters();
      render();
    }
  }catch(_){}
}

// ---------- Filtros dinámicos ----------
function paintFilters(){
  // Área
  const selArea = byId("filterArea");
  selArea.innerHTML = `<option value="">(Todas)</option>` + AREAS.map(a => `<option>${escapeHtml(a)}</option>`).join("");

  // Sexo
  const selSexo = byId("filterSexo");
  const sexOptions = SEXOS.map(code => {
    const label = code === "H" ? "Hombre" : "Mujer";
    return `<option value="${code}">${label}</option>`;
  }).join("");
  selSexo.innerHTML = `<option value="">(Todos)</option>${sexOptions}`;
}

// ---------- Lógica de filtrado / orden ----------
function filteredRows(){
  const areaSel = toLowerNoAccents((byId("filterArea").value || "").trim());
  const sexSel = (byId("filterSexo").value || "").trim().toUpperCase(); // H|M|''
  const vigSel = (byId("filterVigencia").value || "").trim().toLowerCase(); // vigente|vencido|''
  const name = toLowerNoAccents((byId("filterName").value || "").trim());

  return ROWS.filter(r => {
    if(areaSel && toLowerNoAccents(r.Area) !== areaSel) return false;
    if(sexSel){
      const {code} = normalizeSexoValue(r.Sexo);
      if(code !== sexSel) return false;
    }
    if(vigSel){
      if((r.__vig || "desconocido") !== vigSel) return false;
    }
    if(name && !toLowerNoAccents(r.Nombre_Trabajador).includes(name)) return false;
    return true;
  });
}

function sortRows(list){
  const k = SORT.key, dir = SORT.dir === "asc" ? 1 : -1;
  return list.slice().sort((a,b) => {
    let av = a[k], bv = b[k];

    if(k === "ID"){
      const na = parseInt(av||0,10)||0, nb = parseInt(bv||0,10)||0;
      return (na - nb) * dir;
    }
    if(k === "Fecha"){
      const da = a.__date ? a.__date.getTime() : 0;
      const db = b.__date ? b.__date.getTime() : 0;
      return (da - db) * dir;
    }
    if(k === "Vigencia"){
      const order = { "vigente":2, "vencido":1, "desconocido":0 };
      return (order[(a.__vig||"desconocido")] - order[(b.__vig||"desconocido")]) * dir;
    }
    // texto
    av = toLowerNoAccents(String(av||""));
    bv = toLowerNoAccents(String(bv||""));
    if(av < bv) return -1*dir;
    if(av > bv) return  1*dir;
    return 0;
  });
}

// ---------- Render ----------
function render(){
  const all = sortRows(filteredRows());
  const total = all.length;
  const maxPage = Math.max(1, Math.ceil(total / PER_PAGE));
  if(PAGE > maxPage) PAGE = maxPage;

  const start = (PAGE - 1) * PER_PAGE;
  const data = all.slice(start, start + PER_PAGE);

  // Stats
  const vigentes = all.filter(r => r.__vig === "vigente").length;
  const vencidos = all.filter(r => r.__vig === "vencido").length;
  byId("statTotal").textContent = total;
  byId("statVig").textContent = vigentes;
  byId("statVen").textContent = vencidos;

  byId("countRows").textContent = data.length;
  byId("countRowsTotal").textContent = total;
  byId("pageCur").textContent = PAGE;
  byId("pageMax").textContent = maxPage;

  const tbody = byId("tblBody");
  if(!data.length){
    tbody.innerHTML = `<tr><td colspan="8" class="text-center text-muted py-4">
      <i class="bi bi-inboxes"></i> No hay registros para mostrar.
    </td></tr>`;
    return;
  }

  tbody.innerHTML = data.map(r => {
    const sexNorm = normalizeSexoValue(r.Sexo);
    const sexLabel = sexNorm.label || escapeHtml(r.Sexo);
    const vigBadge = r.__vig === "vigente"
      ? `<span class="badge bg-success">Vigente</span>`
      : r.__vig === "vencido"
      ? `<span class="badge bg-danger">Vencido</span>`
      : `<span class="badge bg-secondary">N/D</span>`;
    const fechaTxt = r.Fecha || (r.__date ? r.__date.toLocaleDateString() : "");
    return `
      <tr>
        <td>${escapeHtml(r.ID)}</td>
        <td>${escapeHtml(fechaTxt)}</td>
        <td>${escapeHtml(r.Area)}</td>
        <td>${escapeHtml(r.Nombre_Trabajador)}</td>
        <td>${escapeHtml(r.Puesto_de_Trabajo)}</td>
        <td>${sexLabel}</td>
        <td>${vigBadge}</td>
        <td>
          <a class="btn btn-sm btn-primary" href="cuestionario.html?id=${encodeURIComponent(r.ID)}">
            <i class="bi bi-eye"></i> Ver
          </a>
        </td>
      </tr>
    `;
  }).join("");
}

// ---------- Acciones ----------
function resetFilters(){
  byId("filterArea").value = "";
  byId("filterSexo").value = "";
  byId("filterVigencia").value = "";
  byId("filterName").value = "";
  PAGE = 1; render();
}
function resetPageAndRender(){ PAGE = 1; render(); }

function exportCsv(){
  const rows = sortRows(filteredRows());
  if(!rows.length){
    alert("No hay datos para exportar."); return;
  }
  const headers = ["ID","Fecha","Area","Puesto_de_Trabajo","Nombre_Trabajador","Sexo","Vigencia","Dias_desde"];
  const csv = [
    headers.join(","),
    ...rows.map(r => {
      const vig = r.__vig || "";
      const dias = (r.__dias==null ? "" : r.__dias);
      const vals = [
        r.ID, r.Fecha, r.Area, r.Puesto_de_Trabajo, r.Nombre_Trabajador, r.Sexo, vig, dias
      ].map(s => `"${String(s??"").replace(/"/g,'""')}"`);
      return vals.join(",");
    })
  ].join("\n");

  const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `cuestionarios_vista_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a); a.click();
  URL.revokeObjectURL(a.href); a.remove();
  /* ===== Ajuste dinámico para header sticky (evita solapamiento) ===== */
function setStickyTop(){
  const nav = document.querySelector('.navbar');
  // margen extra de 8px para que “respire”
  const stickyTop = (nav ? nav.offsetHeight : 56) + 8;
  document.documentElement.style.setProperty('--tableStickyTop', `${stickyTop}px`);
}

// calcula al cargar y cuando cambia el layout
window.addEventListener('load', setStickyTop);
window.addEventListener('resize', setStickyTop);
// por si Bootstrap colapsa/expande la navbar (menú hamburguesa)
document.addEventListener('shown.bs.collapse', setStickyTop, true);
document.addEventListener('hidden.bs.collapse', setStickyTop, true);
}
