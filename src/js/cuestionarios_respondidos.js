// cuestionarios_respondidos.js (ultra-robusto + filtro por SEXO + búsqueda por nombre)
//
// - Prueba múltiples rutas del XLSX + cache-buster
// - Detección flexible de hoja "Respuestas"
// - Normaliza encabezados y mapea a claves lógicas
// - Filtros: Área, Sexo y Nombre (case/acentos-insensitive)
// - Fallback: permite cargar un archivo local si falla el fetch
// - Lee cache local si existe para hidratar rápido la UI

document.addEventListener("DOMContentLoaded", () => {
  ensureUiHooks();
  tryLoadFromCacheFirst();     // pinta algo si hay cache
  loadExcelAuto();             // luego intenta el archivo real

  byId("filterArea").addEventListener("change", render);
  byId("filterSexo").addEventListener("change", render);
  byId("filterName").addEventListener("input", render);
});

let ROWS = [];
let AREAS = [];
let SEXOS = []; // normalizados a 'H' / 'M' con etiquetas 'Hombre' / 'Mujer'

// ---------- Utils ----------
function byId(id){ return document.getElementById(id); }
function showError(msg){
  const el = byId("loadError");
  if(!el) return;
  el.classList.remove("d-none");
  el.innerHTML = `
    <div class="alert alert-warning d-flex align-items-start gap-2 mb-3">
      <i class="bi bi-exclamation-triangle mt-1"></i>
      <div>
        <div class="fw-semibold">No se pudo leer el Excel de cuestionarios.</div>
        <div class="small">${escapeHtml(msg)}</div>
        <div class="mt-2">
          <button id="btnPickXlsx" type="button" class="btn btn-sm btn-outline-primary">
            <i class="bi bi-file-earmark-spreadsheet"></i> Cargar Excel manualmente
          </button>
          <input id="fileXlsx" type="file" accept=".xlsx,.xls" class="d-none">
        </div>
      </div>
    </div>
  `;
  const btn = byId("btnPickXlsx");
  const inp = byId("fileXlsx");
  if(btn && inp){
    btn.addEventListener("click", () => inp.click());
    inp.addEventListener("change", handleLocalXlsx);
  }
}
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
function escapeHtml(str){
  return String(str ?? "").replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
function getParam(k){
  try{
    const u = new URL(location.href);
    return u.searchParams.get(k);
  }catch{ return null; }
}
function toLowerNoAccents(s){
  return String(s||"")
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim();
}
function normalizeHeader(h){
  if(h == null) return "";
  let s = String(h).replace(/\uFEFF/g,"").trim();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  s = s.replace(/\s+/g,'_');
  s = s.replace(/[^\w]/g,'_');
  s = s.replace(/_+/g,'_').replace(/^_|_$/g,'');
  return s.toLowerCase();
}
function rowToObj(headersNorm, rowArr){
  const o = {};
  headersNorm.forEach((keyNorm, idx) => {
    o[keyNorm] = rowArr[idx] == null ? "" : rowArr[idx];
  });
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
  out.ID = String(out.ID || "").trim();
  out.Fecha = String(out.Fecha || "").trim();
  out.Area = String(out.Area || "").trim();
  out.Puesto_de_Trabajo = String(out.Puesto_de_Trabajo || "").trim();
  out.Nombre_Trabajador = String(out.Nombre_Trabajador || "").trim();
  out.Sexo = String(out.Sexo || "").trim();
  return out;
}

// Normaliza sexo a 'H' | 'M' | '' y etiqueta a mostrar
function normalizeSexoValue(s){
  const v = toLowerNoAccents(s);
  if(!v) return {code:"", label:""};
  if(v.startsWith("h")) return {code:"H", label:"Hombre"};
  if(v.startsWith("m")) return {code:"M", label:"Mujer"};
  // si viniera algo raro como 'femenino'/'masculino'
  if(v.startsWith("masc")) return {code:"H", label:"Hombre"};
  if(v.startsWith("fem"))  return {code:"M", label:"Mujer"};
  return {code:"", label:""};
}

// ---------- Carga principal ----------
async function loadExcelAuto(){
  try{
    if(typeof XLSX === "undefined"){
      showError("SheetJS (XLSX) no está cargado en la página.");
      return;
    }

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

    if(!parsed){
      showError(`No se pudo descargar/parsing del Excel. Último error: ${escapeHtml(lastErr||"desconocido")}`);
      return;
    }

    const { rows } = parsed;
    ROWS = rows;

    // Áreas únicas (respetando primer casing visto)
    const seenAreas = new Map();
    for(const r of ROWS){
      const key = toLowerNoAccents(r.Area);
      if(r.Area && !seenAreas.has(key)) seenAreas.set(key, r.Area);
    }
    AREAS = Array.from(seenAreas.values()).sort((a,b)=> a.localeCompare(b));

    // Sexos únicos normalizados
    const seenSex = new Set();
    for(const r of ROWS){
      const {code} = normalizeSexoValue(r.Sexo);
      if(code) seenSex.add(code);
    }
    SEXOS = Array.from(seenSex.values()); // p.ej. ['H','M']

    paintFilters();
    render();

  }catch(e){
    showError(e.message || "Error desconocido al cargar el Excel.");
  }
}

function addCacheBuster(path){
  const sep = path.includes("?") ? "&" : "?";
  return `${path}${sep}v=${Date.now()}`;
}

function parseWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  // localizar hoja "Respuestas" de forma flexible
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

// ---------- Fallback: archivo local ----------
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
      render();
    }catch(e){
      showError("Error leyendo el archivo local: " + escapeHtml(e.message||"desconocido"));
    }
  };
  reader.onerror = () => showError("No se pudo leer el archivo local.");
  reader.readAsArrayBuffer(file);
}

// ---------- Pintar filtros dinámicos ----------
function paintFilters(){
  // Áreas
  const selArea = byId("filterArea");
  selArea.innerHTML = `<option value="">(Todas)</option>` + AREAS.map(a => `<option>${escapeHtml(a)}</option>`).join("");

  // Param ?area=
  const qArea = getParam("area");
  if(qArea){
    const found = AREAS.find(a => toLowerNoAccents(a) === toLowerNoAccents(qArea));
    if(found) selArea.value = found;
  }

  // Sexos
  const selSexo = byId("filterSexo");
  const sexOptions = SEXOS.map(code => {
    const label = code === "H" ? "Hombre" : "Mujer";
    return `<option value="${code}">${label}</option>`;
  }).join("");
  selSexo.innerHTML = `<option value="">(Todos)</option>${sexOptions}`;
}

// ---------- Render ----------
function render(){
  const areaSel = (byId("filterArea").value || "").trim();
  const areaKey = toLowerNoAccents(areaSel);

  const sexSel = (byId("filterSexo").value || "").trim().toUpperCase(); // 'H' | 'M' | ''

  const name = (byId("filterName").value || "").toLowerCase().trim();
  const tbody = byId("tblBody");

  const data = ROWS.filter(r => {
    // Área
    if(areaKey && toLowerNoAccents(r.Area) !== areaKey) return false;
    // Sexo (normalizado)
    if(sexSel){
      const {code} = normalizeSexoValue(r.Sexo);
      if(code !== sexSel) return false;
    }
    // Nombre
    if(name && !toLowerNoAccents(r.Nombre_Trabajador).includes(name)) return false;

    return true;
  });

  if(!data.length){
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">
      <i class="bi bi-inboxes"></i> No hay registros para mostrar.
    </td></tr>`;
    return;
  }

  tbody.innerHTML = data.map(r => {
    const sexNorm = normalizeSexoValue(r.Sexo);
    const sexLabel = sexNorm.label || escapeHtml(r.Sexo);
    return `
      <tr>
        <td>${escapeHtml(r.ID)}</td>
        <td>${escapeHtml(r.Fecha)}</td>
        <td>${escapeHtml(r.Area)}</td>
        <td>${escapeHtml(r.Nombre_Trabajador)}</td>
        <td>${escapeHtml(r.Puesto_de_Trabajo)}</td>
        <td>${sexLabel}</td>
        <td>
          <a class="btn btn-sm btn-primary" href="cuestionario.html?id=${encodeURIComponent(r.ID)}">
            <i class="bi bi-eye"></i> Ver
          </a>
        </td>
      </tr>
    `;
  }).join("");
}

// ---------- Cache local para hidratar rápido ----------
const CACHE_KEY_RESP = "RESPUESTAS_CACHE_V1";
function tryLoadFromCacheFirst(){
  try{
    const cache = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "null");
    if(cache && Array.isArray(cache.rows)){
      ROWS = cache.rows.slice();

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
window.addEventListener("cuestionariosCacheUpdated", tryLoadFromCacheFirst);
