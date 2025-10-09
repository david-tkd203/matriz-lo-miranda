// cuestionarios_respondidos.js (versión ultra-robusta + fallback manual)
//
// Qué mejora:
// - Prueba múltiples rutas del XLSX (relativa, absoluta y local) + cache-buster
// - Validación de existencia de XLSX (status y tamaño)
// - Detección de hoja "Respuestas" ignorando mayúsculas/acentos/espacios
// - Normalización de encabezados (trim, BOM, acentos, espacios)
// - Filtro ?area= case-insensitive y sin acentos
// - UI: mensaje de error visible + selector de archivo local para fallback

document.addEventListener("DOMContentLoaded", () => {
  ensureUiHooks();
  loadExcelAuto();
  byId("filterArea").addEventListener("change", render);
  byId("filterName").addEventListener("input", render);
  byId("btnPickXlsx").addEventListener("click", () => byId("fileXlsx").click());
  byId("fileXlsx").addEventListener("change", handleLocalXlsx);
});

let ROWS = [];
let AREAS = [];

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
  // Re-agregar handlers si el bloque se regeneró
  const btn = byId("btnPickXlsx");
  const inp = byId("fileXlsx");
  if(btn && inp){
    btn.addEventListener("click", () => inp.click());
    inp.addEventListener("change", handleLocalXlsx);
  }
}

function ensureUiHooks(){
  // Inserta contenedor de error si no existe
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
  let s = String(h).replace(/\uFEFF/g,"").trim(); // BOM + trim
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // quita acentos
  s = s.replace(/\s+/g,'_');         // espacios -> _
  s = s.replace(/[^\w]/g,'_');       // otros -> _
  s = s.replace(/_+/g,'_').replace(/^_|_$/g,''); // colapsa _
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
  const pick = (cands) => {
    for(const k of cands){ if(k in objNorm) return objNorm[k]; }
    return "";
  };
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
  // Tipo seguro
  out.ID = String(out.ID || "").trim();
  out.Fecha = String(out.Fecha || "").trim();
  out.Area = String(out.Area || "").trim();
  out.Puesto_de_Trabajo = String(out.Puesto_de_Trabajo || "").trim();
  out.Nombre_Trabajador = String(out.Nombre_Trabajador || "").trim();
  out.Sexo = String(out.Sexo || "").trim();
  return out;
}

// ---------- Carga principal ----------
async function loadExcelAuto(){
  try{
    if(typeof XLSX === "undefined"){
      showError("SheetJS (XLSX) no está cargado en la página.");
      return;
    }

    // Construimos una lista de rutas a intentar
    const basePaths = [];
    if (window.RESP_XLSX) basePaths.push(window.RESP_XLSX);
    // Alternativas razonables (según tu estructura / Netlify)
    basePaths.push("../source/respuestas_cuestionario.xlsx");
    basePaths.push("/source/respuestas_cuestionario.xlsx");
    basePaths.push("./respuestas_cuestionario.xlsx");

    let parsed = null, pickedUrl = null, lastErr = null;

    for(const p of basePaths){
      try{
        const url = addCacheBuster(p);
        const res = await fetch(url, { cache: "no-store" });
        if(!res.ok){
          lastErr = `HTTP ${res.status} ${res.statusText} en ${url}`;
          continue;
        }
        const buf = await res.arrayBuffer();
        if(!buf || buf.byteLength < 50){ // XLSX real suele ser > 1KB
          lastErr = `Archivo vacío o muy pequeño en ${url}`;
          continue;
        }
        parsed = parseWorkbook(buf);
        pickedUrl = url;
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

    // Áreas únicas
    const seen = new Map();
    for(const r of ROWS){
      const key = toLowerNoAccents(r.Area);
      if(r.Area && !seen.has(key)) seen.set(key, r.Area);
    }
    AREAS = Array.from(seen.values()).sort((a,b)=> a.localeCompare(b));

    // Pintar select
    const sel = byId("filterArea");
    sel.innerHTML = `<option value="">(Todas)</option>` + AREAS.map(a => `<option>${escapeHtml(a)}</option>`).join("");

    // ?area= soporte
    const qArea = getParam("area");
    if(qArea){
      const targetKey = toLowerNoAccents(qArea);
      const found = AREAS.find(a => toLowerNoAccents(a) === targetKey);
      if(found) sel.value = found;
    }

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
  // lee workbook y selecciona hoja "Respuestas" (flexible: sin acentos/espacios/case)
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  let ws = null;

  // 1) búsqueda directa
  ws = wb.Sheets["Respuestas"];
  if(!ws){
    // 2) buscar por similitud
    const goal = toLowerNoAccents("Respuestas").replace(/\s+/g,"");
    for(const name of wb.SheetNames){
      const norm = toLowerNoAccents(name).replace(/\s+/g,"");
      if(norm === goal || norm.includes("respuesta")){ ws = wb.Sheets[name]; break; }
    }
  }
  // 3) último recurso: primera hoja
  ws = ws || wb.Sheets[wb.SheetNames[0]];
  if(!ws) throw new Error("El libro no contiene hojas.");

  // Convertir a tabla 2D para limpiar headers
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

      // Áreas
      const seen = new Map();
      for(const r of ROWS){
        const key = toLowerNoAccents(r.Area);
        if(r.Area && !seen.has(key)) seen.set(key, r.Area);
      }
      AREAS = Array.from(seen.values()).sort((a,b)=> a.localeCompare(b));

      const sel = byId("filterArea");
      sel.innerHTML = `<option value="">(Todas)</option>` + AREAS.map(a => `<option>${escapeHtml(a)}</option>`).join("");
      byId("loadError").classList.add("d-none");
      render();
    }catch(e){
      showError("Error leyendo el archivo local: " + escapeHtml(e.message||"desconocido"));
    }
  };
  reader.onerror = () => showError("No se pudo leer el archivo local.");
  reader.readAsArrayBuffer(file);
}

// ---------- Render ----------
function render(){
  const areaSel = (byId("filterArea").value || "").trim();
  const areaKey = toLowerNoAccents(areaSel);
  const name = (byId("filterName").value || "").toLowerCase().trim();
  const tbody = byId("tblBody");

  const data = ROWS.filter(r => {
    if(areaKey && toLowerNoAccents(r.Area) !== areaKey) return false;
    if(name && !toLowerNoAccents(r.Nombre_Trabajador).includes(name)) return false;
    return true;
  });

  if(!data.length){
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">
      <i class="bi bi-inboxes"></i> No hay registros para mostrar.
    </td></tr>`;
    return;
  }

  tbody.innerHTML = data.map(r => `
    <tr>
      <td>${escapeHtml(r.ID)}</td>
      <td>${escapeHtml(r.Fecha)}</td>
      <td>${escapeHtml(r.Area)}</td>
      <td>${escapeHtml(r.Nombre_Trabajador)}</td>
      <td>${escapeHtml(r.Puesto_de_Trabajo)}</td>
      <td>${escapeHtml(r.Sexo)}</td>
      <td>
        <a class="btn btn-sm btn-primary" href="cuestionario.html?id=${encodeURIComponent(r.ID)}">
          <i class="bi bi-eye"></i> Ver
        </a>
      </td>
    </tr>
  `).join("");
}
