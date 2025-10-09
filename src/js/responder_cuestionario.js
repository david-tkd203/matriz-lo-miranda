// responder_cuestionario.js (sin descarga; persiste en localStorage para la UI)

// Zonas del Nórdico
const ZONAS = [
  "Cuello","Hombro Derecho","Hombro Izquierdo","Codo/antebrazo Derecho","Codo/antebrazo Izquierdo",
  "Muñeca/mano Derecha","Muñeca/mano Izquierda","Espalda Alta","Espalda Baja",
  "Caderas/nalgas/muslos","Rodillas (una o ambas)","Pies/tobillos (uno o ambos)"
];

// Keys para cache local
const CACHE_KEY_RESP = "RESPUESTAS_CACHE_V1";
const CACHE_KEY_PARAMS = "PARAMS_CACHE_V1";

let WORKBOOK = null;
let SHEET_RESP = "Respuestas";
let SHEET_PARAMS = "Parametros";
let HEADERS_RESP = []; // encabezados
let ROWS_RESP = [];    // respuestas (objetos)
let PARAMS = [];       // [{Área, Hombres, Mujeres, Total}]

document.addEventListener("DOMContentLoaded", async () => {
  initDateToday("Fecha");
  renderZones();
  // 1) Intentar cargar cache local primero (si existe)
  loadFromCache();
  // 2) Luego intentar cargar Excel del servidor (para hidratar si hay)
  await tryLoadExcel();
  fillAreas();
  wireForm();
});

function byId(id){ return document.getElementById(id); }
function esc(s){ return String(s??""); }

function initDateToday(id){
  const el = byId(id);
  if(!el) return;
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  el.value = `${yyyy}-${mm}-${dd}`;
}

function renderZones(){
  const wrap = byId("zonesWrap");
  wrap.innerHTML = ZONAS.map(z => `
    <div class="zone-card">
      <h6>${z}</h6>
      <div class="zone-row">
        <div>
          <label class="form-label">Molestias 12 meses</label>
          <select class="form-select" id="${key(z,'12m')}">
            <option>NO</option>
            <option>SI</option>
          </select>
        </div>
        <div>
          <label class="form-label">Incapacidad (12 m)</label>
          <select class="form-select" id="${key(z,'Incap')}">
            <option></option>
            <option>NO</option>
            <option>SI</option>
          </select>
        </div>
        <div>
          <label class="form-label">Escala dolor (12 m) 1–10</label>
          <input type="number" min="1" max="10" class="form-control" id="${key(z,'Dolor12m')}">
        </div>
        <div>
          <label class="form-label">Molestias últimos 7 días</label>
          <select class="form-select" id="${key(z,'7d')}">
            <option></option>
            <option>NO</option>
            <option>SI</option>
          </select>
        </div>
        <div>
          <label class="form-label">Escala dolor (7 d) 1–10</label>
          <input type="number" min="1" max="10" class="form-control" id="${key(z,'Dolor7d')}">
        </div>
      </div>
    </div>
  `).join("");
}

function key(z, suf){ return `${z}__${suf}`; }

function loadFromCache(){
  try{
    const cacheResp = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "null");
    const cacheParams = JSON.parse(localStorage.getItem(CACHE_KEY_PARAMS) || "null");
    if(cacheResp && Array.isArray(cacheResp.headers) && Array.isArray(cacheResp.rows)){
      HEADERS_RESP = cacheResp.headers.slice();
      ROWS_RESP = cacheResp.rows.slice();
    }
    if(cacheParams && Array.isArray(cacheParams)){
      PARAMS = cacheParams.slice();
    }
  }catch(_){}
}

async function tryLoadExcel(){
  try{
    const url = window.RESP_XLSX || "../source/respuestas_cuestionario.xlsx";
    const res = await fetch(url, { cache: "no-store" });
    if(!res.ok){
      // Si no hay Excel en servidor, seguimos sólo con cache local
      WORKBOOK = XLSX.utils.book_new();
      return;
    }
    const buf = await res.arrayBuffer();
    WORKBOOK = XLSX.read(buf, { type: "array" });

    // Garantizar hojas
    if(!WORKBOOK.Sheets["Respuestas"]){
      WORKBOOK.Sheets["Respuestas"] = XLSX.utils.aoa_to_sheet([defaultHeaders()]);
      if(!WORKBOOK.SheetNames.includes("Respuestas")) WORKBOOK.SheetNames.push("Respuestas");
    }
    if(!WORKBOOK.Sheets["Parametros"]){
      WORKBOOK.Sheets["Parametros"] = XLSX.utils.aoa_to_sheet([["Área","Hombres","Mujeres","Total"]]);
      if(!WORKBOOK.SheetNames.includes("Parametros")) WORKBOOK.SheetNames.push("Parametros");
    }

    // Parsear Respuestas desde el archivo y fusionar con cache (sin duplicar IDs)
    const wsResp = WORKBOOK.Sheets["Respuestas"];
    const fileRows = XLSX.utils.sheet_to_json(wsResp, { defval:"" });
    const mapId = new Set(ROWS_RESP.map(r => String(r.ID||"")));
    for(const r of fileRows){
      const id = String(r.ID||"");
      if(id && !mapId.has(id)){
        ROWS_RESP.push(r);
        mapId.add(id);
      }
    }

    // Headers (si no los teníamos)
    if(HEADERS_RESP.length === 0){
      const rows2d = XLSX.utils.sheet_to_json(wsResp, { header:1, defval:"" });
      HEADERS_RESP = rows2d.length ? rows2d[0].map(h => String(h||"")) : defaultHeaders();
    }

    // Parametros
    const wsParams = WORKBOOK.Sheets["Parametros"];
    const fileParams = XLSX.utils.sheet_to_json(wsParams, { defval:"" });
    if(fileParams && fileParams.length) PARAMS = fileParams;

    // Persistir a cache (hidrata UI en otras páginas)
    persistCache();

  }catch(e){
    // si falla lectura, seguimos sólo con cache local
    WORKBOOK = XLSX.utils.book_new();
  }
}

function defaultHeaders(){
  const base = ["ID","Fecha","Area","Puesto_de_Trabajo","Nombre_Trabajador","Sexo","Edad","Diestro_Zurdo",
    "Temporadas_previas","N_temporadas","Tiempo_en_trabajo_meses","Actividad_previa","Otra_actividad","Otra_actividad_cual"];
  const zonaCols = [];
  for(const z of ZONAS){
    zonaCols.push(`${z}__12m`,`${z}__Incap`,`${z}__Dolor12m`,`${z}__7d`,`${z}__Dolor7d`);
  }
  return base.concat(zonaCols);
}

function fillAreas(){
  const sel = byId("Area");
  const setAreas = new Set();

  if(Array.isArray(PARAMS) && PARAMS.length){
    for(const r of PARAMS){
      if(r["Área"]) setAreas.add(String(r["Área"]));
    }
  }
  if(setAreas.size === 0 && Array.isArray(ROWS_RESP)){
    for(const r of ROWS_RESP){
      if(r.Area) setAreas.add(String(r.Area));
    }
  }

  const ordered = Array.from(setAreas).sort((a,b)=> a.localeCompare(b));
  sel.innerHTML = `<option value="">Seleccione...</option>` + ordered.map(a => `<option>${a}</option>`).join("");
}

function wireForm(){
  byId("formNordico").addEventListener("submit", onSubmit);
}

function onSubmit(evt){
  evt.preventDefault();

  const row = buildRowFromForm();
  if(!row.Area || !row.Nombre_Trabajador || !row.Sexo){
    alert("Área, Nombre y Sexo son obligatorios.");
    return;
  }

  if(HEADERS_RESP.length === 0) HEADERS_RESP = defaultHeaders();

  const maxId = ROWS_RESP.reduce((m,r) => Math.max(m, parseInt(r.ID || 0,10)||0), 0);
  row.ID = String(maxId + 1);

  ROWS_RESP.push(row);

  // Recalcular "Parametros" desde ROWS_RESP
  const counts = {};
  for(const r of ROWS_RESP){
    const area = String(r.Area||"").trim();
    const sexo = String(r.Sexo||"").trim();
    if(!area) continue;
    if(!counts[area]) counts[area] = {Hombres:0, Mujeres:0};
    if(/^h/i.test(sexo)) counts[area].Hombres += 1;
    else if(/^m/i.test(sexo)) counts[area].Mujeres += 1;
  }
  PARAMS = Object.keys(counts).sort((a,b)=> a.localeCompare(b)).map(a => ({
    "Área": a,
    "Hombres": counts[a].Hombres||0,
    "Mujeres": counts[a].Mujeres||0,
    "Total": (counts[a].Hombres||0) + (counts[a].Mujeres||0)
  }));

  // Persistir en localStorage para que la lista/detalle lo vean sin recargar del servidor
  persistCache();

  // Notificar a otras pestañas o scripts que usan la cache
  window.dispatchEvent(new CustomEvent("cuestionariosCacheUpdated"));

  // Feedback en UI
  alert(`¡Guardado! Respuesta #${row.ID} agregada en memoria.\nPuedes verla en "Cuestionarios Respondidos".`);
  // Opcional: redirigir directo a la lista:
  // location.href = "cuestionarios_respondidos.html?area=" + encodeURIComponent(row.Area);
}

function persistCache(){
  try{
    localStorage.setItem(CACHE_KEY_RESP, JSON.stringify({
      headers: HEADERS_RESP,
      rows: ROWS_RESP,
      ts: Date.now()
    }));
    localStorage.setItem(CACHE_KEY_PARAMS, JSON.stringify(PARAMS));
  }catch(_){}
}

function buildRowFromForm(){
  const fecha = byId("Fecha").value ? fmtDateDMY(byId("Fecha").value) : todayDMY();
  const row = {
    ID: "",
    Fecha: fecha,
    Area: byId("Area").value.trim(),
    Puesto_de_Trabajo: byId("Puesto_de_Trabajo").value.trim(),
    Nombre_Trabajador: byId("Nombre_Trabajador").value.trim(),
    Sexo: byId("Sexo").value.trim(),
    Edad: byId("Edad").value,
    Diestro_Zurdo: byId("Diestro_Zurdo").value,
    Temporadas_previas: byId("Temporadas_previas").value,
    N_temporadas: byId("N_temporadas").value,
    Tiempo_en_trabajo_meses: byId("Tiempo_en_trabajo_meses").value,
    Actividad_previa: byId("Actividad_previa").value.trim(),
    Otra_actividad: byId("Otra_actividad").value,
    Otra_actividad_cual: byId("Otra_actividad_cual").value.trim()
  };
  for(const z of ZONAS){
    row[`${z}__12m`] = byId(key(z,"12m")).value;
    row[`${z}__Incap`] = byId(key(z,"Incap")).value;
    row[`${z}__Dolor12m`] = byId(key(z,"Dolor12m")).value;
    row[`${z}__7d`] = byId(key(z,"7d")).value;
    row[`${z}__Dolor7d`] = byId(key(z,"Dolor7d")).value;
  }
  return row;
}

function todayDMY(){
  const d = new Date();
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}
function fmtDateDMY(val){
  const [y,m,d] = String(val).split("-");
  if(!y||!m||!d) return todayDMY();
  return `${d}/${m}/${y}`;
}
