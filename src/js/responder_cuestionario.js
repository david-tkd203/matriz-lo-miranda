// responder_cuestionario.js
// - Carga Excel existente (si está disponible)
// - Rellena lista de Áreas desde hoja Parametros (o Respuestas)
// - Arma el formulario de zonas del Nórdico
// - Agrega una fila a "Respuestas" y actualiza "Parametros"
// - Ofrece descarga del Excel actualizado (para que lo subas a src/source/)

const ZONAS = [
  "Cuello","Hombro Derecho","Hombro Izquierdo","Codo/antebrazo Derecho","Codo/antebrazo Izquierdo",
  "Muñeca/mano Derecha","Muñeca/mano Izquierda","Espalda Alta","Espalda Baja",
  "Caderas/nalgas/muslos","Rodillas (una o ambas)","Pies/tobillos (uno o ambos)"
];

let WORKBOOK = null;
let SHEET_RESP = "Respuestas";
let SHEET_PARAMS = "Parametros";
let HEADERS_RESP = []; // encabezados de "Respuestas"
let ROWS_RESP = [];    // datos de "Respuestas" como array de objetos
let PARAMS = [];       // [{Área, Hombres, Mujeres, Total}, ...]

document.addEventListener("DOMContentLoaded", async () => {
  initDateToday("Fecha");
  renderZones();
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

async function tryLoadExcel(){
  try{
    const url = window.RESP_XLSX || "../source/respuestas_cuestionario.xlsx";
    const res = await fetch(url, { cache: "no-store" });
    if(!res.ok){
      byId("loadError").classList.remove("d-none");
      byId("loadError").textContent = "No se pudo cargar el Excel actual; podrás crear uno nuevo al guardar.";
      WORKBOOK = XLSX.utils.book_new();
      return;
    }
    const buf = await res.arrayBuffer();
    WORKBOOK = XLSX.read(buf, { type: "array" });

    // Hoja Respuestas: si no existe, crear vacía con headers
    if(!WORKBOOK.Sheets[SHEET_RESP]){
      const ws = XLSX.utils.aoa_to_sheet([defaultHeaders()]);
      WORKBOOK.Sheets[SHEET_RESP] = ws;
      if(!WORKBOOK.SheetNames.includes(SHEET_RESP)) WORKBOOK.SheetNames.push(SHEET_RESP);
    }
    // Hoja Parametros: si no, crear
    if(!WORKBOOK.Sheets[SHEET_PARAMS]){
      const ws = XLSX.utils.aoa_to_sheet([["Área","Hombres","Mujeres","Total"]]);
      WORKBOOK.Sheets[SHEET_PARAMS] = ws;
      if(!WORKBOOK.SheetNames.includes(SHEET_PARAMS)) WORKBOOK.SheetNames.push(SHEET_PARAMS);
    }

    // Parsear Respuestas
    const wsResp = WORKBOOK.Sheets[SHEET_RESP];
    const rows2d = XLSX.utils.sheet_to_json(wsResp, { header:1, defval:"" });
    if(rows2d.length){
      HEADERS_RESP = rows2d[0].map(h => String(h||""));
      const objs = XLSX.utils.sheet_to_json(wsResp, { defval:"" });
      ROWS_RESP = objs;
    }else{
      HEADERS_RESP = defaultHeaders();
      ROWS_RESP = [];
    }

    // Parsear Parametros
    PARAMS = XLSX.utils.sheet_to_json(WORKBOOK.Sheets[SHEET_PARAMS], { defval:"" });

  }catch(e){
    byId("loadError").classList.remove("d-none");
    byId("loadError").textContent = "Error al leer Excel: " + (e.message||e);
    WORKBOOK = XLSX.utils.book_new();
    // crear hojas vacías
    WORKBOOK.Sheets[SHEET_RESP] = XLSX.utils.aoa_to_sheet([defaultHeaders()]);
    WORKBOOK.SheetNames.push(SHEET_RESP);
    WORKBOOK.Sheets[SHEET_PARAMS] = XLSX.utils.aoa_to_sheet([["Área","Hombres","Mujeres","Total"]]);
    WORKBOOK.SheetNames.push(SHEET_PARAMS);
    HEADERS_RESP = defaultHeaders();
    ROWS_RESP = [];
    PARAMS = [];
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

  // Preferimos hoja Parametros si viene
  if(Array.isArray(PARAMS) && PARAMS.length){
    for(const r of PARAMS){
      if(r["Área"]) setAreas.add(String(r["Área"]));
    }
  }
  // Si no hay, miramos Respuestas existentes
  if(setAreas.size === 0 && Array.isArray(ROWS_RESP)){
    for(const r of ROWS_RESP){
      if(r.Area) setAreas.add(String(r.Area));
    }
  }

  const ordered = Array.from(setAreas).sort((a,b)=> a.localeCompare(b));
  sel.innerHTML = `<option value="">Seleccione...</option>` + ordered.map(a => `<option>${a}</option>`).join("");
}

function wireForm(){
  const f = byId("formNordico");
  f.addEventListener("submit", onSubmit);
}

function onSubmit(evt){
  evt.preventDefault();
  // Construir fila nueva
  const row = buildRowFromForm();
  if(!row.Area || !row.Nombre_Trabajador || !row.Sexo){
    alert("Área, Nombre y Sexo son obligatorios.");
    return;
  }

  // Asegurar encabezados
  if(HEADERS_RESP.length === 0) HEADERS_RESP = defaultHeaders();

  // Normalizar ID
  const maxId = ROWS_RESP.reduce((m,r) => Math.max(m, parseInt(r.ID || 0,10)||0), 0);
  row.ID = String(maxId + 1);

  // Insertar en memoria
  ROWS_RESP.push(row);

  // Actualizar hoja Respuestas
  const rows2d = [HEADERS_RESP].concat(ROWS_RESP.map(r => HEADERS_RESP.map(h => r[h] ?? "")));
  WORKBOOK.Sheets[SHEET_RESP] = XLSX.utils.aoa_to_sheet(rows2d);
  if(!WORKBOOK.SheetNames.includes(SHEET_RESP)) WORKBOOK.SheetNames.push(SHEET_RESP);

  // Actualizar hoja Parametros (reconstruir a partir de ROWS_RESP)
  const counts = {};
  for(const r of ROWS_RESP){
    const area = String(r.Area||"").trim();
    const sexo = String(r.Sexo||"").trim();
    if(!area) continue;
    if(!counts[area]) counts[area] = {Hombres:0, Mujeres:0};
    if(/^h/i.test(sexo)) counts[area].Hombres += 1;
    else if(/^m/i.test(sexo)) counts[area].Mujeres += 1;
  }
  const paramsRows = [["Área","Hombres","Mujeres","Total"]];
  Object.keys(counts).sort((a,b)=> a.localeCompare(b)).forEach(a => {
    const H = counts[a].Hombres||0, M = counts[a].Mujeres||0;
    paramsRows.push([a, H, M, H+M]);
  });
  WORKBOOK.Sheets[SHEET_PARAMS] = XLSX.utils.aoa_to_sheet(paramsRows);
  if(!WORKBOOK.SheetNames.includes(SHEET_PARAMS)) WORKBOOK.SheetNames.push(SHEET_PARAMS);

  // Descargar workbook
  const wbout = XLSX.write(WORKBOOK, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "respuestas_cuestionario.xlsx";
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(a.href);
  a.remove();

  alert("¡Listo! Se descargó el Excel actualizado. Súbelo a src/source/ para verlo en la lista.");
  // Opcional: limpiar
  // byId("formNordico").reset();
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
  // Zonas
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
  // recibe yyyy-mm-dd
  const [y,m,d] = String(val).split("-");
  if(!y||!m||!d) return todayDMY();
  return `${d}/${m}/${y}`;
}
