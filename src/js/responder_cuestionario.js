// responder_cuestionario.js (filas + sliders con colores + caritas clicables + cache local + FIX cascada Incap 12m)

const ZONAS = [
  "Cuello","Hombro Derecho","Hombro Izquierdo","Codo/antebrazo Derecho","Codo/antebrazo Izquierdo",
  "Muñeca/mano Derecha","Muñeca/mano Izquierda","Espalda Alta","Espalda Baja",
  "Caderas/nalgas/muslos","Rodillas (una o ambas)","Pies/tobillos (uno o ambos)"
];

const CACHE_KEY_RESP = "RESPUESTAS_CACHE_V1";
const CACHE_KEY_PARAMS = "PARAMS_CACHE_V1";

let HEADERS_RESP = [];
let ROWS_RESP = [];
let PARAMS = [];

/* 6 pasos inspirados en la figura */
const FACE_STEPS = [
  {v:0,  emoji:"😀", label:"Sin dolor"},
  {v:2,  emoji:"🙂", label:"Leve"},
  {v:4,  emoji:"😐", label:"Moderado"},
  {v:6,  emoji:"☹️", label:"Severo"},
  {v:8,  emoji:"😢", label:"Muy severo"},
  {v:10, emoji:"😭", label:"Máximo"}
];

document.addEventListener("DOMContentLoaded", async () => {
  initDateToday("Fecha");
  renderZonesTable();
  loadFromCache();
  await tryLoadExcel();
  fillAreas();
  wireForm();
});

function byId(id){ return document.getElementById(id); }

/* ========== Fecha por defecto ========== */
function initDateToday(id){
  const el = byId(id);
  if(!el) return;
  const d = new Date();
  el.value = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

/* ========== ZONAS EN FILAS + SLIDERS + CARITAS ========== */
function renderZonesTable(){
  const tbody = byId("zonesTbody");
  tbody.innerHTML = ZONAS.map(z => zoneRowHtml(z)).join("");
  wireZoneLogic();
  wireFacesClicks();
  initSliderFill();
}

function zoneRowHtml(z){
  const id12 = key(z,"12m");
  const idIn = key(z,"Incap");
  const idD12 = key(z,"Dolor12m");
  const id7 = key(z,"7d");
  const idD7 = key(z,"Dolor7d");
  return `
    <tr>
      <td>${z}</td>
      <td>
        <select class="form-select form-select-sm" id="${id12}">
          <option>NO</option>
          <option>SI</option>
        </select>
      </td>
      <td>
        <select class="form-select form-select-sm" id="${idIn}" disabled>
          <option></option>
          <option>NO</option>
          <option>SI</option>
        </select>
      </td>
      <td>
        <div class="slider-wrap">
          <input type="range" min="0" max="10" step="1" value="0" id="${idD12}" disabled>
          <span class="value-badge" id="${idD12}_out">0</span>
        </div>
        <div class="ticks"><span>0</span><span>5</span><span>10</span></div>
        ${facesHtml(idD12)}
      </td>
      <td>
        <select class="form-select form-select-sm" id="${id7}" disabled>
          <option></option>
          <option>NO</option>
          <option>SI</option>
        </select>
      </td>
      <td>
        <div class="slider-wrap">
          <input type="range" min="0" max="10" step="1" value="0" id="${idD7}" disabled>
          <span class="value-badge" id="${idD7}_out">0</span>
        </div>
        <div class="ticks"><span>0</span><span>5</span><span>10</span></div>
        ${facesHtml(idD7)}
      </td>
    </tr>
  `;
}

function facesHtml(targetRangeId){
  return `
    <div class="faces" data-target="${targetRangeId}">
      ${FACE_STEPS.map(s => `
        <button type="button"
          class="face-btn"
          data-target-range="${targetRangeId}"
          data-value="${s.v}"
          title="${s.label}">
          <span class="face-emoji" style="background:${alpha(colorFor(s.v),.18)}; color:${colorFor(s.v)}">${s.emoji}</span>
          <span class="face-label">${s.v}</span>
        </button>
      `).join("")}
    </div>
  `;
}

function key(z, suf){ return `${z}__${suf}`; }

/* ======= APLICA ESTADO DE UNA FILA SEGÚN 12m / INCAP / 7d ======= */
function applyRowState(z){
  const sel12 = byId(key(z,"12m"));
  const selIn = byId(key(z,"Incap"));
  const rngD12 = byId(key(z,"Dolor12m"));
  const sel7  = byId(key(z,"7d"));
  const rngD7 = byId(key(z,"Dolor7d"));

  const si12 = (sel12.value === "SI");

  // Si 12m = NO -> todo bloqueado/limpio
  if(!si12){
    selIn.disabled = true; selIn.value = "";
    rngD12.disabled = true; resetRange(rngD12);
    sel7.disabled  = true; sel7.value = "";
    rngD7.disabled = true; resetRange(rngD7);
    setFacesDisabled(rngD12.id, true);
    setFacesDisabled(rngD7.id,  true);
    return;
  }

  // 12m = SI -> Incapacidad decide el resto
  selIn.disabled = false;

  const incapSI = (selIn.value === "SI");
  if(!incapSI){
    // Incapacidad NO (o vacía) -> bloquear lo que sigue
    rngD12.disabled = true; resetRange(rngD12);
    sel7.disabled   = true; sel7.value = "";
    rngD7.disabled  = true; resetRange(rngD7);
    setFacesDisabled(rngD12.id, true);
    setFacesDisabled(rngD7.id,  true);
    return;
  }

  // Incapacidad = SI -> habilitar Dolor 12m y 7d
  rngD12.disabled = false;
  setFacesDisabled(rngD12.id, false);

  // 7d controla su slider
  sel7.disabled = false;
  const si7 = (sel7.value === "SI");
  rngD7.disabled = !si7;
  setFacesDisabled(rngD7.id, !si7);
  if(!si7) resetRange(rngD7);

  // repintar
  paintFill(rngD12);
  paintFill(rngD7);
}

function wireZoneLogic(){
  for(const z of ZONAS){
    const sel12 = byId(key(z,"12m"));
    const selIn = byId(key(z,"Incap"));
    const sel7  = byId(key(z,"7d"));
    const rngD12 = byId(key(z,"Dolor12m"));
    const rngD7  = byId(key(z,"Dolor7d"));

    sel12.addEventListener("change", () => applyRowState(z));
    selIn.addEventListener("change", () => applyRowState(z));
    sel7 .addEventListener("change", ()  => applyRowState(z));

    // Sliders -> actualiza burbuja/caritas y color
    for(const rng of [rngD12, rngD7]){
      rng.addEventListener("input", () => { updateBadge(rng); paintFill(rng); syncFaces(rng); });
      updateBadge(rng); paintFill(rng); syncFaces(rng);
    }

    // estado inicial
    applyRowState(z);
  }
}

/* ========== Caritas: interacción ========== */
function wireFacesClicks(){
  const tbody = byId("zonesTbody");
  tbody.addEventListener("click", (ev) => {
    const btn = ev.target.closest(".face-btn");
    if(!btn) return;
    const rangeId = btn.dataset.targetRange;
    const rng = byId(rangeId);
    if(!rng || rng.disabled) return;
    rng.value = btn.dataset.value;
    rng.dispatchEvent(new Event("input", {bubbles:true}));
  });
}

function setFacesDisabled(rangeId, disabled){
  const wrap = document.querySelector(`.faces[data-target="${CSS.escape(rangeId)}"]`);
  if(!wrap) return;
  wrap.querySelectorAll(".face-btn").forEach(b => { b.disabled = !!disabled; });
}

function syncFaces(rng){
  const wrap = document.querySelector(`.faces[data-target="${CSS.escape(rng.id)}"]`);
  if(!wrap) return;
  const v = Number(rng.value);
  wrap.querySelectorAll(".face-btn").forEach(b => {
    const val = Number(b.dataset.value);
    b.setAttribute("aria-pressed", String(v === val));
  });
}

/* ========== Slider y burbuja ========== */
function resetRange(rng){
  rng.value = 0;
  updateBadge(rng);
  paintFill(rng);
  syncFaces(rng);
}
function updateBadge(rng){
  const v = Number(rng.value||0);
  const out = byId(`${rng.id}_out`);
  if(out){
    out.textContent = v;
    const color = colorFor(v);
    const bg = alpha(color, .15);
    const br = alpha(color, .35);
    out.style.setProperty('--badge-bg', bg);
    out.style.setProperty('--badge-fg', color);
    out.style.setProperty('--badge-border', br);
    out.style.background = bg;
    out.style.color = color;
    out.style.borderColor = br;
    out.title = bandLabel(v);
  }
}
function initSliderFill(){
  document.querySelectorAll('input[type="range"]').forEach(paintFill);
}
function paintFill(r){
  const min = Number(r.min || 0), max = Number(r.max || 10), val = Number(r.value || 0);
  const pct = ((val - min) * 100) / (max - min);
  r.style.setProperty('--_val', `${pct}%`);
  r.style.setProperty('--range-fill', colorFor(val));
}

/* ========== Colores por banda (0–10) ========== */
function colorFor(v){
  if (v <= 3) return '#2ecc71';     // verde (leve)
  if (v <= 6) return '#f1c40f';     // amarillo (moderado)
  if (v <= 8) return '#e67e22';     // naranjo (intenso)
  return '#e74c3c';                  // rojo (severo)
}
function bandLabel(v){
  if (v <= 3) return 'Leve (0–3)';
  if (v <= 6) return 'Moderado (4–6)';
  if (v <= 8) return 'Intenso (7–8)';
  return 'Severo (9–10)';
}
function alpha(hex, a){
  const h = hex.replace('#','');
  const num = parseInt(h.length===3 ? h.split('').map(c=>c+c).join('') : h, 16);
  const r = (num >> 16) & 255, g = (num >> 8) & 255, b = num & 255;
  return `rgba(${r}, ${g}, ${b}, ${a})`;
}

/* ========== Cache / Excel (hidrata si existe) ========== */
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
    if(!res.ok) return;
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const wsResp = wb.Sheets["Respuestas"];
    if(wsResp){
      const fromFile = XLSX.utils.sheet_to_json(wsResp, { defval:"" });
      const seen = new Set(ROWS_RESP.map(r => String(r.ID||"")));
      for(const r of fromFile){
        const id = String(r.ID||"");
        if(id && !seen.has(id)){ ROWS_RESP.push(r); seen.add(id); }
      }
      if(HEADERS_RESP.length === 0){
        const rows2d = XLSX.utils.sheet_to_json(wsResp, { header:1, defval:"" });
        HEADERS_RESP = rows2d.length ? rows2d[0].map(h => String(h||"")) : defaultHeaders();
      }
    }else{
      if(HEADERS_RESP.length === 0) HEADERS_RESP = defaultHeaders();
    }

    const wsParams = wb.Sheets["Parametros"];
    if(wsParams){
      const fileParams = XLSX.utils.sheet_to_json(wsParams, { defval:"" });
      if(fileParams && fileParams.length) PARAMS = fileParams;
    }

    persistCache();
  }catch(_){}
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

/* ========== Guardado en memoria (sin descarga) ========== */
function wireForm(){
  const form = byId("formNordico");
  form.addEventListener("submit", onSubmit);
  form.addEventListener("submit", (e) => {
    if (!form.checkValidity()){
      e.preventDefault();
      e.stopPropagation();
    }
    form.classList.add('was-validated');
  }, { capture: true });
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

  persistCache();
  window.dispatchEvent(new CustomEvent("cuestionariosCacheUpdated"));

  alert(`¡Guardado! Respuesta #${row.ID} agregada en memoria.\nPuedes verla en "Cuestionarios Respondidos".`);
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
    const v12 = byId(key(z,"12m")).value;
    const vIn = byId(key(z,"Incap")).value;
    const vD12 = byId(key(z,"Dolor12m")).value;
    const v7 = byId(key(z,"7d")).value;
    const vD7 = byId(key(z,"Dolor7d")).value;

    if(v12 === "NO"){
      row[`${z}__12m`] = "NO";
      row[`${z}__Incap`] = "";
      row[`${z}__Dolor12m`] = "";
      row[`${z}__7d`] = "";
      row[`${z}__Dolor7d`] = "";
    }else{
      row[`${z}__12m`] = "SI";
      row[`${z}__Incap`] = vIn || "";

      // *** Regla de cascada corregida: Incapacidad SI habilita dolor/7d; NO (o vacío) los deja vacíos ***
      if(vIn === "SI"){
        row[`${z}__Dolor12m`] = String(vD12 ?? "0");
        if(v7 === "SI"){
          row[`${z}__7d`] = "SI";
          row[`${z}__Dolor7d`] = String(vD7 ?? "0");
        }else if(v7 === "NO"){
          row[`${z}__7d`] = "NO";
          row[`${z}__Dolor7d`] = "";
        }else{
          row[`${z}__7d`] = "";
          row[`${z}__Dolor7d`] = "";
        }
      }else{
        row[`${z}__Dolor12m`] = "";
        row[`${z}__7d`] = "";
        row[`${z}__Dolor7d`] = "";
      }
    }
  }
  return row;
}

/* ========== Helpers ========== */
function todayDMY(){
  const d = new Date();
  return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
}
function fmtDateDMY(val){
  const [y,m,d] = String(val).split("-");
  if(!y||!m||!d) return todayDMY();
  return `${d}/${m}/${y}`;
}
