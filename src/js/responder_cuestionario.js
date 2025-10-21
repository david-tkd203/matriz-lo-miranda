// 3 filas por sección (no tabla): 12m, Incap 12m + dolor, Incap 7d + dolor
// Mantiene cache local y estructura de columnas original (…__12m, …__Incap, …__Dolor12m, …__7d, …__Dolor7d)

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
  renderZones();
  loadFromCache();
  await tryLoadExcel();
  fillAreas();
  wireForm();
});

function byId(id){ return document.getElementById(id); }

/* ===== Fecha ===== */
function initDateToday(id){
  const el = byId(id); if(!el) return;
  const d = new Date();
  el.value = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

/* ===== Render de secciones ===== */
function renderZones(){
  const wrap = byId("zonesWrap");
  wrap.innerHTML = ZONAS.map(z => sectionHtml(z)).join("");
  wireSectionLogic();
  initSliderFill(); // color de sliders
}

function sectionHtml(z){
  const id12 = key(z,"12m");
  const idIn = key(z,"Incap");
  const idD12 = key(z,"Dolor12m");
  const id7  = key(z,"7d");        // aquí usamos “7d” como “Incapacidad 7 días”
  const idD7  = key(z,"Dolor7d");
  return `
    <section class="zone" data-zone="${z}">
      <div class="zone-header">
        <h6 class="zone-title">${z}</h6>
      </div>
      <div class="zone-body">
        <!-- Fila 1: Zona + 12 meses -->
        <div class="zone-row">
          <div class="zone-label"><i class="bi bi-calendar3"></i> Molestias en los últimos 12 meses</div>
          <div>
            <select class="form-select form-select-sm" id="${id12}">
              <option>NO</option><option>SI</option>
            </select>
          </div>
        </div>

        <!-- Fila 2: Incapacidad 12 m + dolor 12 m -->
        <div class="zone-row">
          <div class="zone-label"><i class="bi bi-person-exclamation"></i> Incapacidad (12 meses)</div>
          <div class="d-grid gap-2">
            <div>
              <select class="form-select form-select-sm" id="${idIn}" disabled>
                <option></option><option>NO</option><option>SI</option>
              </select>
            </div>
            <div>
              <div class="slider-wrap">
                <input type="range" min="0" max="10" step="1" value="0" id="${idD12}" disabled>
                <span class="value-badge" id="${idD12}_out">0</span>
              </div>
              <div class="ticks"><span>0</span><span>5</span><span>10</span></div>
              ${facesHtml(idD12)}
            </div>
          </div>
        </div>

        <!-- Fila 3: Incapacidad 7 d + dolor 7 d -->
        <div class="zone-row">
          <div class="zone-label"><i class="bi bi-activity"></i> Incapacidad (últimos 7 días)</div>
          <div class="d-grid gap-2">
            <div>
              <select class="form-select form-select-sm" id="${id7}" disabled>
                <option></option><option>NO</option><option>SI</option>
              </select>
            </div>
            <div>
              <div class="slider-wrap">
                <input type="range" min="0" max="10" step="1" value="0" id="${idD7}" disabled>
                <span class="value-badge" id="${idD7}_out">0</span>
              </div>
              <div class="ticks"><span>0</span><span>5</span><span>10</span></div>
              ${facesHtml(idD7)}
            </div>
          </div>
        </div>
      </div>
    </section>
  `;
}

function facesHtml(rangeId){
  return `
    <div class="faces" data-target="${rangeId}">
      ${FACE_STEPS.map(s => `
        <button type="button" class="face-btn" data-target-range="${rangeId}" data-value="${s.v}" title="${s.label}">
          <span class="face-emoji" style="background:${alpha(colorFor(s.v),.18)}; color:${colorFor(s.v)}">${s.emoji}</span>
          <span class="face-label">${s.v}</span>
        </button>
      `).join("")}
    </div>
  `;
}

function key(z,suf){ return `${z}__${suf}`; }

/* ===== Lógica de cada sección ===== */
function wireSectionLogic(){
  for(const z of ZONAS){
    const sel12 = byId(key(z,"12m"));
    const selIn = byId(key(z,"Incap"));
    const rngD12 = byId(key(z,"Dolor12m"));
    const sel7  = byId(key(z,"7d"));
    const rngD7 = byId(key(z,"Dolor7d"));

    const applyState = () => {
      const si12 = (sel12.value === "SI");
      if(!si12){
        selIn.disabled = true; selIn.value = "";
        rngD12.disabled = true; resetRange(rngD12); setFacesDisabled(rngD12.id, true);
        sel7.disabled = true; sel7.value = "";
        rngD7.disabled = true; resetRange(rngD7); setFacesDisabled(rngD7.id, true);
        return;
      }
      // 12m = SI
      selIn.disabled = false;
      const incap12 = (selIn.value === "SI");
      if(!incap12){
        rngD12.disabled = true; resetRange(rngD12); setFacesDisabled(rngD12.id, true);
        sel7.disabled = true;  sel7.value = "";
        rngD7.disabled = true; resetRange(rngD7); setFacesDisabled(rngD7.id, true);
        return;
      }
      // incap 12m = SI
      rngD12.disabled = false; setFacesDisabled(rngD12.id, false);
      sel7.disabled = false;
      const incap7 = (sel7.value === "SI");
      rngD7.disabled = !incap7; setFacesDisabled(rngD7.id, !incap7);
      if(!incap7) resetRange(rngD7);
      paintFill(rngD12); paintFill(rngD7);
    };

    sel12.addEventListener("change", applyState);
    selIn.addEventListener("change", applyState);
    sel7 .addEventListener("change", applyState);

    // sliders / caritas
    [rngD12, rngD7].forEach(r => {
      r.addEventListener("input", () => { updateBadge(r); paintFill(r); syncFaces(r); });
      updateBadge(r); paintFill(r); syncFaces(r);
    });

    // caritas click
    const container = document.querySelector(`.faces[data-target="${CSS.escape(rngD12.id)}"]`).parentElement.parentElement;
    document.getElementById(rngD12.id).closest('.zone-row').addEventListener("click", onFacesClick);
    document.getElementById(rngD7.id).closest('.zone-row').addEventListener("click", onFacesClick);

    applyState(); // estado inicial
  }
}

function onFacesClick(ev){
  const btn = ev.target.closest(".face-btn");
  if(!btn) return;
  const id = btn.dataset.targetRange;
  const rng = byId(id); if(!rng || rng.disabled) return;
  rng.value = btn.dataset.value;
  rng.dispatchEvent(new Event("input",{bubbles:true}));
}

function setFacesDisabled(rangeId, disabled){
  const wrap = document.querySelector(`.faces[data-target="${CSS.escape(rangeId)}"]`);
  if(!wrap) return;
  wrap.querySelectorAll(".face-btn").forEach(b => b.disabled = !!disabled);
}
function syncFaces(rng){
  const wrap = document.querySelector(`.faces[data-target="${CSS.escape(rng.id)}"]`);
  if(!wrap) return;
  const v = Number(rng.value);
  wrap.querySelectorAll(".face-btn").forEach(b => {
    b.setAttribute("aria-pressed", String(Number(b.dataset.value) === v));
  });
}

/* ===== Slider helpers ===== */
function resetRange(rng){
  rng.value = 0; updateBadge(rng); paintFill(rng); syncFaces(rng);
}
function updateBadge(rng){
  const v = Number(rng.value||0);
  const out = byId(`${rng.id}_out`);
  if(!out) return;
  out.textContent = v;
  const color = colorFor(v), bg = alpha(color,.15), br = alpha(color,.35);
  out.style.setProperty('--badge-bg', bg);
  out.style.setProperty('--badge-fg', color);
  out.style.setProperty('--badge-border', br);
  out.style.background = bg; out.style.color = color; out.style.borderColor = br;
  out.title = bandLabel(v);
}
function initSliderFill(){ document.querySelectorAll('input[type="range"]').forEach(paintFill); }
function paintFill(r){
  const min = Number(r.min||0), max = Number(r.max||10), val = Number(r.value||0);
  const pct = ((val-min)*100)/(max-min);
  r.style.setProperty('--_val', `${pct}%`);
  r.style.setProperty('--range-fill', colorFor(val));
}

/* ===== Colores por banda ===== */
function colorFor(v){ if(v<=3) return '#2ecc71'; if(v<=6) return '#f1c40f'; if(v<=8) return '#e67e22'; return '#e74c3c'; }
function bandLabel(v){ if(v<=3) return 'Leve (0–3)'; if(v<=6) return 'Moderado (4–6)'; if(v<=8) return 'Intenso (7–8)'; return 'Severo (9–10)'; }
function alpha(hex,a){ const h=hex.replace('#',''); const n=parseInt(h.length===3?h.split('').map(c=>c+c).join(''):h,16);
  const r=(n>>16)&255, g=(n>>8)&255, b=n&255; return `rgba(${r}, ${g}, ${b}, ${a})`; }

/* ===== Cache / Excel (hidrata si existe) ===== */
function loadFromCache(){
  try{
    const R = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "null");
    const P = JSON.parse(localStorage.getItem(CACHE_KEY_PARAMS) || "null");
    if(R && Array.isArray(R.headers) && Array.isArray(R.rows)){ HEADERS_RESP = R.headers.slice(); ROWS_RESP = R.rows.slice(); }
    if(P && Array.isArray(P)){ PARAMS = P.slice(); }
  }catch(_){}
}

async function tryLoadExcel(){
  try{
    const url = window.RESP_XLSX || "../source/respuestas_cuestionario.xlsx";
    const res = await fetch(url, { cache:"no-store" });
    if(!res.ok) return;
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:"array" });

    const wsResp = wb.Sheets["Respuestas"];
    if(wsResp){
      const fromFile = XLSX.utils.sheet_to_json(wsResp, { defval:"" });
      const seen = new Set(ROWS_RESP.map(r => String(r.ID||"")));
      for(const r of fromFile){ const id=String(r.ID||""); if(id && !seen.has(id)){ ROWS_RESP.push(r); seen.add(id); } }
      if(HEADERS_RESP.length===0){
        const rows2d = XLSX.utils.sheet_to_json(wsResp, { header:1, defval:"" });
        HEADERS_RESP = rows2d.length ? rows2d[0].map(h => String(h||"")) : defaultHeaders();
      }
    }else if(HEADERS_RESP.length===0){
      HEADERS_RESP = defaultHeaders();
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
  for(const z of ZONAS){ zonaCols.push(`${z}__12m`,`${z}__Incap`,`${z}__Dolor12m`,`${z}__7d`,`${z}__Dolor7d`); }
  return base.concat(zonaCols);
}

function fillAreas(){
  const sel = byId("Area");
  const setAreas = new Set();
  if(Array.isArray(PARAMS) && PARAMS.length){ for(const r of PARAMS){ if(r["Área"]) setAreas.add(String(r["Área"])); } }
  if(setAreas.size===0 && Array.isArray(ROWS_RESP)){ for(const r of ROWS_RESP){ if(r.Area) setAreas.add(String(r.Area)); } }
  const ordered = Array.from(setAreas).sort((a,b)=> a.localeCompare(b));
  sel.innerHTML = `<option value="">Seleccione...</option>` + ordered.map(a => `<option>${a}</option>`).join("");
}

/* ===== Guardado (en memoria) ===== */
function wireForm(){
  const form = byId("formNordico");
  form.addEventListener("submit", onSubmit);
  form.addEventListener("submit", (e) => {
    if (!form.checkValidity()){ e.preventDefault(); e.stopPropagation(); }
    form.classList.add('was-validated');
  }, { capture: true });
}

function onSubmit(evt){
  evt.preventDefault();
  const row = buildRowFromForm();
  if(!row.Area || !row.Nombre_Trabajador || !row.Sexo){ alert("Área, Nombre y Sexo son obligatorios."); return; }

  if(HEADERS_RESP.length===0) HEADERS_RESP = defaultHeaders();
  const maxId = ROWS_RESP.reduce((m,r)=>Math.max(m, parseInt(r.ID||0,10)||0), 0);
  row.ID = String(maxId + 1);
  ROWS_RESP.push(row);

  const counts = {};
  for(const r of ROWS_RESP){
    const area = String(r.Area||"").trim(), sexo = String(r.Sexo||"").trim();
    if(!area) continue;
    if(!counts[area]) counts[area] = {Hombres:0, Mujeres:0};
    if(/^h/i.test(sexo)) counts[area].Hombres += 1; else if(/^m/i.test(sexo)) counts[area].Mujeres += 1;
  }
  PARAMS = Object.keys(counts).sort((a,b)=>a.localeCompare(b)).map(a => ({
    "Área":a, "Hombres":counts[a].Hombres||0, "Mujeres":counts[a].Mujeres||0, "Total":(counts[a].Hombres||0)+(counts[a].Mujeres||0)
  }));

  persistCache();
  window.dispatchEvent(new CustomEvent("cuestionariosCacheUpdated"));
  alert(`¡Guardado! Respuesta #${row.ID} agregada en memoria.\nPuedes verla en "Cuestionarios Respondidos".`);
}

function persistCache(){
  try{
    localStorage.setItem(CACHE_KEY_RESP, JSON.stringify({ headers:HEADERS_RESP, rows:ROWS_RESP, ts:Date.now() }));
    localStorage.setItem(CACHE_KEY_PARAMS, JSON.stringify(PARAMS));
  }catch(_){}
}

function buildRowFromForm(){
  const fecha = byId("Fecha").value ? fmtDateDMY(byId("Fecha").value) : todayDMY();
  const row = {
    ID:"", Fecha:fecha, Area:byId("Area").value.trim(), Puesto_de_Trabajo:byId("Puesto_de_Trabajo").value.trim(),
    Nombre_Trabajador:byId("Nombre_Trabajador").value.trim(), Sexo:byId("Sexo").value.trim(), Edad:byId("Edad").value,
    Diestro_Zurdo:byId("Diestro_Zurdo").value, Temporadas_previas:byId("Temporadas_previas").value,
    N_temporadas:byId("N_temporadas").value, Tiempo_en_trabajo_meses:byId("Tiempo_en_trabajo_meses").value,
    Actividad_previa:byId("Actividad_previa").value.trim(), Otra_actividad:byId("Otra_actividad").value,
    Otra_actividad_cual:byId("Otra_actividad_cual").value.trim()
  };

  for(const z of ZONAS){
    const v12 = byId(key(z,"12m")).value;
    const vIn = byId(key(z,"Incap")).value;
    const vD12 = byId(key(z,"Dolor12m")).value;
    const v7  = byId(key(z,"7d")).value;         // Incapacidad 7 días
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

      if(vIn === "SI"){
        row[`${z}__Dolor12m`] = String(vD12 ?? "0");
        if(v7 === "SI"){
          row[`${z}__7d`] = "SI";
          row[`${z}__Dolor7d`] = String(vD7 ?? "0");
        }else if(v7 === "NO"){
          row[`${z}__7d`] = "NO";
          row[`${z}__Dolor7d`] = "";
        }else{ row[`${z}__7d`] = ""; row[`${z}__Dolor7d`] = ""; }
      }else{
        row[`${z}__Dolor12m`] = "";
        row[`${z}__7d`] = "";
        row[`${z}__Dolor7d`] = "";
      }
    }
  }
  return row;
}

/* Helpers fecha */
function todayDMY(){ const d=new Date(); return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`; }
function fmtDateDMY(val){ const [y,m,d]=String(val).split("-"); if(!y||!m||!d) return todayDMY(); return `${d}/${m}/${y}`; }
