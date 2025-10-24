/* inicial_cuestionarios.js
   Panel derecho: “Cuestionarios por Área” con colapsables por Puesto → Personas
   - Lee respuestas desde ../source/respuestas_cuestionario.xlsx (hoja "Respuestas") o cache local
   - Sincroniza con filtros de la matriz (evento 'matrizFiltersChanged' y window.__MATRIZ_FILTERED_PUESTOS)
   - Muestra estado Vigente/Vencido por persona (por fecha más reciente)
*/

const VALID_DAYS = 365; // vigencia de un cuestionario (días)
let RESP_ROWS = [];     // filas de Respuestas
let HAS_DATA = false;

document.addEventListener("DOMContentLoaded", () => {
  // Hidratar desde cache local (si existe) para dar feedback rápido
  tryLoadCache();
  // Intentar XLSX del servidor
  tryLoadResponses();
  // Reaccionar a cambios de filtros de la matriz
  window.addEventListener('matrizFiltersChanged', () => renderAreasPanel());
});

function tryLoadCache(){
  try{
    const cache = JSON.parse(localStorage.getItem("RESPUESTAS_CACHE_V1") || "null");
    if(cache && Array.isArray(cache.rows)){
      RESP_ROWS = cache.rows.slice();
      HAS_DATA = true;
      renderAreasPanel();
    }
  }catch(_){}
}

async function tryLoadResponses(){
  try{
    if(!window.RESP_XLSX) return;
    const res = await fetch(addCacheBuster(window.RESP_XLSX), { cache:"no-store" });
    if(!res.ok) return;
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type:"array" });
    let ws = wb.Sheets["Respuestas"] || wb.Sheets[wb.SheetNames[0]];
    if(!ws) return;
    const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
    if(Array.isArray(rows) && rows.length){
      RESP_ROWS = rows;
      HAS_DATA = true;
      renderAreasPanel();
    }
  }catch(_){}
}

function addCacheBuster(path){
  const sep = path.includes("?") ? "&" : "?";
  return `${path}${sep}v=${Date.now()}`;
}

/* ======= Render panel ======= */
function renderAreasPanel(){
  const wrap = document.getElementById("areasList");
  const empty = document.getElementById("areasEmpty");
  if(!wrap || !empty) return;

  if(!HAS_DATA || !RESP_ROWS.length){
    wrap.innerHTML = "";
    empty.classList.remove("d-none");
    return;
  }
  empty.classList.add("d-none");

  // Tomar filtros actuales (expuestos por inicial.js)
  const filters = window.__MATRIZ_FILTERS || {};
  const allowed = window.__MATRIZ_FILTERED_PUESTOS || new Set();

  // Filtro por Área (obligatorio si viene)
  let rows = RESP_ROWS.slice();
  if(filters.area){
    rows = rows.filter(r => (r.Area||"").trim() === filters.area);
  }
  // Si hay set de puestos permitidos (por factor / tarea), restringir
  if(allowed.size){
    rows = rows.filter(r => allowed.has(`${(r.Area||"").trim()}|||${(r.Puesto_de_Trabajo||"").trim()}`));
  }else{
    // Si hay filtro de Puesto en la matriz y no hay allowed (por ejemplo sin factor),
    // aplicar al dataset de respuestas igualmente:
    if(filters.puesto){
      rows = rows.filter(r => (r.Puesto_de_Trabajo||"").trim() === filters.puesto);
    }
  }

  // Agrupar por Puesto y, dentro, por Persona → seleccionar fecha más reciente
  const map = new Map(); // puesto -> Map(persona -> {row, status})
  for(const r of rows){
    const puesto = String(r.Puesto_de_Trabajo||"").trim() || "(Sin puesto)";
    const persona = String(r.Nombre_Trabajador||"").trim() || "(Sin nombre)";
    const current = { date: parseFecha(r.Fecha), raw: r };
    if(!map.has(puesto)) map.set(puesto, new Map());
    const pMap = map.get(puesto);
    if(!pMap.has(persona)){
      pMap.set(persona, current);
    }else{
      // conservar el más reciente
      const prev = pMap.get(persona);
      if(current.date && (!prev.date || current.date > prev.date)) pMap.set(persona, current);
    }
  }

  if(map.size === 0){
    wrap.innerHTML = `<div class="alert alert-info small mb-0">No hay cuestionarios que coincidan con los filtros.</div>`;
    return;
  }

  // Construir accordion
  let idx = 0;
  let html = `<div class="accordion" id="accAreas">`;
  for(const [puesto, pMap] of map){
    const id = `accPuesto_${idx++}`;
    // Contar Vigente/Vencido
    let vig = 0, ven = 0;
    const persons = [];
    for(const [persona, info] of pMap){
      const { status } = statusFromDate(info.date);
      if(status === "Vigente") vig++; else ven++;
      persons.push({ persona, info, status });
    }
    persons.sort((a,b)=> a.persona.localeCompare(b.persona));

    html += `
      <div class="accordion-item">
        <h2 class="accordion-header" id="${id}_h">
          <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${id}_c" aria-expanded="false" aria-controls="${id}_c">
            <div class="w-100 d-flex justify-content-between align-items-center">
              <span><i class="bi bi-person-vcard"></i> <strong>${escapeHtml(puesto)}</strong></span>
              <span class="d-flex align-items-center gap-2">
                <span class="badge badge-vigente">Vigentes: ${vig}</span>
                <span class="badge badge-vencido">Vencidos: ${ven}</span>
              </span>
            </div>
          </button>
        </h2>
        <div id="${id}_c" class="accordion-collapse collapse" aria-labelledby="${id}_h" data-bs-parent="#accAreas">
          <div class="accordion-body p-0">
            <ul class="list-group list-group-flush">
              ${persons.map(p => {
                const fstr = formatFecha(infoToFechaString(p.info.raw.Fecha));
                const sex = (p.info.raw.Sexo||"").toString();
                return `
                <li class="list-group-item d-flex justify-content-between align-items-center">
                  <div>
                    <div class="fw-semibold">${escapeHtml(p.persona)}</div>
                    <div class="small text-muted">
                      ${escapeHtml(sex)} · Último: ${escapeHtml(fstr||"-")}
                    </div>
                  </div>
                  <span class="badge ${p.status === "Vigente" ? "badge-vigente" : "badge-vencido"}">${p.status}</span>
                </li>`;
              }).join("")}
            </ul>
          </div>
        </div>
      </div>
    `;
  }
  html += `</div>`;
  wrap.innerHTML = html;
}

/* ======= Fechas / estado ======= */
function parseFecha(s){
  if(!s) return null;
  // soporta dd/mm/yyyy o yyyy-mm-dd
  const t = String(s).trim();
  let d = null;
  if(/^\d{2}\/\d{2}\/\d{4}$/.test(t)){
    const [dd,mm,yy] = t.split('/').map(n=>parseInt(n,10));
    d = new Date(yy, mm-1, dd);
  }else if(/^\d{4}-\d{2}-\d{2}$/.test(t)){
    const [yy,mm,dd] = t.split('-').map(n=>parseInt(n,10));
    d = new Date(yy, mm-1, dd);
  }else{
    const tryDate = new Date(t);
    if(!isNaN(tryDate.getTime())) d = tryDate;
  }
  return (d && !isNaN(d.getTime())) ? d : null;
}
function daysDiff(from, to){
  const MS = 24*60*60*1000;
  const a = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  const b = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  return Math.round((b-a)/MS);
}
function statusFromDate(d){
  if(!d) return { status:"Vencido", days: null };
  const diff = daysDiff(d, new Date());
  return { status: diff <= VALID_DAYS ? "Vigente" : "Vencido", days: diff };
}
function infoToFechaString(s){ return String(s||"").trim(); }
function formatFecha(s){
  const d = parseFecha(s);
  if(!d) return "";
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const yy = d.getFullYear();
  return `${dd}/${mm}/${yy}`;
}

/* ====== Utils ====== */
function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
