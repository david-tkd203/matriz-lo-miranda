/* inicial_cuestionarios.js
   - Lee cuestionarios (Respuestas) desde Excel o cache local
   - Agrupa por Área y calcula estado: Vigente / Vencido (por fecha)
   - Pinta panel derecho con badges y última fecha por área
*/

const RESP_CACHE_KEY = "RESPUESTAS_CACHE_V1";  // el mismo que usa responder_cuestionario.js
const VALID_DAYS = 365;                        // vigencia: 12 meses

document.addEventListener("DOMContentLoaded", () => {
  // Hidratar desde cache primero, luego tratar de leer el Excel
  tryPaintFromCache();
  hydrateFromExcel();
  // refrescar automáticamente si se guarda algo nuevo en el formulario
  window.addEventListener("cuestionariosCacheUpdated", tryPaintFromCache);
});

function byId(id){ return document.getElementById(id); }

/* -------- Lecturas -------- */
function tryPaintFromCache(){
  try{
    const cache = JSON.parse(localStorage.getItem(RESP_CACHE_KEY) || "null");
    if(!cache || !Array.isArray(cache.rows)) return;
    paintAreasPanel(cache.rows);
  }catch(_){}
}

async function hydrateFromExcel(){
  try{
    const paths = [];
    if(window.RESP_XLSX) paths.push(window.RESP_XLSX);
    paths.push("../source/respuestas_cuestionario.xlsx","/source/respuestas_cuestionario.xlsx","./respuestas_cuestionario.xlsx");

    let rows = null, lastErr=null;
    for(const p of paths){
      try{
        const u = addBuster(p);
        const res = await fetch(u, { cache:"no-store" });
        if(!res.ok){ lastErr = `HTTP ${res.status}`; continue; }
        const buf = await res.arrayBuffer();
        if(!buf || buf.byteLength < 50){ lastErr = "archivo vacío/pequeño"; continue; }
        rows = parseRespuestas(buf);
        break;
      }catch(e){ lastErr = e.message || String(e); }
    }
    if(!rows){ /* sin Excel, ya pintamos desde cache si existía */ return; }
    paintAreasPanel(rows);
  }catch(_){}
}

function addBuster(p){ return `${p}${p.includes("?")?"&":"?"}v=${Date.now()}`; }

function parseRespuestas(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type:"array" });
  let ws = wb.Sheets["Respuestas"];
  if(!ws){
    const goal = norm("Respuestas");
    for(const n of wb.SheetNames){
      if(norm(n).includes("respuesta")){ ws = wb.Sheets[n]; break; }
    }
  }
  ws = ws || wb.Sheets[wb.SheetNames[0]];
  if(!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval:"" });
}

function norm(s){
  return String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g,"").toLowerCase();
}

/* -------- Lógica de vigencia -------- */
function parseFechaDMY(s){
  // soporta dd/mm/yyyy, yyyy-mm-dd, dd-mm-yyyy
  const str = String(s||"").trim();
  if(!str) return null;
  let d=null;
  if(/^\d{2}\/\d{2}\/\d{4}$/.test(str)){ // dd/mm/yyyy
    const [dd,mm,yy] = str.split("/").map(n=>parseInt(n,10));
    d = new Date(yy, mm-1, dd);
  }else if(/^\d{4}-\d{2}-\d{2}$/.test(str)){ // yyyy-mm-dd
    const [yy,mm,dd] = str.split("-").map(n=>parseInt(n,10));
    d = new Date(yy, mm-1, dd);
  }else if(/^\d{2}-\d{2}-\d{4}$/.test(str)){ // dd-mm-yyyy
    const [dd,mm,yy] = str.split("-").map(n=>parseInt(n,10));
    d = new Date(yy, mm-1, dd);
  }
  return Number.isFinite(d?.getTime()) ? d : null;
}
function daysDiff(a,b){ return Math.floor((a - b) / (1000*60*60*24)); }
function statusFromDate(d){
  if(!d) return { code:"NA", label:"Sin fecha" };
  const today = new Date();
  const diff = daysDiff(today, d);
  if(diff <= VALID_DAYS) return { code:"OK", label:"Vigente" };
  return { code:"KO", label:"Vencido" };
}
function fmtDMY(d){
  if(!d) return "";
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const yy = d.getFullYear();
  return `${dd}/${mm}/${yy}`;
}

/* -------- Pintado -------- */
function paintAreasPanel(rows){
  const wrap = byId("areasList");
  const empty = byId("areasEmpty");
  if(!wrap) return;

  // agrupar por área
  const map = new Map(); // area -> {H,M,total,vig,venc, lastDate}
  for(const r of rows){
    const area = String(r.Area||"").trim();
    if(!area) continue;
    const sexo = String(r.Sexo||"").trim().toLowerCase();
    const d = parseFechaDMY(r.Fecha);
    const st = statusFromDate(d);

    if(!map.has(area)) map.set(area, {H:0,M:0,total:0,vig:0,venc:0,lastDate:null});
    const o = map.get(area);

    if(/^h/.test(sexo)) o.H += 1;
    else if(/^m/.test(sexo)) o.M += 1;

    o.total += 1;
    if(st.code === "OK") o.vig += 1;
    else if(st.code === "KO") o.venc += 1;

    if(d && (!o.lastDate || d > o.lastDate)) o.lastDate = d;
  }

  const items = Array.from(map.entries()).sort((a,b)=> a[0].localeCompare(b[0]));
  if(items.length === 0){
    wrap.innerHTML = "";
    empty.classList.remove("d-none");
    return;
  }
  empty.classList.add("d-none");

  wrap.innerHTML = `
    <div class="list-group">
      ${items.map(([area, o]) => areaItem(area, o)).join("")}
    </div>
  `;

  // delegación: clic navega a la lista prefiltrada
  wrap.querySelectorAll("[data-area]").forEach(a => {
    a.addEventListener("click", () => {
      location.href = "cuestionarios_respondidos.html?area=" + encodeURIComponent(a.dataset.area);
    });
  });
}

function areaItem(area, o){
  const last = o.lastDate ? fmtDMY(o.lastDate) : "—";
  const statusBadge = o.venc > 0
    ? `<span class="badge text-bg-danger">Vencidos ${o.venc}</span>`
    : `<span class="badge text-bg-success">Vigentes ${o.vig}</span>`;

  return `
    <button type="button" class="list-group-item list-group-item-action d-flex justify-content-between align-items-center"
            data-area="${escapeHtml(area)}" title="Ver cuestionarios de ${escapeHtml(area)}">
      <div class="me-2">
        <div class="fw-semibold">${escapeHtml(area)}</div>
        <div class="small text-muted">
          H ${o.H} · M ${o.M} · Total ${o.total} · Última: ${last}
        </div>
      </div>
      <div class="d-flex align-items-center gap-2">
        ${statusBadge}
        <i class="bi bi-chevron-right text-muted"></i>
      </div>
    </button>
  `;
}

function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[s]);
}
