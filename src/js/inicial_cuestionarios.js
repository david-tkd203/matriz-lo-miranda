/* inicial_cuestionarios.js
   Panel derecho "Cuestionarios por Área"
   - Lee respuestas desde localStorage (cache) y/o Excel (../source/respuestas_cuestionario.xlsx)
   - Calcula vigencia por persona (última respuesta): vigente <= 12 meses
   - Render: acordeón por Área → colapsable por Puesto → lista de Personas con estado
*/

(() => {
  const AREAS_WRAP_ID = "areasList";
  const AREAS_EMPTY_ID = "areasEmpty";
  const CACHE_KEY_RESP = "RESPUESTAS_CACHE_V1";

  let RESP_ROWS = []; // filas de "Respuestas" (normalizadas)
  let INDEX = {};     // estructura por áreas

  document.addEventListener("DOMContentLoaded", () => {
    // 1) Pintar algo desde cache si hay
    tryLoadFromCache();
    // 2) Intentar hidratar desde Excel (si existe)
    tryFetchXlsx().finally(() => {
      buildIndex(); renderAreas();
    });

    // Si otro tab guarda nuevos formularios, refrescar
    window.addEventListener("cuestionariosCacheUpdated", () => {
      tryLoadFromCache();
      buildIndex(); renderAreas();
    });
  });

  /* ======== Carga de datos ======== */
  function tryLoadFromCache(){
    try{
      const cache = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "null");
      if(cache && Array.isArray(cache.rows)) {
        RESP_ROWS = cache.rows.slice();
      }
    }catch(_){}
  }

  async function tryFetchXlsx(){
    try{
      if(!window.RESP_XLSX) return;
      const url = withBuster(window.RESP_XLSX);
      const res = await fetch(url, { cache: "no-store" });
      if(!res.ok) return;
      const buf = await res.arrayBuffer();
      if(!buf || buf.byteLength < 64) return;

      const wb = XLSX.read(buf, { type: "array" });
      const ws = pickRespuestasSheet(wb);
      if(!ws) return;

      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      // Fusionar evitando duplicados por ID (quedarse con el más reciente si no hay fecha)
      const map = new Map(RESP_ROWS.map(r => [String(r.ID||""), r]));
      for(const r of rows){
        const k = String(r.ID||"");
        if(!k){ continue; }
        if(!map.has(k)){
          map.set(k, r);
        }
      }
      RESP_ROWS = Array.from(map.values());
    }catch(_){}
  }

  function pickRespuestasSheet(wb){
    let ws = wb.Sheets?.["Respuestas"];
    if(!ws){
      const goal = norm("Respuestas");
      for(const name of wb.SheetNames){
        if(norm(name).includes("respuesta")){ ws = wb.Sheets[name]; break; }
      }
    }
    return ws || wb.Sheets?.[wb.SheetNames[0]];
  }

  /* ======== Indexación y vigencia ======== */
  function buildIndex(){
    INDEX = {};
    // Agrupar por Área → Puesto → Persona (quedarse con la respuesta más reciente)
    const lastByPerson = new Map(); // key = area|puesto|nombre → fila más reciente
    for(const r of RESP_ROWS){
      const area  = String(r.Area||"").trim();
      const puesto= String(r.Puesto_de_Trabajo||"").trim();
      const nombre= String(r.Nombre_Trabajador||"").trim();
      if(!area || !nombre) continue;

      const key = `${area}||${puesto}||${nombre}`.toLowerCase();
      const prev = lastByPerson.get(key);
      if(!prev){ lastByPerson.set(key, r); continue; }

      const d1 = parseDateFlex(prev.Fecha);
      const d2 = parseDateFlex(r.Fecha);
      if(d2.getTime() > d1.getTime()) lastByPerson.set(key, r);
    }

    // Construir índice
    for(const [,row] of lastByPerson){
      const area  = String(row.Area||"").trim();
      const puesto= String(row.Puesto_de_Trabajo||"").trim() || "(Sin puesto)";
      const nombre= String(row.Nombre_Trabajador||"").trim();
      const sexo  = String(row.Sexo||"").trim();
      const fecha = String(row.Fecha||"").trim();
      const fechaD= parseDateFlex(fecha);
      const vigente = isVigente(fechaD);
      const id = String(row.ID||"");

      if(!INDEX[area]) INDEX[area] = { puestos: {}, total:0, vigentes:0, vencidos:0 };
      if(!INDEX[area].puestos[puesto]) INDEX[area].puestos[puesto] = [];
      INDEX[area].puestos[puesto].push({ id, nombre, sexo, fecha, vigente });

      INDEX[area].total++;
      if(vigente) INDEX[area].vigentes++; else INDEX[area].vencidos++;
    }

    // Ordenar personas por nombre dentro de cada puesto
    Object.values(INDEX).forEach(a => {
      Object.keys(a.puestos).forEach(p => {
        a.puestos[p].sort((x,y) => x.nombre.localeCompare(y.nombre));
      });
    });
  }

  function isVigente(d){
    if(!(d instanceof Date) || isNaN(d)) return false;
    const now = new Date();
    const diff = now.getTime() - d.getTime();
    const days = diff / (1000*60*60*24);
    return days <= 365; // vigente si <= 12 meses
  }

  function parseDateFlex(s){
    const t = String(s||"").trim();
    if(!t) return new Date(0);
    // dd/mm/yyyy
    if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)){
      const [dd,mm,yy] = t.split("/").map(x=>parseInt(x,10));
      return new Date(yy, mm-1, dd);
    }
    // yyyy-mm-dd
    if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(t)){
      const [yy,mm,dd] = t.split("-").map(x=>parseInt(x,10));
      return new Date(yy, mm-1, dd);
    }
    // Date parse fallback
    const d = new Date(t);
    return isNaN(d) ? new Date(0) : d;
  }

  function norm(s){
    return String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'').toLowerCase();
  }

  /* ======== Render ======== */
  function renderAreas(){
    const wrap = document.getElementById(AREAS_WRAP_ID);
    const empty = document.getElementById(AREAS_EMPTY_ID);
    if(!wrap) return;

    const areas = Object.keys(INDEX).sort((a,b)=> a.localeCompare(b));
    if(areas.length === 0){
      wrap.innerHTML = "";
      if(empty) empty.classList.remove("d-none");
      return;
    }
    if(empty) empty.classList.add("d-none");

    const accId = "areasAcc";
    let html = `<div class="accordion" id="${accId}">`;
    areas.forEach((area, i) => {
      const a = INDEX[area];
      const areaId = `area_${i}`;
      html += `
        <div class="accordion-item">
          <h2 class="accordion-header" id="h_${areaId}">
            <button class="accordion-button collapsed d-flex justify-content-between align-items-center" type="button"
                    data-bs-toggle="collapse" data-bs-target="#c_${areaId}" aria-expanded="false" aria-controls="c_${areaId}">
              <span class="me-2">${escapeHtml(area)}</span>
              <span class="ms-auto d-inline-flex gap-1">
                <span class="badge text-bg-success" title="Vigentes">${a.vigentes}</span>
                <span class="badge text-bg-danger"  title="Vencidos">${a.vencidos}</span>
                <span class="badge text-bg-secondary" title="Total">${a.total}</span>
              </span>
            </button>
          </h2>
          <div id="c_${areaId}" class="accordion-collapse collapse" aria-labelledby="h_${areaId}" data-bs-parent="#${accId}">
            <div class="accordion-body p-2">
              ${renderPuestos(area, a.puestos)}
            </div>
          </div>
        </div>`;
    });
    html += `</div>`;
    wrap.innerHTML = html;
  }

  function renderPuestos(area, puestosMap){
    const puestos = Object.keys(puestosMap).sort((a,b)=> a.localeCompare(b));
    if(puestos.length === 0) return `<div class="text-muted small">No hay puestos.</div>`;

    let html = `<div class="list-group list-group-flush">`;
    puestos.forEach((p, idx) => {
      const ppl = puestosMap[p];
      const vig = ppl.filter(x => x.vigente).length;
      const ven = ppl.length - vig;
      const pid = `pst_${norm(area)}_${idx}`;
      html += `
        <div class="list-group-item px-0">
          <div class="d-flex align-items-center">
            <button class="btn btn-sm btn-outline-secondary me-2" type="button" data-bs-toggle="collapse" data-bs-target="#${pid}" aria-expanded="false" aria-controls="${pid}">
              <i class="bi bi-caret-down"></i>
            </button>
            <div class="flex-grow-1">
              <div class="fw-semibold">${escapeHtml(p)}</div>
              <div class="small text-muted">Personas: ${ppl.length}
                · <span class="text-success">Vigentes: ${vig}</span>
                · <span class="text-danger">Vencidos: ${ven}</span>
              </div>
            </div>
          </div>
          <div id="${pid}" class="collapse mt-2">
            ${renderPersonas(ppl)}
          </div>
        </div>`;
    });
    html += `</div>`;
    return html;
  }

  function renderPersonas(list){
    if(!list || list.length===0) return `<div class="text-muted small ms-5">Sin personas.</div>`;
    let html = `<ul class="list-group list-group-flush ms-4">`;
    list.forEach(p => {
      const badge = p.vigente
        ? `<span class="badge rounded-pill text-bg-success">Vigente</span>`
        : `<span class="badge rounded-pill text-bg-danger">Vencido</span>`;
      html += `
        <li class="list-group-item px-0 d-flex align-items-center justify-content-between">
          <div class="me-2">
            <div class="fw-semibold">${escapeHtml(p.nombre)}</div>
            <div class="small text-muted">Sexo: ${escapeHtml(p.sexo || "-")} · Fecha: ${escapeHtml(p.fecha || "-")}</div>
          </div>
          <a class="btn btn-sm btn-outline-primary" href="cuestionario.html?id=${encodeURIComponent(p.id)}">
            <i class="bi bi-eye"></i> Ver ${badge}
          </a>
        </li>`;
    });
    html += `</ul>`;
    return html;
  }

  function escapeHtml(str){
    return String(str ?? "").replace(/[&<>"']/g, s => ({
      '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
    })[s]);
  }
  function withBuster(path){
    const sep = path.includes("?") ? "&" : "?";
    return `${path}${sep}v=${Date.now()}`;
  }
})();
