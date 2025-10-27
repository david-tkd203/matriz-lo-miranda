/* inicial_cuestionarios.js
   - Lee respuestas_cuestionario.xlsx (o cache local del formulario)
   - Muestra panel derecho: por Área -> colapsa por Puesto -> lista personas con estado (Vigente/Próximo/Vencido)
   - Respeta filtros actuales de la matriz (área/puesto/tarea/factor) escuchando 'matrizFiltersChanged'
*/

(() => {
  const VIGENCIA_DIAS = 365;       // vigente si <= 365 días
  const PROXIMO_UMBRAL = 30;       // "próximo a vencer" si faltan <= 30 días

  // cache en memoria
  let ROWS = []; // respuestas del cuestionario

  // inicio
  document.addEventListener("DOMContentLoaded", () => {
    tryLoadFromCacheFirst();
    loadExcelRespuestas(); // hidratar si existe archivo
    // sincronía con filtros de matriz
    document.addEventListener("matrizFiltersChanged", renderPanel);
  });

  function byId(id){ return document.getElementById(id); }

  function tryLoadFromCacheFirst(){
    try{
      const cache = JSON.parse(localStorage.getItem("RESPUESTAS_CACHE_V1") || "null");
      if(cache && Array.isArray(cache.rows)) ROWS = cache.rows.slice();
      renderPanel();
    }catch(_){}
  }

  async function loadExcelRespuestas(){
    try{
      const res = await fetch(window.RESP_XLSX, { cache:"no-store" });
      if(!res.ok) return;
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type:"array" });
      const ws = wb.Sheets["Respuestas"] || wb.Sheets[wb.SheetNames[0]];
      if(!ws) return;
      const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
      // fusionar sin duplicar ID
      const seen = new Set(ROWS.map(r => String(r.ID||"")));
      rows.forEach(r => {
        const id = String(r.ID||"");
        if(id && !seen.has(id)){ ROWS.push(r); seen.add(id); }
      });
      renderPanel();
    }catch(_){}
  }

  function getFilters(){
    if(typeof window.getMatrizFilters === "function") return window.getMatrizFilters();
    return { area:"", puesto:"", tarea:"", factorKey:"", factorState:"" };
  }

  function parseDMY(s){
    // admite dd/mm/yyyy
    const m = String(s||"").match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if(!m) return null;
    const [_,d,mo,y] = m.map(Number);
    return new Date(y, mo-1, d);
  }

  function estadoVigencia(fechaDMY){
    const d = parseDMY(fechaDMY) || new Date(0);
    const hoy = new Date();
    const diff = Math.floor((hoy - d)/(1000*60*60*24));
    if(diff <= VIGENCIA_DIAS - PROXIMO_UMBRAL) return { cls:"status-vigente", label:"Vigente" };
    if(diff <= VIGENCIA_DIAS)                 return { cls:"status-proximo", label:"Próximo a vencer" };
    return { cls:"status-vencido", label:"Vencido" };
  }

  function renderPanel(){
    const { area, puesto, tarea } = getFilters();
    const mount = byId("areasList");
    const empty = byId("areasEmpty");

    // filtrar por área (respuestas pueden no tener tarea exacta; usamos área y puesto si está)
    let data = ROWS.slice();
    if(area)  data = data.filter(r => (r.Area||"") === area);
    if(puesto) data = data.filter(r => (r.Puesto_de_Trabajo||"") === puesto);

    if(!data.length){
      mount.innerHTML = "";
      empty.classList.remove("d-none");
      return;
    }
    empty.classList.add("d-none");

    // agrupar: Area -> Puesto -> personas
    const byPuesto = {};
    data.forEach(r => {
      const p = String(r.Puesto_de_Trabajo||"Sin Puesto");
      if(!byPuesto[p]) byPuesto[p] = [];
      byPuesto[p].push(r);
    });

    // resumen por área (H+M)
    const totalH = data.filter(r => /^h/i.test(String(r.Sexo||""))).length;
    const totalM = data.filter(r => /^m/i.test(String(r.Sexo||""))).length;

    // HTML
    const puestosHtml = Object.keys(byPuesto).sort((a,b)=> a.localeCompare(b)).map((p, idx) => {
      const pid = `p_${idx}`;
      const persons = byPuesto[p].sort((a,b)=> String(a.Nombre_Trabajador||"").localeCompare(String(b.Nombre_Trabajador||"")));
      const personsHtml = persons.map(r => {
        const st = estadoVigencia(r.Fecha);
        return `
          <div class="person-row">
            <div><i class="bi bi-person-circle"></i> ${escapeHtml(r.Nombre_Trabajador||"-")} <span class="text-muted">(${escapeHtml(r.Sexo||"-")})</span></div>
            <div class="status-pill ${st.cls}">${st.label}</div>
          </div>
        `;
      }).join("");

      return `
        <div class="mb-2">
          <button class="puesto-btn" data-bs-toggle="collapse" data-bs-target="#${pid}" aria-expanded="false">
            <span><i class="bi bi-person-badge"></i> ${escapeHtml(p)}</span>
            <i class="bi bi-chevron-down"></i>
          </button>
          <div id="${pid}" class="collapse mt-2">
            <div class="list-group list-group-flush">
              ${personsHtml || `<div class="text-muted small">Sin personas registradas.</div>`}
            </div>
          </div>
        </div>
      `;
    }).join("");

    mount.innerHTML = `
      <div class="area-block">
        <div class="area-head">
          <div><strong>${escapeHtml(area || "Todas las áreas")}</strong></div>
          <div class="d-flex gap-2">
            <span class="pill"><i class="bi bi-people"></i> H ${totalH} · M ${totalM}</span>
          </div>
        </div>
        <div class="mt-2">
          ${puestosHtml || `<div class="text-muted small">No hay puestos para mostrar.</div>`}
        </div>
      </div>
    `;
  }

  function escapeHtml(str){
    return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
  }
})();
