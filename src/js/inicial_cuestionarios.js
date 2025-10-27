/* inicial_cuestionarios.js
  Panel derecho: “Cuestionarios por Área”
  - Lee ../source/respuestas_cuestionario.xlsx (hoja "Respuestas" y "Parametros")
  - Normaliza encabezados y mapea a llaves lógicas (ID, Fecha, Area, Puesto_de_Trabajo, etc.)
  - Renderiza por Área > Puesto_de_Trabajo (collapse), mostrando a las personas con estado:
     Vigente (≤ 335 días), Próximo (336–365), Vencido (> 365)
  - Se sincroniza con los filtros del panel izquierdo:
     * Área y Puesto de trabajo afectan esta vista (tarea/factor no aplican porque no existen en el XLSX de respuestas)
  - Fallback: si no hay XLSX, intenta cache local (RESPUESTAS_CACHE_V1) y puedes habilitar un “seeder” opcional
*/

(function(){
  // ---- Config ----
  const RESP_XLSX = window.RESP_XLSX || "../source/respuestas_cuestionario.xlsx";
  const ENABLE_DEMO_SEED = false; // ponlo en true si quieres autogenerar datos cuando no existan

  // ---- Estado ----
  let RESP_ROWS = [];   // objetos normalizados (ID, Fecha, Area, Puesto_de_Trabajo, Nombre_Trabajador, Sexo, ...)
  let PARAMS = [];      // [{Área,Hombres,Mujeres,Total}, ...] si viene de hoja Parametros

  // ---- Utilidades ----
  const $ = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  function escapeHtml(str){ return String(str ?? "").replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s])); }

  function normalizeHeader(h){
   if(h == null) return "";
   let s = String(h).replace(/\uFEFF/g,"").trim();
   s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
   s = s.replace(/\s+/g,'_').replace(/[^\w]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');
   return s.toLowerCase();
  }
  function rowToObj(headersNorm, rowArr){
   const o = {};
   headersNorm.forEach((k, i)=>{ o[k] = rowArr[i] == null ? "" : rowArr[i]; });
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
    Puesto_de_Trabajo: pick(["puesto_de_trabajo","puesto","cargo"]),  // <- CORRECTO
    Nombre_Trabajador: pick(["nombre_trabajador","nombre","trabajador"]),
    Sexo: pick(["sexo","genero"]),
    Edad: pick(["edad"]),
    Diestro_Zurdo: pick(["diestro_zurdo","lateralidad"]),
    Actividad_previa: pick(["actividad_previa"]),
   };
   // resto de columnas
   for(const [k,v] of Object.entries(objNorm)){
    if(k in out) continue;
    out[k] = v;
   }
   // saneo básico
   for(const k of Object.keys(out)) out[k] = String(out[k] ?? "").trim();
   return out;
  }
  function toLowerNoAccents(s){
   return String(s||"").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
  }

  // dd/mm/yyyy -> Date
  function parseDMY(dmy){
   const m = String(dmy||"").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
   if(!m) return null;
   const dd = +m[1], mm = +m[2], yy = +m[3];
   const dt = new Date(yy, mm-1, dd);
   return Number.isFinite(dt.getTime()) ? dt : null;
  }
  function daysAgo(dt){
   const today = new Date(); today.setHours(0,0,0,0);
   const d = new Date(dt); d.setHours(0,0,0,0);
   return Math.floor((today - d)/86400000);
  }
  function computeStatus(fechaDMY){
   const dt = parseDMY(fechaDMY);
   if(!dt) return {key:"na", label:"Sin fecha", cls:"pill-na"};
   const d = daysAgo(dt);
   if(d <= 335) return {key:"vig", label:"Vigente", cls:"pill-ok"};
   if(d <= 365) return {key:"prox", label:"Próximo a vencer", cls:"pill-warn"};
   return {key:"venc", label:"Vencido", cls:"pill-bad"};
  }

  // ---- Carga ----
  document.addEventListener("DOMContentLoaded", async () => {
   // Escuchar cambios de filtros de la izquierda
   ["filterArea","filterPuesto","filterTarea","filterFactor","filterFactorState"].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.addEventListener("change", renderPanel);
   });

   // 1) Intentar cache local (para que se vea algo rápido)
   tryLoadFromCache();

   // 2) Intentar Excel real
   await loadResponsesXlsx();

   // 3) Si no hay nada y quieres demo
   if(ENABLE_DEMO_SEED && RESP_ROWS.length===0){
    seedFewDemo();
   }

   renderPanel();

   // Si el formulario guarda algo nuevo, refrescamos
   window.addEventListener("cuestionariosCacheUpdated", () => {
    tryLoadFromCache();
    renderPanel();
   });
  });

  function tryLoadFromCache(){
   try{
    const cache = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "null");
    if(cache && Array.isArray(cache.rows)){
      RESP_ROWS = cache.rows.map(r => ({
       ...r,
       // asegurar campos esenciales como strings
       Area: String(r.Area||"").trim(),
       Puesto_de_Trabajo: String(r.Puesto_de_Trabajo||"").trim(),
       Nombre_Trabajador: String(r.Nombre_Trabajador||"").trim(),
       Sexo: String(r.Sexo||"").trim(),
       Fecha: String(r.Fecha||"").trim()
      }));
    }
   }catch(_){}
  }

  async function loadResponsesXlsx(){
   if(typeof XLSX === "undefined") return; // la página ya lo carga en <head>
   const tryPaths = [
    RESP_XLSX,
    "../source/respuestas_cuestionario.xlsx",
    "/source/respuestas_cuestionario.xlsx",
    "./respuestas_cuestionario.xlsx"
   ];
   for(const p of tryPaths){
    try{
      const url = p + (p.includes("?") ? "&" : "?") + "v=" + Date.now();
      const res = await fetch(url, { cache:"no-store" });
      if(!res.ok) continue;
      const buf = await res.arrayBuffer();
      if(!buf || buf.byteLength < 50) continue;

      const wb = XLSX.read(buf, { type:"array" });

      // ---- Respuestas ----
      let wsResp = wb.Sheets["Respuestas"];
      if(!wsResp){
       // buscar por similitud
       const goal = toLowerNoAccents("Respuestas").replace(/\s+/g,"");
       for(const n of wb.SheetNames){
        const norm = toLowerNoAccents(n).replace(/\s+/g,"");
        if(norm === goal || norm.includes("respuesta")) { wsResp = wb.Sheets[n]; break; }
       }
      }
      if(wsResp){
       const rows2D = XLSX.utils.sheet_to_json(wsResp, { header:1, defval:"" });
       if(rows2D.length){
        const headersNorm = rows2D[0].map(normalizeHeader);
        const objs = rows2D.slice(1).map(r => rowToObj(headersNorm, r));
        RESP_ROWS = objs.map(mapToLogical);
       }
      }

      // ---- Parametros (opcional) ----
      let wsParams = wb.Sheets["Parametros"];
      if(wsParams){
       PARAMS = XLSX.utils.sheet_to_json(wsParams, { defval:"" });
      }

      return; // ¡listo!
    }catch(_){}
   }
  }

  // Semillador opcional (solo cache) para ver estados vencidos
  function seedFewDemo(){
   const demo = [
    {ID:"1001", Fecha:"10/10/2025", Area:"Congelados", Puesto_de_Trabajo:"Operario/a túnel", Nombre_Trabajador:"María López", Sexo:"Mujer"},
    {ID:"1002", Fecha:"12/07/2025", Area:"Congelados", Puesto_de_Trabajo:"Operario/a túnel", Nombre_Trabajador:"Juan Pérez", Sexo:"Hombre"},
    {ID:"1003", Fecha:"05/11/2024", Area:"Faenado",    Puesto_de_Trabajo:"Desposte",      Nombre_Trabajador:"Ana Rojas",  Sexo:"Mujer"},
    {ID:"1004", Fecha:"20/05/2024", Area:"Faenado",    Puesto_de_Trabajo:"Desposte",      Nombre_Trabajador:"Carlos Muñoz", Sexo:"Hombre"},
    {ID:"1005", Fecha:"14/03/2023", Area:"Embalaje",   Puesto_de_Trabajo:"Embalador/a",   Nombre_Trabajador:"Luis González", Sexo:"Hombre"},
   ];
   // Mezclar con lo que ya había (evitar duplicar IDs)
   const seen = new Set(RESP_ROWS.map(r => r.ID));
   for(const r of demo){
    if(!seen.has(r.ID)) RESP_ROWS.push(r);
   }
   // persistir en cache para que otras vistas lo vean
   try{
    const cache = JSON.parse(localStorage.getItem(CACHE_KEY_RESP) || "{}");
    localStorage.setItem(CACHE_KEY_RESP, JSON.stringify({
      headers: Array.isArray(cache.headers) ? cache.headers : [],
      rows: RESP_ROWS,
      ts: Date.now()
    }));
   }catch(_){}
  }

  // ---- Render principal ----
  function renderPanel(){
   const host = document.getElementById("areasList");
   const empty = document.getElementById("areasEmpty");
   if(!host) return;

   const selArea = ($("#filterArea")?.value || "").trim();
   const selPuesto = ($("#filterPuesto")?.value || "").trim();

   // filtrar por Área y Puesto (los mismos filtros de la izquierda)
   let data = RESP_ROWS.slice();
   if(selArea)   data = data.filter(r => r.Area === selArea);
   if(selPuesto) data = data.filter(r => r.Puesto_de_Trabajo === selPuesto);

   if(!data.length){
    host.innerHTML = "";
    empty?.classList.remove("d-none");
    return;
   }
   empty?.classList.add("d-none");

   // Resumen superior: contar vigentes y vencidos (del conjunto filtrado)
   const totalSummary = summarizeStatuses(data);
   const topHtml = `
    <div class="panel-summary d-flex align-items-center gap-3 mb-2">
      <div class="small text-muted">Formularios:</div>
      <div class="d-flex align-items-center gap-2">
       <div class="d-flex align-items-center"><strong class="me-1 small">Vigentes</strong>${statusPill("vig", totalSummary.vig)}</div>
       <div class="d-flex align-items-center"><strong class="me-1 small">Vencidos</strong>${statusPill("venc", totalSummary.venc)}</div>
      </div>
    </div>
   `;

   // Agrupar Área > Puesto_de_Trabajo
   const byArea = groupBy(data, r => r.Area || "(Sin área)");
   // construir HTML
   let html = "";
   for(const [area, arrA] of byArea){
    const byPuesto = groupBy(arrA, r => r.Puesto_de_Trabajo || "(Sin puesto)");
    // resumen de estados por Área
    const summary = summarizeStatuses(arrA);
    html += `
      <div class="area-block">
       <div class="area-head d-flex justify-content-between align-items-center">
        <h6 class="mb-1">${escapeHtml(area)}</h6>
        <div class="d-flex gap-1">
          ${statusPill("vig", summary.vig)}
          ${statusPill("prox", summary.prox)}
          ${statusPill("venc", summary.venc)}
        </div>
       </div>
       <div class="list-group list-group-flush">
        ${Array.from(byPuesto).map(([puesto, arrP]) => puestoItem(area, puesto, arrP)).join("")}
       </div>
      </div>
    `;
   }
   host.innerHTML = topHtml + html;

   // Hook para toggles (no es necesario si usamos data-bs-target, pero dejamos por accesibilidad)
   host.addEventListener("click", (ev) => {
    const btn = ev.target.closest("[data-bs-toggle='collapse']");
    if(!btn) return;
    const sel = btn.getAttribute("data-bs-target");
    if(!sel) return;
    const col = document.querySelector(sel);
    if(!col) return;
    const inst = bootstrap.Collapse.getOrCreateInstance(col);
    inst.toggle();
   }, { once: true });
  }

  function groupBy(arr, keyFn){
   const map = new Map();
   for(const it of arr){
    const k = keyFn(it);
    if(!map.has(k)) map.set(k, []);
    map.get(k).push(it);
   }
   return map;
  }

  function summarizeStatuses(list){
   let vig=0, prox=0, venc=0;
   for(const r of list){
    const st = computeStatus(r.Fecha).key;
    if(st==="vig") vig++; else if(st==="prox") prox++; else if(st==="venc") venc++;
   }
   return {vig, prox, venc};
  }

  function statusPill(type, n){
   const map = {
    vig: {cls:"pill-ok", label:"Vigente"},
    prox:{cls:"pill-warn", label:"Próx."},
    venc:{cls:"pill-bad", label:"Vencido"}
   };
   const m = map[type] || {cls:"pill-na", label:"N/A"};
   return `<span class="pill ${m.cls}" title="${m.label}">${n}</span>`;
  }

  function puestoItem(area, puesto, arr){
   const id = "col_"+ hashId(area+"__"+puesto);
   const persons = arr.map(p => {
    const st = computeStatus(p.Fecha);
    const dias = (() => {
      const dt = parseDMY(p.Fecha); return dt ? daysAgo(dt) : null;
    })();
    return `
      <div class="person-item">
       <div class="person-main">
        <div class="fw-semibold">${escapeHtml(p.Nombre_Trabajador || "-")}</div>
        <div class="small text-muted">${escapeHtml(p.Sexo||"-")} · ${escapeHtml(p.Puesto_de_Trabajo||"-")}</div>
       </div>
       <div class="person-meta text-end">
        <div class="small text-muted">${escapeHtml(p.Fecha||"-")}${dias!=null?` · ${dias}d`:""}</div>
        <span class="pill ${st.cls}">${st.label}</span>
       </div>
      </div>
    `;
   }).join("");

   return `
    <button class="list-group-item list-group-item-action d-flex justify-content-between align-items-center"
          type="button" data-bs-toggle="collapse" data-bs-target="#${id}" aria-expanded="false">
      <span class="text-truncate"><i class="bi bi-person-vcard"></i> ${escapeHtml(puesto)}</span>
      <span class="badge rounded-pill bg-secondary">${arr.length}</span>
    </button>
    <div id="${id}" class="collapse">
      <div class="p-2">
       ${persons}
      </div>
    </div>
   `;
  }

  function hashId(s){
   let h = 0; for(let i=0;i<s.length;i++){ h=((h<<5)-h)+s.charCodeAt(i); h|=0; }
   return "x"+Math.abs(h);
  }
})();
