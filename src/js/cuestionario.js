// cuestionario.js
const ZONAS = ["Cuello","Hombro Derecho","Hombro Izquierdo","Codo/antebrazo Derecho","Codo/antebrazo Izquierdo",
    "Muñeca/mano Derecha","Muñeca/mano Izquierda","Espalda Alta","Espalda Baja",
    "Caderas/nalgas/muslos","Rodillas (una o ambas)","Pies/tobillos (uno o ambos)"];

document.addEventListener("DOMContentLoaded", async () => {
    const id = getParam("id");
    const data = await loadExcel();
    const row = data.find(r => String(r.ID) === String(id));
    renderRow(row || null);
});

function getParam(k){
    const u = new URL(location.href);
    return u.searchParams.get(k);
}

async function loadExcel(){
    const res = await fetch(window.RESP_XLSX);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets["Respuestas"] || wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function renderRow(r){
    const root = document.getElementById("detailRoot");
    if(!r){
        root.innerHTML = `<div class="alert alert-warning">No se encontró el cuestionario.</div>`;
        return;
    }

    const header = `
        <div class="card border-0 shadow-sm mb-3">
        <div class="card-body">
            <div class="d-flex justify-content-between align-items-center">
            <h5 class="mb-0">${escapeHtml(r.Nombre_Trabajador)} <span class="badge bg-secondary">${escapeHtml(r.Sexo)}</span></h5>
            <span class="text-muted"><i class="bi bi-calendar3"></i> ${escapeHtml(r.Fecha||"")}</span>
            </div>
            <div class="grid-2 mt-3 small">
            <div><strong>Área:</strong> ${escapeHtml(r.Area||"-")}</div>
            <div><strong>Puesto:</strong> ${escapeHtml(r.Puesto_de_Trabajo||"-")}</div>
            <div><strong>Edad:</strong> ${escapeHtml(r.Edad||"-")}</div>
            <div><strong>Diestro/Zurdo:</strong> ${escapeHtml(r.Diestro_Zurdo||"-")}</div>
            <div><strong>Tiempo en el trabajo (meses):</strong> ${escapeHtml(r.Tiempo_en_trabajo_meses||"-")}</div>
            <div><strong>Temporadas previas:</strong> ${escapeHtml(r.Temporadas_previas||"-")} ${r.N_temporadas?`(${escapeHtml(r.N_temporadas)})`:""}</div>
            <div><strong>Actividad previa:</strong> ${escapeHtml(r.Actividad_previa||"-")}</div>
            <div><strong>Otra actividad:</strong> ${escapeHtml(r.Otra_actividad||"-")} ${r.Otra_actividad_cual?`(${escapeHtml(r.Otra_actividad_cual)})`:""}</div>
            </div>
        </div>
        </div>
    `;

    const rows = ZONAS.map(z => {
        const v12 = r[`${z}__12m`] || "";
        const incap = r[`${z}__Incap`] || "";
        const d12 = r[`${z}__Dolor12m`] || "";
        const v7 = r[`${z}__7d`] || "";
        const d7 = r[`${z}__Dolor7d`] || "";
        const badge = v12 === "SI" ? "badge-yes" : (v12 === "NO" ? "badge-no" : "badge-na");
        return `<tr>
            <th>${escapeHtml(z)}</th>
            <td><span class="badge ${badge}">${v12||"N/A"}</span></td>
            <td>${escapeHtml(incap||"")}</td>
            <td>${escapeHtml(d12||"")}</td>
            <td>${escapeHtml(v7||"")}</td>
            <td>${escapeHtml(d7||"")}</td>
            </tr>`;
    }).join("");

    const table = `
        <div class="card border-0 shadow-sm">
        <div class="card-body">
            <div class="d-flex align-items-center gap-2 mb-2">
            <i class="bi bi-clipboard2-pulse"></i><h6 class="mb-0">Cuestionario Nórdico de Síntomas Musculoesqueléticos</h6>
            </div>
            <div class="table-responsive">
            <table class="table zone-table">
                <thead class="table-light">
                <tr>
                    <th>Zona</th>
                    <th>Molestias 12 meses</th>
                    <th>Incapacidad (12 m)</th>
                    <th>Escala Dolor (12 m)</th>
                    <th>Molestias últimos 7 días</th>
                    <th>Escala Dolor (7 d)</th>
                </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
            </div>
            <div class="small text-muted">
            Hallazgo positivo: dolor &gt; 3 en últimos 7 días (según protocolo).
            </div>
        </div>
        </div>
    `;

    root.innerHTML = header + table;
}

function escapeHtml(str){
    return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}
