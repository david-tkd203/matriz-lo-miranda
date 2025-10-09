// cuestionarios_respondidos.js
document.addEventListener("DOMContentLoaded", () => {
    loadExcel();
    document.getElementById("filterArea").addEventListener("change", render);
    document.getElementById("filterName").addEventListener("input", render);
    });

    let ROWS = [];
    let AREAS = [];

    function getParam(k){
    const u = new URL(location.href);
    return u.searchParams.get(k);
    }

    async function loadExcel(){
    try{
        const res = await fetch(window.RESP_XLSX);
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets["Respuestas"] || wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        ROWS = json;
        AREAS = [...new Set(ROWS.map(r => r.Area).filter(Boolean))].sort((a,b)=> a.localeCompare(b));

        const sel = document.getElementById("filterArea");
        sel.innerHTML = `<option value="">(Todas)</option>` + AREAS.map(a => `<option>${escapeHtml(a)}</option>`).join("");

        // Si viene ?area= en la URL, aplicarlo
        const qArea = getParam("area");
        if(qArea && AREAS.includes(qArea)){
        sel.value = qArea;
        }

        render();
    }catch(e){
        console.error("Error cargando Excel", e);
    }
    }

    function render(){
    const area = document.getElementById("filterArea").value || "";
    const name = (document.getElementById("filterName").value || "").toLowerCase().trim();
    const tbody = document.getElementById("tblBody");

    const data = ROWS.filter(r => {
        if(area && r.Area !== area) return false;
        if(name && !String(r.Nombre_Trabajador||"").toLowerCase().includes(name)) return false;
        return true;
    });

    tbody.innerHTML = data.map(r => `
        <tr>
        <td>${r.ID}</td>
        <td>${escapeHtml(r.Fecha || "")}</td>
        <td>${escapeHtml(r.Area || "")}</td>
        <td>${escapeHtml(r.Nombre_Trabajador || "")}</td>
        <td>${escapeHtml(r.Puesto_de_Trabajo || "")}</td>
        <td>${escapeHtml(r.Sexo || "")}</td>
        <td>
            <a class="btn btn-sm btn-primary" href="cuestionario.html?id=${encodeURIComponent(r.ID)}">
            <i class="bi bi-eye"></i> Ver
            </a>
        </td>
        </tr>
    `).join("");
}

function escapeHtml(str){
    return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}
