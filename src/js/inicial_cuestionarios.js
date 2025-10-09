// inicial_cuestionarios.js
// Panel lateral en inicial.html con "Cuestionarios por Área"

document.addEventListener("DOMContentLoaded", async () => {
    await renderCuestionariosPorArea();
});

async function renderCuestionariosPorArea(){
    const wrap = document.getElementById("areasList");
    const empty = document.getElementById("areasEmpty");
    if(!wrap) return;

    try{
        const res = await fetch(window.RESP_XLSX);
        if(!res.ok) throw new Error("No se encontró el Excel de cuestionarios");
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type:"array" });
        const ws = wb.Sheets["Respuestas"] || wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });

        if(!rows.length){
        empty.classList.remove("d-none");
        wrap.innerHTML = "";
        return;
        }

        const byArea = {};
        for(const r of rows){
        const area = r.Area || "Sin área";
        if(!byArea[area]) byArea[area] = { total:0, H:0, M:0 };
        byArea[area].total += 1;
        if(String(r.Sexo).toLowerCase().startsWith("h")) byArea[area].H += 1;
        if(String(r.Sexo).toLowerCase().startsWith("m")) byArea[area].M += 1;
        }

        const areas = Object.keys(byArea).sort((a,b)=> byArea[b].total - byArea[a].total);
        const html = [
        '<div class="list-group list-group-flush">',
        ...areas.map(a => {
            const info = byArea[a];
            const link = `cuestionarios_respondidos.html?area=${encodeURIComponent(a)}`;
            return `
            <a class="list-group-item list-group-item-action d-flex justify-content-between align-items-center" href="${link}">
                <div>
                <div class="fw-semibold">${escapeHtml(a)}</div>
                <div class="small text-muted">H ${info.H} · M ${info.M}</div>
                </div>
                <span class="badge bg-primary rounded-pill">${info.total}</span>
            </a>
            `;
        }),
        '</div>'
        ].join("");

        wrap.innerHTML = html;
        empty.classList.add("d-none");

    }catch(e){
        console.warn("Cuestionarios por área: ", e.message);
        empty.classList.remove("d-none");
        if(wrap) wrap.innerHTML = "";
    }
}

function escapeHtml(str){
    return String(str).replace(/[&<>"']/g, s => ({
        '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
    })[s]);
}
