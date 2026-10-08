// --- MÓDULO INDEPENDIENTE PARA CANAL DE AGENCIAS ---

function cargarAgenciasDesdeWorkbook(workbook) {
  if (typeof leerHojaExcel !== "function") {
    console.warn("La función 'leerHojaExcel' no está disponible en app.js");
    return;
  }

  // Lee la pestaña "Agencias" del Excel usando "kpi" como clave
  const actividadAgencias = leerHojaExcel(workbook, "Agencias", "kpi");

  if (actividadAgencias && actividadAgencias.length > 0) {
    renderizarActividadAgencias(actividadAgencias);
  } else {
    console.warn("No se obtuvieron filas de la hoja 'Agencias'.");
  }
}

function renderizarActividadAgencias(datos) {
  // 1. Búsqueda directa del tbody en la tarjeta de agencias
  let tbody = document.querySelector("#tabla-agencias tbody");

  // Si no le habías puesto id a la tabla, buscamos por el encabezado de la tarjeta
  if (!tbody) {
    const cards = document.querySelectorAll(".card");
    for (const card of cards) {
      const h2 = card.querySelector("h2");
      if (h2 && h2.textContent.toLowerCase().includes("actividad")) {
        tbody = card.querySelector("tbody");
        break;
      }
    }
  }

  if (!tbody) {
    console.error("No se encontró la tabla de Agencias en el HTML.");
    return;
  }

  if (!datos || datos.length === 0) return;

  tbody.innerHTML = "";

  datos.forEach((row) => {
    // Obtiene el nombre del KPI
    const kpi = row.kpi || Object.values(row)[0] || "";
    if (!kpi) return;

    // Helper para limpiar nulos, undefined o NaN
    const val = (campo) => {
      const v = row[campo];
      if (
        v === undefined ||
        v === null ||
        v === "" ||
        String(v).toLowerCase() === "nan"
      ) {
        return "—";
      }
      return v;
    };

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${kpi}</strong></td>
      <td class="text-center">${val("enero")}</td>
      <td class="text-center">${val("febrero")}</td>
      <td class="text-center">${val("marzo")}</td>
      <td class="text-center">${val("abril")}</td>
      <td class="text-center">${val("mayo")}</td>
      <td class="text-center">${val("junio")}</td>
      <td class="text-center">${val("julio")}</td>
      <td class="text-center">${val("agosto")}</td>
      <td class="text-center">${val("septiembre")}</td>
      <td class="text-center">${val("octubre")}</td>
      <td class="text-center">${val("noviembre")}</td>
      <td class="text-center">${val("diciembre")}</td>
    `;
    tbody.appendChild(tr);
  });
}
