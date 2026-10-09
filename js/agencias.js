// Renderizado del apartado "Actividad y conversión del canal de agencias".
(function () {
  "use strict";

  function normalizarClaveAgencias(clave) {
    return String(clave ?? "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[\s_-]+/g, "");
  }

  function buscarTablaAgencias() {
    let tabla = document.querySelector("#tabla-agencias");
    if (tabla) return tabla;

    const cards = Array.from(document.querySelectorAll(".card"));
    const card = cards.find((elemento) =>
      (elemento.querySelector("h2")?.textContent || "")
        .toLowerCase()
        .includes("actividad y conversión del canal de agencias"),
    );

    return card?.querySelector("table") || null;
  }

  function renderizarActividadAgencias(datos) {
    const tabla = buscarTablaAgencias();
    const tbody = tabla?.querySelector("tbody");

    if (!tbody) {
      console.error(
        '[Agencias] No se encontró la tabla. Comprueba que index.html contiene <table id="tabla-agencias">.',
      );
      return;
    }

    if (!Array.isArray(datos) || datos.length === 0) {
      console.warn('[Agencias] La hoja "Agencias" no contiene filas de datos.');
      tbody.innerHTML =
        '<tr><td colspan="13">No hay datos de agencias en el Excel.</td></tr>';
      return;
    }

    tbody.innerHTML = "";

    const meses = [
      ["enero"],
      ["febrero"],
      ["marzo"],
      ["abril"],
      ["mayo"],
      ["junio"],
      ["julio"],
      ["agosto"],
      ["septiembre", "setiembre"],
      ["octubre"],
      ["noviembre"],
      ["diciembre"],
    ];

    let filasPintadas = 0;

    datos.forEach((fila) => {
      const normalizada = {};

      Object.entries(fila || {}).forEach(([clave, valor]) => {
        normalizada[normalizarClaveAgencias(clave)] = valor;
      });

      const kpi = normalizada.kpi ?? Object.values(normalizada)[0] ?? "";

      if (String(kpi).trim() === "") return;

      const tr = document.createElement("tr");
      const tdKpi = document.createElement("td");
      const strong = document.createElement("strong");

      strong.textContent = String(kpi);
      tdKpi.appendChild(strong);
      tr.appendChild(tdKpi);

      meses.forEach((variantes) => {
        let valor;

        for (const mes of variantes) {
          const clave = normalizarClaveAgencias(mes);

          if (
            normalizada[clave] !== undefined &&
            normalizada[clave] !== null &&
            String(normalizada[clave]).trim() !== ""
          ) {
            valor = normalizada[clave];
            break;
          }
        }

        const td = document.createElement("td");
        td.className = "text-center";
        td.textContent = valor === undefined ? "—" : String(valor);

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
      filasPintadas++;
    });

    if (filasPintadas === 0) {
      tbody.innerHTML =
        '<tr><td colspan="13">No se han encontrado KPI. Revisa la primera columna de la hoja Agencias.</td></tr>';
    }

    console.info(`[Agencias] Tabla renderizada: ${filasPintadas} KPI.`);
  }

  function cargarAgenciasDesdeWorkbook(workbook) {
    if (!workbook || !workbook.Sheets) {
      console.error("[Agencias] No se ha recibido un libro Excel válido.");
      return;
    }

    const nombreHoja = workbook.SheetNames?.find(
      (nombre) => normalizarClaveAgencias(nombre) === "agencias",
    );

    const hoja = nombreHoja ? workbook.Sheets[nombreHoja] : null;

    if (!hoja) {
      console.error(
        '[Agencias] No existe una hoja llamada "Agencias". Hojas encontradas:',
        workbook.SheetNames || [],
      );

      const tbody = buscarTablaAgencias()?.querySelector("tbody");

      if (tbody) {
        tbody.innerHTML =
          '<tr><td colspan="13">No se encuentra la hoja Agencias en el Excel.</td></tr>';
      }

      return;
    }

    const datos = XLSX.utils.sheet_to_json(hoja, {
      defval: "",
      raw: true,
      blankrows: false,
    });

    console.info(
      `[Agencias] Leídas ${datos.length} filas desde la hoja "${nombreHoja}".`,
    );

    renderizarActividadAgencias(datos);
  }

  // Permite que app.js acceda a estas funciones.
  window.cargarAgenciasDesdeWorkbook = cargarAgenciasDesdeWorkbook;

  window.renderizarActividadAgencias = renderizarActividadAgencias;
})();
