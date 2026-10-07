// --- FUENTE DE DATOS LOCAL ---
// El Excel debe estar en la carpeta "plantillas" junto a index.html.
const EXCEL_LOCAL_URL = "plantillas/datos_mes.xlsx";

// ======================================================
// PROTECCIÓN DE ACCESO
// Contraseña real: 2026
// Guardada como hash SHA-256 para no mostrarla en claro
// ======================================================

const HASH_CORRECTO =
  "d405d092f5f0f38f92d76f1b8d2ec4c8d11f6f7d839a3c8670fbe3ca98b6d548";

function validarAcceso() {
  const pass = document.getElementById("passwordInput").value;

  if (pass === "2026") {
    document.getElementById("pantalla-login").style.display = "none";
    document.querySelector(".dashboard-container").style.display = "block";
  } else {
    document.getElementById("mensajeError").textContent =
      "Contraseña incorrecta";
  }
}

function cerrarSesion() {
  document.getElementById("pantalla-login").style.display = "block";
  document.querySelector(".dashboard-container").style.display = "none";
}

// --- DATOS POR DEFECTO (PORTFOLIO BASE) ---
const PORTFOLIO_DEFAULT = [
  {
    Promocion: "SOLARIS",
    Zona: "MARIA",
    Vendidas: 80,
    Cobradas: 80,
    TotUnid: 80,
    PteVender: 0,
    PctVenta: 1,
    EstProduccion: 0,
  },
  {
    Promocion: "ARBITANA",
    Zona: "MARIA",
    Vendidas: 24,
    Cobradas: 21,
    TotUnid: 25,
    PteVender: 1,
    PctVenta: 0.96,
    EstProduccion: 0,
  },
  {
    Promocion: "INESIA",
    Zona: "MARIA",
    Vendidas: 24,
    Cobradas: 23,
    TotUnid: 25,
    PteVender: 1,
    PctVenta: 0.96,
    EstProduccion: 0,
  },
  {
    Promocion: "PORTAMARE",
    Zona: "RUTH",
    Vendidas: 18,
    Cobradas: 18,
    TotUnid: 22,
    PteVender: 4,
    PctVenta: 0.82,
    EstProduccion: 51276,
  },
  {
    Promocion: "DOÑA RAFAELA",
    Zona: "MARIA",
    Vendidas: 5,
    Cobradas: 5,
    TotUnid: 10,
    PteVender: 5,
    PctVenta: 0.5,
    EstProduccion: 70667,
  },
  {
    Promocion: "BE GRAND",
    Zona: "MARIA",
    Vendidas: 8,
    Cobradas: 7,
    TotUnid: 13,
    PteVender: 5,
    PctVenta: 0.62,
    EstProduccion: 207250,
  },
  {
    Promocion: "LAKE ESSENCE",
    Zona: "RUTH",
    Vendidas: 6,
    Cobradas: 6,
    TotUnid: 14,
    PteVender: 8,
    PctVenta: 0.43,
    EstProduccion: 211726.5,
  },
  {
    Promocion: "ALTOASIS",
    Zona: "RUTH",
    Vendidas: 21,
    Cobradas: 10,
    TotUnid: 83,
    PteVender: 62,
    PctVenta: 0.25,
    EstProduccion: 1415909.5,
  },
  {
    Promocion: "CASARES BAY",
    Zona: "RUTH",
    Vendidas: 3,
    Cobradas: 0,
    TotUnid: 9,
    PteVender: 6,
    PctVenta: 0.33,
    EstProduccion: 92820,
  },
  {
    Promocion: "BALCON DEL MEDITERRANEO",
    Zona: "RUTH",
    Vendidas: 8,
    Cobradas: 4,
    TotUnid: 37,
    PteVender: 29,
    PctVenta: 0.22,
    EstProduccion: 1202313.76,
  },
  {
    Promocion: "VILLAS PALATINO",
    Zona: "MARIA",
    Vendidas: 0,
    Cobradas: 0,
    TotUnid: 4,
    PteVender: 4,
    PctVenta: 0,
    EstProduccion: 166625,
  },
];

// --- MODALES ---
function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) {
    modal.style.display = "flex";
    document.body.style.overflow = "hidden";
    if (id === "modal-captaciones") calcularKPIsCaptaciones();
  } else {
    alert("El informe (" + id + ") aún no está disponible.");
  }
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) {
    modal.style.display = "none";
    document.body.style.overflow = "auto";
  }
}

window.onclick = function (event) {
  if (event.target.classList.contains("modal-overlay")) {
    event.target.style.display = "none";
    document.body.style.overflow = "auto";
  }
};

// --- FORMATEADOR Y UTILIDADES ---
function formatearNombreMes(mesVal) {
  if (!mesVal) return "";
  if (typeof mesVal === "string" && !mesVal.includes("-")) return mesVal;

  const dateObj = new Date(mesVal);
  if (!isNaN(dateObj.getTime())) {
    const meses = [
      "Enero",
      "Febrero",
      "Marzo",
      "Abril",
      "Mayo",
      "Junio",
      "Julio",
      "Agosto",
      "Septiembre",
      "Octubre",
      "Noviembre",
      "Diciembre",
    ];
    return meses[dateObj.getUTCMonth()] + " " + dateObj.getUTCFullYear();
  }
  return String(mesVal);
}

// Convertidor flexible de cadenas/números en formato es-ES a números flotantes
function parseNumeroEs(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  let str = String(val)
    .replace(/[^0-9,-]/g, "")
    .replace(",", ".");
  return parseFloat(str) || 0;
}

// --- PARSER CSV ADAPTABLE (SOPORTA COMAS Y PUNTO Y COMA) ---
function parseCSVToJSON(csvText) {
  if (!csvText) return [];

  const lines = csvText.trim().split(/\r\n|\n/);
  if (lines.length < 2) return [];

  // Detectar si el separador es coma (,) o punto y coma (;)
  const primeraLinea = lines[0];
  const numComas = (primeraLinea.match(/,/g) || []).length;
  const numPuntosComa = (primeraLinea.match(/;/g) || []).length;
  const separador = numPuntosComa > numComas ? ";" : ",";

  const regexSeparador = new RegExp(`${separador}(?=(?:(?:[^"]*"){2})*[^"]*$)`);

  // Normalizar cabeceras a minúsculas y sin acentos
  const headers = primeraLinea.split(regexSeparador).map((h) =>
    h
      .replace(/^"(.*)"$/, "$1")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
  );

  const result = [];

  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const currentline = lines[i].split(regexSeparador);
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = currentline[j]
        ? currentline[j].replace(/^"(.*)"$/, "$1").trim()
        : "";
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return result;
}

// --- RENDERIZADO Y SUMA DE TOTALES DE PRODUCCIÓN ---
function renderizarProduccion(datos) {
  const tbody = document.querySelector("#tabla-produccion tbody");
  if (!tbody || !datos || datos.length === 0) return;
  tbody.innerHTML = "";

  let totalAcumulado = 0,
    totalGilmar = 0,
    totalAgencias = 0,
    totalUnidades = 0;

  datos.forEach((row) => {
    const tr = document.createElement("tr");
    const nombreMes = formatearNombreMes(row.mes || row.periodo);

    const prod = parseNumeroEs(
      row.prodtotal || row.prod_total || row.total || row.prod
    );
    const gilmar = parseNumeroEs(row.gilmar);
    const agencias = parseNumeroEs(row.agencias);
    const unidades = parseNumeroEs(row.unidades || row.ventas);

    totalAcumulado += prod;
    totalGilmar += gilmar;
    totalAgencias += agencias;
    totalUnidades += unidades;

    if (
      String(row.esseptiembre || row.esagosto || "").toUpperCase() === "TRUE"
    ) {
      tr.classList.add("highlight-julio");
    }

    tr.innerHTML = `
      <td><strong>${nombreMes}</strong></td>
      <td class="text-right">${prod.toLocaleString("es-ES", {
        style: "currency",
        currency: "EUR",
      })}</td>
      <td class="text-right">${gilmar.toLocaleString("es-ES", {
        style: "currency",
        currency: "EUR",
      })}</td>
      <td class="text-right">${agencias.toLocaleString("es-ES", {
        style: "currency",
        currency: "EUR",
      })}</td>
      <td class="text-right">${unidades.toLocaleString("es-ES")}</td>
    `;
    tbody.appendChild(tr);
  });

  // Fila de Totales
  const trTotal = document.createElement("tr");
  trTotal.style.backgroundColor = "var(--bg)";
  trTotal.style.fontWeight = "bold";
  trTotal.innerHTML = `
    <td>TOTAL ACUMULADO</td>
    <td class="text-right text-accent">${totalAcumulado.toLocaleString(
      "es-ES",
      { style: "currency", currency: "EUR" }
    )}</td>
    <td class="text-right">${totalGilmar.toLocaleString("es-ES", {
      style: "currency",
      currency: "EUR",
    })}</td>
    <td class="text-right">${totalAgencias.toLocaleString("es-ES", {
      style: "currency",
      currency: "EUR",
    })}</td>
    <td class="text-right">${totalUnidades.toLocaleString("es-ES")}</td>
  `;
  tbody.appendChild(trTotal);

  actualizarCardsSuperiores(
    totalAcumulado,
    totalGilmar,
    totalAgencias,
    totalUnidades
  );
}

function renderizarPortfolio(datos) {
  const tbody = document.querySelector("#tabla-portfolio tbody");
  if (!tbody || !datos || datos.length === 0) return;
  tbody.innerHTML = "";

  let totalVendidas = 0,
    totalCobradas = 0,
    totalPteCobrar = 0,
    totalUnidades = 0,
    totalPteVender = 0,
    totalEstProduccion = 0;

  datos.forEach((row) => {
    const tr = document.createElement("tr");

    let pct = parseNumeroEs(row.pctventa || row.pct);
    if (pct > 0 && pct <= 1) pct = Math.round(pct * 100);
    else pct = Math.round(pct);

    const vendidas = parseNumeroEs(row.vendidas);
    const cobradas = parseNumeroEs(row.cobradas);
    const pteCobrar = parseNumeroEs(
      row.ptedecobro || row.pte_cobro || row.ptecobrar
    );
    const totUnid = parseNumeroEs(
      row.totunid || row.totunidades || row.totalunidades
    );
    const pteVender = parseNumeroEs(row.ptevender || row.pte_vender);
    const estProduccion = parseNumeroEs(
      row.estproduccion || row.est_produccion || row.produccion
    );

    totalVendidas += vendidas;
    totalCobradas += cobradas;
    totalPteCobrar += pteCobrar;
    totalUnidades += totUnid;
    totalPteVender += pteVender;
    totalEstProduccion += estProduccion;

    const promocion = row.promocion || "";
    const zona = row.zona || "";

    tr.innerHTML = `
      <td><strong>${promocion}</strong></td>
      <td>${zona}</td>
      <td class="text-right">${vendidas}</td>
      <td class="text-right">${cobradas}</td>
      <td class="text-right">${pteCobrar}</td>
      <td class="text-right">${totUnid}</td>
      <td class="text-right">${pteVender}</td>
      <td>
        <div class="progress-container">
          <span>${pct}%</span>
          <div class="progress-bar">
            <div class="progress-fill ${
              pct === 100 ? "complete" : ""
            }" style="width: ${pct}%;"></div>
          </div>
        </div>
      </td>
      <td class="text-right">${
        estProduccion > 0
          ? estProduccion.toLocaleString("es-ES", {
              style: "currency",
              currency: "EUR",
            })
          : "—"
      }</td>
    `;
    tbody.appendChild(tr);
  });

  const pctGlobal =
    totalUnidades > 0 ? Math.round((totalVendidas / totalUnidades) * 100) : 0;

  const trTotal = document.createElement("tr");
  trTotal.style.backgroundColor = "var(--bg)";
  trTotal.style.fontWeight = "bold";
  trTotal.innerHTML = `
    <td>TOTAL PORTFOLIO</td>
    <td>—</td>
    <td class="text-right">${totalVendidas}</td>
    <td class="text-right">${totalCobradas}</td>
    <td class="text-right">${totalPteCobrar}</td>
    <td class="text-right">${totalUnidades}</td>
    <td class="text-right">${totalPteVender}</td>
    <td>
      <div class="progress-container">
        <span>${pctGlobal}%</span>
        <div class="progress-bar">
          <div class="progress-fill ${
            pctGlobal === 100 ? "complete" : ""
          }" style="width: ${pctGlobal}%;"></div>
        </div>
      </div>
    </td>
    <td class="text-right text-accent">${totalEstProduccion.toLocaleString(
      "es-ES",
      { style: "currency", currency: "EUR" }
    )}</td>
  `;
  tbody.appendChild(trTotal);

  const elemHeaderEuros = document.getElementById("header-total-euros");
  const elemHeaderUnidades = document.getElementById("header-total-unidades");
  if (elemHeaderEuros)
    elemHeaderEuros.textContent = totalEstProduccion.toLocaleString("es-ES", {
      style: "currency",
      currency: "EUR",
    });
  if (elemHeaderUnidades) elemHeaderUnidades.textContent = totalPteVender;
}
function renderizarPteCobro(datos) {
  const tbody = document.querySelector("#tabla-pte-cobro tbody");
  if (!tbody || !datos || datos.length === 0) return;
  tbody.innerHTML = "";

  let totalUnidadesPte = 0;
  let totalProduccionPte = 0;

  datos.forEach((row) => {
    // Procura a promoção em qualquer variação de nome de chave
    const promocion =
      row.promocion || row.promocion || Object.values(row)[0] || "";
    if (!promocion) return;

    // Obtém as unidades cobradas pendentes
    const unidades = parseNumeroEs(
      row.ptedecobrounidades ||
        row["ptedecobro,unidades"] ||
        row.ptedecobro ||
        row.unidades
    );

    // Obtém o valor em euros da produção pendente
    const prodPte = parseNumeroEs(
      row.cantidadenproduccionpte ||
        row["cantidadenproduccionpte,€"] ||
        row.produccionpte ||
        row.cantidadenproduccionpte
    );

    totalUnidadesPte += unidades;
    totalProduccionPte += prodPte;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${promocion}</strong></td>
      <td class="text-right">${unidades}</td>
      <td class="text-right">${prodPte.toLocaleString("es-ES", {
        style: "currency",
        currency: "EUR",
      })}</td>
    `;
    tbody.appendChild(tr);
  });

  // Linha de Totais da tabela
  const trTotal = document.createElement("tr");
  trTotal.style.backgroundColor = "var(--bg)";
  trTotal.style.fontWeight = "bold";
  trTotal.innerHTML = `
    <td>TOTAL PENDIENTE DE COBRO</td>
    <td class="text-right">${totalUnidadesPte}</td>
    <td class="text-right text-accent">${totalProduccionPte.toLocaleString(
      "es-ES",
      { style: "currency", currency: "EUR" }
    )}</td>
  `;
  tbody.appendChild(trTotal);
}


function actualizarCardsSuperiores(
  totalAcumulado,
  totalGilmar,
  totalAgencias,
  totalUnidades
) {
  const elemProdAcum = document.querySelector(".kpi-card:nth-child(1) .value");
  const elemUnidades = document.querySelector(
    ".kpi-card:nth-child(1) .subtext"
  );

  if (elemProdAcum) {
    elemProdAcum.textContent = totalAcumulado.toLocaleString("es-ES", {
      style: "currency",
      currency: "EUR",
    });
  }
  if (elemUnidades) {
    elemUnidades.innerHTML = `Equivalente a <strong>${totalUnidades.toLocaleString(
      "es-ES"
    )} ventas unidades</strong>`;
  }

  const elemGilmarKpi = document.querySelector(".kpi-card:nth-child(2) .value");
  const elemDesglose = document.querySelector(
    ".kpi-card:nth-child(2) .subtext"
  );

  if (elemGilmarKpi && totalAcumulado > 0) {
    // Destacamos la cantidad de Gilmar en el valor principal de la tarjeta
    elemGilmarKpi.textContent = totalGilmar.toLocaleString("es-ES", {
      style: "currency",
      currency: "EUR",
    });
  }

  if (elemDesglose && totalAcumulado > 0) {
    const pctGilmar = ((totalGilmar / totalAcumulado) * 100)
      .toFixed(1)
      .replace(".", ",");
    const pctAgencias = ((totalAgencias / totalAcumulado) * 100)
      .toFixed(1)
      .replace(".", ",");

    // Mostramos porcentajes y en minúsculas la aportación de agencias
    elemDesglose.innerHTML = `<strong>Gilmar (${pctGilmar}%)</strong> | <span style="font-size:0.9em; text-transform: lowercase;">${totalAgencias.toLocaleString(
      "es-ES",
      { style: "currency", currency: "EUR" }
    )} aportación agencias (${pctAgencias}%)</span>`;
  }
}

// --- RENDERIZADO Y SUMA DE TOTALES DE PORTFOLIO ---
function renderizarPortfolio(datos) {
  const tbody = document.querySelector("#tabla-portfolio tbody");
  if (!tbody || !datos || datos.length === 0) return;
  tbody.innerHTML = "";

  let totalVendidas = 0,
    totalCobradas = 0,
    totalUnidades = 0,
    totalPteVender = 0,
    totalEstProduccion = 0;

  datos.forEach((row) => {
    const tr = document.createElement("tr");

    let pct = parseNumeroEs(row.pctventa || row.pct);
    if (pct > 0 && pct <= 1) pct = Math.round(pct * 100);
    else pct = Math.round(pct);

    const vendidas = parseNumeroEs(row.vendidas);
    const cobradas = parseNumeroEs(row.cobradas);
    const totUnid = parseNumeroEs(
      row.totunid || row.totunidades || row.totalunidades
    );
    const pteVender = parseNumeroEs(row.ptevender || row.pte_vender);
    const estProduccion = parseNumeroEs(
      row.estproduccion || row.est_produccion || row.produccion
    );

    totalVendidas += vendidas;
    totalCobradas += cobradas;
    totalUnidades += totUnid;
    totalPteVender += pteVender;
    totalEstProduccion += estProduccion;

    const promocion = row.promocion || "";
    const zona = row.zona || "";

    tr.innerHTML = `
            <td><strong>${promocion}</strong></td>
            <td>${zona}</td>
            <td class="text-right">${vendidas}</td>
            <td class="text-right">${cobradas}</td>
            <td class="text-right">${totUnid}</td>
            <td class="text-right">${pteVender}</td>
            <td>
                <div class="progress-container">
                    <span>${pct}%</span>
                    <div class="progress-bar"><div class="progress-fill ${
                      pct === 100 ? "complete" : ""
                    }" style="width: ${pct}%;"></div></div>
                </div>
            </td>
            <td class="text-right">${
              estProduccion > 0
                ? estProduccion.toLocaleString("es-ES", {
                    style: "currency",
                    currency: "EUR",
                  })
                : "—"
            }</td>
        `;
    tbody.appendChild(tr);
  });

  const pctGlobal =
    totalUnidades > 0 ? Math.round((totalVendidas / totalUnidades) * 100) : 0;

  // Fila con la suma acumulada del portfolio
  const trTotal = document.createElement("tr");
  trTotal.style.backgroundColor = "var(--bg)";
  trTotal.style.fontWeight = "bold";
  trTotal.innerHTML = `
        <td>TOTAL PORTFOLIO</td>
        <td>—</td>
        <td class="text-right">${totalVendidas}</td>
        <td class="text-right">${totalCobradas}</td>
        <td class="text-right">${totalUnidades}</td>
        <td class="text-right">${totalPteVender}</td>
        <td>
            <div class="progress-container">
                <span>${pctGlobal}%</span>
                <div class="progress-bar"><div class="progress-fill ${
                  pctGlobal === 100 ? "complete" : ""
                }" style="width: ${pctGlobal}%;"></div></div>
            </div>
        </td>
        <td class="text-right text-accent">${totalEstProduccion.toLocaleString(
          "es-ES",
          { style: "currency", currency: "EUR" }
        )}</td>
    `;
  tbody.appendChild(trTotal);

  // Actualizar encabezados del Dashboard
  const elemHeaderEuros = document.getElementById("header-total-euros");
  const elemHeaderUnidades = document.getElementById("header-total-unidades");
  if (elemHeaderEuros)
    elemHeaderEuros.textContent = totalEstProduccion.toLocaleString("es-ES", {
      style: "currency",
      currency: "EUR",
    });
  if (elemHeaderUnidades) elemHeaderUnidades.textContent = totalPteVender;
}

// --- RENDERIZADO DE CAPTACIONES ---
function renderizarCaptaciones(datos) {
  const tbody = document.querySelector("#table-captaciones tbody");
  if (!tbody || !datos || datos.length === 0) return;
  tbody.innerHTML = "";

  datos.forEach((row) => {
    const tr = document.createElement("tr");
    const fase = row.fase || "";
    const esAvanzada = fase.toLowerCase().includes("avanzada");

    let linkDossier = String(row.dossierlink || row.dossier || "").trim();
    if (linkDossier.includes("](")) {
      const match = linkDossier.match(/\((https?:\/\/[^\)]+)\)/);
      if (match) linkDossier = match[1];
    }

    tr.innerHTML = `
            <td>${row.responsable || ""}</td>
            <td>${formatearNombreMes(row.periodo)}</td>
            <td>${row.empresa || ""}</td>
            <td>${row.ubicacion || ""}</td>
            <td>${row.proyecto || ""}</td>
            <td>${row.contacto || ""}</td>
            <td>${row.estado || ""}</td>
            <td>${
              esAvanzada
                ? `<span style="color: #16a34a; font-weight:bold;">${fase}</span>`
                : fase
            }</td>
            <td>${row.observaciones || ""}</td>
            <td>${
              linkDossier && linkDossier.startsWith("http")
                ? `<a href="${linkDossier}" target="_blank" class="text-accent" style="font-weight:bold;">📄 Dossier</a>`
                : "—"
            }</td>
        `;
    tbody.appendChild(tr);
  });

  calcularKPIsCaptaciones();
}

function calcularKPIsCaptaciones() {
  const tabla = document.getElementById("table-captaciones");
  if (!tabla) return;

  const filas = tabla.querySelectorAll("tbody tr");
  let total = filas.length,
    avanzadas = 0,
    cerradas = 0;

  filas.forEach((fila) => {
    const txt = fila.innerText.toLowerCase();
    if (txt.includes("cerrada sin éxito") || txt.includes("denegado"))
      cerradas++;
    if (txt.includes("avanzada")) avanzadas++;
  });

  if (document.getElementById("kpi-total-oportunidades"))
    document.getElementById("kpi-total-oportunidades").innerText = total;
  if (document.getElementById("kpi-activas-seguimiento"))
    document.getElementById("kpi-activas-seguimiento").innerText =
      total - cerradas;
  if (document.getElementById("kpi-fase-avanzada"))
    document.getElementById("kpi-fase-avanzada").innerText = avanzadas;
}

// --- CARGA DE DATOS DESDE EXCEL LOCAL ---

function normalizarClave(clave) {
  return String(clave ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\s_-]+/g, "");
}

function normalizarFilaExcel(fila) {
  const salida = {};
  Object.keys(fila || {}).forEach((clave) => {
    salida[normalizarClave(clave)] = fila[clave];
  });
  return salida;
}

function limpiarFilasExcel(filas, campoPrincipal) {
  return (filas || []).map(normalizarFilaExcel).filter((fila) => {
    const valor = fila[campoPrincipal];
    return valor !== undefined && valor !== null && String(valor).trim() !== "";
  });
}

function leerHojaExcel(workbook, nombreHoja, campoPrincipal) {
  const hoja = workbook.Sheets[nombreHoja];
  if (!hoja) {
    console.warn(`No se encuentra la hoja "${nombreHoja}" en el Excel.`);
    return [];
  }

  const filas = XLSX.utils.sheet_to_json(hoja, {
    defval: "",
    raw: true,
    blankrows: false,
  });

  return limpiarFilasExcel(filas, campoPrincipal);
}

function mostrarErrorExcel(mensaje) {
  console.error(mensaje);

  const zona = document.getElementById("zona-carga-excel");
  if (zona) {
    zona.style.display = "block";
    zona.innerHTML = `
            <div style="padding:12px; border:1px solid #fca5a5; background:#fef2f2; color:#991b1b; border-radius:8px; margin-bottom:10px;">
                <strong>No se pudo cargar el Excel automáticamente.</strong><br>
                ${mensaje}<br>
                <small>Si estás abriendo index.html directamente con doble clic, utiliza un servidor local (por ejemplo VS Code + Live Server).</small>
            </div>
            <input type="file" id="input-excel" accept=".xlsx,.xls" onchange="procesarExcel(event)">
        `;
  }
}

function cargarWorkbookEnDashboard(workbook) {
  try {
    const produccion = leerHojaExcel(workbook, "Produccion", "mes");
    const portfolio = leerHojaExcel(workbook, "Portfolio", "promocion");
    const captaciones = leerHojaExcel(workbook, "Captaciones", "responsable");
    const pteCobro = leerHojaExcel(workbook, "Pte de cobro", "promocion");

    console.log("Excel cargado correctamente:", {
      Produccion: produccion.length,
      Portfolio: portfolio.length,
      Captaciones: captaciones.length,
      PteCobro: pteCobro.length,
    });

    if (produccion.length) renderizarProduccion(produccion);
    if (portfolio.length) renderizarPortfolio(portfolio);
    if (captaciones.length) renderizarCaptaciones(captaciones);
    if (pteCobro.length) renderizarPteCobro(pteCobro);

    return true;
  } catch (error) {
    mostrarErrorExcel("Error procesando el archivo: " + error.message);
    return false;
  }
}

function procesarExcel(event) {
  const archivo = event.target.files?.[0];
  if (!archivo) return;

  const lector = new FileReader();

  lector.onload = function (e) {
    try {
      const datos = new Uint8Array(e.target.result);
      const workbook = XLSX.read(datos, { type: "array", cellDates: true });
      cargarWorkbookEnDashboard(workbook);
    } catch (error) {
      mostrarErrorExcel(
        "No se pudo leer el Excel seleccionado: " + error.message
      );
    }
  };

  lector.onerror = function () {
    mostrarErrorExcel("El navegador no pudo leer el archivo seleccionado.");
  };

  lector.readAsArrayBuffer(archivo);
}

async function cargarExcelLocal() {
  try {
    // Evita que el navegador reutilice una copia antigua del Excel.
    const url = EXCEL_LOCAL_URL + "?v=" + Date.now();

    const respuesta = await fetch(url, {
      cache: "no-store",
    });

    if (!respuesta.ok) {
      throw new Error(`HTTP ${respuesta.status} al abrir ${EXCEL_LOCAL_URL}`);
    }

    const arrayBuffer = await respuesta.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
      type: "array",
      cellDates: true,
    });

    return cargarWorkbookEnDashboard(workbook);
  } catch (error) {
    mostrarErrorExcel(
      'No se ha podido abrir "plantillas/datos_mes.xlsx". ' +
        "Comprueba que el archivo esté exactamente en la carpeta plantillas. " +
        "Detalle: " +
        error.message
    );
    return false;
  }
}

document.addEventListener("DOMContentLoaded", async function () {
  const dashboard = document.querySelector(".dashboard-container");

  if (sessionStorage.getItem("acceso_autorizado") === "true") {
    document.getElementById("pantalla-login").style.display = "none";

    if (dashboard) {
      dashboard.style.display = "block";
    }
  } else {
    if (dashboard) {
      dashboard.style.display = "none";
    }
  }

  renderizarPortfolio(PORTFOLIO_DEFAULT);

  await cargarExcelLocal();
});
