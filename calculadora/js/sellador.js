function formatNumber(value, options = {}) {
  return Number(value).toLocaleString("es-GT", options);
}

function getModoCalculo() {
  return document.getElementById("modoCalculo")?.value || "empaque";
}

function calcularDesdeEmpaque() {
  const piezas = Number(document.getElementById("piezas").value);
  const porCaja = Number(document.getElementById("porCaja").value);
  const largoPulg = Number(document.getElementById("largo").value);
  const anchoPulg = Number(document.getElementById("ancho").value);
  const forma = document.getElementById("formaEmpaque").value;

  if (!Number.isFinite(piezas) || !Number.isFinite(porCaja) || !Number.isFinite(largoPulg) || !Number.isFinite(anchoPulg) ||
      piezas <= 0 || porCaja <= 0 || largoPulg <= 0 || anchoPulg <= 0) {
    return { error: "Completa todos los campos de empaque con números válidos mayores a cero." };
  }

  const cajas = Math.ceil(piezas / porCaja);
  const largoCm = (largoPulg * 2.54) + 4;
  const anchoCm = (anchoPulg * 2.54) + 4;

  let consumo = 0;
  if (forma === "Normal") {
    consumo = largoCm * 2;
  } else {
    consumo = (largoCm * 2) + (anchoCm * 4);
  }

  const consumoMetros = consumo / 100;
  const totalMetros = consumoMetros * cajas;
  const selladoresNecesarios = Math.ceil(totalMetros / 100);

  return {
    modo: "Datos por empaque",
    forma,
    cajas,
    consumoMetros,
    totalMetros,
    selladoresNecesarios,
    piezas,
    porCaja,
    largoPulg,
    anchoPulg
  };
}

function calcularDesdeConsumo() {
  const cajas = Number(document.getElementById("cajasManual").value);
  const consumoMetros = Number(document.getElementById("consumoManual").value);
  const forma = document.getElementById("formaConsumo").value;

  if (!Number.isFinite(cajas) || !Number.isFinite(consumoMetros) || cajas <= 0 || consumoMetros <= 0) {
    return { error: "Completa la cantidad de cajas y el consumo con números válidos mayores a cero." };
  }

  const totalMetros = consumoMetros * cajas;
  const selladoresNecesarios = Math.ceil(totalMetros / 100);

  return {
    modo: "Ya tengo consumo",
    forma,
    cajas,
    consumoMetros,
    totalMetros,
    selladoresNecesarios
  };
}

function metricCard(label, value, caption = "") {
  return `
    <article class="result-metric-card${label === "Selladores requeridos" ? " result-metric-card-accent" : ""}">
      <span>${label}</span>
      <strong>${value}</strong>
      <small>${caption}</small>
    </article>
  `;
}

function detailRow(label, value) {
  return `
    <div class="result-detail-row">
      <span>${label}</span>
      <strong>${value}</strong>
    </div>
  `;
}

function renderPlaceholder(message) {
  return `
    <article class="result-placeholder-card">
      <strong>Vista previa del cálculo</strong>
      <p>${message}</p>
    </article>
  `;
}

function renderResultado(data) {
  const resultado = document.getElementById("resultado");
  if (!resultado) return;

  if (data.error) {
    resultado.innerHTML = renderPlaceholder(data.error);
    return;
  }

  const consumoCaja = formatNumber(data.consumoMetros, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const totalMetros = formatNumber(data.totalMetros, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const cajas = formatNumber(data.cajas);
  const selladores = formatNumber(data.selladoresNecesarios);

  let detailRows = '';
  if (data.modo === "Datos por empaque") {
    detailRows += detailRow("Modo de captura", data.modo);
    detailRows += detailRow("Sellado de caja", data.forma);
    detailRows += detailRow("Cantidad de pcs", formatNumber(data.piezas));
    detailRows += detailRow("Pcs por caja", formatNumber(data.porCaja));
    detailRows += detailRow("Largo × Ancho", `${formatNumber(data.largoPulg, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}" × ${formatNumber(data.anchoPulg, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}"`);
  } else {
    detailRows += detailRow("Modo de captura", data.modo);
    detailRows += detailRow("Sellado de caja", data.forma);
    detailRows += detailRow("Consumo ingresado", `${consumoCaja} m por caja`);
    detailRows += detailRow("Cantidad de cajas", cajas);
  }

  resultado.innerHTML = `
    <div class="result-summary-grid">
      ${metricCard("Total de cajas", cajas, "cajas")}
      ${metricCard("Consumo por caja", consumoCaja, "metros")}
      ${metricCard("Metros totales", totalMetros, "metros")}
      ${metricCard("Selladores requeridos", selladores, "rollos aprox.")}
    </div>

    <div class="result-detail-grid">
      <article class="result-detail-card">
        <strong>Resumen del cálculo</strong>
        <p>El cálculo se actualiza automáticamente mientras completas los datos. No necesitas presionar ningún botón.</p>
        <div class="result-detail-meta">
          ${detailRows}
        </div>
      </article>
    </div>
  `;
}

function toggleModoSellador() {
  const modo = getModoCalculo();
  const empaque = document.getElementById("modoEmpaque");
  const consumo = document.getElementById("modoConsumo");
  const cajasCalculadas = document.getElementById("cajasCalculadas");
  const resultado = document.getElementById("resultado");

  empaque.classList.toggle("hidden", modo !== "empaque");
  consumo.classList.toggle("hidden", modo !== "consumo");

  if (modo === "empaque") {
    document.getElementById("cajasManual").value = "";
    document.getElementById("consumoManual").value = "";
  } else {
    document.getElementById("piezas").value = "";
    document.getElementById("porCaja").value = "";
    document.getElementById("largo").value = "";
    document.getElementById("ancho").value = "";
    cajasCalculadas.value = "";
  }

  resultado.innerHTML = "";
  calcularSelladores();
}

function calcularSelladores() {
  const modo = getModoCalculo();
  const cajasCalculadas = document.getElementById("cajasCalculadas");
  const data = modo === "consumo" ? calcularDesdeConsumo() : calcularDesdeEmpaque();

  if (modo === "empaque") {
    cajasCalculadas.value = data.error ? "" : data.cajas;
  }

  renderResultado(data);
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("modoCalculo")?.addEventListener("change", toggleModoSellador);

  ["piezas", "porCaja", "largo", "ancho"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", calcularSelladores);
  });
  document.getElementById("formaEmpaque")?.addEventListener("change", calcularSelladores);

  ["cajasManual", "consumoManual"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", calcularSelladores);
  });
  document.getElementById("formaConsumo")?.addEventListener("change", calcularSelladores);

  toggleModoSellador();
});
