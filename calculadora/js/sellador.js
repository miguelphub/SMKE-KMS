function calcularSelladores() {
  const piezas = Number(document.getElementById("piezas").value);
  const porCaja = Number(document.getElementById("porCaja").value);
  const largoPulg = Number(document.getElementById("largo").value);
  const anchoPulg = Number(document.getElementById("ancho").value);
  const forma = document.getElementById("forma").value;
  const resultado = document.getElementById("resultado");
  const cajasInput = document.getElementById("cajas");

  if (!Number.isFinite(piezas) || !Number.isFinite(porCaja) || !Number.isFinite(largoPulg) || !Number.isFinite(anchoPulg) ||
      piezas <= 0 || porCaja <= 0 || largoPulg <= 0 || anchoPulg <= 0) {
    cajasInput.value = "";
    resultado.textContent = "Completa todos los campos con números válidos mayores a cero.";
    return;
  }

  const cajas = Math.ceil(piezas / porCaja);
  cajasInput.value = cajas;

  const largoCm = (largoPulg * 2.54) + 4;
  const anchoCm = (anchoPulg * 2.54) + 4;

  let consumo = 0;
  if (forma === "Normal") {
    consumo = largoCm * 2;
  } else if (forma === "H") {
    consumo = (largoCm * 2) + (anchoCm * 4);
  }

  const consumoMetros = Math.round((consumo / 100) * 100) / 100;
  const totalMetros = consumoMetros * cajas;
  const selladoresNecesarios = Math.ceil(totalMetros / 100);

  resultado.textContent =
    `Cajas: ${cajas.toLocaleString("es-GT")}. ` +
    `Consumo por caja: ${consumoMetros.toLocaleString("es-GT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} m. ` +
    `Selladores requeridos: ${selladoresNecesarios.toLocaleString("es-GT")}.`;
}

document.addEventListener("DOMContentLoaded", () => {
  ["piezas", "porCaja", "largo", "ancho"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", calcularSelladores);
  });
  document.getElementById("forma")?.addEventListener("change", calcularSelladores);
});
