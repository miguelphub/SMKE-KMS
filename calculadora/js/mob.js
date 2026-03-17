function formatNumber(value, digits = 2) {
  return Number(value || 0).toLocaleString("en-US", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function calcularMobilon() {
  const piezas = Number(document.getElementById("piezas").value);
  const consumoCm = Number(document.getElementById("consumo").value);
  const resultado = document.getElementById("resultado");
  const totalYardasInput = document.getElementById("totalYardas");

  if (!Number.isFinite(piezas) || !Number.isFinite(consumoCm) || piezas <= 0 || consumoCm <= 0) {
    totalYardasInput.value = "";
    resultado.textContent = "Completa ambos campos con valores mayores a cero.";
    return;
  }

  const totalCm = piezas * consumoCm;
  const yardas = totalCm / 91.44;
  const yardasRedondeadas = Math.round(yardas * 100) / 100;
  const bolsasMobilon = Math.ceil(yardasRedondeadas / 1000) * 1000;

  totalYardasInput.value = yardasRedondeadas;

  resultado.textContent =
    `Total estimado: ${formatNumber(yardasRedondeadas)} yardas. ` +
    `Bolsas de mobilon requeridas: ${bolsasMobilon.toLocaleString("en-US")}.`;
}

document.addEventListener("DOMContentLoaded", () => {
  ["piezas", "consumo"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", calcularMobilon);
  });
});
