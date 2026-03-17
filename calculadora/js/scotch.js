function calcularRollos() {
  const bolsas = Number(document.getElementById("bolsas").value);
  const cintaPorBolsa = Number(document.getElementById("cintaPorBolsa").value);
  const largoRollo = 5000;
  const resultado = document.getElementById("resultado");

  if (!Number.isFinite(bolsas) || !Number.isFinite(cintaPorBolsa) || bolsas <= 0 || cintaPorBolsa <= 0) {
    resultado.textContent = "Agrega valores válidos mayores a cero.";
    return;
  }

  const totalCinta = bolsas * cintaPorBolsa;
  const rollosNecesarios = Math.ceil(totalCinta / largoRollo);

  resultado.textContent =
    `Se necesitan ${totalCinta.toLocaleString("es-GT")} cm de scotch. ` +
    `Rollos requeridos: ${rollosNecesarios.toLocaleString("es-GT")}.`;
}

document.addEventListener("DOMContentLoaded", () => {
  ["bolsas", "cintaPorBolsa"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", calcularRollos);
  });
});
