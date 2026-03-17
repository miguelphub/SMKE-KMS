const el = (id) => document.getElementById(id);

const MOBILON_TEMPLATE_PATH = "plantillas/PLANTILLA-MOB.xlsx";
const PROVIDER_PRICES = {
  "SUMINISTROS ARYEL": 0.0065,
  JCA: 0.009
};

const mobilonState = {
  imageBase64: null,
  imageExtension: null
};

function todayYMD() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addDays(ymd, days) {
  if (!ymd) return "";
  const [y, m, d] = ymd.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + days);
  const yy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

function formatDateExcel(ymd) {
  if (!ymd) return "";
  const [y, m, d] = ymd.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function formatCurrency(value, digits = 4) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "$0.0000";
  return num.toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function formatCurrencyAmount(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "$0.00";
  return num.toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatInt(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "0";
  return num.toLocaleString("en-US", { maximumFractionDigits: 0 });
}

function parseNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function poDigitsOnly(raw) {
  const s = String(raw || "").trim();
  const digits = s.replace(/\D+/g, "");
  return digits || s;
}

function safeFileName(name) {
  return String(name || "archivo").replace(/[\/:*?"<>|]+/g, "_");
}

function setMessage(text = "", type = "") {
  const msg = el("message");
  msg.textContent = text;
  msg.className = `message${type ? ` ${type}` : ""}`;
}

function forceDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function setWorkbookMeta(workbook) {
  workbook.creator = "Miguelnmms";
  workbook.lastModifiedBy = "Miguelnmms";
  workbook.created = new Date();
  workbook.modified = new Date();
}

async function blobToDataURL(blob) {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function setTodayDefaults() {
  const today = todayYMD();
  el("fechaPo").value = today;
  el("fechaProd").value = addDays(today, 7);
}

function roundToNextThousand(value) {
  const num = Math.max(0, Math.ceil(parseNumber(value)));
  if (!num) return 0;
  return Math.ceil(num / 1000) * 1000;
}

function getPrice(provider) {
  return PROVIDER_PRICES[provider] || PROVIDER_PRICES["SUMINISTROS ARYEL"];
}

function getFormData() {
  const provider = el("proveedor").value || "SUMINISTROS ARYEL";
  const requestedQty = parseNumber(el("yardas").value);
  const qtyOrder = roundToNextThousand(requestedQty);
  const price = getPrice(provider);

  return {
    poBase: poDigitsOnly(el("po").value),
    provider,
    fechaPo: el("fechaPo").value || todayYMD(),
    fechaProd: el("fechaProd").value || addDays(el("fechaPo").value || todayYMD(), 7),
    styleRef: (el("styleRef").value || "").trim(),
    color: (el("color").value || "").trim(),
    requestedQty,
    qtyOrder,
    price,
    amount: qtyOrder * price,
    imageBase64: mobilonState.imageBase64,
    imageExtension: mobilonState.imageExtension
  };
}

function updateSummary() {
  const data = getFormData();
  el("qtyOrder").value = data.qtyOrder ? formatInt(data.qtyOrder) : "";
  el("sumProveedor").textContent = data.provider;
  el("sumPrecio").textContent = formatCurrency(data.price);
  el("sumQty").textContent = formatInt(data.qtyOrder);
  el("sumAmount").textContent = formatCurrencyAmount(data.amount);

  if (data.requestedQty > 0 && data.qtyOrder !== data.requestedQty) {
    setMessage(`La cantidad se ajustó a ${formatInt(data.qtyOrder)} porque Mobilon solo se compra en múltiplos de 1,000.`, "success");
  } else {
    setMessage("");
  }
}

async function handlePaste(event) {
  const items = [...(event.clipboardData?.items || [])];
  const imageItem = items.find((item) => item.type.startsWith("image/"));
  if (!imageItem) return;
  event.preventDefault();

  const file = imageItem.getAsFile();
  if (!file) return;

  mobilonState.imageBase64 = await blobToDataURL(file);
  mobilonState.imageExtension = file.type.includes("png")
    ? "png"
    : file.type.includes("webp")
      ? "webp"
      : "jpeg";

  updatePastePreview();
}

function updatePastePreview() {
  const preview = el("pastePreview");
  const empty = el("pasteEmpty");
  const clearBtn = el("clearPaste");

  if (mobilonState.imageBase64) {
    preview.src = mobilonState.imageBase64;
    preview.classList.remove("hidden");
    empty.classList.add("hidden");
    clearBtn.classList.remove("hidden");
  } else {
    preview.removeAttribute("src");
    preview.classList.add("hidden");
    empty.classList.remove("hidden");
    clearBtn.classList.add("hidden");
  }
}

function clearPastedImage() {
  mobilonState.imageBase64 = null;
  mobilonState.imageExtension = null;
  updatePastePreview();
}

function validateData(data) {
  if (!data.poBase) return "Coloca el número de PO antes de descargar.";
  if (!data.styleRef) return "Escribe el style del item.";
  if (!data.color) return "Escribe el color del Mobilon.";
  if (data.requestedQty <= 0) return "La cantidad de yardas debe ser mayor que cero.";
  if (data.qtyOrder <= 0) return "No pude calcular la cantidad final de compra.";
  return "";
}

async function loadTemplateWorkbook() {
  const response = await fetch(MOBILON_TEMPLATE_PATH);
  if (!response.ok) {
    throw new Error(`No se pudo cargar la plantilla ${MOBILON_TEMPLATE_PATH}`);
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  setWorkbookMeta(workbook);
  return workbook;
}

function fillTemplate(workbook, data) {
  const ws = workbook.getWorksheet(1);
  ws.getCell("C1").value = `${data.poBase}-MOB`;
  ws.getCell("C2").value = data.provider;
  ws.getCell("C3").value = formatDateExcel(data.fechaPo);
  ws.getCell("C4").value = formatDateExcel(data.fechaProd);
  ws.getCell("C7").value = data.styleRef;
  ws.getCell("D7").value = data.color;
  ws.getCell("E7").value = data.qtyOrder;
  ws.getCell("F7").value = data.price;
  ws.getCell("G7").value = data.amount;
  ws.getCell("E8").value = data.qtyOrder;
  ws.getCell("G8").value = data.amount;

  ws.getCell("C3").numFmt = "dd/mmm/yy";
  ws.getCell("C4").numFmt = "dd/mmm/yy";
  ws.getCell("E7").numFmt = "#,##0";
  ws.getCell("E8").numFmt = "#,##0";
  ws.getCell("F7").numFmt = "$#,##0.0000";
  ws.getCell("G7").numFmt = "$#,##0.00";
  ws.getCell("G8").numFmt = "$#,##0.00";

  if (data.imageBase64) {
    const imageId = workbook.addImage({
      base64: data.imageBase64,
      extension: data.imageExtension || "png"
    });

    ws.addImage(imageId, {
      tl: { col: 0.14, row: 6.18 },
      ext: { width: 96, height: 138 },
      editAs: "oneCell"
    });
  }
}

async function downloadExcel() {
  const data = getFormData();
  const error = validateData(data);
  if (error) {
    setMessage(error, "error");
    return;
  }

  try {
    const workbook = await loadTemplateWorkbook();
    fillTemplate(workbook, data);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    forceDownload(blob, safeFileName(`${data.poBase}-MOB.xlsx`));
    setMessage(`Listo. Se generó el archivo ${data.poBase}-MOB.xlsx`, "success");
  } catch (error) {
    console.error(error);
    setMessage("Hubo un problema generando el Excel de Mobilon. Ábrelo con Live Server para que cargue bien la plantilla.", "error");
  }
}

function bindEvents() {
  ["po", "proveedor", "fechaPo", "fechaProd", "styleRef", "color", "yardas"].forEach((id) => {
    el(id).addEventListener("input", updateSummary);
    el(id).addEventListener("change", updateSummary);
  });

  el("fechaPo").addEventListener("change", (event) => {
    el("fechaProd").value = addDays(event.target.value, 7);
    updateSummary();
  });

  const zone = el("pasteZone");
  zone.addEventListener("paste", handlePaste);
  zone.addEventListener("click", () => zone.focus());
  el("clearPaste").addEventListener("click", (event) => {
    event.stopPropagation();
    clearPastedImage();
  });

  el("downloadBtn").addEventListener("click", downloadExcel);
}

function init() {
  setTodayDefaults();
  bindEvents();
  updateSummary();
}

document.addEventListener("DOMContentLoaded", init);
