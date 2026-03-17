const el = (id) => document.getElementById(id);

const MOBILON_TEMPLATE_PATH = "plantillas/PLANTILLA-MOB.xlsx";
const PROVIDER_PRICES = {
  "SUMINISTROS ARYEL": 0.0065,
  JCA: 0.009
};

const mobilonState = {
  styles: []
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
  return String(name || "archivo").replace(/[\\/:*?"<>|]+/g, "_");
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

function ensureStyleSlots(count) {
  const safeCount = Math.min(20, Math.max(1, Number(count) || 1));
  while (mobilonState.styles.length < safeCount) {
    mobilonState.styles.push({
      styleRef: "",
      requestedQty: "",
      qtyOrder: 0,
      imageBase64: null,
      imageExtension: null
    });
  }
  mobilonState.styles = mobilonState.styles.slice(0, safeCount);
  return safeCount;
}

function getGeneralData() {
  const provider = el("proveedor").value || "SUMINISTROS ARYEL";
  return {
    poBase: poDigitsOnly(el("po").value),
    provider,
    price: getPrice(provider),
    fechaPo: el("fechaPo").value || todayYMD(),
    fechaProd: el("fechaProd").value || addDays(el("fechaPo").value || todayYMD(), 7)
  };
}

function getStyleData(index) {
  const style = mobilonState.styles[index] || {};
  const requestedQty = parseNumber(style.requestedQty);
  const qtyOrder = roundToNextThousand(requestedQty);
  return {
    styleRef: String(style.styleRef || "").trim(),
    color: "CLEAR",
    requestedQty,
    qtyOrder,
    amount: qtyOrder * getGeneralData().price,
    imageBase64: style.imageBase64 || null,
    imageExtension: style.imageExtension || null
  };
}

function renderStyles() {
  const container = el("stylesContainer");
  const countValue = el("qtyStyles").value;
  const count = ensureStyleSlots(countValue);

  container.innerHTML = mobilonState.styles.map((style, index) => {
    const data = getStyleData(index);
    const hasImage = Boolean(style.imageBase64);
    const previewMarkup = hasImage
      ? `<img class="paste-zone__preview" src="${style.imageBase64}" alt="Vista previa del item ${index + 1}" />`
      : `<div class="paste-zone__empty">
          <i class="fa-regular fa-image"></i>
          <strong>Pega aquí la imagen</strong>
          <small>Haz clic en esta zona y usa Ctrl + V</small>
        </div>`;

    return `
      <article class="glass-subpanel style-entry style-entry--mobilon" data-index="${index}">
        <div class="style-entry__head">
          <div>
            <small>Estilo #${index + 1}</small>
            <h3>Información del estilo</h3>
          </div>
          <div class="style-entry__actions">
            <span class="style-entry__badge">Se coloca en la misma hoja</span>
          </div>
        </div>

        <div class="style-entry__body">
          <div class="style-grid mobilon-style-grid">
            <label>
              <span>Style</span>
              <input type="text" data-field="styleRef" data-index="${index}" value="${escapeHtml(style.styleRef || "")}" placeholder="Ej. BLC524Y" />
            </label>

            <label>
              <span>Yardas requeridas</span>
              <input type="number" min="0" step="1" data-field="requestedQty" data-index="${index}" value="${style.requestedQty ?? ""}" placeholder="Ej. 15000" />
            </label>

            <label>
              <span>Qty order total</span>
              <input type="text" value="${data.qtyOrder ? formatInt(data.qtyOrder) : ""}" readonly placeholder="0" />
            </label>

            <div class="mobilon-locked">
              <span>Color fijo</span>
              <div class="mobilon-locked__value">CLEAR</div>
            </div>

            <div class="mobilon-note">
              <i class="fa-solid fa-circle-info"></i>
              <span>Si escribes una cantidad que no cae exacta, la orden de este estilo se sube al siguiente múltiplo de 1,000.</span>
            </div>

            <div class="paste-wrap paste-wrap--compact" data-index="${index}">
              <span class="paste-wrap__label">Imagen del item (pegar desde portapapeles)</span>
              <div class="paste-zone" data-paste-zone="${index}" tabindex="0" role="button" aria-label="Pega la imagen del item del estilo ${index + 1}">
                ${previewMarkup}
                <button class="paste-zone__clear ${hasImage ? "" : "hidden"}" type="button" data-clear-image="${index}">Limpiar</button>
              </div>
            </div>
          </div>
        </div>
      </article>
    `;
  }).join("");

  bindStyleEvents();
  updateProviderLabel();
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function updateProviderLabel() {
  const provider = el("proveedor").value || "SUMINISTROS ARYEL";
  el("providerPriceLabel").textContent = formatCurrency(getPrice(provider));
}

function updateStyleField(index, field, value) {
  if (!mobilonState.styles[index]) return;
  mobilonState.styles[index][field] = value;
}

function refreshStyleEntry(index) {
  const article = document.querySelector(`.style-entry[data-index="${index}"]`);
  if (!article) return;
  const qtyField = article.querySelector("[readonly]");
  if (qtyField) {
    const data = getStyleData(index);
    qtyField.value = data.qtyOrder ? formatInt(data.qtyOrder) : "";
  }
}

async function handlePaste(event, index) {
  const items = [...(event.clipboardData?.items || [])];
  const imageItem = items.find((item) => item.type.startsWith("image/"));
  if (!imageItem) return;
  event.preventDefault();

  const file = imageItem.getAsFile();
  if (!file || !mobilonState.styles[index]) return;

  mobilonState.styles[index].imageBase64 = await blobToDataURL(file);
  mobilonState.styles[index].imageExtension = file.type.includes("png")
    ? "png"
    : file.type.includes("webp")
      ? "webp"
      : "jpeg";

  renderStyles();
}

function clearPastedImage(index) {
  if (!mobilonState.styles[index]) return;
  mobilonState.styles[index].imageBase64 = null;
  mobilonState.styles[index].imageExtension = null;
  renderStyles();
}

function getAllData() {
  const general = getGeneralData();
  const styles = mobilonState.styles.map((_, index) => getStyleData(index));
  return { ...general, styles };
}

function validateData(data) {
  if (!data.poBase) return "Coloca el número de PO antes de descargar.";
  if (!data.styles.length) return "Agrega al menos un estilo.";

  for (let index = 0; index < data.styles.length; index += 1) {
    const style = data.styles[index];
    if (!style.styleRef) return `Escribe el style del estilo #${index + 1}.`;
    if (style.requestedQty <= 0) return `La cantidad de yardas del estilo #${index + 1} debe ser mayor que cero.`;
    if (style.qtyOrder <= 0) return `No pude calcular la cantidad final del estilo #${index + 1}.`;
  }

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

function copyCellStyle(fromCell, toCell) {
  toCell.style = JSON.parse(JSON.stringify(fromCell.style || {}));
  if (fromCell.numFmt) toCell.numFmt = fromCell.numFmt;
  if (fromCell.alignment) toCell.alignment = JSON.parse(JSON.stringify(fromCell.alignment));
  if (fromCell.fill) toCell.fill = JSON.parse(JSON.stringify(fromCell.fill));
  if (fromCell.border) toCell.border = JSON.parse(JSON.stringify(fromCell.border));
  if (fromCell.font) toCell.font = JSON.parse(JSON.stringify(fromCell.font));
}

function cloneRowLayout(ws, sourceRowIndex, targetRowIndex) {
  const sourceRow = ws.getRow(sourceRowIndex);
  const targetRow = ws.getRow(targetRowIndex);
  targetRow.height = sourceRow.height;

  for (let col = 1; col <= 7; col += 1) {
    const sourceCell = ws.getCell(sourceRowIndex, col);
    const targetCell = ws.getCell(targetRowIndex, col);
    copyCellStyle(sourceCell, targetCell);
    targetCell.value = null;
  }
}

function fillTemplate(workbook, data) {
  const ws = workbook.getWorksheet(1);
  const styles = data.styles;
  const extraRows = Math.max(0, styles.length - 1);
  const baseItemRow = 7;
  const baseTotalRow = 8;

  if (ws.hasMerges && ws.getCell("A8").isMerged) {
    ws.unMergeCells("A8:D8");
  }

  if (extraRows > 0) {
    ws.spliceRows(baseTotalRow, 0, ...Array.from({ length: extraRows }, () => [null, null, null, null, null, null, null]));
    for (let i = 0; i < extraRows; i += 1) {
      cloneRowLayout(ws, baseItemRow, baseItemRow + 1 + i);
    }
  }

  const totalRowIndex = baseItemRow + styles.length;
  ws.mergeCells(`A${totalRowIndex}:D${totalRowIndex}`);

  ws.getCell("C1").value = `${data.poBase}-MOB`;
  ws.getCell("C2").value = data.provider;
  ws.getCell("C3").value = formatDateExcel(data.fechaPo);
  ws.getCell("C4").value = formatDateExcel(data.fechaProd);

  ws.getCell("C3").numFmt = "dd/mmm/yy";
  ws.getCell("C4").numFmt = "dd/mmm/yy";

  styles.forEach((style, index) => {
    const rowIndex = baseItemRow + index;
    ws.getCell(`B${rowIndex}`).value = "MOBILON\nTAPE";
    ws.getCell(`C${rowIndex}`).value = style.styleRef;
    ws.getCell(`D${rowIndex}`).value = "CLEAR";
    ws.getCell(`E${rowIndex}`).value = style.qtyOrder;
    ws.getCell(`F${rowIndex}`).value = data.price;
    ws.getCell(`G${rowIndex}`).value = style.amount;

    ws.getCell(`E${rowIndex}`).numFmt = "#,##0";
    ws.getCell(`F${rowIndex}`).numFmt = "$#,##0.0000";
    ws.getCell(`G${rowIndex}`).numFmt = "$#,##0.00";

    if (style.imageBase64) {
      const imageId = workbook.addImage({
        base64: style.imageBase64,
        extension: style.imageExtension || "png"
      });

      ws.addImage(imageId, {
        tl: { col: 0.14, row: (rowIndex - 1) + 0.18 },
        ext: { width: 96, height: 138 },
        editAs: "oneCell"
      });
    }
  });

  ws.getCell(`A${totalRowIndex}`).value = "TOTAL";
  ws.getCell(`E${totalRowIndex}`).value = styles.reduce((sum, style) => sum + style.qtyOrder, 0);
  ws.getCell(`G${totalRowIndex}`).value = styles.reduce((sum, style) => sum + style.amount, 0);
  ws.getCell(`E${totalRowIndex}`).numFmt = "#,##0";
  ws.getCell(`G${totalRowIndex}`).numFmt = "$#,##0.00";
}

async function downloadExcel() {
  const data = getAllData();
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

function bindStyleEvents() {
  document.querySelectorAll("[data-field]").forEach((input) => {
    input.addEventListener("input", (event) => {
      const index = Number(event.currentTarget.dataset.index);
      const field = event.currentTarget.dataset.field;
      updateStyleField(index, field, event.currentTarget.value);
      refreshStyleEntry(index);
      setMessage("");
    });
  });

  document.querySelectorAll("[data-paste-zone]").forEach((zone) => {
    const index = Number(zone.dataset.pasteZone);
    zone.addEventListener("paste", (event) => handlePaste(event, index));
    zone.addEventListener("click", () => zone.focus());
  });

  document.querySelectorAll("[data-clear-image]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      clearPastedImage(Number(event.currentTarget.dataset.clearImage));
    });
  });
}

function bindEvents() {
  ["po", "proveedor", "fechaPo", "fechaProd"].forEach((id) => {
    el(id).addEventListener("input", () => {
      updateProviderLabel();
      setMessage("");
    });
    el(id).addEventListener("change", () => {
      updateProviderLabel();
      setMessage("");
    });
  });

  el("qtyStyles").addEventListener("input", () => {
    renderStyles();
    setMessage("");
  });

  el("qtyStyles").addEventListener("blur", () => {
    const count = ensureStyleSlots(el("qtyStyles").value);
    el("qtyStyles").value = count;
    renderStyles();
  });

  el("fechaPo").addEventListener("change", (event) => {
    el("fechaProd").value = addDays(event.target.value, 7);
  });

  el("downloadBtn").addEventListener("click", downloadExcel);
}

function init() {
  setTodayDefaults();
  bindEvents();
  renderStyles();
  updateProviderLabel();
}

document.addEventListener("DOMContentLoaded", init);
