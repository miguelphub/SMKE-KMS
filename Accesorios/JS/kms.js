const el = (id) => document.getElementById(id);

const SIZE_SETS = {
  regulares: ["XXS", "XS", "S", "M", "L", "XL", "XXL"],
  toddler: ["12M", "18M", "2T", "3T", "4T", "5T"],
  oldnavy: []
};

const PLUS_SIZES = ["1X", "2X", "3X"];

const COLOR_TAG_OPTIONS = {
  pink: {
    label: "Pink",
    description: "HANG TAG PINK / TAMAÑO  38MM*83MMM",
    imagePath: "assets/color_tag_pink.png"
  },
  grey: {
    label: "Grey",
    description: "HANG TAG GREY / TAMAÑO  38MM*83MMM",
    imagePath: "assets/color_tag_grey.png"
  }
};

const assetBase64Cache = {};
const accessoryState = { styles: [] };

// =========================
// Helpers generales
// =========================
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

function poDigitsOnly(raw) {
  const s = String(raw || "").trim();
  const digits = s.replace(/\D+/g, "");
  return digits || s;
}

function safeFileName(name) {
  return String(name || "archivo").replace(/[\\/:*?"<>|]+/g, "_");
}

function sizeToId(size) {
  return String(size).replace(/[^a-zA-Z0-9_]/g, "");
}

function qtyVal(id) {
  const node = el(id);
  if (!node) return null;
  const v = node.value;
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function setWorkbookMeta(workbook) {
  workbook.creator = "Miguelnmms";
  workbook.lastModifiedBy = "Miguelnmms";
  workbook.created = new Date();
  workbook.modified = new Date();
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

function borderStyle(style) {
  return style ? { style } : undefined;
}

function makeBorder({ top = "thin", left = "thin", bottom = "thin", right = "thin" } = {}) {
  return {
    top: borderStyle(top),
    left: borderStyle(left),
    bottom: borderStyle(bottom),
    right: borderStyle(right)
  };
}

function cellAddr(row, col) {
  return `${ExcelJS.Workbook.xlsx ? "" : ""}`;
}

function getColorTagConfig(colorKey) {
  return COLOR_TAG_OPTIONS[colorKey] || COLOR_TAG_OPTIONS.pink;
}

async function blobToDataURL(blob) {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function fetchAssetAsBase64(path) {
  if (assetBase64Cache[path]) return assetBase64Cache[path];

  const res = await fetch(path);
  if (!res.ok) {
    throw new Error(`No se pudo cargar el recurso: ${path}`);
  }

  const blob = await res.blob();
  const base64 = await blobToDataURL(blob);
  assetBase64Cache[path] = base64;
  return base64;
}

function hasBaseQty(data) {
  return data.baseFiltered.length > 0;
}

function hasPlusQty(data) {
  return data.sizeType === "regulares" && data.plusEnabled && data.plusFiltered.length > 0;
}

function hasHeatTransferQty(data) {
  return hasBaseQty(data) || hasPlusQty(data);
}

function hasSizeStripQty(data) {
  return hasBaseQty(data);
}

function hasHangTagQty(data) {
  return Number.isFinite(data.hangTagQty) && data.hangTagQty > 0;
}

function hasColorTagQty(data) {
  return Number.isFinite(data.colorTagQty) && data.colorTagQty > 0;
}

// =========================
// Compras seleccionadas
// =========================
function getSelectedCompras() {
  const selected = [];
  if (el("buyHT")?.checked) selected.push("heat_transfer");
  if (el("buySS")?.checked) selected.push("size_strip");
  if (el("buyCT")?.checked) selected.push("color_tag");
  if (el("buyHG")?.checked) selected.push("hang_tag");

  const sel = el("compras");
  if (sel) {
    [...sel.options].forEach(o => {
      o.selected = selected.includes(o.value);
    });
  }

  return selected;
}

function isHTSelected() {
  return !!el("buyHT")?.checked;
}

function isCTSelected() {
  return !!el("buyCT")?.checked;
}

function isHGSelected() {
  return !!el("buyHG")?.checked;
}

// =========================
// Inputs dinámicos
// =========================
function renderStyleNameInputs(n) {
  const box = el("styleNamesBox");
  box.innerHTML = "";

  for (let i = 1; i <= n; i++) {
    const wrap = document.createElement("label");
    wrap.innerHTML = `
      Referencia estilo #${i} (opcional)
      <input id="styleName${i}" type="text" placeholder="(opcional)" />
    `;
    box.appendChild(wrap);
  }
}

function ensureAccessoryStateSize(n) {
  while (accessoryState.styles.length < n) {
    accessoryState.styles.push({ imageBase64: null, imageExtension: null });
  }
  if (accessoryState.styles.length > n) {
    accessoryState.styles = accessoryState.styles.slice(0, n);
  }
}

function captureCurrentStyleValues() {
  const cards = [...document.querySelectorAll(".styleCard")];
  return cards.map((card, idx) => {
    const i = idx + 1;
    const key = el(`sizeType_${i}`)?.value || "regulares";
    const sizes = SIZE_SETS[key] || SIZE_SETS.regulares;
    const baseQty = {};
    sizes.forEach((s) => {
      baseQty[s] = el(`q${sizeToId(s)}_${i}`)?.value || "";
    });
    baseQty["1X"] = el(`q1X_${i}`)?.value || "";
    baseQty["2X"] = el(`q2X_${i}`)?.value || "";
    baseQty["3X"] = el(`q3X_${i}`)?.value || "";
    return {
      styleName: el(`styleName${i}`)?.value || "",
      itemText: el(`item_${i}`)?.value || "",
      color: el(`color_${i}`)?.value || "",
      content: el(`content_${i}`)?.value || "",
      ctColor: el(`ctColor_${i}`)?.value || "pink",
      ctQty: el(`ctQty_${i}`)?.value || "",
      hgQty: el(`hgQty_${i}`)?.value || "",
      sizeType: key,
      plus: el(`plus_${i}`)?.value || "no",
      baseQty
    };
  });
}

function applyAccessoryPreservedValues(values = []) {
  values.forEach((item, idx) => {
    const i = idx + 1;
    if (el(`styleName${i}`)) el(`styleName${i}`).value = item.styleName || "";
    if (!el(`item_${i}`)) return;
    el(`item_${i}`).value = item.itemText || "";
    el(`color_${i}`).value = item.color || "";
    el(`content_${i}`).value = item.content || "";
    el(`ctColor_${i}`).value = item.ctColor || "pink";
    el(`ctQty_${i}`).value = item.ctQty || "";
    if (el(`hgQty_${i}`)) el(`hgQty_${i}`).value = item.hgQty || "";
    el(`sizeType_${i}`).value = item.sizeType || "regulares";

    const sizeTypeSel = el(`sizeType_${i}`);
    const sizes = SIZE_SETS[sizeTypeSel.value || "regulares"] || SIZE_SETS.regulares;
    const table = el(`table_${i}`);
    table.style.setProperty("--cols", String(sizes.length));
    el(`thead_${i}`).innerHTML = sizes.map(s => `<div>${s}</div>`).join("");
    el(`tqty_${i}`).innerHTML = sizes.map(s => {
      const id = `q${sizeToId(s)}_${i}`;
      return `<input id="${id}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />`;
    }).join("");

    el(`plus_${i}`).value = item.plus || "no";
    sizes.forEach((s) => {
      const node = el(`q${sizeToId(s)}_${i}`);
      if (node) node.value = item.baseQty?.[s] || "";
    });
    ["1X","2X","3X"].forEach((s) => {
      const node = el(`q${s}_${i}`);
      if (node) node.value = item.baseQty?.[s] || "";
    });

    refreshStyleOptionVisibility(i);
    updateAccessoryPastePreview(i);
  });
}

async function handleAccessoryPaste(event, index) {
  const items = [...(event.clipboardData?.items || [])];
  const imageItem = items.find((item) => item.type.startsWith("image/"));
  if (!imageItem) return;
  event.preventDefault();

  const file = imageItem.getAsFile();
  if (!file) return;

  const base64 = await blobToDataURL(file);
  const extension = file.type.includes("png") ? "png" : file.type.includes("webp") ? "webp" : "jpeg";
  accessoryState.styles[index - 1] = { imageBase64: base64, imageExtension: extension };
  updateAccessoryPastePreview(index);
}

function clearAccessoryPastedImage(index) {
  accessoryState.styles[index - 1] = { imageBase64: null, imageExtension: null };
  updateAccessoryPastePreview(index);
}

function updateAccessoryPastePreview(index) {
  const imageState = accessoryState.styles[index - 1] || { imageBase64: null, imageExtension: null };
  const preview = el(`pastePreview_${index}`);
  const empty = el(`pasteEmpty_${index}`);
  const clearBtn = el(`clearPaste_${index}`);
  if (!preview || !empty || !clearBtn) return;
  if (imageState.imageBase64) {
    preview.src = imageState.imageBase64;
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

function refreshStyleOptionVisibility(i) {
  const colorWrap = el(`colorWrap_${i}`);
  const contentWrap = el(`contentWrap_${i}`);
  const plusWrap = el(`plusWrap_${i}`);
  const plusTable = el(`plusTable_${i}`);
  const sizeTypeSel = el(`sizeType_${i}`);
  const plusSel = el(`plus_${i}`);
  const ctColorWrap = el(`ctColorWrap_${i}`);
  const ctQtyWrap = el(`ctQtyWrap_${i}`);
  const hgQtyWrap = el(`hgQtyWrap_${i}`);
  const pasteWrap = el(`pasteWrap_${i}`);
  const sizeTypeWrap = el(`sizeTypeWrap_${i}`);
  const plusArea = el(`plusArea_${i}`);
  const sizeNote = el(`sizeNote_${i}`);
  const table = el(`table_${i}`);

  const htOn = isHTSelected();
  const ssOn = !!el("buySS")?.checked;
  const ctOn = isCTSelected();
  const hgOn = isHGSelected();
  const selectedType = sizeTypeSel?.value || "regulares";
  const isRegular = selectedType === "regulares";
  const isOldNavy = selectedType === "oldnavy";
  const showSizeSection = htOn || ssOn;

  if (colorWrap) {
    colorWrap.style.display = htOn ? "flex" : "none";
    colorWrap.style.flexDirection = "column";
  }

  if (contentWrap) {
    contentWrap.style.display = htOn ? "flex" : "none";
    contentWrap.style.flexDirection = "column";
  }

  if (sizeTypeWrap) {
    sizeTypeWrap.style.display = showSizeSection ? "flex" : "none";
    sizeTypeWrap.style.flexDirection = "column";
  }

  if (plusWrap) {
    plusWrap.style.display = (showSizeSection && htOn && isRegular) ? "flex" : "none";
    plusWrap.style.flexDirection = "column";
    if (!(showSizeSection && htOn && isRegular) && plusSel) {
      plusSel.value = "no";
    }
  }

  if (plusArea) {
    plusArea.classList.toggle("hidden", !showSizeSection || isOldNavy);
  }

  if (plusTable) {
    plusTable.hidden = !(showSizeSection && htOn && isRegular && plusSel?.value === "si");
  }

  if (table) {
    table.classList.toggle("hidden", !showSizeSection || isOldNavy);
  }

  if (sizeNote) {
    sizeNote.classList.toggle("hidden", !showSizeSection || !isOldNavy);
  }

  if (ctColorWrap) {
    ctColorWrap.style.display = ctOn ? "flex" : "none";
    ctColorWrap.style.flexDirection = "column";
  }

  if (ctQtyWrap) {
    ctQtyWrap.style.display = ctOn ? "flex" : "none";
    ctQtyWrap.style.flexDirection = "column";
  }

  if (hgQtyWrap) {
    hgQtyWrap.style.display = hgOn ? "flex" : "none";
    hgQtyWrap.style.flexDirection = "column";
  }

  if (pasteWrap) {
    pasteWrap.classList.toggle("hidden", !htOn);
  }
}

function renderStyles(n) {
  const cont = el("stylesContainer");
  if (!cont) return;

  const preserved = captureCurrentStyleValues();
  ensureAccessoryStateSize(n);
  cont.innerHTML = "";

  for (let i = 1; i <= n; i++) {
    const card = document.createElement("div");
    card.className = "styleCard";

    card.innerHTML = `
      <div class="styleHeader">
        <h3>Estilo #${i}</h3>
        <div class="styleMeta">* Cantidades vacías no se incluyen en el Excel</div>
      </div>

      <div class="grid">
        <label class="full">
          ITEM / H-TRANSFER / SIZE STRIP / TAG
          <input id="item_${i}" type="text" placeholder="(opcional)" />
        </label>

        <label id="colorWrap_${i}">
          Color de transfer <span class="onlyHT">(solo Heat Transfer)</span>
          <input id="color_${i}" type="text" placeholder="White / Black / Green..." />
        </label>

        <label id="contentWrap_${i}">
          Contenido de transfer <span class="onlyHT">(solo Heat Transfer)</span>
          <select id="content_${i}">
            <option value="" selected>Selecciona una opción</option>
            <option value="100% COTTON">100% COTTON</option>
            <option value="50% COTTON 50% POLYESTER">50% COTTON 50% POLYESTER</option>
            <option value="90% COTTON 10% POLYESTER">90% COTTON 10% POLYESTER</option>
          </select>
        </label>

        <label id="ctColorWrap_${i}">
          Color de Color Tag <span class="onlyCT">(solo Color Tag)</span>
          <select id="ctColor_${i}">
            <option value="pink" selected>Pink</option>
            <option value="grey">Grey</option>
          </select>
        </label>

        <label id="ctQtyWrap_${i}">
          Cantidad Color Tag <span class="onlyCT">(manual, ya no sale de tallas)</span>
          <input id="ctQty_${i}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />
        </label>

        <label id="hgQtyWrap_${i}">
          Cantidad Hang Tag <span class="onlyCT">(manual)</span>
          <input id="hgQty_${i}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />
        </label>

        <label id="sizeTypeWrap_${i}">
          Tipo de tallas
          <select id="sizeType_${i}">
            <option value="regulares" selected>Regulares</option>
            <option value="toddler">Toddler</option>
            <option value="oldnavy">OLD NAVY</option>
          </select>
        </label>

        <label id="plusWrap_${i}">
          Tallas plus (solo Regulares y Heat Transfer)
          <select id="plus_${i}">
            <option value="no" selected>No</option>
            <option value="si">Sí</option>
          </select>
        </label>
      </div>

      <div class="table" id="table_${i}">
        <div class="thead" id="thead_${i}"></div>
        <div class="tqty" id="tqty_${i}"></div>
      </div>

      <div class="sizePending hidden" id="sizeNote_${i}">OLD NAVY queda agregado como opción, pero las tallas todavía están pendientes de definirse.</div>

      <div class="plusBox" id="plusArea_${i}">
        <div class="table plusTable" id="plusTable_${i}" hidden>
          <div class="thead" style="--cols:3;">
            <div>1X</div>
            <div>2X</div>
            <div>3X</div>
          </div>
          <div class="tqty" style="--cols:3;">
            <input id="q1X_${i}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />
            <input id="q2X_${i}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />
            <input id="q3X_${i}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />
          </div>
        </div>
      </div>

      <div class="stylePasteWrap hidden" id="pasteWrap_${i}">
        <span class="pasteTitle">Imagen del transfer (pegar desde portapapeles)</span>
        <div class="pasteZoneAcc" id="pasteZone_${i}" tabindex="0" role="button" aria-label="Pega una imagen para el estilo ${i}">
          <div class="pasteZoneAcc__empty" id="pasteEmpty_${i}">
            <i class="fa-regular fa-image"></i>
            <strong>Pega aquí la imagen</strong>
            <small>Haz clic en esta zona y usa Ctrl + V</small>
          </div>
          <img id="pastePreview_${i}" class="pasteZoneAcc__preview hidden" alt="Vista previa del estilo ${i}" />
          <button id="clearPaste_${i}" class="pasteZoneAcc__clear hidden" type="button">Limpiar</button>
        </div>
      </div>
    `;

    cont.appendChild(card);

    const sizeTypeSel = el(`sizeType_${i}`);
    const plusSel = el(`plus_${i}`);

    function renderTable() {
      const key = sizeTypeSel.value || "regulares";
      const sizes = SIZE_SETS[key] || SIZE_SETS.regulares;

      const table = el(`table_${i}`);
      table.style.setProperty("--cols", String(Math.max(sizes.length, 1)));

      if (sizes.length === 0) {
        el(`thead_${i}`).innerHTML = "";
        el(`tqty_${i}`).innerHTML = "";
        refreshStyleOptionVisibility(i);
        return;
      }

      el(`thead_${i}`).innerHTML = sizes.map(s => `<div>${s}</div>`).join("");
      el(`tqty_${i}`).innerHTML = sizes.map(s => {
        const id = `q${sizeToId(s)}_${i}`;
        return `<input id="${id}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />`;
      }).join("");

      refreshStyleOptionVisibility(i);
    }

    sizeTypeSel.addEventListener("change", renderTable);
    plusSel.addEventListener("change", () => refreshStyleOptionVisibility(i));

    const pasteZone = el(`pasteZone_${i}`);
    pasteZone.addEventListener("paste", (event) => handleAccessoryPaste(event, i));
    pasteZone.addEventListener("click", () => pasteZone.focus());
    el(`clearPaste_${i}`).addEventListener("click", (event) => {
      event.stopPropagation();
      clearAccessoryPastedImage(i);
    });

    renderTable();
  }

  applyAccessoryPreservedValues(preserved);
  refreshAllStyleOptionVisibility();
}

function refreshAllStyleOptionVisibility() {
  let n = parseInt(el("qtyStyles")?.value, 10);
  if (!Number.isFinite(n) || n <= 0) n = 1;
  if (n > 20) n = 20;

  for (let i = 1; i <= n; i++) {
    refreshStyleOptionVisibility(i);
  }
}

function getStyleData(i) {
  const itemText = (el(`item_${i}`)?.value || "").trim();
  const refText = (el(`styleName${i}`)?.value || "").trim();
  const color = (el(`color_${i}`)?.value || "").trim();
  const content = (el(`content_${i}`)?.value || "").trim();
  const sizeType = el(`sizeType_${i}`)?.value || "regulares";
  const plusEnabled = (el(`plus_${i}`)?.value || "no") === "si";
  const colorTagColor = el(`ctColor_${i}`)?.value || "pink";
  const colorTagQty = qtyVal(`ctQty_${i}`);
  const hangTagQty = qtyVal(`hgQty_${i}`);

  const baseSizes = SIZE_SETS[sizeType] || [];
  const qtyBase = {};

  for (const s of baseSizes) {
    qtyBase[s] = qtyVal(`q${sizeToId(s)}_${i}`);
  }

  const baseFiltered = baseSizes.filter(s => qtyBase[s] !== null);

  const plusQty = {
    "1X": qtyVal(`q1X_${i}`),
    "2X": qtyVal(`q2X_${i}`),
    "3X": qtyVal(`q3X_${i}`)
  };

  const plusFiltered = PLUS_SIZES.filter(s => plusQty[s] !== null);

  const imageState = accessoryState.styles[i - 1] || { imageBase64: null, imageExtension: null };

  return {
    itemText,
    refText,
    color,
    content,
    sizeType,
    plusEnabled,
    qtyBase,
    plusQty,
    baseFiltered,
    plusFiltered,
    colorTagColor,
    colorTagQty,
    hangTagQty,
    imageBase64: imageState.imageBase64,
    imageExtension: imageState.imageExtension
  };
}

function sumStyleQty(styleData, includePlus = true) {
  let total = 0;

  for (const s of styleData.baseFiltered) {
    total += Number(styleData.qtyBase[s] || 0);
  }

  if (includePlus && styleData.sizeType === "regulares" && styleData.plusEnabled) {
    for (const s of styleData.plusFiltered) {
      total += Number(styleData.plusQty[s] || 0);
    }
  }

  return total;
}

// =========================
// Helpers Excel
// =========================
function thinBorder() {
  return {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  };
}

function mediumBorder() {
  return {
    top: { style: "medium" },
    left: { style: "medium" },
    bottom: { style: "medium" },
    right: { style: "medium" }
  };
}

function setCell(ws, addr, value) {
  const c = ws.getCell(addr);
  c.value = value;
  return c;
}

function applyBorderRange(ws, r1, c1, r2, c2, border = thinBorder()) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getRow(r).getCell(c).border = border;
    }
  }
}

function fillRange(ws, r1, c1, r2, c2, color) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getRow(r).getCell(c).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color }
      };
    }
  }
}

function centerRange(ws, r1, c1, r2, c2) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getRow(r).getCell(c).alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true
      };
    }
  }
}

function setFontRange(ws, r1, c1, r2, c2, font) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getRow(r).getCell(c).font = {
        ...(ws.getRow(r).getCell(c).font || {}),
        ...font
      };
    }
  }
}

function setNumberFormatRange(ws, r1, c1, r2, c2, numFmt) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getRow(r).getCell(c).numFmt = numFmt;
    }
  }
}

function setOuterBorder(ws, r1, c1, r2, c2, style = "medium") {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getRow(r).getCell(c);
      const current = cell.border || {};
      const next = {
        top: current.top,
        left: current.left,
        bottom: current.bottom,
        right: current.right
      };

      if (r === r1) next.top = { style };
      if (r === r2) next.bottom = { style };
      if (c === c1) next.left = { style };
      if (c === c2) next.right = { style };

      cell.border = next;
    }
  }
}

function setupHeader8(ws, poText, fechaPoYMD, fechaProdYMD, lugar) {
  ws.columns = [
    { width: 18 },
    { width: 24 },
    { width: 16 },
    { width: 30 },
    { width: 12 },
    { width: 10 },
    { width: 10 },
    { width: 12 }
  ];

  ws.mergeCells("A1:B1");
  ws.mergeCells("C1:H1");
  ws.mergeCells("A2:B2");
  ws.mergeCells("C2:H2");
  ws.mergeCells("A3:B3");
  ws.mergeCells("C3:H3");
  ws.mergeCells("A4:B4");
  ws.mergeCells("C4:H4");

  setCell(ws, "A1", "PO #");
  setCell(ws, "C1", poText);
  setCell(ws, "A2", "FECHA DE PO:");
  setCell(ws, "C2", formatDateExcel(fechaPoYMD));
  setCell(ws, "A3", "NECESITO LA PRODUCCION:");
  setCell(ws, "C3", formatDateExcel(fechaProdYMD));
  setCell(ws, "A4", "LUGAR DE ENTREGA:");
  setCell(ws, "C4", lugar || "PENDIENTE");

  ws.getCell("C2").numFmt = "dd/mmm/yy";
  ws.getCell("C3").numFmt = "dd/mmm/yy";

  applyBorderRange(ws, 1, 1, 4, 8, thinBorder());
  centerRange(ws, 1, 1, 4, 8);
  setFontRange(ws, 1, 1, 4, 8, { bold: true, size: 14 });
  ws.getCell("C4").font = { bold: true, color: { argb: "FFFF0000" }, size: 14 };
}

function setupHeader7(ws, poText, fechaPoYMD, fechaProdYMD, lugar) {
  ws.columns = [
    { width: 18 },
    { width: 24 },
    { width: 22 },
    { width: 12 },
    { width: 10 },
    { width: 10 },
    { width: 12 }
  ];

  ws.mergeCells("A1:B1");
  ws.mergeCells("C1:G1");
  ws.mergeCells("A2:B2");
  ws.mergeCells("C2:G2");
  ws.mergeCells("A3:B3");
  ws.mergeCells("C3:G3");
  ws.mergeCells("A4:B4");
  ws.mergeCells("C4:G4");

  setCell(ws, "A1", "PO #");
  setCell(ws, "C1", poText);
  setCell(ws, "A2", "FECHA DE PO:");
  setCell(ws, "C2", formatDateExcel(fechaPoYMD));
  setCell(ws, "A3", "NECESITO LA PRODUCCION:");
  setCell(ws, "C3", formatDateExcel(fechaProdYMD));
  setCell(ws, "A4", "LUGAR DE ENTREGA:");
  setCell(ws, "C4", lugar || "PENDIENTE");

  ws.getCell("C2").numFmt = "dd/mmm/yy";
  ws.getCell("C3").numFmt = "dd/mmm/yy";

  applyBorderRange(ws, 1, 1, 4, 7, thinBorder());
  centerRange(ws, 1, 1, 4, 7);
  setFontRange(ws, 1, 1, 4, 7, { bold: true, size: 14 });
  ws.getCell("C4").font = { bold: true, color: { argb: "FFFF0000" }, size: 14 };
}

function setupHeader6(ws, poText, fechaPoYMD, fechaProdYMD, lugar) {
  ws.columns = [
    { width: 18 },
    { width: 22 },
    { width: 38 },
    { width: 10 },
    { width: 10 },
    { width: 12 }
  ];

  ws.mergeCells("A1:B1");
  ws.mergeCells("C1:F1");
  ws.mergeCells("A2:B2");
  ws.mergeCells("C2:F2");
  ws.mergeCells("A3:B3");
  ws.mergeCells("C3:F3");
  ws.mergeCells("A4:B4");
  ws.mergeCells("C4:F4");

  setCell(ws, "A1", "PO #");
  setCell(ws, "C1", poText);
  setCell(ws, "A2", "FECHA DE PO:");
  setCell(ws, "C2", formatDateExcel(fechaPoYMD));
  setCell(ws, "A3", "NECESITO LA PRODUCCION:");
  setCell(ws, "C3", formatDateExcel(fechaProdYMD));
  setCell(ws, "A4", "LUGAR DE ENTREGA:");
  setCell(ws, "C4", lugar || "PENDIENTE");

  ws.getCell("C2").numFmt = "dd/mmm/yy";
  ws.getCell("C3").numFmt = "dd/mmm/yy";

  applyBorderRange(ws, 1, 1, 4, 6, thinBorder());
  centerRange(ws, 1, 1, 4, 6);
  setFontRange(ws, 1, 1, 4, 6, { bold: true, size: 14 });
  ws.getCell("C4").font = { bold: true, color: { argb: "FFFF0000" }, size: 14 };
}

function setupColorTagSheet(ws, poText, fechaPoYMD, fechaProdYMD, lugar) {
  ws.columns = [
    { width: 30.43 },
    { width: 36.71 },
    { width: 65.71 },
    { width: 22.0 },
    { width: 19.29 },
    { width: 23.71 },
    { width: 23.0 }
  ];

  ws.getRow(1).height = 32.45;
  ws.getRow(2).height = 32.45;
  ws.getRow(3).height = 32.45;
  ws.getRow(4).height = 32.45;
  ws.getRow(5).height = 29.25;

  ws.mergeCells("A1:B1");
  ws.mergeCells("C1:D1");
  ws.mergeCells("A2:B2");
  ws.mergeCells("C2:D2");
  ws.mergeCells("A3:B3");
  ws.mergeCells("C3:D3");
  ws.mergeCells("A4:B4");
  ws.mergeCells("C4:D4");

  setCell(ws, "A1", "PO #");
  setCell(ws, "C1", poText);
  setCell(ws, "A2", "FECHA DE PO:");
  setCell(ws, "C2", formatDateExcel(fechaPoYMD));
  setCell(ws, "A3", "NECESITO LA PRODUCCION:");
  setCell(ws, "C3", formatDateExcel(fechaProdYMD));
  setCell(ws, "A4", "LUGAR DE ENTREGA:");
  setCell(ws, "C4", lugar || "PENDIENTE");
 
  ws.getCell("C2").numFmt = "dd/mmm/yy";
  ws.getCell("C3").numFmt = "dd/mmm/yy";

  fillRange(ws, 1, 1, 4, 4, "FFD9D9D9");
  centerRange(ws, 1, 1, 4, 4);
  applyBorderRange(ws, 1, 1, 4, 4, thinBorder());
  setOuterBorder(ws, 1, 1, 4, 4, "medium");

  setFontRange(ws, 1, 1, 4, 2, { name: "Tahoma", size: 20, bold: false });
  setFontRange(ws, 1, 3, 4, 4, { name: "Tahoma", size: 20, bold: true });
  ws.getCell("C4").font = { name: "Tahoma", size: 20, bold: true, color: { argb: "FFFF0000" } };
}

// =========================
// Bloques HT
// =========================
function writeHTBlock(ws, startRow, st) {
  const baseSizes = st.data.baseFiltered.slice();
  const plusSizes = (st.data.sizeType === "regulares" && st.data.plusEnabled)
    ? st.data.plusFiltered.slice()
    : [];

  const sizes = [...baseSizes, ...plusSizes];
  const qtyMap = { ...st.data.qtyBase, ...st.data.plusQty };

  const dataStart = startRow + 1;
  const dataEnd = dataStart + sizes.length - 1;
  const totalRow = dataEnd + 1;

  setCell(ws, `A${startRow}`, "H-TRANSFER");
  setCell(ws, `B${startRow}`, "ITEM #");
  setCell(ws, `C${startRow}`, "color de transfer");
  setCell(ws, `D${startRow}`, "CONTENIDO DE TRANSFER");
  setCell(ws, `E${startRow}`, "SIZE");
  setCell(ws, `F${startRow}`, "QTY");
  setCell(ws, `G${startRow}`, "UPPR");
  setCell(ws, `H${startRow}`, "TTL $");

  fillRange(ws, startRow, 1, startRow, 8, "FFD9D9D9");
  centerRange(ws, startRow, 1, startRow, 8);
  setFontRange(ws, startRow, 1, startRow, 8, { bold: true });

  ws.mergeCells(`A${dataStart}:A${dataEnd}`);
  ws.mergeCells(`B${dataStart}:B${dataEnd}`);
  ws.mergeCells(`C${dataStart}:C${dataEnd}`);
  ws.mergeCells(`D${dataStart}:D${dataEnd}`);

  setCell(ws, `A${dataStart}`, st.data.itemText || "");
  setCell(ws, `B${dataStart}`, "");
  setCell(ws, `C${dataStart}`, st.data.color || "");
  addPastedImageToHT(ws, st, dataStart, dataEnd);
  setCell(ws, `D${dataStart}`, st.data.content || "");

  ws.getCell(`D${dataStart}`).font = {
    bold: true,
    color: { argb: "FFFF0000" }
  };

  centerRange(ws, dataStart, 1, dataEnd, 4);
  setFontRange(ws, dataStart, 1, dataEnd, 4, { bold: true, size: 12 });

  for (let i = 0; i < sizes.length; i++) {
    const r = dataStart + i;
    const size = sizes[i];
    const qty = Number(qtyMap[size] || 0);

    setCell(ws, `E${r}`, size);
    setCell(ws, `F${r}`, qty);
    setCell(ws, `G${r}`, null);
    ws.getCell(`H${r}`).value = { formula: `F${r}*G${r}` };
    ws.getCell(`H${r}`).numFmt = '$#,##0.00';

    centerRange(ws, r, 5, r, 8);
    setFontRange(ws, r, 5, r, 8, { bold: true });
  }

  fillRange(ws, totalRow, 1, totalRow, 8, "FFFFFF00");
  setCell(ws, `B${totalRow}`, st.data.refText || "");
  setCell(ws, `E${totalRow}`, "TOTAL");
  ws.getCell(`F${totalRow}`).value = { formula: `SUM(F${dataStart}:F${dataEnd})` };
  ws.getCell(`H${totalRow}`).value = { formula: `SUM(H${dataStart}:H${dataEnd})` };
  ws.getCell(`H${totalRow}`).numFmt = '$#,##0.00';

  centerRange(ws, totalRow, 1, totalRow, 8);
  setFontRange(ws, totalRow, 1, totalRow, 8, { bold: true });

  applyBorderRange(ws, startRow, 1, totalRow, 8, thinBorder());

  return totalRow;
}

// =========================
// Bloques SS
// =========================
function writeSSBlock(ws, startRow, st) {
  const sizes = st.data.baseFiltered.slice();
  const qtyMap = { ...st.data.qtyBase };

  const dataStart = startRow + 1;
  const dataEnd = dataStart + sizes.length - 1;
  const totalRow = dataEnd + 1;

  setCell(ws, `A${startRow}`, "SIZE STRIP");
  setCell(ws, `B${startRow}`, "ITEM #");
  setCell(ws, `C${startRow}`, "STYLE");
  setCell(ws, `D${startRow}`, "SIZE");
  setCell(ws, `E${startRow}`, "QTY");
  setCell(ws, `F${startRow}`, "UPPR");
  setCell(ws, `G${startRow}`, "TTL $");

  fillRange(ws, startRow, 1, startRow, 7, "FFD9D9D9");
  centerRange(ws, startRow, 1, startRow, 7);
  setFontRange(ws, startRow, 1, startRow, 7, { bold: true });

  ws.mergeCells(`A${dataStart}:A${dataEnd}`);
  ws.mergeCells(`B${dataStart}:B${dataEnd}`);
  ws.mergeCells(`C${dataStart}:C${dataEnd}`);

  setCell(ws, `A${dataStart}`, st.data.itemText || "");
  setCell(ws, `B${dataStart}`, "");
  setCell(ws, `C${dataStart}`, "");

  centerRange(ws, dataStart, 1, dataEnd, 3);
  setFontRange(ws, dataStart, 1, dataEnd, 3, { bold: true, size: 12 });

  for (let i = 0; i < sizes.length; i++) {
    const r = dataStart + i;
    const size = sizes[i];
    const qty = Number(qtyMap[size] || 0);

    setCell(ws, `D${r}`, size);
    setCell(ws, `E${r}`, qty);
    setCell(ws, `F${r}`, null);
    ws.getCell(`G${r}`).value = { formula: `E${r}*F${r}` };
    ws.getCell(`G${r}`).numFmt = '$#,##0.00';

    centerRange(ws, r, 4, r, 7);
    setFontRange(ws, r, 4, r, 7, { bold: true });
  }

  fillRange(ws, totalRow, 1, totalRow, 7, "FFFFFF00");
  setCell(ws, `B${totalRow}`, st.data.refText || "");
  setCell(ws, `D${totalRow}`, "TOTAL");
  ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${dataStart}:E${dataEnd})` };
  ws.getCell(`G${totalRow}`).value = { formula: `SUM(G${dataStart}:G${dataEnd})` };
  ws.getCell(`G${totalRow}`).numFmt = '$#,##0.00';

  centerRange(ws, totalRow, 1, totalRow, 7);
  setFontRange(ws, totalRow, 1, totalRow, 7, { bold: true });

  applyBorderRange(ws, startRow, 1, totalRow, 7, thinBorder());

  return totalRow;
}

// =========================
// Bloque Color Tag nuevo
// =========================
async function writeColorTagBlock(ws, startRow, st) {
  const dataStart = startRow + 1;
  const dataEnd = dataStart + 4;
  const totalRow = dataEnd + 1;
  const cfg = getColorTagConfig(st.data.colorTagColor);

  ws.getRow(startRow).height = 26.25;
  for (let r = dataStart; r <= dataEnd; r++) {
    ws.getRow(r).height = 47.25;
  }
  ws.getRow(totalRow).height = 35.25;

  setCell(ws, `A${startRow}`, "COLOR TAG");
  setCell(ws, `B${startRow}`, "ITEM #");
  setCell(ws, `C${startRow}`, "DESCRIPTION");
  setCell(ws, `D${startRow}`, "TAG");
  setCell(ws, `E${startRow}`, "QTY");
  setCell(ws, `F${startRow}`, "U / PRC");
  setCell(ws, `G${startRow}`, "TTL $");

  fillRange(ws, startRow, 1, startRow, 7, "FFD9D9D9");
  centerRange(ws, startRow, 1, startRow, 7);
  setFontRange(ws, startRow, 1, startRow, 7, { name: "Tahoma", size: 18, bold: true });
  applyBorderRange(ws, startRow, 1, startRow, 7, thinBorder());
  setOuterBorder(ws, startRow, 1, startRow, 7, "medium");

  ws.mergeCells(`A${dataStart}:A${dataEnd}`);
  ws.mergeCells(`B${dataStart}:B${dataEnd}`);
  ws.mergeCells(`C${dataStart}:C${dataEnd}`);
  ws.mergeCells(`D${dataStart}:D${dataEnd}`);
  ws.mergeCells(`E${dataStart}:E${dataEnd}`);
  ws.mergeCells(`F${dataStart}:F${dataEnd}`);
  ws.mergeCells(`G${dataStart}:G${dataEnd}`);

  setCell(ws, `A${dataStart}`, "");
  setCell(ws, `B${dataStart}`, st.data.itemText || "");
  setCell(ws, `C${dataStart}`, cfg.description);
  setCell(ws, `D${dataStart}`, "");
  setCell(ws, `E${dataStart}`, Number(st.data.colorTagQty || 0));
  setCell(ws, `F${dataStart}`, 0.0090);
  ws.getCell(`G${dataStart}`).value = { formula: `E${dataStart}*F${dataStart}` };

  centerRange(ws, dataStart, 1, dataEnd, 7);
  applyBorderRange(ws, dataStart, 1, dataEnd, 7, thinBorder());
  setOuterBorder(ws, dataStart, 1, dataEnd, 7, "medium");

  setFontRange(ws, dataStart, 1, dataEnd, 2, { name: "Tahoma", size: 18, bold: true });
  setFontRange(ws, dataStart, 3, dataEnd, 3, { name: "Tahoma", size: 16, bold: true });
  setFontRange(ws, dataStart, 4, dataEnd, 7, { name: "Tahoma", size: 20, bold: true });

  ws.getCell(`F${dataStart}`).numFmt = '$#,##0.0000';
  ws.getCell(`G${dataStart}`).numFmt = '$#,##0.00';

  fillRange(ws, totalRow, 1, totalRow, 3, "FFFFFF00");
  centerRange(ws, totalRow, 1, totalRow, 7);
  applyBorderRange(ws, totalRow, 1, totalRow, 7, thinBorder());
  setOuterBorder(ws, totalRow, 1, totalRow, 7, "medium");
  setFontRange(ws, totalRow, 1, totalRow, 7, { name: "Tahoma", size: 18, bold: true });

  setCell(ws, `B${totalRow}`, st.data.refText || "");
  setCell(ws, `D${totalRow}`, "TOTAL");
  ws.getCell(`E${totalRow}`).value = { formula: `E${dataStart}` };
  ws.getCell(`G${totalRow}`).value = { formula: `G${dataStart}` };
  ws.getCell(`G${totalRow}`).numFmt = '$#,##0.00';

  const base64 = await fetchAssetAsBase64(cfg.imagePath);
  const imageId = ws.workbook.addImage({
    base64,
    extension: "png"
  });

  // Imagen centrada y con tamaño normal dentro del bloque TAG
  ws.addImage(imageId, {
  tl: { col: 3.7, row: (dataStart - 1) + 0.28 },
    ext: { width: 113, height: 222 },
    editAs: "oneCell"
  });

  return totalRow;
}

// =========================
// Bloques Hang Tag
// =========================
function writeHangTagBlock(ws, startRow, st, title) {
  const totalQty = Number(st.data.hangTagQty || 0);

  const dataStart = startRow + 1;
  const dataEnd = dataStart + 3;
  const totalRow = dataEnd + 1;

  setCell(ws, `A${startRow}`, "PO:");
  setCell(ws, `B${startRow}`, "ITEM #");
  setCell(ws, `C${startRow}`, title);
  setCell(ws, `D${startRow}`, "QTY");
  setCell(ws, `E${startRow}`, "UPPRC");
  setCell(ws, `F${startRow}`, "TTL $");

  fillRange(ws, startRow, 1, startRow, 6, "FFD9D9D9");
  centerRange(ws, startRow, 1, startRow, 6);
  setFontRange(ws, startRow, 1, startRow, 6, { bold: true });

  ws.mergeCells(`A${dataStart}:A${dataEnd}`);
  ws.mergeCells(`B${dataStart}:B${dataEnd}`);
  ws.mergeCells(`C${dataStart}:C${dataEnd}`);
  ws.mergeCells(`D${dataStart}:D${dataEnd}`);
  ws.mergeCells(`E${dataStart}:E${dataEnd}`);
  ws.mergeCells(`F${dataStart}:F${dataEnd}`);

  setCell(ws, `A${dataStart}`, st.data.itemText || "");
  setCell(ws, `B${dataStart}`, "");
  setCell(ws, `C${dataStart}`, "");
  setCell(ws, `D${dataStart}`, totalQty || 0);
  setCell(ws, `E${dataStart}`, null);
  ws.getCell(`F${dataStart}`).value = { formula: `D${dataStart}*E${dataStart}` };
  ws.getCell(`F${dataStart}`).numFmt = '$#,##0.00';

  centerRange(ws, dataStart, 1, dataEnd, 6);
  setFontRange(ws, dataStart, 1, dataEnd, 6, { bold: true });

  fillRange(ws, totalRow, 1, totalRow, 6, "FFFFFF00");
  setCell(ws, `B${totalRow}`, st.data.refText || "");
  ws.getCell(`D${totalRow}`).value = { formula: `D${dataStart}` };
  ws.getCell(`F${totalRow}`).value = { formula: `F${dataStart}` };
  ws.getCell(`F${totalRow}`).numFmt = '$#,##0.00';

  centerRange(ws, totalRow, 1, totalRow, 6);
  setFontRange(ws, totalRow, 1, totalRow, 6, { bold: true });

  applyBorderRange(ws, startRow, 1, totalRow, 6, thinBorder());

  return totalRow;
}

function addPastedImageToHT(ws, st, dataStart, dataEnd) {
  if (!st?.data?.imageBase64) return;
  const imageId = ws.workbook.addImage({
    base64: st.data.imageBase64,
    extension: st.data.imageExtension || "png"
  });

  ws.addImage(imageId, {
    tl: { col: 1.08, row: (dataStart - 1) + 0.16 },
    ext: { width: 92, height: Math.max(72, ((dataEnd - dataStart) + 1) * 26) },
    editAs: "oneCell"
  });
}

// =========================
// Workbooks
// =========================
async function buildHeatTransferWorkbook(payload) {
  const wb = new ExcelJS.Workbook();
  setWorkbookMeta(wb);
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("HT");
  setupHeader8(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeHTBlock(ws, row, st);
    totalRows.push(endRow);
    row = endRow + 2;
  }

  if (totalRows.length > 0) {
    const generalRow = row;
    ws.mergeCells(`A${generalRow}:E${generalRow}`);
    ws.getCell(`A${generalRow}`).value = "TOTAL GENERAL:";
    ws.getCell(`F${generalRow}`).value = {
      formula: totalRows.map(r => `F${r}`).join("+")
    };
    ws.getCell(`H${generalRow}`).value = {
      formula: totalRows.map(r => `H${r}`).join("+")
    };
    ws.getCell(`H${generalRow}`).numFmt = '$#,##0.00';

    fillRange(ws, generalRow, 1, generalRow, 8, "FFFFFF00");
    centerRange(ws, generalRow, 1, generalRow, 8);
    setFontRange(ws, generalRow, 1, generalRow, 8, { bold: true });
    applyBorderRange(ws, generalRow, 1, generalRow, 8, mediumBorder());
  }

  return wb;
}

async function buildSizeStripWorkbook(payload) {
  const wb = new ExcelJS.Workbook();
  setWorkbookMeta(wb);
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("SS");
  setupHeader7(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeSSBlock(ws, row, st);
    totalRows.push(endRow);
    row = endRow + 2;
  }

  if (totalRows.length > 0) {
    const generalRow = row;
    ws.mergeCells(`A${generalRow}:D${generalRow}`);
    ws.getCell(`A${generalRow}`).value = "TOTAL GENERAL:";
    ws.getCell(`E${generalRow}`).value = {
      formula: totalRows.map(r => `E${r}`).join("+")
    };
    ws.getCell(`G${generalRow}`).value = {
      formula: totalRows.map(r => `G${r}`).join("+")
    };
    ws.getCell(`G${generalRow}`).numFmt = '$#,##0.00';

    fillRange(ws, generalRow, 1, generalRow, 7, "FFFFFF00");
    centerRange(ws, generalRow, 1, generalRow, 7);
    setFontRange(ws, generalRow, 1, generalRow, 7, { bold: true });
    applyBorderRange(ws, generalRow, 1, generalRow, 7, mediumBorder());
  }

  return wb;
}

async function buildColorTagWorkbook(payload) {
  const wb = new ExcelJS.Workbook();
  setWorkbookMeta(wb);
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("CL");
  setupColorTagSheet(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = await writeColorTagBlock(ws, row, st);
    totalRows.push(endRow);

    const blankRow = endRow + 1;
    ws.getRow(blankRow).height = 21;
    row = endRow + 2;
  }

  if (totalRows.length > 0) {
    const generalRow = row;
    ws.getRow(generalRow).height = 48;

    ws.mergeCells(`A${generalRow}:D${generalRow}`);
    ws.getCell(`A${generalRow}`).value = "TOTAL GENERAL:";
    ws.getCell(`E${generalRow}`).value = {
      formula: totalRows.map(r => `E${r}`).join("+")
    };
    ws.getCell(`G${generalRow}`).value = {
      formula: totalRows.map(r => `G${r}`).join("+")
    };
    ws.getCell(`G${generalRow}`).numFmt = '$#,##0.00';

    fillRange(ws, generalRow, 1, generalRow, 4, "FFFFFF00");
    centerRange(ws, generalRow, 1, generalRow, 7);
    applyBorderRange(ws, generalRow, 1, generalRow, 7, thinBorder());
    setOuterBorder(ws, generalRow, 1, generalRow, 7, "medium");

    setFontRange(ws, generalRow, 1, generalRow, 4, { name: "Arial", size: 22, bold: true });
    setFontRange(ws, generalRow, 5, generalRow, 7, { name: "Tahoma", size: 20, bold: true });
  }

  return wb;
}

async function buildHangTagWorkbook(payload) {
  const wb = new ExcelJS.Workbook();
  setWorkbookMeta(wb);
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("TAG");
  setupHeader6(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeHangTagBlock(ws, row, st, "HANG TAG");
    totalRows.push(endRow);
    row = endRow + 2;
  }

  if (totalRows.length > 0) {
    const generalRow = row;
    ws.mergeCells(`A${generalRow}:C${generalRow}`);
    ws.getCell(`A${generalRow}`).value = "TOTAL GENERAL:";
    ws.getCell(`D${generalRow}`).value = {
      formula: totalRows.map(r => `D${r}`).join("+")
    };
    ws.getCell(`F${generalRow}`).value = {
      formula: totalRows.map(r => `F${r}`).join("+")
    };
    ws.getCell(`F${generalRow}`).numFmt = '$#,##0.00';

    fillRange(ws, generalRow, 1, generalRow, 6, "FFFFFF00");
    centerRange(ws, generalRow, 1, generalRow, 6);
    setFontRange(ws, generalRow, 1, generalRow, 6, { bold: true });
    applyBorderRange(ws, generalRow, 1, generalRow, 6, mediumBorder());
  }

  return wb;
}

async function downloadWorkbook(wb, filename) {
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  forceDownload(blob, filename);
}

// =========================
// Init
// =========================
(function initDates() {
  el("fechaPo").value = todayYMD();
  el("fechaProd").value = addDays(el("fechaPo").value, 7);

  el("fechaPo").addEventListener("change", () => {
    el("fechaProd").value = addDays(el("fechaPo").value, 7);
  });
})();

(function initStyleCount() {
  const qty = el("qtyStyles");

  function update() {
    let n = parseInt(qty?.value, 10);
    if (!Number.isFinite(n) || n <= 0) n = 1;
    if (n > 20) n = 20;

    renderStyleNameInputs(n);
    renderStyles(n);
  }

  qty?.addEventListener("input", update);
  update();
})();

(function initBuyChecks() {
  ["buyHT", "buySS", "buyCT", "buyHG"].forEach(id => {
    el(id)?.addEventListener("change", () => {
      getSelectedCompras();
      refreshAllStyleOptionVisibility();
    });
  });

  getSelectedCompras();
  refreshAllStyleOptionVisibility();
})();

// =========================
// Descargar
// =========================
el("download").addEventListener("click", async () => {
  try {
    const compras = getSelectedCompras();

    if (compras.length === 0) {
      alert("Selecciona al menos una opción.");
      return;
    }

    const poBase = poDigitsOnly(el("po")?.value || "");
    if (!poBase) {
      alert("Ingresa el PO #.");
      return;
    }

    const fechaPoYMD = el("fechaPo").value || todayYMD();
    const fechaProdYMD = el("fechaProd").value || addDays(fechaPoYMD, 7);
    const lugar = (el("lugar").value || "").trim() || "PENDIENTE";

    let n = parseInt(el("qtyStyles")?.value, 10);
    if (!Number.isFinite(n) || n <= 0) n = 1;
    if (n > 20) n = 20;

    const allStyles = [];
    for (let i = 1; i <= n; i++) {
      allStyles.push({
        idx: i,
        data: getStyleData(i)
      });
    }

    const htStyles = allStyles.filter(st => hasHeatTransferQty(st.data));
    const ssStyles = allStyles.filter(st => hasSizeStripQty(st.data));
    const ctStyles = allStyles.filter(st => hasColorTagQty(st.data));
    const hgStyles = allStyles.filter(st => hasHangTagQty(st.data));

    const missing = [];
    let generatedAny = false;

    if (compras.includes("heat_transfer")) {
      if (htStyles.length === 0) {
        missing.push("Heat Transfer: ingresa cantidades por talla.");
      } else {
        const wbHT = await buildHeatTransferWorkbook({
          poText: `${poBase}-HT`,
          fechaPoYMD,
          fechaProdYMD,
          lugar,
          styles: htStyles
        });
        await downloadWorkbook(wbHT, safeFileName(`${poBase}-HT.xlsx`));
        generatedAny = true;
        await sleep(350);
      }
    }

    if (compras.includes("size_strip")) {
      if (ssStyles.length === 0) {
        missing.push("Size Strip: ingresa cantidades por talla.");
      } else {
        const wbSS = await buildSizeStripWorkbook({
          poText: `${poBase}-SS`,
          fechaPoYMD,
          fechaProdYMD,
          lugar,
          styles: ssStyles
        });
        await downloadWorkbook(wbSS, safeFileName(`${poBase}-SS.xlsx`));
        generatedAny = true;
        await sleep(350);
      }
    }

    if (compras.includes("color_tag")) {
      if (ctStyles.length === 0) {
        missing.push("Color Tag: ingresa la cantidad manual por estilo.");
      } else {
        const wbCT = await buildColorTagWorkbook({
          poText: `${poBase}-CL`,
          fechaPoYMD,
          fechaProdYMD,
          lugar,
          styles: ctStyles
        });
        await downloadWorkbook(wbCT, safeFileName(`${poBase}-CL.xlsx`));
        generatedAny = true;
        await sleep(350);
      }
    }

    if (compras.includes("hang_tag")) {
      if (hgStyles.length === 0) {
        missing.push("Hang Tag: ingresa la cantidad manual por estilo.");
      } else {
        const wbHG = await buildHangTagWorkbook({
          poText: `${poBase}-TAG`,
          fechaPoYMD,
          fechaProdYMD,
          lugar,
          styles: hgStyles
        });
        await downloadWorkbook(wbHG, safeFileName(`${poBase}-TAG.xlsx`));
        generatedAny = true;
        await sleep(350);
      }
    }

    if (!generatedAny) {
      alert(missing.join("\n") || "No hay datos suficientes para generar el archivo.");
      return;
    }

    if (missing.length > 0) {
      alert(`Se generaron los archivos disponibles.\n\nPendiente de completar:\n- ${missing.join("\n- ")}`);
    }

  } catch (err) {
    console.error(err);
    alert(err?.message || String(err));
  }
});