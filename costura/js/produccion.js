const el = (id) => document.getElementById(id);

const PRICE_RULES = {
  TH_5000: 1.32,
  TH_2500: 0.66,
  THX_5000: 0.95,
  THX_2500: 0.475
};

const HEADER_IMAGE_PATH = "assets/kmh_header_clean.png";
const state = { styles: [], expanded: {} };
let lastStyleCount = 1;

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

function formatNumber(value, digits = 2) {
  const num = Number(value);
  if (!Number.isFinite(num)) return digits === 0 ? "0" : "0.00";
  return num.toLocaleString("es-GT", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function formatCurrency(value, digits = 2) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "$0.00";
  return num.toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
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

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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

async function fetchAssetAsBase64(path) {
  const res = await fetch(path);
  if (!res.ok) throw new Error(`No se pudo cargar el recurso ${path}`);
  return await blobToDataURL(await res.blob());
}

function setTodayDefaults() {
  const today = todayYMD();
  el("fechaPo").value = today;
  el("fechaProd").value = addDays(today, 7);
}

function normalizeStyleCount(raw, fallback = 1) {
  const value = parseInt(String(raw || "").trim(), 10);
  if (!Number.isFinite(value)) return fallback;
  return Math.min(20, Math.max(1, value));
}

function getStyleCountFromInput(fallback = lastStyleCount) {
  const raw = el("qtyStyles")?.value || "";
  return normalizeStyleCount(raw, fallback || 1);
}

function captureCurrentStyleValues() {
  const cards = [...document.querySelectorAll(".style-entry")];
  return cards.map((card) => {
    const index = Number(card.dataset.index);
    return {
      styleRef: el(`styleRef_${index}`)?.value || "",
      color: el(`color_${index}`)?.value || "",
      pieces: el(`pieces_${index}`)?.value || "",
      consumo: el(`consumo_${index}`)?.value || "",
      approvalType: el(`approvalType_${index}`)?.value || "igualar",
      matchPo: el(`matchPo_${index}`)?.value || "",
      matchColor: el(`matchColor_${index}`)?.value || ""
    };
  });
}

function ensureStateSize(n) {
  while (state.styles.length < n) {
    state.styles.push({ imageBase64: null, imageExtension: null });
  }
  if (state.styles.length > n) {
    state.styles = state.styles.slice(0, n);
  }

  for (let index = 1; index <= n; index += 1) {
    if (!(index in state.expanded)) {
      state.expanded[index] = false;
    }
  }

  Object.keys(state.expanded).forEach((key) => {
    if (Number(key) > n) delete state.expanded[key];
  });
}

function applyPreservedValues(values = []) {
  values.forEach((item, idx) => {
    const index = idx + 1;
    if (!el(`styleRef_${index}`)) return;
    el(`styleRef_${index}`).value = item.styleRef || "";
    el(`color_${index}`).value = item.color || "";
    el(`pieces_${index}`).value = item.pieces || "";
    el(`consumo_${index}`).value = item.consumo || "";
    el(`approvalType_${index}`).value = item.approvalType || "igualar";
    el(`matchPo_${index}`).value = item.matchPo || "";
    el(`matchColor_${index}`).value = item.matchColor || "";
    toggleApprovalFields(index);
  });

  state.styles.forEach((_, idx) => updatePastePreview(idx + 1));
}

function renderStyleCards(count) {
  const preserved = captureCurrentStyleValues();
  ensureStateSize(count);

  const host = el("stylesContainer");
  host.innerHTML = "";

  for (let index = 1; index <= count; index += 1) {
    const previewExpanded = !!state.expanded[index];
    const article = document.createElement("article");
    article.className = "style-entry glass-subpanel";
    article.dataset.index = index;
    article.innerHTML = `
      <div class="style-entry__head">
        <div>
          <small>Estilo #${index}</small>
          <h3>Información del estilo</h3>
        </div>
        <div class="style-entry__actions">
          <span class="style-entry__badge">Se coloca en la misma hoja</span>
          <button class="style-entry__toggle" id="toggleStyle_${index}" type="button">${previewExpanded ? "Ocultar cálculo" : "Ver cálculo del estilo"}</button>
        </div>
      </div>

      <div class="style-entry__body" id="styleBody_${index}">
        <div class="style-grid">
          <label>
            <span>Style</span>
            <input id="styleRef_${index}" type="text" placeholder="Ej. TSBAGV3FNF" />
          </label>

          <label>
            <span>Color</span>
            <input id="color_${index}" type="text" placeholder="Ej. CASTLEROCK / verde" />
          </label>

          <label>
            <span>Cantidad de piezas</span>
            <input id="pieces_${index}" type="number" min="0" step="1" placeholder="0" />
          </label>

          <label>
            <span>Consumo de hilo (cm por pieza)</span>
            <input id="consumo_${index}" type="number" min="0" step="0.01" placeholder="0.00" />
          </label>

          <label>
            <span>Tipo de aprobación</span>
            <select id="approvalType_${index}">
              <option value="igualar" selected>Igualar PO</option>
              <option value="muestra">Se mandará muestra de tela</option>
              <option value="pendiente">Pendiente muestra de tela o igualación</option>
            </select>
          </label>

          <label id="matchPoWrap_${index}">
            <span>PO a igualar</span>
            <input id="matchPo_${index}" type="text" inputmode="numeric" placeholder="Ej. 11485" />
          </label>

          <label id="matchColorWrap_${index}">
            <span>Color de esa PO</span>
            <input id="matchColor_${index}" type="text" placeholder="Ej. CASTLEROCK" />
          </label>

          <div class="paste-wrap paste-wrap--compact">
            <span class="paste-wrap__label">Imagen del item (pegar desde portapapeles)</span>
            <div class="paste-zone" id="pasteZone_${index}" tabindex="0" role="button" aria-label="Pega una imagen para el estilo ${index}">
              <div class="paste-zone__empty" id="pasteEmpty_${index}">
                <i class="fa-regular fa-image"></i>
                <strong>Pega aquí la imagen</strong>
                <small>Haz clic en esta zona y usa Ctrl + V</small>
              </div>
              <img id="pastePreview_${index}" class="paste-zone__preview hidden" alt="Vista previa del item ${index}" />
              <button id="clearPaste_${index}" class="paste-zone__clear hidden" type="button">Limpiar</button>
            </div>
          </div>

          <div class="style-preview">
            <div class="style-preview__head">
              <strong>Vista de cálculo del estilo</strong>
              <button class="style-preview__toggle" id="togglePreview_${index}" type="button">${previewExpanded ? "Ocultar cálculo" : "Ver cálculo del estilo"}</button>
            </div>
            <div class="style-preview__body${previewExpanded ? "" : " hidden"}" id="previewBody_${index}">
              <div class="style-mini-summary" id="previewSummary_${index}"></div>
              <div class="style-mini-plan" id="previewPlan_${index}"></div>
            </div>
          </div>
        </div>
      </div>
    `;
    host.appendChild(article);
  }

  bindStyleEvents(count);
  applyPreservedValues(preserved);
  renderPlan();
}

function bindStyleEvents(count) {
  for (let index = 1; index <= count; index += 1) {
    ["styleRef", "color", "pieces", "consumo", "matchPo", "matchColor"].forEach((name) => {
      el(`${name}_${index}`).addEventListener("input", renderPlan);
    });

    el(`approvalType_${index}`).addEventListener("change", () => {
      toggleApprovalFields(index);
      renderPlan();
    });

    el(`togglePreview_${index}`).addEventListener("click", () => toggleStyleCard(index));
    if (el(`toggleStyle_${index}`)) {
      el(`toggleStyle_${index}`).addEventListener("click", () => toggleStyleCard(index));
    }

    const zone = el(`pasteZone_${index}`);
    zone.addEventListener("paste", (event) => handlePaste(event, index));
    zone.addEventListener("click", () => zone.focus());
    el(`clearPaste_${index}`).addEventListener("click", (event) => {
      event.stopPropagation();
      clearPastedImage(index);
      renderPlan();
    });
  }
}

function toggleApprovalFields(index) {
  const isEqual = el(`approvalType_${index}`).value === "igualar";
  el(`matchPoWrap_${index}`).classList.toggle("hidden", !isEqual);
  el(`matchColorWrap_${index}`).classList.toggle("hidden", !isEqual);
}

function toggleStyleCard(index, force) {
  const next = typeof force === "boolean" ? force : !state.expanded[index];
  state.expanded[index] = next;
  const body = el(`previewBody_${index}`);
  if (body) body.classList.toggle("hidden", !next);
  [el(`togglePreview_${index}`), el(`toggleStyle_${index}`)].forEach((btn) => {
    if (btn) btn.textContent = next ? "Ocultar cálculo" : "Ver cálculo del estilo";
  });
}

async function handlePaste(event, index) {
  const items = [...(event.clipboardData?.items || [])];
  const imageItem = items.find((item) => item.type.startsWith("image/"));
  if (!imageItem) return;
  event.preventDefault();

  const file = imageItem.getAsFile();
  if (!file) return;

  const base64 = await blobToDataURL(file);
  const extension = file.type.includes("png")
    ? "png"
    : file.type.includes("webp")
      ? "webp"
      : "jpeg";

  state.styles[index - 1] = {
    imageBase64: base64,
    imageExtension: extension
  };

  updatePastePreview(index);
  renderPlan();
}

function clearPastedImage(index) {
  state.styles[index - 1] = { imageBase64: null, imageExtension: null };
  updatePastePreview(index);
}

function updatePastePreview(index) {
  const imageState = state.styles[index - 1] || { imageBase64: null, imageExtension: null };
  const preview = el(`pastePreview_${index}`);
  const empty = el(`pasteEmpty_${index}`);
  const clearBtn = el(`clearPaste_${index}`);

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

function getStyleEntries() {
  const count = getStyleCountFromInput(lastStyleCount);
  const entries = [];

  for (let index = 1; index <= count; index += 1) {
    const styleRef = (el(`styleRef_${index}`)?.value || "").trim();
    const color = (el(`color_${index}`)?.value || "").trim();
    const piecesRaw = el(`pieces_${index}`)?.value;
    const consumoRaw = el(`consumo_${index}`)?.value;
    const pieces = parseNumber(piecesRaw);
    const consumoCm = parseNumber(consumoRaw);
    const approvalType = el(`approvalType_${index}`)?.value || "igualar";
    const matchPo = poDigitsOnly(el(`matchPo_${index}`)?.value);
    const matchColor = (el(`matchColor_${index}`)?.value || "").trim();
    const totalMeters = (pieces * consumoCm) / 100;
    const imageState = state.styles[index - 1] || { imageBase64: null, imageExtension: null };

    const hasAny = [styleRef, color, piecesRaw, consumoRaw, matchPo, matchColor].some((v) => String(v || "").trim() !== "") || !!imageState.imageBase64;
    const isValid = color && pieces > 0 && consumoCm > 0;

    entries.push({
      index,
      styleRef,
      color,
      pieces,
      consumoCm,
      totalMeters,
      thMeters: totalMeters * 0.7,
      thxMeters: totalMeters * 0.3,
      approvalType,
      matchPo,
      matchColor,
      imageBase64: imageState.imageBase64,
      imageExtension: imageState.imageExtension,
      hasAny,
      isValid
    });
  }

  return entries;
}
function validateEntries(entries) {
  for (const entry of entries) {
    if (entry.hasAny && !entry.isValid) {
      return `El estilo #${entry.index} está incompleto. Necesita color, piezas y consumo.`;
    }
    if (entry.isValid && entry.approvalType === "igualar" && (!entry.matchPo || !entry.matchColor)) {
      return `El estilo #${entry.index} necesita PO y color de referencia para "Igualar PO".`;
    }
  }
  if (!entries.some((entry) => entry.isValid)) {
    return "Necesito al menos un estilo válido para generar el Excel.";
  }
  return "";
}

function distributeQty(totalQty, styles, metersKey) {
  const positive = styles.filter((style) => style[metersKey] > 0);
  const result = new Map(styles.map((style) => [style.index, 0]));
  if (!positive.length || totalQty <= 0) return result;

  const totalMeters = positive.reduce((sum, style) => sum + style[metersKey], 0);
  const raw = positive.map((style) => {
    const share = (style[metersKey] / totalMeters) * totalQty;
    const base = Math.floor(share);
    return {
      index: style.index,
      base,
      frac: share - base
    };
  });

  let assigned = raw.reduce((sum, item) => sum + item.base, 0);
  raw.sort((a, b) => b.frac - a.frac || a.index - b.index);

  for (let i = 0; assigned < totalQty; i += 1) {
    raw[i % raw.length].base += 1;
    assigned += 1;
  }

  raw.forEach((item) => result.set(item.index, item.base));
  return result;
}

function buildPurchasePlan(entries = getStyleEntries()) {
  const validStyles = entries.filter((entry) => entry.isValid);
  if (!validStyles.length) return [];

  const totalThMeters = validStyles.reduce((sum, entry) => sum + entry.thMeters, 0);
  const totalThxMeters = validStyles.reduce((sum, entry) => sum + entry.thxMeters, 0);
  const thQty5000Total = Math.round(totalThMeters / 50);
  const thQty2500Total = Math.round(totalThMeters / 25);
  const thxQty5000Total = Math.round(totalThxMeters / 50);
  const isMinimumMode = thQty5000Total <= 30;

  if (isMinimumMode) {
    const totalQty = Math.max(thQty2500Total, 60);
    const shares = distributeQty(totalQty, validStyles, "thMeters");
    const rows = validStyles.map((style) => {
      const qty = shares.get(style.index) || 0;
      return {
        ...style,
        qty,
        unitPrice: PRICE_RULES.TH_2500,
        amount: qty * PRICE_RULES.TH_2500,
        sizeLabel: "2,500 MTS",
        threadLabel: "TEX 27"
      };
    });

    return [{
      suffix: "TH",
      threadLabel: "TEX 27",
      sizeLabel: "2,500 MTS",
      totalQty,
      unitPrice: PRICE_RULES.TH_2500,
      totalAmount: totalQty * PRICE_RULES.TH_2500,
      totalMeters: totalThMeters,
      minMode: true,
      rows
    }];
  }

  const thShares = distributeQty(thQty5000Total, validStyles, "thMeters");
  const thxShares = distributeQty(thxQty5000Total, validStyles, "thxMeters");

  const thRows = validStyles.map((style) => {
    const qty = thShares.get(style.index) || 0;
    return {
      ...style,
      qty,
      unitPrice: PRICE_RULES.TH_5000,
      amount: qty * PRICE_RULES.TH_5000,
      sizeLabel: "5,000 MTS",
      threadLabel: "TEX 27"
    };
  });

  const thxRows = validStyles.map((style) => {
    const qty = thxShares.get(style.index) || 0;
    return {
      ...style,
      qty,
      unitPrice: PRICE_RULES.THX_5000,
      amount: qty * PRICE_RULES.THX_5000,
      sizeLabel: "5,000 MTS",
      threadLabel: "40/2"
    };
  });

  return [
    {
      suffix: "TH",
      threadLabel: "TEX 27",
      sizeLabel: "5,000 MTS",
      totalQty: thQty5000Total,
      unitPrice: PRICE_RULES.TH_5000,
      totalAmount: thQty5000Total * PRICE_RULES.TH_5000,
      totalMeters: totalThMeters,
      minMode: false,
      rows: thRows
    },
    {
      suffix: "THX",
      threadLabel: "40/2",
      sizeLabel: "5,000 MTS",
      totalQty: thxQty5000Total,
      unitPrice: PRICE_RULES.THX_5000,
      totalAmount: thxQty5000Total * PRICE_RULES.THX_5000,
      totalMeters: totalThxMeters,
      minMode: false,
      rows: thxRows
    }
  ];
}

function getOrderData() {
  const entries = getStyleEntries();
  const validStyles = entries.filter((entry) => entry.isValid);
  return {
    poBase: poDigitsOnly(el("po").value),
    fechaPo: el("fechaPo").value || todayYMD(),
    fechaProd: el("fechaProd").value || addDays(el("fechaPo").value || todayYMD(), 7),
    lugar: (el("lugar").value || "PENDIENTE").trim() || "PENDIENTE",
    entries,
    validStyles,
    totalMeters: validStyles.reduce((sum, entry) => sum + entry.totalMeters, 0),
    totalThMeters: validStyles.reduce((sum, entry) => sum + entry.thMeters, 0),
    totalThxMeters: validStyles.reduce((sum, entry) => sum + entry.thxMeters, 0)
  };
}

function renderStylePreview(entry, stylePlans) {
  const summaryHost = el(`previewSummary_${entry.index}`);
  const planHost = el(`previewPlan_${entry.index}`);
  if (!summaryHost || !planHost) return;

  if (!entry.hasAny) {
    summaryHost.innerHTML = "";
    planHost.innerHTML = '<div class="style-mini-placeholder">Este estilo todavía está vacío. Cuando coloques color, piezas y consumo, aquí aparecerá el cálculo.</div>';
    return;
  }

  if (!entry.isValid) {
    summaryHost.innerHTML = "";
    planHost.innerHTML = '<div class="style-mini-placeholder">Faltan datos en este estilo. Completa color, cantidad de piezas y consumo de hilo para ver el cálculo.</div>';
    return;
  }

  const visiblePlans = stylePlans.filter((item) => (item.row?.qty || 0) > 0 || item.usedMeters > 0);
  const modeText = visiblePlans.length === 1 && visiblePlans[0].minMode ? "TH mínimo" : "TH + THX";
  const modeHint = visiblePlans.length === 1 && visiblePlans[0].minMode
    ? "Solo TH en 2,500 MTS"
    : "Compra normal en 5,000 MTS";

  summaryHost.innerHTML = `
    <article class="style-mini-card">
      <span>Total de hilo</span>
      <strong>${formatNumber(entry.totalMeters)}</strong>
      <small>metros</small>
    </article>
    <article class="style-mini-card">
      <span>TH · TEX 27 (70%)</span>
      <strong>${formatNumber(entry.thMeters)}</strong>
      <small>metros</small>
    </article>
    <article class="style-mini-card">
      <span>THX · 40/2 (30%)</span>
      <strong>${formatNumber(entry.thxMeters)}</strong>
      <small>metros</small>
    </article>
    <article class="style-mini-card style-mini-card--accent">
      <span>Modo de compra</span>
      <strong>${modeText}</strong>
      <small>${modeHint}</small>
    </article>
  `;

  if (!visiblePlans.length) {
    planHost.innerHTML = '<div class="style-mini-placeholder">Este estilo ya está completo, pero todavía no alcanzó una distribución visible de conos en la vista previa.</div>';
    return;
  }

  planHost.innerHTML = visiblePlans.map((item) => `
    <article class="plan-card">
      <span class="plan-card__badge ${item.minMode ? "min" : ""}">
        <i class="fa-solid fa-file-excel"></i>
        ${item.minMode ? "Mínimo automático" : "Compra normal"}
      </span>
      <strong>${safeFileName(`${(poDigitsOnly(el("po")?.value) || "orden")}-${item.suffix}.xlsx`)}</strong>
      <p>${item.threadLabel} · ${item.sizeLabel} · estilo #${entry.index}</p>
      <div class="plan-meta">
        <div>
          <span>Qty order total</span>
          <strong>${formatNumber(item.row.qty, 0)}</strong>
        </div>
        <div>
          <span>Precio</span>
          <strong>${formatCurrency(item.row.unitPrice, item.row.unitPrice === 0.475 ? 3 : 2)}</strong>
        </div>
        <div>
          <span>Monto total</span>
          <strong>${formatCurrency(item.row.amount)}</strong>
        </div>
        <div>
          <span>Metros usados</span>
          <strong>${formatNumber(item.usedMeters)}</strong>
        </div>
      </div>
    </article>
  `).join("");
}

function renderPlan() {
  setMessage("");
  const data = getOrderData();
  const plan = buildPurchasePlan(data.entries);
  const stylePlanMap = new Map();

  plan.forEach((item) => {
    item.rows.forEach((row) => {
      if (!stylePlanMap.has(row.index)) stylePlanMap.set(row.index, []);
      stylePlanMap.get(row.index).push({
        suffix: item.suffix,
        threadLabel: item.threadLabel,
        sizeLabel: item.sizeLabel,
        minMode: item.minMode,
        row,
        usedMeters: item.suffix === "TH" ? row.thMeters : row.thxMeters
      });
    });
  });

  data.entries.forEach((entry) => {
    renderStylePreview(entry, stylePlanMap.get(entry.index) || []);
  });
}
function thinBorder() {
  return {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  };
}

function setCell(ws, addr, value) {
  const cell = ws.getCell(addr);
  cell.value = value;
  return cell;
}

function setFill(cell, argb = "FFFFFFFF") {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb }
  };
}

function setBorder(cell) {
  cell.border = thinBorder();
}

function setAlignment(cell, horizontal = "center", vertical = "middle", wrapText = true) {
  cell.alignment = { horizontal, vertical, wrapText };
}

function setFont(cell, options = {}) {
  cell.font = {
    name: "Arial",
    size: 12,
    ...options
  };
}

function styleRange(ws, startRow, startCol, endRow, endCol, callback) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      callback(ws.getRow(row).getCell(col), row, col);
    }
  }
}

function buildApprovalText(row, suffix) {
  if (row.approvalType === "muestra") {
    return "SE MANDARA MUESTRA DE TELA";
  }

  if (row.approvalType === "pendiente") {
    return "PENDIENTE MUESTRA DE TELA O IGUALACIÓN.";
  }

  const po = row.matchPo || "00000";
  const color = (row.matchColor || row.color || "").toUpperCase();
  return `IGUALAR AL MISMO COLOR DE LA ORDEN ${po}-${suffix} ${color}`.trim();
}

function prepareSheetBase(ws, itemRowsCount, approvalRowsCount) {
  ws.columns = [
    { width: 21.33 },
    { width: 22.89 },
    { width: 20.55 },
    { width: 20.11 },
    { width: 19.33 },
    { width: 20.11 },
    { width: 18.0 }
  ];

  ws.views = [{ showGridLines: false, zoomScale: 85 }];

  ws.getRow(1).height = 27;
  ws.getRow(2).height = 27;
  ws.getRow(3).height = 27;
  ws.getRow(4).height = 27;
  ws.getRow(5).height = 51.75;
  ws.getRow(6).height = 36;

  const firstItemRow = 7;
  for (let i = 0; i < itemRowsCount; i += 1) {
    const topRow = firstItemRow + (i * 2);
    const bottomRow = topRow + 1;
    ws.getRow(topRow).height = 132;
    ws.getRow(bottomRow).height = 20.25;

    [2, 3, 4, 5, 6, 7].forEach((col) => {
      ws.mergeCells(topRow, col, bottomRow, col);
    });
  }

  const totalRow = firstItemRow + (itemRowsCount * 2);
  ws.getRow(totalRow).height = 29.25;
  ws.getRow(totalRow + 1).height = 18;
  const approvalTitleRow = totalRow + 2;
  const approvalHeaderRow = approvalTitleRow + 1;

  ws.getRow(approvalTitleRow).height = 21;
  ws.getRow(approvalHeaderRow).height = 27;
  ws.mergeCells(`B${approvalTitleRow}:E${approvalTitleRow}`);
  ws.mergeCells(`C${approvalHeaderRow}:E${approvalHeaderRow}`);

  let cursor = approvalHeaderRow + 1;
  for (let i = 0; i < approvalRowsCount; i += 1) {
    ws.getRow(cursor).height = 31.5;
    ws.mergeCells(`C${cursor}:E${cursor}`);
    cursor += 1;
  }

  ws.getRow(cursor).height = 18;
  ws.getRow(cursor + 1).height = 39;
  ws.mergeCells(`A${cursor + 1}:B${cursor + 1}`);
  ws.mergeCells(`C${cursor + 1}:G${cursor + 1}`);

  ws.mergeCells("A1:B1");
  ws.mergeCells("C1:D1");
  ws.mergeCells("E1:G3");
  ws.mergeCells("A2:B2");
  ws.mergeCells("C2:D2");
  ws.mergeCells("A3:B3");
  ws.mergeCells("C3:D3");
  ws.mergeCells("A4:B4");
  ws.mergeCells("C4:D4");

  styleRange(ws, 1, 1, cursor + 1, 7, (cell) => setAlignment(cell));

  return {
    totalRow,
    approvalTitleRow,
    approvalHeaderRow,
    approvalStartRow: approvalHeaderRow + 1,
    lugarRow: cursor + 1
  };
}

function writeHeader(ws, data, item, layout) {
  const poValue = `${data.poBase || "orden"}-${item.suffix}`;
  const lugar = data.lugar || "PENDIENTE";

  setCell(ws, "A1", "PO #");
  setCell(ws, "C1", poValue);
  setCell(ws, "A2", "PROVEEDOR");
  setCell(ws, "C2", "TEXTILES BYCEL,S.A.");
  setCell(ws, "A3", "FECHA DE PO:");
  setCell(ws, "C3", formatDateExcel(data.fechaPo));
  setCell(ws, "A4", "NECESITO LA PRODUCCION:");
  setCell(ws, "C4", formatDateExcel(data.fechaProd));

  ws.getCell("C3").numFmt = "dd/mmm/yy";
  ws.getCell("C4").numFmt = "dd/mmm/yy";

  styleRange(ws, 1, 1, 4, 4, (cell) => {
    setFill(cell, "FFFFFFFF");
    setBorder(cell);
    setAlignment(cell);
  });

  ["A1", "A2", "A3", "A4"].forEach((addr) => setFont(ws.getCell(addr), { size: 14 }));
  ["C1", "C2", "C3", "C4"].forEach((addr) => setFont(ws.getCell(addr), { size: 14, bold: true }));

  setCell(ws, `A${layout.lugarRow}`, "LUGAR DE ENTREGA:");
  setCell(ws, `C${layout.lugarRow}`, lugar);
  styleRange(ws, layout.lugarRow, 1, layout.lugarRow, 7, (cell) => {
    setBorder(cell);
    setFill(cell, "FFFFFFFF");
    setAlignment(cell);
  });
  setFont(ws.getCell(`A${layout.lugarRow}`), { size: 16, bold: true });
  setFont(ws.getCell(`C${layout.lugarRow}`), {
    size: lugar.toUpperCase() === "PENDIENTE" ? 22 : 16,
    bold: true,
    color: { argb: lugar.toUpperCase() === "PENDIENTE" ? "FFFF0000" : "FF000000" }
  });
}

function writeTable(ws, item, layout) {
  const headers = [
    ["A6", "ITEM #"],
    ["B6", "STYLE"],
    ["C6", "COLOR"],
    ["D6", "SIZE"],
    ["E6", "TOTAL QTY\nORDER"],
    ["F6", "U/PRICE"],
    ["G6", "AMOUNT"]
  ];

  headers.forEach(([addr, text]) => {
    const cell = setCell(ws, addr, text);
    setFill(cell, "FFB7B7B7");
    setBorder(cell);
    setAlignment(cell);
    setFont(cell, { size: 12, bold: true });
  });

  const startRow = 7;
  item.rows.forEach((rowItem, idx) => {
    const topRow = startRow + (idx * 2);
    const bottomRow = topRow + 1;

    styleRange(ws, topRow, 1, bottomRow, 7, (cell) => {
      setBorder(cell);
      setFill(cell, "FFFFFFFF");
      setAlignment(cell);
    });

    setCell(ws, `A${bottomRow}`, rowItem.threadLabel);
    setCell(ws, `B${topRow}`, rowItem.styleRef || "");
    setCell(ws, `C${topRow}`, rowItem.color || "");
    setCell(ws, `D${topRow}`, rowItem.sizeLabel);
    setCell(ws, `E${topRow}`, rowItem.qty);
    setCell(ws, `F${topRow}`, rowItem.unitPrice);
    setCell(ws, `G${topRow}`, rowItem.amount);

    setFont(ws.getCell(`A${bottomRow}`), { size: 16 });
    setFont(ws.getCell(`B${topRow}`), { size: 16, bold: true });
    setFont(ws.getCell(`C${topRow}`), { size: 16 });
    setFont(ws.getCell(`D${topRow}`), { size: 16 });
    setFont(ws.getCell(`E${topRow}`), { size: 16 });
    setFont(ws.getCell(`F${topRow}`), { size: 16 });
    setFont(ws.getCell(`G${topRow}`), { size: 16 });

    ws.getCell(`E${topRow}`).numFmt = "#,##0";
    ws.getCell(`F${topRow}`).numFmt = rowItem.unitPrice === 0.475 ? "$#,##0.000" : "$#,##0.00";
    ws.getCell(`G${topRow}`).numFmt = "$#,##0.00";
  });

  styleRange(ws, layout.totalRow, 4, layout.totalRow, 5, (cell) => {
    setBorder(cell);
    setFill(cell, "FFFFFFFF");
    setAlignment(cell);
  });
  styleRange(ws, layout.totalRow, 7, layout.totalRow, 7, (cell) => {
    setBorder(cell);
    setFill(cell, "FFFFFFFF");
    setAlignment(cell);
  });

  setCell(ws, `D${layout.totalRow}`, "TOTAL");
  setCell(ws, `E${layout.totalRow}`, item.totalQty);
  setCell(ws, `G${layout.totalRow}`, item.totalAmount);
  setFont(ws.getCell(`D${layout.totalRow}`), { size: 16, bold: true });
  setFont(ws.getCell(`E${layout.totalRow}`), { size: 16, bold: true });
  setFont(ws.getCell(`G${layout.totalRow}`), { size: 16, bold: true });
  ws.getCell(`E${layout.totalRow}`).numFmt = "#,##0";
  ws.getCell(`G${layout.totalRow}`).numFmt = "$#,##0.00";
}

function writeApproval(ws, item, layout) {
  setCell(ws, `B${layout.approvalTitleRow}`, "APROBACIÓN");
  setCell(ws, `B${layout.approvalHeaderRow}`, "COLOR");
  setCell(ws, `C${layout.approvalHeaderRow}`, "DETALLE DE APROBACIÓN");

  styleRange(ws, layout.approvalTitleRow, 2, layout.approvalTitleRow, 5, (cell) => {
    setBorder(cell);
    setFill(cell, "FFFFFFFF");
    setAlignment(cell);
  });
  styleRange(ws, layout.approvalHeaderRow, 2, layout.approvalHeaderRow, 5, (cell) => {
    setBorder(cell);
    setFill(cell, "FFB7B7B7");
    setAlignment(cell);
  });
  setFont(ws.getCell(`B${layout.approvalTitleRow}`), { size: 16, bold: true });
  setFont(ws.getCell(`B${layout.approvalHeaderRow}`), { size: 14, bold: true });
  setFont(ws.getCell(`C${layout.approvalHeaderRow}`), { size: 14, bold: true });

  let cursor = layout.approvalStartRow;
  item.rows.forEach((rowItem) => {
    const approvalDetail = buildApprovalText(rowItem, item.suffix);
    styleRange(ws, cursor, 2, cursor, 5, (cell) => {
      setBorder(cell);
      setFill(cell, "FFFFFFFF");
      setAlignment(cell);
    });

    setCell(ws, `B${cursor}`, rowItem.color || "");
    setCell(ws, `C${cursor}`, approvalDetail);

    setFont(ws.getCell(`B${cursor}`), { size: 12 });
    setFont(ws.getCell(`C${cursor}`), { size: 10, bold: true, color: { argb: "FFFF0000" } });
    cursor += 1;
  });
}

async function addImages(workbook, ws, item, layout, headerBase64) {
  if (headerBase64) {
    const headerImageId = workbook.addImage({
      base64: headerBase64,
      extension: "png"
    });

    ws.addImage(headerImageId, {
      tl: { col: 4.02, row: 0.02 },
      ext: { width: 420, height: 136 }
    });
  }

  item.rows.forEach((rowItem, idx) => {
    if (!rowItem.imageBase64 || !rowItem.imageExtension) return;
    const imageId = workbook.addImage({
      base64: rowItem.imageBase64,
      extension: rowItem.imageExtension
    });

    const topRow = 6.08 + (idx * 2);
    ws.addImage(imageId, {
      tl: { col: 0.16, row: topRow },
      ext: { width: 84, height: 132 }
    });
  });
}

async function buildWorkbook(item, data, headerBase64) {
  const workbook = new ExcelJS.Workbook();
  setWorkbookMeta(workbook);

  const ws = workbook.addWorksheet(item.suffix);
  const layout = prepareSheetBase(ws, item.rows.length, item.rows.length);
  writeHeader(ws, data, item, layout);
  writeTable(ws, item, layout);
  writeApproval(ws, item, layout);
  await addImages(workbook, ws, item, layout, headerBase64);

  return workbook;
}

async function downloadExcels() {
  setMessage("");
  const data = getOrderData();
  const error = validateEntries(data.entries);
  if (error) {
    setMessage(error, "error");
    return;
  }

  if (!data.poBase) {
    setMessage("Coloca el número de PO antes de descargar.", "error");
    return;
  }

  const plan = buildPurchasePlan(data.entries);
  if (!plan.length) {
    setMessage("No hay estilos válidos para generar el Excel.", "error");
    return;
  }

  try {
    const headerBase64 = await fetchAssetAsBase64(HEADER_IMAGE_PATH);

    for (let index = 0; index < plan.length; index += 1) {
      const item = plan[index];
      const workbook = await buildWorkbook(item, data, headerBase64);
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      forceDownload(blob, safeFileName(`${data.poBase}-${item.suffix}.xlsx`));
      if (index < plan.length - 1) await sleep(350);
    }

    setMessage(`Listo. Se generó ${plan.length === 1 ? "1 archivo" : `${plan.length} archivos`} para la PO ${data.poBase}.`, "success");
  } catch (error) {
    console.error(error);
    setMessage("Hubo un problema generando el Excel. Abre el proyecto con Live Server o un servidor local para que carguen bien los recursos.", "error");
  }
}

function bindGlobalEvents() {
  ["po", "fechaPo", "fechaProd", "lugar"].forEach((id) => {
    el(id).addEventListener("input", renderPlan);
  });

  el("fechaPo").addEventListener("change", (event) => {
    const value = event.target.value;
    el("fechaProd").value = addDays(value, 7);
    renderPlan();
  });

  const applyStyleCount = ({ commit = false } = {}) => {
    const input = el("qtyStyles");
    const raw = (input.value || "").trim();
    if (!raw) {
      if (commit) {
        lastStyleCount = 1;
        renderStyleCards(1);
      }
      return;
    }

    const value = normalizeStyleCount(raw, lastStyleCount);
    if (commit) input.value = String(value);
    if (value !== lastStyleCount || commit) {
      lastStyleCount = value;
      renderStyleCards(value);
    }
  };

  el("qtyStyles").addEventListener("input", () => applyStyleCount());
  el("qtyStyles").addEventListener("change", () => applyStyleCount({ commit: true }));
  el("qtyStyles").addEventListener("blur", () => applyStyleCount({ commit: true }));

  el("downloadBtn").addEventListener("click", downloadExcels);
}

function init() {
  setTodayDefaults();
  bindGlobalEvents();
  lastStyleCount = 1;
  renderStyleCards(1);
}

init();
