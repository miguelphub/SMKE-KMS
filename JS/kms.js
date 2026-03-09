const el = (id) => document.getElementById(id);

const SIZE_SETS = {
  regulares: ["XXS", "XS", "S", "M", "L", "XL", "XXL"],
  toddler: ["12M", "18M", "2T", "3T", "4T", "5T"]
};

const PLUS_SIZES = ["1X", "2X", "3X"];

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

function refreshStyleOptionVisibility(i) {
  const colorWrap = el(`colorWrap_${i}`);
  const contentWrap = el(`contentWrap_${i}`);
  const plusWrap = el(`plusWrap_${i}`);
  const plusTable = el(`plusTable_${i}`);
  const sizeTypeSel = el(`sizeType_${i}`);
  const plusSel = el(`plus_${i}`);

  const htOn = isHTSelected();
  const isRegular = (sizeTypeSel?.value || "regulares") === "regulares";

  if (colorWrap) {
    colorWrap.style.display = htOn ? "flex" : "none";
    colorWrap.style.flexDirection = "column";
  }

  if (contentWrap) {
    contentWrap.style.display = htOn ? "flex" : "none";
    contentWrap.style.flexDirection = "column";
  }

  if (plusWrap) {
    plusWrap.style.display = (htOn && isRegular) ? "flex" : "none";
    plusWrap.style.flexDirection = "column";
    if (!htOn && plusSel) {
      plusSel.value = "no";
    }
  }

  if (plusTable) {
    plusTable.hidden = !(htOn && isRegular && plusSel?.value === "si");
  }
}

function renderStyles(n) {
  const cont = el("stylesContainer");
  if (!cont) return;

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
          <input id="content_${i}" type="text" placeholder="Ej: 100% Cotton" />
        </label>

        <label>
          Tipo de tallas
          <select id="sizeType_${i}">
            <option value="regulares" selected>Regulares</option>
            <option value="toddler">Toddler</option>
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
    `;

    cont.appendChild(card);

    const sizeTypeSel = el(`sizeType_${i}`);
    const plusSel = el(`plus_${i}`);

    function renderTable() {
      const key = sizeTypeSel.value || "regulares";
      const sizes = SIZE_SETS[key] || SIZE_SETS.regulares;

      const table = el(`table_${i}`);
      table.style.setProperty("--cols", String(sizes.length));

      el(`thead_${i}`).innerHTML = sizes.map(s => `<div>${s}</div>`).join("");
      el(`tqty_${i}`).innerHTML = sizes.map(s => {
        const id = `q${sizeToId(s)}_${i}`;
        return `<input id="${id}" type="number" min="0" step="1" placeholder="-" inputmode="numeric" />`;
      }).join("");

      refreshStyleOptionVisibility(i);
    }

    sizeTypeSel.addEventListener("change", renderTable);
    plusSel.addEventListener("change", () => refreshStyleOptionVisibility(i));

    renderTable();
  }

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

  const baseSizes = SIZE_SETS[sizeType] || SIZE_SETS.regulares;
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
    plusFiltered
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

function setupHeader8(ws, poText, fechaPoYMD, fechaProdYMD, lugar) {
  ws.columns = [
    { width: 18 }, // A
    { width: 24 }, // B
    { width: 16 }, // C
    { width: 30 }, // D
    { width: 12 }, // E
    { width: 10 }, // F
    { width: 10 }, // G
    { width: 12 }  // H
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

  // ITEM # grande queda vacío; la referencia va abajo en la fila amarilla
  setCell(ws, `A${dataStart}`, st.data.itemText || "");
  setCell(ws, `B${dataStart}`, "");
  setCell(ws, `C${dataStart}`, st.data.color || "");
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

  // Referencia del estilo debajo de ITEM #
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

  // ITEM # grande queda vacío; la referencia va abajo en la fila amarilla
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

  // Referencia del estilo debajo de ITEM #
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
// Bloques CT / HG
// =========================
function writeTagBlock(ws, startRow, st, title) {
  const totalQty = sumStyleQty(st.data, true);

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

  // ITEM # grande queda vacío; la referencia va abajo en la fila amarilla
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

  // Referencia del estilo debajo de ITEM #
  setCell(ws, `B${totalRow}`, st.data.refText || "");

  ws.getCell(`D${totalRow}`).value = { formula: `D${dataStart}` };
  ws.getCell(`F${totalRow}`).value = { formula: `F${dataStart}` };
  ws.getCell(`F${totalRow}`).numFmt = '$#,##0.00';

  centerRange(ws, totalRow, 1, totalRow, 6);
  setFontRange(ws, totalRow, 1, totalRow, 6, { bold: true });

  applyBorderRange(ws, startRow, 1, totalRow, 6, thinBorder());

  return totalRow;
}

// =========================
// Workbooks
// =========================
async function buildHeatTransferWorkbook(payload) {
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("HT");
  setupHeader8(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeHTBlock(ws, row, st);
    totalRows.push(endRow);

    // fila en blanco entre bloques
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
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("SS");
  setupHeader7(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeSSBlock(ws, row, st);
    totalRows.push(endRow);

    // fila en blanco entre bloques
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
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("CL");
  setupHeader6(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeTagBlock(ws, row, st, "COLOR TAG");
    totalRows.push(endRow);

    // fila en blanco entre bloques
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

async function buildHangTagWorkbook(payload) {
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("TAG");
  setupHeader6(ws, payload.poText, payload.fechaPoYMD, payload.fechaProdYMD, payload.lugar);

  let row = 6;
  const totalRows = [];

  for (const st of payload.styles) {
    const endRow = writeTagBlock(ws, row, st, "HANG TAG");
    totalRows.push(endRow);

    // fila en blanco entre bloques
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

    const styles = [];

    for (let i = 1; i <= n; i++) {
      const data = getStyleData(i);

      const hasAnyBase = data.baseFiltered.length > 0;
      const hasAnyPlus = isHTSelected() && data.sizeType === "regulares" && data.plusEnabled && data.plusFiltered.length > 0;

      if (!hasAnyBase && !hasAnyPlus) continue;

      styles.push({
        idx: i,
        data
      });
    }

    if (styles.length === 0) {
      alert("Ingresa al menos una cantidad en alguna talla.");
      return;
    }

    if (compras.includes("heat_transfer")) {
      const wbHT = await buildHeatTransferWorkbook({
        poText: `${poBase}-HT`,
        fechaPoYMD,
        fechaProdYMD,
        lugar,
        styles
      });
      await downloadWorkbook(wbHT, safeFileName(`${poBase}-HT.xlsx`));
      await sleep(350);
    }

    if (compras.includes("size_strip")) {
      const wbSS = await buildSizeStripWorkbook({
        poText: `${poBase}-SS`,
        fechaPoYMD,
        fechaProdYMD,
        lugar,
        styles
      });
      await downloadWorkbook(wbSS, safeFileName(`${poBase}-SS.xlsx`));
      await sleep(350);
    }

    if (compras.includes("color_tag")) {
      const wbCT = await buildColorTagWorkbook({
        poText: `${poBase}-CL`,
        fechaPoYMD,
        fechaProdYMD,
        lugar,
        styles
      });
      await downloadWorkbook(wbCT, safeFileName(`${poBase}-CL.xlsx`));
      await sleep(350);
    }

    if (compras.includes("hang_tag")) {
      const wbHG = await buildHangTagWorkbook({
        poText: `${poBase}-TAG`,
        fechaPoYMD,
        fechaProdYMD,
        lugar,
        styles
      });
      await downloadWorkbook(wbHG, safeFileName(`${poBase}-TAG.xlsx`));
      await sleep(350);
    }

  } catch (err) {
    console.error(err);
    alert(err?.message || String(err));
  }
});