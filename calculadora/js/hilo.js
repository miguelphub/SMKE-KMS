const rows = [
  { id: 1, process: 'CUELLO', stitch: 'Overlock', iso: 514, width: '1/4"', multiful: 20, remark: 57 },
  { id: 2, process: 'CUELLO', stitch: 'CERRADORA', iso: 514, width: '1/4"', multiful: 15, remark: 57 },
  { id: 3, process: 'VOCAMANGA', stitch: 'Overlock', iso: 514, width: '1/4"', multiful: 20, remark: 52 },
  { id: 4, process: 'MANGA', stitch: 'sambo', iso: 504, width: '1/4"', multiful: 26, remark: 39 },
  { id: 5, process: 'LADO', stitch: 'Overlock', iso: 514, width: '1/4"', multiful: 20, remark: 44 },
  { id: 6, process: 'ABAJO', stitch: 'sambo', iso: 504, width: '1/4"', multiful: 26, remark: 97 },
  { id: 7, process: 'CINTA DE CUELLO', stitch: 'CERRADORA', iso: 514, width: '1/4"', multiful: 15, remark: 62 },
  { id: 8, process: 'HOMBRO', stitch: 'sambo', iso: 504, width: '1/4"', multiful: 26, remark: 21 }
];

const els = {
  tableBody: document.getElementById('tableBody'),
  rowTemplate: document.getElementById('rowTemplate'),
  size: document.getElementById('size'),
  sizePreview: document.getElementById('sizePreview'),
  sizeTitle: document.getElementById('sizeTitle'),
  lossPercent: document.getElementById('lossPercent'),
  thread2: document.getElementById('thread2'),
  thread1Meters: document.getElementById('thread1Meters'),
  thread2Meters: document.getElementById('thread2Meters'),
  threadTotalMeters: document.getElementById('threadTotalMeters'),
  totalNet: document.getElementById('totalNet'),
  totalWithMultiful: document.getElementById('totalWithMultiful'),
  totalConsumption: document.getElementById('totalConsumption'),
  clearBtn: document.getElementById('clearBtn'),
  fillExampleBtn: document.getElementById('fillExampleBtn'),
  date: document.getElementById('date')
};

function formatNumber(value, digits = 2) {
  if (!Number.isFinite(value)) return '0.00';
  return value.toLocaleString('es-GT', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function parseNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function setToday() {
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth() + 1).padStart(2, '0');
  const d = String(today.getDate()).padStart(2, '0');
  els.date.value = `${y}-${m}-${d}`;
}

function renderRows() {
  els.tableBody.innerHTML = '';

  rows.forEach((row) => {
    const fragment = els.rowTemplate.content.cloneNode(true);
    const tr = fragment.querySelector('tr');
    tr.dataset.id = row.id;

    fragment.querySelector('.col-index').textContent = row.id;
    fragment.querySelector('.col-process').textContent = row.process;
    fragment.querySelector('.col-stitch').textContent = row.stitch;
    fragment.querySelector('.col-iso').textContent = row.iso;
    fragment.querySelector('.col-width').textContent = row.width;
    fragment.querySelector('.col-multiful').textContent = row.multiful;
    fragment.querySelector('.col-remark').textContent = row.remark;

    const netInput = fragment.querySelector('.net-input');
    netInput.addEventListener('input', calculateAll);

    els.tableBody.appendChild(fragment);
  });
}

function getRowInputs() {
  return [...document.querySelectorAll('.net-input')];
}

function calculateAll() {
  const lossFactor = 1 + (parseNumber(els.lossPercent.value) / 100);
  const thread2Meters = parseNumber(els.thread2.value);
  const inputs = getRowInputs();

  let totalNet = 0;
  let totalWithMultiful = 0;
  let totalConsumption = 0;

  inputs.forEach((input, index) => {
    const row = rows[index];
    const netValue = parseNumber(input.value);
    const withMultiful = netValue * row.multiful;
    const consumption = withMultiful * lossFactor;

    totalNet += netValue;
    totalWithMultiful += withMultiful;
    totalConsumption += consumption;

    const tr = input.closest('tr');
    tr.querySelector('.col-with-multiful').textContent = formatNumber(withMultiful);
    tr.querySelector('.col-consumption').textContent = formatNumber(consumption);
  });

  const thread1Meters = totalConsumption / 100;
  const totalMeters = thread1Meters + thread2Meters;

  els.totalNet.textContent = formatNumber(totalNet);
  els.totalWithMultiful.textContent = formatNumber(totalWithMultiful);
  els.totalConsumption.textContent = formatNumber(totalConsumption);
  els.thread1Meters.textContent = formatNumber(thread1Meters);
  els.thread2Meters.textContent = formatNumber(thread2Meters);
  els.threadTotalMeters.textContent = formatNumber(totalMeters);

  [els.thread1Meters, els.thread2Meters, els.threadTotalMeters].forEach((el) => {
    el.classList.remove('flash-update');
    void el.offsetWidth;
    el.classList.add('flash-update');
  });
}

function syncSize() {
  const value = (els.size.value || 'M').trim().toUpperCase();
  els.sizePreview.textContent = value || 'M';
  els.sizeTitle.textContent = value || 'M';
}

function clearAll() {
  getRowInputs().forEach((input) => {
    input.value = '';
  });
  els.thread2.value = '1';
  els.lossPercent.value = '10';
  calculateAll();
}

function fillExample() {
  const example = [1, 0, 1.2, 0, 0, 0, 26, 0];
  getRowInputs().forEach((input, index) => {
    input.value = example[index] ?? '';
  });
  calculateAll();
}

function bindEvents() {
  els.lossPercent.addEventListener('input', calculateAll);
  els.thread2.addEventListener('input', calculateAll);
  els.size.addEventListener('input', syncSize);
  els.clearBtn.addEventListener('click', clearAll);
  els.fillExampleBtn.addEventListener('click', fillExample);
}

function init() {
  setToday();
  renderRows();
  bindEvents();
  syncSize();
  calculateAll();
}

init();
