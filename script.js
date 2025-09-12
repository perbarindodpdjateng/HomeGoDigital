let previousData = '';
let originalData = [];
let scrollPaused = false;

const $ = id => document.getElementById(id);

async function fetchData() {
  try {
    const res = await fetch('https://script.google.com/macros/s/AKfycbxUszW9ZIJp-T4NZNftrkfpxcV2WKTaL3hZ8tU7pE7oJl2hbtNLU_xqLYIn9Y3oAAHuOw/exec', { cache: 'no-store' });
    const data = await res.json();
    const dataStr = JSON.stringify(data);
    if (dataStr !== previousData) {
      previousData = dataStr;
      originalData = data;
      renderTable(data);
      $('loading').style.display = 'none';
    }
  } catch (e) {
    console.error("Gagal memuat:", e);
    $('loading').textContent = "⚠️ Gagal memuat data.";
  }
}

function renderTable(data) {
  const table = $('dataTable');
  const fragment = document.createDocumentFragment();
  data.forEach((row, i) => {
    const tr = document.createElement('tr');
    row.forEach((cell, j) => {
      const tag = i === 0 ? 'th' : 'td';
      const el = document.createElement(tag);
      if (i === 0) {
        el.textContent = cell;
      } else if (j === 4 && typeof cell === "string" && cell.includes("google.com/maps")) {
        el.textContent = "🏦";
      } else if (typeof cell === "string" && cell.startsWith("http")) {
        el.innerHTML = `<a href="${cell}" target="_blank">🌐</a>`;
      } else {
        el.textContent = cell;
      }
      tr.appendChild(el);
    });
    fragment.appendChild(tr);
  });
  table.innerHTML = '';
  table.appendChild(fragment);
  updateRowCount(data);
}

function updateRowCount(data) {
  $('rowCount').textContent = `(${data.length - 1})`;
}

function rollingMarquee() {
  if (scrollPaused) return;
  const table = $('dataTable');
  const firstRow = table.rows[1];
  if (!firstRow) return;
  firstRow.classList.add("fade-out");
  setTimeout(() => {
    table.deleteRow(1);
    table.appendChild(firstRow);
    firstRow.classList.remove("fade-out");
  }, 1000);
}

function toggleMarquee() {
  scrollPaused = !scrollPaused;
  $('marqueeBtn').textContent = scrollPaused ? "▶️" : "⏸";
}

function debounce(fn, delay) {
  let t;
  return function (...args) {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), delay);
  };
}

$('searchInput').addEventListener("input", debounce(function () {
  const kw = this.value.toLowerCase();
  const filtered = originalData.filter((r, i) =>
    i === 0 || r.some(c => String(c).toLowerCase().includes(kw))
  );
  renderTable(filtered);
}, 300));

function exportCSV() {
  const csv = originalData.map(r => r.map(v => `"${v}"`).join(',')).join('\n');
  downloadBlob(csv, 'rekap.csv', 'text/csv');
}

function exportXLS() {
  const html = `<table>${originalData.map(r => `<tr>${r.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</table>`;
  downloadBlob(html, 'rekap.xls', 'application/vnd.ms-excel');
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement('a'), { href: url, download: filename });
  a.click();
  URL.revokeObjectURL(url);
}

fetchData();
setInterval(fetchData, 15000);
setInterval(rollingMarquee, 7000);
