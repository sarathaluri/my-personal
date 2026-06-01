const state = {
  data: null,
  currentToolId: null,
  search: '',
  status: 'all'
};

const el = (id) => document.getElementById(id);
const money = (value) => new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(Number(value || 0));
const titleCase = (text) => String(text).replaceAll('_', ' ');
const statusClass = (value) => `s-${String(value).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')}`;

function getCurrentTool() {
  return state.data.tools.find((tool) => tool.id === state.currentToolId) || state.data.tools[0];
}

function filteredRecords(tool = getCurrentTool()) {
  return tool.records.filter((record) => {
    const matchesSearch = !state.search || Object.values(record).some((value) => String(value).toLowerCase().includes(state.search));
    const matchesStatus = state.status === 'all' || String(record[tool.statusField]) === state.status;
    return matchesSearch && matchesStatus;
  });
}

function renderToolOptions() {
  el('tool-select').innerHTML = state.data.tools
    .map((tool) => `<option value="${tool.id}">${tool.title}</option>`)
    .join('');
  el('tool-select').value = state.currentToolId;
}

function renderStatusOptions(tool) {
  const statuses = [...new Set(tool.records.map((record) => String(record[tool.statusField])))];
  el('status-filter').innerHTML = ['<option value="all">All statuses</option>']
    .concat(statuses.map((status) => `<option value="${status}">${status}</option>`))
    .join('');
  el('status-filter').value = state.status;
}

function renderHeader(tool) {
  el('source-file').textContent = `Source: ${tool.sourceFile}`;
  el('tool-title').textContent = tool.title;
  el('tool-description').textContent = tool.description;
  el('dataset-date').textContent = state.data.generatedOn;
}

function renderKpis(tool) {
  el('kpi-grid').innerHTML = tool.kpis.map((kpi) => `
    <article class="kpi-card">
      <span>${kpi.label}</span>
      <strong>${money(kpi.value)}${kpi.suffix}</strong>
    </article>
  `).join('');
}

function renderRiskChart(tool, records) {
  const counts = records.reduce((acc, record) => {
    const key = String(record[tool.statusField]);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const max = Math.max(...Object.values(counts), 1);
  const colors = { LOW: 'var(--good)', MEDIUM: 'var(--warn)', HIGH: 'var(--bad)', CRITICAL: 'var(--critical)', DISCLOSED: 'var(--good)', GAP: 'var(--bad)', '0-30 DAYS': 'var(--good)', '31-60 DAYS': 'var(--warn)', '61-90 DAYS': 'var(--warn)', '>90 DAYS': 'var(--bad)' };
  el('risk-count').textContent = `${records.length} matching rows`;
  el('risk-chart').innerHTML = Object.entries(counts).map(([status, count]) => `
    <div class="risk-row">
      <strong>${status}</strong>
      <div class="risk-bar"><div class="risk-fill" style="width:${(count / max) * 100}%;background:${colors[status] || 'var(--brand)'}"></div></div>
      <span>${count}</span>
    </div>
  `).join('') || '<p>No records match the filters.</p>';
}

function recordName(record) {
  return record.Scheme_Name || record.Contractor || record.Project_Name || record.Bill_ID || 'Record';
}

function renderAmountChart(tool, records) {
  const topRecords = [...records]
    .sort((a, b) => Number(b[tool.amountField] || 0) - Number(a[tool.amountField] || 0))
    .slice(0, 6);
  const max = Math.max(...topRecords.map((record) => Number(record[tool.amountField] || 0)), 1);
  el('amount-label').textContent = titleCase(tool.amountField);
  el('amount-chart').innerHTML = topRecords.map((record) => `
    <div class="amount-row">
      <div class="amount-meta"><strong>${recordName(record)}</strong><span>₹${money(record[tool.amountField])} Cr</span></div>
      <div class="amount-track"><div class="amount-fill" style="width:${(Number(record[tool.amountField] || 0) / max) * 100}%"></div></div>
    </div>
  `).join('') || '<p>No records match the filters.</p>';
}

function renderTable(tool, records) {
  const columns = Object.keys(tool.records[0] || {});
  el('record-count').textContent = `${records.length} of ${tool.records.length} records shown`;
  el('data-table').innerHTML = `
    <thead><tr>${columns.map((column) => `<th>${titleCase(column)}</th>`).join('')}</tr></thead>
    <tbody>
      ${records.map((record) => `<tr>${columns.map((column) => tableCell(tool, column, record[column])).join('')}</tr>`).join('')}
    </tbody>
  `;
}

function tableCell(tool, column, value) {
  if (column === tool.statusField || column === 'Stall_Risk' || column === 'MSME') {
    return `<td><span class="badge ${statusClass(value)}">${value}</span></td>`;
  }
  return `<td>${value}</td>`;
}

function renderPortfolioSummary() {
  const rows = state.data.tools.flatMap((tool) => tool.records.map((record) => ({ tool, record })));
  const critical = rows.filter(({ tool, record }) => ['CRITICAL', 'HIGH', 'GAP', '>90 DAYS'].includes(String(record[tool.statusField]))).length;
  el('portfolio-total').textContent = `${rows.length} records`;
  el('portfolio-risk').textContent = `${critical} records need priority attention across the four finance workstreams.`;
}

function renderAll() {
  const tool = getCurrentTool();
  const records = filteredRecords(tool);
  renderToolOptions();
  renderStatusOptions(tool);
  renderHeader(tool);
  renderKpis(tool);
  renderRiskChart(tool, records);
  renderAmountChart(tool, records);
  renderTable(tool, records);
  renderPortfolioSummary();
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function downloadCsv() {
  const tool = getCurrentTool();
  const records = filteredRecords(tool);
  const columns = Object.keys(tool.records[0] || {});
  const csv = [columns.join(','), ...records.map((record) => columns.map((column) => csvEscape(record[column])).join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${tool.id}-dataset.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

async function loadDashboardData() {
  try {
    const response = await fetch('data/dashboard-data.json');
    if (!response.ok) {
      throw new Error(`Dataset request failed with ${response.status}`);
    }
    return response.json();
  } catch (error) {
    if (window.DASHBOARD_DATA) {
      return window.DASHBOARD_DATA;
    }
    throw error;
  }
}

async function boot() {
  state.data = await loadDashboardData();
  state.currentToolId = state.data.tools[0].id;

  el('tool-select').addEventListener('change', (event) => {
    state.currentToolId = event.target.value;
    state.status = 'all';
    state.search = '';
    el('search-input').value = '';
    renderAll();
  });
  el('search-input').addEventListener('input', (event) => {
    state.search = event.target.value.trim().toLowerCase();
    renderAll();
  });
  el('status-filter').addEventListener('change', (event) => {
    state.status = event.target.value;
    renderAll();
  });
  el('download-table').addEventListener('click', downloadCsv);
  el('download-current').addEventListener('click', downloadCsv);

  renderAll();
}

boot().catch((error) => {
  document.body.innerHTML = `<main class="shell"><article class="panel"><h1>Unable to load dashboard</h1><p>${error.message}</p></article></main>`;
});
