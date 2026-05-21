/**
 * Static Excel table viewer for Cloudflare Pages.
 * Uses SheetJS in the browser to parse FPSBeneficiaryDetailsbalapora.xlsx.
 */
(() => {
  const EXCEL_FILE = 'FPSBeneficiaryDetailsbalapora.xlsx';

  const table = document.getElementById('dataTable');
  const tableHead = table.querySelector('thead');
  const tableBody = table.querySelector('tbody');
  const searchInput = document.getElementById('searchInput');
  const statusMessage = document.getElementById('statusMessage');
  const rowCount = document.getElementById('rowCount');

  let rows = [];
  let headers = [];

  /**
   * Initializes the app once DOM is ready.
   */
  async function init() {
    try {
      const workbook = await loadWorkbook(EXCEL_FILE);
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      headers = rows.length ? Object.keys(rows[0]) : [];

      if (!rows.length || !headers.length) {
        showStatus('No data rows found in the Excel file.', true);
        table.style.display = 'none';
        return;
      }

      renderTable(headers, rows);
      updateRowCount(rows.length, rows.length);
      showStatus(`Loaded ${rows.length} rows from “${firstSheetName}”.`, false);
      bindSearch();
    } catch (error) {
      console.error(error);
      showStatus(`Unable to load ${EXCEL_FILE}. Ensure it exists in the repo root.`, true);
      table.style.display = 'none';
    }
  }

  /**
   * Fetches and parses the workbook file.
   */
  async function loadWorkbook(fileName) {
    const response = await fetch(fileName);
    if (!response.ok) {
      throw new Error(`Fetch failed with status ${response.status}`);
    }

    const arrayBuffer = await response.arrayBuffer();
    return XLSX.read(arrayBuffer, { type: 'array' });
  }

  /**
   * Renders table headers and row body.
   */
  function renderTable(columnNames, rowData) {
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    const headerRow = document.createElement('tr');
    columnNames.forEach((name) => {
      const th = document.createElement('th');
      th.textContent = name;
      headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);

    const fragment = document.createDocumentFragment();
    rowData.forEach((row) => {
      const tr = document.createElement('tr');
      columnNames.forEach((name) => {
        const td = document.createElement('td');
        td.textContent = row[name];
        tr.appendChild(td);
      });
      fragment.appendChild(tr);
    });

    tableBody.appendChild(fragment);
    table.style.display = 'table';
  }

  /**
   * Filters rows on every keystroke.
   */
  function bindSearch() {
    searchInput.addEventListener('input', () => {
      const query = searchInput.value.trim().toLowerCase();
      const filteredRows = !query
        ? rows
        : rows.filter((row) =>
            headers.some((header) => String(row[header]).toLowerCase().includes(query))
          );

      renderTable(headers, filteredRows);
      updateRowCount(filteredRows.length, rows.length);
      showStatus(
        filteredRows.length
          ? `Showing ${filteredRows.length} matching row(s).`
          : 'No matching rows found.',
        filteredRows.length === 0
      );
    });
  }

  function updateRowCount(visible, total) {
    rowCount.textContent = `${visible} / ${total} rows`;
  }

  function showStatus(message, isError) {
    statusMessage.textContent = message;
    statusMessage.style.color = isError ? '#b91c1c' : '#475569';
  }

  document.addEventListener('DOMContentLoaded', init);
})();
