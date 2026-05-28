/**
 * Village-wise static Excel viewer for Cloudflare Pages.
 * Add more village files to VILLAGE_SOURCES when new workbooks are available.
 */
(() => {
  const VILLAGE_SOURCES = [
    {
      id: 'balapora',
      name: 'Balapora',
      file: 'FPSBeneficiaryDetailsbalapora.xlsx',
    },
    {
      id: 'ganowpora',
      name: 'GANOWPORA',
      file: 'FPSBeneficiaryDetailGANOWPORA.xlsx',
    },
  ];

  const TABLE_COLUMNS = [
    'S.No.',
    'Ration Card No.',
    'Family Head',
    'M.S. No.',
    'Member Name',
    'Member ID',
    "Member's Age",
    'UID No.',
    'Relation with HoF',
    'Mother Name',
    'Father Name',
    'Gender',
    'Scheme',
    'FPS',
  ];

  const table = document.getElementById('dataTable');
  const tableHead = table.querySelector('thead');
  const tableBody = table.querySelector('tbody');
  const activeVillageName = document.getElementById('activeVillageName');
  const activeVillageNote = document.getElementById('activeVillageNote');
  const summaryGrid = document.getElementById('summaryGrid');
  const searchInputs = {
    serial: document.getElementById('serialSearch'),
    rationCard: document.getElementById('rationCardSearch'),
    familyHead: document.getElementById('familyHeadSearch'),
    memberName: document.getElementById('memberNameSearch'),
    uidLast4: document.getElementById('uidLast4Search'),
  };
  const villageFilter = document.getElementById('villageFilter');
  const schemeFilter = document.getElementById('schemeFilter');
  const fpsFilter = document.getElementById('fpsFilter');
  const clearSearchButton = document.getElementById('clearSearchButton');
  const statusMessage = document.getElementById('statusMessage');
  const rowCount = document.getElementById('rowCount');
  const filterHint = document.getElementById('filterHint');

  let villages = [];
  let activeVillageId = VILLAGE_SOURCES[0].id;

  async function init() {
    bindFilters();

    try {
      villages = await Promise.all(VILLAGE_SOURCES.map(loadVillage));
      activeVillageId = villages[0]?.id || activeVillageId;

      if (!villages.length || villages.every((village) => village.rows.length === 0)) {
        showStatus('No beneficiary rows found in the configured village files.', true);
        table.style.display = 'none';
        return;
      }

      renderVillageFilter();
      setActiveVillage(activeVillageId);
      showStatus('Village data loaded', false);
    } catch (error) {
      console.error(error);
      showStatus('Unable to load the village Excel data. Check that the workbook exists in the repo root.', true);
      table.style.display = 'none';
    }
  }

  async function loadVillage(source) {
    const workbook = await loadWorkbook(source.file);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const rows = parseBeneficiaryRows(matrix, source);

    return {
      ...source,
      sheetName,
      rows,
      stats: getVillageStats(rows),
    };
  }

  async function loadWorkbook(fileName) {
    const response = await fetch(fileName);
    if (!response.ok) {
      throw new Error(`Fetch failed for ${fileName} with status ${response.status}`);
    }

    const arrayBuffer = await response.arrayBuffer();
    return XLSX.read(arrayBuffer, { type: 'array' });
  }

  function parseBeneficiaryRows(matrix, village) {
    const headerIndex = matrix.findIndex((row) =>
      String(row[0]).trim().toLowerCase() === 's.no.' &&
      String(row[1]).trim().toLowerCase() === 'ration card no.'
    );

    if (headerIndex === -1) {
      return [];
    }

    const columns = getColumnIndexes(matrix[headerIndex]);
    const rows = [];
    let currentFps = '';
    let currentFpsCode = '';
    let currentScheme = '';
    let currentCard = '';
    let currentFamilyHead = '';

    matrix.slice(headerIndex + 1).forEach((rawRow) => {
      const row = rawRow.map((value) => String(value).trim());
      const firstCell = row[0] || '';

      if (!row.some(Boolean)) {
        return;
      }

      if (firstCell.startsWith('FPS Detail:')) {
        const detail = firstCell.replace('FPS Detail:', '').trim();
        const match = detail.match(/^(.*)\(([^)]+)\)$/);
        currentFps = match ? match[1].trim() : detail;
        currentFpsCode = match ? match[2].trim() : '';
        return;
      }

      if (firstCell.startsWith('Scheme Name')) {
        currentScheme = firstCell
          .replace('Scheme Name', '')
          .replace(':', '')
          .replace(/\[\d+\]/g, '')
          .trim();
        return;
      }

      if (firstCell.startsWith('Total Ration Cards') || firstCell.startsWith('Note:-')) {
        return;
      }

      const memberId = readColumn(row, columns.memberId);

      if (!/^\d+$/.test(memberId)) {
        return;
      }

      currentCard = readColumn(row, columns.card) || currentCard;
      currentFamilyHead = readColumn(row, columns.familyHead) || currentFamilyHead;

      rows.push({
        Village: village.name,
        VillageId: village.id,
        'S.No.': readColumn(row, columns.serial),
        'Ration Card No.': currentCard,
        'Family Head': currentFamilyHead,
        'M.S. No.': readColumn(row, columns.msNo),
        'Member Name': readColumn(row, columns.memberName),
        'Member ID': memberId,
        "Member's Age": readColumn(row, columns.age),
        'UID No.': readColumn(row, columns.uid),
        'Relation with HoF': readColumn(row, columns.relation),
        'Mother Name': readColumn(row, columns.mother),
        'Father Name': readColumn(row, columns.father),
        Gender: readColumn(row, columns.gender),
        Scheme: currentScheme || 'Unspecified',
        FPS: currentFps || 'Unspecified',
        'FPS Code': currentFpsCode,
      });
    });

    return rows;
  }

  function getColumnIndexes(headerRow) {
    return {
      serial: findColumn(headerRow, ['S.No.']),
      card: findColumn(headerRow, ['Ration Card No.']),
      familyHead: findColumn(headerRow, ['Family Head']),
      msNo: findColumn(headerRow, ['M.S. No.']),
      memberName: findColumn(headerRow, ['Member Name(in Eng)', 'Member Name']),
      memberId: findColumn(headerRow, ['Member ID']),
      age: findColumn(headerRow, ["Member's Age*", "Member's Age"]),
      uid: findColumn(headerRow, ['UID No.']),
      relation: findColumn(headerRow, ['Relation with HoF']),
      mother: findColumn(headerRow, ['Mother Name']),
      father: findColumn(headerRow, ['Father Name']),
      gender: findColumn(headerRow, ['Gender']),
    };
  }

  function findColumn(headerRow, candidates) {
    const normalizedCandidates = candidates.map(normalizeHeader);
    return headerRow.findIndex((header) => normalizedCandidates.includes(normalizeHeader(header)));
  }

  function normalizeHeader(value) {
    return String(value).trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  function readColumn(row, index) {
    return index >= 0 ? row[index] || '' : '';
  }

  function getVillageStats(rows) {
    const cards = uniqueCount(rows, 'Ration Card No.');
    const fps = uniqueCount(rows, 'FPS');
    const schemes = uniqueCount(rows, 'Scheme');
    const female = rows.filter((row) => row.Gender.toLowerCase() === 'female').length;
    const male = rows.filter((row) => row.Gender.toLowerCase() === 'male').length;

    return {
      members: rows.length,
      cards,
      fps,
      schemes,
      female,
      male,
    };
  }

  function uniqueCount(rows, key) {
    return new Set(rows.map((row) => row[key]).filter(Boolean)).size;
  }

  function bindFilters() {
    [...Object.values(searchInputs), schemeFilter, fpsFilter].forEach((control) => {
      control.addEventListener('input', renderActiveVillage);
      control.addEventListener('change', renderActiveVillage);
    });

    villageFilter.addEventListener('change', () => {
      setActiveVillage(villageFilter.value);
    });

    clearSearchButton.addEventListener('click', () => {
      resetSearchFields();
      schemeFilter.value = '';
      fpsFilter.value = '';
      renderActiveVillage();
    });

    tableBody.addEventListener('click', handleCopyClick);
  }

  function renderVillageFilter() {
    villageFilter.innerHTML = '';

    villages.forEach((village) => {
      const option = document.createElement('option');
      option.value = village.id;
      option.textContent = `${village.name} (${formatNumber(village.stats.members)})`;
      villageFilter.appendChild(option);
    });
  }

  function setActiveVillage(villageId) {
    activeVillageId = villageId;
    const village = getActiveVillage();

    if (!village) {
      return;
    }

    resetSearchFields();
    activeVillageName.textContent = village.name;
    activeVillageNote.textContent = `${formatNumber(village.stats.members)} members across ${formatNumber(
      village.stats.cards
    )} ration cards`;
    villageFilter.value = villageId;

    populateSelect(schemeFilter, village.rows, 'Scheme', 'All schemes');
    populateSelect(fpsFilter, village.rows, 'FPS', 'All FPS');
    renderSummary(village);
    renderActiveVillage();
  }

  function populateSelect(select, rows, key, defaultLabel) {
    const values = [...new Set(rows.map((row) => row[key]).filter(Boolean))].sort();
    select.innerHTML = `<option value="">${defaultLabel}</option>`;

    values.forEach((value) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
  }

  function renderSummary(village) {
    const stats = village.stats;
    const cardsLabel = stats.cards === 1 ? 'Ration card' : 'Ration cards';
    const fpsLabel = stats.fps === 1 ? 'FPS' : 'FPS units';
    const schemeLabel = stats.schemes === 1 ? 'Scheme' : 'Schemes';
    const tiles = [
      ['Members', stats.members],
      [cardsLabel, stats.cards],
      [fpsLabel, stats.fps],
      [schemeLabel, stats.schemes],
      ['Female', stats.female],
      ['Male', stats.male],
    ];

    summaryGrid.innerHTML = tiles
      .map(
        ([label, value]) => `
          <article class="summary-tile">
            <span>${escapeHtml(label)}</span>
            <strong>${formatNumber(value)}</strong>
          </article>
        `
      )
      .join('');
  }

  function renderActiveVillage() {
    const village = getActiveVillage();

    if (!village) {
      return;
    }

    const rows = getFilteredRows(village.rows);
    renderTable(TABLE_COLUMNS, rows);
    rowCount.textContent = `${formatNumber(rows.length)} / ${formatNumber(village.rows.length)} records`;
    filterHint.textContent = getFilterHint(rows.length, village.rows.length, village.name);

    if (!rows.length) {
      showStatus('No matching records for the selected village filters.', true);
    } else {
      showStatus('Village data loaded', false);
    }
  }

  function getFilteredRows(rows) {
    const filters = getSearchFilters();
    const scheme = schemeFilter.value;
    const fps = fpsFilter.value;

    return rows.filter((row) => {
      const matchesScheme = !scheme || row.Scheme === scheme;
      const matchesFps = !fps || row.FPS === fps;
      const matchesSearch =
        matchesColumn(row, 'S.No.', filters.serial) &&
        matchesColumn(row, 'Ration Card No.', filters.rationCard) &&
        matchesColumn(row, 'Family Head', filters.familyHead) &&
        matchesColumn(row, 'Member Name', filters.memberName) &&
        matchesUidLast4(row, filters.uidLast4);

      return matchesScheme && matchesFps && matchesSearch;
    });
  }

  function getSearchFilters() {
    return {
      serial: normalizeSearch(searchInputs.serial.value),
      rationCard: normalizeSearch(searchInputs.rationCard.value),
      familyHead: normalizeSearch(searchInputs.familyHead.value),
      memberName: normalizeSearch(searchInputs.memberName.value),
      uidLast4: onlyDigits(searchInputs.uidLast4.value).slice(-4),
    };
  }

  function matchesColumn(row, column, query) {
    return !query || normalizeSearch(row[column]).includes(query);
  }

  function matchesUidLast4(row, query) {
    if (!query) {
      return true;
    }

    return getLastFourDigits(row['UID No.']).includes(query);
  }

  function getLastFourDigits(value) {
    return onlyDigits(value).slice(-4);
  }

  function normalizeSearch(value) {
    return String(value || '').trim().toLowerCase();
  }

  function onlyDigits(value) {
    return String(value || '').replace(/\D/g, '');
  }

  function resetSearchFields() {
    Object.values(searchInputs).forEach((input) => {
      input.value = '';
    });
  }

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

        if (name === 'Ration Card No.') {
          renderCopyCell(td, row[name] || '');
        } else {
          td.textContent = row[name] || '';
        }

        tr.appendChild(td);
      });
      fragment.appendChild(tr);
    });

    tableBody.appendChild(fragment);
    table.style.display = 'table';
  }

  function renderCopyCell(cell, value) {
    cell.className = 'copy-cell';
    cell.dataset.copyValue = value;

    const valueText = document.createElement('span');
    valueText.className = 'copy-value';
    valueText.textContent = value;

    const button = createCopyButton(value);
    cell.append(valueText, button);
  }

  function createCopyButton(value) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'copy-button';
    button.dataset.copyValue = value;
    button.setAttribute('aria-label', `Copy ration card number ${value}`);
    button.title = 'Copy ration card number';
    button.textContent = 'Copy';
    return button;
  }

  async function handleCopyClick(event) {
    const button = event.target.closest('.copy-button');

    if (!button) {
      return;
    }

    const value = button.dataset.copyValue || '';
    if (!value) {
      return;
    }

    try {
      const result = await copyToClipboard(value, button);
      showCopyState(button, result === 'copied' ? 'Copied' : 'Selected');
    } catch (error) {
      console.error(error);
      showCopyState(button, 'Selected');
    }
  }

  async function copyToClipboard(value, button) {
    if (window.navigator?.clipboard?.writeText) {
      await window.navigator.clipboard.writeText(value);
      return 'copied';
    }

    const textArea = document.createElement('textarea');
    textArea.value = value;
    textArea.setAttribute('readonly', '');
    textArea.style.position = 'fixed';
    textArea.style.left = '-9999px';
    document.body.appendChild(textArea);
    textArea.select();
    const copied = document.execCommand('copy');
    textArea.remove();

    if (!copied) {
      selectCopyValue(button);
      return 'selected';
    }

    return 'copied';
  }

  function selectCopyValue(button) {
    const valueElement = button.closest('.copy-cell')?.querySelector('.copy-value');

    if (!valueElement || !window.getSelection || !document.createRange) {
      return;
    }

    const range = document.createRange();
    range.selectNodeContents(valueElement);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function showCopyState(button, label) {
    const originalLabel = button.textContent;
    button.textContent = label;
    button.classList.add('is-copied');

    window.setTimeout(() => {
      button.textContent = originalLabel;
      button.classList.remove('is-copied');
    }, 1200);
  }

  function getActiveVillage() {
    return villages.find((village) => village.id === activeVillageId);
  }

  function getFilterHint(visible, total, villageName) {
    if (visible === total) {
      return `Showing all beneficiary records for ${villageName}.`;
    }

    return `Filtered village view for ${villageName}.`;
  }

  function showStatus(message, isError) {
    statusMessage.textContent = message;
    statusMessage.classList.toggle('is-error', isError);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat('en-IN').format(value || 0);
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (char) => {
      const replacements = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;',
      };
      return replacements[char];
    });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
