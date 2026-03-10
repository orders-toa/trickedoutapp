// Navigation
document.querySelectorAll('.nav-item').forEach((item) => {
  item.addEventListener('click', () => {
    const pageId = item.dataset.page;

    document.querySelectorAll('.nav-item').forEach((i) => i.classList.remove('active'));
    document.querySelectorAll('.page').forEach((p) => p.classList.remove('active'));

    item.classList.add('active');
    document.getElementById(`page-${pageId}`)?.classList.add('active');

    if (pageId === 'page2') {
      loadStackRanker();
    }
    if (pageId === 'home') {
      loadHomePanels();
    }
  });
});

// State
let selectedTab = null;

const generateBtn = document.getElementById('generateBtn');
const codeBox = document.getElementById('codeBox');
const statusMsg = document.getElementById('statusMsg');
const tabSelector = document.getElementById('tabSelector');
const toggleStoreDrawerBtn = document.getElementById('toggleStoreDrawerBtn');
const closeStoreDrawerBtn = document.getElementById('closeStoreDrawerBtn');
const storeDrawerStoreSelect = document.getElementById('storeDrawerStoreSelect');
const storeDrawerEmployeeSelect = document.getElementById('storeDrawerEmployeeSelect');
const storeDrawerStatus = document.getElementById('storeDrawerStatus');
const employeeSalesNumber = document.getElementById('employeeSalesNumber');
const employeeHighlightsSummary = document.getElementById('employeeHighlightsSummary');
const employeeHighlightsList = document.getElementById('employeeHighlightsList');
const storeStatsSummary = document.getElementById('storeStatsSummary');
const storeStatsGrid = document.getElementById('storeStatsGrid');
const connDot = document.getElementById('connDot');
const connLabel = document.getElementById('connLabel');
const employeeNameInput = document.getElementById('employeeNameInput');
const customerNameInput = document.getElementById('customerNameInput');
const notesInput = document.getElementById('notesInput');
const stackRankerStatus = document.getElementById('stackRankerStatus');
const refreshStackRankerBtn = document.getElementById('refreshStackRankerBtn');
const suggestionNameInput = document.getElementById('suggestionNameInput');
const suggestionTextInput = document.getElementById('suggestionTextInput');
const suggestionTextLabel = document.getElementById('suggestionTextLabel');
const isShoutOutCheckbox = document.getElementById('isShoutOutCheckbox');
const shoutOutForGroup = document.getElementById('shoutOutForGroup');
const shoutOutForInput = document.getElementById('shoutOutForInput');
const shoutOutLocationGroup = document.getElementById('shoutOutLocationGroup');
const shoutOutLocationInput = document.getElementById('shoutOutLocationInput');
const suggestionStatus = document.getElementById('suggestionStatus');
const submitSuggestionBtn = document.getElementById('submitSuggestionBtn');
const homeShoutOutMessage = document.getElementById('homeShoutOutMessage');
const homeShoutOutMeta = document.getElementById('homeShoutOutMeta');
const homeShoutOutReactionCount = document.getElementById('homeShoutOutReactionCount');
const companyLeadersSummary = document.getElementById('companyLeadersSummary');
const companyLeadersGrid = document.getElementById('companyLeadersGrid');
const reactLikeBtn = document.getElementById('reactLikeBtn');
const reactLoveBtn = document.getElementById('reactLoveBtn');
const reactLaughBtn = document.getElementById('reactLaughBtn');
const reactClapBtn = document.getElementById('reactClapBtn');
const supplyItemGrid = document.getElementById('supplyItemGrid');
const addAllToCartBtn = document.getElementById('addAllToCartBtn');
const clearCatalogQtyBtn = document.getElementById('clearCatalogQtyBtn');
const otherRequestInput = document.getElementById('otherRequestInput');
const addOtherRequestBtn = document.getElementById('addOtherRequestBtn');
const openCartBtn = document.getElementById('openCartBtn');
const cartCountBadge = document.getElementById('cartCountBadge');
const supplyStepCatalog = document.getElementById('supplyStepCatalog');
const supplyStepCart = document.getElementById('supplyStepCart');
const supplyStepStore = document.getElementById('supplyStepStore');
const supplyStepSuccess = document.getElementById('supplyStepSuccess');
const cartEmptyMsg = document.getElementById('cartEmptyMsg');
const cartTable = document.getElementById('cartTable');
const cartTableBody = document.getElementById('cartTableBody');
const backToCatalogBtn = document.getElementById('backToCatalogBtn');
const proceedToStoreBtn = document.getElementById('proceedToStoreBtn');
const backToCartBtn = document.getElementById('backToCartBtn');
const submitSupplyOrderBtn = document.getElementById('submitSupplyOrderBtn');
const newSupplyOrderBtn = document.getElementById('newSupplyOrderBtn');
const supplySuccessMessage = document.getElementById('supplySuccessMessage');
const supplyStoreSelect = document.getElementById('supplyStoreSelect');
const orderSummary = document.getElementById('orderSummary');
const supplyCatalogStatus = document.getElementById('supplyCatalogStatus');
const supplyCartStatus = document.getElementById('supplyCartStatus');
const supplySubmitStatus = document.getElementById('supplySubmitStatus');

let isGenerating = false;
let shoutOutRotationTimer = null;
let shoutOutFadeTimer = null;
let shoutOutRotationIndex = 0;
let recentShoutOuts = [];
let storeEmployeeDirectory = [];
let highlightRequestToken = 0;
let supplyCart = {};
let supplyOtherRequests = [];
let nextOtherRequestId = 1;

const SUPPLY_ITEMS = [
  '70% Isopropyl',
  '99% Isopropyl',
  'Air Freshener Spray',
  'Avery Labels',
  'Baby Shampoo',
  'Blue Tape',
  'Box Cutter',
  'Canned Air',
  'Clorox Wipes',
  'Counterfeit Pen',
  'Hand Sanitizer',
  'Hand Soap',
  'Paper Towels',
  'Pens',
  'Printer Paper',
  'Rubber Bands',
  'Scissors',
  'Sharpies',
  'Spray Away Glass Cleaner',
  'Install Spray Bottle',
  'Squeegee',
  'Stainless Steel Spray',
  'Staples',
  'Sticky Notes',
  'Swiffer Duster Pad',
  'Swiffer Wet Jet Cleaner',
  'Swiffer Wet Jet Pads',
  'Toilet Bowl Cleaner',
  'Toilet Paper',
  'Toner (Printer Ink)',
  'Tooth Brush',
  'Trash Bags',
  'X-Acto Knife',
  'Guitar Picks',
  'TOA Black Bags',
  'TOA White Bags',
  'TOA White Phone Packaging',
  'Deposit Bags',
  'Deposit Slips',
];
const MAX_OTHER_REQUESTS = 10;

const reactionConfig = [
  { key: 'like', emoji: '👍' },
  { key: 'love', emoji: '❤️' },
  { key: 'laugh', emoji: '😂' },
  { key: 'clap', emoji: '👏' },
];

function setReactionButtonsDisabled(disabled) {
  [reactLikeBtn, reactLoveBtn, reactLaughBtn, reactClapBtn]
    .filter(Boolean)
    .forEach((btn) => {
      btn.disabled = disabled;
    });
}

function isNameValid(value) {
  return value.trim().length >= 4;
}

function isEmployeeValid(value) {
  return value.trim().length > 0;
}

function updateGenerateButtonState() {
  const validEmployee = isEmployeeValid(employeeNameInput?.value || '');
  const validCustomer = isNameValid(customerNameInput?.value || '');
  const ready = Boolean(selectedTab) && validEmployee && validCustomer && !isGenerating;
  if (generateBtn) generateBtn.disabled = !ready;
}

function setStoreDrawerStatus(msg, type = '') {
  if (!storeDrawerStatus) return;
  storeDrawerStatus.textContent = msg;
  storeDrawerStatus.style.color = type === 'error' ? 'var(--accent2)' : 'var(--muted)';
}

function setEmployeeHighlightsSummary(message, type = '') {
  if (!employeeHighlightsSummary) return;
  employeeHighlightsSummary.textContent = message;
  employeeHighlightsSummary.style.color = type === 'error' ? 'var(--accent2)' : 'var(--text)';
}

function clearEmployeeHighlightsList() {
  if (!employeeHighlightsList) return;
  employeeHighlightsList.innerHTML = '';
}

function formatMetricNumber(value) {
  return Number(value || 0).toLocaleString(undefined, {
    maximumFractionDigits: 2,
  });
}

function formatMoney(value) {
  return Number(value || 0).toLocaleString(undefined, {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  });
}

function formatPercentTenths(value) {
  const numeric = Number(value || 0);
  const rounded = Math.round(numeric * 10) / 10;
  return `${rounded.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: 1,
  })}%`;
}

function formatCompanyLeaderValue(item) {
  if (item.metricType === 'glassAttachmentRate') return formatPercentTenths(item.value);
  return formatMetricNumber(item.value);
}

function setCompanyLeadersSummary(message, type = '') {
  if (!companyLeadersSummary) return;
  companyLeadersSummary.textContent = message;
  companyLeadersSummary.style.color = type === 'error' ? 'var(--accent2)' : 'var(--text)';
}

function setEmployeeSalesNumber(value) {
  if (!employeeSalesNumber) return;
  const raw = Number(value);
  employeeSalesNumber.textContent = Number.isFinite(raw)
    ? raw.toLocaleString(undefined, { maximumFractionDigits: 0 })
    : '--';
}

function setStoreStatsSummary(message, type = '') {
  if (!storeStatsSummary) return;
  storeStatsSummary.textContent = message;
  storeStatsSummary.style.color = type === 'error' ? 'var(--accent2)' : 'var(--text)';
}

function clearStoreStatsGrid() {
  if (!storeStatsGrid) return;
  storeStatsGrid.innerHTML = '';
}

function renderStoreStatsCards(stats) {
  if (!storeStatsGrid) return;
  const cards = [
    {
      label: 'Products Sold',
      value: formatMetricNumber(stats.totalUnits),
      sub: `${formatMetricNumber(stats.invoiceCount)} invoices`,
    },
    {
      label: 'Store Rank',
      value: `#${stats.rank}`,
      sub: `Top ${stats.topPercent}% by GP/invoice`,
    },
    {
      label: 'Avg GP/Invoice',
      value: formatMoney(stats.avgGrossProfitPerInvoice),
      sub: 'Gross profit per invoice',
    },
    {
      label: 'Total Gross Profit',
      value: formatMoney(stats.totalNetProfit),
      sub: `Top product: ${String(stats.topProduct?.name || 'N/A')}`,
    },
  ];

  storeStatsGrid.innerHTML = cards
    .map((card) => `
      <div class="store-stat-card">
        <div class="store-stat-label">${escapeHtml(card.label)}</div>
        <div class="store-stat-value">${escapeHtml(card.value)}</div>
        <div class="store-stat-sub">${escapeHtml(card.sub)}</div>
      </div>
    `)
    .join('');
}

async function loadStoreStats(storeName) {
  if (!window.toaAPI?.getStoreStats) return;

  const selectedStore = String(storeName || '').trim();
  if (!selectedStore) {
    clearStoreStatsGrid();
    setStoreStatsSummary('Store Highlights');
    return;
  }

  clearStoreStatsGrid();
  setStoreStatsSummary('Store Highlights');

  try {
    const res = await window.toaAPI.getStoreStats({ storeName: selectedStore });
    if (!res.success) throw new Error(res.message || 'Failed to load store stats.');

    if (!res.stats) {
      clearStoreStatsGrid();
      setStoreStatsSummary('Store Highlights');
      return;
    }

    setStoreStatsSummary('Store Highlights');
    renderStoreStatsCards(res.stats);
  } catch (err) {
    clearStoreStatsGrid();
    setStoreStatsSummary(`Failed to load store stats: ${err.message}`, 'error');
  }
}

function renderEmployeeHighlights(items) {
  if (!employeeHighlightsList) return;
  employeeHighlightsList.innerHTML = items
    .map((item) => {
      const isGlassAttachmentRate = item.metricType === 'glassAttachmentRate';
      const rankText = isGlassAttachmentRate
        ? (item.isLeader
            ? 'Leads company glass attachment rate'
            : `Top ${item.topPercent}% in company glass attachment rate`)
        : (item.isLeader
            ? 'Leads company sales'
            : `Top ${item.topPercent}% in company sales`);
      const valueText = isGlassAttachmentRate
        ? formatPercentTenths(item.qty)
        : formatMetricNumber(item.qty);
      const detailText = isGlassAttachmentRate
        ? `${rankText} (${formatMetricNumber(item.numerator)}/${formatMetricNumber(item.denominator)})`
        : rankText;
      return `
        <div class="employee-highlight-item ${item.isLeader ? 'leader' : ''}">
          <div class="employee-highlight-name">${escapeHtml(item.product)}</div>
          <div class="employee-highlight-value">${escapeHtml(valueText)}</div>
          <div class="employee-highlight-meta">
            ${escapeHtml(detailText)}
          </div>
        </div>
      `;
    })
    .join('');
}

async function loadEmployeeHighlights(employeeName) {
  if (!window.toaAPI?.getEmployeeHighlights) return;

  const selectedName = String(employeeName || '').trim();
  if (!selectedName) {
    setEmployeeSalesNumber(null);
    clearEmployeeHighlightsList();
    setEmployeeHighlightsSummary('Highlights');
    return;
  }

  const currentToken = ++highlightRequestToken;
  clearEmployeeHighlightsList();
  setEmployeeHighlightsSummary('Highlights');

  try {
    const res = await window.toaAPI.getEmployeeHighlights({ employeeName: selectedName });
    if (currentToken !== highlightRequestToken) return;
    if (!res.success) throw new Error(res.message || 'Failed to load employee highlights.');

    const items = Array.isArray(res.items) ? res.items : [];
    setEmployeeSalesNumber(res.summary?.saleCount ?? 0);
    if (items.length === 0) {
      clearEmployeeHighlightsList();
      setEmployeeHighlightsSummary('Highlights');
      setStoreDrawerStatus(`No stats found for ${selectedName} in S/P Month to Date.`, 'error');
      return;
    }

    setEmployeeHighlightsSummary('Highlights');
    renderEmployeeHighlights(items);
  } catch (err) {
    setEmployeeSalesNumber(null);
    clearEmployeeHighlightsList();
    setEmployeeHighlightsSummary(`Failed to load highlights: ${err.message}`, 'error');
    setStoreDrawerStatus(`Stats failed: ${err.message}`, 'error');
  }
}

function populateEmployeeOptions(storeName) {
  if (!storeDrawerEmployeeSelect) return;

  const storeEntry = storeEmployeeDirectory.find((entry) => entry.store === storeName);
  const employees = storeEntry?.employees || [];

  if (!storeName) {
    storeDrawerEmployeeSelect.innerHTML = '<option value="">Select a store first...</option>';
    storeDrawerEmployeeSelect.disabled = true;
    setStoreDrawerStatus('Select a store to load employees.');
    if (employeeNameInput) employeeNameInput.value = '';
    loadEmployeeHighlights('');
    loadStoreStats('');
    updateGenerateButtonState();
    return;
  }

  if (employees.length === 0) {
    storeDrawerEmployeeSelect.innerHTML = '<option value="">No employees found for this store</option>';
    storeDrawerEmployeeSelect.disabled = true;
    setStoreDrawerStatus(`No employees found for ${storeName}.`, 'error');
    if (employeeNameInput) employeeNameInput.value = '';
    loadEmployeeHighlights('');
    loadStoreStats(storeName);
    updateGenerateButtonState();
    return;
  }

  storeDrawerEmployeeSelect.disabled = false;
  storeDrawerEmployeeSelect.innerHTML = `
    <option value="">Select an employee...</option>
    ${employees.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join('')}
  `;
  storeDrawerEmployeeSelect.value = '';
  setStoreDrawerStatus(`${employees.length} employees for ${storeName}.`);
  if (employeeNameInput) employeeNameInput.value = '';
  loadEmployeeHighlights('');
  loadStoreStats(storeName);
  updateGenerateButtonState();
}

async function loadStoreEmployeeDirectory() {
  if (!window.toaAPI?.getStoreEmployeeDirectory || !storeDrawerStoreSelect) return;

  storeDrawerStoreSelect.innerHTML = '<option value="">Loading stores...</option>';
  storeDrawerStoreSelect.disabled = true;
  if (storeDrawerEmployeeSelect) {
    storeDrawerEmployeeSelect.innerHTML = '<option value="">Select a store first...</option>';
    storeDrawerEmployeeSelect.disabled = true;
  }

  try {
    const res = await window.toaAPI.getStoreEmployeeDirectory();
    if (!res.success) throw new Error(res.message || 'Failed to load store directory.');

    storeEmployeeDirectory = Array.isArray(res.directory) ? res.directory : [];
    if (storeEmployeeDirectory.length === 0) {
      storeDrawerStoreSelect.innerHTML = '<option value="">No stores found</option>';
      setStoreDrawerStatus('No store/employee rows found in columns B:C.', 'error');
      return;
    }

    storeDrawerStoreSelect.innerHTML = `
      <option value="">Select a store...</option>
      ${storeEmployeeDirectory
        .map((entry) => `<option value="${escapeHtml(entry.store)}">${escapeHtml(entry.store)}</option>`)
        .join('')}
    `;
    storeDrawerStoreSelect.disabled = false;
    setStoreDrawerStatus(`Loaded ${storeEmployeeDirectory.length} stores. Select one to continue.`);
  } catch (err) {
    storeDrawerStoreSelect.innerHTML = '<option value="">Failed to load stores</option>';
    setStoreDrawerStatus(`Failed: ${err.message}`, 'error');
  }
}

// Load tabs from Google Sheet
async function loadTabs() {
  try {
    const res = await window.toaAPI.getTabs();
    if (!res.success) throw new Error(res.message);

    connDot.classList.add('connected');
    connLabel.textContent = 'Connected';

    tabSelector.innerHTML = '';

    res.tabs.forEach((tab) => {
      const btn = document.createElement('button');
      btn.className = 'tab-btn';
      btn.dataset.tab = tab;
      btn.innerHTML = `<span class="dot"></span>${tab}`;

      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach((b) => b.classList.remove('selected'));
        btn.classList.add('selected');
        selectedTab = tab;
        codeBox.innerHTML = `<div class="code-placeholder">Ready to generate a ${tab} code</div>`;
        setStatus('', '');
        updateGenerateButtonState();
      });

      tabSelector.appendChild(btn);
    });
  } catch (err) {
    connDot.classList.add('error');
    connLabel.textContent = 'Disconnected';
    tabSelector.innerHTML = `<div style="color:#ef4444;font-size:12px;">Failed to load: ${err.message}</div>`;
  }
}

// Generate code (and immediately mark as used)
generateBtn?.addEventListener('click', async () => {
  if (!selectedTab || isGenerating) return;

  const employeeName = employeeNameInput.value.trim();
  const customerName = customerNameInput.value.trim();
  const notes = notesInput.value.trim();

  if (!isEmployeeValid(employeeName) || !isNameValid(customerName)) {
    setStatus('Select an employee and enter a customer name (at least 4 characters).', 'error');
    updateGenerateButtonState();
    return;
  }

  isGenerating = true;
  updateGenerateButtonState();
  generateBtn.innerHTML = '<span class="spinner"></span>Generating...';
  setStatus('Generating and logging code...', 'info');

  try {
    const res = await window.toaAPI.getNextCode(selectedTab);
    if (!res.success) {
      setStatus(res.message, 'error');
      return;
    }

    const code = res.code;
    const saveRes = await window.toaAPI.markCodeUsed({
      code,
      tabName: selectedTab,
      employeeName,
      customerName,
      notes,
    });

    if (!saveRes.success) {
      setStatus(saveRes.message, 'error');
      return;
    }

    codeBox.innerHTML = `
      <div class="code-value">${code}</div>
      <div class="code-type-badge">${selectedTab} Code</div>
    `;

    setStatus(`Code generated and logged for ${customerName}.`, 'success');
    notesInput.value = '';
  } catch (err) {
    setStatus('Error: ' + err.message, 'error');
  } finally {
    isGenerating = false;
    generateBtn.innerHTML = 'Generate Code';
    updateGenerateButtonState();
  }
});

async function loadStackRanker() {
  const wrap = document.querySelector('.stack-ranker-wrap');
  if (!wrap || !stackRankerStatus || !refreshStackRankerBtn) return;

  stackRankerStatus.textContent = 'Loading Stack Ranker...';
  wrap.innerHTML = '<table class="stack-ranker-table" id="stackRankerTable"></table>';
  refreshStackRankerBtn.disabled = true;
  refreshStackRankerBtn.innerHTML = '<span class="spinner"></span>Loading';

  try {
    const res = await window.toaAPI.getStackRanker();
    if (!res.success) throw new Error(res.message || 'Failed to load Stack Ranker');

    const rows = res.values || [];
    const styledCells = res.cells || [];
    const images = res.images || [];

    if (images.length > 0) {
      const imageCards = images
        .map(
          (url) => `
            <div class="stack-ranker-image-card">
              <img src="${escapeHtml(url)}" alt="Stack Ranker image" loading="lazy" />
            </div>
          `
        )
        .join('');

      wrap.innerHTML = `<div class="stack-ranker-images" id="stackRankerTable">${imageCards}</div>`;
      const sourceName = res.sheetName || 'Stack Ranker';
      stackRankerStatus.textContent = `Loaded ${images.length} image(s) from ${sourceName}.`;
      return;
    }
    const tableEl = document.getElementById('stackRankerTable');
    if (rows.length === 0) {
      tableEl.innerHTML =
        '<tbody><tr><td class="stack-ranker-empty">No data found in Stack Ranker tab.</td></tr></tbody>';
      stackRankerStatus.textContent = 'No data available.';
      return;
    }

    const hiddenColumnIndexes = new Set([0, 12, 13]); // A, M, N
    const maxCols = Math.max(...rows.map((r) => r.length), 1);
    const normalized = rows.map((r) => {
      const row = [...r];
      while (row.length < maxCols) row.push('');
      return row;
    });

    const normalizedStyles = styledCells.map((row) => {
      const styleRow = [...row];
      while (styleRow.length < maxCols) styleRow.push({ text: '', style: {} });
      return styleRow;
    });

    const tbodyRows = normalized
      .slice(1) // Hide sheet row 1 in UI only.
      .map((row, rowIndex) => {
        const rowStyles = normalizedStyles[rowIndex + 1] || [];
        const cellsHtml = row
          .map((cell, cellIndex) => ({ cell, cellIndex }))
          .filter(({ cellIndex }) => !hiddenColumnIndexes.has(cellIndex))
          .map(({ cell, cellIndex }) => {
            const style = styleToInline(rowStyles[cellIndex]?.style);
            return `<td${style ? ` style="${style}"` : ''}>${escapeHtml(cell)}</td>`;
          })
          .join('');
        return `<tr>${cellsHtml}</tr>`;
      })
      .join('');
    tableEl.innerHTML = `<tbody>${tbodyRows}</tbody>`;
    const sourceName = res.sheetName || 'Stack Ranker';
    stackRankerStatus.textContent = `Loaded ${rows.length} rows from ${sourceName}.`;
  } catch (err) {
    stackRankerStatus.textContent = `Failed to load Stack Ranker: ${err.message}`;
    const tableEl = document.getElementById('stackRankerTable');
    if (tableEl) tableEl.innerHTML = '';
  } finally {
    refreshStackRankerBtn.disabled = false;
    refreshStackRankerBtn.innerHTML = 'Refresh';
  }
}

function setStatus(msg, type) {
  statusMsg.textContent = msg;
  statusMsg.className = `status-msg ${type}`;
}

function setSuggestionStatus(msg, type = '') {
  if (!suggestionStatus) return;
  suggestionStatus.textContent = msg;
  suggestionStatus.className = `suggestion-status ${type}`.trim();
}

function updateSuggestionMode() {
  const isShoutOut = Boolean(isShoutOutCheckbox?.checked);
  if (shoutOutForGroup) {
    shoutOutForGroup.classList.toggle('suggestion-hidden', !isShoutOut);
  }
  if (shoutOutLocationGroup) {
    shoutOutLocationGroup.classList.toggle('suggestion-hidden', !isShoutOut);
  }
  if (suggestionTextLabel) {
    suggestionTextLabel.textContent = isShoutOut ? 'Shout Out Message' : 'Suggestion';
  }
  if (suggestionTextInput) {
    suggestionTextInput.placeholder = isShoutOut
      ? 'Write the shout out you want to send...'
      : 'What should we add or improve?';
  }
  if (submitSuggestionBtn) {
    submitSuggestionBtn.textContent = isShoutOut ? 'Submit Shout Out' : 'Submit Suggestion';
  }
}

function stopShoutOutRotation() {
  if (shoutOutRotationTimer) {
    clearInterval(shoutOutRotationTimer);
    shoutOutRotationTimer = null;
  }
  if (shoutOutFadeTimer) {
    clearTimeout(shoutOutFadeTimer);
    shoutOutFadeTimer = null;
  }
}

function renderCurrentShoutOut(withFade = false) {
  if (!homeShoutOutMessage || !homeShoutOutMeta) return;

  if (recentShoutOuts.length === 0) {
    homeShoutOutMessage.textContent =
      'Gordon received a shout-out for creating such an awesome in-house application to simplify our lives at Tricked Out!';
    homeShoutOutMeta.textContent = '';
    if (homeShoutOutReactionCount) homeShoutOutReactionCount.textContent = '';
    setReactionButtonsDisabled(true);
    return;
  }

  const item = recentShoutOuts[shoutOutRotationIndex % recentShoutOuts.length];
  const applyContent = () => {
    const recognizedName = item.shoutOutFor?.trim() || 'A team member';
    const writtenBy = item.fromName?.trim() || 'Anonymous';
    homeShoutOutMessage.textContent = `${recognizedName} was recognized!`;

    let details = item.message || '';
    if (item.location) {
      details += `\nLocation: ${item.location}`;
    }
    details += `\n- ${writtenBy}`;
    homeShoutOutMeta.textContent = details;
    const reactions = item.reactions || {};
    if (homeShoutOutReactionCount) {
      const chips = reactionConfig
        .map(({ key, emoji }) => {
          const count = Number.isFinite(reactions[key]) ? reactions[key] : 0;
          if (count <= 0) return '';
          return `<span class="reaction-chip">${emoji} ${count}</span>`;
        })
        .filter(Boolean)
        .join('');
      homeShoutOutReactionCount.innerHTML = chips || '<span class="reaction-chip">No reactions yet</span>';
    }
    setReactionButtonsDisabled(false);
  };

  if (!withFade) {
    applyContent();
    return;
  }

  homeShoutOutMessage.classList.add('fade-out');
  homeShoutOutMeta.classList.add('fade-out');
  if (shoutOutFadeTimer) clearTimeout(shoutOutFadeTimer);
  shoutOutFadeTimer = setTimeout(() => {
    applyContent();
    homeShoutOutMessage.classList.remove('fade-out');
    homeShoutOutMeta.classList.remove('fade-out');
  }, 470);
}

function startShoutOutRotation() {
  stopShoutOutRotation();
  if (recentShoutOuts.length <= 1) return;

  shoutOutRotationTimer = setInterval(() => {
    shoutOutRotationIndex = (shoutOutRotationIndex + 1) % recentShoutOuts.length;
    renderCurrentShoutOut(true);
  }, 15000);
}

async function loadRecentShoutOuts() {
  if (!homeShoutOutMessage || !homeShoutOutMeta) return;

  homeShoutOutMessage.textContent = 'Loading shout outs...';
  homeShoutOutMeta.textContent = '';
  stopShoutOutRotation();

  try {
    const res = await window.toaAPI.getRecentShoutOuts();
    if (!res.success) throw new Error(res.message || 'Failed to load shout outs.');

    recentShoutOuts = res.shoutOuts || [];
    shoutOutRotationIndex = 0;
    renderCurrentShoutOut();
    startShoutOutRotation();
  } catch (err) {
    homeShoutOutMessage.textContent = `Failed to load shout outs: ${err.message}`;
    homeShoutOutMeta.textContent = '';
  }
}

function renderCompanyLeaders(items) {
  if (!companyLeadersGrid) return;
  companyLeadersGrid.innerHTML = items
    .map((item) => {
      const locationText = String(item.location || '').trim();
      const nameText = String(item.employeeName || '').trim();
      const detailText = item.metricType === 'glassAttachmentRate'
        ? `${item.metricName} (${formatMetricNumber(item.numerator)}/${formatMetricNumber(item.denominator)})`
        : item.metricName;
      return `
        <div class="store-stat-card">
          <div class="store-stat-label">${escapeHtml(detailText)}</div>
          <div class="store-stat-value">${escapeHtml(formatCompanyLeaderValue(item))}</div>
          <div class="store-stat-sub">
            <div class="company-leader-name">${escapeHtml(nameText)}</div>
            <div class="company-leader-location">${escapeHtml(locationText)}</div>
          </div>
        </div>
      `;
    })
    .join('');
}

async function loadCompanyLeaders() {
  if (!window.toaAPI?.getCompanyLeaders || !companyLeadersSummary || !companyLeadersGrid) return;

  setCompanyLeadersSummary('Company Leaders');
  companyLeadersGrid.innerHTML = '';

  try {
    const res = await window.toaAPI.getCompanyLeaders();
    if (!res.success) throw new Error(res.message || 'Failed to load company leaders.');

    const leaders = Array.isArray(res.leaders) ? res.leaders : [];
    if (leaders.length === 0) {
      setCompanyLeadersSummary('Company Leaders');
      return;
    }

    setCompanyLeadersSummary('Company Leaders');
    renderCompanyLeaders(leaders);
  } catch (err) {
    setCompanyLeadersSummary(`Failed to load leaders: ${err.message}`, 'error');
    companyLeadersGrid.innerHTML = '';
  }
}

function loadHomePanels() {
  loadRecentShoutOuts();
  loadCompanyLeaders();
}

function setSupplyStatus(target, msg, type = '') {
  if (!target) return;
  target.textContent = msg;
  target.className = `supply-status ${type}`.trim();
}

function getSupplyCartTotalQty() {
  const standardTotal = Object.values(supplyCart).reduce((sum, qty) => sum + qty, 0);
  const otherTotal = supplyOtherRequests.reduce((sum, row) => sum + row.qty, 0);
  return standardTotal + otherTotal;
}

function updateCartBadge() {
  if (!cartCountBadge) return;
  cartCountBadge.textContent = String(getSupplyCartTotalQty());
}

function showSupplyStep(step) {
  [supplyStepCatalog, supplyStepCart, supplyStepStore, supplyStepSuccess]
    .filter(Boolean)
    .forEach((el) => el.classList.add('hidden'));

  if (step === 'catalog') supplyStepCatalog?.classList.remove('hidden');
  if (step === 'cart') supplyStepCart?.classList.remove('hidden');
  if (step === 'store') supplyStepStore?.classList.remove('hidden');
  if (step === 'success') supplyStepSuccess?.classList.remove('hidden');
}

function renderSupplyCatalog() {
  if (!supplyItemGrid) return;
  supplyItemGrid.innerHTML = SUPPLY_ITEMS.map((item) => `
    <div class="supply-item-card">
      <div class="supply-item-name">${escapeHtml(item)}</div>
      <input
        type="number"
        min="0"
        step="1"
        class="supply-qty-input catalog-qty-input"
        data-item="${escapeHtml(item)}"
        value="0"
      />
    </div>
  `).join('');
}

function getCatalogQuantities() {
  const quantities = {};
  document.querySelectorAll('.catalog-qty-input').forEach((input) => {
    const item = String(input.dataset.item || '');
    const value = Number.parseInt(input.value, 10);
    if (!item) return;
    quantities[item] = Number.isNaN(value) ? 0 : Math.max(0, value);
  });
  return quantities;
}

function renderCart() {
  if (!cartTableBody || !cartEmptyMsg || !cartTable) return;

  const standardEntries = Object.entries(supplyCart)
    .filter(([, qty]) => qty > 0)
    .sort(([a], [b]) => a.localeCompare(b));
  const otherEntries = supplyOtherRequests.filter((row) => row.qty > 0);

  if (standardEntries.length === 0 && otherEntries.length === 0) {
    cartEmptyMsg.style.display = '';
    cartTable.style.display = 'none';
    cartTableBody.innerHTML = '';
    setSupplyStatus(supplyCartStatus, '');
    updateCartBadge();
    return;
  }

  cartEmptyMsg.style.display = 'none';
  cartTable.style.display = '';
  const standardRows = standardEntries.map(([item, qty]) => `
      <tr>
        <td>${escapeHtml(item)}</td>
        <td>
          <input
            type="number"
            min="0"
            step="1"
            class="supply-qty-input cart-qty-input"
            data-kind="standard"
            data-key="${escapeHtml(item)}"
            value="${qty}"
          />
        </td>
        <td>
          <button class="btn btn-secondary cart-remove-btn" data-kind="standard" data-key="${escapeHtml(item)}">Remove</button>
        </td>
      </tr>
    `)
    .join('');
  const otherRows = otherEntries.map((row) => `
      <tr>
        <td>Other: ${escapeHtml(row.description)}</td>
        <td>
          <input
            type="number"
            min="0"
            step="1"
            class="supply-qty-input cart-qty-input"
            data-kind="other"
            data-key="${row.id}"
            value="${row.qty}"
          />
        </td>
        <td>
          <button class="btn btn-secondary cart-remove-btn" data-kind="other" data-key="${row.id}">Remove</button>
        </td>
      </tr>
    `).join('');
  cartTableBody.innerHTML = `${standardRows}${otherRows}`;
  updateCartBadge();
}

function renderOrderSummary() {
  if (!orderSummary) return;

  const standardEntries = Object.entries(supplyCart).filter(([, qty]) => qty > 0);
  const otherEntries = supplyOtherRequests.filter((row) => row.qty > 0);
  if (standardEntries.length === 0 && otherEntries.length === 0) {
    orderSummary.textContent = 'No items selected.';
    return;
  }

  const standardSummary = standardEntries
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([item, qty]) => `${item}: ${qty}`)
    .join(' | ');
  const otherSummary = otherEntries
    .map((row) => `Other (${row.description}): ${row.qty}`)
    .join(' | ');
  const fullSummary = [standardSummary, otherSummary].filter(Boolean).join(' | ');
  orderSummary.textContent = `Items: ${fullSummary}`;
}

function resetSupplyOrderFlow() {
  supplyCart = {};
  supplyOtherRequests = [];
  nextOtherRequestId = 1;
  document.querySelectorAll('.catalog-qty-input').forEach((input) => {
    input.value = '0';
  });
  if (otherRequestInput) otherRequestInput.value = '';
  if (supplyStoreSelect) supplyStoreSelect.value = '';
  if (supplySuccessMessage) supplySuccessMessage.textContent = 'Your order was submitted.';
  setSupplyStatus(supplyCatalogStatus, '');
  setSupplyStatus(supplyCartStatus, '');
  setSupplyStatus(supplySubmitStatus, '');
  renderCart();
  renderOrderSummary();
  showSupplyStep('catalog');
}

function initSupplyOrders() {
  if (!supplyItemGrid) return;

  renderSupplyCatalog();
  renderCart();
  renderOrderSummary();
  updateCartBadge();

  openCartBtn?.addEventListener('click', () => {
    renderCart();
    showSupplyStep('cart');
  });

  addAllToCartBtn?.addEventListener('click', () => {
    const source = getCatalogQuantities();
    let addedItems = 0;

    SUPPLY_ITEMS.forEach((item) => {
      const qty = source[item] || 0;
      if (qty <= 0) return;
      supplyCart[item] = (supplyCart[item] || 0) + qty;
      addedItems += 1;
    });

    if (addedItems === 0) {
      setSupplyStatus(supplyCatalogStatus, 'Enter quantities greater than 0 before adding to cart.', 'error');
      return;
    }

    setSupplyStatus(supplyCatalogStatus, `Added ${addedItems} item types to cart.`, 'success');
    renderCart();
    renderOrderSummary();
  });

  clearCatalogQtyBtn?.addEventListener('click', () => {
    document.querySelectorAll('.catalog-qty-input').forEach((input) => {
      input.value = '0';
    });
    setSupplyStatus(supplyCatalogStatus, 'Catalog quantities cleared.');
  });

  addOtherRequestBtn?.addEventListener('click', () => {
    const description = String(otherRequestInput?.value || '').trim();
    if (description.length < 2) {
      setSupplyStatus(supplyCatalogStatus, 'Enter a description for the other request.', 'error');
      return;
    }
    if (supplyOtherRequests.length >= MAX_OTHER_REQUESTS) {
      setSupplyStatus(supplyCatalogStatus, `Only ${MAX_OTHER_REQUESTS} other requests can be added per order.`, 'error');
      return;
    }

    supplyOtherRequests.push({
      id: nextOtherRequestId++,
      description,
      qty: 1,
    });
    if (otherRequestInput) otherRequestInput.value = '';
    setSupplyStatus(supplyCatalogStatus, 'Other request added to cart.', 'success');
    renderCart();
    renderOrderSummary();
  });

  backToCatalogBtn?.addEventListener('click', () => {
    showSupplyStep('catalog');
  });

  proceedToStoreBtn?.addEventListener('click', () => {
    if (getSupplyCartTotalQty() <= 0) {
      setSupplyStatus(supplyCartStatus, 'Your cart is empty.', 'error');
      return;
    }
    setSupplyStatus(supplyCartStatus, '');
    renderOrderSummary();
    showSupplyStep('store');
  });

  backToCartBtn?.addEventListener('click', () => {
    renderCart();
    showSupplyStep('cart');
  });

  cartTableBody?.addEventListener('input', (event) => {
    const target = event.target;
    if (!target?.classList?.contains('cart-qty-input')) return;
    const kind = String(target.dataset.kind || '');
    const key = String(target.dataset.key || '');
    const value = Number.parseInt(target.value, 10);
    const qty = Number.isNaN(value) ? 0 : Math.max(0, value);
    if (kind === 'standard') {
      if (!key) return;
      if (qty <= 0) delete supplyCart[key];
      else supplyCart[key] = qty;
    } else if (kind === 'other') {
      const id = Number.parseInt(key, 10);
      if (!Number.isInteger(id)) return;
      const row = supplyOtherRequests.find((entry) => entry.id === id);
      if (!row) return;
      if (qty <= 0) {
        supplyOtherRequests = supplyOtherRequests.filter((entry) => entry.id !== id);
      } else {
        row.qty = qty;
      }
    } else {
      return;
    }
    renderCart();
    renderOrderSummary();
  });

  cartTableBody?.addEventListener('click', (event) => {
    const target = event.target;
    if (!target?.classList?.contains('cart-remove-btn')) return;
    const kind = String(target.dataset.kind || '');
    const key = String(target.dataset.key || '');
    if (kind === 'standard') {
      if (!key) return;
      delete supplyCart[key];
    } else if (kind === 'other') {
      const id = Number.parseInt(key, 10);
      if (!Number.isInteger(id)) return;
      supplyOtherRequests = supplyOtherRequests.filter((entry) => entry.id !== id);
    } else {
      return;
    }
    renderCart();
    renderOrderSummary();
  });

  submitSupplyOrderBtn?.addEventListener('click', async () => {
    if (!window.toaAPI?.submitSupplyOrder) {
      setSupplyStatus(supplySubmitStatus, 'This app version does not support supply orders yet.', 'error');
      return;
    }

    if (getSupplyCartTotalQty() <= 0) {
      setSupplyStatus(supplySubmitStatus, 'Your cart is empty.', 'error');
      return;
    }

    const store = String(supplyStoreSelect?.value || '').trim();
    if (!store) {
      setSupplyStatus(supplySubmitStatus, 'Please select a store location.', 'error');
      return;
    }

    submitSupplyOrderBtn.disabled = true;
    submitSupplyOrderBtn.innerHTML = '<span class="spinner"></span>Submitting';
    setSupplyStatus(supplySubmitStatus, 'Submitting order...');

    try {
      const res = await window.toaAPI.submitSupplyOrder({
        store,
        quantities: supplyCart,
        otherRequests: supplyOtherRequests,
      });

      if (!res.success) {
        setSupplyStatus(supplySubmitStatus, res.message || 'Failed to submit order.', 'error');
        return;
      }

      if (supplySuccessMessage) {
        const submittedAt = new Date().toLocaleString();
        supplySuccessMessage.textContent = `Your order was submitted on ${submittedAt}.`;
      }
      setSupplyStatus(supplySubmitStatus, 'Order submitted.', 'success');
      showSupplyStep('success');
    } catch (err) {
      setSupplyStatus(supplySubmitStatus, `Failed to submit order: ${err.message}`, 'error');
    } finally {
      submitSupplyOrderBtn.disabled = false;
      submitSupplyOrderBtn.innerHTML = 'Submit Order';
    }
  });

  newSupplyOrderBtn?.addEventListener('click', () => {
    resetSupplyOrderFlow();
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function styleToInline(styleObj) {
  if (!styleObj) return '';

  const rules = [];
  if (styleObj.backgroundColor) rules.push(`background-color:${styleObj.backgroundColor}`);
  if (styleObj.color) rules.push(`color:${styleObj.color}`);
  if (styleObj.fontSize) rules.push(`font-size:${styleObj.fontSize}`);
  if (styleObj.fontWeight) rules.push(`font-weight:${styleObj.fontWeight}`);
  if (styleObj.fontStyle) rules.push(`font-style:${styleObj.fontStyle}`);
  if (styleObj.textDecoration) rules.push(`text-decoration:${styleObj.textDecoration}`);
  if (styleObj.textAlign) rules.push(`text-align:${styleObj.textAlign}`);
  if (styleObj.verticalAlign) rules.push(`vertical-align:${styleObj.verticalAlign}`);
  if (styleObj.borderTop) rules.push(`border-top:${styleObj.borderTop}`);
  if (styleObj.borderRight) rules.push(`border-right:${styleObj.borderRight}`);
  if (styleObj.borderBottom) rules.push(`border-bottom:${styleObj.borderBottom}`);
  if (styleObj.borderLeft) rules.push(`border-left:${styleObj.borderLeft}`);

  return rules.join(';');
}

customerNameInput?.addEventListener('input', updateGenerateButtonState);
refreshStackRankerBtn?.addEventListener('click', loadStackRanker);
toggleStoreDrawerBtn?.addEventListener('click', () => {
  document.body.classList.toggle('store-drawer-open');
});
closeStoreDrawerBtn?.addEventListener('click', () => {
  document.body.classList.remove('store-drawer-open');
});
storeDrawerStoreSelect?.addEventListener('change', () => {
  const selectedStore = String(storeDrawerStoreSelect.value || '');
  populateEmployeeOptions(selectedStore);
});
storeDrawerEmployeeSelect?.addEventListener('change', () => {
  const selectedEmployee = String(storeDrawerEmployeeSelect.value || '').trim();
  if (employeeNameInput) employeeNameInput.value = selectedEmployee;
  if (selectedEmployee) {
    setStoreDrawerStatus(`Loading stats for ${selectedEmployee}...`);
  } else {
    setStoreDrawerStatus('Select an employee to view stats.');
  }
  loadEmployeeHighlights(selectedEmployee);
  updateGenerateButtonState();
});
submitSuggestionBtn?.addEventListener('click', async () => {
  const suggestion = suggestionTextInput?.value.trim() || '';
  const name = suggestionNameInput?.value.trim() || '';
  const isShoutOut = Boolean(isShoutOutCheckbox?.checked);
  const shoutOutFor = shoutOutForInput?.value.trim() || '';
  const location = shoutOutLocationInput?.value.trim() || '';
  const entryType = isShoutOut ? 'Shout Out' : 'Suggestion';

  if (suggestion.length < 4) {
    setSuggestionStatus(`Please enter at least 4 characters for your ${entryType.toLowerCase()}.`, 'error');
    return;
  }
  if (isShoutOut && shoutOutFor.length < 2) {
    setSuggestionStatus('Please enter who the shout out is for.', 'error');
    return;
  }

  submitSuggestionBtn.disabled = true;
  submitSuggestionBtn.innerHTML = '<span class="spinner"></span>Submitting';
  setSuggestionStatus(`Submitting ${entryType.toLowerCase()}...`);

  try {
    const res = await window.toaAPI.submitSuggestion({
      name,
      suggestion,
      type: entryType,
      shoutOutFor,
      location,
    });
    if (!res.success) {
      setSuggestionStatus(res.message || `Failed to submit ${entryType.toLowerCase()}.`, 'error');
      return;
    }

    setSuggestionStatus(`${entryType} sent. Thank you!`, 'success');
    if (suggestionTextInput) suggestionTextInput.value = '';
    if (shoutOutForInput) shoutOutForInput.value = '';
    if (shoutOutLocationInput) shoutOutLocationInput.value = '';
    if (isShoutOut) loadRecentShoutOuts();
  } catch (err) {
    setSuggestionStatus(`Failed to submit ${entryType.toLowerCase()}: ${err.message}`, 'error');
  } finally {
    submitSuggestionBtn.disabled = false;
    submitSuggestionBtn.innerHTML = isShoutOut ? 'Submit Shout Out' : 'Submit Suggestion';
  }
});
isShoutOutCheckbox?.addEventListener('change', updateSuggestionMode);
async function submitShoutOutReaction(reactionType) {
  if (recentShoutOuts.length === 0) return;

  const item = recentShoutOuts[shoutOutRotationIndex % recentShoutOuts.length];
  if (!item?.rowNumber) return;

  setReactionButtonsDisabled(true);

  try {
    const res = await window.toaAPI.reactToShoutOut({ rowNumber: item.rowNumber, reactionType });
    if (!res.success) throw new Error(res.message || 'Failed to add reaction.');

    item.reactions = res.reactions || item.reactions || {};
    renderCurrentShoutOut();
  } catch (err) {
    if (homeShoutOutReactionCount) {
      homeShoutOutReactionCount.textContent = `Reaction failed: ${err.message}`;
    }
  } finally {
    if (recentShoutOuts.length > 0) setReactionButtonsDisabled(false);
  }
}

reactLikeBtn?.addEventListener('click', () => submitShoutOutReaction('like'));
reactLoveBtn?.addEventListener('click', () => submitShoutOutReaction('love'));
reactLaughBtn?.addEventListener('click', () => submitShoutOutReaction('laugh'));
reactClapBtn?.addEventListener('click', () => submitShoutOutReaction('clap'));

// Init
loadTabs();
loadStoreEmployeeDirectory();
updateGenerateButtonState();
updateSuggestionMode();
loadHomePanels();
initSupplyOrders();
