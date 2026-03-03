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
      loadRecentShoutOuts();
    }
  });
});

// State
let selectedTab = null;

const generateBtn = document.getElementById('generateBtn');
const codeBox = document.getElementById('codeBox');
const statusMsg = document.getElementById('statusMsg');
const tabSelector = document.getElementById('tabSelector');
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
const reactLikeBtn = document.getElementById('reactLikeBtn');
const reactLoveBtn = document.getElementById('reactLoveBtn');
const reactLaughBtn = document.getElementById('reactLaughBtn');
const reactClapBtn = document.getElementById('reactClapBtn');

let isGenerating = false;
let shoutOutRotationTimer = null;
let shoutOutFadeTimer = null;
let shoutOutRotationIndex = 0;
let recentShoutOuts = [];

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

function updateGenerateButtonState() {
  const validEmployee = isNameValid(employeeNameInput?.value || '');
  const validCustomer = isNameValid(customerNameInput?.value || '');
  const ready = Boolean(selectedTab) && validEmployee && validCustomer && !isGenerating;
  if (generateBtn) generateBtn.disabled = !ready;
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

  if (!isNameValid(employeeName) || !isNameValid(customerName)) {
    setStatus('Employee Name and Customer Name must each be at least 4 characters.', 'error');
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

employeeNameInput?.addEventListener('input', updateGenerateButtonState);
customerNameInput?.addEventListener('input', updateGenerateButtonState);
refreshStackRankerBtn?.addEventListener('click', loadStackRanker);
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
updateGenerateButtonState();
updateSuggestionMode();
loadRecentShoutOuts();
