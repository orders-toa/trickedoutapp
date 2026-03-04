const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { google } = require('googleapis');
const { autoUpdater } = require('electron-updater');

// ── Google Sheets config ──────────────────────────────────────────────────────
const SHEET_ID = '1nywGhJ50rrntwoGwO9zgPaVOsEpHM1iTOG7O_pXHKrQ';
const STACK_RANKER_SHEET_ID = '1oKJbQA12JIvNQdjNiWmd3jb2pmQytuxbRh_snGWvbRk';
const STACK_RANKER_TAB_NAME = 'Stack Ranker';
const SUGGESTIONS_TAB_NAME = 'Suggestions';
const SUPPLY_ORDERS_TAB_NAME = 'Supply Orders';
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');

const SUPPLY_ORDER_ITEMS = [
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

const SUPPLY_ORDER_LOCATIONS = [
  'Magic Valley',
  'Grand Teton',
  'Newgate',
  'Layton',
  'Station Park',
  'Valley Fair',
  'South Town',
  'University',
  'Provo',
];
const MAX_OTHER_REQUESTS = 10;

function rgbToCss(color) {
  if (!color) return null;
  const r = Math.round((color.red || 0) * 255);
  const g = Math.round((color.green || 0) * 255);
  const b = Math.round((color.blue || 0) * 255);
  const a = color.alpha;
  if (typeof a === 'number' && a < 1) return `rgba(${r}, ${g}, ${b}, ${a})`;
  return `rgb(${r}, ${g}, ${b})`;
}

function borderToCss(border) {
  if (!border || !border.style || border.style === 'NONE') return null;

  const color = rgbToCss(border.colorStyle?.rgbColor) || rgbToCss(border.color) || 'rgb(0, 0, 0)';
  let width = 1;
  if (border.style === 'SOLID_MEDIUM') width = 2;
  if (border.style === 'SOLID_THICK') width = 3;

  let line = 'solid';
  if (border.style === 'DOTTED') line = 'dotted';
  if (border.style === 'DASHED') line = 'dashed';
  if (border.style === 'DOUBLE') line = 'double';

  return `${width}px ${line} ${color}`;
}

function columnToLetters(column) {
  let num = column;
  let letters = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    num = Math.floor((num - rem - 1) / 26);
  }
  return letters;
}

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_PATH,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: 'v4', auth: authClient });
}

async function ensureSuggestionsSheet(sheets) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
    fields: 'sheets(properties(title))',
  });
  const existingTitles = (meta.data.sheets || [])
    .map((s) => s.properties?.title)
    .filter(Boolean);

  const exists = existingTitles.some(
    (title) => title.trim().toLowerCase() === SUGGESTIONS_TAB_NAME.toLowerCase()
  );

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: SUGGESTIONS_TAB_NAME } } }],
      },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SUGGESTIONS_TAB_NAME}!A1:H1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [['Timestamp', 'Type', 'Name', 'Email', 'Shout Out For', 'Location', 'Message', 'Reactions']],
      },
    });
  }
}

async function ensureSupplyOrdersSheet(sheets) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
    fields: 'sheets(properties(sheetId,title))',
  });

  const existingSheet = (meta.data.sheets || []).find(
    (s) =>
      String(s.properties?.title || '').trim().toLowerCase() ===
      SUPPLY_ORDERS_TAB_NAME.toLowerCase()
  );

  const headers = [
    'Submitted At',
    'Store',
    'Standard Items',
    ...Array.from({ length: MAX_OTHER_REQUESTS }, (_, idx) => `Other Request ${idx + 1}`),
  ];
  const lastCol = columnToLetters(headers.length);

  if (existingSheet?.properties?.sheetId) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SUPPLY_ORDERS_TAB_NAME}!A1:${lastCol}1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [headers],
      },
    });
    return existingSheet.properties.sheetId;
  }

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title: SUPPLY_ORDERS_TAB_NAME } } }],
    },
  });

  const newSheetId = addRes.data?.replies?.[0]?.addSheet?.properties?.sheetId;
  if (!Number.isInteger(newSheetId)) {
    throw new Error('Failed to create Supply Orders sheet.');
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SUPPLY_ORDERS_TAB_NAME}!A1:${lastCol}1`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [headers],
    },
  });

  return newSheetId;
}

function parseSheetTimestamp(rawValue) {
  const text = String(rawValue || '').trim();
  if (!text) return null;

  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) return direct;

  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!match) return null;

  const month = Number(match[1]);
  const day = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += 2000;

  const parsed = new Date(year, month - 1, day);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed;
}

function parseReactionMap(rawValue) {
  const text = String(rawValue ?? '').trim();
  if (!text) return {};

  const asNumber = Number.parseInt(text, 10);
  if (!Number.isNaN(asNumber) && String(asNumber) === text) {
    return { clap: Math.max(asNumber, 0) };
  }

  const map = {};
  text.split('|').forEach((part) => {
    const [keyRaw, valueRaw] = part.split(':');
    const key = String(keyRaw || '').trim();
    const value = Number.parseInt(String(valueRaw || '').trim(), 10);
    if (!key || Number.isNaN(value) || value <= 0) return;
    map[key] = value;
  });
  return map;
}

function encodeReactionMap(map) {
  const keys = Object.keys(map || {});
  if (keys.length === 0) return '0';
  return keys
    .filter((key) => Number.isFinite(map[key]) && map[key] > 0)
    .map((key) => `${key}:${Math.floor(map[key])}`)
    .join('|') || '0';
}

// ── IPC: Get next available code from a tab ───────────────────────────────────
ipcMain.handle('get-next-code', async (event, tabName) => {
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${tabName}!A:A`,
    });

    const rows = res.data.values || [];
    if (rows.length === 0) return { success: false, message: 'No codes available.' };

    // First available code
    const code = rows[0][0];
    if (!code) return { success: false, message: 'No codes available.' };

    return { success: true, code, tabName };
  } catch (err) {
    console.error(err);
    return { success: false, message: 'Failed to fetch code: ' + err.message };
  }
});

// ── IPC: Mark code as used (move to Used tab) ─────────────────────────────────
ipcMain.handle('mark-code-used', async (event, { code, tabName, employeeName, customerName, notes }) => {
  try {
    const sheets = await getSheetsClient();

    // 1. Find the row in the source tab
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${tabName}!A:A`,
    });
    const rows = res.data.values || [];
    const rowIndex = rows.findIndex(r => r[0] === code);
    if (rowIndex === -1) return { success: false, message: 'Code not found.' };

    // 2. Append to Used tab
    const now = new Date();
    const usedDate = now.toLocaleDateString();
    const usedTime = now.toLocaleTimeString();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: 'Used!A:G',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          code,
          employeeName || 'Unknown',
          customerName || 'Unknown',
          usedDate,
          usedTime,
          tabName,
          notes || '',
        ]],
      },
    });

    // 3. Delete from source tab
    // Get sheet metadata to find sheetId
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheet = meta.data.sheets.find(s => s.properties.title === tabName);
    if (!sheet) return { success: false, message: 'Tab not found.' };

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheet.properties.sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex,
              endIndex: rowIndex + 1,
            },
          },
        }],
      },
    });

    return { success: true };
  } catch (err) {
    console.error(err);
    return { success: false, message: 'Failed to mark used: ' + err.message };
  }
});

// ── IPC: Get available tabs ───────────────────────────────────────────────────
ipcMain.handle('get-tabs', async () => {
  try {
    const sheets = await getSheetsClient();
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const tabs = meta.data.sheets
      .map(s => s.properties.title)
      .filter(
        (t) =>
          t !== 'Used' &&
          t !== SUGGESTIONS_TAB_NAME &&
          t !== SUPPLY_ORDERS_TAB_NAME
      );
    return { success: true, tabs };
  } catch (err) {
    return { success: false, message: err.message };
  }
});

ipcMain.handle('submit-suggestion', async (event, payload) => {
  try {
    const suggestion = String(payload?.suggestion || '').trim();
    const name = String(payload?.name || '').trim();
    const email = String(payload?.email || '').trim();
    const type = String(payload?.type || 'Suggestion').trim();
    const shoutOutFor = String(payload?.shoutOutFor || '').trim();
    const location = String(payload?.location || '').trim();

    if (suggestion.length < 4) {
      return { success: false, message: 'Suggestion must be at least 4 characters.' };
    }
    if (type.toLowerCase() === 'shout out' && shoutOutFor.length < 2) {
      return { success: false, message: 'Shout out recipient is required.' };
    }

    const sheets = await getSheetsClient();
    await ensureSuggestionsSheet(sheets);

    const now = new Date();
    const timestamp = `${now.toLocaleDateString()} ${now.toLocaleTimeString()}`;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${SUGGESTIONS_TAB_NAME}!A:H`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[timestamp, type || 'Suggestion', name || 'Anonymous', email || '', shoutOutFor, location, suggestion, 0]],
      },
    });

    return { success: true };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to submit suggestion.' };
  }
});

ipcMain.handle('submit-supply-order', async (event, payload) => {
  try {
    const store = String(payload?.store || '').trim();
    const quantitiesRaw = payload?.quantities && typeof payload.quantities === 'object'
      ? payload.quantities
      : {};

    if (!SUPPLY_ORDER_LOCATIONS.includes(store)) {
      return { success: false, message: 'Please select a valid store location.' };
    }

    const quantities = {};
    let totalQty = 0;

    for (const item of SUPPLY_ORDER_ITEMS) {
      const raw = quantitiesRaw[item];
      const qtyNum = Number.parseInt(String(raw ?? ''), 10);
      const qty = Number.isNaN(qtyNum) ? 0 : Math.max(0, qtyNum);
      quantities[item] = qty;
      totalQty += qty;
    }

    const otherRaw = Array.isArray(payload?.otherRequests) ? payload.otherRequests : [];
    const otherRequests = [];
    otherRaw.forEach((entry) => {
      if (otherRequests.length >= MAX_OTHER_REQUESTS) return;
      const description = String(entry?.description || '').trim();
      const qtyRaw = Number.parseInt(String(entry?.qty ?? ''), 10);
      const qty = Number.isNaN(qtyRaw) ? 0 : Math.max(0, qtyRaw);
      if (!description || qty <= 0) return;
      otherRequests.push(qty > 1 ? `${description} (x${qty})` : description);
    });

    if (totalQty <= 0 && otherRequests.length === 0) {
      return { success: false, message: 'Add at least one item before submitting.' };
    }

    const sheets = await getSheetsClient();
    const sheetId = await ensureSupplyOrdersSheet(sheets);

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId,
                dimension: 'ROWS',
                startIndex: 1,
                endIndex: 2,
              },
              inheritFromBefore: false,
            },
          },
        ],
      },
    });

    const now = new Date();
    const submittedAt = `${now.toLocaleDateString()} ${now.toLocaleTimeString()}`;
    const standardItemsSummary = SUPPLY_ORDER_ITEMS
      .filter((item) => quantities[item] > 0)
      .map((item) => `${item} - ${quantities[item]}`)
      .join(', ');

    const rowValues = [
      submittedAt,
      store,
      standardItemsSummary,
      ...Array.from({ length: MAX_OTHER_REQUESTS }, (_, idx) => otherRequests[idx] || ''),
    ];
    const lastCol = columnToLetters(rowValues.length);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SUPPLY_ORDERS_TAB_NAME}!A2:${lastCol}2`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [rowValues],
      },
    });

    return { success: true };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to submit order.' };
  }
});

ipcMain.handle('get-recent-shout-outs', async () => {
  try {
    const sheets = await getSheetsClient();
    await ensureSuggestionsSheet(sheets);

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SUGGESTIONS_TAB_NAME}!A:H`,
    });

    const rows = res.data.values || [];
    if (rows.length <= 1) return { success: true, shoutOuts: [] };

    const now = new Date();
    const cutoff = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const shoutOuts = [];

    for (let i = 1; i < rows.length; i += 1) {
      const row = rows[i];
      const timestamp = row[0] || '';
      const type = String(row[1] || '').trim().toLowerCase();
      const fromName = row[2] || 'Anonymous';
      const shoutOutFor = row[4] || '';
      const hasLocationColumn = row.length >= 7;
      const location = hasLocationColumn ? (row[5] || '') : '';
      const message = hasLocationColumn ? (row[6] || '') : (row[5] || '');
      const reactionsRaw = row[7];
      const reactions = parseReactionMap(reactionsRaw);

      if (type !== 'shout out') continue;
      if (!message) continue;

      const parsedDate = parseSheetTimestamp(timestamp);
      if (!parsedDate || parsedDate < cutoff) continue;

      shoutOuts.push({
        rowNumber: i + 1,
        timestamp,
        fromName,
        shoutOutFor,
        location,
        message,
        reactions,
      });
    }

    shoutOuts.sort((a, b) => {
      const aDate = parseSheetTimestamp(a.timestamp)?.getTime() || 0;
      const bDate = parseSheetTimestamp(b.timestamp)?.getTime() || 0;
      return bDate - aDate;
    });

    return { success: true, shoutOuts };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to load shout outs.', shoutOuts: [] };
  }
});

ipcMain.handle('react-to-shout-out', async (event, payload) => {
  try {
    const rowNumber = Number(payload?.rowNumber);
    const reactionTypeRaw = String(payload?.reactionType || '').trim().toLowerCase();
    const allowedTypes = new Set(['like', 'love', 'laugh', 'clap']);
    if (!Number.isInteger(rowNumber) || rowNumber < 2) {
      return { success: false, message: 'Invalid shout out row.' };
    }
    if (!allowedTypes.has(reactionTypeRaw)) {
      return { success: false, message: 'Invalid reaction type.' };
    }

    const sheets = await getSheetsClient();
    await ensureSuggestionsSheet(sheets);

    const currentRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SUGGESTIONS_TAB_NAME}!H${rowNumber}:H${rowNumber}`,
    });

    const currentRaw = currentRes.data.values?.[0]?.[0] ?? '0';
    const reactionMap = parseReactionMap(currentRaw);
    reactionMap[reactionTypeRaw] = (Number.isFinite(reactionMap[reactionTypeRaw]) ? reactionMap[reactionTypeRaw] : 0) + 1;
    const encoded = encodeReactionMap(reactionMap);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SUGGESTIONS_TAB_NAME}!H${rowNumber}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[encoded]],
      },
    });

    return { success: true, reactions: reactionMap };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to add reaction.' };
  }
});

// IPC: Get Stack Ranker table data (view-only)
ipcMain.handle('get-stack-ranker', async () => {
  try {
    const sheets = await getSheetsClient();
    const metaRes = await sheets.spreadsheets.get({
      spreadsheetId: STACK_RANKER_SHEET_ID,
      fields: 'sheets(properties(title))',
    });
    const sheetTitles = (metaRes.data.sheets || [])
      .map((s) => s.properties?.title)
      .filter(Boolean);
    if (sheetTitles.length === 0) {
      return { success: false, message: 'Stack Ranker spreadsheet has no sheets/tabs.' };
    }

    const preferred =
      sheetTitles.find((t) => t === STACK_RANKER_TAB_NAME) ||
      sheetTitles.find((t) => t.toLowerCase().trim() === STACK_RANKER_TAB_NAME.toLowerCase()) ||
      sheetTitles[0];

    const valuesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: STACK_RANKER_SHEET_ID,
      range: `${preferred}!A:ZZ`,
    });
    const formulaRes = await sheets.spreadsheets.values.get({
      spreadsheetId: STACK_RANKER_SHEET_ID,
      range: `${preferred}!A:ZZ`,
      valueRenderOption: 'FORMULA',
    });

    const values = valuesRes.data.values || [];
    const formulaRows = formulaRes.data.values || [];
    const imageUrlSet = new Set();

    // Pull image URLs from IMAGE("...") formulas and direct image links in cells.
    formulaRows.forEach((row) => {
      row.forEach((cell) => {
        if (!cell) return;
        const text = String(cell).trim();

        const imageFormulaMatch = text.match(/^=IMAGE\(\s*"([^"]+)"/i);
        if (imageFormulaMatch?.[1]) {
          imageUrlSet.add(imageFormulaMatch[1]);
          return;
        }

        const imageUrlMatch = text.match(/^https?:\/\/\S+/i);
        if (
          imageUrlMatch &&
          (/\.(png|jpe?g|gif|webp|bmp|svg)(\?|$)/i.test(text) ||
            /drive\.google\.com/i.test(text) ||
            /googleusercontent\.com/i.test(text))
        ) {
          imageUrlSet.add(text);
        }
      });
    });

    const images = Array.from(imageUrlSet);
    if (values.length === 0) return { success: true, values: [], cells: [], images, sheetName: preferred };

    const rowCount = values.length;
    const colCount = Math.max(...values.map((r) => r.length), 1);
    const lastCol = columnToLetters(colCount);
    const gridRange = `${preferred}!A1:${lastCol}${rowCount}`;

    const gridRes = await sheets.spreadsheets.get({
      spreadsheetId: STACK_RANKER_SHEET_ID,
      ranges: [gridRange],
      includeGridData: true,
      fields: 'sheets(data(rowData(values(formattedValue,effectiveValue,effectiveFormat))))',
    });

    const rowData = gridRes.data.sheets?.[0]?.data?.[0]?.rowData || [];
    const cells = rowData.map((row) => {
      const rowValues = row.values || [];
      const normalized = [];

      for (let i = 0; i < colCount; i += 1) {
        const cell = rowValues[i] || {};
        const format = cell.effectiveFormat || {};
        const textFormat = format.textFormat || {};
        const borders = format.borders || {};

        const bgColor =
          rgbToCss(format.backgroundColorStyle?.rgbColor) || rgbToCss(format.backgroundColor);
        const textColor =
          rgbToCss(textFormat.foregroundColorStyle?.rgbColor) || rgbToCss(textFormat.foregroundColor);

        normalized.push({
          text: cell.formattedValue || '',
          style: {
            backgroundColor: bgColor,
            color: textColor,
            fontSize: textFormat.fontSize ? `${textFormat.fontSize}px` : null,
            fontWeight: textFormat.bold ? '700' : null,
            fontStyle: textFormat.italic ? 'italic' : null,
            textDecoration: textFormat.underline ? 'underline' : null,
            textAlign: format.horizontalAlignment ? format.horizontalAlignment.toLowerCase() : null,
            verticalAlign: format.verticalAlignment ? format.verticalAlignment.toLowerCase() : null,
            borderTop: borderToCss(borders.top),
            borderRight: borderToCss(borders.right),
            borderBottom: borderToCss(borders.bottom),
            borderLeft: borderToCss(borders.left),
          },
        });
      }

      return normalized;
    });

    return { success: true, values, cells, images, sheetName: preferred };
  } catch (err) {
    const baseMsg = err?.message || 'Unknown error';
    if (baseMsg.includes('Requested entity was not found')) {
      return {
        success: false,
        message:
          'Stack Ranker sheet not found. Verify spreadsheet ID, confirm it is shared with the service account, and confirm the tab exists.',
      };
    }
    return { success: false, message: baseMsg };
  }
});

// ── IPC: Window controls ──────────────────────────────────────────────────────
ipcMain.on('close-app', () => { if (mainWindow) mainWindow.close(); });
ipcMain.on('minimize-app', () => { if (mainWindow) mainWindow.minimize(); });
ipcMain.on('toggle-maximize', () => {
  if (!mainWindow) return;
  if (mainWindow.isMaximized()) mainWindow.unmaximize();
  else mainWindow.maximize();
});

// ── Window ────────────────────────────────────────────────────────────────────
let mainWindow;

function setupAutoUpdater() {
  if (!app.isPackaged) return;

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  let updatePromptShown = false;

  autoUpdater.on('error', (err) => {
    console.error('Auto-update error:', err?.message || err);
  });

  autoUpdater.on('checking-for-update', () => {
    console.log('Checking for update...');
  });

  autoUpdater.on('update-available', (info) => {
    console.log(`Update available: ${info.version}`);
  });

  autoUpdater.on('update-not-available', (info) => {
    console.log(`No update available. Current version: ${info.version || app.getVersion()}`);
  });

  autoUpdater.on('update-downloaded', async (info) => {
    console.log(`Update downloaded: ${info.version}. Will install on app quit.`);

    if (updatePromptShown) return;
    updatePromptShown = true;

    try {
      const result = await dialog.showMessageBox(mainWindow || null, {
        type: 'info',
        buttons: ['Restart Now', 'Later'],
        defaultId: 0,
        cancelId: 1,
        title: 'Update Ready',
        message: `Version ${info.version} is ready to install.`,
        detail: 'Restart now to apply the update, or choose Later to install when the app closes.',
      });

      if (result.response === 0) {
        autoUpdater.quitAndInstall();
      }
    } catch (err) {
      console.error('Failed to show update prompt:', err?.message || err);
    }
  });

  // Small delay so window/UI initialization is not blocked.
  setTimeout(() => {
    autoUpdater.checkForUpdates().catch((err) => {
      console.error('Failed to check for updates:', err?.message || err);
    });
  }, 4000);
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 720,
    minWidth: 900,
    minHeight: 600,
    frame: false,
    titleBarStyle: 'hidden',
    backgroundColor: '#0a0a0f',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    icon: path.join(__dirname, 'assets', 'icon.ico'),
  });

  mainWindow.loadFile('src/index.html');
}

app.whenReady().then(createWindow);
app.whenReady().then(setupAutoUpdater);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
