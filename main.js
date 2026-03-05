const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');
const { autoUpdater } = require('electron-updater');

// ── Google Sheets config ──────────────────────────────────────────────────────
const SHEET_ID = '1nywGhJ50rrntwoGwO9zgPaVOsEpHM1iTOG7O_pXHKrQ';
const STACK_RANKER_SHEET_ID = '1oKJbQA12JIvNQdjNiWmd3jb2pmQytuxbRh_snGWvbRk';
const STACK_RANKER_TAB_NAME = 'Stack Ranker';
const SALES_REPORT_TAB_NAME = 'S/P Month to Date';
const SALES_REPORT_STORE_COL_INDEX = 1; // Column B
const SALES_REPORT_PRODUCT_COL_INDEX = 4; // Column E
const SALES_REPORT_QTY_COL_INDEX = 7; // Column H
const SALES_REPORT_SELL_PRICE_COL_INDEX = 10; // Column K
const SALES_REPORT_NET_PROFIT_COL_INDEX = 13; // Column N
const SALES_REPORT_EMPLOYEE_COL_INDEX = 2; // Column C
const MAX_EMPLOYEE_HIGHLIGHTS = 6;
const MAX_COMPANY_LEADERS = 8;
const GLASS_ATTACHMENT_PRIOR_WEIGHT = 5;
const COMPANY_TOP_PRODUCTS_FOR_EMPLOYEE_HIGHLIGHTS = 20;
const EMPLOYEE_HIGHLIGHT_MAX_TOP_PERCENT = 50;
const OVERRIDE_CODE_TAB_ALLOWLIST = ['Apple Pay', 'Refund', 'Discount'];
const SUGGESTIONS_TAB_NAME = 'Suggestions';
const SUPPLY_ORDERS_TAB_NAME = 'Supply Orders';
const DEFAULT_CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
const employeeHighlightsCache = {
  key: '',
  expiresAt: 0,
  data: null,
};

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
  const envCredentialsPath = String(process.env.TOA_GOOGLE_CREDENTIALS || '').trim();
  const credentialsPath = envCredentialsPath || DEFAULT_CREDENTIALS_PATH;
  if (!fs.existsSync(credentialsPath)) {
    throw new Error(
      `Google credentials file not found at "${credentialsPath}". ` +
      'Place credentials.json in the app root or set TOA_GOOGLE_CREDENTIALS to the full file path.'
    );
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: credentialsPath,
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

function escapeSheetTitleForA1(title) {
  return `'${String(title || '').replace(/'/g, "''")}'`;
}

function buildSheetRange(title, a1Range) {
  return `${escapeSheetTitleForA1(title)}!${a1Range}`;
}

function normalizeName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function parseQuantity(raw) {
  const cleaned = String(raw ?? '')
    .replace(/,/g, '')
    .trim();
  const qty = Number.parseFloat(cleaned);
  if (!Number.isFinite(qty)) return 0;
  return qty > 0 ? qty : 0;
}

function parseCurrency(raw) {
  const cleaned = String(raw ?? '')
    .replace(/[$,]/g, '')
    .trim();
  const value = Number.parseFloat(cleaned);
  if (!Number.isFinite(value)) return 0;
  return value;
}

function getEmployeeHighlightBucket(productName) {
  const text = String(productName || '').toLowerCase();
  if (text.includes('glass install')) return 'Glass Install';
  if (text.includes('lens cover')) return 'Lens Cover';
  return String(productName || '').trim();
}

function isRepairProductName(productName) {
  const text = String(productName || '').toLowerCase();
  return (
    text.includes('display')
    || text.includes('back glass')
    || text.includes('frame')
    || text.includes('battery replacement')
    || text.includes('repair')
  );
}

function resolveEmployeePrimaryStore(aggregate, employeeKey) {
  const byStore = aggregate.employeeStoreCounts.get(employeeKey);
  if (!byStore || byStore.size === 0) return '';

  let winningStoreKey = '';
  let winningCount = -1;
  byStore.forEach((count, storeKey) => {
    if (count > winningCount) {
      winningStoreKey = storeKey;
      winningCount = count;
    }
  });

  return aggregate.storeDisplayNames.get(winningStoreKey) || '';
}

function getWeightedGlassAttachmentRatePercent(numerator, denominator, companyRatePercent) {
  const num = Number(numerator) || 0;
  const den = Number(denominator) || 0;
  if (den <= 0) return null;
  const priorRate = Math.max(0, Number(companyRatePercent) || 0) / 100;
  const adjusted = (num + (priorRate * GLASS_ATTACHMENT_PRIOR_WEIGHT)) / (den + GLASS_ATTACHMENT_PRIOR_WEIGHT);
  return adjusted * 100;
}

function buildEmployeeHighlightItems(aggregate, employeeKey) {
  const topCompanyProducts = Array.from(aggregate.companyTotalsByProduct.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, COMPANY_TOP_PRODUCTS_FOR_EMPLOYEE_HIGHLIGHTS)
    .map(([product]) => product);
  const allowedProducts = new Set(topCompanyProducts);

  const groupedCompanyTotals = new Map();
  topCompanyProducts.forEach((product) => {
    const bucket = getEmployeeHighlightBucket(product);
    const qty = aggregate.companyTotalsByProduct.get(product) || 0;
    groupedCompanyTotals.set(bucket, (groupedCompanyTotals.get(bucket) || 0) + qty);
  });

  const groupedEmployeeTotals = new Map();
  aggregate.employeeProductTotals.forEach((productMap, key) => {
    const bucketMap = new Map();
    productMap.forEach((qty, product) => {
      if (!allowedProducts.has(product)) return;
      const bucket = getEmployeeHighlightBucket(product);
      bucketMap.set(bucket, (bucketMap.get(bucket) || 0) + qty);
    });
    if (bucketMap.size > 0) groupedEmployeeTotals.set(key, bucketMap);
  });

  const selectedGroupedTotals = groupedEmployeeTotals.get(employeeKey);
  if (!selectedGroupedTotals || selectedGroupedTotals.size === 0) {
    return [];
  }

  const items = [];
  for (const [bucket, qty] of selectedGroupedTotals.entries()) {
    const companyTotal = groupedCompanyTotals.get(bucket) || 0;
    const sellers = [];
    groupedEmployeeTotals.forEach((bucketMap) => {
      const employeeQty = bucketMap.get(bucket) || 0;
      if (employeeQty > 0) sellers.push(employeeQty);
    });
    if (sellers.length === 0) continue;

    const higherCount = sellers.filter((amount) => amount > qty).length;
    const rank = higherCount + 1;
    const sellerCount = sellers.length;
    const topPercent = sellerCount <= 1
      ? 1
      : Math.round((higherCount / (sellerCount - 1)) * 99) + 1;
    const sharePercent = companyTotal > 0 ? (qty / companyTotal) * 100 : 0;
    if (topPercent > EMPLOYEE_HIGHLIGHT_MAX_TOP_PERCENT) continue;

    items.push({
      product: bucket,
      qty: Math.round(qty),
      companyTotal: Math.round(companyTotal),
      rank,
      sellerCount,
      topPercent,
      sharePercent: Math.round(sharePercent * 10) / 10,
      isLeader: higherCount === 0,
    });
  }

  items.sort((a, b) => {
    if (a.isLeader !== b.isLeader) return a.isLeader ? -1 : 1;
    if (b.qty !== a.qty) return b.qty - a.qty;
    if (a.topPercent !== b.topPercent) return a.topPercent - b.topPercent;
    return b.sharePercent - a.sharePercent;
  });

  const glassStats = aggregate.employeeGlassAttachmentStats.get(employeeKey) || {
    numerator: 0,
    denominator: 0,
    ratePercent: null,
    glassInstallAddonInvoices: 0,
  };
  let glassAttachmentItem = null;
  if (glassStats.denominator > 0 && Number.isFinite(glassStats.ratePercent)) {
    const glassCohort = Array.from(aggregate.employeeGlassAttachmentStats.values())
      .filter((row) => row.denominator > 0 && Number.isFinite(row.ratePercent));
    const cohortTotals = glassCohort.reduce((acc, row) => {
      acc.numerator += Number(row.numerator) || 0;
      acc.denominator += Number(row.denominator) || 0;
      return acc;
    }, { numerator: 0, denominator: 0 });
    const companyRatePercent = cohortTotals.denominator > 0
      ? (cohortTotals.numerator / cohortTotals.denominator) * 100
      : 0;
    const selectedWeightedScore = getWeightedGlassAttachmentRatePercent(
      glassStats.numerator,
      glassStats.denominator,
      companyRatePercent
    );
    const weightedScores = glassCohort
      .map((row) => getWeightedGlassAttachmentRatePercent(row.numerator, row.denominator, companyRatePercent))
      .filter((score) => Number.isFinite(score));
    const higherCount = weightedScores.filter((score) => score > selectedWeightedScore).length;
    const rank = higherCount + 1;
    const sellerCount = weightedScores.length;
    const topPercent = sellerCount <= 1
      ? 1
      : Math.round((higherCount / (sellerCount - 1)) * 99) + 1;
    glassAttachmentItem = {
      product: 'Glass Attachment Rate',
      metricType: 'glassAttachmentRate',
      qty: Math.round(glassStats.ratePercent * 100) / 100,
      rank,
      sellerCount,
      topPercent,
      sharePercent: 0,
      isLeader: higherCount === 0,
      numerator: glassStats.numerator,
      denominator: glassStats.denominator,
      weightedScore: Number(selectedWeightedScore) || 0,
    };
  }

  const productHighlightLimit = glassAttachmentItem
    ? Math.max(MAX_EMPLOYEE_HIGHLIGHTS - 1, 0)
    : MAX_EMPLOYEE_HIGHLIGHTS;
  const topItems = items.slice(0, productHighlightLimit);
  if (glassAttachmentItem) topItems.push(glassAttachmentItem);
  return topItems;
}

async function getSalesAggregateData(sheets, tabTitle) {
  const cacheKey = `${STACK_RANKER_SHEET_ID}:${tabTitle}`;
  if (employeeHighlightsCache.key === cacheKey && Date.now() < employeeHighlightsCache.expiresAt) {
    return employeeHighlightsCache.data;
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: STACK_RANKER_SHEET_ID,
    range: buildSheetRange(tabTitle, 'A:N'),
  });
  const rows = res.data.values || [];

  const companyTotalsByProduct = new Map();
  const employeeProductTotals = new Map();
  const overallByEmployee = new Map();
  const employeeInvoiceIds = new Map();
  const employeeDisplayNames = new Map();
  const overallByStore = new Map();
  const netProfitByStore = new Map();
  const storeProductTotals = new Map();
  const storeEmployees = new Map();
  const storeInvoiceIds = new Map();
  const storeInvoiceMetrics = new Map();
  const storeDisplayNames = new Map();
  const employeeStoreCounts = new Map();
  const employeeInvoiceMetrics = new Map();
  let companyTotalQty = 0;

  rows.forEach((row) => {
    const invoiceId = String(row?.[0] || '').trim();
    const storeName = String(row?.[SALES_REPORT_STORE_COL_INDEX] || '').trim();
    const storeKey = normalizeName(storeName);
    const employeeName = String(row?.[SALES_REPORT_EMPLOYEE_COL_INDEX] || '').trim();
    const employeeKey = normalizeName(employeeName);
    const product = String(row?.[SALES_REPORT_PRODUCT_COL_INDEX] || '').trim();
    const qty = parseQuantity(row?.[SALES_REPORT_QTY_COL_INDEX]);
    const sellingPrice = parseCurrency(row?.[SALES_REPORT_SELL_PRICE_COL_INDEX]);
    const netProfit = parseCurrency(row?.[SALES_REPORT_NET_PROFIT_COL_INDEX]);

    if (!employeeKey || !product || qty <= 0) return;
    if (!employeeDisplayNames.has(employeeKey)) employeeDisplayNames.set(employeeKey, employeeName);

    companyTotalQty += qty;
    companyTotalsByProduct.set(product, (companyTotalsByProduct.get(product) || 0) + qty);
    overallByEmployee.set(employeeKey, (overallByEmployee.get(employeeKey) || 0) + qty);
    if (!employeeInvoiceIds.has(employeeKey)) employeeInvoiceIds.set(employeeKey, new Set());
    if (invoiceId) employeeInvoiceIds.get(employeeKey).add(invoiceId);

    if (!employeeProductTotals.has(employeeKey)) employeeProductTotals.set(employeeKey, new Map());
    const productMap = employeeProductTotals.get(employeeKey);
    productMap.set(product, (productMap.get(product) || 0) + qty);

    if (storeKey) {
      if (!employeeStoreCounts.has(employeeKey)) employeeStoreCounts.set(employeeKey, new Map());
      const storeCountMap = employeeStoreCounts.get(employeeKey);
      storeCountMap.set(storeKey, (storeCountMap.get(storeKey) || 0) + 1);
    }

    if (invoiceId) {
      if (!employeeInvoiceMetrics.has(employeeKey)) employeeInvoiceMetrics.set(employeeKey, new Map());
      const invoiceMap = employeeInvoiceMetrics.get(employeeKey);
      if (!invoiceMap.has(invoiceId)) {
        invoiceMap.set(invoiceId, {
          products: new Set(),
          hasAnyGlassInstall: false,
          hasZeroDollarGlassInstall: false,
          hasGp5to15GlassInstall: false,
        });
      }
      const invoiceMetrics = invoiceMap.get(invoiceId);
      invoiceMetrics.products.add(product);
      const productText = product.toLowerCase();
      if (productText.includes('glass install')) {
        invoiceMetrics.hasAnyGlassInstall = true;
        if (Math.abs(sellingPrice) < 0.000001) invoiceMetrics.hasZeroDollarGlassInstall = true;
        if (netProfit >= 5 && netProfit <= 15) invoiceMetrics.hasGp5to15GlassInstall = true;
      }
    }

    if (storeKey) {
      if (!storeDisplayNames.has(storeKey)) storeDisplayNames.set(storeKey, storeName);
      overallByStore.set(storeKey, (overallByStore.get(storeKey) || 0) + qty);
      netProfitByStore.set(storeKey, (netProfitByStore.get(storeKey) || 0) + netProfit);

      if (!storeProductTotals.has(storeKey)) storeProductTotals.set(storeKey, new Map());
      const storeProductMap = storeProductTotals.get(storeKey);
      storeProductMap.set(product, (storeProductMap.get(product) || 0) + qty);

      if (!storeEmployees.has(storeKey)) storeEmployees.set(storeKey, new Set());
      storeEmployees.get(storeKey).add(employeeKey);

      if (!storeInvoiceIds.has(storeKey)) storeInvoiceIds.set(storeKey, new Set());
      if (invoiceId) storeInvoiceIds.get(storeKey).add(invoiceId);

      if (invoiceId) {
        if (!storeInvoiceMetrics.has(storeKey)) storeInvoiceMetrics.set(storeKey, new Map());
        const invoiceMap = storeInvoiceMetrics.get(storeKey);
        if (!invoiceMap.has(invoiceId)) {
          invoiceMap.set(invoiceId, {
            netProfitTotal: 0,
            hasRepairItem: false,
          });
        }
        const invoiceMetrics = invoiceMap.get(invoiceId);
        invoiceMetrics.netProfitTotal += netProfit;
        if (isRepairProductName(product)) invoiceMetrics.hasRepairItem = true;
      }
    }
  });

  const employeeGlassAttachmentStats = new Map();
  employeeInvoiceMetrics.forEach((invoiceMap, employeeKey) => {
    let numerator = 0;
    let denominator = 0;
    let glassInstallAddonInvoices = 0;
    invoiceMap.forEach((invoiceMetrics) => {
      const isAddon = invoiceMetrics.products.size >= 2;
      if (!isAddon) return;
      if (invoiceMetrics.hasAnyGlassInstall) glassInstallAddonInvoices += 1;
      if (invoiceMetrics.hasZeroDollarGlassInstall) numerator += 1;
      if (invoiceMetrics.hasAnyGlassInstall) denominator += 1;
    });
    const ratePercent = denominator > 0 ? (numerator / denominator) * 100 : null;
    employeeGlassAttachmentStats.set(employeeKey, {
      numerator,
      denominator,
      ratePercent,
      glassInstallAddonInvoices,
    });
  });

  const aggregate = {
    companyTotalsByProduct,
    employeeProductTotals,
    overallByEmployee,
    employeeInvoiceIds,
    employeeDisplayNames,
    overallByStore,
    netProfitByStore,
    storeProductTotals,
    storeEmployees,
    storeInvoiceIds,
    storeInvoiceMetrics,
    storeDisplayNames,
    employeeStoreCounts,
    employeeGlassAttachmentStats,
    companyTotalQty,
  };

  employeeHighlightsCache.key = cacheKey;
  employeeHighlightsCache.expiresAt = Date.now() + 60 * 1000;
  employeeHighlightsCache.data = aggregate;
  return aggregate;
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
    const tabsByNormalized = new Map(
      (meta.data.sheets || [])
        .map((s) => String(s.properties?.title || '').trim())
        .filter(Boolean)
        .map((title) => [title.toLowerCase(), title])
    );

    const tabs = OVERRIDE_CODE_TAB_ALLOWLIST
      .map((allowedTitle) => tabsByNormalized.get(allowedTitle.toLowerCase()) || null)
      .filter(Boolean);

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

ipcMain.handle('get-store-employee-directory', async () => {
  try {
    const sheets = await getSheetsClient();

    const rangeRes = await sheets.spreadsheets.values.get({
      spreadsheetId: STACK_RANKER_SHEET_ID,
      range: buildSheetRange(SALES_REPORT_TAB_NAME, 'B:C'),
    });

    const rows = rangeRes.data.values || [];
    const byStore = new Map();

    rows.forEach((row) => {
      const store = String(row?.[0] || '').trim();
      const employee = String(row?.[1] || '').trim();
      if (!store || !employee) return;

      if (!byStore.has(store)) byStore.set(store, new Set());
      byStore.get(store).add(employee);
    });

    const stores = Array.from(byStore.keys()).sort((a, b) => a.localeCompare(b));
    const directory = stores.map((store) => ({
      store,
      employees: Array.from(byStore.get(store)).sort((a, b) => a.localeCompare(b)),
    }));

    return { success: true, directory, sourceTab: SALES_REPORT_TAB_NAME };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to load store/employee directory.' };
  }
});

ipcMain.handle('get-employee-highlights', async (_event, payload) => {
  try {
    const employeeName = String(payload?.employeeName || '').trim();
    const employeeKey = normalizeName(employeeName);
    if (!employeeKey) {
      return { success: false, message: 'Employee name is required.' };
    }

    const sheets = await getSheetsClient();
    const tabTitle = SALES_REPORT_TAB_NAME;

    const aggregate = await getSalesAggregateData(sheets, tabTitle);
    const saleCount = aggregate.employeeInvoiceIds.get(employeeKey)?.size || 0;
    const topItems = buildEmployeeHighlightItems(aggregate, employeeKey);
    if (topItems.length === 0) {
      return {
        success: true,
        sourceTab: tabTitle,
        employeeName,
        items: [],
        summary: {
          saleCount,
          leaderCount: 0,
          companyEmployeeCount: aggregate.employeeProductTotals.size,
        },
      };
    }

    const leaderCount = topItems.filter((row) => row.isLeader).length;

    return {
      success: true,
      sourceTab: tabTitle,
      employeeName,
      items: topItems,
      summary: {
        saleCount,
        leaderCount,
        companyEmployeeCount: aggregate.employeeProductTotals.size,
      },
    };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to calculate employee highlights.' };
  }
});

ipcMain.handle('get-company-leaders', async () => {
  try {
    const sheets = await getSheetsClient();
    const tabTitle = SALES_REPORT_TAB_NAME;
    const aggregate = await getSalesAggregateData(sheets, tabTitle);

    const leaders = [];
    aggregate.employeeProductTotals.forEach((_value, employeeKey) => {
      const items = buildEmployeeHighlightItems(aggregate, employeeKey);
      if (items.length === 0) return;

      const employeeName = aggregate.employeeDisplayNames?.get(employeeKey) || employeeKey;
      const location = resolveEmployeePrimaryStore(aggregate, employeeKey);

      items.forEach((item) => {
        if (!item.isLeader) return;
        const rankingScore = item.metricType === 'glassAttachmentRate'
          ? (Number(item.weightedScore) || 0)
          : (Number(item.qty) || 0);
        leaders.push({
          employeeName,
          location,
          metricName: item.product,
          metricType: item.metricType || 'sales',
          value: item.qty,
          rankingScore,
          topPercent: item.topPercent,
          numerator: item.numerator,
          denominator: item.denominator,
        });
      });
    });

    leaders.sort((a, b) => {
      if (b.rankingScore !== a.rankingScore) return b.rankingScore - a.rankingScore;
      if ((b.denominator || 0) !== (a.denominator || 0)) return (b.denominator || 0) - (a.denominator || 0);
      if (b.value !== a.value) return b.value - a.value;
      if (a.metricType !== b.metricType) return a.metricType.localeCompare(b.metricType);
      if (a.metricName !== b.metricName) return a.metricName.localeCompare(b.metricName);
      return a.employeeName.localeCompare(b.employeeName);
    });

    let selectedLeaders = leaders.slice(0, MAX_COMPANY_LEADERS);
    const glassLeader = leaders.find((row) => row.metricType === 'glassAttachmentRate');
    const hasGlassLeader = selectedLeaders.some((row) => row.metricType === 'glassAttachmentRate');
    if (glassLeader && !hasGlassLeader) {
      if (selectedLeaders.length < MAX_COMPANY_LEADERS) {
        selectedLeaders.push(glassLeader);
      } else if (selectedLeaders.length > 0) {
        selectedLeaders[selectedLeaders.length - 1] = glassLeader;
      } else {
        selectedLeaders = [glassLeader];
      }
    }

    selectedLeaders.sort((a, b) => {
      if (b.rankingScore !== a.rankingScore) return b.rankingScore - a.rankingScore;
      if ((b.denominator || 0) !== (a.denominator || 0)) return (b.denominator || 0) - (a.denominator || 0);
      if (b.value !== a.value) return b.value - a.value;
      return a.employeeName.localeCompare(b.employeeName);
    });

    return {
      success: true,
      sourceTab: tabTitle,
      leaders: selectedLeaders,
    };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to load company leaders.', leaders: [] };
  }
});

ipcMain.handle('get-store-stats', async (_event, payload) => {
  try {
    const storeName = String(payload?.storeName || '').trim();
    const storeKey = normalizeName(storeName);
    if (!storeKey) return { success: false, message: 'Store name is required.' };

    const sheets = await getSheetsClient();
    const tabTitle = SALES_REPORT_TAB_NAME;
    const aggregate = await getSalesAggregateData(sheets, tabTitle);

    const totalUnits = aggregate.overallByStore.get(storeKey) || 0;
    const totalNetProfit = aggregate.netProfitByStore.get(storeKey) || 0;
    const overallStoreCount = aggregate.overallByStore.size;
    if (totalUnits <= 0 || overallStoreCount === 0) {
      return {
        success: true,
        sourceTab: tabTitle,
        storeName,
        stats: null,
      };
    }

    const allInvoiceCount = aggregate.storeInvoiceIds.get(storeKey)?.size || 0;
    const invoiceMetricsForStore = aggregate.storeInvoiceMetrics.get(storeKey) || new Map();
    const eligibleSelectedInvoices = Array.from(invoiceMetricsForStore.values())
      .filter((row) => !row.hasRepairItem);
    const gpInvoiceCount = eligibleSelectedInvoices.length;
    const eligibleNetProfit = eligibleSelectedInvoices
      .reduce((sum, row) => sum + (row.netProfitTotal || 0), 0);
    const avgGrossProfitPerInvoice = gpInvoiceCount > 0 ? eligibleNetProfit / gpInvoiceCount : 0;

    const allStoreMetrics = Array.from(aggregate.storeInvoiceMetrics.entries())
      .map(([key, invoiceMap]) => {
        const eligibleInvoices = Array.from(invoiceMap.values()).filter((row) => !row.hasRepairItem);
        const invoices = eligibleInvoices.length;
        const gp = eligibleInvoices.reduce((sum, row) => sum + (row.netProfitTotal || 0), 0);
        const avg = invoices > 0 ? gp / invoices : 0;
        return {
          key,
          invoices,
          avg,
        };
      })
      .filter((row) => row.invoices > 0);
    const rankStoreCount = allStoreMetrics.length;
    const higherCount = gpInvoiceCount > 0
      ? allStoreMetrics.filter((row) => row.avg > avgGrossProfitPerInvoice).length
      : rankStoreCount;
    const rank = higherCount + 1;
    const topPercent = gpInvoiceCount <= 0
      ? 100
      : (rankStoreCount <= 1
          ? 1
          : Math.round((higherCount / (rankStoreCount - 1)) * 99) + 1);
    const topProductMap = aggregate.storeProductTotals.get(storeKey) || new Map();
    let topProductName = '';
    let topProductQty = 0;
    topProductMap.forEach((qty, product) => {
      if (qty > topProductQty) {
        topProductQty = qty;
        topProductName = product;
      }
    });

    return {
      success: true,
      sourceTab: tabTitle,
      storeName: aggregate.storeDisplayNames.get(storeKey) || storeName,
      stats: {
        totalUnits: Math.round(totalUnits),
        rank,
        storeCount: rankStoreCount,
        topPercent,
        invoiceCount: allInvoiceCount,
        gpInvoiceCount,
        totalNetProfit: Math.round(totalNetProfit * 100) / 100,
        avgGrossProfitPerInvoice: Math.round(avgGrossProfitPerInvoice * 100) / 100,
        topProduct: {
          name: topProductName || 'N/A',
          qty: Math.round(topProductQty),
        },
      },
    };
  } catch (err) {
    return { success: false, message: err?.message || 'Failed to load store stats.' };
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
