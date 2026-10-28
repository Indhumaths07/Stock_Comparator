import * as XLSX_BASE from 'xlsx';
import XLSX_STYLE from 'xlsx-js-style';
import { ColumnMapping, ComparisonItem, ComparisonSummary, MatchMode, UploadedFile } from '../types';

// Use xlsx-js-style for rich Excel generation with colors, and xlsx for parsing
const XLSX_WRITER = XLSX_STYLE || XLSX_BASE;
const XLSX_READER = XLSX_BASE;

/**
 * Normalizes text for case-insensitive and whitespace-tolerant matching
 */
export function normalizeText(val: any): string {
  if (val === undefined || val === null) return '';
  return String(val)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

export function normalizeKey(val: any): string {
  if (val === undefined || val === null) return '';
  return String(val)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Intelligent Spreadsheet Parser
 * Accurately extracts headers and rows from Wondersoft ERP and E-Commerce files.
 * If the file starts directly with data (like the 10-column e-commerce format):
 * Men | Bottom Wear | Pant Mens Aim Cotton Narrow-28 | AIM | LIGHT BLUE GREEN | 28, 30, 32... | Cotton | 5-P006-08 | 770 | 25
 * it automatically recognizes the structure and provides named headers!
 */
export function extractHeadersAndRows(worksheet: XLSX_BASE.WorkSheet): {
  headers: string[];
  rawRows: Record<string, any>[];
  headerRowIndex: number;
} {
  const sheetData: any[][] = XLSX_READER.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    blankrows: false,
  });

  if (!sheetData || sheetData.length === 0) {
    return { headers: [], rawRows: [], headerRowIndex: 0 };
  }

  // Header keywords for retail ERP and e-commerce stock files
  const headerKeywords = [
    'toon label',
    'toonlabel',
    'town label',
    'label',
    'brand',
    'style no',
    'styleno',
    'style_no',
    'style',
    'item code',
    'itemcode',
    'code',
    'product name',
    'productname',
    'item name',
    'item_name',
    'description',
    'desc',
    'colour',
    'color',
    'shade',
    'size',
    'sixe',
    'closing stock',
    'closing_stock',
    'cl.qty',
    'cl. stock',
    'cl stock',
    'closing qty',
    'total stock quantity',
    'total stock',
    'stock quantity',
    'stock',
    'qty',
    'column',
    'fabric',
    'material',
    'price',
    'mrp',
    'rate',
    'category',
    'department',
    'main category',
    'product type',
    'store name',
    'storename',
    'store',
    'branch',
    'ean',
    'barcode',
  ];

  let bestRowIndex = -1;
  let maxScore = 0;

  const scanLimit = Math.min(25, sheetData.length);
  for (let r = 0; r < scanLimit; r++) {
    const row = sheetData[r] || [];
    if (!Array.isArray(row) || row.length === 0) continue;

    const nonEmptyCells = row.map((c) => String(c ?? '').trim()).filter(Boolean);
    if (nonEmptyCells.length < 2) continue; // title banner like "GM FASHIONS"

    let keywordMatches = 0;
    for (const cell of nonEmptyCells) {
      const lower = cell.toLowerCase().replace(/[^a-z0-9]/g, ' ');
      if (headerKeywords.some((kw) => lower.includes(kw) || kw.includes(lower))) {
        keywordMatches++;
      }
    }

    const textCellCount = nonEmptyCells.filter((c) => isNaN(Number(c))).length;
    const score = keywordMatches * 25 + textCellCount * 2 + nonEmptyCells.length;

    if (score > maxScore && keywordMatches >= 2) {
      maxScore = score;
      bestRowIndex = r;
    }
  }

  // Check if row 0 is raw e-commerce apparel data WITHOUT a header row
  // Format: [Department, Category, Product Name, Toon Label, Colour, Size, Column, Style No, Price, Stock]
  const row0 = sheetData[0] || [];
  const isRow0ApparelData =
    bestRowIndex === -1 &&
    row0.length >= 7 &&
    typeof row0[2] === 'string' &&
    row0[2].length > 3 && // Product Name
    (typeof row0[3] === 'string' && row0[3].trim().length >= 2) && // Toon Label (e.g. AIM)
    (typeof row0[4] === 'string' && /blue|green|grey|navy|black|white|red|yellow|pink|cotton|denim/i.test(row0[4])) && // Colour
    !isNaN(Number(String(row0[row0.length - 1] ?? '').trim())); // Ending with stock quantity

  if (isRow0ApparelData) {
    // Standard synthetic headers matching user's exact 10-column e-commerce layout
    const syntheticHeaders = [
      'Main Category',
      'Product Type',
      'Product Name',
      'Toon Label',
      'Colour',
      'Size',
      'Column',
      'Style No',
      'Price',
      'Total Stock Quantity',
    ];

    // Ensure we have headers for all columns
    const headers: string[] = [];
    for (let c = 0; c < row0.length; c++) {
      headers.push(syntheticHeaders[c] || `Column_${c + 1}`);
    }

    const rawRows: Record<string, any>[] = [];
    for (let r = 0; r < sheetData.length; r++) {
      const row = sheetData[r] || [];
      if (!Array.isArray(row) || row.length === 0) continue;
      const hasData = row.some((cell) => String(cell ?? '').trim() !== '');
      if (!hasData) continue;

      const rowObj: Record<string, any> = {};
      for (let c = 0; c < headers.length; c++) {
        rowObj[headers[c]] = row[c] !== undefined ? row[c] : '';
      }
      rawRows.push(rowObj);
    }

    return {
      headers,
      rawRows,
      headerRowIndex: 0,
    };
  }

  // Fallback: If no keyword matched, find first row with >= 3 text cells
  if (bestRowIndex === -1) {
    for (let r = 0; r < scanLimit; r++) {
      const row = sheetData[r] || [];
      const textCells = row
        .map((c) => String(c ?? '').trim())
        .filter((c) => c && isNaN(Number(c)));
      if (textCells.length >= 3) {
        bestRowIndex = r;
        break;
      }
    }
  }

  if (bestRowIndex === -1) bestRowIndex = 0;

  const headerRow = sheetData[bestRowIndex] || [];

  // Find last non-empty column index to trim trailing blank columns
  let lastNonEmptyCol = -1;
  for (let c = headerRow.length - 1; c >= 0; c--) {
    if (String(headerRow[c] ?? '').trim() !== '') {
      lastNonEmptyCol = c;
      break;
    }
  }

  const effectiveColCount = lastNonEmptyCol >= 0 ? lastNonEmptyCol + 1 : headerRow.length;
  const headers: string[] = [];
  const seenHeaders = new Map<string, number>();

  for (let c = 0; c < effectiveColCount; c++) {
    let rawH = String(headerRow[c] ?? '').trim();
    if (!rawH || rawH.startsWith('__EMPTY')) {
      rawH = `Column_${c + 1}`;
    }

    let uniqueH = rawH;
    if (seenHeaders.has(rawH.toLowerCase())) {
      const count = seenHeaders.get(rawH.toLowerCase())! + 1;
      seenHeaders.set(rawH.toLowerCase(), count);
      uniqueH = `${rawH}_${count}`;
    } else {
      seenHeaders.set(rawH.toLowerCase(), 1);
    }
    headers.push(uniqueH);
  }

  // Construct rawRows starting from the row after headers
  const rawRows: Record<string, any>[] = [];
  for (let r = bestRowIndex + 1; r < sheetData.length; r++) {
    const row = sheetData[r] || [];
    if (!Array.isArray(row) || row.length === 0) continue;

    const hasData = row.some((cell) => String(cell ?? '').trim() !== '');
    if (!hasData) continue;

    const firstCell = String(row[0] ?? '').trim().toLowerCase();
    const secondCell = String(row[1] ?? '').trim().toLowerCase();
    if (
      firstCell.startsWith('total') ||
      firstCell.startsWith('grand total') ||
      secondCell.startsWith('total') ||
      secondCell.startsWith('grand total')
    ) {
      continue;
    }

    const rowObj: Record<string, any> = {};
    for (let c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = row[c] !== undefined ? row[c] : '';
    }
    rawRows.push(rowObj);
  }

  return {
    headers,
    rawRows,
    headerRowIndex: bestRowIndex,
  };
}

export async function parseSpreadsheetFile(file: File): Promise<UploadedFile> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX_READER.read(buffer, {
    type: 'array',
    cellDates: true,
    cellNF: false,
    cellText: false,
  });

  const sheets = workbook.SheetNames;
  const selectedSheet = sheets[0] || '';
  const worksheet = workbook.Sheets[selectedSheet];

  const { headers, rawRows, headerRowIndex } = worksheet
    ? extractHeadersAndRows(worksheet)
    : { headers: [], rawRows: [], headerRowIndex: 0 };

  return {
    file,
    name: file.name,
    size: file.size,
    sheets,
    selectedSheet,
    headers,
    rawRows,
    headerRowIndex,
  };
}

export function changeSpreadsheetSheet(
  workbookBuffer: ArrayBuffer,
  sheetName: string
): { headers: string[]; rawRows: Record<string, any>[] } {
  const workbook = XLSX_READER.read(workbookBuffer, { type: 'array' });
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) return { headers: [], rawRows: [] };

  const { headers, rawRows } = extractHeadersAndRows(worksheet);
  return { headers, rawRows };
}

/**
 * Checks whether a column contains barcode/EAN-like numbers (e.g. 8905000083746)
 */
function isNumericBarcodeSample(rawRows: Record<string, any>[], colName: string): boolean {
  if (!colName || rawRows.length === 0) return false;
  let matches = 0;
  const sampleCount = Math.min(20, rawRows.length);
  for (let i = 0; i < sampleCount; i++) {
    const rawVal = rawRows[i]?.[colName];
    if (rawVal === undefined || rawVal === null) continue;
    const val = String(rawVal).trim().replace(/\.0+$/, '');
    if (/^\d{8,16}$/.test(val) || /^[+-]?\d+(\.\d+)?[eE][+-]?\d+$/.test(val)) {
      matches++;
    }
  }
  return matches >= 1;
}

/**
 * Intelligent Column Auto-Detection:
 * Defaults to 'toon_label' match mode as requested.
 * Detects Toon Label, Style No, Color, Size, Closing Stock, and E-Commerce fields.
 */
export function autoDetectColumns(
  wondersoftHeaders: string[],
  ecommerceHeaders: string[],
  wondersoftRows: Record<string, any>[] = [],
  ecommerceRows: Record<string, any>[] = []
): ColumnMapping {
  // 1. Separate Store Name detection
  const wsStore =
    findHeader(wondersoftHeaders, ['storename', 'store name', 'store', 'branch', 'shop']) || '';
  const ecomStore =
    findHeader(ecommerceHeaders, ['storename', 'store name', 'store', 'branch', 'shop']) || '';

  // 2. Toon Label detection (Primary matching key!)
  const wsToon =
    findHeader(wondersoftHeaders, [
      'toon label',
      'toonlabel',
      'town label',
      'label',
      'brand',
      'sub brand',
      'vendor',
    ]) || '';

  const ecomToon =
    findHeader(ecommerceHeaders, [
      'toon label',
      'toonlabel',
      'town label',
      'label',
      'brand',
    ]) || '';

  // 3. Style No detection
  const wsStyle =
    findHeader(wondersoftHeaders, [
      'style no',
      'styleno',
      'style_no',
      'style',
      'item code',
      'itemcode',
      'item_code',
      'article no',
      'article',
      'sku',
    ]) || '';

  const ecomStyle =
    findHeader(ecommerceHeaders, [
      'style no',
      'styleno',
      'style_no',
      'style',
      'item no',
      'sku',
      'item code',
      'code',
    ]) || '';

  // 4. Stock Quantities
  // STRICT RULE: Reject Op.Qty, Sale.Qty, Pur.Qty, Adj.Qty, Return Qty!
  // Closing stock (Cl.Qty / Closing Stock) is the true ERP stock!
  const isInvalidStockHeader = (h: string) =>
    /op\.?qty|opening|pur\.?qty|purchase|sale\.?qty|sales|adj\.?qty|adjust|ret\.?qty|return/i.test(h);

  const cleanWsStockHeaders = wondersoftHeaders.filter((h) => !isInvalidStockHeader(h));

  const wsStock =
    findHeader(cleanWsStockHeaders, [
      'cl.qty',
      'cl. stock',
      'cl stock',
      'clqty',
      'clstock',
      'closing stock',
      'closing qty',
      'closing',
      'closing_stock',
      'closing_qty',
      'bal.qty',
      'balance qty',
      'bal qty',
      'balance',
      'current stock',
      'physical stock',
      'stock',
      'qty',
    ]) || cleanWsStockHeaders.find((h) => /closing|cl\.?qty|stock|balance/i.test(h)) || '';

  const ecomStock =
    findHeader(ecommerceHeaders, [
      'total stock quantity',
      'total stock qnatity',
      'total stock',
      'stock quantity',
      'stock',
      'qty',
      'quantity',
      'available',
      'inventory',
    ]) ||
    ecommerceHeaders.find((h) => /^stock$/i.test(h)) ||
    '';

  // 5. Product Name
  const ecomName =
    findHeader(ecommerceHeaders, [
      'productname',
      'product name',
      'item name',
      'item_name',
      'description',
      'desc',
      'title',
    ]) || '';

  const wsName =
    findHeader(wondersoftHeaders, [
      'item name',
      'item_name',
      'product name',
      'description',
      'item desc',
      'desc',
    ]) || '';

  // 6. E-Commerce specific details
  const ecomMainCategory =
    findHeader(ecommerceHeaders, [
      'main category',
      'maincategory',
      'main_category',
      'department',
      'category',
      'gender',
    ]) || '';

  const ecomProductType =
    findHeader(ecommerceHeaders, [
      'product type',
      'producttype',
      'product_type',
      'type',
      'sub category',
      'subtype',
    ]) || '';

  const ecomColor =
    findHeader(ecommerceHeaders, ['colour', 'color', 'shade']) || '';
  const wsColor =
    findHeader(wondersoftHeaders, ['colour', 'color', 'shade']) || '';

  const ecomSize =
    findHeader(ecommerceHeaders, ['size', 'sixe', 'item size', 'variant']) || '';
  const wsSize =
    findHeader(wondersoftHeaders, ['size', 'sixe', 'item size']) || '';

  const ecomColumnExtra =
    findHeader(ecommerceHeaders, ['column', 'col', 'fabric', 'material']) || '';

  const ecomPrice =
    findHeader(ecommerceHeaders, ['price', 'mrp', 'rate', 'selling price', 'amount']) || '';

  // Optional EAN / Barcodes
  const isExcludedFromEan = (h: string) =>
    /store|branch|shop|name|product|label|size|sixe|colour|color|style|qty|stock|rate|mrp|price|amount/i.test(
      h
    );

  const ecomEanCandidates = ecommerceHeaders.filter((h) => !isExcludedFromEan(h));
  const ecomEan =
    findHeader(ecommerceHeaders, ['ean', 'eancode', 'ean code', 'barcode', 'bar code']) ||
    ecomEanCandidates.find((h) => isNumericBarcodeSample(ecommerceRows, h)) ||
    '';

  const wsEanCandidates = wondersoftHeaders.filter((h) => !isExcludedFromEan(h));
  const wsEan =
    findHeader(wondersoftHeaders, ['barcode', 'bar code', 'ean', 'eancode', 'ean code']) ||
    wsEanCandidates.find((h) => isNumericBarcodeSample(wondersoftRows, h)) ||
    '';

  // Determine optimal match mode:
  // If Toon Label is present in either file, default to 'toon_label' (user request!)
  let matchMode: MatchMode = 'toon_label';
  if (ecomToon || wsToon) {
    matchMode = 'toon_label';
  } else if (ecomStyle || wsStyle) {
    matchMode = 'style_no';
  } else if (ecomEan && wsEan) {
    matchMode = 'ean';
  }

  return {
    matchMode,
    toonLabelColumn: ecomToon || wsToon,
    wondersoftToonLabelColumn: wsToon,
    ecommerceToonLabelColumn: ecomToon,
    styleNoColumn: ecomStyle || wsStyle,
    wondersoftStyleNoColumn: wsStyle,
    ecommerceStyleNoColumn: ecomStyle,
    skuColumn: ecomEan || wsEan || ecomStyle || wsStyle || '',
    wondersoftSkuColumn: wsEan,
    ecommerceSkuColumn: ecomEan,
    nameColumn: ecomName || wsName || '',
    wondersoftNameColumn: wsName,
    ecommerceNameColumn: ecomName,
    wondersoftQtyColumn: wsStock,
    ecommerceQtyColumn: ecomStock,
    mainCategoryColumn: ecomMainCategory,
    productTypeColumn: ecomProductType,
    wondersoftColorColumn: wsColor,
    ecommerceColorColumn: ecomColor,
    colorColumn: ecomColor || wsColor,
    wondersoftSizeColumn: wsSize,
    ecommerceSizeColumn: ecomSize,
    sizeColumn: ecomSize || wsSize,
    priceColumn: ecomPrice,
    wondersoftStoreNameColumn: wsStore,
    ecommerceStoreNameColumn: ecomStore,
    storeNameColumn: ecomStore || wsStore,
    columnExtraColumn: ecomColumnExtra,
    storeComparisonMode: 'all',
  };
}

function findHeader(headers: string[], candidates: string[]): string | null {
  const cleanList = headers.map((h) => ({
    original: h,
    lower: h.toLowerCase().replace(/[^a-z0-9]/g, ' ').trim(),
  }));

  for (const cand of candidates) {
    const cLower = cand.toLowerCase().replace(/[^a-z0-9]/g, ' ').trim();
    const exact = cleanList.find((item) => item.lower === cLower);
    if (exact) return exact.original;
  }

  for (const cand of candidates) {
    const cLower = cand.toLowerCase().replace(/[^a-z0-9]/g, ' ').trim();
    const partial = cleanList.find((item) => item.lower.includes(cLower) || cLower.includes(item.lower));
    if (partial) return partial.original;
  }

  return null;
}

export function normalizeIdentifier(val: any): string {
  if (val === undefined || val === null) return '';
  let str = '';
  if (typeof val === 'number') {
    if (Number.isFinite(val) && Math.floor(val) === val) {
      try {
        str = BigInt(val).toString();
      } catch {
        str = val.toLocaleString('fullwide', { useGrouping: false });
      }
    } else {
      str = String(val);
    }
  } else {
    str = String(val);
  }

  str = str.trim().replace(/\.0+$/, '');

  if (/^[+-]?\d+(\.\d+)?[eE][+-]?\d+$/.test(str)) {
    const num = Number(str);
    if (!isNaN(num)) {
      try {
        str = BigInt(Math.round(num)).toString();
      } catch {
        str = num.toLocaleString('fullwide', { useGrouping: false });
      }
    }
  }

  return str.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
}

/**
 * Extracts list of distinct store names from rows
 */
export function extractUniqueStores(rows: Record<string, any>[], storeCol?: string): string[] {
  if (!storeCol) return [];
  const set = new Set<string>();
  for (const row of rows) {
    const val = String(row[storeCol] ?? '').trim();
    if (val) set.add(val);
  }
  return Array.from(set).sort();
}

/**
 * Entry holding aggregated stock for all Wondersoft ERP rows matching an item / style
 * (Crucial for summing stock across multiple sizes: e.g. 28:5, 30:5, 32:6, 34:4, 36:5 = 25)
 */
interface WondersoftAggregatedEntry {
  rows: Record<string, any>[];
  totalQty: number;
  sizes: { size: string; qty: number }[];
  toonLabel: string;
  styleNo: string;
  color: string;
  name: string;
  storeName: string;
}

/**
 * Core Stock Comparison Engine
 * Supports:
 * - Matching by Toon Label, Toon Label + Style No, Style No, or EAN
 * - Correctly CALCULATES Wondersoft stock by SUMMING all size rows for that product
 * - Provides breakdown of sizes contributing to the stock
 * - Computes difference = targetQty - ecommerceQty
 */
export function compareStockData(
  wondersoftRows: Record<string, any>[],
  ecommerceRows: Record<string, any>[],
  mapping: {
    matchMode?: MatchMode;
    wondersoftSkuCol?: string;
    ecommerceSkuCol?: string;
    wondersoftToonLabelCol?: string;
    ecommerceToonLabelCol?: string;
    wondersoftStyleNoCol?: string;
    ecommerceStyleNoCol?: string;
    wondersoftColorCol?: string;
    ecommerceColorCol?: string;
    wondersoftSizeCol?: string;
    ecommerceSizeCol?: string;
    wondersoftNameCol?: string;
    ecommerceNameCol?: string;
    wondersoftQtyCol: string;
    ecommerceQtyCol: string;
    mainCategoryCol?: string;
    productTypeCol?: string;
    wondersoftStoreNameCol?: string;
    ecommerceStoreNameCol?: string;
    storeNameCol?: string;
    toonLabelCol?: string;
    styleNoCol?: string;
    colorCol?: string;
    sizeCol?: string;
    priceCol?: string;
    columnExtraCol?: string;
    storeComparisonMode?: 'all' | 'exact_store' | 'filter_store';
    selectedStore?: string;
  }
): { items: ComparisonItem[]; summary: ComparisonSummary } {
  const availableStoresSet = new Set<string>();
  const distinctToonLabelsSet = new Set<string>();

  const wsHeaders = wondersoftRows.length > 0 ? Object.keys(wondersoftRows[0]) : [];

  // Determine effective column names
  const effectiveWsToonCol =
    mapping.wondersoftToonLabelCol ||
    mapping.toonLabelCol ||
    findHeader(wsHeaders, ['toon label', 'toonlabel', 'town label', 'label', 'brand']) ||
    '';

  const effectiveEcomToonCol =
    mapping.ecommerceToonLabelCol ||
    mapping.toonLabelCol ||
    '';

  const effectiveWsStyleCol =
    mapping.wondersoftStyleNoCol ||
    mapping.styleNoCol ||
    findHeader(wsHeaders, ['style no', 'styleno', 'style_no', 'style', 'item code', 'itemcode', 'article no']) ||
    '';

  const effectiveEcomStyleCol =
    mapping.ecommerceStyleNoCol ||
    mapping.styleNoCol ||
    '';

  const effectiveWsNameCol =
    mapping.wondersoftNameCol ||
    findHeader(wsHeaders, ['item name', 'item_name', 'product name', 'description']) ||
    '';

  const effectiveWsColorCol =
    mapping.wondersoftColorCol ||
    mapping.colorCol ||
    findHeader(wsHeaders, ['colour', 'color', 'shade']) ||
    '';

  const effectiveWsSizeCol =
    mapping.wondersoftSizeCol ||
    mapping.sizeCol ||
    findHeader(wsHeaders, ['size', 'sixe']) ||
    '';

  // Self-healing Wondersoft Stock column
  let effectiveWsQtyCol = mapping.wondersoftQtyCol;
  const isInvalidQty = (col: string) =>
    /op\.?qty|opening|pur\.?qty|purchase|sale\.?qty|sales|ret\.?qty|return|adj\.?qty|adjust/i.test(col);

  let hasNonZeroStock = false;
  if (effectiveWsQtyCol && !isInvalidQty(effectiveWsQtyCol)) {
    for (const r of wondersoftRows) {
      if (parseNumber(r[effectiveWsQtyCol]) > 0) {
        hasNonZeroStock = true;
        break;
      }
    }
  }

  if (!hasNonZeroStock) {
    const stockCandidates = [
      'cl.qty',
      'cl. stock',
      'cl stock',
      'clqty',
      'clstock',
      'closing stock',
      'closing qty',
      'closing',
      'closing_stock',
      'closing_qty',
      'bal.qty',
      'balance qty',
      'bal qty',
      'balance',
      'current stock',
      'physical stock',
      'stock',
      'qty',
    ];

    for (const cand of stockCandidates) {
      const found = wsHeaders.find((h) => {
        const cleanH = h.toLowerCase().replace(/[^a-z0-9]/g, ' ').trim();
        const cleanCand = cand.toLowerCase().replace(/[^a-z0-9]/g, ' ').trim();
        return (cleanH === cleanCand || cleanH.includes(cleanCand)) && !isInvalidQty(h);
      });
      if (found) {
        const checkPositive = wondersoftRows.some((r) => parseNumber(r[found]) > 0);
        if (checkPositive) {
          effectiveWsQtyCol = found;
          break;
        }
      }
    }
  }

  // Multi-index Maps for Wondersoft ERP
  // Aggregates rows with multiple sizes into single calculated stock!
  const wsByToonAndStyle = new Map<string, WondersoftAggregatedEntry>();
  const wsByToonAndColor = new Map<string, WondersoftAggregatedEntry>();
  const wsByStyleOnly = new Map<string, WondersoftAggregatedEntry>();
  const wsByToonOnly = new Map<string, WondersoftAggregatedEntry>();
  const wsByEan = new Map<string, WondersoftAggregatedEntry>();

  const addToEntry = (
    map: Map<string, WondersoftAggregatedEntry>,
    key: string,
    row: Record<string, any>,
    qty: number,
    size: string,
    toon: string,
    style: string,
    color: string,
    name: string,
    store: string
  ) => {
    if (!key) return;
    const existing = map.get(key);
    if (existing) {
      existing.rows.push(row);
      existing.totalQty += qty;
      if (size) existing.sizes.push({ size, qty });
    } else {
      map.set(key, {
        rows: [row],
        totalQty: qty,
        sizes: size ? [{ size, qty }] : [],
        toonLabel: toon,
        styleNo: style,
        color,
        name,
        storeName: store,
      });
    }
  };

  for (const row of wondersoftRows) {
    const rawToon = effectiveWsToonCol ? String(row[effectiveWsToonCol] || '').trim() : '';
    const normToon = normalizeKey(rawToon);

    const rawStyle = effectiveWsStyleCol ? String(row[effectiveWsStyleCol] || '').trim() : '';
    const normStyle = normalizeKey(rawStyle);

    const rawColor = effectiveWsColorCol ? String(row[effectiveWsColorCol] || '').trim() : '';
    const normColor = normalizeKey(rawColor);

    const rawName = effectiveWsNameCol ? String(row[effectiveWsNameCol] || '').trim() : '';

    const rawSize = effectiveWsSizeCol ? String(row[effectiveWsSizeCol] || '').trim() : '';

    const rawEan = mapping.wondersoftSkuCol ? String(row[mapping.wondersoftSkuCol] || '').trim() : '';
    const normEan = normalizeIdentifier(rawEan);

    const storeName = mapping.wondersoftStoreNameCol
      ? String(row[mapping.wondersoftStoreNameCol] || '').trim()
      : '';
    const normStore = normalizeKey(storeName);

    if (storeName) availableStoresSet.add(storeName);
    if (rawToon) distinctToonLabelsSet.add(rawToon);

    const qtyVal = parseNumber(row[effectiveWsQtyCol]);

    // 1. Toon + Style index
    if (normToon && normStyle) {
      addToEntry(wsByToonAndStyle, `${normToon}:::${normStyle}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      if (normStore) {
        addToEntry(wsByToonAndStyle, `${normStore}:::${normToon}:::${normStyle}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      }
    }

    // 2. Toon + Color index
    if (normToon && normColor) {
      addToEntry(wsByToonAndColor, `${normToon}:::${normColor}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      if (normStore) {
        addToEntry(wsByToonAndColor, `${normStore}:::${normToon}:::${normColor}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      }
    }

    // 3. Style No only index
    if (normStyle) {
      addToEntry(wsByStyleOnly, normStyle, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      if (normStore) {
        addToEntry(wsByStyleOnly, `${normStore}:::${normStyle}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      }
    }

    // 4. Toon Label only index (Aggregates all styles/sizes under this Toon Label)
    if (normToon) {
      addToEntry(wsByToonOnly, normToon, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      if (normStore) {
        addToEntry(wsByToonOnly, `${normStore}:::${normToon}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      }
    }

    // 5. EAN index
    if (normEan) {
      addToEntry(wsByEan, normEan, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      const stripped = normEan.replace(/^0+/, '');
      if (stripped && stripped !== normEan) {
        addToEntry(wsByEan, stripped, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      }
      if (normStore) {
        addToEntry(wsByEan, `${normStore}:::${normEan}`, row, qtyVal, rawSize, rawToon, rawStyle, rawColor, rawName, storeName);
      }
    }
  }

  // Iterate E-Commerce items and compare
  const comparisonItems: ComparisonItem[] = [];

  let stockDecreasedCount = 0;
  let outOfStockCount = 0;
  let inSyncCount = 0;
  let stockIncreasedCount = 0;
  let onlyInWondersoftCount = 0;
  let onlyInEcommerceCount = 0;
  let totalUnitsDecreased = 0;

  let idx = 0;
  for (const row of ecommerceRows) {
    idx++;

    const rawToon = effectiveEcomToonCol ? String(row[effectiveEcomToonCol] || '').trim() : '';
    const normToon = normalizeKey(rawToon);

    const rawStyle = effectiveEcomStyleCol ? String(row[effectiveEcomStyleCol] || '').trim() : '';
    const normStyle = normalizeKey(rawStyle);

    const rawColor = mapping.ecommerceColorCol || mapping.colorCol
      ? String(row[mapping.ecommerceColorCol || mapping.colorCol || ''] || '').trim()
      : '';
    const normColor = normalizeKey(rawColor);

    const productName = mapping.ecommerceNameCol
      ? String(row[mapping.ecommerceNameCol] || '').trim()
      : '';

    const mainCategory = mapping.mainCategoryCol
      ? String(row[mapping.mainCategoryCol] || '').trim()
      : '';
    const productType = mapping.productTypeCol
      ? String(row[mapping.productTypeCol] || '').trim()
      : '';
    const size = mapping.ecommerceSizeCol || mapping.sizeCol
      ? String(row[mapping.ecommerceSizeCol || mapping.sizeCol || ''] || '').trim()
      : '';
    const columnExtra = mapping.columnExtraCol
      ? String(row[mapping.columnExtraCol] || '').trim()
      : '';
    const priceVal = mapping.priceCol ? parseNumber(row[mapping.priceCol]) : undefined;

    const rawEan = mapping.ecommerceSkuCol ? String(row[mapping.ecommerceSkuCol] || '').trim() : '';
    const normEan = normalizeIdentifier(rawEan);

    const storeName = mapping.ecommerceStoreNameCol || mapping.storeNameCol
      ? String(row[mapping.ecommerceStoreNameCol || mapping.storeNameCol || ''] || '').trim()
      : '';
    const normStore = normalizeKey(storeName);

    if (storeName) availableStoresSet.add(storeName);
    if (rawToon) distinctToonLabelsSet.add(rawToon);

    // Filter store if selected
    if (
      mapping.storeComparisonMode === 'filter_store' &&
      mapping.selectedStore &&
      mapping.selectedStore !== 'all' &&
      storeName.toLowerCase() !== mapping.selectedStore.toLowerCase()
    ) {
      continue;
    }

    const ecomQty = parseNumber(row[mapping.ecommerceQtyCol]);

    // Matching Algorithm:
    // If exact_store mode is active, try store-prefixed keys first
    const isExactStore = mapping.storeComparisonMode === 'exact_store' && normStore;
    let matchedEntry: WondersoftAggregatedEntry | undefined;

    // Tier 1: Toon Label + Style No (Best & most accurate match!)
    if (normToon && normStyle) {
      if (isExactStore) {
        matchedEntry = wsByToonAndStyle.get(`${normStore}:::${normToon}:::${normStyle}`);
      }
      if (!matchedEntry) {
        matchedEntry = wsByToonAndStyle.get(`${normToon}:::${normStyle}`);
      }
    }

    // Tier 2: Style No alone (e.g. 5-P006-08)
    if (!matchedEntry && normStyle) {
      if (isExactStore) {
        matchedEntry = wsByStyleOnly.get(`${normStore}:::${normStyle}`);
      }
      if (!matchedEntry) {
        matchedEntry = wsByStyleOnly.get(normStyle);
      }
    }

    // Tier 3: Toon Label + Colour (e.g. AIM + LIGHT BLUE GREEN)
    if (!matchedEntry && normToon && normColor) {
      if (isExactStore) {
        matchedEntry = wsByToonAndColor.get(`${normStore}:::${normToon}:::${normColor}`);
      }
      if (!matchedEntry) {
        matchedEntry = wsByToonAndColor.get(`${normToon}:::${normColor}`);
      }
    }

    // Tier 4: EAN if present
    if (!matchedEntry && normEan) {
      if (isExactStore) {
        matchedEntry = wsByEan.get(`${normStore}:::${normEan}`);
      }
      if (!matchedEntry) {
        matchedEntry =
          wsByEan.get(normEan) ||
          wsByEan.get(normEan.replace(/^0+/, '')) ||
          wsByEan.get('0' + normEan);
      }
    }

    // Tier 5: If matching by Toon Label alone and this is the only entry or exact match
    if (!matchedEntry && normToon) {
      if (isExactStore) {
        matchedEntry = wsByToonOnly.get(`${normStore}:::${normToon}`);
      }
      if (!matchedEntry && mapping.matchMode === 'toon_label') {
        matchedEntry = wsByToonOnly.get(normToon);
      }
    }

    if (matchedEntry) {
      const wsQty = matchedEntry.totalQty;
      const diff = wsQty - ecomQty;
      const isOos = wsQty <= 0 && ecomQty > 0;
      const hasDec = wsQty < ecomQty;
      const inSync = wsQty === ecomQty;

      let status: ComparisonItem['status'] = 'in_sync';
      if (isOos) {
        status = 'out_of_stock';
        outOfStockCount++;
        stockDecreasedCount++;
        totalUnitsDecreased += ecomQty - wsQty;
      } else if (hasDec) {
        status = 'stock_decreased';
        stockDecreasedCount++;
        totalUnitsDecreased += ecomQty - wsQty;
      } else if (inSync) {
        status = 'in_sync';
        inSyncCount++;
      } else {
        status = 'stock_increased';
        stockIncreasedCount++;
      }

      // Generate human-friendly size breakdown if multiple sizes exist
      let matchedBreakdown = '';
      if (matchedEntry.sizes.length > 1) {
        const sizeParts = matchedEntry.sizes
          .slice(0, 6)
          .map((s) => `${s.size}:${s.qty}`)
          .join(', ');
        matchedBreakdown = `${matchedEntry.sizes.length} sizes: ${sizeParts}${
          matchedEntry.sizes.length > 6 ? '...' : ''
        }`;
      } else if (matchedEntry.rows.length > 1) {
        matchedBreakdown = `${matchedEntry.rows.length} ERP rows summed`;
      }

      comparisonItems.push({
        id: `ecom-${idx}-${normToon}-${normStyle || normEan || idx}`,
        sku: rawStyle || rawEan || `ITEM-${idx}`,
        ean: rawEan,
        mainCategory,
        productType,
        storeName: storeName || matchedEntry.storeName || '',
        name: productName || matchedEntry.name || 'Product',
        toonLabel: rawToon || matchedEntry.toonLabel || '',
        color: rawColor || matchedEntry.color || '',
        size,
        columnExtra,
        styleNo: rawStyle || matchedEntry.styleNo || '',
        price: priceVal,
        wondersoftQty: wsQty,
        ecommerceQty: ecomQty,
        targetQty: wsQty, // New target stock set to calculated Wondersoft closing stock
        difference: diff,
        status,
        hasDecreased: hasDec,
        isOutOfStock: isOos,
        isInSync: inSync,
        matchedCount: matchedEntry.rows.length,
        matchedBreakdown,
        sourceWondersoftRow: matchedEntry.rows[0],
        sourceWondersoftRows: matchedEntry.rows,
        sourceEcommerceRow: row,
      });
    } else {
      // Product not found in Wondersoft ERP file
      onlyInEcommerceCount++;
      const diff = 0 - ecomQty;
      const isOos = ecomQty > 0;

      comparisonItems.push({
        id: `ecom-only-${idx}-${normToon}-${normStyle || normEan || idx}`,
        sku: rawStyle || rawEan || `ITEM-${idx}`,
        ean: rawEan,
        mainCategory,
        productType,
        storeName,
        name: productName || 'Product',
        toonLabel: rawToon,
        color: rawColor,
        size,
        columnExtra,
        styleNo: rawStyle,
        price: priceVal,
        wondersoftQty: 0,
        ecommerceQty: ecomQty,
        targetQty: 0,
        difference: diff,
        status: isOos ? 'out_of_stock' : 'in_sync',
        hasDecreased: ecomQty > 0,
        isOutOfStock: isOos,
        isInSync: ecomQty === 0,
        matchedCount: 0,
        sourceEcommerceRow: row,
      });

      if (ecomQty > 0) {
        stockDecreasedCount++;
        totalUnitsDecreased += ecomQty;
      }
    }
  }

  return {
    items: comparisonItems,
    summary: {
      totalCompared: comparisonItems.length,
      stockDecreasedCount,
      outOfStockCount,
      inSyncCount,
      stockIncreasedCount,
      onlyInWondersoftCount,
      onlyInEcommerceCount,
      totalUnitsDecreased,
      availableStores: Array.from(availableStoresSet).sort(),
      distinctToonLabels: Array.from(distinctToonLabelsSet).sort(),
      totalToonLabels: distinctToonLabelsSet.size,
    },
  };
}

export function parseNumber(val: any): number {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : Math.round(val);
  const cleaned = String(val).replace(/,/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : Math.round(num);
}

export function formatStatusLabel(status: ComparisonItem['status']): string {
  switch (status) {
    case 'out_of_stock':
      return 'Out of Stock';
    case 'stock_decreased':
      return 'Stock Decreased';
    case 'in_sync':
      return 'In Sync';
    case 'stock_increased':
      return 'Stock Increased';
    case 'only_in_wondersoft':
      return 'Only in ERP';
    case 'only_in_ecommerce':
      return 'Only in E-Com';
    default:
      return status;
  }
}

/**
 * Downloads the updated E-Commerce file in original format with COLORED cell highlighting!
 * User requirement:
 * "aprom download panrapa na send panna xl la stock data changes iruntha atha vera colour la denote panni kattu ok va"
 * - All original columns are 100% preserved in exact original order.
 * - Changed stock cells are highlighted in distinct Amber/Yellow (or Red for OOS) so changes immediately stand out!
 * - Matches rows by Toon Label + Style No, Style No, Color, or index.
 */
export function downloadModifiedEcommerceFile({
  originalEcommerceRows,
  ecommerceQtyCol,
  comparisonItems,
  onlyChanged = false,
  baseName = 'ecommerce_inventory_updated',
}: {
  originalEcommerceRows: Record<string, any>[];
  ecommerceSkuCol?: string;
  ecommerceQtyCol: string;
  comparisonItems: ComparisonItem[];
  onlyChanged?: boolean;
  format?: 'xlsx';
  baseName?: string;
}) {
  // Multi-key lookup maps for robust matching
  const lookupByToonAndStyle = new Map<string, ComparisonItem>();
  const lookupByStyle = new Map<string, ComparisonItem>();
  const lookupByToonAndColor = new Map<string, ComparisonItem>();
  const lookupByIndex = new Map<number, ComparisonItem>();

  for (let i = 0; i < comparisonItems.length; i++) {
    const item = comparisonItems[i];
    lookupByIndex.set(i, item);

    const normToon = normalizeKey(item.toonLabel);
    const normStyle = normalizeKey(item.styleNo);
    const normColor = normalizeKey(item.color);

    if (normToon && normStyle) {
      lookupByToonAndStyle.set(`${normToon}:::${normStyle}`, item);
    }
    if (normStyle) {
      lookupByStyle.set(normStyle, item);
    }
    if (normToon && normColor) {
      lookupByToonAndColor.set(`${normToon}:::${normColor}`, item);
    }
  }

  const originalHeaders =
    originalEcommerceRows.length > 0 ? Object.keys(originalEcommerceRows[0]) : [];

  const stockColIndex = originalHeaders.indexOf(ecommerceQtyCol);

  const updatedRows: Record<string, any>[] = [];
  const rowHighlightIndices: {
    rowIndex: number;
    isOos: boolean;
    isInSync: boolean;
    isChanged: boolean;
  }[] = [];

  for (let r = 0; r < originalEcommerceRows.length; r++) {
    const originalRow = originalEcommerceRows[r];

    // Find row's toon label, style, color
    let rowToon = '';
    let rowStyle = '';
    let rowColor = '';

    for (const h of originalHeaders) {
      const lower = h.toLowerCase().replace(/[^a-z0-9]/g, ' ');
      if (!rowToon && (lower.includes('toon') || lower.includes('label') || lower.includes('brand'))) {
        rowToon = String(originalRow[h] || '').trim();
      }
      if (!rowStyle && (lower.includes('style') || lower.includes('item no') || lower.includes('item code'))) {
        rowStyle = String(originalRow[h] || '').trim();
      }
      if (!rowColor && (lower.includes('color') || lower.includes('colour') || lower.includes('shade'))) {
        rowColor = String(originalRow[h] || '').trim();
      }
    }

    const normToon = normalizeKey(rowToon);
    const normStyle = normalizeKey(rowStyle);
    const normColor = normalizeKey(rowColor);

    const matched =
      (normToon && normStyle ? lookupByToonAndStyle.get(`${normToon}:::${normStyle}`) : undefined) ||
      (normStyle ? lookupByStyle.get(normStyle) : undefined) ||
      (normToon && normColor ? lookupByToonAndColor.get(`${normToon}:::${normColor}`) : undefined) ||
      lookupByIndex.get(r);

    if (matched) {
      const isChanged: boolean = Boolean(
        matched.hasDecreased ||
        matched.isOutOfStock ||
        matched.difference !== 0 ||
        matched.isManuallyEdited
      );

      const isInSync: boolean = matched.difference === 0;

      if (onlyChanged && !isChanged) {
        continue;
      }

      const newRow: Record<string, any> = {};
      for (const col of originalHeaders) {
        newRow[col] = originalRow[col];
      }
      newRow[ecommerceQtyCol] = matched.targetQty;
      updatedRows.push(newRow);

      rowHighlightIndices.push({
        rowIndex: updatedRows.length, // 1-based data row in worksheet
        isOos: matched.targetQty <= 0,
        isInSync,
        isChanged,
      });
    } else {
      if (!onlyChanged) {
        const newRow: Record<string, any> = {};
        for (const col of originalHeaders) {
          newRow[col] = originalRow[col];
        }
        updatedRows.push(newRow);
      }
    }
  }

  const filename = `${baseName}_${onlyChanged ? 'changed_only' : 'updated'}_${
    new Date().toISOString().split('T')[0]
  }.xlsx`;

  const worksheet = XLSX_WRITER.utils.json_to_sheet(updatedRows, { header: originalHeaders });

  // Apply header styling
  for (let c = 0; c < originalHeaders.length; c++) {
    const cellAddress = XLSX_WRITER.utils.encode_cell({ r: 0, c });
    if (worksheet[cellAddress]) {
      const isStockHeader = c === stockColIndex;
      worksheet[cellAddress].s = {
        fill: { fgColor: { rgb: isStockHeader ? '1E40AF' : '1E293B' } },
        font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: {
          bottom: { style: 'medium', color: { rgb: '0F172A' } },
        },
      };
    }
  }

  // Apply colored cell highlighting:
  // - Difference 0 (In-Sync): Soft Emerald Green (D1FAE5 / 065F46)
  // - Out of Stock (0): Soft Red (FEE2E2 / B91C1C)
  // - Changed Stock (diff > 0): Soft Yellow (FEF08A / 854D0E)
  for (const info of rowHighlightIndices) {
    const r = info.rowIndex;
    if (stockColIndex >= 0) {
      const stockCellAddr = XLSX_WRITER.utils.encode_cell({ r, c: stockColIndex });
      if (worksheet[stockCellAddr]) {
        if (info.isInSync) {
          worksheet[stockCellAddr].s = {
            fill: { fgColor: { rgb: 'D1FAE5' } }, // Soft Emerald Green
            font: {
              color: { rgb: '065F46' }, // Dark Emerald Green
              bold: true,
              sz: 11,
            },
            border: {
              top: { style: 'thin', color: { rgb: '10B981' } },
              bottom: { style: 'thin', color: { rgb: '10B981' } },
              left: { style: 'thin', color: { rgb: '10B981' } },
              right: { style: 'thin', color: { rgb: '10B981' } },
            },
            alignment: { horizontal: 'center' },
          };
        } else if (info.isChanged) {
          worksheet[stockCellAddr].s = {
            fill: { fgColor: { rgb: info.isOos ? 'FEE2E2' : 'FEF08A' } }, // Soft Red or Yellow
            font: {
              color: { rgb: info.isOos ? 'B91C1C' : '854D0E' }, // Dark Red or Dark Amber
              bold: true,
              sz: 11,
            },
            border: {
              top: { style: 'medium', color: { rgb: info.isOos ? 'EF4444' : 'EAB308' } },
              bottom: { style: 'medium', color: { rgb: info.isOos ? 'EF4444' : 'EAB308' } },
              left: { style: 'medium', color: { rgb: info.isOos ? 'EF4444' : 'EAB308' } },
              right: { style: 'medium', color: { rgb: info.isOos ? 'EF4444' : 'EAB308' } },
            },
            alignment: { horizontal: 'center' },
          };
        }
      }
    }
  }

  const workbook = XLSX_WRITER.utils.book_new();
  XLSX_WRITER.utils.book_append_sheet(workbook, worksheet, 'Updated Stock');
  XLSX_WRITER.writeFile(workbook, filename);

  return { rowCount: updatedRows.length, filename };
}

/**
 * Generates an Excel Stock Report matching the user's requirements:
 * Main Category, Product Type, Product Name, Toon Label, Style No, Colour, Size, E-Com Stock, Wondersoft Stock, Diff
 * In-Sync (Difference 0) items highlighted in soft green!
 */
export function exportChangedStockReport(
  items: ComparisonItem[],
  _ignoredFormat?: any,
  filename?: string,
  options?: { onlyChanged?: boolean }
) {
  const onlyChanged = options?.onlyChanged ?? false;
  const filteredItems = onlyChanged
    ? items.filter(
        (i) => i.hasDecreased || i.isOutOfStock || i.difference !== 0 || i.isManuallyEdited
      )
    : items;

  const dateStr = new Date().toISOString().split('T')[0];
  const finalFilename =
    filename ||
    (onlyChanged
      ? `Changed_Stock_Report_${dateStr}.xlsx`
      : `Stock_Comparison_Report_${dateStr}.xlsx`);

  const reportRows = filteredItems.map((item) => ({
    'Main Category': item.mainCategory || '',
    'Product Name': item.name || '',
    'Toon Label': item.toonLabel || '',
    'Style No': item.styleNo || item.sku || '',
    'Colour': item.color || '',
    'Size': item.size || '',
    'E-Com Stock': item.ecommerceQty,
    'Wondersoft Stock': item.wondersoftQty,
    'Diff': Math.abs(item.difference),
  }));

  const worksheet = XLSX_WRITER.utils.json_to_sheet(reportRows);

  const headerCols = [
    'Main Category',
    'Product Name',
    'Toon Label',
    'Style No',
    'Colour',
    'Size',
    'E-Com Stock',
    'Wondersoft Stock',
    'Diff',
  ];

  // Style headers
  for (let c = 0; c < headerCols.length; c++) {
    const addr = XLSX_WRITER.utils.encode_cell({ r: 0, c });
    if (worksheet[addr]) {
      worksheet[addr].s = {
        fill: { fgColor: { rgb: '1E3A8A' } }, // Royal Blue
        font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: {
          bottom: { style: 'medium', color: { rgb: '0F172A' } },
        },
      };
    }
  }

  // Highlight Diff column (index 8) and Wondersoft Stock (index 7)
  const diffColIndex = 8;
  const wsStockColIndex = 7;

  for (let r = 0; r < filteredItems.length; r++) {
    const item = filteredItems[r];
    const diffCellAddr = XLSX_WRITER.utils.encode_cell({ r: r + 1, c: diffColIndex });
    const wsStockCellAddr = XLSX_WRITER.utils.encode_cell({ r: r + 1, c: wsStockColIndex });
    const isZeroDiff = item.difference === 0;
    const isOos = item.wondersoftQty <= 0;

    if (worksheet[diffCellAddr]) {
      if (isZeroDiff) {
        worksheet[diffCellAddr].s = {
          fill: { fgColor: { rgb: 'D1FAE5' } }, // Soft green
          font: { color: { rgb: '065F46' }, bold: true, sz: 11 },
          border: {
            top: { style: 'thin', color: { rgb: '10B981' } },
            bottom: { style: 'thin', color: { rgb: '10B981' } },
            left: { style: 'thin', color: { rgb: '10B981' } },
            right: { style: 'thin', color: { rgb: '10B981' } },
          },
          alignment: { horizontal: 'center' },
        };
      } else {
        worksheet[diffCellAddr].s = {
          fill: { fgColor: { rgb: isOos ? 'FEE2E2' : 'FEF08A' } },
          font: { color: { rgb: isOos ? 'DC2626' : '854D0E' }, bold: true, sz: 11 },
          border: {
            top: { style: 'thin', color: { rgb: isOos ? 'EF4444' : 'EAB308' } },
            bottom: { style: 'thin', color: { rgb: isOos ? 'EF4444' : 'EAB308' } },
            left: { style: 'thin', color: { rgb: isOos ? 'EF4444' : 'EAB308' } },
            right: { style: 'thin', color: { rgb: isOos ? 'EF4444' : 'EAB308' } },
          },
          alignment: { horizontal: 'center' },
        };
      }
    }

    if (worksheet[wsStockCellAddr] && isZeroDiff) {
      worksheet[wsStockCellAddr].s = {
        fill: { fgColor: { rgb: 'ECFDF5' } },
        font: { color: { rgb: '047857' }, bold: true, sz: 11 },
        alignment: { horizontal: 'center' },
      };
    }
  }

  worksheet['!cols'] = [
    { wch: 18 }, // Main Category
    { wch: 34 }, // Product Name
    { wch: 14 }, // Toon Label
    { wch: 16 }, // Style No
    { wch: 20 }, // Colour
    { wch: 20 }, // Size
    { wch: 14 }, // E-Com Stock
    { wch: 18 }, // Wondersoft Stock
    { wch: 12 }, // Diff
  ];

  const workbook = XLSX_WRITER.utils.book_new();
  XLSX_WRITER.utils.book_append_sheet(workbook, worksheet, 'Stock Comparison');
  XLSX_WRITER.writeFile(workbook, finalFilename);

  return { rowCount: filteredItems.length, filename: finalFilename };
}

/**
 * Aggregates comparison items by distinct Toon Label
 * Enables "calculate by Toon Label" grouped perspective
 */
export interface ToonLabelSummaryItem {
  toonLabel: string;
  itemCount: number;
  totalEcommerceQty: number;
  totalWondersoftQty: number;
  netDifference: number;
  outOfStockCount: number;
  decreasedCount: number;
  inSyncCount: number;
  styles: string[];
}

export function aggregateByToonLabel(items: ComparisonItem[]): ToonLabelSummaryItem[] {
  const map = new Map<string, ToonLabelSummaryItem>();

  for (const item of items) {
    const label = item.toonLabel ? item.toonLabel.trim() : '(Unlabeled)';
    const existing = map.get(label);

    const isOos = item.isOutOfStock;
    const isDec = item.hasDecreased;
    const isInSync = item.isInSync;

    if (existing) {
      existing.itemCount++;
      existing.totalEcommerceQty += item.ecommerceQty;
      existing.totalWondersoftQty += item.wondersoftQty;
      existing.netDifference += item.wondersoftQty - item.ecommerceQty;
      if (isOos) existing.outOfStockCount++;
      if (isDec) existing.decreasedCount++;
      if (isInSync) existing.inSyncCount++;
      if (item.styleNo && !existing.styles.includes(item.styleNo)) {
        existing.styles.push(item.styleNo);
      }
    } else {
      map.set(label, {
        toonLabel: label,
        itemCount: 1,
        totalEcommerceQty: item.ecommerceQty,
        totalWondersoftQty: item.wondersoftQty,
        netDifference: item.wondersoftQty - item.ecommerceQty,
        outOfStockCount: isOos ? 1 : 0,
        decreasedCount: isDec ? 1 : 0,
        inSyncCount: isInSync ? 1 : 0,
        styles: item.styleNo ? [item.styleNo] : [],
      });
    }
  }

  return Array.from(map.values()).sort((a, b) => b.totalEcommerceQty - a.totalEcommerceQty);
}
