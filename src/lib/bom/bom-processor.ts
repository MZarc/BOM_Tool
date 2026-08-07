/**
 * BOM Processor Module
 * Handles Excel parsing, designator expansion, normalization, and validation.
 * Optimized with Map/Set for O(n) performance on large BOMs.
 */

import * as XLSX from 'xlsx';

export interface BOMRow {
  mpn: string;
  qty: number;
  designators: string[];
  description?: string;
  duplicate?: boolean;
  duplicateDesignators?: Set<string>;
  qtyMismatch?: boolean;
  missing?: boolean;
  extra?: boolean;
  qtyMismatchResult?: boolean;
  swappedResult?: boolean;
  swappedDesignators?: Set<string>;
  comparisonQtyInfo?: { expected: number; actual: number };
  comparisonMPNInfo?: { expected: string; actual: string };
  rawDesignators?: string;
  modifiedFields?: Set<'mpn' | 'qty' | 'designators' | 'description'>;
  originalRowIdx: number;
  originalData?: {
    mpn: string;
    qty: number;
    designators: string[];
    description: string;
  };
}

export interface ColumnMapping {
  mpn: string | null;
  qty: string | null;
  designators: string | null;
  description?: string | null;
}

export interface ParseResult {
  rows: BOMRow[];
  rawData: Record<string, unknown>[]; // Objects with __rowIdx
  rawAOA: unknown[][]; // Original array of arrays
  headerRowIdx: number;
  columnMapping: ColumnMapping;
  needsManualMapping: boolean;
  availableColumns: string[];
  filename: string;
  merges: XLSX.Range[];
  sheetName?: string;
  buffer?: ArrayBuffer;
}

// All known aliases for each field type — used both for auto-mapping AND header-row detection scoring.
const MPN_ALIASES = [
  'manufacturer part number', 'manufacturerpartnumber', 'manufacturer_part_number',
  'part number', 'partnumber', 'part_number', 'part name', 'partname',
  'part number0', 'part number1', 'partnumber0', 'partnumber1',
  'mpn', 'mfr pn', 'mfr_pn', 'manufacturer pn', 'mfr part no',
  'item_number', 'part_order_code', 'order_code', 'component_pn',
  'part_num', 'part_no', 'mfr_part_num',
];
const QTY_ALIASES = [
  'qty', 'quantity', 'count', 'amount', 'total', 'total qty', 'total quantity',
  'qnty', 'q', 'qty/board', 'qty/ board', 'qty per board', 'required qty',
];
const REF_ALIASES = [
  'ref', 'reference', 'designator', 'designators', 'refdes', 'ref des', 'ref designator',
  'position', 'position no', 'position no.', 'components', 'references',
  'ref no', 'reference designator', 'reference designators', 'location', 'placement',
  'ref_des', 'ref_designator', 'position_no',
];
const DESC_ALIASES = [
  'description', 'desc', 'component description', 'part description', 'remarks', 'comment', 'product description',
];

// Combined set for header row scoring
const ALL_BOM_KEYWORDS = new Set([...MPN_ALIASES, ...QTY_ALIASES, ...REF_ALIASES, ...DESC_ALIASES]);

export class BOMProcessor {
  static get MPN_ALIASES() { return MPN_ALIASES; }
  static get QTY_ALIASES() { return QTY_ALIASES; }
  static get REF_ALIASES() { return REF_ALIASES; }

  /** Auto-detect column mapping from Excel headers */
  /** Auto-detect column mapping from Excel headers */
  static autoMapHeaders(headers: string[]): ColumnMapping {
    const mapping: ColumnMapping = { mpn: null, qty: null, designators: null, description: null };
    
    const findBest = (aliases: string[]) => {
      const lowerHeaders = headers.map(h => ({ original: h, lower: h.toLowerCase().trim() }));
      
      // Phase 1: Exact match
      for (const a of aliases) {
        const found = lowerHeaders.find(h => h.lower === a);
        if (found) return found.original;
      }
      
      // Phase 2: Start with
      for (const a of aliases) {
        const found = lowerHeaders.find(h => h.lower.startsWith(a));
        if (found) return found.original;
      }
      
      // Phase 3: Contains (only for longer aliases to avoid false positives)
      for (const a of aliases) {
        if (a.length < 3) continue;
        const found = lowerHeaders.find(h => h.lower.includes(a));
        if (found) return found.original;
      }
      
      return null;
    };

    mapping.mpn = findBest(MPN_ALIASES);
    mapping.qty = findBest(QTY_ALIASES);
    mapping.designators = findBest(REF_ALIASES);
    mapping.description = findBest(DESC_ALIASES);

    return mapping;
  }

  /** Score a row by how many cells look like BOM column header keywords */
  static scoreRowAsHeader(row: unknown[]): number {
    let score = 0;
    for (const cell of row) {
      if (typeof cell !== 'string' || cell.trim().length < 2) continue;
      const lower = cell.toLowerCase().trim();
      if (ALL_BOM_KEYWORDS.has(lower)) {
        score += 3; // exact match — very likely a header cell
      } else if ([...ALL_BOM_KEYWORDS].some(k => lower.includes(k) || k.includes(lower.split(' ')[0]))) {
        score += 1; // partial match
      } else if (cell.trim().length > 2) {
        score += 0.1; // any non-empty string adds a tiny bit
      }
    }
    return score;
  }

  /** Check if all three required columns are mapped */
  static isMappingComplete(mapping: ColumnMapping): boolean {
    return !!(mapping.mpn && mapping.qty && mapping.designators);
  }

  /** Find the best header row in a 2D array */
  static findHeaderRow(raw: unknown[][]): number {
    const SCAN_DEPTH = Math.min(raw.length, 25);
    let bestScore = 0;
    let headerRowIdx = -1;

    for (let i = 0; i < SCAN_DEPTH; i++) {
      const row = raw[i] ?? [];
      const score = this.scoreRowAsHeader(row);
      if (score > bestScore) {
        bestScore = score;
        headerRowIdx = i;
      }
    }

    if (headerRowIdx === -1 || bestScore < 1) {
      for (let i = 0; i < SCAN_DEPTH; i++) {
        const row = raw[i] ?? [];
        const nonEmpty = (row as unknown[]).filter((c: unknown) => typeof c === 'string' && (c as string).trim().length > 1);
        if (nonEmpty.length >= 3) return i;
      }
    }

    return headerRowIdx === -1 ? 0 : headerRowIdx;
  }

  /** Alias for autoMapHeaders */
  static guessMapping(headers: string[]): ColumnMapping {
    return this.autoMapHeaders(headers);
  }

  /** Alias for parseBOMWithMapping */
  static applyMapping(rawData: Record<string, unknown>[], mapping: ColumnMapping, headerRowIdx: number = 0): BOMRow[] {
    return this.parseBOMWithMapping(rawData, mapping, headerRowIdx);
  }

  /**
   * Normalize designator string:
   * - Replace spaces/semicolons with commas
   * - Remove duplicate commas, trim
   */
  static normalizeDesignatorString(input: string): string {
    if (!input) return '';
    return input
      .replace(/\s*;\s*/g, ',')
      .replace(/\n+/g, ',')
      // Smart space split: Only split spaces if preceded by a digit (e.g. "C5 C6" -> "C5, C6")
      // but NOT if preceded by a letter (e.g. "R C 3" stays "R C 3")
      .replace(/(\d)\s+/g, '$1,')
      // Clean up multiple commas, trim
      .replace(/,{2,}/g, ',')
      .replace(/^,|,$/g, '')
      .toUpperCase()
      .trim();
  }

  /**
   * Expand designator strings like "R1-R5", "R1...R5", "R1 to R5", "C001-C010".
   * Preserves leading zeros. Multi-letter prefixes (TR, SW) are treated correctly.
   */
  static expandDesignators(input: string): { expanded: string[]; isExpanded: boolean } {
    if (!input) return { expanded: [], isExpanded: false };
    // Split ONLY by comma to fulfill user request to prioritize commas and preserve spaces
    const parts = input.split(',').map(s => s.trim()).filter(Boolean);
    const expanded: string[] = [];
    let lastPrefix = '';

    for (const part of parts) {
      // 1. Check for ranges like R1-R5, R1 to R5, R C 1-R C 5, etc.
      // Updated regex to support spaces in prefixes (e.g., "R C")
      const rangeMatch = part.match(/^([A-Za-z][A-Za-z ]*?)(\d+)\s*(-|to|\.{1,}|[~\u2013\u2014])\s*(?:[A-Za-z][A-Za-z ]*?)?(\d+)$/i);
      if (rangeMatch) {
        const prefix = rangeMatch[1];
        lastPrefix = prefix;
        const startNumStr = rangeMatch[2];
        const endNumStr = rangeMatch[4];

        const padding = Math.max(startNumStr.length, endNumStr.length);
        const start = Math.min(parseInt(startNumStr, 10), parseInt(endNumStr, 10));
        const end = Math.max(parseInt(startNumStr, 10), parseInt(endNumStr, 10));
        for (let i = start; i <= end; i++) {
          expanded.push(`${prefix}${String(i).padStart(padding, '0')}`);
        }
      } 
      // 2. Check for naked numbers (prefix inheritance like R1, 2, 3)
      else if (/^\d+$/.test(part) && lastPrefix) {
        expanded.push(`${lastPrefix}${part}`);
      }
      // 3. Regular designator (store its prefix for later)
      else {
        const prefixMatch = part.match(/^([A-Za-z][A-Za-z ]*?)(\d+)$/i);
        if (prefixMatch) lastPrefix = prefixMatch[1];
        expanded.push(part);
      }
    }
    return { expanded, isExpanded: parts.length !== expanded.length };
  }

  /** Remove duplicate designators within a list */
  static removeDuplicatesInList(designators: string[]): { cleaned: string[]; duplicates: string[] } {
    const seen = new Set<string>();
    const cleaned: string[] = [];
    const duplicates: string[] = [];
    for (const d of designators) {
      if (seen.has(d)) duplicates.push(d);
      else { seen.add(d); cleaned.push(d); }
    }
    return { cleaned, duplicates };
  }

  /** Parse raw data rows using a given column mapping */
  static parseBOMWithMapping(rawData: Record<string, unknown>[], mapping: ColumnMapping, headerRowIdx: number = 0): BOMRow[] {
    const globalCounts = new Map<string, number>();
    
    // Pass 1: Minimal counting for duplicate detection (ONLY split by comma, no expansion)
    rawData.forEach(row => {
      const rawDesStr = mapping.designators ? String(row[mapping.designators] ?? '') : '';
      const parts = rawDesStr ? rawDesStr.split(',').map(s => s.trim()).filter(Boolean) : [];
      parts.forEach(d => {
        const k = d.trim().toUpperCase();
        if (k) globalCounts.set(k, (globalCounts.get(k) || 0) + 1);
      });
    });

    const result: BOMRow[] = [];
    rawData.forEach((row, idx) => {
      const mpn = mapping.mpn ? String(row[mapping.mpn] ?? '').trim() : '';
      const rawQtyVal = mapping.qty ? row[mapping.qty] : undefined;
      const qtyStr = String(rawQtyVal ?? '0');
      // Fix: parseFloat then floor to handle some scientific notation or string decimals
      const qty = Math.floor(parseFloat(qtyStr)) || 0;
      const description = mapping.description ? String(row[mapping.description] ?? '').trim() : '';
      const rawDesStr = mapping.designators ? String(row[mapping.designators] ?? '') : '';
      
      // Pass 2: Strictly comma-based designator list to satisfy "no automatic format" requirement
      const designators = rawDesStr ? rawDesStr.split(',').map(s => s.trim()).filter(Boolean) : [];

      const rowDups = new Set<string>();
      designators.forEach(d => {
        if (globalCounts.get(d.trim().toUpperCase())! > 1) {
          rowDups.add(d.trim().toUpperCase());
        }
      });

      const finalQty = qty > 0 ? qty : designators.length;

      result.push({
        mpn,
        qty: finalQty,
        designators: designators,
        description,
        rawDesignators: rawDesStr,
        duplicate: rowDups.size > 0,
        duplicateDesignators: rowDups,
        qtyMismatch: designators.length > 0 && designators.length !== finalQty,
        modifiedFields: new Set(), // No automatic formatting, so no modified fields initially
        originalRowIdx: (row.__rowIdx as number ?? (idx + headerRowIdx + 1)) + 1,
        originalData: {
          mpn,
          qty: finalQty,
          designators: [...designators],
          description: description || ''
        }
      });
    });
    return result;
  }

  /**
   * Parse directly from an XLSX WorkSheet, correctly handling sheets
   * that have title blocks, logos, or metadata before the real header row.
   *
   * Strategy:
   *  1. Read as raw 2D array (no header interpretation)
   *  2. Score each of the first 20 rows by how many cells match BOM keywords
   *  3. Pick the row with the highest score — that is the real header row
   *  4. Fall back to first row with ≥3 non-empty strings if no keyword match
   *
   * @param sheetToJson - pass `XLSX.utils.sheet_to_json` from the caller
   */
  static parseSheet(
    ws: XLSX.WorkSheet,
    sheetToJson: (ws: XLSX.WorkSheet, opts: XLSX.Sheet2JSONOpts) => unknown[][]
  ): ParseResult {
    const EMPTY_RESULT: ParseResult = {
      rows: [], rawData: [], rawAOA: [], headerRowIdx: 0,
      columnMapping: { mpn: null, qty: null, designators: null, description: null },
      needsManualMapping: false, availableColumns: [], filename: '', merges: []
    };

    const raw: unknown[][] = sheetToJson(ws, { header: 1, defval: '' });
    if (!raw.length) return EMPTY_RESULT;

    // --- Step 1: Find header row by scoring ---
    const SCAN_DEPTH = Math.min(raw.length, 25);
    let bestScore = 0;
    let headerRowIdx = -1;

    for (let i = 0; i < SCAN_DEPTH; i++) {
      const row = raw[i] ?? [];
      const score = this.scoreRowAsHeader(row);
      if (score > bestScore) {
        bestScore = score;
        headerRowIdx = i;
      }
    }

    // Fallback: first row with ≥3 non-empty string cells
    if (headerRowIdx === -1 || bestScore < 1) {
      for (let i = 0; i < SCAN_DEPTH; i++) {
        const row = raw[i] ?? [];
        const nonEmpty = (row as unknown[]).filter((c: unknown) => typeof c === 'string' && (c as string).trim().length > 1);
        if (nonEmpty.length >= 3) { headerRowIdx = i; break; }
      }
    }

    if (headerRowIdx === -1) headerRowIdx = 0;

    // --- Step 2: Build header array with positional placeholders for blank cells ---
    const headerRow = raw[headerRowIdx] as unknown[];
    const headers: string[] = headerRow.map((h: unknown, idx: number) => {
      const str = String(h ?? '').replace(/\r?\n/g, ' ').trim();
      return str.length > 0 ? str : `_COL_${idx}`;
    });

    // --- Step 3: Build data rows ---
    const dataRows: Record<string, unknown>[] = [];
    for (let i = headerRowIdx + 1; i < raw.length; i++) {
      const row = raw[i] ?? [] as unknown[];
      // Skip rows that are entirely empty
      const hasData = (row as unknown[]).some((c: unknown) => c !== '' && c !== null && c !== undefined);
      if (!hasData) continue;
      // Skip rows that look like footer/signature lines (all strings, no numbers at all)
      const numericCells = (row as unknown[]).filter((c: unknown) => typeof c === 'number');
      const stringCells = (row as unknown[]).filter((c: unknown) => typeof c === 'string' && (c as string).trim().length > 1);
      if (stringCells.length > 0 && numericCells.length === 0 && stringCells.length < 3) {
        // Likely a notes/signature row — skip
        continue;
      }
      const obj: Record<string, unknown> = { __rowIdx: i };
      headers.forEach((h, idx) => { obj[h] = (row as unknown[])[idx] ?? ''; });
      dataRows.push(obj);
    }

    const mapping = this.autoMapHeaders(headers);
    const rows = this.parseBOMWithMapping(dataRows, mapping, headerRowIdx);
    const needsManualMapping = !this.isMappingComplete(mapping);
    const merges = ws['!merges'] || [];

    return {
      rows,
      rawData: dataRows,
      rawAOA: raw,
      headerRowIdx,
      columnMapping: mapping,
      needsManualMapping,
      availableColumns: headers.filter(h => !h.startsWith('_COL_')),
      filename: '',
      merges,
    };
  }

  /** Re-apply processing to already-parsed rows. */
  static reprocess(rows: BOMRow[]): BOMRow[] {
    const counts = new Map<string, number>();
    rows.forEach(r => r.designators.forEach(d => {
      const k = d.trim().toUpperCase();
      counts.set(k, (counts.get(k) || 0) + 1);
    }));

    return rows.map(row => {
      const { expanded, isExpanded } = this.expandDesignators(row.designators.join(', '));
      const { cleaned } = this.removeDuplicatesInList(expanded);
      const rowDups = new Set<string>();
      cleaned.forEach(d => {
        if (counts.get(d.trim().toUpperCase())! > 1) rowDups.add(d.trim().toUpperCase());
      });
      
      const modifiedFields = new Set<'mpn' | 'qty' | 'designators' | 'description'>(row.modifiedFields || []);
      if (isExpanded) modifiedFields.add('designators');
      
      return { 
        ...row, 
        designators: cleaned, 
        duplicate: rowDups.size > 0, 
        duplicateDesignators: rowDups,
        qtyMismatch: cleaned.length > 0 && cleaned.length !== row.qty, 
        modifiedFields 
      };
    });
  }
  /** Find overlapping designators between two SEPARATE sets of rows */
  static findOverlappingDesignators(rowsA: BOMRow[], rowsB: BOMRow[]): { duplicates: string[] } {
    const setA = new Set(rowsA.flatMap(r => r.designators));
    const duplicates: string[] = [];
    for (const r of rowsB) {
      for (const d of r.designators) {
        if (setA.has(d)) duplicates.push(d);
      }
    }
    return { duplicates };
  }

  /** Find designators that appear more than once within a SINGLE set of rows */
  static findSelfDuplicates(rows: BOMRow[]): Set<string> {
    const counts = new Map<string, number>();
    for (const r of rows) {
      for (const d of r.designators) {
        if (!d) continue;
        const clean = d.trim().toUpperCase();
        counts.set(clean, (counts.get(clean) || 0) + 1);
      }
    }
    const dups = new Set<string>();
    for (const [d, count] of counts.entries()) {
      if (count > 1) dups.add(d);
    }
    return dups;
  }

  /** Recalculate duplicate designators across all rows after an edit */
  static recalculateDuplicates(rows: BOMRow[]): BOMRow[] {
    const globalCounts = new Map<string, number>();
    rows.forEach(r => {
      r.designators.forEach(d => {
        const k = d.trim().toUpperCase();
        if (k) globalCounts.set(k, (globalCounts.get(k) || 0) + 1);
      });
    });

    return rows.map(r => {
      const rowDups = new Set<string>();
      r.designators.forEach(d => {
        const k = d.trim().toUpperCase();
        if (k && (globalCounts.get(k) || 0) > 1) {
          rowDups.add(k);
        }
      });
      return {
        ...r,
        duplicate: rowDups.size > 0,
        duplicateDesignators: rowDups
      };
    });
  }
}
