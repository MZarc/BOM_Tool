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
// Precomputed array — avoids re-spreading the Set for every cell while scoring.
const ALL_BOM_KEYWORDS_ARRAY = [...ALL_BOM_KEYWORDS];

// Headers that must never be treated as the quantity column (e.g. "Total Price").
const QTY_BLOCKED = /price|cost|rate|usd|\$/;

export class BOMProcessor {
  static get MPN_ALIASES() { return MPN_ALIASES; }
  static get QTY_ALIASES() { return QTY_ALIASES; }
  static get REF_ALIASES() { return REF_ALIASES; }

  /** Auto-detect column mapping from Excel headers */
  static autoMapHeaders(headers: string[]): ColumnMapping {
    const mapping: ColumnMapping = { mpn: null, qty: null, designators: null, description: null };

    const lowerHeaders = headers.map((h, index) => ({
      original: h,
      lower: String(h ?? '').toLowerCase().trim(),
      index,
    }));
    // Track consumed columns so the same header can never fill two roles.
    const used = new Set<number>();

    const findBest = (aliases: string[], blocked?: RegExp): string | null => {
      const candidates = lowerHeaders.filter(
        h => h.lower.length > 0 && !used.has(h.index) && (!blocked || !blocked.test(h.lower))
      );

      // Phase 1: Exact match
      for (const a of aliases) {
        const found = candidates.find(h => h.lower === a);
        if (found) { used.add(found.index); return found.original; }
      }

      // Phase 2: Starts with (single-character aliases are too greedy)
      for (const a of aliases) {
        if (a.length < 2) continue;
        const found = candidates.find(h => h.lower.startsWith(a));
        if (found) { used.add(found.index); return found.original; }
      }

      // Phase 3: Contains (only for longer aliases to avoid false positives)
      for (const a of aliases) {
        if (a.length < 3) continue;
        const found = candidates.find(h => h.lower.includes(a));
        if (found) { used.add(found.index); return found.original; }
      }

      return null;
    };

    mapping.mpn = findBest(MPN_ALIASES);
    mapping.qty = findBest(QTY_ALIASES, QTY_BLOCKED);
    mapping.designators = findBest(REF_ALIASES);
    mapping.description = findBest(DESC_ALIASES);

    return mapping;
  }

  /** Score a row by how many cells look like BOM column header keywords */
  static scoreRowAsHeader(row: unknown[]): number {
    let score = 0;
    for (const cell of row) {
      if (typeof cell !== 'string') continue;
      const lower = cell.toLowerCase().trim();
      if (lower.length < 2) continue;
      if (ALL_BOM_KEYWORDS.has(lower)) {
        score += 3; // exact match — very likely a header cell
      } else {
        const firstWord = lower.split(/\s+/)[0];
        const partial = ALL_BOM_KEYWORDS_ARRAY.some(
          k => lower.includes(k) || (firstWord.length >= 3 && k.startsWith(firstWord))
        );
        score += partial ? 1 : 0.1; // partial / any other non-empty string
      }
    }
    return score;
  }

  /**
   * Single source of truth for "row quantity does not match its designators".
   * A positive quantity with zero designators is a mismatch too (you cannot
   * place parts that have no reference). Every mode uses this exact rule.
   */
  static isQtyMismatch(qty: number, designators: string[]): boolean {
    return qty !== designators.length;
  }

  /** Check if all three required columns are mapped */
  static isMappingComplete(mapping: ColumnMapping): boolean {
    return !!(mapping.mpn && mapping.qty && mapping.designators);
  }

  /** Build a stable, duplicate-safe header list (matches XLSX's `_1` suffix convention). */
  static buildHeaders(headerRow: unknown[]): string[] {
    const seen = new Map<string, number>();
    return headerRow.map((h, idx) => {
      let name = String(h ?? '').replace(/\r?\n/g, ' ').trim();
      if (!name) name = `_COL_${idx}`;
      const count = seen.get(name) ?? 0;
      seen.set(name, count + 1);
      return count === 0 ? name : `${name}_${count}`;
    });
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

    // Only trust a row that actually matched BOM keywords. Weak noise scores
    // (0.1 per generic string) must not beat the positional fallback.
    if (bestScore >= 1 && headerRowIdx !== -1) return headerRowIdx;

    for (let i = 0; i < SCAN_DEPTH; i++) {
      const row = raw[i] ?? [];
      const nonEmpty = (row as unknown[]).filter((c: unknown) => typeof c === 'string' && (c as string).trim().length > 1);
      if (nonEmpty.length >= 3) return i;
    }

    return 0;
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
      // 0. Decimal values like "R1.5" are single designators, not ranges
      if (/^[A-Za-z][A-Za-z ]*\d+\.\d+$/.test(part)) {
        const prefixMatch = part.match(/^([A-Za-z][A-Za-z ]*?)(\d+)$/i);
        if (prefixMatch) lastPrefix = prefixMatch[1];
        expanded.push(part);
        continue;
      }
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

  /** Parse a quantity cell, tolerating "1,000" and "1 pcs" style values. */
  private static parseQtyValue(raw: unknown): number {
    if (typeof raw === 'number') return Math.floor(raw) || 0;
    let str = String(raw ?? '').trim();
    if (!str) return 0;
    // Strip thousands separators only when they follow the 3-digit grouping rule
    // (so European decimals like "1,5" are left alone).
    if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(str)) str = str.replace(/,/g, '');
    const n = parseFloat(str);
    return Number.isFinite(n) ? Math.floor(n) : 0;
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
      const qty = mapping.qty ? this.parseQtyValue(row[mapping.qty]) : 0;
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
        qtyMismatch: this.isQtyMismatch(finalQty, designators),
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

    // --- Step 1: Find header row by scoring (shared with the PnP flow) ---
    const headerRowIdx = this.findHeaderRow(raw);

    // --- Step 2: Build header array with positional placeholders for blank cells ---
    const headers = this.buildHeaders(raw[headerRowIdx] ?? []);

    // --- Step 3: Build data rows ---
    // Only fully empty rows are skipped. Sparse data rows (e.g. a text qty and
    // no description) are still valid BOM rows and must not be dropped.
    const dataRows: Record<string, unknown>[] = [];
    for (let i = headerRowIdx + 1; i < raw.length; i++) {
      const row = raw[i] ?? [] as unknown[];
      const hasData = (row as unknown[]).some((c: unknown) => c !== '' && c !== null && c !== undefined);
      if (!hasData) continue;
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
        qtyMismatch: this.isQtyMismatch(row.qty, cleaned), 
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
