/**
 * BOM Comparison Module
 * Efficiently compares two BOM datasets using Map/Set for O(n) performance.
 * Supports: missing parts, extra parts, qty mismatch, designator overlap, swapped components.
 */

import { BOMRow } from './bom-processor';

export interface DesignatorOverlap {
  mpn: string;
  intersection: string[];
  missingInTarget: string[];
  extraInTarget: string[];
}

export interface SwappedComponent {
  designator: string;
  sourceMPN: string;
  targetMPN: string;
  sourceRowIdx: number;
  targetRowIdx: number;
}

export interface BOMDiff {
  missing: BOMRow[];
  extra: BOMRow[];
  qtyMismatch: {
    mpn: string;
    expectedQty: number;
    actualQty: number;
    expectedDesignators: string[];
    actualDesignators: string[];
    sourceRowIndices: number[];
    targetRowIndices: number[];
  }[];
  designatorOverlap: DesignatorOverlap[];
  swapped: SwappedComponent[];
  duplicates: { source: BOMRow[]; target: BOMRow[] };
  internalQtyErrors: { source: BOMRow[]; target: BOMRow[] };
}

export interface ComparisonResults {
  source: BOMRow[];
  target: BOMRow[];
  diff: BOMDiff;
}

export class BOMCompare {
  /**
   * Full comparison of two BOMs using MPN as key.
   * Uses Map and Set for O(n) performance — no nested loops.
   */
  static compare(sourceBOM: BOMRow[], targetBOM: BOMRow[]): ComparisonResults {
    // Clone rows to avoid in-place mutation of the original state objects.
    // Sets are cloned too — otherwise `swappedDesignators` would accumulate
    // stale entries on the original rows across repeated comparisons.
    const cloneRow = (r: BOMRow): BOMRow => ({
      ...r,
      designators: [...r.designators],
      duplicateDesignators: new Set(r.duplicateDesignators ?? []),
      swappedDesignators: new Set(r.swappedDesignators ?? []),
      modifiedFields: new Set(r.modifiedFields ?? []),
    });
    const sourceCopy = sourceBOM.map(cloneRow);
    const targetCopy = targetBOM.map(cloneRow);

    const diff: BOMDiff = {
      missing: [],
      extra: [],
      qtyMismatch: [],
      designatorOverlap: [],
      swapped: [],
      duplicates: {
        source: sourceCopy.filter(r => r.duplicate),
        target: targetCopy.filter(r => r.duplicate)
      },
      internalQtyErrors: {
        source: sourceCopy.filter(r => r.qtyMismatch),
        target: targetCopy.filter(r => r.qtyMismatch)
      }
    };

    const normalize = (s: string) => String(s || '').trim().toUpperCase();

    // Reset all comparison flags and data first
    sourceCopy.forEach(r => { 
      r.missing = false; 
      r.swappedResult = false; 
      r.qtyMismatchResult = false; 
      r.comparisonQtyInfo = undefined;
      r.comparisonMPNInfo = undefined;
      r.swappedDesignators = new Set();
    });
    targetCopy.forEach(r => { 
      r.extra = false; 
      r.swappedResult = false; 
      r.qtyMismatchResult = false; 
      r.comparisonQtyInfo = undefined;
      r.comparisonMPNInfo = undefined;
      r.swappedDesignators = new Set();
    });

    // 1. Build lookup maps for fast O(n) access
    const sourceByMPN = new Map<string, BOMRow[]>();
    sourceCopy.forEach(r => {
      const k = normalize(r.mpn);
      if (!k) return; // skip empty MPN — blank cells are NOT the same part
      if (!sourceByMPN.has(k)) sourceByMPN.set(k, []);
      sourceByMPN.get(k)!.push(r);
    });

    const targetByMPN = new Map<string, BOMRow[]>();
    targetCopy.forEach(r => {
      const k = normalize(r.mpn);
      if (!k) return; // skip empty MPN
      if (!targetByMPN.has(k)) targetByMPN.set(k, []);
      targetByMPN.get(k)!.push(r);
    });

    const sourceByDes = new Map<string, { mpn: string; rows: BOMRow[] }>();
    sourceCopy.forEach(row => {
      row.designators.forEach(d => {
        const k = normalize(d);
        if (!sourceByDes.has(k)) sourceByDes.set(k, { mpn: row.mpn, rows: [] });
        sourceByDes.get(k)!.rows.push(row);
      });
    });

    const targetByDes = new Map<string, { mpn: string; rows: BOMRow[] }>();
    targetCopy.forEach(row => {
      row.designators.forEach(d => {
        const k = normalize(d);
        if (!targetByDes.has(k)) targetByDes.set(k, { mpn: row.mpn, rows: [] });
        targetByDes.get(k)!.rows.push(row);
      });
    });

    // 2. Identify Swaps (Designator-first)
    const allDesignators = new Set([...sourceByDes.keys(), ...targetByDes.keys()]);
    allDesignators.forEach(des => {
      const s = sourceByDes.get(des);
      const t = targetByDes.get(des);

      if (s && t && normalize(s.mpn) && normalize(t.mpn) && normalize(s.mpn) !== normalize(t.mpn)) {
        // Find the specific rows containing this designator
        const sourceRow = s.rows.find(r => r.designators.some(d => normalize(d) === des))!;
        const targetRow = t.rows.find(r => r.designators.some(d => normalize(d) === des))!;

        diff.swapped.push({
          designator: des,
          sourceMPN: s.mpn,
          targetMPN: t.mpn,
          sourceRowIdx: sourceRow.originalRowIdx,
          targetRowIdx: targetRow.originalRowIdx
        });
        
        s.rows.forEach(r => {
          r.swappedResult = true;
          r.comparisonMPNInfo = { expected: s.mpn, actual: t.mpn };
          if (!r.swappedDesignators) r.swappedDesignators = new Set();
          r.swappedDesignators.add(des.toUpperCase());
        });

        t.rows.forEach(r => {
          r.swappedResult = true;
          r.comparisonMPNInfo = { expected: s.mpn, actual: t.mpn };
          if (!r.swappedDesignators) r.swappedDesignators = new Set();
          r.swappedDesignators.add(des.toUpperCase());
        });
      }
    });

    // 3. MPN-based Comparison
    sourceByMPN.forEach((sRows, mpn) => {
      const tRows = targetByMPN.get(mpn);
      // If MPN not in target, it's missing from target
      if (!tRows) {
        sRows.forEach(r => {
          r.missing = true;
          diff.missing.push(r);
        });
      } else {
        // MPN exists, check total quantities
        const sTotal = sRows.reduce((sum, r) => sum + r.qty, 0);
        const tTotal = tRows.reduce((sum, r) => sum + r.qty, 0);

        if (sTotal !== tTotal) {
          sRows.forEach(r => {
            r.qtyMismatchResult = true;
            r.comparisonQtyInfo = { expected: sTotal, actual: tTotal };
          });
          tRows.forEach(r => {
            r.qtyMismatchResult = true;
            r.comparisonQtyInfo = { expected: sTotal, actual: tTotal };
          });
          diff.qtyMismatch.push({
            mpn,
            expectedQty: sTotal,
            actualQty: tTotal,
            expectedDesignators: sRows.flatMap(r => r.designators),
            actualDesignators: tRows.flatMap(r => r.designators),
            sourceRowIndices: sRows.map(r => r.originalRowIdx),
            targetRowIndices: tRows.map(r => r.originalRowIdx)
          });
        }

        // Designator overlap for parts that exist on both sides
        const sDes = new Set(sRows.flatMap(r => r.designators.map(normalize)));
        const tDes = new Set(tRows.flatMap(r => r.designators.map(normalize)));
        const intersection = [...sDes].filter(d => tDes.has(d));
        const missingInTarget = [...sDes].filter(d => !tDes.has(d));
        const extraInTarget = [...tDes].filter(d => !sDes.has(d));
        if (missingInTarget.length || extraInTarget.length) {
          diff.designatorOverlap.push({
            mpn,
            intersection,
            missingInTarget,
            extraInTarget,
          });
        }
      }
    });

    // Extra tagging: in target but NOT in source
    targetByMPN.forEach((tRows, mpn) => {
      if (!sourceByMPN.has(mpn)) {
        tRows.forEach(r => {
          r.extra = true;
          diff.extra.push(r);
        });
      }
    });

    // Rows without an MPN cannot be matched by part number. Surface them as
    // missing/extra instead of silently dropping them from the comparison.
    sourceCopy.forEach(r => {
      if (!normalize(r.mpn) && !r.missing) {
        r.missing = true;
        diff.missing.push(r);
      }
    });
    targetCopy.forEach(r => {
      if (!normalize(r.mpn) && !r.extra) {
        r.extra = true;
        diff.extra.push(r);
      }
    });

    return { source: sourceCopy, target: targetCopy, diff };
  }
}
