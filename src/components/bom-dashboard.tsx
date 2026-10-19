import React, { useState, useCallback, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { BOMProcessor, BOMRow, ColumnMapping, ParseResult } from '@/lib/bom/bom-processor';
import { BOMCompare, BOMDiff } from '@/lib/bom/bom-compare';
import { PnPRow } from '@/lib/bom/pnp-parser';
import {
  Upload, ArrowRightLeft, Database, Save, RefreshCw,
  Trash2, Columns, CheckCircle2, X, ChevronDown, ChevronRight, Undo2, Redo2,
  GitCompareArrows, AlignJustify, ShieldCheck, Download,
  AlertCircle, Info, ExternalLink, MapPin, Minus, Search, Bug
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { PnPValidator, PnPFileData } from './pnp-validator';
import { PendingMapping, ColumnMapDialog } from './ui/MappingDialogs';
import { FilterDropdown } from './ui/FilterDropdown';
import { ModeSearch, PnPSearchRow } from './ui/UniversalSearch';

export interface SheetPickerState {
  sheetNames: string[];
  workbookBinary: string;
  filename: string;
  onSheetSelected: (name: string) => void;
}

// ─── Types ────────────────────────────────────────────────────────────────────
/** Fixed row height (px). Must stay in sync with the row markup so the
 *  virtualized window math and the spacer rows line up exactly. */
const ROW_HEIGHT = 64;

type Tab = 'analyzer' | 'comparator' | 'pnp';
type ToolId = 'normalize' | 'expand' | 'validate' | 'duplicates' | 'fix_mismatch' | 'fix_duplicate';


interface FileState {
  rows: BOMRow[];
  rawData: Record<string, unknown>[];
  rawAOA: unknown[][];
  headerRowIdx: number;
  columnMapping: ColumnMapping;
  availableColumns: string[];
  filename: string;
  merges: XLSX.Range[];
  buffer: ArrayBuffer;
  sheetName?: string;
}

interface StatusModalState {
  type: 'success' | 'error' | 'info';
  title: string;
  message: string;
}

const emptyFile = (): FileState => ({
  rows: [], rawData: [], rawAOA: [], headerRowIdx: 0,
  columnMapping: { mpn: null, qty: null, designators: null, description: null },
  availableColumns: [], filename: '', merges: [],
  buffer: new ArrayBuffer(0)
});

// ─── Status Modal ────────────────────────────────────────────────────────────
function StatusModal({
  type, title, message, onClose
}: {
  type: 'success' | 'error' | 'info';
  title: string;
  message: string;
  onClose: () => void;
}) {
  const icons = {
    success: <CheckCircle2 className="h-8 w-8 text-emerald-400" />,
    error: <AlertCircle className="h-8 w-8 text-red-400" />,
    info: <Info className="h-8 w-8 text-blue-400" />,
  };

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
      <div className="absolute inset-0 bg-black/60" />
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="relative z-10 glass-card p-8 w-full max-w-lg text-center space-y-4"
        style={{ background: 'rgba(10,18,36,0.98)', border: '1px solid rgba(255,255,255,0.1)' }}
      >
        <div className="flex justify-center">{icons[type]}</div>
        <div className="space-y-1">
          <h3 className="text-xl font-bold text-white tracking-tight">{title}</h3>
          <p className="text-sm text-white/50 leading-relaxed break-words px-4">{message}</p>
        </div>
        <button
          onClick={onClose}
          className="w-full glass-button py-2.5 text-sm font-bold mt-2"
        >
          Got it
        </button>
      </motion.div>
    </div>
  );
}



// ─── Shared UI Helpers Moved ──────────────────────────────────────────────────


// ─── Tool Bar ─────────────────────────────────────────────────────────────────
function ToolBar({
  onRunTool, onRemap, filename,
  onUndo, onRedo, canUndo, canRedo,
  onSearch
}: {
  onRunTool: (tool: ToolId) => void;
  onRemap: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  filename: string;
  onSearch: () => void;
}) {
  const tools: { id: ToolId; label: string; icon: React.ReactNode; action: () => void; color: string; title: string; description: string }[] = [
    { id: 'normalize', label: 'Normalize', icon: <AlignJustify className="h-3.5 w-3.5" />, action: () => onRunTool('normalize'), color: 'hover:bg-blue-500/20 hover:border-blue-400/30 hover:text-blue-300', title: 'Normalize Designators', description: 'Clean designators: replace spaces with commas and fix formatting' },
    { id: 'expand', label: 'Expand', icon: <ArrowRightLeft className="h-3.5 w-3.5" />, action: () => onRunTool('expand'), color: 'hover:bg-amber-500/20 hover:border-amber-400/30 hover:text-amber-300', title: 'Expand Ranges', description: 'Expand ranges: convert R1-R5 into R1,R2,R3,R4,R5' },
    { id: 'validate', label: 'Validate', icon: <ShieldCheck className="h-4 w-4" />, action: () => onRunTool('validate'), color: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30 hover:bg-emerald-500/25 hover:border-emerald-400/60 shadow-[0_0_15px_rgba(16,185,129,0.1)]', title: 'Full Validation', description: 'Final Validation: check for duplicates and qty mismatches' },
  ];

  return (
    <div className="w-full flex items-center justify-between gap-4">
      {/* Left group: file info + tools */}
      <div className="flex items-center gap-2">
        <div className="flex items-center gap-2 pr-4 border-r border-white/10" title={filename || 'bom.xlsx'}>
          <div className="p-2 rounded-xl bg-white/5 border border-white/10">
            <Database className="h-4 w-4 text-white/40" />
          </div>
          <div className="flex flex-col">
            <span className="text-[10px] font-black uppercase tracking-widest text-white/20 leading-none mb-1">Active File</span>
            <span className="text-xs text-white/60 font-mono truncate max-w-[150px] leading-none">{filename || 'bom.xlsx'}</span>
          </div>
        </div>
        <div className="flex items-center gap-2 ml-2">
          {tools.map(t => (
            <button
              key={t.id}
              onClick={t.action}
              title={t.description}
              className={`flex shrink-0 items-center gap-2 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest whitespace-nowrap transition-all duration-300 cursor-pointer border ${t.id === 'validate' ? t.color : 'text-white/40 border-white/5 hover:bg-white/5 hover:text-white hover:border-white/20'} ${t.color}`}
            >
              {t.icon}
              {t.label}
            </button>
          ))}
        </div>
        <button
          onClick={onSearch}
          title="Search within current database mode"
          className="flex shrink-0 items-center gap-2 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest whitespace-nowrap transition-all duration-300 cursor-pointer border border-primary/20 bg-primary/5 text-primary hover:bg-primary/10 hover:border-primary/40 shadow-sm ml-4"
        >
          <Search className="h-3.5 w-3.5" />
          Search
        </button>
      </div>

      <div className="flex-1" />

      {/* Right group: actions */}
      <div className="flex items-center gap-1.5">
        <button
          onClick={onRemap}
          title="Manually select columns for MPN, Qty, and Designators"
          className="flex shrink-0 items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-white/50 whitespace-nowrap hover:text-white bg-white/5 hover:bg-white/10 transition-all duration-200 cursor-pointer"
        >
          <Columns className="h-3.5 w-3.5" />
          Remap
        </button>
        <div className="w-[1px] h-4 bg-white/10 mx-1" />
        <button
          onClick={onUndo}
          disabled={!canUndo}
          title="Undo last change (Ctrl+Z)"
          className="p-1.5 rounded-lg text-white/40 hover:text-white hover:bg-white/10 disabled:opacity-20 disabled:cursor-not-allowed transition-all cursor-pointer"
        >
          <Undo2 className="h-4 w-4" />
        </button>
        <button
          onClick={onRedo}
          disabled={!canRedo}
          title="Redo next change (Ctrl+Y)"
          className="p-1.5 rounded-lg text-white/40 hover:text-white hover:bg-white/10 disabled:opacity-20 disabled:cursor-not-allowed transition-all cursor-pointer"
        >
          <Redo2 className="h-4 w-4" />
        </button>
      </div>
    </div>
  );
}

// --- Double Click Input Component ---
function DoubleClickInput({
  value,
  onUpdate,
  ariaLabel,
  title,
  className,
  type = 'text',
  placeholder = '—',
  isNumber = false
}: {
  value: string | number;
  onUpdate: (val: string | number) => void;
  ariaLabel?: string;
  title?: string;
  className?: string;
  type?: string;
  placeholder?: string;
  isNumber?: boolean;
}) {
  const [editing, setEditing] = useState(false);
  const [localVal, setLocalVal] = useState(String(value ?? ''));

  React.useEffect(() => {
    if (!editing) setLocalVal(String(value ?? ''));
  }, [value, editing]);

  if (!editing) {
    return (
      <div
        onDoubleClick={() => setEditing(true)}
        title={title || String(value)}
        className={`w-full h-full px-2 flex items-center justify-center select-none cursor-pointer text-center ${className}`}
      >
        <span className="min-w-0 break-words leading-snug line-clamp-2">{String(value) || placeholder}</span>
      </div>
    );
  }

  return (
    <input
      type={type}
      aria-label={ariaLabel}
      autoFocus
      title={title}
      className={`w-full h-full bg-white/10 px-2 py-1 text-center focus:outline-none transition-all ${className}`}
      value={localVal}
      onChange={e => setLocalVal(e.target.value)}
      onBlur={() => {
        setEditing(false);
        if (isNumber) {
          onUpdate(parseInt(localVal) || 0);
        } else {
          onUpdate(localVal);
        }
      }}
      onKeyDown={e => {
        if (e.key === 'Enter') {
          setEditing(false);
          if (isNumber) {
            onUpdate(parseInt(localVal) || 0);
          } else {
            onUpdate(localVal);
          }
        }
      }}
      placeholder={placeholder}
    />
  );
}

// --- Designator Input Sub-component ---
function DesignatorInput({ 
  value, 
  onUpdate, 
  isModified, 
  rowIdx,
  row
}: { 
  value: string[]; 
  onUpdate: (val: string[]) => void; 
  isModified?: boolean;
  rowIdx: number;
  row: BOMRow;
}) {
  const [local, setLocal] = useState(value.join(', '));
  const [isFocused, setIsFocused] = useState(false);

  React.useEffect(() => {
    if (!isFocused) setLocal(value.join(', '));
  }, [value, isFocused]);

  if (!isFocused) {
    return (
      <div 
        onDoubleClick={() => setIsFocused(true)}
        title={value.join(', ')}
        className={`w-full h-full flex items-center justify-center px-2 py-1 text-[11px] font-semibold font-mono group-hover:text-white rounded-lg transition-all text-center [text-shadow:0_1px_2px_rgba(0,0,0,0.5)] cursor-text overflow-hidden ${
          isModified ? 'bg-indigo-500/10 ring-1 ring-inset ring-indigo-500/40 shadow-[0_0_10px_rgba(129,140,248,0.2)]' : ''
        }`}
      >
        <span className="w-full min-w-0 line-clamp-2 leading-snug break-words">
          {value.map((d, i) => {
            const clean = d.trim().toUpperCase();
            const isDup = row.duplicateDesignators?.has(clean);

            return (
              <span key={i}>
                <span className={isDup ? 'text-orange-400 font-black underline underline-offset-4 decoration-orange-500/50' : 'text-white/80'}>
                  {d}
                </span>
                {i < value.length - 1 && <span className="opacity-30">, </span>}
              </span>
            );
          })}
          {value.length === 0 && <span className="opacity-20">—</span>}
        </span>
      </div>
    );
  }

  return (
    <input
      aria-label={`Row ${rowIdx + 1} Designators`}
      autoFocus
      onBlur={() => {
        setIsFocused(false);
        onUpdate(local.split(',').map(s => s.trim()).filter(Boolean));
      }}
      className="w-full h-full bg-white/10 px-2 py-1 text-[11px] font-black font-mono text-white focus:outline-none rounded-lg transition-all text-center [text-shadow:0_1px_2px_rgba(0,0,0,0.5)]"
      value={local}
      onChange={e => setLocal(e.target.value)}
      placeholder="—"
    />
  );
}

// ─── Editable BOM Grid ────────────────────────────────────────────────────────
function EditableBOMGrid({
  data, onUpdate, type = 'analyzer', highlightLogic, label, onRemap, onSyncFromMaster,
  onUndo, onRedo, canUndo, canRedo,
  filters = {}, onFilterChange, jumpTarget,
  initialScroll, onScrollChange
}: {
  data: BOMRow[];
  onUpdate: (rowIdx: number, f: keyof BOMRow, v: BOMRow[keyof BOMRow]) => void;
  type?: 'analyzer' | 'ref' | 'target';
  highlightLogic?: (row: BOMRow) => string;
  label?: string;
  onRemap?: () => void;
  onSyncFromMaster?: (row: BOMRow, i: number) => void;
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  filters?: Record<string, Set<string>>;
  onFilterChange?: (column: string, selected: Set<string>) => void;
  jumpTarget?: { rowIdx: number; timestamp: number } | null;
  initialScroll?: number;
  onScrollChange?: (top: number) => void;
}) {
  const tableRef = useRef<HTMLDivElement>(null);
  const [openFilter, setOpenFilter] = useState<string | null>(null);

  const statusBadge = useCallback((row: BOMRow) => {
    const problems: string[] = [];
    let priorityColor = 'bg-emerald-400/50';
    let labelText = 'OK';

    // 1. Comparison Results (High Priority - Context Aware)
    if (row.missing) {
      problems.push('Part exists in Reference but is MISSING in Target BOM');
      priorityColor = 'bg-pink-500 shadow-[0_0_10px_rgba(236,72,153,0.5)]';
      labelText = 'Missing';
    }
    if (row.extra) {
      problems.push('Extra component (found in Target but NOT in Reference)');
      priorityColor = 'bg-teal-400 shadow-[0_0_10px_rgba(45,212,191,0.5)]';
      labelText = 'Extra';
    }
    if (row.swappedResult) {
      const { expected, actual } = row.comparisonMPNInfo || {};
      if (type === 'ref') {
        problems.push(`Swapped: Target uses different MPN (${actual || 'Unknown'})`);
      } else {
        problems.push(`Swapped: Reference uses different MPN (${expected || 'Unknown'})`);
      }
      priorityColor = 'bg-indigo-400 shadow-[0_0_10px_rgba(129,140,248,0.5)]';
      labelText = 'Swapped';
    }
    if (row.qtyMismatchResult) {
      const { expected, actual } = row.comparisonQtyInfo || {};
      if (type === 'ref') {
        problems.push(`Qty Mismatch: Ref has ${expected}, Target has ${actual}`);
      } else {
        problems.push(`Qty Mismatch: Target has ${actual}, Reference has ${expected}`);
      }
      priorityColor = 'bg-sky-400 shadow-[0_0_10px_rgba(56,189,248,0.5)]';
      labelText = 'Qty Delta';
    }

    // 2. Internal Consistency (Overrides comparison if issues found)
    if (row.duplicate) {
      problems.push('Duplicate designators found in this file');
      priorityColor = 'bg-orange-500 shadow-[0_0_10px_rgba(249,115,22,0.5)]';
      labelText = 'Duplicate';
    }
    if (row.qtyMismatch) {
      problems.push('Internal Qty mismatch (Row Qty ≠ Designator count)');
      priorityColor = 'bg-amber-400 shadow-[0_0_10px_rgba(251,191,36,0.5)]';
      labelText = 'Qty Error';
    }
    
    // 3. Critical Override (Both internal issues)
    if (row.duplicate && row.qtyMismatch) {
      priorityColor = 'bg-rose-500 shadow-[0_0_11px_rgba(244,63,94,0.5)]';
      labelText = 'Critical';
    }

    return { 
      dot: priorityColor, 
      title: problems.length > 0 ? problems.join('\n• ') : 'No issues found',
      label: labelText,
      hasIssues: true,
      problems: problems.length > 0 ? problems : 
        type === 'analyzer' ? ['Component is internally consistent'] : ['Component matches correctly with the other BOM']
    };
  }, [type]);

  // Compute unique values for each column based on the FULL data
  const columnUniqueValues = useMemo(() => {
    const values: Record<string, string[]> = {
      mpn: [],
      qty: [],
      description: [],
      designators: [],
      status: []
    };
    
    data.forEach(row => {
      values.mpn.push(row.mpn || '');
      values.qty.push(String(row.qty || 0));
      values.description.push(row.description || '');
      values.designators.push(row.designators.join(', '));
      values.status.push(statusBadge(row).label);
    });

    return values;
  }, [data, statusBadge]);

  // Apply filters to data
  const filteredData = useMemo(() => {
    if (!filters || Object.keys(filters).length === 0) return data;
    
    return data.filter(row => {
      // Check each filter
      for (const [col, selected] of Object.entries(filters)) {
        if (!selected || selected.size === 0) continue;
        
        let val = '';
        if (col === 'mpn') val = row.mpn || '';
        else if (col === 'qty') val = String(row.qty || 0);
        else if (col === 'description') val = row.description || '';
        else if (col === 'designators') val = row.designators.join(', ');
        else if (col === 'status') val = statusBadge(row).label;
        
        if (!selected.has(val)) return false;
      }
      return true;
    });
  }, [data, filters, statusBadge]);

  // Handle Jumping and Highlighting
  const [highlightedRow, setHighlightedRow] = useState<number | null>(null);

  useEffect(() => {
    if (!jumpTarget || !tableRef.current) return;

    const ROW_HEIGHT = 64;
    const container = tableRef.current;
    
    const targetIdx = filteredData.findIndex(r => r.originalRowIdx === jumpTarget.rowIdx);
    const idxToScroll = targetIdx !== -1 ? targetIdx : Math.max(0, jumpTarget.rowIdx - 1);
    const targetScrollTop = Math.max(0, (idxToScroll * ROW_HEIGHT) - (container.clientHeight / 3));
    
    // Synchronously jump DOM container and React scrollTop state together
    container.scrollTop = targetScrollTop;
    setScrollTop(targetScrollTop);
    setHighlightedRow(jumpTarget.rowIdx);

    // Micro-align on next frame if row element is in DOM
    const raf = requestAnimationFrame(() => {
      const rowElem = container.querySelector<HTMLElement>(`#row-${jumpTarget.rowIdx}`);
      if (rowElem) {
        rowElem.scrollIntoView({ block: 'nearest' });
      }
    });

    const timer = setTimeout(() => setHighlightedRow(null), 5000);
    return () => {
      cancelAnimationFrame(raf);
      clearTimeout(timer);
    };
  }, [jumpTarget, filteredData]);

  const [scrollTop, setScrollTop] = useState(0);

  const hasRestored = useRef(false);
  const lastReportedScroll = useRef<number>(0);
  const scrollTimeout = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Restore scroll position ONCE on mount
  useEffect(() => {
    if (tableRef.current && initialScroll !== undefined && !hasRestored.current) {
      tableRef.current.scrollTop = initialScroll;
      setScrollTop(initialScroll);
      lastReportedScroll.current = initialScroll;
      hasRestored.current = true;
    }
  }, [initialScroll]); // Dependency added for lint, Ref ensures 'once' logic

  // Throttled scroll tracking & stepped windowing to prevent scroll recoil
  const handleScroll = useCallback(() => {
    if (tableRef.current) {
      const top = tableRef.current.scrollTop;
      // Only update state when scroll shifts by > 150px to eliminate per-pixel DOM shifts & bounce-back
      setScrollTop(prev => Math.abs(top - prev) > 150 ? top : prev);
      
      // If the change is significant or it's the end of a scroll, report it
      if (Math.abs(top - lastReportedScroll.current) > 50) {
        lastReportedScroll.current = top;
        onScrollChange?.(top);
      } else {
        // Fallback: report eventual position after scroll stops
        if (scrollTimeout.current) clearTimeout(scrollTimeout.current);
        scrollTimeout.current = setTimeout(() => {
          if (tableRef.current) onScrollChange?.(tableRef.current.scrollTop);
        }, 150);
      }
    }
  }, [onScrollChange]);

  const visibleWindow = useMemo(() => {
    if (filteredData.length <= 50) {
      return { startIndex: 0, endIndex: filteredData.length, topPadding: 0, bottomPadding: 0 };
    }
    const ROW_HEIGHT = 64;
    const OVERSCAN = 25; // Generous 1250px buffer prevents scroll edge clipping
    const startIndex = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN);
    const endIndex = Math.min(filteredData.length, Math.ceil((scrollTop + 750) / ROW_HEIGHT) + OVERSCAN);
    const topPadding = startIndex * ROW_HEIGHT;
    const bottomPadding = (filteredData.length - endIndex) * ROW_HEIGHT;
    return { startIndex, endIndex, topPadding, bottomPadding };
  }, [filteredData.length, scrollTop]);

  return (
    <div className="glass-card overflow-hidden border-white/5 h-full flex flex-col shadow-2xl">
      {label && (
        <div className="px-4 py-3 border-b border-white/5 shrink-0 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3 min-w-0 flex-1">
            <span 
              className="text-[10px] font-black uppercase tracking-[0.2em] text-white/30 truncate"
              title={label}
            >
              {label}
            </span>
            {onUndo && (
              <div className="flex items-center gap-1 bg-white/10 border border-white/20 rounded-lg p-0.5 shrink-0 scale-90 origin-left">
                <button 
                  onClick={onUndo} 
                  disabled={!canUndo} 
                  className="flex items-center gap-1 px-2 py-1 rounded-md hover:bg-white/20 text-white/60 hover:text-white disabled:opacity-20 transition-all cursor-pointer"
                >
                  <Undo2 className="h-3 w-3" />
                  <span className="text-[9px] font-black uppercase">Undo</span>
                </button>
                <div className="w-[1px] h-2 bg-white/10" />
                <button 
                  onClick={onRedo} 
                  disabled={!canRedo} 
                  className="flex items-center gap-1 px-2 py-1 rounded-md hover:bg-white/20 text-white/60 hover:text-white disabled:opacity-20 transition-all cursor-pointer"
                >
                  <Redo2 className="h-3 w-3" />
                  <span className="text-[9px] font-black uppercase">Redo</span>
                </button>
              </div>
            )}
          </div>
          
          {onRemap && (
            <button
              onClick={onRemap}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest text-white/40 hover:text-white hover:bg-white/10 transition-all cursor-pointer"
            >
              <Columns className="h-3 w-3" />
              Remap Fields
            </button>
          )}
        </div>
      )}
      <div 
        ref={tableRef} 
        onScroll={handleScroll}
        className="flex-1 overflow-y-auto overflow-x-hidden bom-scroll-stable"
      >
        <table className="w-full table-fixed text-center border-collapse select-none">
          <colgroup>
            <col style={{ width: '5%' }} />
            <col style={{ width: '19%' }} />
            <col style={{ width: '8%' }} />
            <col style={{ width: '30%' }} />
            <col style={{ width: '26%' }} />
            <col style={{ width: '12%' }} />
          </colgroup>
          <thead className="sticky top-0 z-[80] shadow-[0_4px_12px_rgba(0,0,0,0.5)]">
            <tr className="bg-[#0b101b] border-b border-white/20">
              <th className="py-2 px-1 text-[10px] uppercase tracking-wider text-white/80 font-bold border-r border-white/10 text-center leading-tight overflow-hidden" title="Index">#</th>
              
              <th className="py-2 px-1 text-[10px] uppercase tracking-wider text-white/80 font-bold border-r border-white/10 text-center relative group overflow-hidden" title="Manufacturer Part Number">
                <div className="flex items-center justify-center gap-1 flex-wrap leading-tight">
                  MPN
                  <FilterDropdown 
                    column="MPN" 
                    values={columnUniqueValues.mpn} 
                    selectedValues={filters.mpn || new Set()} 
                    onFilterChange={(sel) => onFilterChange?.('mpn', sel)}
                    isOpen={openFilter === 'mpn'}
                    onOpenChange={(open) => setOpenFilter(open ? 'mpn' : null)}
                  />
                </div>
              </th>

              <th className="py-2 px-1 text-[10px] uppercase tracking-wider text-white/80 font-bold border-r border-white/10 text-center relative group overflow-hidden" title="Quantity">
                <div className="flex items-center justify-center gap-1 flex-wrap leading-tight">
                  Qty
                  <FilterDropdown 
                    column="Qty" 
                    values={columnUniqueValues.qty} 
                    selectedValues={filters.qty || new Set()} 
                    onFilterChange={(sel) => onFilterChange?.('qty', sel)}
                    isOpen={openFilter === 'qty'}
                    onOpenChange={(open) => setOpenFilter(open ? 'qty' : null)}
                  />
                </div>
              </th>

              <th className="py-2 px-1 text-[10px] uppercase tracking-wider text-white/80 font-bold border-r border-white/10 text-center relative group overflow-hidden" title="Description">
                <div className="flex items-center justify-center gap-1 flex-wrap leading-tight">
                  Description
                  <FilterDropdown 
                    column="Description" 
                    values={columnUniqueValues.description} 
                    selectedValues={filters.description || new Set()} 
                    onFilterChange={(sel) => onFilterChange?.('description', sel)}
                    isOpen={openFilter === 'description'}
                    onOpenChange={(open) => setOpenFilter(open ? 'description' : null)}
                  />
                </div>
              </th>

              <th className="py-2 px-1 text-[10px] uppercase tracking-wider text-white/80 font-bold border-r border-white/10 text-center relative group overflow-hidden" title="Designators / References">
                <div className="flex items-center justify-center gap-1 flex-wrap leading-tight">
                  Designator
                  <FilterDropdown 
                    column="Designator" 
                    values={columnUniqueValues.designators} 
                    selectedValues={filters.designators || new Set()} 
                    onFilterChange={(sel) => onFilterChange?.('designators', sel)}
                    isOpen={openFilter === 'designators'}
                    onOpenChange={(open) => setOpenFilter(open ? 'designators' : null)}
                  />
                </div>
              </th>

              <th className="py-2 px-1 text-[10px] uppercase tracking-wider text-white/80 font-bold text-center relative group overflow-hidden" title="Component Status">
                <div className="flex items-center justify-center gap-1 flex-wrap leading-tight">
                  Status
                  <FilterDropdown 
                    column="Status" 
                    values={columnUniqueValues.status} 
                    selectedValues={filters.status || new Set()} 
                    onFilterChange={(sel) => onFilterChange?.('status', sel)}
                    isOpen={openFilter === 'status'}
                    onOpenChange={(open) => setOpenFilter(open ? 'status' : null)}
                  />
                </div>
              </th>
            </tr>
          </thead>
          <tbody>
            {visibleWindow.topPadding > 0 && (
              <tr style={{ height: `${visibleWindow.topPadding}px` }} aria-hidden="true" />
            )}
            {filteredData.slice(visibleWindow.startIndex, visibleWindow.endIndex).map((row, i) => {
              const badge = statusBadge(row);
              const highlight = highlightLogic?.(row) || '';
              
              return (
                <tr
                  key={row.originalRowIdx}
                  id={`row-${row.originalRowIdx}`}
                  style={{ height: `${ROW_HEIGHT}px` }}
                  className={`border-b border-white/20 hover:bg-white/[0.06] transition-colors group bg-black/15 ${highlight} align-middle relative hover:z-[70] ${
                    highlightedRow === row.originalRowIdx ? 'bg-primary/20 ring-2 ring-primary/50 relative z-[75]' : ''
                  }`}
                >
                  <td className="p-0 text-[11px] text-white/40 font-mono text-center border-r border-white/20 align-middle overflow-hidden">
                    <div className="flex items-center justify-center h-full w-full">{row.originalRowIdx}</div>
                  </td>
                  <td className="p-0 border-r border-white/20 align-middle overflow-hidden text-center">
                    <DoubleClickInput
                      ariaLabel={`Row ${i + 1} MPN`}
                      title={row.mpn}
                      className={`text-[11px] font-semibold font-mono ${
                        row.swappedResult ? 'text-indigo-400' : 'text-white/90'
                      }`}
                      value={row.mpn}
                      onUpdate={val => onUpdate(row.originalRowIdx, 'mpn', val)}
                      placeholder="—"
                    />
                  </td>
                  <td className="p-0 border-r border-white/20 align-middle text-center overflow-hidden">
                    <DoubleClickInput
                      type="number"
                      isNumber
                      ariaLabel={`Row ${i + 1} Quantity`}
                      title={String(row.qty)}
                      className={`text-xs font-semibold font-mono ${
                        highlight.includes('color-sky') || row.qtyMismatchResult ? 'text-sky-400' : 
                        row.qtyMismatch ? 'text-amber-400' : 'text-white/90'
                      }`}
                      value={row.qty}
                      onUpdate={val => onUpdate(row.originalRowIdx, 'qty', val)}
                      placeholder="0"
                    />
                  </td>
                  <td className="p-0 border-r border-white/20 align-middle overflow-hidden text-center">
                    <DoubleClickInput
                      ariaLabel={`Row ${i + 1} Description`}
                      title={row.description || ''}
                      className="text-[11px] font-semibold font-mono text-white/40 group-hover:text-white/70"
                      value={row.description || ''}
                      onUpdate={val => onUpdate(row.originalRowIdx, 'description', val)}
                      placeholder="—"
                    />
                  </td>
                  <td className="p-0 border-r border-white/20 overflow-hidden text-center align-middle">
                    <DesignatorInput 
                      value={row.designators} 
                      onUpdate={(val) => onUpdate(row.originalRowIdx, 'designators', val)} 
                      isModified={row.modifiedFields?.has('designators')}
                      rowIdx={row.originalRowIdx}
                      row={row}
                    />
                  </td>
                  <td className="p-0 align-middle text-center overflow-visible">
                    <div className="flex flex-col items-center justify-center py-2 group/status relative h-full w-full">
                      <div className={`h-3 w-3 rounded-full transition-all duration-500 shadow-glow ${badge.dot}`} />
                      
                      <div className="absolute right-[calc(100%+16px)] top-1/2 -translate-y-1/2 w-64 p-4 rounded-2xl bg-[#080c14] border border-white/20 shadow-[0_20px_50px_rgba(0,0,0,0.8)] opacity-0 group-hover/status:opacity-100 transition-all duration-300 pointer-events-none z-[110] text-left">
                        <div className="text-[10px] font-black text-white/40 uppercase tracking-[0.2em] mb-3 border-b border-white/10 pb-2">Status Ledger</div>
                        <div className="space-y-2">
                          {badge.problems.map((p, idx) => (
                            <div key={idx} className="flex gap-2.5 text-[11px] text-white/80 leading-relaxed font-bold">
                              <span className={badge.dot.split(' ')[0].replace('bg-', 'text-')}>•</span>
                              {p}
                            </div>
                          ))}
                        </div>
                      </div>

                      {onSyncFromMaster && highlightLogic?.(row).includes('yellow') && (
                        <button
                          onClick={() => onSyncFromMaster(row, row.originalRowIdx)}
                          title="Sync MPN from Reference (Master)"
                          className="p-1.5 rounded-md bg-yellow-500/20 text-yellow-400 hover:bg-yellow-500/40 transition-colors cursor-pointer group/sync mt-2"
                        >
                          <Download className="h-3.5 w-3.5 group-hover/sync:scale-110" />
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
            {visibleWindow.bottomPadding > 0 && (
              <tr style={{ height: `${visibleWindow.bottomPadding}px` }} aria-hidden="true" />
            )}
          </tbody>
        </table>
        {filteredData.length === 0 && (
          <div className="flex-1 flex flex-col items-center justify-center py-24 text-white/40 text-xs font-mono space-y-4">
            <div className="p-6 rounded-[2rem] bg-white/5 border border-white/5 opacity-50">
              <Search className="h-12 w-12" />
            </div>
            <p className="font-black uppercase tracking-widest italic">{data.length === 0 ? 'No rows found' : 'No rows match your filters'}</p>
            {data.length > 0 && (
              <button 
                onClick={() => {
                  onFilterChange?.('mpn', new Set());
                  onFilterChange?.('qty', new Set());
                  onFilterChange?.('description', new Set());
                  onFilterChange?.('designators', new Set());
                  onFilterChange?.('status', new Set());
                }}
                className="px-4 py-2 rounded-xl bg-white/5 border border-white/10 text-primary font-black uppercase text-[10px] tracking-widest hover:bg-white/10 transition-all cursor-pointer"
              >
                Clear All Filters
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Upload Zone ──────────────────────────────────────────────────────────────
export function UploadZone({ 
  id, 
  onUpload, 
  label, 
  icon = <Upload className="h-5 w-5" />, 
  isSmall = false, 
  hasFile = false, 
  accept = ".xlsx,.xls,.csv" 
}: { 
  id: string;
  onUpload: (files: FileList) => void;
  label: string;
  isSmall?: boolean;
  icon?: React.ReactNode;
  hasFile?: boolean;
  accept?: string;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = React.useState(false);

  const handleDrag = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setIsDragging(true);
    } else if (e.type === 'dragleave') {
      setIsDragging(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      onUpload(e.dataTransfer.files);
    }
  }, [onUpload]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onUpload(e.target.files);
    }
  };

  return (
    <div 
      className={`relative group w-fit mx-auto transition-all duration-300 ${isDragging ? 'scale-105' : ''}`}
      onDragEnter={handleDrag}
      onDragOver={handleDrag}
      onDragLeave={handleDrag}
      onDrop={handleDrop}
    >
      <label htmlFor={id} className="cursor-pointer">
        <input
          ref={inputRef}
          id={id}
          type="file"
          accept={accept}
          onChange={handleChange}
          className="hidden"
        />
        <div className={`glass-button flex items-center gap-2.5 transition-all duration-500 group-hover:shadow-glow ${
          isDragging ? 'border-emerald-500/50 bg-emerald-500/10 shadow-glow ring-2 ring-emerald-500/20' : ''
        } ${isSmall ? 'px-5 py-2.5' : 'px-8 py-3.5'}`}>
          <div className={`transition-transform duration-500 ${isDragging ? '-translate-y-1' : 'group-hover:-translate-y-0.5'} ${hasFile ? 'text-emerald-400' : 'text-primary'}`}>
            {hasFile ? <CheckCircle2 className={isSmall ? 'h-4 w-4' : 'h-5 w-5'} /> : (icon || <Upload className={isSmall ? 'h-4 w-4' : 'h-5 w-5'} />)}
          </div>
          <span className={`font-bold uppercase tracking-wider ${isSmall ? 'text-xs' : 'text-sm'} ${isDragging ? 'text-emerald-400' : ''}`}>
            {isDragging ? 'Drop File Here' : label}
          </span>
        </div>
      </label>
    </div>
  );
}

// ─── Accordion Section Component ──────────────────────────────────────────────
function AccordionSection({ 
  id, title, count, isOpen, onToggle, children 
}: { 
  id: string; 
  title: string; 
  count: number; 
  isOpen: boolean;
  onToggle: (id: string | null) => void;
  children: React.ReactNode;
}) {
  return (
    <div className={`transition-all duration-300 rounded-2xl border ${isOpen ? 'bg-white/5 border-white/10 shadow-lg' : 'bg-transparent border-transparent hover:bg-white/5 hover:border-white/5'}`}>
      <button
        onClick={() => onToggle(isOpen ? null : id)}
        className={`w-full flex items-center justify-between p-3.5 transition-all duration-300 group sticky top-[-13px] z-30 rounded-t-2xl ${
          isOpen 
            ? 'bg-slate-900 border-b border-white/10 shadow-lg' 
            : 'bg-slate-900 border-b border-transparent hover:bg-white/5'
        }`}
      >
        <div className="flex items-center gap-3">
          <span className="text-xs font-black uppercase tracking-widest text-white/80 group-hover:text-white transition-colors">
            {title}
            {count > 0 && <span className="ml-2 py-0.5 px-1.5 rounded-md bg-white/5 text-[9px] font-mono text-white/40">{count}</span>}
          </span>
        </div>
        <div className="text-white/20 group-hover:text-white/60 transition-colors">
          {isOpen ? <ChevronDown className="h-4 w-4" /> : <ChevronRight className="h-4 w-4" />}
        </div>
      </button>

      <AnimatePresence initial={false}>
        {isOpen && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.3, ease: [0.23, 1, 0.32, 1] }}
            className="overflow-hidden"
          >
            <div className="px-3.5 pb-4 space-y-2">
              <div className="h-px bg-white/5 mb-3 mx-1" />
              {children}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

// ─── Comparison Result Panel ──────────────────────────────────────────────────
function ComparisonResults({ 
  diff, 
  onJump 
}: { 
  diff: BOMDiff; 
  onJump: (rowIdx: number, source: string) => void 
}) {
  const [open, setOpen] = useState<string | null>('summary');

  return (
    <div className="flex flex-col h-full bg-slate-950/40 border border-white/10 rounded-3xl overflow-hidden shadow-2xl backdrop-blur-none">
      <div className="px-5 py-4 border-b border-white/5 flex items-center justify-between shrink-0 bg-white/5">
        <h3 className="text-[10px] font-black uppercase tracking-[0.2em] text-white/60 font-mono">Differential Report</h3>
      </div>

      <div className="flex-1 overflow-y-auto p-3 space-y-2 bom-scroll-stable">
        <AccordionSection 
          id="summary" 
          title="Summary" 
          count={0} 
          isOpen={open === 'summary'}
          onToggle={setOpen}
        >
          <div className="grid grid-cols-2 gap-2">
            {[
              { label: 'Missing', count: diff.missing.length, color: 'text-pink-500', tooltip: "Parts found in the Reference (Master) BOM but missing from your Target file." },
              { label: 'Extra', count: diff.extra.length, color: 'text-teal-400', tooltip: "Parts present in your Target file that do not exist in the Reference (Master) BOM." },
              { label: 'Δ Qty', count: diff.qtyMismatch.length, color: 'text-sky-400', tooltip: "Identical parts found in both files but with a different total count." },
              { label: 'Swapped', count: diff.swapped.length, color: 'text-indigo-400', tooltip: "Components where the designator (e.g., R1) matches, but the part number is different from the Master." },
              { label: 'Dup', count: diff.duplicates.target.length + diff.duplicates.source.length, color: 'text-orange-500', tooltip: "Self-duplicates found within the files." },
              { label: 'Qty Err', count: diff.internalQtyErrors.target.length + diff.internalQtyErrors.source.length, color: 'text-amber-500', tooltip: "Internal quantity mismatches within the files." },
            ].map(s => (
              <div key={s.label} title={s.tooltip} className="p-3 rounded-xl bg-white/5 border border-white/5 text-center cursor-help">
                <div className={`text-xl font-black ${s.color}`}>{s.count}</div>
                <div className="text-[8px] uppercase font-bold tracking-widest text-white/30 mt-0.5 whitespace-nowrap overflow-hidden text-ellipsis">{s.label}</div>
              </div>
            ))}
          </div>
        </AccordionSection>

        {diff.missing.length > 0 && (
          <AccordionSection id="missing" title="Missing" count={diff.missing.length} isOpen={open === 'missing'} onToggle={setOpen}>
            {diff.missing.map((m, i) => (
              <div 
                key={i} 
                onClick={() => onJump(m.originalRowIdx, 'ref')}
                className="p-2.5 rounded-lg bg-pink-500/5 border border-pink-500/10 text-[11px] font-medium text-pink-200/70 cursor-pointer hover:bg-pink-500/10 transition-all group"
              >
                <div className="flex justify-between items-start mb-0.5">
                  <span className="font-black text-pink-400 group-hover:text-pink-300 transition-colors">{m.mpn}</span>
                  <span className="text-[9px] font-bold text-white/30 italic">Row #{m.originalRowIdx}</span>
                </div>
                {m.designators.join(', ')}
              </div>
            ))}
          </AccordionSection>
        )}

        {diff.extra.length > 0 && (
          <AccordionSection id="extra" title="Extra" count={diff.extra.length} isOpen={open === 'extra'} onToggle={setOpen}>
            {diff.extra.map((e, i) => (
              <div 
                key={i} 
                onClick={() => onJump(e.originalRowIdx, 'target')}
                className="p-2.5 rounded-lg bg-teal-500/5 border border-teal-500/10 text-[11px] font-medium text-teal-200/70 cursor-pointer hover:bg-teal-500/10 transition-all group"
              >
                <div className="flex justify-between items-start mb-0.5">
                  <span className="font-black text-teal-400 group-hover:text-teal-300 transition-colors">{e.mpn}</span>
                  <span className="text-[9px] font-bold text-white/30 italic">Row #{e.originalRowIdx}</span>
                </div>
                {e.designators.join(', ')}
              </div>
            ))}
          </AccordionSection>
        )}

        {diff.qtyMismatch.length > 0 && (
          <AccordionSection id="qty" title="Δ Qty / Mismatch" count={diff.qtyMismatch.length} isOpen={open === 'qty'} onToggle={setOpen}>
            {diff.qtyMismatch.map((q, i) => (
              <div 
                key={i} 
                onClick={() => onJump(q.targetRowIndices[0], 'target')}
                className="p-2.5 rounded-lg bg-sky-500/5 border border-sky-500/10 text-[11px] font-medium text-sky-200/70 cursor-pointer hover:bg-sky-500/10 transition-all group"
              >
                <div className="flex justify-between items-start mb-1">
                  <span className="font-black text-sky-400 group-hover:text-sky-300 transition-colors">{q.mpn}</span>
                  <div className="flex flex-col items-end gap-0.5">
                    <span className="text-[10px] font-black px-1.5 py-0.5 rounded bg-sky-500/20 text-sky-300">Δ {q.actualQty - q.expectedQty}</span>
                    <span className="text-[8px] font-bold text-white/20 italic">Ref #{q.sourceRowIndices.join(',')} · Target #{q.targetRowIndices.join(',')}</span>
                  </div>
                </div>
                <p className="text-[10px] opacity-60">Ref: {q.expectedQty} · Target: {q.actualQty}</p>
              </div>
            ))}
          </AccordionSection>
        )}

        {diff.swapped.length > 0 && (
          <AccordionSection id="swapped" title="Swapped" count={diff.swapped.length} isOpen={open === 'swapped'} onToggle={setOpen}>
            {diff.swapped.map((s, i) => (
              <div 
                key={i} 
                onClick={() => onJump(s.targetRowIdx, 'target')}
                className="p-2.5 rounded-lg bg-indigo-500/5 border border-indigo-500/10 text-[11px] font-medium text-indigo-200/70 cursor-pointer hover:bg-indigo-500/10 transition-all group"
              >
                <div className="flex justify-between items-start mb-1">
                  <span className="font-black text-indigo-400 group-hover:text-indigo-300 transition-colors">{s.designator}</span>
                  <span className="text-[9px] font-bold text-white/30 italic">Ref #{s.sourceRowIdx} · Target #{s.targetRowIdx}</span>
                </div>
                <div className="flex items-center gap-2 opacity-80">
                  <span className="line-through text-white/40">{s.sourceMPN}</span>
                  <ArrowRightLeft className="h-3 w-3 text-indigo-500/50" />
                  <span className="text-white">{s.targetMPN}</span>
                </div>
              </div>
            ))}
          </AccordionSection>
        )}

        {(diff.duplicates.source.length > 0 || diff.duplicates.target.length > 0) && (
          <AccordionSection id="duplicates" title="Duplicates (Internal)" count={diff.duplicates.source.length + diff.duplicates.target.length} isOpen={open === 'duplicates'} onToggle={setOpen}>
            {[...diff.duplicates.source.map(r => ({...r, side: 'Ref'})), ...diff.duplicates.target.map(r => ({...r, side: 'Target'}))].map((d, i) => (
              <div 
                key={i} 
                onClick={() => onJump(d.originalRowIdx, d.side === 'Ref' ? 'ref' : 'target')}
                className="p-2.5 rounded-lg bg-orange-500/5 border border-orange-500/10 text-[11px] font-medium text-orange-200/70 cursor-pointer hover:bg-orange-500/10 transition-all group"
              >
                <div className="flex justify-between items-start mb-0.5 text-[9px] uppercase font-bold tracking-tighter opacity-40 italic">
                  <span>{d.side}</span>
                  <span>Row #{d.originalRowIdx}</span>
                </div>
                <span className="font-black text-orange-400 group-hover:text-orange-300 transition-colors block mb-0.5">{d.mpn}</span>
                {d.designators.join(', ')}
              </div>
            ))}
          </AccordionSection>
        )}

        {(diff.internalQtyErrors.source.length > 0 || diff.internalQtyErrors.target.length > 0) && (
          <AccordionSection id="qtyErrors" title="Internal Qty Errors" count={diff.internalQtyErrors.source.length + diff.internalQtyErrors.target.length} isOpen={open === 'qtyErrors'} onToggle={setOpen}>
            {[...diff.internalQtyErrors.source.map(r => ({...r, side: 'Ref'})), ...diff.internalQtyErrors.target.map(r => ({...r, side: 'Target'}))].map((q, i) => (
              <div 
                key={i} 
                onClick={() => onJump(q.originalRowIdx, q.side === 'Ref' ? 'ref' : 'target')}
                className="p-2.5 rounded-lg bg-yellow-500/5 border border-yellow-500/10 text-[11px] font-medium text-yellow-200/70 cursor-pointer hover:bg-yellow-500/10 transition-all group"
              >
                <div className="flex justify-between items-start mb-1 text-[9px] uppercase font-bold tracking-tighter opacity-40 italic">
                  <span>{q.side}</span>
                  <span>Row #{q.originalRowIdx}</span>
                </div>
                <span className="font-black text-yellow-400 group-hover:text-yellow-300 transition-colors block mb-0.5">{q.mpn}</span>
                <p className="text-[10px] opacity-60">Row Qty: {q.qty} · D-Count: {q.designators.length} </p>
              </div>
            ))}
          </AccordionSection>
        )}
      </div>
    </div>
  );
}

// ─── Sheet Picker Modal ──────────────────────────────────────────────────────
export function SheetPicker({ state, onCancel }: { state: SheetPickerState; onCancel: () => void }) {
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-6 sm:p-12">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onCancel}
        className="absolute inset-0 bg-slate-950/80"
      />
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-lg glass-card p-10 space-y-8 border-white/10"
      >
        <div className="text-center space-y-2">
          <Database className="h-10 w-10 text-emerald-400 mx-auto mb-2" />
          <h2 className="text-3xl font-black text-white">Select Sheet</h2>
          <p className="text-white/30 text-[10px] uppercase font-black tracking-[.2em]">Processing {state.filename}</p>
        </div>

        <div className="grid gap-3 max-h-[40vh] overflow-y-auto pr-2 bom-scroll-stable">
          {state.sheetNames.map(name => (
            <button
              key={name}
              onClick={() => state.onSheetSelected(name)}
              className="group relative p-5 rounded-[1.5rem] bg-white/5 border border-white/5 hover:border-emerald-500/30 hover:bg-emerald-500/5 transition-all text-left overflow-hidden cursor-pointer"
            >
              <div className="flex items-center justify-between">
                <span className="text-sm font-black text-white/50 group-hover:text-emerald-400 transition-colors uppercase tracking-tight">{name}</span>
                <ChevronRight className="h-5 w-5 text-white/10 group-hover:text-emerald-500 transition-all group-hover:translate-x-1" />
              </div>
            </button>
          ))}
        </div>

        <button 
          onClick={onCancel}
          className="w-full py-4 text-[10px] font-black uppercase tracking-widest text-white/20 hover:text-white transition-all cursor-pointer"
        >
          Cancel
        </button>
      </motion.div>
    </div>
  );
}

// ─── Help / About Dialog ──────────────────────────────────────────────────────
export function AboutDialog({ onClose }: { onClose: () => void }) {
  const openExternal = (url: string) => {
    if (window.ipcRenderer) {
      window.ipcRenderer.invoke('open-url', url);
    } else {
      window.open(url, '_blank');
    }
  };

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-6 sm:p-12">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/80"
      />
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-xl bg-slate-900 border border-white/10 rounded-[2.5rem] shadow-2xl overflow-hidden"
      >
        <div className="absolute top-0 left-0 w-full h-1.5 bg-gradient-to-r from-emerald-500 via-blue-400 to-indigo-500" />
        
        <div className="p-10 text-center space-y-8">
          <div className="space-y-4">
            <div className="relative w-24 h-24 mx-auto mb-6 group">
              <div className="absolute inset-0 bg-emerald-500/20 rounded-3xl blur-2xl group-hover:bg-emerald-500/40 transition-all duration-700" />
              <div className="relative glass-card w-full h-full flex items-center justify-center bg-white/5 border-white/10 rounded-[2rem] shadow-xl group-hover:scale-110 transition-transform duration-500">
                <img src="./icon.png" alt="BOM Tool" className="w-full h-full object-contain brightness-125 transition-all duration-500" />
              </div>
            </div>
            
            <h2 className="text-4xl font-black tracking-tight text-white mb-2">BOM Tool</h2>
            <div className="flex items-center justify-center gap-2">
              <span className="px-3 py-1 rounded-full bg-white/5 border border-white/10 text-[10px] font-black uppercase tracking-widest text-white/40">Version 1.2.5</span>
              <span className="text-[10px] font-black text-emerald-500/60 uppercase tracking-widest">Stable Release</span>
            </div>
          </div>

          <div className="space-y-6 max-w-sm mx-auto">
            <div className="relative p-6 rounded-[2rem] bg-[#080d1a]/80 border border-white/15 shadow-[0_8px_32px_rgba(0,0,0,0.5),inset_0_1px_1px_rgba(255,255,255,0.15)] flex flex-col items-center justify-center space-y-3">
              <span className="text-[9px] font-black uppercase tracking-[0.3em] text-amber-200/60 leading-none">
                DEVELOPED BY
              </span>
              <h3 className="text-3xl font-black uppercase tracking-[0.18em] bg-gradient-to-r from-amber-100 via-white to-amber-300 bg-clip-text text-transparent [text-shadow:0_0_20px_rgba(251,191,36,0.3)]">
                MEET MISTRY
              </h3>
              <div className="flex items-center justify-center gap-4 pt-1">
                <button 
                  onClick={() => openExternal('https://meetmistry.vercel.app')}
                  title="Visit Developer Portfolio"
                  className="p-3 rounded-2xl bg-white/5 border border-white/10 text-white/60 hover:text-white hover:bg-white/10 transition-all group shadow-sm cursor-pointer"
                >
                  <ExternalLink className="h-5 w-5" />
                </button>
                <button 
                  onClick={() => openExternal('mailto:meetzarc@gmail.com')}
                  className="px-6 py-2.5 rounded-2xl bg-emerald-500/10 border border-primary/20 text-emerald-400 hover:bg-emerald-500/20 transition-all font-bold text-xs flex items-center gap-2 cursor-pointer"
                >
                  <Bug className="h-4 w-4" />
                  Report a bug
                </button>
              </div>
            </div>
          </div>

          <button 
            onClick={onClose}
            className="w-full max-w-[140px] px-8 py-3 rounded-2xl bg-white text-slate-950 font-black text-xs uppercase tracking-widest hover:bg-blue-50 transition-all shadow-xl hover:scale-105 active:scale-95 cursor-pointer"
          >
            Got it
          </button>
        </div>
      </motion.div>
    </div>
  );
}


function WindowControls() {
  const handleMinimize = () => window.ipcRenderer.send('window-minimize');
  const handleClose = () => window.ipcRenderer.send('window-close');

  return (
    <div className="flex items-center gap-1 no-drag">
      <button 
        onClick={handleMinimize}
        title="Minimize"
        className="p-2.5 rounded-xl text-white/60 hover:text-white hover:bg-white/10 border border-transparent hover:border-white/10 transition-all cursor-pointer group"
      >
        <Minus className="h-4 w-4 transition-transform group-hover:scale-110" strokeWidth={2.5} />
      </button>
      <button 
        onClick={handleClose}
        title="Close"
        className="p-2.5 rounded-xl text-white/60 hover:text-red-400 hover:bg-red-500/10 border border-transparent hover:border-red-500/20 transition-all cursor-pointer group"
      >
        <X className="h-4 w-4 transition-transform group-hover:scale-110" strokeWidth={2.5} />
      </button>
    </div>
  );
}


// ─── Main Dashboard ───────────────────────────────────────────────────────────
export function BOMDashboard() {
  const [activeTab, setActiveTab] = useState<Tab>('analyzer');
  const [uiScale, setUiScale] = useState(1);

  useEffect(() => {
    function handleResize() {
      const w = window.innerWidth;
      // Using 1440px as the 'Golden Width' where the UI looks perfect.
      if (w < 1440) {
        setUiScale(w / 1440);
      } else {
        setUiScale(1);
      }
    }
    window.addEventListener('resize', handleResize);
    handleResize();
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Analyzer state
  const [analyzerFile, setAnalyzerFile] = useState<FileState>(emptyFile());
  const [analyzerHistory, setAnalyzerHistory] = useState<FileState[]>([emptyFile()]);
  const [analyzerPointer, setAnalyzerPointer] = useState(0);
  const [analyzerPending, setAnalyzerPending] = useState<PendingMapping | null>(null);

  // Comparator state
  const [refFile, setRefFile] = useState<FileState>(emptyFile());
  const [refHistory, setRefHistory] = useState<FileState[]>([emptyFile()]);
  const [refPointer, setRefPointer] = useState(0);

  const [targetFile, setTargetFile] = useState<FileState>(emptyFile());
  const [targetHistory, setTargetHistory] = useState<FileState[]>([emptyFile()]);
  const [targetPointer, setTargetPointer] = useState(0);

  const [refPending, setRefPending] = useState<PendingMapping | null>(null);
  const [targetPending, setTargetPending] = useState<PendingMapping | null>(null);

  // PnP Validator state
  const [pnpFile, setPnpFile] = useState<PnPFileData<PnPRow> | null>(null);
  const [pnpBomFile, setPnpBomFile] = useState<PnPFileData<BOMRow> | null>(null);
  const [pnpHasValidated, setPnpHasValidated] = useState(false);


  // Status modal state
  const [statusModal, setStatusModal] = useState<StatusModalState | null>(null);

  // Filter states
  const [analyzerFilters, setAnalyzerFilters] = useState<Record<string, Set<string>>>({});
  const [refFilters, setRefFilters] = useState<Record<string, Set<string>>>({});
  const [targetFilters, setTargetFilters] = useState<Record<string, Set<string>>>({});

  // Universal Search state
  const [showSearch, setShowSearch] = useState(false);
  const [jumpTarget, setJumpTarget] = useState<{ rowIdx: number; source: string; timestamp: number } | null>(null);

  // Scroll Preservation State
  const [analyzerScroll, setAnalyzerScroll] = useState(0);
  const [pnpScroll, setPnpScroll] = useState(0);
  const [refScroll, setRefScroll] = useState(0);
  const [targetScroll, setTargetScroll] = useState(0);

  // Keyboard Shortcuts
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      // Ctrl+F or Cmd+F
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        setShowSearch(true);
      }
    };
    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => window.removeEventListener('keydown', handleGlobalKeyDown);
  }, []);

  const handleSearchResultClick = useCallback((result: { source: string; rowIdx: number }) => {
    // Search is mode-specific — no tab switch needed.
    // Just set the jump target for the current mode's grid and close.
    // Map 'reference' from UniversalSearch to 'ref' used in dashboards
    const source = result.source === 'reference' ? 'ref' : result.source;
    setJumpTarget({ rowIdx: result.rowIdx, source, timestamp: Date.now() });
    setShowSearch(false);
  }, []);

  // --- Search Sources Logic ---
  const searchSources = useMemo(() => {
    if (activeTab === 'analyzer') {
      return { analyzer: analyzerFile.rows.length > 0 ? { rows: analyzerFile.rows, filename: analyzerFile.filename } : undefined };
    }
    if (activeTab === 'comparator') {
      return { 
        reference: refFile.rows.length > 0 ? { rows: refFile.rows, filename: refFile.filename } : undefined,
        target: targetFile.rows.length > 0 ? { rows: targetFile.rows, filename: targetFile.filename } : undefined 
      };
    }
    if (activeTab === 'pnp') {
      // For PnP, we search both the Placement and the BOM
      return {
        pnp: (pnpFile && pnpFile.rows) ? { rows: pnpFile.rows as unknown as PnPSearchRow[], filename: pnpFile.filename } : undefined,
        reference: (pnpBomFile && pnpBomFile.rows) ? { rows: pnpBomFile.rows, filename: pnpBomFile.filename } : undefined
      };
    }
    return {};
  }, [activeTab, analyzerFile, refFile, targetFile, pnpFile, pnpBomFile]);

  const searchTitle = useMemo(() => {
    if (activeTab === 'analyzer') return 'Analyzer Database Search';
    if (activeTab === 'comparator') return 'Comparator Database Search';
    if (activeTab === 'pnp') return 'PnP Audit Search';
    return 'Database Search';
  }, [activeTab]);

  // Clear jump target whenever the active tab changes so stale highlights
  // never bleed into a different mode's grid.
  useEffect(() => {
    setJumpTarget(null);
  }, [activeTab]);


  // Activity Log State
  const [activities, setActivities] = useState<{
    id: string, 
    action: string, 
    result: string, 
    timestamp: string, 
    type: 'info' | 'success',
    mode: 'analyzer' | 'comparator',
    details?: { mpn: string; row: string; details: string; type: string }[]
  }[]>([]);

  const addActivity = (action: string, result: string, type: 'info' | 'success' = 'info', details?: { mpn: string; row: string; details: string; type: string }[]) => {
    setActivities(prev => [{
      id: Math.random().toString(36).substring(7),
      action,
      result,
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      type,
      mode: activeTab as 'analyzer' | 'comparator',
      details: (details as { mpn: string; row: string; details: string; type: string }[])
    }, ...prev].slice(0, 50));
  };

  const analyzerActivities = useMemo(() => activities.filter(a => a.mode === 'analyzer'), [activities]);

  // Whether the user has explicitly started a comparison. The diff itself is
  // always derived live from the current rows, so the report panel, the status
  // dots and the summary counters can never disagree.
  const [hasCompared, setHasCompared] = useState(false);

  // Reactive Comparison Engine
  const comparisonResult = useMemo(() => {
    if (activeTab === 'comparator' && refFile.rows.length && targetFile.rows.length) {
      return BOMCompare.compare(refFile.rows, targetFile.rows);
    }
    return null;
  }, [refFile.rows, targetFile.rows, activeTab]);

  const comparison = hasCompared ? (comparisonResult?.diff ?? null) : null;

  const handleStartComparison = () => {
    if (comparisonResult) {
      setHasCompared(true);
      const diffCount = comparisonResult.diff.missing.length + comparisonResult.diff.extra.length;
      addActivity('Comparison Complete', `Identified ${diffCount} primary discrepancies between BOMs.`, 'success');
    }
  };

  // Sheet picker state (single shared picker)
  const [sheetPicker, setSheetPicker] = useState<SheetPickerState | null>(null);

  // Auto-prompt comparison when both files are loaded
  useEffect(() => {
    if (activeTab === 'comparator' && refFile.rows.length > 0 && targetFile.rows.length > 0 && !comparison) {
      setStatusModal({
        type: 'success',
        title: 'Analysis Ready',
        message: 'Both BOM files have been successfully processed and verified. You can now proceed with the comparison.'
      });
    }
  }, [refFile.rows.length, targetFile.rows.length, comparison, activeTab]);

  // State for About dialog
  const [showAbout, setShowAbout] = useState(false);

  // ── History Helpers ──────────────────────────────────────────────────────────
  const commit = (type: 'analyzer' | 'ref' | 'target', next: FileState) => {
    if (type === 'analyzer') {
      const nextH = [...analyzerHistory.slice(0, analyzerPointer + 1), next];
      setAnalyzerHistory(nextH);
      setAnalyzerPointer(nextH.length - 1);
      setAnalyzerFile(next);
    } else if (type === 'ref') {
      const nextH = [...refHistory.slice(0, refPointer + 1), next];
      setRefHistory(nextH);
      setRefPointer(nextH.length - 1);
      setRefFile(next);
    } else {
      const nextH = [...targetHistory.slice(0, targetPointer + 1), next];
      setTargetHistory(nextH);
      setTargetPointer(nextH.length - 1);
      setTargetFile(next);
    }
  };

  const undo = (type: 'analyzer' | 'ref' | 'target') => {
    if (type === 'analyzer' && analyzerPointer > 0) {
      setAnalyzerPointer(p => p - 1);
      setAnalyzerFile(analyzerHistory[analyzerPointer - 1]);
    } else if (type === 'ref' && refPointer > 0) {
      setRefPointer(p => p - 1);
      setRefFile(refHistory[refPointer - 1]);
    } else if (type === 'target' && targetPointer > 0) {
      setTargetPointer(p => p - 1);
      setTargetFile(targetHistory[targetPointer - 1]);
    }
  };

  const redo = (type: 'analyzer' | 'ref' | 'target') => {
    if (type === 'analyzer' && analyzerPointer < analyzerHistory.length - 1) {
      setAnalyzerPointer(p => p + 1);
      setAnalyzerFile(analyzerHistory[analyzerPointer + 1]);
    } else if (type === 'ref' && refPointer < refHistory.length - 1) {
      setRefPointer(p => p + 1);
      setRefFile(refHistory[refPointer + 1]);
    } else if (type === 'target' && targetPointer < targetHistory.length - 1) {
      setTargetPointer(p => p + 1);
      setTargetFile(targetHistory[targetPointer + 1]);
    }
  };
  // Core parse function: reads a File, shows sheet picker if multi-sheet,
  // then calls onParsed with the chosen sheet's ParseResult.
  const parseFile = useCallback((
    file: File,
    onParsed: (result: ParseResult, filename: string) => void
  ) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(arrayBuffer, { type: 'array' });

        const doParseSheet = (ws: XLSX.WorkSheet, sheetName: string) => {
          const result = BOMProcessor.parseSheet(ws, XLSX.utils.sheet_to_json);
          onParsed({ ...result, sheetName, buffer: arrayBuffer }, file.name);
        };

        if (wb.SheetNames.length > 1) {
          setSheetPicker({
            sheetNames: wb.SheetNames,
            workbookBinary: '',
            filename: file.name,
            onSheetSelected: (name) => {
              setSheetPicker(null);
              doParseSheet(wb.Sheets[name], name);
            },
          });
        } else {
          const name = wb.SheetNames[0];
          doParseSheet(wb.Sheets[name], name);
        }
      } catch (err) {
        console.error('Failed to parse BOM:', err);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const resetAnalyzerHistory = (file: FileState) => {
    setAnalyzerFile(file);
    setAnalyzerHistory([file]);
    setAnalyzerPointer(0);
    setAnalyzerFilters({});
  };
  const resetRefHistory = (file: FileState) => {
    setRefFile(file);
    setRefHistory([file]);
    setRefPointer(0);
    setRefFilters({});
    setHasCompared(false);
  };
  const resetTargetHistory = (file: FileState) => {
    setTargetFile(file);
    setTargetHistory([file]);
    setTargetPointer(0);
    setTargetFilters({});
    setHasCompared(false);
  };

  // ── Analyzer Upload ─────────────────────────────────────────────────────────
  const handleAnalyzerUpload = (files: FileList) => {
    const file = files[0];
    if (!file) return;
    parseFile(file, (result, filename) => {
      if (result.needsManualMapping) {
        setAnalyzerPending({ 
          rawData: result.rawData, 
          rawAOA: result.rawAOA,
          headerRowIdx: result.headerRowIdx,
          mapping: result.columnMapping, 
          availableColumns: result.availableColumns, 
          filename,
          merges: result.merges,
          buffer: result.buffer as ArrayBuffer,
          sheetName: result.sheetName
        });
      } else {
        resetAnalyzerHistory({ 
          rows: result.rows, 
          rawData: result.rawData, 
          rawAOA: result.rawAOA,
          headerRowIdx: result.headerRowIdx,
          columnMapping: result.columnMapping, 
          availableColumns: result.availableColumns, 
          filename,
          merges: result.merges,
          buffer: result.buffer as ArrayBuffer,
          sheetName: result.sheetName
        });
      }
    });
  };

  const confirmAnalyzerMapping = (m: ColumnMapping) => {
    if (!analyzerPending) return;
    const rows = BOMProcessor.parseBOMWithMapping(analyzerPending.rawData, m);
    resetAnalyzerHistory({ 
      rows, 
      rawData: analyzerPending.rawData, 
      rawAOA: analyzerPending.rawAOA,
      headerRowIdx: analyzerPending.headerRowIdx,
      columnMapping: m, 
      availableColumns: analyzerPending.availableColumns, 
      filename: analyzerPending.filename,
      merges: analyzerPending.merges,
      buffer: analyzerPending.buffer,
      sheetName: analyzerPending.sheetName
    });
    setAnalyzerPending(null);
  };

  // ── Comparator Uploads ──────────────────────────────────────────────────────
  const handleRefUpload = (files: FileList) => {
    const file = files[0];
    if (!file) return;
    parseFile(file, (result, filename) => {
      if (result.needsManualMapping) {
        setRefPending({ 
          rawData: result.rawData, 
          rawAOA: result.rawAOA,
          headerRowIdx: result.headerRowIdx,
          mapping: result.columnMapping, 
          availableColumns: result.availableColumns, 
          filename,
          merges: result.merges,
          buffer: result.buffer as ArrayBuffer,
          sheetName: result.sheetName
        });
      } else {
        resetRefHistory({ 
          rows: result.rows, 
          rawData: result.rawData, 
          rawAOA: result.rawAOA,
          headerRowIdx: result.headerRowIdx,
          columnMapping: result.columnMapping, 
          availableColumns: result.availableColumns, 
          filename,
          merges: result.merges,
          buffer: result.buffer as ArrayBuffer,
          sheetName: result.sheetName
        });
      }
    });
  };

  const handleTargetUpload = (files: FileList) => {
    const file = files[0];
    if (!file) return;
    parseFile(file, (result, filename) => {
      if (result.needsManualMapping) {
        setTargetPending({ 
          rawData: result.rawData, 
          rawAOA: result.rawAOA,
          headerRowIdx: result.headerRowIdx,
          mapping: result.columnMapping, 
          availableColumns: result.availableColumns, 
          filename,
          merges: result.merges,
          buffer: result.buffer as ArrayBuffer,
          sheetName: result.sheetName
        });
      } else {
        resetTargetHistory({ 
          rows: result.rows, 
          rawData: result.rawData, 
          rawAOA: result.rawAOA,
          headerRowIdx: result.headerRowIdx,
          columnMapping: result.columnMapping, 
          availableColumns: result.availableColumns, 
          filename,
          merges: result.merges,
          buffer: result.buffer as ArrayBuffer,
          sheetName: result.sheetName
        });
      }
    });
  };

  const confirmRefMapping = (m: ColumnMapping) => {
    if (!refPending) return;
    const rows = BOMProcessor.parseBOMWithMapping(refPending.rawData, m, refPending.headerRowIdx);
    resetRefHistory({ 
      rows, 
      rawData: refPending.rawData, 
      rawAOA: refPending.rawAOA,
      headerRowIdx: refPending.headerRowIdx,
      columnMapping: m, 
      availableColumns: refPending.availableColumns, 
      filename: refPending.filename,
      merges: refPending.merges,
      buffer: refPending.buffer,
      sheetName: refPending.sheetName
    });
    setRefPending(null);
  };

  const confirmTargetMapping = (m: ColumnMapping) => {
    if (!targetPending) return;
    const rows = BOMProcessor.parseBOMWithMapping(targetPending.rawData, m, targetPending.headerRowIdx);
    resetTargetHistory({ 
      rows, 
      rawData: targetPending.rawData, 
      rawAOA: targetPending.rawAOA,
      headerRowIdx: targetPending.headerRowIdx,
      columnMapping: m, 
      availableColumns: targetPending.availableColumns, 
      filename: targetPending.filename,
      merges: targetPending.merges,
      buffer: targetPending.buffer,
      sheetName: targetPending.sheetName
    });
    setTargetPending(null);
  };

  // ── Row Updates ─────────────────────────────────────────────────────────────
  const updateRow = (type: 'analyzer' | 'ref' | 'target', originalRowIdx: number, field: keyof BOMRow, val: BOMRow[keyof BOMRow]) => {
    const file = type === 'analyzer' ? analyzerFile : type === 'ref' ? refFile : targetFile;
    const nextRows = [...file.rows];
    const realIdx = nextRows.findIndex(r => r.originalRowIdx === originalRowIdx);
    if (realIdx === -1) return;
    
    const nextRow = { ...nextRows[realIdx], [field]: val };
    
    // Live Row Validation: Recalculate Qty Mismatch for the modified row
    nextRow.qtyMismatch = BOMProcessor.isQtyMismatch(nextRow.qty, nextRow.designators);
    
    // Add field to modified set ONLY IF it changed from the ORIGINAL data
    const modifiedFields = new Set(nextRow.modifiedFields || []);
    const { originalData: orig } = nextRow;
    
    let isDifferent = false;
    if (orig) {
      if (field === 'designators') {
        const currentDesArr = (val as string[]).map(d => d.trim().toUpperCase());
        const originalDesArr = (orig.designators || []).map(d => d.trim().toUpperCase());
        isDifferent = currentDesArr.join(',') !== originalDesArr.join(',');
      } else if (field === 'qty') {
        isDifferent = Number(val) !== orig.qty;
      } else {
        isDifferent = String(val).trim() !== String((orig as Record<string, unknown>)[field] || '').trim();
      }
    } else {
      isDifferent = true; // Fallback if no original data
    }
    
    if (isDifferent) {
      modifiedFields.add(field as 'mpn' | 'qty' | 'designators' | 'description');
    } else {
      modifiedFields.delete(field as 'mpn' | 'qty' | 'designators' | 'description');
    }
    nextRow.modifiedFields = modifiedFields;
    
    nextRows[realIdx] = nextRow;
    const recalculatedRows = BOMProcessor.recalculateDuplicates(nextRows);
    commit(type, { ...file, rows: recalculatedRows, merges: file.merges });
  };

  const handleSyncFromMaster = (row: BOMRow, targetIdx: number) => {
    // Find a matching reference row by searching for ANY of the current row's designators
    const refRow = refFile.rows.find(r => 
      r.designators.some(rd => row.designators.some(td => td.trim().toUpperCase() === rd.trim().toUpperCase()))
    );

    if (refRow) {
      const nextTarget = [...targetFile.rows];
      nextTarget[targetIdx] = { 
        ...nextTarget[targetIdx], 
        mpn: refRow.mpn,
        description: refRow.description
      };
      commit('target', { ...targetFile, rows: nextTarget, merges: targetFile.merges });
    }
  };

  // ── Tools ───────────────────────────────────────────────────────────────────
  const runAnalyzerTool = (tool: ToolId) => {
    let nextRows = [...analyzerFile.rows];
    let msg = "";

    switch (tool) {
      case 'validate': {
        const validNextRows = nextRows.map(r => ({
          ...r,
          qtyMismatch: BOMProcessor.isQtyMismatch(r.qty, r.designators)
        }));
        const valDups = BOMProcessor.findSelfDuplicates(validNextRows);
        nextRows = validNextRows.map(r => ({
          ...r,
          duplicate: r.designators.some(d => valDups.has(d.trim().toUpperCase()))
        }));
        
        const details: { mpn: string; row: string; details: string; type: string }[] = [];
        nextRows.forEach(r => {
          if (r.duplicate) {
            const myDups = r.designators.filter(d => valDups.has(d.trim().toUpperCase())).join(', ');
            details.push({ mpn: r.mpn, row: String(r.originalRowIdx), details: `Duplicate: ${myDups}`, type: 'duplicate' });
          }
          if (r.qtyMismatch) {
            details.push({ mpn: r.mpn, row: String(r.originalRowIdx), details: `Qty Error: ${r.qty} Pcs. vs ${r.designators.length} Refs.`, type: 'qtyMismatch' });
          }
        });

        const dupCount = nextRows.filter(r => r.duplicate).length;
        const qtyErrCount = nextRows.filter(r => r.qtyMismatch).length;
        const issues = dupCount + qtyErrCount;
        msg = `Validation complete: ${issues} issue${issues !== 1 ? 's' : ''} found.`;
        addActivity('Validate BOM', msg, issues > 0 ? 'info' : 'success', details);
        break;
      }
      case 'duplicates': {
        const selfDups = BOMProcessor.findSelfDuplicates(nextRows);
        nextRows = nextRows.map(r => ({
          ...r,
          duplicate: r.designators.some(d => selfDups.has(d.trim().toUpperCase()))
        }));
        
        const details: { mpn: string; row: string; details: string; type: string }[] = [];
        nextRows.forEach(r => {
          if (r.duplicate) {
            const myDups = r.designators.filter(d => selfDups.has(d.trim().toUpperCase())).join(', ');
            details.push({ mpn: r.mpn, row: String(r.originalRowIdx), details: `Duplicate: ${myDups}`, type: 'duplicate' });
          }
        });

        const dupRowsCount = nextRows.filter(r => r.duplicate).length;
        msg = `Duplicate check complete: ${dupRowsCount} rows affected.`;
        addActivity('Check Duplicates', msg, dupRowsCount > 0 ? 'info' : 'success', details);
        break;
      }
      case 'normalize': {
        let normCount = 0;
        const details: { mpn: string; row: string; details: string; type: string }[] = [];
        nextRows = nextRows.map(r => {
          const before = r.designators.join(', ');
          const { expanded: nextDes } = BOMProcessor.expandDesignators(BOMProcessor.normalizeDesignatorString(before));
          const after = nextDes.join(', ');
          if (before === after) return r;
          
          normCount++;
          details.push({
            mpn: r.mpn,
            row: String(r.originalRowIdx),
            details: `Cleaned: [${before}] ➔ [${after}]`,
            type: 'normalize'
          });
          const modifiedFields = new Set(r.modifiedFields || []);
          modifiedFields.add('designators');
          return { ...r, designators: nextDes, modifiedFields };
        });
        msg = normCount > 0
          ? `Normalized: ${normCount} row${normCount > 1 ? 's' : ''} cleaned.`
          : `Normalize: No changes needed — all rows already clean.`;
        addActivity('Normalize Designators', msg, normCount > 0 ? 'success' : 'info', details);
        break;
      }
      case 'expand': {
        let expCount = 0;
        const details: { mpn: string; row: string; details: string; type: string }[] = [];
        nextRows = nextRows.map(r => {
          const before = r.designators.join(', ');
          const { expanded: nextDes, isExpanded } = BOMProcessor.expandDesignators(before);
          const after = nextDes.join(', ');
          if (before === after) return r;
          
          expCount++;
          details.push({
            mpn: r.mpn,
            row: String(r.originalRowIdx),
            details: `Expanded: [${before}] ➔ [${after}]`,
            type: 'expand'
          });
          const modifiedFields = new Set(r.modifiedFields || []);
          if (isExpanded) modifiedFields.add('designators');
          return { ...r, designators: nextDes, modifiedFields };
        });
        msg = expCount > 0
          ? `Expanded: ${expCount} row${expCount > 1 ? 's' : ''} had ranges expanded.`
          : `Expand: No ranges found — all designators already expanded.`;
        addActivity('Expand Ranges', msg, expCount > 0 ? 'success' : 'info', details);
        break;
      }
    }

    commit('analyzer', { ...analyzerFile, rows: nextRows });
    setStatusModal({ type: 'info', title: 'Tool Execution', message: msg });
  };

  const runComparatorTool = (tool: ToolId) => {
    let msg = '';
    const validateFile = (file: FileState) => {
      const dups = BOMProcessor.findSelfDuplicates(file.rows);
      return file.rows.map(r => {
        const rowDups = new Set<string>();
        r.designators.forEach(d => {
          if (dups.has(d.trim().toUpperCase())) rowDups.add(d.trim().toUpperCase());
        });
        return {
          ...r,
          qtyMismatch: BOMProcessor.isQtyMismatch(r.qty, r.designators),
          duplicate: rowDups.size > 0,
          duplicateDesignators: rowDups
        };
      });
    };

    switch (tool) {
      case 'validate': {
        const nextRefRows = validateFile(refFile);
        const nextTargetRows = validateFile(targetFile);
        commit('ref', { ...refFile, rows: nextRefRows });
        commit('target', { ...targetFile, rows: nextTargetRows });
        const totalIssues = nextRefRows.filter(r => r.duplicate || r.qtyMismatch).length + 
                           nextTargetRows.filter(r => r.duplicate || r.qtyMismatch).length;
        msg = `Found ${totalIssues} consistency issues across Ref & Target files.`;
        addActivity('Validate (Comparator)', msg, totalIssues > 0 ? 'info' : 'success');
        break;
      }
      case 'normalize': {
        let normCount = 0;
        const details: { mpn: string; row: string; details: string; type: string; source: 'ref' | 'target' }[] = [];
        
        const normalizeFile = (file: FileState, source: 'ref' | 'target') => {
          return file.rows.map(r => {
            const before = r.designators.join(', ');
            const { expanded: nextDes } = BOMProcessor.expandDesignators(BOMProcessor.normalizeDesignatorString(before));
            const after = nextDes.join(', ');
            if (before === after) return r;
            normCount++;
            details.push({
              mpn: r.mpn,
              row: String(r.originalRowIdx),
              details: `Cleaned (${source}): [${before}] ➔ [${after}]`,
              type: 'normalize',
              source
            });
            const modifiedFields = new Set(r.modifiedFields || []);
            modifiedFields.add('designators');
            return { ...r, designators: nextDes, modifiedFields };
          });
        };

        const nextRefRows = normalizeFile(refFile, 'ref');
        const nextTargetRows = normalizeFile(targetFile, 'target');
        commit('ref', { ...refFile, rows: nextRefRows });
        commit('target', { ...targetFile, rows: nextTargetRows });
        msg = normCount > 0 ? `Normalized ${normCount} rows across both files.` : "No normalization needed.";
        addActivity('Normalize (Comparator)', msg, normCount > 0 ? 'success' : 'info', details);
        break;
      }
      case 'expand': {
        let expCount = 0;
        const details: { mpn: string; row: string; details: string; type: string; source: 'ref' | 'target' }[] = [];
        
        const expandFile = (file: FileState, source: 'ref' | 'target') => {
          return file.rows.map(r => {
            const before = r.designators.join(', ');
            const { expanded: nextDes, isExpanded } = BOMProcessor.expandDesignators(before);
            const after = nextDes.join(', ');
            if (before === after) return r;
            expCount++;
            details.push({
              mpn: r.mpn,
              row: String(r.originalRowIdx),
              details: `Expanded (${source}): [${before}] ➔ [${after}]`,
              type: 'expand',
              source
            });
            const modifiedFields = new Set(r.modifiedFields || []);
            if (isExpanded) modifiedFields.add('designators');
            return { ...r, designators: nextDes, modifiedFields };
          });
        };

        const nextRefRows = expandFile(refFile, 'ref');
        const nextTargetRows = expandFile(targetFile, 'target');
        commit('ref', { ...refFile, rows: nextRefRows });
        commit('target', { ...targetFile, rows: nextTargetRows });
        msg = expCount > 0 ? `Expanded ${expCount} rows across both files.` : "No ranges to expand.";
        addActivity('Expand (Comparator)', msg, expCount > 0 ? 'success' : 'info', details);
        break;
      }
    }

    setStatusModal({ type: 'info', title: 'Compare Tool Results', message: msg });
  };


  // ── Save ────────────────────────────────────────────────────────────────────
  const handleSave = async (file: FileState, fallbackName: string) => {
    try {
      const { columnMapping: mapping, headerRowIdx, rows: editedRows, rawData: objects, buffer: originalBuffer, sheetName } = file;
      if (!mapping.mpn) throw new Error("No MPN column mapped. Cannot update data.");
      if (!originalBuffer || originalBuffer.byteLength === 0) throw new Error("Original file buffer not available for high-fidelity export.");

      // 1. Load original template using ExcelJS to preserve styles/logos/borders
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(originalBuffer);
      const sheet = workbook.getWorksheet(sheetName || 1);
      
      if (!sheet) throw new Error(`Could not find sheet "${sheetName || 'default'}" in original workbook.`);

      // 2. Map mapping header strings to 1-indexed ExcelJS column numbers
      const headerRow = sheet.getRow(headerRowIdx + 1);
      const getColIndex = (name: string | null | undefined) => {
        if (!name) return -1;
        let found = -1;
        headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (String(cell.value || '').trim() === name) found = colNumber;
        });
        return found;
      };

      const mpnCol = getColIndex(mapping.mpn);
      const qtyCol = getColIndex(mapping.qty);
      const desCol = getColIndex(mapping.designators);
      const descCol = getColIndex(mapping.description);

      // 3. Update data rows using absolute __rowIdx from the original parse
      objects.forEach((obj, idx) => {
        const edited = editedRows[idx];
        if (!edited) return;
        
        const rIdx = (obj as Record<string, number>).__rowIdx + 1; // 1-indexed for ExcelJS
        const row = sheet.getRow(rIdx);

        if (mpnCol !== -1) row.getCell(mpnCol).value = edited.mpn;
        if (qtyCol !== -1) row.getCell(qtyCol).value = edited.qty;
        if (desCol !== -1) row.getCell(desCol).value = edited.designators.join(', ');
        if (descCol !== -1) row.getCell(descCol).value = edited.description || '';
        
        row.commit();
      });

      // 4. Comparison Results Sheet
      if (activeTab === 'comparator' && comparison) {
        // Remove old comparison sheet if it exists to avoid duplicates on re-save
        const existing = workbook.getWorksheet('Comparison Analysis');
        if (existing) workbook.removeWorksheet(existing.id);

        const analysisSheet = workbook.addWorksheet('Comparison Analysis');

        // Column Config with initial defaults (will be auto-resized later)
        analysisSheet.columns = [
          { header: 'CATEGORY', key: 'category', width: 20 },
          { header: 'ITEM (MPN / DESIGNATOR)', key: 'item', width: 35 },
          { header: 'EXCEL ROW(S)', key: 'rows', width: 15 },
          { header: 'DETAILS / DESCRIPTION', key: 'details', width: 50 },
          { header: 'SEVERITY', key: 'severity', width: 12 }
        ];

        // 1. Title Row
        analysisSheet.insertRow(1, ['BOM COMPARISON ANALYSIS REPORT']);
        analysisSheet.mergeCells(1, 1, 1, 5);
        const titleCell = analysisSheet.getCell(1, 1);
        titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } }; // Deep Blue
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        analysisSheet.getRow(1).height = 40;

        // 2. Style the Header Row (Now Row 2)
        const headerRowRow = analysisSheet.getRow(2);
        headerRowRow.height = 25;
        headerRowRow.eachCell((cell) => {
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } }; // Slate 700
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = { 
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };
        });

        analysisSheet.views = [{ state: 'frozen', ySplit: 2 }];

        const addRow = (data: { category: string; item: string; rows: string; details: string; severity: string }, colors: { bg: string; text: string }) => {
          const r = analysisSheet.addRow(data);
          r.height = 20;
          r.eachCell((cell, colNumber) => {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.font = { size: 10, color: { argb: colors.text } };
            cell.border = { 
              top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
            if (colNumber === 5) { // Severity column background
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.bg } };
              cell.font = { bold: true, color: { argb: colors.text } };
            }
          });
          return r;
        };

        const CATEGORY_STYLES = {
          critical: { bg: 'FFFEE2E2', text: 'FFB91C1C' }, // Red
          warning: { bg: 'FFFFF7ED', text: 'FFC2410C' }, // Orange
          info: { bg: 'FFF0F9FF', text: 'FF0369A1' },    // Sky
          success: { bg: 'FFF0FDF4', text: 'FF15803D' }  // Emerald
        };

        // --- Data Populate ---
        if (comparison.missing.length) {
          comparison.missing.forEach(m => addRow({ 
            category: 'MISSING', 
            item: m.mpn, 
            rows: String(m.originalRowIdx), 
            details: `Found in Reference but completely missing from Target BOM.`, 
            severity: 'CRITICAL' 
          }, CATEGORY_STYLES.critical));
        }

        if (comparison.swapped.length) {
          comparison.swapped.forEach(s => addRow({ 
            category: 'SWAPPED', 
            item: s.designator, 
            rows: `${s.sourceRowIdx} / ${s.targetRowIdx}`, 
            details: `Reference MPN: ${s.sourceMPN} | Target MPN: ${s.targetMPN}`, 
            severity: 'CRITICAL' 
          }, CATEGORY_STYLES.critical));
        }

        if (comparison.qtyMismatch.length) {
          comparison.qtyMismatch.forEach(q => addRow({ 
            category: 'QTY DELTA', 
            item: q.mpn, 
            rows: `${q.sourceRowIndices.join(',')} | ${q.targetRowIndices.join(',')}`, 
            details: `Reference Qty: ${q.expectedQty} | Target Qty: ${q.actualQty} (Delta: ${q.actualQty - q.expectedQty})`, 
            severity: 'INFO' 
          }, CATEGORY_STYLES.info));
        }

        if (comparison.extra.length) {
          comparison.extra.forEach(e => addRow({ 
            category: 'EXTRA', 
            item: e.mpn, 
            rows: String(e.originalRowIdx), 
            details: `Present in Target file but does not exist in the Reference BOM.`, 
            severity: 'INFO' 
          }, CATEGORY_STYLES.info));
        }

        if (comparison.duplicates.source.length || comparison.duplicates.target.length) {
          [...comparison.duplicates.source.map(r => ({ ...r, side: 'REF' })), ...comparison.duplicates.target.map(r => ({ ...r, side: 'TARGET' }))].forEach(d => addRow({ 
            category: 'DUPLICATE', 
            item: d.mpn, 
            rows: String(d.originalRowIdx), 
            details: `[${d.side}] Duplicate designators found within this file: ${d.designators.join(', ')}`, 
            severity: 'WARNING' 
          }, CATEGORY_STYLES.warning));
        }

        if (comparison.internalQtyErrors.source.length || comparison.internalQtyErrors.target.length) {
          [...comparison.internalQtyErrors.source.map(r => ({ ...r, side: 'REF' })), ...comparison.internalQtyErrors.target.map(r => ({ ...r, side: 'TARGET' }))].forEach(q => addRow({ 
            category: 'ERROR', 
            item: q.mpn, 
            rows: String(q.originalRowIdx), 
            details: `[${q.side}] Internal Quantity mismatch: Row Qty (${q.qty}) differs from Designator count (${q.designators.length})`, 
            severity: 'WARNING' 
          }, CATEGORY_STYLES.warning));
        }

        // AUTO-FIT COLUMNS: Calculate max length for each column
        analysisSheet.columns.forEach((column) => {
          let maxLength = 0;
          column.eachCell?.({ includeEmpty: true }, (cell) => {
            const columnLength = cell.value ? String(cell.value).length : 0;
            if (columnLength > maxLength) {
              maxLength = columnLength;
            }
          });
          // Set width with some padding (approx 1.2 padding factor + buffer)
          column.width = Math.min(Math.max(12, maxLength + 5), 80); 
        });
      }

      // 5. Analyzer Results Sheet
      if (activeTab === 'analyzer' && activities.length > 0) {
        const lastActivity = activities[0];
        
        // Remove old analyzer report sheet if it exists
        const existing = workbook.getWorksheet('Analyzer Report');
        if (existing) workbook.removeWorksheet(existing.id);

        const analyzerSheet = workbook.addWorksheet('Analyzer Report');

        analyzerSheet.columns = [
          { header: 'ITEM (MPN / DESIGNATOR)', key: 'mpn', width: 35 },
          { header: 'EXCEL ROW', key: 'row', width: 15 },
          { header: 'DETAILS / DESCRIPTION', key: 'details', width: 60 }
        ];

        // 1. Title Row
        analyzerSheet.insertRow(1, [`BOM ANALYZER REPORT: ${lastActivity.action.toUpperCase()}`]);
        analyzerSheet.mergeCells(1, 1, 1, 3);
        const titleCell = analyzerSheet.getCell(1, 1);
        titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D9488' } }; // Teal 600
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        analyzerSheet.getRow(1).height = 40;

        // Date / Result Summary Row
        analyzerSheet.insertRow(2, [`Summary: ${lastActivity.result}`]);
        analyzerSheet.mergeCells(2, 1, 2, 3);
        const subTitleCell = analyzerSheet.getCell(2, 1);
        subTitleCell.font = { italic: true, size: 11, color: { argb: 'FF334155' } };
        subTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        analyzerSheet.getRow(2).height = 25;

        // 2. Style the Header Row (Now Row 3)
        const headerRowRow = analyzerSheet.getRow(3);
        headerRowRow.height = 25;
        headerRowRow.eachCell((cell) => {
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } }; // Slate 700
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = { 
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };
        });

        analyzerSheet.views = [{ state: 'frozen', ySplit: 3 }];

        const addAnalyzerRow = (data: { mpn: string; row: string; details: string }, isSummary = false) => {
          const r = analyzerSheet.addRow(data);
          r.height = 20;
          r.eachCell((cell) => {
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { size: 10, color: { argb: 'FF0F172A' } };
            if (isSummary) {
               cell.font.italic = true;
               cell.font.color = { argb: 'FF64748B' };
            }
            cell.border = { 
              top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
          });
          return r;
        };

        if (lastActivity.details && lastActivity.details.length > 0) {
          lastActivity.details.forEach(d => {
            addAnalyzerRow({
              mpn: d.mpn || '-',
              row: d.row || '-',
              details: d.details || '-'
            });
          });
        } else {
          // If no specific details, just put a placeholder row
          addAnalyzerRow({
            mpn: '-',
            row: '-',
            details: 'No specific row-level modifications logged for this action.'
          }, true);
        }

        // AUTO-FIT
        analyzerSheet.columns.forEach((column) => {
          let maxLength = 0;
          column.eachCell?.({ includeEmpty: true }, (cell, rowNum) => {
            // Skip title heading row from auto-fit
            if (rowNum <= 2) return;
            const columnLength = cell.value ? String(cell.value).length : 0;
            if (columnLength > maxLength) {
              maxLength = columnLength;
            }
          });
          column.width = Math.min(Math.max(15, maxLength + 5), 100); 
        });
      }

      // 6. File Naming
      const originalFullName = file.filename || `${fallbackName}.xlsx`;
      let baseName = originalFullName.replace(/^BT_/, '');
      if (!baseName.toLowerCase().endsWith('.xlsx') && !baseName.toLowerCase().endsWith('.xls')) {
        baseName += '.xlsx';
      }
      const finalName = `BT_${baseName}`;

      // 6. Generate buffer and save
      const outBuffer = await workbook.xlsx.writeBuffer();
      const result = await window.ipcRenderer.invoke('save-excel', { buffer: outBuffer, filename: finalName });

      if (result.success) {
        setStatusModal({
          type: 'success',
          title: 'BOM Exported Successfully',
          message: `File saved to: ${result.path}`
        });
      } else if (result.error) {
        throw new Error(result.error);
      }
    } catch (err) {
      console.error('Save error:', err);
      setStatusModal({
        type: 'error',
        title: 'Export Failed',
        message: err instanceof Error ? err.message : 'An unknown error occurred during Excel generation.'
      });
    }
  };



  // ─── Render ─────────────────────────────────────────────────────────────────
  const analyzerHasData = analyzerFile.rows.length > 0;
  const bothComparatorFilesLoaded = refFile.rows.length > 0 && targetFile.rows.length > 0;

  return (
    <div 
      className="bg-transparent text-white selection:bg-primary/40 overflow-hidden flex flex-col"
      style={{ 
        zoom: uiScale,
        // When zoom < 1, 100vh shrinks visually. We compensate by pre-expanding
        // height and width using inverse scaling: 100/zoom. This guarantees the
        // container always fills the full physical screen after zoom is applied.
        height: `${100 / uiScale}vh`,
        width: `${100 / uiScale}vw`,
        minWidth: '1440px'
      } as React.CSSProperties}
    >
      <div className="flex-1 flex flex-col pt-3 px-3 pb-4 overflow-hidden relative">
        {statusModal && (
          <StatusModal
            {...statusModal}
            onClose={() => setStatusModal(null)}
          />
        )}

        {/* Header - Integrated Drag Region (Full Width) */}
        <header 
          className="w-full flex items-center justify-between gap-4 mb-4 shrink-0 transition-all duration-500 pl-1 drag"
        >
          <motion.div 
            initial={{ opacity: 0, x: -30 }} 
            animate={{ opacity: 1, x: 0 }}
            className="no-drag"
          >
            <div className="flex items-center gap-5">
              {/* Logo Group */}
              <div 
                className="relative group flex-shrink-0 cursor-pointer"
                onClick={() => setShowAbout(true)}
                title="About BOM Tool"
              >
                <div className="relative glass-card p-2 bg-white/5 border-white/10 rounded-2xl shadow-2xl group-hover:bg-white/10 group-hover:border-white/20 transition-all duration-300">
                  <img 
                    src="./icon.png" 
                    alt="BOM Tool Logo" 
                    className="h-10 w-10 object-contain brightness-125 drop-shadow-[0_0_15px_rgba(59,130,246,0.3)] transition-transform duration-500 group-hover:scale-110" 
                  />
                </div>
              </div>

              {/* Brand Title & Credit Pill */}
              <div className="flex flex-col gap-1">
                <div className="flex items-center gap-4">
                  <h1 className="text-4xl font-black tracking-tight whitespace-nowrap shrink-0 bg-clip-text text-transparent bg-gradient-to-br from-white via-white/90 to-white/40 select-none">
                    BOM Tool
                  </h1>
                  
                  {/* High-Fidelity Luxury Developer Badge */}
                  <div className="relative group/pill cursor-default select-none shrink-0">
                    <div className="absolute -inset-0.5 bg-gradient-to-r from-amber-500/20 via-blue-500/20 to-emerald-500/20 rounded-xl blur-sm opacity-40 group-hover/pill:opacity-100 transition-opacity duration-500" />
                    <div className="relative flex flex-col items-center justify-center px-3.5 py-1.5 bg-[#090f1d]/90 backdrop-blur-md border border-white/15 rounded-xl shadow-[0_4px_20px_rgba(0,0,0,0.5),inset_0_1px_1px_rgba(255,255,255,0.15)] group-hover/pill:border-amber-400/40 transition-all duration-300">
                      <span className="text-[7.5px] font-black uppercase tracking-[0.25em] text-amber-200/60 group-hover/pill:text-amber-200/90 transition-colors leading-none mb-0.5 whitespace-nowrap">
                        DEVELOPED BY
                      </span>
                      <span className="text-[10.5px] font-black uppercase tracking-[0.16em] bg-gradient-to-r from-amber-100 via-white to-amber-300 bg-clip-text text-transparent group-hover/pill:[text-shadow:0_0_12px_rgba(251,191,36,0.6)] transition-all leading-none whitespace-nowrap">
                        MEET MISTRY
                      </span>
                    </div>
                  </div>
                </div>
                <p className="text-blue-100/70 font-black text-[10px] tracking-[0.2em] uppercase flex items-center whitespace-nowrap">ULTRA-FAST HARMONIZED BOM INTELLIGENCE</p>
              </div>
            </div>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, x: 30 }}
            animate={{ opacity: 1, x: 0 }}
            className="flex items-center gap-4 no-drag"

          >
            {/* Quick Actions Tile — only when a file is loaded */}
            <AnimatePresence>
              {activeTab === 'analyzer' && analyzerHasData && (
                <motion.div
                  initial={{ opacity: 0, scale: 0.85 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 0.85 }}
                  className="flex items-center gap-1.5 bg-[#0a101d] p-1.5 rounded-2xl border border-white/10 shadow-2xl"
                >
                  <button
                    onClick={() => {
                      if (activeTab === 'analyzer') handleSave(analyzerFile, 'analyzed_bom');
                      else if (targetFile.rows.length) handleSave(targetFile, 'target_bom');
                      else if (refFile.rows.length) handleSave(refFile, 'reference_bom');
                    }}
                    title="Save / Export current BOM"
                    className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold text-emerald-300 bg-emerald-500/15 border border-emerald-500/25 hover:bg-emerald-500/25 hover:border-emerald-400/50 transition-all duration-200 cursor-pointer"
                  >
                    <Save className="h-3.5 w-3.5" />
                    Save
                  </button>
                  <button
                    onClick={() => {
                      if (activeTab === 'analyzer') commit('analyzer', emptyFile());
                      else { setRefFile(emptyFile()); setTargetFile(emptyFile()); setHasCompared(false); }
                    }}
                    title="Clear all loaded data"
                    className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold text-white/40 hover:text-red-400 hover:bg-red-500/10 border border-transparent hover:border-red-400/20 transition-all duration-200 cursor-pointer"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    Clear
                  </button>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Tab Nav */}
            <nav className="bg-[#0a101d] p-1.5 rounded-2xl border border-white/10 shadow-2xl pointer-events-auto flex gap-3 items-center">
              <div className="flex gap-1">
                <TabButton active={activeTab === 'analyzer'} onClick={() => setActiveTab('analyzer')} icon={<Database className="h-4 w-4" />} label="BOM Analyzer" />
                <TabButton active={activeTab === 'pnp'} onClick={() => setActiveTab('pnp')} icon={<MapPin className="h-4 w-4" />} label="PnP Validator" />
                <TabButton active={activeTab === 'comparator'} onClick={() => setActiveTab('comparator')} icon={<GitCompareArrows className="h-4 w-4" />} label="Comparator" />
              </div>
            </nav>

            {/* Window Controls */}
            <div className="h-10 w-[1px] bg-white/10 mx-1" />
            <WindowControls />
          </motion.div>
        </header>

      {/* Main — Full Width to match Header */}
      <main className="w-full flex-1 overflow-hidden">
        <AnimatePresence mode="wait">
          {/* ═══ ANALYZER ═══════════════════════════════════════════════════════ */}
          {activeTab === 'analyzer' && (
            <motion.div
              key="analyzer"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.02 }}
              className="h-full flex flex-col gap-4"
            >
              {!analyzerHasData ? (
                /* Upload state */
                <div className="flex-1 flex flex-col items-center justify-center glass-card">
                  <div className="text-center space-y-6">
                    <div className="space-y-2">
                      <h2 className="text-3xl font-bold tracking-tight bg-gradient-to-r from-white via-blue-100 to-white bg-clip-text text-transparent">
                        BOM Analyzer
                      </h2>
                      <p className="text-blue-100/50 text-sm font-medium">Upload an Excel or CSV file to begin analysis</p>
                    </div>
                    <UploadZone id="analyzer-upload" onUpload={handleAnalyzerUpload} label="Choose BOM File" />
                    <p className="text-white/30 text-xs font-mono">.xlsx · .xls · .csv</p>
                  </div>
                </div>
              ) : (
                /* Data state */
                <>
                  {/* Main Analyzer Content split between Grid and side-panel */}
                  <div className="flex-1 flex gap-4 min-h-0">
                    <div className="flex-1 flex flex-col gap-3 min-w-0">
                      {/* Toolbar within data state */}
                      <div className="flex items-center justify-between gap-4 shrink-0 bg-[#0d1527] p-3 px-4 rounded-2xl border border-white/10 shadow-2xl">
                        <ToolBar 
                          filename={analyzerFile.filename}
                          onRunTool={runAnalyzerTool}
                          onUndo={() => undo('analyzer')}
                          onRedo={() => redo('analyzer')}
                          canUndo={analyzerPointer > 0}
                          canRedo={analyzerPointer < analyzerHistory.length - 1}
                          onSearch={() => setShowSearch(true)}
                          onRemap={() => setAnalyzerPending({
                            rawData: analyzerFile.rawData,
                            rawAOA: analyzerFile.rawAOA,
                            headerRowIdx: analyzerFile.headerRowIdx,
                            mapping: analyzerFile.columnMapping,
                            availableColumns: analyzerFile.availableColumns,
                            filename: analyzerFile.filename,
                            merges: analyzerFile.merges,
                            buffer: analyzerFile.buffer,
                            sheetName: analyzerFile.sheetName
                          })}
                        />
                      </div>

                      {/* Status legend */}
                      <div className="flex items-center gap-4 px-3 py-1.5 shrink-0 bg-white/5 rounded-xl border border-white/5">
                        {[
                          { dot: 'bg-emerald-500', label: 'OK' },
                          { dot: 'bg-slate-400', label: 'Value Auto-Filled' },
                          { dot: 'bg-orange-500', label: 'DUP (SELF)' },
                          { dot: 'bg-amber-400', label: 'QTY ERR' },
                          { dot: 'bg-rose-500', label: 'Both' },
                        ].map(s => (
                          <div key={s.label} className="flex items-center gap-1.5 shrink-0 whitespace-nowrap">
                            <div className={`h-2 w-2 rounded-full shrink-0 ${s.dot}`} />
                            <span className="text-[10px] text-white/70 font-bold uppercase tracking-widest">{s.label}</span>
                          </div>
                        ))}
                        <div className="flex-1" />
                        <div className="flex items-center gap-6 pr-2">
                          <div className="flex flex-col items-end shrink-0 whitespace-nowrap">
                            <span className="text-[8px] font-black uppercase tracking-widest text-white/20 mb-0.5">Board Scale</span>
                            <span className="text-[10px] text-blue-400 font-black uppercase tracking-widest">
                              {analyzerFile.rows.reduce((acc, r) => acc + (Number(r.qty) || 0), 0)} Total Pop.
                            </span>
                          </div>
                          <div className="w-[1px] h-6 bg-white/5 mx-1" />
                          <div className="flex flex-col items-end shrink-0 whitespace-nowrap">
                            <span className="text-[8px] font-black uppercase tracking-widest text-white/20 mb-0.5">Grid Visibility</span>
                            <span className="text-[10px] text-emerald-400 font-black uppercase tracking-widest">
                              {analyzerFile.rows.length} Rows Loaded
                            </span>
                          </div>
                          <div className="w-[1px] h-6 bg-white/5 mx-1" />
                          <div className="flex flex-col items-end shrink-0 whitespace-nowrap">
                            <span className="text-[8px] font-black uppercase tracking-widest text-white/20 mb-0.5">BOM Health</span>
                            <span className={`text-[10px] font-black uppercase tracking-widest italic ${
                              (analyzerFile.rows.filter(r => r.duplicate).length + analyzerFile.rows.filter(r => r.qtyMismatch).length) > 0
                                ? 'text-red-400' : 'text-emerald-400'
                            }`}>
                              {analyzerFile.rows.filter(r => r.duplicate).length + analyzerFile.rows.filter(r => r.qtyMismatch).length} Total Issues
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* Grid */}
                      <div className="flex-1 min-h-0">
                        <EditableBOMGrid
                          type="analyzer"
                          data={analyzerFile.rows}
                          onUpdate={(rowIdx, f, v) => updateRow('analyzer', rowIdx, f, v)}
                          filters={analyzerFilters}
                          onFilterChange={(col, sel) => setAnalyzerFilters(prev => ({ ...prev, [col]: sel }))}
                          jumpTarget={activeTab === 'analyzer' && jumpTarget?.source === 'analyzer' ? jumpTarget : null}
                          initialScroll={analyzerScroll}
                          onScrollChange={setAnalyzerScroll}
                        />
                      </div>
                    </div>

                    {/* Activity Panel */}
                    <aside className="w-80 flex flex-col gap-3 shrink-0">
                      <div className="flex-1 glass-card border-white/5 flex flex-col min-h-0 overflow-hidden bg-slate-900/40 backdrop-blur-2xl">
                        <div className="p-4 border-b border-white/5 flex items-center justify-between shrink-0">
                          <div className="flex items-center gap-2">
                            <RefreshCw className="h-4 w-4 text-primary" />
                            <h3 className="text-xs font-black uppercase tracking-[0.2em] text-white/60">Activity Log</h3>
                          </div>
                          <button 
                            onClick={() => setActivities(prev => prev.filter(a => a.mode !== 'analyzer'))}
                            className="text-[10px] font-bold text-white/20 hover:text-white/40 uppercase tracking-widest cursor-pointer"
                          >
                            Clear
                          </button>
                        </div>

                        {/* File Summary Results */}
                        <div className="grid grid-cols-2 gap-2 p-3 pb-4 bg-white/5 border-b border-white/5">
                          {[
                            { label: 'Dup (Self)', count: analyzerFile.rows.filter(r => r.duplicate).length, color: 'text-orange-400', tip: "Designators appearing multiple times in this project." },
                            { label: 'Qty Err', count: analyzerFile.rows.filter(r => r.qtyMismatch).length, color: 'text-yellow-400', tip: "Inventory Qty doesn't match the number of Designators." },
                          ].map(s => (
                            <div 
                              key={s.label} 
                              title={s.tip}
                              className="p-3 rounded-2xl bg-white/5 border border-white/5 text-center shadow-inner group transition-all duration-300 hover:bg-white/10 cursor-help"
                            >
                              <div className={`text-xl font-black ${s.color} transition-transform group-hover:scale-110`}>{s.count}</div>
                              <div className="text-[8px] uppercase font-bold tracking-[0.2em] text-white/30 mt-1 whitespace-nowrap overflow-hidden text-ellipsis">{s.label}</div>
                            </div>
                          ))}
                        </div>

                        <div className="flex-1 overflow-y-auto p-3 space-y-3 bom-scroll-stable">
                          <AnimatePresence initial={false}>
                            {analyzerActivities.length === 0 ? (
                              <div className="h-full flex flex-col items-center justify-center text-center p-6 space-y-2 opacity-50">
                                <Info className="h-8 w-8 text-white/40" />
                                <p className="text-[10px] uppercase font-bold tracking-widest text-white/60">No recent activity</p>
                              </div>
                            ) : analyzerActivities.map((act) => (
                              <motion.div
                                key={act.id}
                                initial={{ opacity: 0, x: 10 }}
                                animate={{ opacity: 1, x: 0 }}
                                className={`p-3 rounded-xl border border-white/5 bg-white/5 space-y-2`}
                              >
                                <div className="sticky top-[-13px] z-30 -mx-3 -mt-3 p-3 mb-1 bg-slate-900 border-b border-white/10 rounded-t-xl flex items-center justify-between transition-all shadow-lg shadow-black/20">
                                  <span className="text-[10px] font-black uppercase tracking-widest text-white/80">{act.action}</span>
                                  <span className="text-[9px] font-mono text-white/20">{act.timestamp}</span>
                                </div>
                                <div className="px-0.5 pb-1">
                                  <p className={`text-[11px] leading-relaxed font-bold ${act.type === 'success' ? 'text-emerald-400' : 'text-blue-300'}`}>
                                    {act.result}
                                  </p>
                                </div>
                                
                                {act.details && act.details.length > 0 && (
                                  <div className="mt-3 space-y-1.5 pt-3 border-t border-white/5">
                                    <div className="text-[8px] font-black uppercase tracking-[0.2em] text-white/20 mb-2">Issue Breakdown</div>
                                    <div className="space-y-2 max-h-60 overflow-y-auto pr-1 bom-scroll-stable">
                                      {act.details.map((d, idx) => (
                                        <div 
                                          key={idx} 
                                          onClick={() => setJumpTarget({ rowIdx: Number(d.row), source: 'analyzer', timestamp: Date.now() })}
                                          className="p-2 rounded-lg bg-black/20 border border-white/5 group hover:border-white/20 transition-all cursor-pointer shadow-sm hover:shadow-md"
                                        >
                                          <div className="flex items-center justify-between gap-2 mb-1">
                                            <span className="text-[10px] font-black text-white/90 truncate group-hover:text-primary transition-colors">{d.mpn}</span>
                                            <span className="text-[9px] font-mono text-white/30 shrink-0">Row #{d.row}</span>
                                          </div>
                                          {(() => {
                                            const isDup = d.details?.toLowerCase().includes('duplicate');
                                            const colorClass = isDup ? 'text-orange-500' : 'text-yellow-400';
                                            const dotClass = isDup ? 'bg-orange-500 shadow-[0_0_8px_rgba(249,115,22,0.6)]' : 'bg-yellow-400 shadow-[0_0_8px_rgba(250,204,21,0.5)]';
                                            
                                            return (
                                              <div className={`flex items-start gap-1.5 text-[10px] leading-tight font-black ${colorClass}`}>
                                                <div className={`mt-1 h-1.5 w-1.5 rounded-full shrink-0 ${dotClass}`} />
                                                {d.details}
                                              </div>
                                            );
                                          })()}
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                )}
                              </motion.div>
                            ))}
                          </AnimatePresence>
                        </div>
                      </div>
                    </aside>
                  </div>
                </>
              )}
            </motion.div>
          )}

          {/* ═══ PNP VALIDATOR ═════════════════════════════════════════════════════ */}
          {activeTab === 'pnp' && (
            <motion.div
              key="pnp"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.02 }}
              className="h-full flex flex-col gap-4"
            >
              <PnPValidator 
                pnpFile={pnpFile} 
                setPnpFile={setPnpFile}
                bomFile={pnpBomFile}
                setBomFile={setPnpBomFile}
                hasValidated={pnpHasValidated}
                setHasValidated={setPnpHasValidated}
                onSearch={undefined}
                jumpTarget={activeTab === 'pnp' && jumpTarget?.source === 'pnp' ? jumpTarget : null}
                initialScroll={pnpScroll}
                onScrollChange={setPnpScroll}
              />
            </motion.div>
          )}

          {/* ═══ COMPARATOR ═════════════════════════════════════════════════════ */}
          {activeTab === 'comparator' && (
            <motion.div
              key="comparator"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="h-full flex flex-col gap-4"
            >
              {/* Upload panels: shown when at least one file is missing */}
              {(refFile.rows.length === 0 || targetFile.rows.length === 0) ? (
                /* Upload state for both sides */
                <div className="flex-1 flex flex-col gap-4">
                  <div className="flex-1 grid grid-cols-1 md:grid-cols-2 gap-6">
                    {/* Reference */}
                    <div className="glass-card p-10 flex flex-col items-center justify-center space-y-6">
                      <div className="text-center space-y-2">
                        <h3 className="text-2xl font-bold tracking-tight bg-gradient-to-r from-white via-blue-200 to-white bg-clip-text text-transparent">
                          Reference BOM
                        </h3>
                        <p className="text-white/30 text-xs uppercase tracking-widest font-medium">Master baseline</p>
                      </div>
                      <UploadZone
                        id="ref-upload"
                        onUpload={handleRefUpload}
                        label={refFile.rows.length ? 'Change File' : 'Upload Reference'}
                        icon={<Database className="h-4 w-4" />}
                        isSmall
                        hasFile={refFile.rows.length > 0}
                      />
                      {refFile.rows.length > 0 && (
                        <div className="flex flex-col items-center gap-3">
                          <p className="text-xs text-emerald-400/60 font-mono">{refFile.filename} · {refFile.rows.length} rows</p>
                          <button 
                            onClick={() => { setRefFile(emptyFile()); setHasCompared(false); }}
                            className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 hover:bg-red-400/5 transition-all cursor-pointer border border-transparent hover:border-red-400/20"
                          >
                            <Trash2 className="h-3 w-3" />
                            Clear Ref File
                          </button>
                        </div>
                      )}
                    </div>

                    {/* Target */}
                    <div className="glass-card p-10 flex flex-col items-center justify-center space-y-6">
                      <div className="text-center space-y-2">
                        <h3 className="text-2xl font-bold tracking-tight bg-gradient-to-r from-white via-emerald-200 to-white bg-clip-text text-transparent">
                          Your BOM
                        </h3>
                        <p className="text-white/30 text-xs uppercase tracking-widest font-medium">Target implementation</p>
                      </div>
                      <UploadZone
                        id="target-upload"
                        onUpload={handleTargetUpload}
                        label={targetFile.rows.length ? 'Change File' : 'Upload Target'}
                      />
                      {targetFile.rows.length > 0 && (
                        <div className="flex flex-col items-center gap-3">
                          <p className="text-xs text-emerald-400/60 font-mono">{targetFile.filename} · {targetFile.rows.length} rows</p>
                          <button 
                            onClick={() => { setTargetFile(emptyFile()); setHasCompared(false); }}
                            className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 hover:bg-red-400/5 transition-all cursor-pointer border border-transparent hover:border-red-400/20"
                          >
                            <Trash2 className="h-3 w-3" />
                            Clear Target File
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              ) : bothComparatorFilesLoaded && !comparison ? (
                /* Both files loaded — show Compare CTA */
                <div className="flex-1 flex flex-col items-center justify-center gap-6">
                  <div className="glass-card p-10 text-center max-w-md space-y-5">
                    <div className="flex items-center justify-center gap-3 mb-2">
                      <div className="h-2 w-2 rounded-full bg-blue-400" />
                      <span className="text-xs text-white/40 font-mono truncate max-w-[150px]" title={refFile.filename}>{refFile.filename}</span>
                      <ArrowRightLeft className="h-4 w-4 text-white/20" />
                      <span className="text-xs text-white/40 font-mono truncate max-w-[150px]" title={targetFile.filename}>{targetFile.filename}</span>
                      <div className="h-2 w-2 rounded-full bg-emerald-400" />
                    </div>
                    <div className="space-y-4">
                      <div className="p-4 rounded-2xl bg-emerald-500/10 border border-emerald-500/20 shadow-glow">
                        <p className="text-sm text-emerald-400 font-bold">Both BOMs Loaded Successfully</p>
                        <p className="text-xs text-emerald-400/60 mt-1 font-mono tracking-tight">{refFile.rows.length} Ref Rows ➔ {targetFile.rows.length} Target Rows</p>
                      </div>
                      
                      <button
                        onClick={handleStartComparison}
                        className="w-full relative group/btn"
                      >
                        <div className="absolute -inset-1 bg-gradient-to-r from-emerald-500 to-blue-500 rounded-2xl blur opacity-25 group-hover/btn:opacity-50 transition duration-1000 group-hover/btn:duration-200" />
                        <div className="relative glass-button w-full px-8 py-4 rounded-2xl flex items-center justify-center gap-3 bg-slate-900 border border-emerald-500/30 group-hover/btn:border-emerald-500/50 transition-all">
                          <CheckCircle2 className="h-5 w-5 text-emerald-400" />
                          <span className="text-sm font-black uppercase tracking-[0.2em] text-white">Compare & Analyze Now</span>
                          <ChevronRight className="h-4 w-4 text-white/40 group-hover/btn:translate-x-1 transition-transform" />
                        </div>
                      </button>
                    </div>
                    <p className="text-[10px] text-white/40 font-bold uppercase tracking-wider">Or change your files below before comparing</p>
                    <div className="flex gap-4 justify-center">
                      <div className="flex flex-col items-center gap-1">
                        <UploadZone
                          id="ref-upload-2"
                          onUpload={handleRefUpload}
                          label="Change Ref"
                          isSmall
                        />
                        <button 
                          onClick={() => { setRefFile(emptyFile()); setHasCompared(false); }}
                          className="text-[9px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 transition-all cursor-pointer"
                        >
                          Clear
                        </button>
                      </div>
                      
                      <div className="w-[1px] h-6 bg-white/5 mx-2" />
                      <div className="flex flex-col items-center gap-1">
                        <UploadZone
                          id="target-upload-2"
                          onUpload={handleTargetUpload}
                          label="Change Target"
                          isSmall
                        />
                        <button 
                          onClick={() => { setTargetFile(emptyFile()); setHasCompared(false); }}
                          className="text-[9px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 transition-all cursor-pointer"
                        >
                          Clear
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <>
                  <div className="glass-card px-4 py-3 shrink-0 flex flex-col gap-3">
                    <div className="flex items-center gap-3">
                      <button
                        onClick={() => runComparatorTool('normalize')}
                        title="Clean designators across both Reference and Target BOMs"
                        className="flex shrink-0 items-center gap-2 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest text-white/40 border border-white/5 hover:bg-white/5 hover:text-white hover:border-white/20 transition-all duration-200 cursor-pointer shadow-lg"
                      >
                        <AlignJustify className="h-3.5 w-3.5" />
                        Normalize
                      </button>

                      <button
                        onClick={() => runComparatorTool('expand')}
                        title="Expand all designator ranges for accurate comparison"
                        className="flex shrink-0 items-center gap-2 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest text-white/40 border border-white/5 hover:bg-white/5 hover:text-white hover:border-white/20 transition-all duration-200 cursor-pointer shadow-lg"
                      >
                        <ArrowRightLeft className="h-3.5 w-3.5" />
                        Expand
                      </button>

                      <button
                        onClick={() => runComparatorTool('validate')}
                        title="Check both BOMs for duplicate designators and quantity mismatches"
                        className="flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold text-emerald-400 border border-emerald-500/20 bg-emerald-500/10 hover:bg-emerald-500/20 hover:border-emerald-500/40 transition-all duration-200 cursor-pointer shadow-lg shrink-0"
                      >
                        <ShieldCheck className="h-4 w-4" />
                        Validate
                      </button>

                      <button
                        onClick={() => setShowSearch(true)}
                        title="Search within current database mode"
                        className="flex shrink-0 items-center gap-2 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest whitespace-nowrap transition-all duration-300 cursor-pointer border border-primary/20 bg-primary/5 text-primary hover:bg-primary/10 hover:border-primary/40 shadow-sm ml-4"
                      >
                        <Search className="h-3.5 w-3.5" />
                        Search
                      </button>

                      <div className="flex-1" />

                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          onClick={() => handleSave(refFile, 'reference_bom')}
                          title="Export the Reference BOM with your edits"
                          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-blue-300 border border-blue-400/30 bg-blue-500/20 hover:bg-blue-500/30 transition-all cursor-pointer whitespace-nowrap"
                        >
                          <Save className="h-3.5 w-3.5" /> Save Ref
                        </button>
                        <button
                          onClick={() => handleSave(targetFile, 'target_bom')}
                          title="Export the Target BOM with your edits"
                          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-emerald-300 border border-emerald-400/30 bg-emerald-500/20 hover:bg-emerald-500/30 transition-all cursor-pointer whitespace-nowrap"
                        >
                          <Save className="h-3.5 w-3.5" /> Save Target
                        </button>

                        <button
                          onClick={() => { setRefFile(emptyFile()); setTargetFile(emptyFile()); setHasCompared(false); }}
                          title="Clear both files and reset the comparator"
                          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-white/40 hover:text-red-400 hover:bg-red-500/10 border border-transparent hover:border-red-400/30 transition-all cursor-pointer whitespace-nowrap"
                        >
                          <Trash2 className="h-3.5 w-3.5" /> Clear
                        </button>
                      </div>
                    </div>

                    {/* Line 2: Status Legend */}
                    <div className="flex items-center gap-4 px-3 py-1.5 shrink-0 bg-white/5 rounded-xl border border-white/5">
                      <div className="flex items-center gap-4 border-r border-white/10 pr-4">
                        {[
                          { dot: 'bg-emerald-500', label: 'OK' },
                          { dot: 'bg-slate-400', label: 'Auto-Filled' },
                          { dot: 'bg-orange-500', label: 'DUP (SELF)' },
                          { dot: 'bg-amber-400', label: 'QTY ERR' },
                          { dot: 'bg-rose-500', label: 'Both' },
                        ].map(s => (
                          <div key={s.label} className="flex shrink-0 whitespace-nowrap items-center gap-1.5">
                            <div className={`h-2 w-2 rounded-full ${s.dot}`} />
                            <span className="text-[10px] text-white/70 font-bold uppercase tracking-widest">{s.label}</span>
                          </div>
                        ))}
                      </div>
                      <div className="flex items-center gap-4">
                        {[
                          { dot: 'bg-pink-500', label: 'Missing' },
                          { dot: 'bg-teal-400', label: 'Extra' },
                          { dot: 'bg-sky-400', label: 'Δ Qty' },
                          { dot: 'bg-indigo-400', label: 'Swapped' },
                        ].map(s => (
                          <div key={s.label} className="flex shrink-0 whitespace-nowrap items-center gap-1.5 opacity-80">
                            <div className={`h-2 w-2 rounded-full ${s.dot} shadow-[0_0_8px_rgba(255,255,255,0.1)]`} />
                            <span className="text-[10px] text-white/90 font-black uppercase tracking-widest italic">{s.label}</span>
                          </div>
                        ))}
                      </div>
                      <div className="flex-1 min-w-[24px]" />
                      {(() => {
                        const totalIssues = comparison 
                          ? (comparison.missing.length + 
                             comparison.extra.length + 
                             comparison.qtyMismatch.length + 
                             comparison.swapped.length + 
                             comparison.duplicates.source.length + 
                             comparison.duplicates.target.length + 
                             comparison.internalQtyErrors.source.length + 
                             comparison.internalQtyErrors.target.length)
                          : (refFile.rows.filter(r => r.duplicate).length + 
                             refFile.rows.filter(r => r.qtyMismatch).length +
                             targetFile.rows.filter(r => r.duplicate).length + 
                             targetFile.rows.filter(r => r.qtyMismatch).length);
                        
                        return (
                          <span className={`text-[10px] font-black uppercase tracking-widest italic shrink-0 whitespace-nowrap transition-colors duration-300 ${
                            totalIssues > 0 ? 'text-rose-500' : 'text-emerald-400'
                          }`}>
                            {totalIssues} Total issues found
                          </span>
                        );
                      })()}
                    </div>
                  </div>

                    <div className={`flex-1 min-h-0 grid gap-3 ${comparison ? 'grid-cols-[minmax(0,1fr)_minmax(0,1fr)_340px]' : 'grid-cols-2'}`} id="comparator-layout">
                      {/* Reference Master Grid */}
                      <div className="flex flex-col min-h-0 min-w-0" id="ref-grid-container">
                        <EditableBOMGrid 
                          type="ref"
                          label={`REFERENCE — ${refFile.filename}`} 
                          data={comparisonResult?.source || refFile.rows}
                          filters={refFilters}
                          onFilterChange={(col, sel) => setRefFilters(prev => ({ ...prev, [col]: sel }))}
                          onUpdate={(rowIdx, f, v) => updateRow('ref', rowIdx, f, v)}
                          canUndo={refPointer > 0}
                          canRedo={refPointer < refHistory.length - 1}
                          onUndo={() => undo('ref')}
                          onRedo={() => redo('ref')}
                          onRemap={() => setRefPending({ 
                            rawData: refFile.rawData, 
                            rawAOA: refFile.rawAOA || [], 
                            headerRowIdx: refFile.headerRowIdx,
                            mapping: refFile.columnMapping, 
                            availableColumns: refFile.availableColumns || [],
                            filename: refFile.filename,
                            buffer: refFile.buffer,
                            sheetName: refFile.sheetName,
                            merges: refFile.merges
                          })}
                          onScrollChange={setRefScroll}
                          initialScroll={refScroll}
                          jumpTarget={activeTab === 'comparator' && jumpTarget?.source === 'ref' ? jumpTarget : null}
                          highlightLogic={(row) => {
                            if (row.missing) return 'bg-pink-500/10 border-l-2 border-l-pink-500 shadow-[inset_4px_0_12px_-4px_rgba(236,72,153,0.2)] color-pink';
                            if (row.swappedResult) return 'bg-indigo-500/5 border-l-2 border-l-indigo-400 color-indigo';
                            if (row.qtyMismatchResult) return 'bg-sky-500/5 border-l-2 border-l-sky-400 color-sky';
                            return '';
                          }}
                        />
                      </div>

                      {/* Target Grid */}
                      <div className="flex flex-col min-h-0 min-w-0" id="target-grid-container">
                        <EditableBOMGrid 
                          type="target"
                          label={`TARGET — ${targetFile.filename}`}
                          data={comparisonResult?.target || targetFile.rows}
                          filters={targetFilters}
                          onFilterChange={(col, sel) => setTargetFilters(prev => ({ ...prev, [col]: sel }))}
                          onUpdate={(rowIdx, f, v) => updateRow('target', rowIdx, f, v)}
                          canUndo={targetPointer > 0}
                          canRedo={targetPointer < targetHistory.length - 1}
                          onUndo={() => undo('target')}
                          onRedo={() => redo('target')}
                          onRemap={() => setTargetPending({ 
                            rawData: targetFile.rawData, 
                            rawAOA: targetFile.rawAOA || [], 
                            headerRowIdx: targetFile.headerRowIdx,
                            mapping: targetFile.columnMapping, 
                            availableColumns: targetFile.availableColumns || [],
                            filename: targetFile.filename,
                            buffer: targetFile.buffer,
                            sheetName: targetFile.sheetName,
                            merges: targetFile.merges
                          })}
                          onSyncFromMaster={handleSyncFromMaster}
                          onScrollChange={setTargetScroll}
                          initialScroll={targetScroll}
                          jumpTarget={activeTab === 'comparator' && jumpTarget?.source === 'target' ? jumpTarget : null}
                          highlightLogic={(row) => {
                            if (row.extra) return 'bg-teal-500/10 border-l-2 border-l-teal-400 shadow-[inset_4px_0_12px_-4px_rgba(20,184,166,0.2)] color-teal';
                            if (row.qtyMismatchResult) return 'bg-sky-500/5 border-l-2 border-l-sky-400 color-sky';
                            if (row.swappedResult) return 'bg-indigo-500/5 border-l-2 border-l-indigo-400 color-indigo';
                            return '';
                          }}
                        />
                      </div>

                      {comparison && (
                        <aside className="w-[340px] flex-shrink-0 flex flex-col min-h-0 shadow-2xl">
                          <ComparisonResults 
                            diff={comparison} 
                            onJump={(rowIdx, source) => setJumpTarget({ rowIdx, source, timestamp: Date.now() })}
                          />
                        </aside>
                      )}
                    </div>
                  </>
                )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Sheet Picker — appears first for multi-sheet files */}
      <AnimatePresence>
        {sheetPicker && (
          <SheetPicker
            key="sheet-picker"
            state={sheetPicker}
            onCancel={() => setSheetPicker(null)}
          />
        )}
      </AnimatePresence>

      {/* Column Mapping Dialogs */}
      <AnimatePresence>
        {analyzerPending && (
          <ColumnMapDialog
            key="analyzer-map"
            columns={analyzerPending.availableColumns}
            mapping={analyzerPending.mapping}
            onConfirm={confirmAnalyzerMapping}
            onCancel={() => setAnalyzerPending(null)}
          />
        )}
        {refPending && (
          <ColumnMapDialog
            key="ref-map"
            columns={refPending.availableColumns}
            mapping={refPending.mapping}
            onConfirm={confirmRefMapping}
            onCancel={() => setRefPending(null)}
          />
        )}
        {targetPending && (
          <ColumnMapDialog
            key="target-map"
            columns={targetPending.availableColumns}
            mapping={targetPending.mapping}
            onConfirm={confirmTargetMapping}
            onCancel={() => setTargetPending(null)}
          />
        )}
        {showAbout && (
          <AboutDialog key="about-modal" onClose={() => setShowAbout(false)} />
        )}
        {showSearch && (
          <ModeSearch 
            onClose={() => setShowSearch(false)} 
            dataSources={searchSources}
            onResultClick={handleSearchResultClick}
            title={searchTitle}
          />
        )}
      </AnimatePresence>
    </div>
  </div>
  );
}

// ── Tab Button ────────────────────────────────────────────────────────────────
export function TabButton({ active, onClick, icon, label }: { active: boolean; onClick: () => void; icon: React.ReactNode; label: string }) {
  return (
    <button
      onClick={onClick}
      className={`flex shrink-0 whitespace-nowrap items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-extrabold transition-all duration-300 cursor-pointer
        ${active
          ? 'bg-white/10 text-white shadow-[0_0_15px_rgba(255,255,255,0.05)]'
          : 'text-white/35 hover:text-white/60 hover:bg-white/5'
        }`}
    >
      {icon}
      {label}
    </button>
  );
}
