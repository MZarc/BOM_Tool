import React, { useState, useMemo, useEffect, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { 
  CheckCircle2, AlertCircle, 
  ArrowRightLeft, MapPin, Search, Trash2, X, Info, ShieldCheck, Database
} from 'lucide-react';
import { AnimatePresence } from 'framer-motion';

import { PnPParser, PnPRow } from '@/lib/bom/pnp-parser';
import { BOMProcessor, BOMRow, ColumnMapping } from '@/lib/bom/bom-processor';
import { UploadZone, SheetPicker, SheetPickerState } from './bom-dashboard';
import { ColumnMapDialog, PendingMapping } from './ui/MappingDialogs';
import { FilterDropdown } from './ui/FilterDropdown';

interface ValidationResult {
  designator: string;
  pnpDesc: string;
  bomDesc: string;
  status: 'match' | 'mismatch' | 'not_in_bom' | 'not_in_pnp';
  details?: string;
  rowIdx?: number;
}

export interface PnPFileData<T> {
  filename: string;
  rows: T[];
  source?: PendingMapping;
}

interface PnPValidatorProps {
  pnpFile: PnPFileData<PnPRow> | null;
  setPnpFile: (file: PnPFileData<PnPRow> | null) => void;
  bomFile: PnPFileData<BOMRow> | null;
  setBomFile: (file: PnPFileData<BOMRow> | null) => void;
  hasValidated: boolean;
  setHasValidated: (v: boolean) => void;
  onSearch?: () => void;
  jumpTarget?: { rowIdx: number; timestamp: number } | null;
  initialScroll?: number;
  onScrollChange?: (top: number) => void;
}

export function PnPValidator({
  pnpFile, setPnpFile,
  bomFile, setBomFile,
  hasValidated, setHasValidated,
  onSearch, jumpTarget,
  initialScroll, onScrollChange
}: PnPValidatorProps) {
  
  const [searchQuery, setSearchQuery] = useState('');
  const [filter, setFilter] = useState<'all' | 'mismatch' | 'missing'>('all');
  const [statusFilters, setStatusFilters] = useState<Set<string>>(new Set());
  const [openFilter, setOpenFilter] = useState<boolean>(false);
  const tableRef = useRef<HTMLDivElement>(null);

  // Sheet Picker Status
  const [sheetPicker, setSheetPicker] = useState<SheetPickerState | null>(null);
  const [pnpPending, setPnpPending] = useState<PendingMapping | null>(null);
  const [bomPending, setBomPending] = useState<PendingMapping | null>(null);

  const applyPnpMapping = (data: PendingMapping, m: ColumnMapping) => {
    const rows = data.rawData.map(row => {
      const designator = m.designators ? String(row[m.designators] || '') : '';
      const description = m.description ? String(row[m.description] || '') : '';
      return { designator, description };
    }).filter(r => r.designator);

    setPnpFile({ 
      filename: data.filename, 
      rows, 
      source: { ...data, mapping: m } 
    });
    setHasValidated(false);
    setPnpPending(null);
  };

  const handlePnpMapping = (m: ColumnMapping) => {
    if (!pnpPending) return;
    applyPnpMapping(pnpPending, m);
  };

  const handlePnpUpload = async (files: FileList) => {
    const file = files[0];
    if (!file) return;

    try {
      if (file.name.toLowerCase().endsWith('.txt') || file.name.toLowerCase().endsWith('.csv')) {
        const text = await file.text();
        const rows = PnPParser.parseAltium(text);
        setPnpFile({ filename: file.name, rows });
        setHasValidated(false);
      } else {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        const sheetNames = wb.SheetNames;

        const processSheet = (name: string) => {
          const ws = wb.Sheets[name];
          const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][];
          const headerIdx = BOMProcessor.findHeaderRow(sheetData);
          const headers = (sheetData[headerIdx] || []) as string[];
          const initialMapping = BOMProcessor.autoMapHeaders(headers);
          const rawData = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: '' }) as Record<string, unknown>[];
          
          const pendingData: PendingMapping = {
            rawData,
            rawAOA: sheetData,
            headerRowIdx: headerIdx,
            mapping: initialMapping,
            availableColumns: headers,
            filename: file.name,
            merges: ws['!merges'] || [],
            buffer: buffer,
            sheetName: name
          };

          // If we found designators, we can auto-apply!
          if (initialMapping.designators) {
            applyPnpMapping(pendingData, initialMapping);
          } else {
            setPnpPending(pendingData);
          }
        };

        if (sheetNames.length > 1) {
          setSheetPicker({
            sheetNames,
            filename: file.name,
            workbookBinary: '',
            onSheetSelected: (name) => {
              processSheet(name);
              setSheetPicker(null);
            }
          });
        } else if (sheetNames.length === 1) {
          processSheet(sheetNames[0]);
        }
      }
    } catch (err) {
      console.error("PnP Upload Error:", err);
    }
  };

  const applyBomMapping = (data: PendingMapping, m: ColumnMapping) => {
    const res = BOMProcessor.applyMapping(data.rawData, m, data.headerRowIdx);
    setBomFile({ 
      filename: data.filename, 
      rows: res, 
      source: { ...data, mapping: m } 
    });
    setHasValidated(false);
    setBomPending(null);
  };

  const handleBomMapping = (m: ColumnMapping) => {
    if (!bomPending) return;
    applyBomMapping(bomPending, m);
  };

  const handleBomUpload = async (files: FileList) => {
    const file = files[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
      const sheetNames = wb.SheetNames;

      const processSheet = (name: string) => {
        const ws = wb.Sheets[name];
        const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][];
        const headerIdx = BOMProcessor.findHeaderRow(sheetData);
        const headers = (sheetData[headerIdx] || []) as string[];
        const initialMapping = BOMProcessor.autoMapHeaders(headers);
        const rawData = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: '' }) as Record<string, unknown>[];

        const pendingData: PendingMapping = {
          rawData,
          rawAOA: sheetData,
          headerRowIdx: headerIdx,
          mapping: initialMapping,
          availableColumns: headers,
          filename: file.name,
          merges: ws['!merges'] || [],
          buffer: buffer,
          sheetName: name
        };

        // For BOM, we need MPN, QTY, and Designators
        if (initialMapping.mpn && initialMapping.qty && initialMapping.designators) {
          applyBomMapping(pendingData, initialMapping);
        } else {
          setBomPending(pendingData);
        }
      };

      if (sheetNames.length > 1) {
        setSheetPicker({
          sheetNames,
          filename: file.name,
          workbookBinary: '',
          onSheetSelected: (name) => {
            processSheet(name);
            setSheetPicker(null);
          }
        });
      } else if (sheetNames.length === 1) {
        processSheet(sheetNames[0]);
      }
    } catch (err) {
      console.error("BOM Upload Error:", err);
    }
  };

  const clearAll = () => {
    setPnpFile(null);
    setBomFile(null);
    setHasValidated(false);
  };

  // --- Comparison Logic ---
  const results = useMemo(() => {
    if (!pnpFile || !bomFile || !hasValidated) return [];

    const statsMap = new Map<string, ValidationResult>();

    // 1. Index BOM by Designator
    const bomMap = new Map<string, { desc: string; row: number }>();
    bomFile.rows.forEach(r => {
      r.designators.forEach(d => {
        bomMap.set(d.trim().toUpperCase(), { 
          desc: r.description || '', 
          row: r.originalRowIdx 
        });
      });
    });

    // Check if we should skip description check (if PnP has no description mapped or all empty)
    const pnpHasAnyDesc = pnpFile.rows.some(r => r.description && r.description.trim().length > 0);
    const skipDescCheck = !pnpHasAnyDesc;

    // 2. Scan PnP and compare to BOM
    const processedDesignators = new Set<string>();
    pnpFile.rows.forEach(p => {
      const d = p.designator.trim().toUpperCase();
      if (!d) return;
      processedDesignators.add(d);
      
      const bomData = bomMap.get(d);
      const pnpDesc = p.description || '';
      const bomDesc = bomData?.desc || '';

      let status: ValidationResult['status'] = 'not_in_bom';
      if (bomData) {
        if (skipDescCheck) {
          status = 'match'; // If no description in PnP, we only care if it exists in BOM
        } else {
          status = pnpDesc.trim().toLowerCase() === bomDesc.trim().toLowerCase() ? 'match' : 'mismatch';
        }
      }

      statsMap.set(d, {
        designator: p.designator,
        pnpDesc,
        bomDesc,
        status,
        rowIdx: bomData?.row
      });
    });

    // 3. Scan BOM for things missing in PnP
    bomMap.forEach((data, d) => {
      if (!processedDesignators.has(d)) {
        statsMap.set(d, {
          designator: d,
          pnpDesc: '',
          bomDesc: data.desc,
          status: 'not_in_pnp',
          rowIdx: data.row
        });
      }
    });

    return Array.from(statsMap.values()).sort((a,b) => a.designator.localeCompare(b.designator, undefined, { numeric: true }));
  }, [pnpFile, bomFile, hasValidated]);

  // Handle Jumping and Highlighting in results
  const [highlightedRow, setHighlightedRow] = useState<number | null>(null);

  useEffect(() => {
    if (!jumpTarget || !hasValidated || !tableRef.current) return;

    const raf = requestAnimationFrame(() => {
      const container = tableRef.current!;
      const rowElem = container.querySelector<HTMLElement>(`#row-${jumpTarget.rowIdx}`);
      if (!rowElem) return;

      // Scroll the overflow container so the row is centered in the visible area
      const containerRect = container.getBoundingClientRect();
      const rowRect = rowElem.getBoundingClientRect();
      const offset = rowRect.top - containerRect.top + container.scrollTop;
      const center = offset - container.clientHeight / 2 + rowRect.height / 2;
      container.scrollTo({ top: Math.max(0, center), behavior: 'smooth' });

      setHighlightedRow(jumpTarget.rowIdx);
      const timer = setTimeout(() => setHighlightedRow(null), 5000);
      (rowElem as HTMLElement & { _hlTimer?: ReturnType<typeof setTimeout> })._hlTimer = timer;
    });

    return () => cancelAnimationFrame(raf);
  }, [jumpTarget, hasValidated]);

  const hasRestored = useRef(false);
  const lastReportedScroll = useRef<number>(0);
  const scrollTimeout = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Restore scroll on mount
  useEffect(() => {
    if (tableRef.current && initialScroll !== undefined && !hasRestored.current) {
      tableRef.current.scrollTop = initialScroll;
      lastReportedScroll.current = initialScroll;
      hasRestored.current = true;
    }
  }, [initialScroll]); // Dependency added for lint, Ref ensures 'once' logic

  // Throttled scroll tracking
  const handleScroll = useCallback(() => {
    if (tableRef.current) {
      const top = tableRef.current.scrollTop;
      if (Math.abs(top - lastReportedScroll.current) > 50) {
        lastReportedScroll.current = top;
        onScrollChange?.(top);
      } else {
        if (scrollTimeout.current) clearTimeout(scrollTimeout.current);
        scrollTimeout.current = setTimeout(() => {
          if (tableRef.current) onScrollChange?.(tableRef.current.scrollTop);
        }, 150);
      }
    }
  }, [onScrollChange]);

  const filteredResults = useMemo(() => {
    return results.filter(r => {
      // 1. Keyword search
      const matchesSearch = r.designator.toLowerCase().includes(searchQuery.toLowerCase()) ||
                          r.pnpDesc.toLowerCase().includes(searchQuery.toLowerCase()) ||
                          r.bomDesc.toLowerCase().includes(searchQuery.toLowerCase());
      
      // 2. Old-style toggle filters
      const matchesFilter = 
        filter === 'all' ? true :
        filter === 'mismatch' ? r.status === 'mismatch' :
        r.status === 'not_in_bom' || r.status === 'not_in_pnp';
      
      // 3. New-style status header multi-select filter
      const matchesStatusHeader = statusFilters.size === 0 || statusFilters.has(r.status);

      return matchesSearch && matchesFilter && matchesStatusHeader;
    });
  }, [results, searchQuery, filter, statusFilters]);

  // Unique status values for filter dropdown
  const uniqueStatuses = useMemo(() => {
    return ['match', 'mismatch', 'not_in_bom', 'not_in_pnp'];
  }, []);

  const stats = useMemo(() => {
    return {
      total: results.length,
      matches: results.filter(r => r.status === 'match').length,
      mismatches: results.filter(r => r.status === 'mismatch').length,
      missing: results.filter(r => r.status === 'not_in_bom' || r.status === 'not_in_pnp').length
    };
  }, [results]);

  // --- Rendering ---
  return (
    <div className="flex-1 flex flex-col gap-4 overflow-hidden">
      {(!pnpFile || !bomFile) ? (
        /* Initialization State */
        <div className="flex-1 flex flex-col gap-4">
           <div className="flex-1 grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="glass-card p-10 flex flex-col items-center justify-center space-y-6">
                <div className="text-center space-y-2">
                  <h3 className="text-2xl font-bold tracking-tight bg-gradient-to-r from-white via-blue-200 to-white bg-clip-text text-transparent">
                    Pick & Place File
                  </h3>
                  <p className="text-white/30 text-xs uppercase tracking-widest font-medium">Placement Data Source</p>
                </div>
                <UploadZone 
                  id="pnp-upload" 
                  onUpload={handlePnpUpload} 
                  label={pnpFile ? 'Change File' : 'Upload PnP Data'} 
                  icon={<MapPin className="h-4 w-4" />}
                  hasFile={!!pnpFile}
                  accept=".txt,.csv,.xlsx,.xls"
                />
                {pnpFile && (
                  <div className="flex flex-col items-center gap-3">
                    <p className="text-xs text-emerald-400/60 font-mono">{pnpFile.filename} · {pnpFile.rows.length} rows</p>
                    <div className="flex items-center gap-2">
                      {pnpFile.source && (
                        <button 
                          onClick={() => setPnpPending(pnpFile.source || null)}
                          className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest text-emerald-400/60 hover:text-emerald-400 hover:bg-emerald-400/5 transition-all cursor-pointer border border-transparent hover:border-emerald-400/20"
                        >
                          <Database className="h-3 w-3" />
                          Remap
                        </button>
                      )}
                      <button 
                        onClick={() => setPnpFile(null)}
                        className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 hover:bg-red-400/5 transition-all cursor-pointer border border-transparent hover:border-red-400/20"
                      >
                        <Trash2 className="h-3 w-3" />
                        Clear
                      </button>
                    </div>
                  </div>
                )}
              </div>

              <div className="glass-card p-10 flex flex-col items-center justify-center space-y-6">
                <div className="text-center space-y-2">
                  <h3 className="text-2xl font-bold tracking-tight bg-gradient-to-r from-white via-emerald-200 to-white bg-clip-text text-transparent">
                    Master BOM
                  </h3>
                  <p className="text-white/30 text-xs uppercase tracking-widest font-medium">Standard Reference File</p>
                </div>
                <UploadZone 
                  id="bom-upload" 
                  onUpload={handleBomUpload} 
                  label={bomFile ? 'Change File' : 'Upload Master BOM'} 
                  hasFile={!!bomFile}
                />
                {bomFile && (
                  <div className="flex flex-col items-center gap-3">
                    <p className="text-xs text-emerald-400/60 font-mono">{bomFile.filename} · {bomFile.rows.length} rows</p>
                    <div className="flex items-center gap-2">
                      {bomFile.source && (
                        <button 
                          onClick={() => setBomPending(bomFile.source || null)}
                          className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest text-emerald-400/60 hover:text-emerald-400 hover:bg-emerald-400/5 transition-all cursor-pointer border border-transparent hover:border-emerald-400/20"
                        >
                          <Database className="h-3 w-3" />
                          Remap
                        </button>
                      )}
                      <button 
                        onClick={() => setBomFile(null)}
                        className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 hover:bg-red-400/5 transition-all cursor-pointer border border-transparent hover:border-red-400/20"
                      >
                        <Trash2 className="h-3 w-3" />
                        Clear
                      </button>
                    </div>
                  </div>
                )}
              </div>
           </div>
        </div>
      ) : (pnpFile && bomFile && !hasValidated) ? (
        <div className="flex-1 flex flex-col items-center justify-center gap-6">
          <div className="glass-card p-10 text-center max-w-md space-y-5">
            <div className="flex items-center justify-center gap-3 mb-2">
              <div className="h-2 w-2 rounded-full bg-blue-400" />
              <span className="text-xs text-white/40 font-mono break-all whitespace-normal max-w-[150px] leading-tight" title={pnpFile.filename}>{pnpFile.filename}</span>
              <ArrowRightLeft className="h-4 w-4 text-white/20" />
              <span className="text-xs text-white/40 font-mono break-all whitespace-normal max-w-[150px] leading-tight" title={bomFile.filename}>{bomFile.filename}</span>
              <div className="h-2 w-2 rounded-full bg-emerald-400" />
            </div>
            <div>
              <p className="text-sm text-white font-semibold">Source files locked — ready for audit</p>
              <p className="text-xs text-white/30 mt-1">{pnpFile.rows.length} placement rows · {bomFile.rows.length} BOM rows</p>
            </div>
            <button
              onClick={() => setHasValidated(true)}
              className="glass-button px-8 py-3 text-sm font-bold flex items-center gap-2 mx-auto cursor-pointer"
            >
              <CheckCircle2 className="h-4 w-4 text-emerald-400" />
              Start Audit Now
            </button>
            <p className="text-[10px] text-white/40 font-bold uppercase tracking-wider">Or change your files below before audits</p>
            
            <div className="flex gap-6 justify-center">
              <div className="flex flex-col items-center gap-1">
                <UploadZone 
                  id="pnp-change" 
                  onUpload={handlePnpUpload} 
                  label="Change PnP" 
                  isSmall 
                  accept=".txt,.csv,.xlsx,.xls" 
                />
                <div className="flex items-center gap-2">
                  {pnpFile?.source && (
                    <button 
                      onClick={() => setPnpPending(pnpFile.source || null)}
                      className="text-[9px] font-black uppercase tracking-widest text-emerald-400/40 hover:text-emerald-400 transition-all cursor-pointer"
                    >
                      Remap
                    </button>
                  )}
                  <button 
                    onClick={() => { setPnpFile(null); setHasValidated(false); }}
                    className="text-[9px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 transition-all cursor-pointer"
                  >
                    Clear
                  </button>
                </div>
              </div>

              <div className="w-[1px] h-6 bg-white/5 mx-2" />

              <div className="flex flex-col items-center gap-1">
                <UploadZone 
                  id="bom-change" 
                  onUpload={handleBomUpload} 
                  label="Change BOM" 
                  isSmall 
                />
                <div className="flex items-center gap-2">
                  {bomFile?.source && (
                    <button 
                      onClick={() => setBomPending(bomFile.source || null)}
                      className="text-[9px] font-black uppercase tracking-widest text-emerald-400/40 hover:text-emerald-400 transition-all cursor-pointer"
                    >
                      Remap
                    </button>
                  )}
                  <button 
                    onClick={() => { setBomFile(null); setHasValidated(false); }}
                    className="text-[9px] font-black uppercase tracking-widest text-white/20 hover:text-red-400/60 transition-all cursor-pointer"
                  >
                    Clear
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : (
        /* Results State */
        <>
          {/* Header Controls */}
          <div className="glass-card p-4 flex flex-col md:flex-row items-center justify-between gap-4 border-white/5">
            <div className="flex items-center gap-4 flex-1">
              <div className="flex items-center gap-4 pr-4 border-r border-white/10">
                <div className="flex flex-col items-end">
                  <span className="text-[10px] text-white/30 font-black uppercase tracking-widest">PnP File</span>
                  <span className="text-xs text-blue-400 font-mono break-all whitespace-normal max-w-[120px] leading-tight text-right" title={pnpFile?.filename || ''}>{pnpFile?.filename || '—'}</span>
                </div>
                <ArrowRightLeft className="h-4 w-4 text-white/10" />
                <div className="flex flex-col">
                  <span className="text-[10px] text-white/30 font-black uppercase tracking-widest">BOM File</span>
                  <span className="text-xs text-emerald-400 font-mono break-all whitespace-normal max-w-[120px] leading-tight text-left" title={bomFile?.filename || ''}>{bomFile?.filename || '—'}</span>
                </div>
              </div>

              <div className="relative group max-w-sm flex-1" title="Search by designator or description">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-white/20 group-focus-within:text-blue-400 transition-colors" />
                <input 
                  type="text"
                  placeholder="Filter by Designator or Description..."
                  className="w-full bg-white/5 border border-white/5 rounded-xl pl-10 pr-4 py-2 text-xs font-bold text-white/40 placeholder:text-white/20 focus:text-white/60 focus:outline-none focus:border-blue-500/50 transition-all font-sans"
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                />
              </div>

              {onSearch && (
                <button
                  onClick={onSearch}
                  title="Search within current auditing mode"
                  className="flex items-center justify-center p-2 rounded-xl bg-white/5 border border-white/10 text-white/40 hover:text-white hover:bg-white/10 hover:border-white/20 transition-all cursor-pointer shadow-sm ml-1"
                >
                  <Search className="h-4 w-4" />
                </button>
              )}

              <div className="flex items-center gap-2 mx-2">
                {pnpFile?.source && (
                  <button 
                    onClick={() => setPnpPending(pnpFile.source || null)}
                    className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-blue-500/10 border border-blue-500/10 text-[9px] font-black uppercase tracking-widest text-blue-400 hover:text-blue-300 hover:bg-blue-500/20 transition-all cursor-pointer"
                    title="Remap PnP Columns"
                  >
                    <ArrowRightLeft className="h-3 w-3" />
                    Remap PnP
                  </button>
                )}
                {bomFile?.source && (
                  <button 
                    onClick={() => setBomPending(bomFile.source || null)}
                    className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-emerald-500/10 border border-emerald-500/10 text-[9px] font-black uppercase tracking-widest text-emerald-400 hover:text-emerald-300 hover:bg-emerald-500/20 transition-all cursor-pointer"
                    title="Remap BOM Columns"
                  >
                    <Database className="h-3 w-3" />
                    Remap BOM
                  </button>
                )}
              </div>
            </div>

            <div className="flex items-center gap-3">
              <div className="flex bg-white/5 p-1 rounded-xl border border-white/10" title="Filter validation results">
                <FilterBtn active={filter === 'all'} onClick={() => setFilter('all')} label="All" />
                <FilterBtn active={filter === 'mismatch'} onClick={() => setFilter('mismatch')} label="Mismatches" tooltip="Show only description mismatches" />
                <FilterBtn active={filter === 'missing'} onClick={() => setFilter('missing')} label="Missing" tooltip="Show items missing in either BOM or PnP" />
              </div>
              <button 
                onClick={() => {
                   clearAll();
                   setStatusFilters(new Set());
                }}
                title="Clear all files and reset validation"
                className="p-2.5 rounded-xl text-white/40 hover:text-red-400 hover:bg-red-400/10 border border-transparent hover:border-red-400/20 transition-all cursor-pointer"
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </div>
          </div>

          <div className="flex-1 flex gap-4 min-h-0">
            <div className="flex-1 glass-card border-white/5 flex flex-col overflow-hidden">
              <div 
                ref={tableRef} 
                onScroll={handleScroll}
                className="flex-1 overflow-y-auto bom-scroll-stable"
              >
                {/* Sticky Header */}
                <div className="grid grid-cols-[100px_1fr_1fr_150px] border-b border-white/30 bg-[#080d19] shadow-xl bom-thead shrink-0">
                  <span className="text-[10px] font-black uppercase tracking-widest text-white/50 py-3 text-center border-r border-white/30 truncate">Designator</span>
                  <span className="text-[10px] font-black uppercase tracking-widest text-white/50 py-3 text-center border-r border-white/30 truncate">PnP Description</span>
                  <span className="text-[10px] font-black uppercase tracking-widest text-white/50 py-3 text-center border-r border-white/30 truncate">BOM Description</span>
                  <div className="text-[10px] font-black uppercase tracking-widest text-white/50 py-3 text-center flex items-center justify-center gap-2 relative group">
                    Status
                    <FilterDropdown 
                        column="Status" 
                        values={uniqueStatuses} 
                        selectedValues={statusFilters} 
                        onFilterChange={setStatusFilters}
                        isOpen={openFilter}
                        onOpenChange={setOpenFilter}
                    />
                  </div>
                </div>
                {filteredResults.map((r) => (
                  <div 
                    key={r.designator} 
                    id={`row-${r.rowIdx}`}
                    className={`grid grid-cols-[100px_1fr_1fr_150px] border-b border-white/20 items-center transition-colors hover:bg-white/5 group [will-change:background-color] [backface-visibility:hidden] [transform:translateZ(0)] ${
                      highlightedRow === r.rowIdx ? 'bg-primary/20 ring-2 ring-primary/50 relative z-10' : ''
                    }`}
                  >
                    <span className="text-[11px] font-black font-mono text-white/90 text-center py-3 border-r border-white/20 h-full flex items-center justify-center shrink-0">{r.designator}</span>
                    <span className={`text-[10px] font-bold text-center py-3 border-r border-white/20 px-2 h-full flex items-center justify-center ${r.status === 'not_in_bom' || r.status === 'mismatch' ? 'text-orange-300' : 'text-white/30'}`}>
                      {r.pnpDesc || '—'}
                    </span>
                    <span className={`text-[10px] font-bold text-center py-3 border-r border-white/20 px-2 h-full flex items-center justify-center ${r.status === 'not_in_pnp' || r.status === 'mismatch' ? 'text-amber-300' : 'text-white/30'}`}>
                      {r.bomDesc || '—'}
                    </span>
                    <div className="flex items-center justify-center h-full">
                      <StatusBadge status={r.status} />
                    </div>
                  </div>
                ))}
                {filteredResults.length === 0 && (
                  <div className="p-20 flex flex-col items-center justify-center text-center opacity-30">
                    <Search className="h-10 w-10 mb-4" />
                    <p className="text-xs font-bold uppercase tracking-widest">No validation entries found</p>
                  </div>
                )}
              </div>
            </div>

            <aside className="w-80 flex flex-col gap-4">
              <div className="glass-card p-6 border-white/5 space-y-6 text-center shadow-inner bg-[#0a1224]">
                <div className="inline-flex items-center justify-center p-4 rounded-3xl bg-blue-500/10 border border-blue-500/20 mb-2">
                   <ShieldCheck className="h-6 w-6 text-blue-400" />
                </div>
                <div>
                   <h3 className="text-xl font-black text-white italic tracking-tighter">Validation Audit</h3>
                   <p className="text-[10px] font-black uppercase tracking-[0.2em] text-white/30">Master Sync Status</p>
                </div>
                
                <div className="grid grid-cols-1 gap-3 text-left pt-6 border-t border-white/5">
                    <StatRow label="Matches" value={stats.matches} color="text-emerald-400" icon={<CheckCircle2 className="h-4 w-4" />} />
                    <StatRow label="Mismatches" value={stats.mismatches} color="text-orange-400" icon={<AlertCircle className="h-4 w-4" />} />
                    <StatRow label="Missing in BOM" value={results.filter(r => r.status === 'not_in_bom').length} color="text-red-400" icon={<X className="h-4 w-4" />} />
                    <StatRow label="Missing in PnP" value={results.filter(r => r.status === 'not_in_pnp').length} color="text-blue-400" icon={<Info className="h-4 w-4" />} />
                </div>

                <div className="pt-6 border-t border-white/5">
                  <div className="flex items-center justify-between p-4 rounded-2xl bg-white/5 border border-white/5 shadow-inner">
                    <div className="flex flex-col text-left">
                      <span className="text-[10px] font-black uppercase text-white/30 tracking-widest">Accuracy</span>
                      <span className={`text-sm font-black tracking-tighter ${stats.mismatches === 0 && stats.missing === 0 ? 'text-emerald-400' : 'text-orange-400'}`}>
                        {stats.mismatches === 0 && stats.missing === 0 ? '100% Verified' : `${Math.round((stats.matches / results.length) * 100) || 0}% Correct`}
                      </span>
                    </div>
                    {stats.mismatches === 0 && stats.missing === 0 ? (
                      <CheckCircle2 className="h-6 w-6 text-emerald-400 shadow-[0_0_15px_rgba(52,211,153,0.3)]" />
                    ) : (
                      <AlertCircle className="h-6 w-6 text-orange-400" />
                    )}
                  </div>
                </div>
              </div>
            </aside>
          </div>
        </>
      )}

      {sheetPicker && <SheetPicker state={sheetPicker} onCancel={() => setSheetPicker(null)} />}

      <AnimatePresence mode="wait">
        {pnpPending && (
          <ColumnMapDialog 
            key="pnp-map-excel"
            title="Map PnP Columns"
            columns={pnpPending.availableColumns}
            mapping={pnpPending.mapping}
            requiredFields={['designators']}
            onConfirm={handlePnpMapping}
            onCancel={() => setPnpPending(null)}
          />
        )}
        {bomPending && (
          <ColumnMapDialog 
            key="bom-pnp-map-excel"
            title="Map Master BOM Columns"
            columns={bomPending.availableColumns}
            mapping={bomPending.mapping}
            requiredFields={['mpn', 'qty', 'designators']}
            onConfirm={handleBomMapping}
            onCancel={() => setBomPending(null)}
          />
        )}
      </AnimatePresence>
    </div>
  );
}

function FilterBtn({ active, onClick, label, tooltip }: { active: boolean; onClick: () => void; label: string; tooltip?: string }) {
  return (
    <button
      onClick={onClick}
      title={tooltip}
      className={`px-4 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all cursor-pointer
        ${active ? 'bg-white/10 text-white shadow-[0_0_10px_rgba(255,255,255,0.05)]' : 'text-white/30 hover:text-white/60'}`}
    >
      {label}
    </button>
  );
}

function StatRow({ label, value, color, icon }: { label: string; value: number; color: string; icon: React.ReactNode }) {
  return (
    <div className="flex items-center justify-between group">
      <div className="flex items-center gap-2">
        <div className={`${color} opacity-80 group-hover:scale-110 transition-transform`}>{icon}</div>
        <span className="text-[11px] font-bold text-white/50 group-hover:text-white transition-colors uppercase tracking-tight">{label}</span>
      </div>
      <span className={`text-xs font-black ${color}`}>{value}</span>
    </div>
  );
}

function StatusBadge({ status }: { status: ValidationResult['status'] }) {
  const configs = {
    match: { label: 'Match', color: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/20' },
    mismatch: { label: 'Mismatch', color: 'bg-orange-500/20 text-orange-400 border-orange-500/20' },
    not_in_bom: { label: 'Missing in BOM', color: 'bg-red-500/20 text-red-400 border-red-500/20' },
    not_in_pnp: { label: 'Missing in PnP', color: 'bg-blue-500/20 text-blue-400 border-blue-500/20' },
  };
  const config = configs[status];
  return (
    <span className={`px-2 py-0.5 rounded-md text-[8px] font-black uppercase tracking-[0.15em] border ${config.color} whitespace-nowrap`}>
      {config.label}
    </span>
  );
}
