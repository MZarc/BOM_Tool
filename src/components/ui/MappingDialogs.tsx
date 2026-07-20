import { useState, useRef, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { ChevronDown, CheckCircle2, Columns, X } from 'lucide-react';
import { ColumnMapping } from '@/lib/bom/bom-processor';
import * as XLSX from 'xlsx';

// ─── Shared Types ────────────────────────────────────────────────────────────
export interface PendingMapping {
  rawData: Record<string, unknown>[];
  rawAOA: unknown[][];
  headerRowIdx: number;
  mapping: ColumnMapping;
  availableColumns: string[];
  filename: string;
  merges: XLSX.Range[];
  buffer: ArrayBuffer;
  sheetName?: string;
}

// ─── Custom Dropdown ──────────────────────────────────────────────────────────
export function CustomDropdown({
  value,
  options,
  onChange,
  placeholder = "Select Column..."
}: {
  value: string | null;
  options: string[];
  onChange: (val: string | null) => void;
  placeholder?: string;
}) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div className="relative" ref={containerRef}>
      <button
        type="button"
        title={value || placeholder}
        onClick={() => setIsOpen(!isOpen)}
        className={`w-full flex items-center justify-between bg-white/[0.03] border rounded-xl px-4 py-3 text-sm font-bold text-white transition-all hover:bg-white/[0.07] cursor-pointer group/btn
          ${isOpen ? 'border-primary/50 ring-2 ring-primary/20' : 'border-white/10'}`}
      >
        <span className={!value ? 'text-white/30' : 'text-white'}>
          {value || placeholder}
        </span>
        <ChevronDown 
          className={`h-4 w-4 text-white/20 transition-transform duration-300 group-hover/btn:text-white/40 ${isOpen ? 'rotate-180 text-primary/60' : ''}`} 
        />
      </button>

      <AnimatePresence>
        {isOpen && (
          <motion.div
            initial={{ opacity: 0, y: -10, scale: 0.98 }}
            animate={{ opacity: 1, y: 5, scale: 1 }}
            exit={{ opacity: 0, y: -10, scale: 0.98 }}
            className="absolute z-[130] w-full mt-1 bg-slate-900 border border-white/10 rounded-xl overflow-hidden shadow-[0_20px_50px_rgba(0,0,0,0.5)]"
          >
            <div className="max-h-60 overflow-y-auto bom-scroll-stable py-1">
              <button
                type="button"
                onClick={() => { onChange(null); setIsOpen(false); }}
                className="w-full text-left px-4 py-2.5 text-sm font-bold text-white/30 hover:bg-white/5 transition-colors"
                title={`Clear ${placeholder}`}
              >
                {placeholder}
              </button>
              {options.map((opt) => (
                <button
                  key={opt}
                  type="button"
                  title={opt}
                  onClick={() => { onChange(opt); setIsOpen(false); }}
                  className={`w-full text-left px-4 py-2.5 text-sm font-bold transition-all flex items-center justify-between group/opt
                    ${(value as string | null) === opt 
                      ? 'bg-primary/20 text-primary-foreground border-l-2 border-primary' 
                      : 'text-white/70 hover:bg-white/5 hover:text-white'}`}
                >
                  <span>{opt}</span>
                  {(value as string | null) === opt && <CheckCircle2 className="h-3.5 w-3.5 text-primary" />}
                </button>
              ))}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

// ─── Column Map Dialog ────────────────────────────────────────────────────────
// ─── Column Map Dialog ────────────────────────────────────────────────────────
export function ColumnMapDialog({
  columns, 
  mapping: initialMapping, 
  onConfirm, 
  onCancel, 
  title = "Map Your BOM Columns",
  requiredFields = ['mpn', 'qty', 'designators']
}: {
  columns: string[];
  mapping: ColumnMapping;
  onConfirm: (m: ColumnMapping) => void;
  onCancel: () => void;
  title?: string;
  requiredFields?: (keyof ColumnMapping)[];
}) {
  const [mapping, setMapping] = useState<ColumnMapping>(initialMapping);

  const isValid = requiredFields.every(field => mapping[field] !== null);
  const isRequired = (field: keyof ColumnMapping) => requiredFields.includes(field);

  return (
    <div className="fixed inset-0 z-[150] flex items-center justify-center p-4">
      <div className="absolute inset-0 bg-black/80" />
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 30 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="relative z-10 glass-card p-8 w-full max-w-xl space-y-6 bg-[#060812] border border-white/10 shadow-2xl"
      >
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded-2xl bg-primary/10 flex items-center justify-center border border-primary/20">
              <Columns className="h-5 w-5 text-primary shadow-[0_0_15px_rgba(59,130,246,0.5)]" />
            </div>
            <div>
              <h2 className="text-xl font-bold tracking-tight text-white">{title}</h2>
              <p className="text-[10px] uppercase font-black tracking-widest text-white/30">Define column roles for precise data extraction</p>
            </div>
          </div>
          <button onClick={onCancel} className="p-2.5 rounded-xl text-white/20 hover:text-white hover:bg-white/5 transition-all cursor-pointer" title="Cancel mapping">
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="space-y-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-primary/60 flex items-center gap-1.5 ml-1">
              <div className="h-1 w-1 rounded-full bg-primary" /> MPN / Part Number {isRequired('mpn') && '*'}
            </label>
            <CustomDropdown 
              value={mapping.mpn as string | null} 
              options={columns} 
              onChange={val => setMapping(prev => ({ ...prev, mpn: val }))} 
              placeholder="Primary Part Identifier"
            />
          </div>
          <div className="space-y-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-emerald-400/60 flex items-center gap-1.5 ml-1">
              <div className="h-1 w-1 rounded-full bg-emerald-400" /> Quantity (Qty) {isRequired('qty') && '*'}
            </label>
            <CustomDropdown 
              value={mapping.qty as string | null} 
              options={columns} 
              onChange={val => setMapping(prev => ({ ...prev, qty: val }))} 
              placeholder="Total Assembly Quantity"
            />
          </div>
          <div className="space-y-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-blue-300/60 flex items-center gap-1.5 ml-1">
              <div className="h-1 w-1 rounded-full bg-blue-300" /> Designators {isRequired('designators') && '*'}
            </label>
            <CustomDropdown 
              value={mapping.designators as string | null} 
              options={columns} 
              onChange={val => setMapping(prev => ({ ...prev, designators: val }))} 
              placeholder="Reference Designator List"
            />
          </div>
          <div className="space-y-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-white/40 flex items-center gap-1.5 ml-1">
              <div className="h-1 w-1 rounded-full bg-white/40" /> Description / Value {isRequired('description') && '*'}
            </label>
            <CustomDropdown 
              value={mapping.description as string | null} 
              options={columns} 
              onChange={val => setMapping(prev => ({ ...prev, description: val }))} 
              placeholder="Component Details / Value"
            />
          </div>
        </div>

        <div className="flex gap-4 pt-4 border-t border-white/5">
          <button
            onClick={onCancel}
            className="flex-1 px-8 py-3 rounded-2xl bg-white/5 text-white/50 font-bold text-xs uppercase tracking-widest hover:bg-white/10 hover:text-white transition-all cursor-pointer"
            title="Cancel mapping process"
          >
            Cancel
          </button>
          <button
            onClick={() => onConfirm(mapping)}
            disabled={!isValid}
            className={`flex-1 px-8 py-3 rounded-2xl font-black text-xs uppercase tracking-widest transition-all shadow-xl
              ${isValid 
                ? 'bg-primary text-white hover:bg-primary/80 hover:scale-[1.02] active:scale-95 cursor-pointer' 
                : 'bg-white/5 text-white/20 cursor-not-allowed opacity-50'}`}
            title={isValid ? "Apply column mapping" : "Please map all required columns"}
          >
            Confirm Mapping
          </button>
        </div>
      </motion.div>
    </div>
  );
}
