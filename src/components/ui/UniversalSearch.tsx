import { useState, useMemo, useRef, useEffect } from 'react';
import { Search, X, FileText, Database, MapPin, ChevronRight, CornerDownRight, AlertCircle } from 'lucide-react';
import { motion } from 'framer-motion';
import { BOMRow } from '@/lib/bom/bom-processor';

interface SearchResult {
  id: string;
  source: 'analyzer' | 'reference' | 'target' | 'pnp';
  filename: string;
  mpn: string;
  qty: number;
  description: string;
  designators: string[];
  rowIdx: number;
}

export interface PnPSearchRow {
  designator: string;
  pnpDesc: string;
  bomDesc: string;
  status: string;
  rowIdx?: number;
}

interface UniversalSearchProps {
  onClose: () => void;
  dataSources: {
    analyzer?: { rows: BOMRow[], filename: string },
    reference?: { rows: BOMRow[], filename: string },
    target?: { rows: BOMRow[], filename: string },
    pnp?: { rows: PnPSearchRow[], filename: string },
  };
  onResultClick?: (result: SearchResult) => void;
  title: string;
}

export function ModeSearch({ onClose, dataSources, onResultClick, title }: UniversalSearchProps) {
  const [query, setQuery] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const [selectedIndex, setSelectedIndex] = useState(0);

  const results = useMemo(() => {
    if (query.length < 2) return [];

    const q = query.toLowerCase().trim();
    const allResults: SearchResult[] = [];

    const searchInSource = (rows: (BOMRow | PnPSearchRow)[], source: SearchResult['source'], filename: string) => {
      rows.forEach((row, idx) => {
        let matches = false;
        
        if (source === 'pnp') {
          // Special matching for PnP Validator results
          const r = row as PnPSearchRow;
          const matchesDes = r.designator.toLowerCase().includes(q);
          const matchesPnp = r.pnpDesc.toLowerCase().includes(q);
          const matchesBom = r.bomDesc.toLowerCase().includes(q);
          const matchesStat = r.status.toLowerCase().includes(q);
          matches = matchesDes || matchesPnp || matchesBom || matchesStat;

          if (matches) {
            allResults.push({
              id: `${source}-${idx}-${r.rowIdx || idx}`,
              source,
              filename,
              mpn: r.designator, // Use designator as MPN title for PnP
              qty: 1,
              description: `${r.status.toUpperCase()}: ${r.pnpDesc || r.bomDesc}`,
              designators: [r.designator],
              rowIdx: r.rowIdx || idx,
            });
          }
        } else {
          // Standard BOMRow matching
          const r = row as BOMRow;
          const matchesMpn = r.mpn.toLowerCase().includes(q);
          const matchesDesc = r.description?.toLowerCase().includes(q);
          const matchesDes = r.designators.some(d => d.toLowerCase().includes(q));
          matches = matchesMpn || matchesDesc || matchesDes;

          if (matches) {
            allResults.push({
              id: `${source}-${idx}-${r.originalRowIdx}`,
              source,
              filename,
              mpn: r.mpn,
              qty: r.qty,
              description: r.description || '',
              designators: r.designators,
              rowIdx: r.originalRowIdx,
            });
          }
        }
      });
    };

    if (dataSources.analyzer) searchInSource(dataSources.analyzer.rows, 'analyzer', dataSources.analyzer.filename);
    if (dataSources.reference) searchInSource(dataSources.reference.rows, 'reference', dataSources.reference.filename);
    if (dataSources.target) searchInSource(dataSources.target.rows, 'target', dataSources.target.filename);
    if (dataSources.pnp) searchInSource(dataSources.pnp.rows, 'pnp', dataSources.pnp.filename);

    return allResults.slice(0, 50); // Limit results for performance
  }, [query, dataSources]);

  // Reset selection index when query changes
  useEffect(() => {
    setSelectedIndex(0);
  }, [query]);

  // Keyboard Navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIndex(prev => (prev + 1) % Math.max(results.length, 1));
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIndex(prev => (prev - 1 + results.length) % Math.max(results.length, 1));
      }
      if (e.key === 'Enter' && results[selectedIndex]) {
        e.preventDefault();
        onResultClick?.(results[selectedIndex]);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [results, selectedIndex, onClose, onResultClick]);

  // Initial Focus
  useEffect(() => {
    const timer = setTimeout(() => {
      inputRef.current?.focus();
    }, 100);
    return () => clearTimeout(timer);
  }, []);

  const sourceConfig = {
    analyzer: { icon: <Database className="h-3 w-3 text-blue-400" />, label: 'Analyzer', color: 'bg-blue-500/10 text-blue-400 border-blue-500/20' },
    reference: { icon: <FileText className="h-3 w-3 text-emerald-400" />, label: 'Ref', color: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' },
    target: { icon: <MapPin className="h-3 w-3 text-indigo-400" />, label: 'Target', color: 'bg-indigo-500/10 text-indigo-400 border-indigo-500/20' },
    pnp: { icon: <MapPin className="h-3 w-3 text-orange-400" />, label: 'PnP', color: 'bg-orange-500/10 text-orange-400 border-orange-500/20' },
  };

  const resultsContainerRef = useRef<HTMLDivElement>(null);

  // Auto-scroll results with keyboard
  useEffect(() => {
    if (resultsContainerRef.current) {
      const container = resultsContainerRef.current;
      const selectedElem = container.children[0]?.children[selectedIndex] as HTMLElement;
      if (selectedElem) {
        selectedElem.scrollIntoView({
          behavior: 'smooth',
          block: 'nearest',
        });
      }
    }
  }, [selectedIndex]);

  return (
    <div className="fixed inset-0 z-[200] flex items-start justify-center pt-24 p-6 sm:p-12 no-drag">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/90 backdrop-blur-sm"
      />
      
      <motion.div
        initial={{ opacity: 0, scale: 0.95, y: -20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: -20 }}
        className="relative w-full max-w-2xl bg-[#0a0f1d] border border-white/10 rounded-[2rem] shadow-[0_30px_100px_rgba(0,0,0,0.8)] overflow-hidden flex flex-col max-h-[80vh]"
      >
        <div className="p-6 pb-2">
          <div className="relative group">
            <div className="absolute -inset-1 bg-gradient-to-r from-primary to-blue-600 rounded-2xl blur-md opacity-20 group-hover:opacity-40 transition-all duration-700" />
            <div className="relative flex items-center bg-white/5 border border-white/10 rounded-2xl p-4 shadow-inner">
              <Search className="h-6 w-6 text-white/30 mr-4" />
              <input
                ref={inputRef}
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search MPNs, Designators, or Descriptions..."
                className="w-full bg-transparent text-xl font-bold text-white placeholder-white/20 focus:outline-none"
              />
              {query && (
                <button 
                  onClick={() => setQuery('')}
                  title="Clear search"
                  className="p-1 rounded-lg hover:bg-white/10 text-white/30 hover:text-white transition-all cursor-pointer"
                >
                  <X className="h-5 w-5" />
                </button>
              )}
            </div>
          </div>
          
          <div className="mt-3 px-1 flex items-center justify-between text-[9px] font-black uppercase tracking-[0.2em] text-white/20">
            <div>{title}</div>
            {query.length >= 2 && <div>{results.length} Matches Found</div>}
          </div>
        </div>

        <div ref={resultsContainerRef} className="flex-1 overflow-y-auto p-4 space-y-3 bom-scroll-stable">
          {query.length === 0 ? (
            <div className="py-20 flex flex-col items-center justify-center space-y-4 opacity-30">
              <div className="p-6 rounded-[2rem] bg-white/5 border border-white/5">
                <Search className="h-12 w-12 text-white/40" />
              </div>
              <p className="text-sm font-bold uppercase tracking-widest text-center">Type something to search <br/> within the current mode</p>
            </div>
          ) : query.length < 2 ? (
            <div className="py-12 flex flex-col items-center justify-center space-y-3 opacity-20">
              <AlertCircle className="h-6 w-6" />
              <p className="text-[10px] font-black uppercase tracking-[0.2em]">Min 2 characters required</p>
            </div>
          ) : results.length === 0 ? (
            <div className="py-20 flex flex-col items-center justify-center space-y-4 opacity-30">
              <div className="p-6 rounded-[2rem] bg-white/5 border border-white/5">
                <X className="h-12 w-12 text-white/40" />
              </div>
              <p className="text-sm font-bold uppercase tracking-widest text-center">No results found for<br/><span className="text-white">"{query}"</span></p>
            </div>
          ) : (
            <div className="space-y-2">
              {results.map((res, i) => (
                <button
                  key={res.id}
                  onClick={() => onResultClick?.(res)}
                  className={`w-full p-4 rounded-2xl border transition-all text-left group flex items-start gap-4 cursor-pointer focus:outline-none ${
                    selectedIndex === i 
                      ? 'bg-white/15 border-primary shadow-[0_0_20px_rgba(255,107,0,0.15)] z-10' 
                      : 'bg-white/5 border-white/5 hover:border-primary/40 hover:bg-white/10'
                  }`}
                >
                  <div className={`mt-0.5 px-2 py-0.5 rounded-lg border text-[8px] font-black uppercase tracking-widest ${sourceConfig[res.source].color}`}>
                    {sourceConfig[res.source].label}
                  </div>
                  
                  <div className="flex-1 min-w-0 space-y-1">
                    <div className="flex items-center justify-between">
                      <h4 className="text-sm font-black text-white truncate leading-tight group-hover:text-primary transition-colors">
                        {res.mpn}
                      </h4>
                      <span className="text-[10px] font-mono text-white/20 whitespace-nowrap">Row #{res.rowIdx}</span>
                    </div>
                    
                    <p className="text-[10px] text-white/40 truncate font-bold uppercase tracking-wide leading-relaxed">
                      {res.description || 'No description available'}
                    </p>
                    
                    <div className="flex flex-wrap gap-1 mt-2">
                      <CornerDownRight className="h-2.5 w-2.5 text-white/20 mt-0.5" />
                      {res.designators.slice(0, 8).map((d, i) => (
                        <span key={i} className="text-[10px] font-mono text-white/60 bg-white/10 px-1 rounded transition-colors group-hover:text-white group-hover:bg-primary/20">
                          {d}{i === 7 && res.designators.length > 8 ? '...' : ''}
                        </span>
                      ))}
                      {res.designators.length > 8 && <span className="text-[9px] font-black text-white/20 uppercase tracking-tighter">+{res.designators.length - 8} more</span>}
                    </div>
                  </div>
                  
                  <div className="self-center">
                    <ChevronRight className="h-5 w-5 text-white/10 group-hover:text-primary transition-all group-hover:translate-x-1" />
                  </div>
                </button>
              ))}
            </div>
          )}
        </div>

        <div className="p-4 border-t border-white/5 flex items-center justify-between bg-black/40">
          <div className="flex items-center gap-4 text-[9px] font-black uppercase tracking-widest text-white/20">
            <div className="flex items-center gap-1.5"><kbd className="px-1.5 py-0.5 rounded bg-white/10 text-white/40 border border-white/10">ESC</kbd> to close</div>
            <div className="flex items-center gap-1.5"><kbd className="px-1.5 py-0.5 rounded bg-white/10 text-white/40 border border-white/10">↑↓</kbd> to navigate</div>
            <div className="flex items-center gap-1.5"><kbd className="px-1.5 py-0.5 rounded bg-white/10 text-white/40 border border-white/10">ENTER</kbd> to select</div>
          </div>
          <button 
            onClick={onClose}
            className="text-[10px] font-black uppercase tracking-widest text-white/40 hover:text-white transition-all cursor-pointer"
          >
            Close Search
          </button>
        </div>
      </motion.div>
    </div>
  );
}
