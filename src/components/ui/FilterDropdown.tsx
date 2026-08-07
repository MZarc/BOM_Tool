import { useState, useMemo, useRef, useEffect } from 'react';
import { Filter, Search, X, Check } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

interface FilterDropdownProps {
  column: string;
  values: string[];
  selectedValues: Set<string>;
  onFilterChange: (values: Set<string>) => void;
  isOpen: boolean;
  onOpenChange: (open: boolean) => void;
}

export function FilterDropdown({
  column,
  values,
  selectedValues,
  onFilterChange,
  isOpen,
  onOpenChange
}: FilterDropdownProps) {
  const [search, setSearch] = useState('');
  const buttonRef = useRef<HTMLButtonElement>(null);
  const [coords, setCoords] = useState({ top: 0, left: 0 });
  const [dropUp, setDropUp] = useState(false);
  const [maxListHeight, setMaxListHeight] = useState(256);

  useEffect(() => {
    if (isOpen && buttonRef.current) {
      const rect = buttonRef.current.getBoundingClientRect();
      const spaceBelow = window.innerHeight - rect.bottom;
      const spaceAbove = rect.top;
      // Dropdown total height estimate: search(~70) + selectAll(~50) + list(variable) + footer(~60) + padding(~16) ≈ 340px min
      const DROPDOWN_MIN = 340;

      const shouldFlipUp = spaceBelow < DROPDOWN_MIN && spaceAbove > spaceBelow;

      setDropUp(shouldFlipUp);

      if (shouldFlipUp) {
        // Position above: we'll use bottom offset via top calculated from window height
        const availableHeight = spaceAbove - 8;
        setMaxListHeight(Math.min(256, availableHeight - 200)); // 200 for other chrome
        setCoords({
          top: rect.top - 8, // will be used as "bottom" anchor via CSS transform
          left: Math.min(window.innerWidth - 268, rect.left),
        });
      } else {
        const availableHeight = spaceBelow - 8;
        setMaxListHeight(Math.min(256, availableHeight - 200));
        setCoords({
          top: rect.bottom + 8,
          left: Math.min(window.innerWidth - 268, rect.left),
        });
      }
    }
  }, [isOpen]);


  const uniqueValues = useMemo(() => {
    const sorted = Array.from(new Set(values)).sort((a, b) => 
      a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
    );
    return sorted;
  }, [values]);

  const filteredValues = useMemo(() => {
    return uniqueValues.filter(v => 
      v.toLowerCase().includes(search.toLowerCase())
    );
  }, [uniqueValues, search]);

  const toggleValue = (val: string) => {
    const next = new Set(selectedValues);
    if (next.has(val)) {
      next.delete(val);
    } else {
      next.add(val);
    }
    onFilterChange(next);
  };

  const toggleAll = () => {
    if (selectedValues.size === uniqueValues.length) {
      onFilterChange(new Set());
    } else {
      onFilterChange(new Set(uniqueValues));
    }
  };

  const isAllSelected = selectedValues.size === uniqueValues.length;
  const isFiltered = selectedValues.size < uniqueValues.length;

  return (
    <div className="relative inline-block text-left no-drag">
      <button
        ref={buttonRef}
        onClick={(e) => {
          e.stopPropagation();
          onOpenChange(!isOpen);
        }}
        className={`p-1.5 rounded-md transition-all duration-300 cursor-pointer ${
          isFiltered 
            ? 'text-primary bg-primary/20 shadow-[0_0_15px_rgba(59,130,246,0.3)] border border-primary/30' 
            : 'text-white/30 hover:text-white hover:bg-white/10 border border-transparent'
        }`}
        title={`Filter ${column}`}
      >
        <Filter className={`h-3 w-3 ${isFiltered ? 'fill-current' : ''}`} />
      </button>

      <AnimatePresence>
        {isOpen && (
          <>
            <div className="fixed inset-0 z-[100]" onClick={() => onOpenChange(false)} />
            <motion.div
              initial={{ opacity: 0, y: dropUp ? -10 : 10, scale: 0.95 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: dropUp ? -10 : 10, scale: 0.95 }}
              className="fixed z-[110] w-[260px] bg-[#0c1424] border border-white/10 rounded-2xl shadow-[0_20px_50px_rgba(0,0,0,0.8)] overflow-hidden flex flex-col"
              style={{
                top: coords.top,
                left: coords.left,
                transform: dropUp ? 'translateY(-100%)' : 'none',
              }}
            >
              <div className="p-3 border-b border-white/5 bg-white/5 mx-2 mt-2 rounded-2xl">
                <div className="relative group">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-white/20 group-focus-within:text-primary transition-colors" />
                  <input
                    autoFocus
                    type="text"
                    placeholder={`Search ${column}...`}
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl py-2 pl-9 pr-3 text-xs text-white/90 placeholder-white/20 focus:outline-none focus:border-primary/50 transition-all font-medium"
                  />
                  {search && (
                    <button 
                      onClick={() => setSearch('')}
                      title="Clear search"
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-white/30 hover:text-white"
                    >
                      <X className="h-3 w-3" />
                    </button>
                  )}
                </div>
              </div>

              <div className="px-5 py-3.5 flex items-center justify-between">
                <button 
                  onClick={toggleAll}
                  className="flex items-center gap-2.5 text-[11px] font-semibold text-white/60 hover:text-white transition-all group"
                >
                  <div className={`h-4 w-4 rounded-md border flex items-center justify-center transition-all ${
                    isAllSelected ? 'bg-primary border-primary' : 'bg-white/5 border-white/10 group-hover:border-white/30'
                  }`}>
                    {isAllSelected && <Check className="h-2.5 w-2.5 text-white stroke-[3]" />}
                  </div>
                  Select All
                </button>
                <div className="text-[10px] font-mono text-white/20 font-medium">
                  {selectedValues.size} / {uniqueValues.length}
                </div>
              </div>

              <div className="overflow-y-auto px-2 pb-2 bom-scroll-stable custom-scrollbar" style={{ maxHeight: `${maxListHeight}px` }}>
                {filteredValues.length === 0 ? (
                  <div className="py-12 text-center text-white/10 text-[10px] font-semibold uppercase tracking-[0.2em]">
                    No Results
                  </div>
                ) : (
                  <div className="space-y-0.5">
                    {filteredValues.map((val) => {
                      const isSelected = selectedValues.has(val);
                      const displayVal = val || '(Blank)';
                      return (
                        <button
                          key={val}
                          onClick={() => toggleValue(val)}
                          className={`w-full flex items-center gap-3 px-3 py-2 rounded-xl transition-all text-left group ${
                            isSelected ? 'bg-primary/10' : 'hover:bg-white/5'
                          }`}
                        >
                          <div className={`h-4 w-4 rounded-md border flex items-center justify-center transition-all flex-shrink-0 ${
                            isSelected ? 'bg-primary border-primary shadow-[0_0_10px_rgba(59,130,246,0.3)]' : 'bg-white/5 border-white/10 group-hover:border-white/30'
                          }`}>
                            {isSelected && <Check className="h-2.5 w-2.5 text-white stroke-[3]" />}
                          </div>
                          <span className={`text-[11px] font-medium truncate leading-tight ${
                            isSelected ? 'text-blue-400' : 'text-white/40 group-hover:text-white/70'
                          }`}>
                            {displayVal}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="p-3 border-t border-white/5 flex gap-2 bg-[#050810]/60">
                <button 
                  onClick={() => onOpenChange(false)}
                  className="flex-[1.5] py-2.5 rounded-xl bg-blue-600 text-white text-[10px] font-bold uppercase tracking-widest hover:bg-blue-500 transition-all cursor-pointer shadow-[0_4px_12px_rgba(37,99,235,0.3)] active:scale-95 flex items-center justify-center gap-2"
                >
                  Apply Filters
                </button>
                <button 
                  onClick={() => {
                    onFilterChange(new Set(uniqueValues));
                    onOpenChange(false);
                  }}
                  className="flex-1 py-2.5 rounded-xl bg-white/5 text-white/40 text-[10px] font-bold uppercase tracking-widest hover:bg-white/10 hover:text-white transition-all cursor-pointer border border-white/10 active:scale-95"
                >
                  Reset
                </button>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>
    </div>
  );
}
