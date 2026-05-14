import { useRef, useState, useLayoutEffect } from 'react';
import { createPortal } from 'react-dom';
import { Filter, FilterX } from 'lucide-react';
import type { SortDir, FilterOperator } from '../types/enrollment';

type MenuView = 'main' | 'filter' | 'width';

export type ColumnHeaderDragProps = {
  draggable: boolean;
  onDragStart: () => void;
  onDragOver: (e: React.DragEvent) => void;
  onDragEnd: () => void;
  className?: string;
};

export type ColumnHeaderFilterProps = {
  filterOptions?: string[];
  filterOptionLabels?: Record<string, string>;
  selectedFilters?: Set<string>;
  filterOperator?: FilterOperator;
  onFilterChange?: (next: Set<string>) => void;
  onFilterOperatorChange?: (op: FilterOperator) => void;
  filterShortcuts?: Array<{ label: string; values: Set<string> }>;
};

export type ColumnHeaderMenuProps<K extends string = string> = {
  label: string;
  sortKey: K;
  currentSortKey: K | null;
  currentSortDir: SortDir;
  onSort: (key: K, dir: SortDir) => void;
  columnWidth: number | undefined;
  onColumnWidthChange: (w: number | undefined) => void;
  dragProps?: ColumnHeaderDragProps;
} & ColumnHeaderFilterProps;

export function ColumnHeaderMenu<K extends string = string>({
  label,
  sortKey,
  currentSortKey,
  currentSortDir,
  onSort,
  filterOptions,
  filterOptionLabels,
  selectedFilters,
  filterOperator,
  onFilterChange,
  onFilterOperatorChange,
  filterShortcuts,
  columnWidth,
  onColumnWidthChange,
  dragProps,
}: ColumnHeaderMenuProps<K>) {

  const [open, setOpen] = useState(false);
  const [view, setView] = useState<MenuView>('main');
  const [operatorOpen, setOperatorOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const anchorRef = useRef<HTMLSpanElement>(null);
  const [menuStyle, setMenuStyle] = useState<React.CSSProperties>({});

  // Position the menu using the anchor's bounding rect
  useLayoutEffect(() => {
    if (open && anchorRef.current && menuRef.current) {
      const anchorRect = anchorRef.current.getBoundingClientRect();
      const menu = menuRef.current;
      const menuHeight = menu.offsetHeight;
      const menuWidth = menu.offsetWidth;
      const viewportHeight = window.innerHeight;
      const viewportWidth = window.innerWidth;
      let top = anchorRect.bottom + 2;
      let left = anchorRect.left;
      // If not enough space below, open upwards
      if (top + menuHeight > viewportHeight && anchorRect.top - menuHeight > 0) {
        top = anchorRect.top - menuHeight - 2;
      }
      // Clamp left to viewport
      if (left + menuWidth > viewportWidth) {
        left = viewportWidth - menuWidth - 8;
      }
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setMenuStyle({ position: 'fixed', top, left, zIndex: 1000, minWidth: 200 });
    }
  }, [open, view]);

  const close = () => { setOpen(false); setView('main'); setOperatorOpen(false); };

  const isSorted = currentSortKey === sortKey;
  const hasFilter = selectedFilters && selectedFilters.size > 0;

  const toggle = (val: string) => {
    if (!selectedFilters || !onFilterChange) return;
    const next = new Set(selectedFilters);
    if (next.has(val)) next.delete(val); else next.add(val);
    onFilterChange(next);
  };

  return (
    <th
      className={`col-hdr-menu-th${dragProps?.className ? ' ' + dragProps.className : ''}`}
      style={{ position: 'relative', cursor: 'grab', minWidth: columnWidth ? `${columnWidth}px` : undefined, width: columnWidth ? `${columnWidth}px` : undefined }}
      draggable={dragProps?.draggable}
      onDragStart={dragProps?.onDragStart}
      onDragOver={dragProps?.onDragOver}
      onDragEnd={dragProps?.onDragEnd}
    >
      <span
        className="col-hdr-label"
        ref={anchorRef}
        onClick={() => { setOpen(o => !o); setView('main'); }}
        style={{ userSelect: 'none' }}
      >
        {label}
        {isSorted && <span className="col-hdr-sort-indicator">{currentSortDir === 'asc' ? ' ↑' : ' ↓'}</span>}
        {hasFilter && <span className="col-hdr-filter-indicator" title="Filtered">&#x25BC;</span>}
        <span className="col-hdr-chevron">&#x25BE;</span>
      </span>

      {open && createPortal(
        <>
          <div className="chm-backdrop" onClick={close} />
          <div className="chm-panel" ref={menuRef} style={menuStyle} onClick={e => e.stopPropagation()}>
            {view === 'main' && (
              <>
                <button className="chm-item" onClick={() => { onSort(sortKey, 'asc'); close(); }}>
                  <span className="chm-icon">↑</span> A to Z
                </button>
                <button className="chm-item" onClick={() => { onSort(sortKey, 'desc'); close(); }}>
                  <span className="chm-icon">↓</span> Z to A
                </button>
                {filterOptions && (
                  <>
                    <div className="chm-divider" />
                    <button className="chm-item" onClick={() => setView('filter')}>
                      <span className="chm-icon"><Filter size={14} /></span> Filter by
                    </button>
                    {filterShortcuts && filterShortcuts.map(sc => (
                      <button key={sc.label} className="chm-item chm-item-shortcut" onClick={() => { onFilterChange!(sc.values); close(); }}>
                        <span className="chm-icon">&#x2713;</span> {sc.label}
                      </button>
                    ))}
                    {hasFilter && (
                      <button className="chm-item chm-item-clear" onClick={() => { onFilterChange!(new Set()); close(); }}>
                        <span className="chm-icon"><FilterX size={14} /></span> Clear filter
                      </button>
                    )}
                  </>
                )}
                <div className="chm-divider" />
                <button className="chm-item" onClick={() => setView('width')}>
                  <span className="chm-icon">↔</span> Column width
                </button>
              </>
            )}

            {view === 'filter' && filterOptions && selectedFilters && onFilterChange && (
              <div className="chm-filter-view">
                <div className="chm-filter-header">
                  <h4>Filter by</h4>
                  <button className="chm-close" onClick={close}>✕</button>
                </div>
                {onFilterOperatorChange && (
                  <div className="chm-operator-wrapper">
                    <button className="chm-operator-btn" onClick={() => setOperatorOpen(o => !o)}>
                      {filterOperator === 'notEquals' ? 'Does not equal' : 'Equals'}
                      <span className="chm-operator-chevron">&#x25BE;</span>
                    </button>
                    {operatorOpen && (
                      <div className="chm-operator-dropdown">
                        <button className={`chm-operator-opt${filterOperator === 'equals' ? ' active' : ''}`} onClick={() => { onFilterOperatorChange('equals'); setOperatorOpen(false); }}>Equals</button>
                        <button className={`chm-operator-opt${filterOperator === 'notEquals' ? ' active' : ''}`} onClick={() => { onFilterOperatorChange('notEquals'); setOperatorOpen(false); }}>Does not equal</button>
                      </div>
                    )}
                  </div>
                )}
                <div className="chm-values">
                  {filterOptions.map(opt => (
                    <label key={opt} className="chm-value-item">
                      <input type="checkbox" checked={selectedFilters.has(opt)} onChange={() => toggle(opt)} />
                      <span>{filterOptionLabels?.[opt] ?? opt}</span>
                    </label>
                  ))}
                </div>
                <div className="chm-filter-actions">
                  <button className="chm-apply" onClick={close}>Apply</button>
                  <button className="chm-clear" onClick={() => { onFilterChange(new Set()); close(); }}>Clear filter</button>
                </div>
              </div>
            )}

            {view === 'width' && (
              <div className="chm-width-view">
                <div className="chm-filter-header">
                  <h4>Column width</h4>
                  <button className="chm-close" onClick={close}>✕</button>
                </div>
                <label className="chm-width-label">Preferred width</label>
                <input
                  className="chm-width-input"
                  type="number"
                  min={40}
                  max={600}
                  value={columnWidth ?? ''}
                  placeholder="Auto"
                  onChange={e => {
                    const v = e.target.value ? Number(e.target.value) : undefined;
                    onColumnWidthChange(v);
                  }}
                />
                <div className="chm-filter-actions">
                  <button className="chm-apply" onClick={close}>Apply</button>
                  <button className="chm-clear" onClick={() => { onColumnWidthChange(undefined); close(); }}>Reset</button>
                </div>
              </div>
            )}
          </div>
        </>,
        document.body
      )}
    </th>
  );
}
