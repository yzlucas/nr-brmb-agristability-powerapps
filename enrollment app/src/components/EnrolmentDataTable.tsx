import { useRef } from 'react';
import type { DragEvent } from 'react';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import type { FilterOperator, SortDir, SortKey } from '../types/enrollment';
import { ALL_COLUMNS } from '../constants/columns';
import { ColumnHeaderMenu, type ColumnHeaderFilterProps } from './ColumnHeaderMenu';
import { renderCell } from './renderCell';
import { formatEnrolmentStatusDisplay } from '../utils/helpers';

type Props = {
  allRowsCount: number;
  pagedRows: Vsi_participantprogramyears[];
  visibleColumnKeys: SortKey[];
  allPageSelected: boolean;
  somePageSelected: boolean;
  onToggleSelectAll: () => void;
  selectedIds: Set<string>;
  onToggleSelect: (id: string) => void;
  onRangeSelect: (ids: string[], checked: boolean) => void;
  colDragIdx: number | null;
  onColDragStart: (index: number) => void;
  onColDragOver: (event: DragEvent, index: number) => void;
  onColDragEnd: () => void;
  taskStatusOptions: string[];
  taskStatusFilter: Set<string>;
  taskFilterOp: FilterOperator;
  onTaskStatusFilterChange: (value: Set<string>) => void;
  onTaskFilterOperatorChange: (value: FilterOperator) => void;
  enrolStatusOptions: string[];
  enrolStatusFilter: Set<string>;
  enrolFilterOp: FilterOperator;
  onEnrolStatusFilterChange: (value: Set<string>) => void;
  onEnrolFilterOperatorChange: (value: FilterOperator) => void;
  yearOptions: string[];
  yearFilter: Set<string>;
  onYearFilterChange: (value: Set<string>) => void;
  ownerOptions: string[];
  ownerFilter: Set<string>;
  onOwnerFilterChange: (value: Set<string>) => void;
  ownerFilterShortcuts?: Array<{ label: string; values: Set<string> }>;
  sortKey: SortKey | null;
  sortDir: SortDir;
  onSort: (key: SortKey, dir: SortDir) => void;
  columnWidths: Partial<Record<SortKey, number>>;
  onColumnWidthChange: (key: SortKey) => (width: number | undefined) => void;
  avatarUrls: Record<string, string>;
  coreAppId: string | null;
  coreBaseUrl: string | null;
};

export function EnrolmentDataTable({
  allRowsCount,
  pagedRows,
  visibleColumnKeys,
  allPageSelected,
  somePageSelected,
  onToggleSelectAll,
  selectedIds,
  onToggleSelect,
  onRangeSelect,
  colDragIdx,
  onColDragStart,
  onColDragOver,
  onColDragEnd,
  taskStatusOptions,
  taskStatusFilter,
  taskFilterOp,
  onTaskStatusFilterChange,
  onTaskFilterOperatorChange,
  enrolStatusOptions,
  enrolStatusFilter,
  enrolFilterOp,
  onEnrolStatusFilterChange,
  onEnrolFilterOperatorChange,
  yearOptions,
  yearFilter,
  onYearFilterChange,
  ownerOptions,
  ownerFilter,
  onOwnerFilterChange,
  ownerFilterShortcuts,
  sortKey,
  sortDir,
  onSort,
  columnWidths,
  onColumnWidthChange,
  avatarUrls,
  coreAppId,
  coreBaseUrl,
}: Props) {
  const lastClickedIdxRef = useRef<number>(-1);
  const isEmptyState = allRowsCount === 0 || pagedRows.length === 0;

  return (
    <div className={`enrolment-table-container${isEmptyState ? ' is-empty' : ''}`}>
      <table className="enrolment-table">
        <thead>
          <tr>
            <th style={{ width: '2rem' }}>
              <input
                type="checkbox"
                checked={allPageSelected}
                ref={el => { if (el) el.indeterminate = somePageSelected && !allPageSelected; }}
                onChange={onToggleSelectAll}
              />
            </th>
            {visibleColumnKeys.map((k, colIdx) => {
              const def = ALL_COLUMNS.find(c => c.key === k)!;
              const extra: ColumnHeaderFilterProps = {};

              if (k === 'taskStatus') {
                extra.filterOptions = taskStatusOptions;
                extra.selectedFilters = taskStatusFilter;
                extra.filterOperator = taskFilterOp;
                extra.onFilterChange = onTaskStatusFilterChange;
                extra.onFilterOperatorChange = onTaskFilterOperatorChange;
              }

              if (k === 'enrolStatus') {
                extra.filterOptions = enrolStatusOptions;
                extra.filterOptionLabels = Object.fromEntries(enrolStatusOptions.map(o => [o, formatEnrolmentStatusDisplay(o)]));
                extra.selectedFilters = enrolStatusFilter;
                extra.filterOperator = enrolFilterOp;
                extra.onFilterChange = onEnrolStatusFilterChange;
                extra.onFilterOperatorChange = onEnrolFilterOperatorChange;
              }

              if (k === 'year') {
                extra.filterOptions = yearOptions;
                extra.selectedFilters = yearFilter;
                extra.onFilterChange = onYearFilterChange;
              }

              if (k === 'owner') {
                extra.filterOptions = ownerOptions;
                extra.selectedFilters = ownerFilter;
                extra.onFilterChange = onOwnerFilterChange;
                extra.filterShortcuts = ownerFilterShortcuts;
              }

              const dragProps = {
                draggable: true,
                onDragStart: () => onColDragStart(colIdx),
                onDragOver: (event: DragEvent) => onColDragOver(event, colIdx),
                onDragEnd: onColDragEnd,
                className: colDragIdx === colIdx ? 'col-dragging' : undefined,
              };

              if (k === 'sharepoint') {
                return <th key={k} {...dragProps} style={{ cursor: 'grab' }}>{def.label}</th>;
              }

              if (k === 'flagged') {
                return <th key={k} {...dragProps} style={{ cursor: 'grab', width: 28 }}></th>;
              }

              return (
                <ColumnHeaderMenu
                  key={k}
                  label={def?.label || 'N/A'}
                  sortKey={k}
                  currentSortKey={sortKey}
                  currentSortDir={sortDir}
                  onSort={onSort}
                  columnWidth={columnWidths[k]}
                  onColumnWidthChange={onColumnWidthChange(k)}
                  dragProps={dragProps}
                  {...extra}
                />
              );
            })}
          </tr>
        </thead>
        <tbody>
          {allRowsCount === 0 ? (
            <tr><td colSpan={visibleColumnKeys.length + 1} className="enrolment-empty">No records found</td></tr>
          ) : pagedRows.length === 0 ? (
            <tr><td colSpan={visibleColumnKeys.length + 1} className="enrolment-empty">No rows returned</td></tr>
          ) : (
            pagedRows.map((row, index) => {
              const raw = row as unknown as Record<string, unknown>;
              return (
                <tr key={row.vsi_participantprogramyearid ?? index}>
                  <td>
                    <input
                      type="checkbox"
                      checked={selectedIds.has(row.vsi_participantprogramyearid)}
                      onChange={() => {}}
                      onClick={(e: React.MouseEvent<HTMLInputElement>) => {
                        if (e.shiftKey && lastClickedIdxRef.current >= 0) {
                          const start = Math.min(lastClickedIdxRef.current, index);
                          const end = Math.max(lastClickedIdxRef.current, index);
                          const rangeIds = pagedRows
                            .slice(start, end + 1)
                            .map(r => r.vsi_participantprogramyearid);
                          const willBeChecked = !selectedIds.has(row.vsi_participantprogramyearid);
                          onRangeSelect(rangeIds, willBeChecked);
                        } else {
                          onToggleSelect(row.vsi_participantprogramyearid);
                        }
                        lastClickedIdxRef.current = index;
                      }}
                    />
                  </td>
                  {visibleColumnKeys.map(key => {
                    // Always pass _source: 'dashboard' for navigation context
                    const rowWithSource = { ...row, _source: 'dashboard' };
                    return renderCell(key, rowWithSource, raw, avatarUrls, coreAppId, coreBaseUrl);
                  })}
                </tr>
              );
            })
          )}
        </tbody>
      </table>
    </div>
  );
}
