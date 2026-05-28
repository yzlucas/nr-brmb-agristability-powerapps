import { useCallback, useEffect, useMemo, useState, type DragEvent } from 'react';
import { Link } from 'react-router-dom';
import { Columns2, Filter, FilterX, Info, RefreshCw } from 'lucide-react';

import type { SortKey, SortDir, FilterOperator, AdvFilterNode, LogicOp, QuickFilterState } from '../types/enrollment';
import { DEFAULT_VISIBLE_KEYS } from '../constants/columns';
import { countActiveNodes } from '../utils/filterTree';
import { useEnrolmentData, useSortedAndFilteredRows } from '../hooks/useEnrolmentData';
import { useRole } from '../context/RoleContext';
import { resolveCurrentSystemUser } from '../utils/currentUser';
import { clearSaCache } from './SupervisorApprovalPage';
import { useViews } from '../hooks/useViews';

import { ViewsMenu } from '../components/ViewsMenu';
import { EditColumnsPanel } from '../components/EditColumnsPanel';
import { EditFiltersPanel } from '../components/EditFiltersPanel';
import { BulkNoticesModal } from '../components/BulkNoticesModal';
import { AssignOwnerModal } from '../components/AssignOwnerModal';
import { ReferToSupervisorModal } from '../components/ReferToSupervisorModal';
import { ApproveCalculatedFeesModal } from '../components/ApproveCalculatedFeesModal';
import { Toast, nextToastId } from '../components/Toast';
import type { ToastMessage } from '../components/Toast';
import { EnrollmentSearchBar } from '../components/EnrollmentSearchBar';
import { EnrolmentQuickFilters } from '../components/EnrolmentQuickFilters';
import { EnrolmentDataTable } from '../components/EnrolmentDataTable';

import { EnrolmentActionsBar } from '../components/EnrolmentActionsBar';

const PAGE_SIZE = 300;

// Module-level cache — persists filter/sort/pagination state across SPA navigations.
type DashboardFilterCache = {
  visibleColumnKeys: SortKey[];
  columnWidths: Partial<Record<SortKey, number>>;
  sortKey: SortKey | null;
  sortDir: SortDir;
  filters: QuickFilterState;
  searchQuery: string;
  taskStatusFilter: Set<string>;
  enrolStatusFilter: Set<string>;
  yearFilter: Set<string>;
  ownerFilter: Set<string>;
  taskFilterOp: FilterOperator;
  enrolFilterOp: FilterOperator;
  advFilterNodes: AdvFilterNode[];
  advLogicOp: LogicOp;
  currentPage: number;
};
let dashboardFilterCache: DashboardFilterCache | null = null;

export function DashboardHomePage() {
  const { activeRole } = useRole();
  const { rows, setRows, loading, error, avatarUrls, fetchEnrolments, coreAppId, coreBaseUrl, fetchCoreAppId } = useEnrolmentData();

  // Refresh handler is defined after useViews so reloadViews is available

  // Column & sort state
  const [visibleColumnKeys, setVisibleColumnKeys] = useState<SortKey[]>(() => dashboardFilterCache?.visibleColumnKeys ?? [...DEFAULT_VISIBLE_KEYS]);
  const [columnWidths, setColumnWidths] = useState<Partial<Record<SortKey, number>>>(() => dashboardFilterCache?.columnWidths ?? {});
  const [sortKey, setSortKey] = useState<SortKey | null>(() => dashboardFilterCache?.sortKey ?? 'modifiedOn');
  const [sortDir, setSortDir] = useState<SortDir>(() => dashboardFilterCache?.sortDir ?? 'desc');

  // Filter state
  const [filters, setFilters] = useState<QuickFilterState>(() => dashboardFilterCache?.filters ?? {
    verifiedCalc: false,
    unverifiedCalc: false,
    flagged: false,
    partnerships: false,
    fortyFiveDayLetter: false,
    varianceAlert: false,
  });
  const [searchQuery, setSearchQuery] = useState(() => dashboardFilterCache?.searchQuery ?? '');
  const [taskStatusFilter, setTaskStatusFilter] = useState<Set<string>>(() => dashboardFilterCache?.taskStatusFilter ?? new Set());
  const [enrolStatusFilter, setEnrolStatusFilter] = useState<Set<string>>(() => dashboardFilterCache?.enrolStatusFilter ?? new Set());
  const [yearFilter, setYearFilter] = useState<Set<string>>(() => dashboardFilterCache?.yearFilter ?? new Set());
  const [ownerFilter, setOwnerFilter] = useState<Set<string>>(() => dashboardFilterCache?.ownerFilter ?? new Set());
  const [currentUserDisplayName, setCurrentUserDisplayName] = useState<string | null>(null);

  useEffect(() => {
    resolveCurrentSystemUser().then(u => setCurrentUserDisplayName(u.displayName)).catch(() => {});
  }, []);
  const [taskFilterOp, setTaskFilterOp] = useState<FilterOperator>(() => dashboardFilterCache?.taskFilterOp ?? 'equals');
  const [enrolFilterOp, setEnrolFilterOp] = useState<FilterOperator>(() => dashboardFilterCache?.enrolFilterOp ?? 'equals');
  const [advFilterNodes, setAdvFilterNodes] = useState<AdvFilterNode[]>(() => dashboardFilterCache?.advFilterNodes ?? []);
  const [advLogicOp, setAdvLogicOp] = useState<LogicOp>(() => dashboardFilterCache?.advLogicOp ?? 'AND');

  // Pagination & selection
  const [currentPage, setCurrentPage] = useState(() => dashboardFilterCache?.currentPage ?? 1);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());

  // Panel visibility
  const [showEditColumns, setShowEditColumns] = useState(false);
  const [showEditFilters, setShowEditFilters] = useState(false);
  const [showBulkModal, setShowBulkModal] = useState(false);
  const [showAssignModal, setShowAssignModal] = useState(false);
  const [showSupervisorModal, setShowSupervisorModal] = useState(false);
  const [showApproveFeesModal, setShowApproveFeesModal] = useState(false);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);

  const addToast = useCallback((message: string, type: ToastMessage['type'] = 'success') => {
    setToasts(prev => [...prev, { id: nextToastId(), message, type }]);
  }, []);
  const dismissToast = useCallback((id: number) => {
    setToasts(prev => prev.filter(t => t.id !== id));
  }, []);

  const setFiltersAndReset = useCallback((next: QuickFilterState) => {
    setFilters(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setTaskStatusFilterAndReset = useCallback((next: Set<string>) => {
    setTaskStatusFilter(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setEnrolStatusFilterAndReset = useCallback((next: Set<string>) => {
    setEnrolStatusFilter(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setYearFilterAndReset = useCallback((next: Set<string>) => {
    setYearFilter(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setOwnerFilterAndReset = useCallback((next: Set<string>) => {
    setOwnerFilter(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setTaskFilterOpAndReset = useCallback((next: FilterOperator) => {
    setTaskFilterOp(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setEnrolFilterOpAndReset = useCallback((next: FilterOperator) => {
    setEnrolFilterOp(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setAdvFilterNodesAndReset = useCallback((next: AdvFilterNode[]) => {
    setAdvFilterNodes(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);
  const setAdvLogicOpAndReset = useCallback((next: LogicOp) => {
    setAdvLogicOp(next);
    setCurrentPage(1);
    setSelectedIds(new Set());
  }, []);

  // Build the combined node list shown in Edit Filters:
  // quick filters (except flagged) + column header filters + existing adv nodes.
  // Synthetic nodes use negative IDs to avoid collisions with nextFilterId().
  const effectiveFilterNodes = useMemo((): AdvFilterNode[] => {
    let sid = -1;
    const extra: AdvFilterNode[] = [];
    if (filters.verifiedCalc)
      extra.push({ kind: 'row', id: sid--, field: 'enrolStatus', operator: 'equals', values: new Set(['VerifiedENCalculalted']), textValue: '' });
    if (filters.unverifiedCalc)
      extra.push({ kind: 'row', id: sid--, field: 'enrolStatus', operator: 'equals', values: new Set(['UnverifiedENCalculated']), textValue: '' });
    if (filters.fortyFiveDayLetter)
      extra.push({ kind: 'row', id: sid--, field: 'enrolStatus', operator: 'equals', values: new Set(['_45DayLetter']), textValue: '' });
    if (filters.partnerships)
      extra.push({ kind: 'group', id: sid--, logic: 'OR', children: [
        { kind: 'row', id: sid--, field: 'hasPartners', operator: 'equals', values: new Set(['Yes']), textValue: '' },
        { kind: 'row', id: sid--, field: 'inCombinedFarm', operator: 'equals', values: new Set(['Yes']), textValue: '' },
      ]});
    if (taskStatusFilter.size > 0)
      extra.push({ kind: 'row', id: sid--, field: 'taskStatus', operator: taskFilterOp, values: new Set(taskStatusFilter), textValue: '' });
    if (enrolStatusFilter.size > 0)
      extra.push({ kind: 'row', id: sid--, field: 'enrolStatus', operator: enrolFilterOp, values: new Set(enrolStatusFilter), textValue: '' });
    return [...extra, ...advFilterNodes];
  }, [filters, taskStatusFilter, enrolStatusFilter, taskFilterOp, enrolFilterOp, advFilterNodes]);

  // Column drag-and-drop
  const [colDragIdx, setColDragIdx] = useState<number | null>(null);
  const handleColDragStart = (i: number) => setColDragIdx(i);
  const handleColDragOver = (e: DragEvent, i: number) => {
    e.preventDefault();
    if (colDragIdx === null || colDragIdx === i) return;
    setVisibleColumnKeys(prev => {
      const next = [...prev];
      const [moved] = next.splice(colDragIdx, 1);
      next.splice(i, 0, moved);
      return next;
    });
    setColDragIdx(i);
  };
  const handleColDragEnd = () => setColDragIdx(null);

  // Views hook
  const viewSetters = useMemo(() => ({
    setVisibleColumnKeys, setColumnWidths, setSortKey, setSortDir,
    setFilters: setFiltersAndReset,
    setTaskStatusFilter: setTaskStatusFilterAndReset,
    setEnrolStatusFilter: setEnrolStatusFilterAndReset,
    setTaskFilterOp: setTaskFilterOpAndReset,
    setEnrolFilterOp: setEnrolFilterOpAndReset,
    setAdvFilterNodes: setAdvFilterNodesAndReset,
    setAdvLogicOp: setAdvLogicOpAndReset,
    setYearFilter: setYearFilterAndReset,
    setOwnerFilter: setOwnerFilterAndReset,
  }), [
    setFiltersAndReset,
    setTaskStatusFilterAndReset,
    setEnrolStatusFilterAndReset,
    setTaskFilterOpAndReset,
    setEnrolFilterOpAndReset,
    setAdvFilterNodesAndReset,
    setAdvLogicOpAndReset,
    setYearFilterAndReset,
    setOwnerFilterAndReset,
  ]);

  const viewState = useMemo(() => ({
    visibleColumnKeys, columnWidths, sortKey, sortDir, filters,
    taskStatusFilter, enrolStatusFilter, taskFilterOp, enrolFilterOp,
    advFilterNodes, advLogicOp,
  }), [visibleColumnKeys, columnWidths, sortKey, sortDir, filters,
    taskStatusFilter, enrolStatusFilter, taskFilterOp, enrolFilterOp,
    advFilterNodes, advLogicOp]);

  const {
    savedViews, viewsLoading, activeViewId, hasUnsavedChanges, saveError,
    handleSelectView, handleSaveAsNew, handleSaveCurrentView,
    handleDeleteView, handleRenameView, handleResetDefault, reloadViews,
  } = useViews(viewState, viewSetters);

  // Close edit panels whenever a view is applied so they remount with fresh state
  const closePanels = useCallback(() => {
    setShowEditColumns(false);
    setShowEditFilters(false);
  }, []);

  // Persist filter/sort/pagination state so it survives navigating to details/calculation and back.
  useEffect(() => {
    dashboardFilterCache = {
      visibleColumnKeys, columnWidths, sortKey, sortDir,
      filters, searchQuery,
      taskStatusFilter, enrolStatusFilter, yearFilter, ownerFilter,
      taskFilterOp, enrolFilterOp, advFilterNodes, advLogicOp,
      currentPage,
    };
  }, [visibleColumnKeys, columnWidths, sortKey, sortDir, filters, searchQuery,
    taskStatusFilter, enrolStatusFilter, yearFilter, ownerFilter,
    taskFilterOp, enrolFilterOp, advFilterNodes, advLogicOp, currentPage]);

  // Refresh handler for manual reload
  const handleRefresh = useCallback(() => {
    if (typeof fetchEnrolments === 'function') fetchEnrolments();
    if (typeof fetchCoreAppId === 'function') fetchCoreAppId();
    reloadViews(false);
    setSortKey('modifiedOn');
    setSortDir('desc');
  }, [fetchEnrolments, fetchCoreAppId, reloadViews]);

  const handleSelectViewAndClose = useCallback((id: string | null) => {
    closePanels();
    handleSelectView(id);
  }, [closePanels, handleSelectView]);

  const handleResetDefaultAndClose = useCallback(() => {
    closePanels();
    handleResetDefault();
  }, [closePanels, handleResetDefault]);

  const handleDeleteViewAndClose = useCallback((id: string) => {
    closePanels();
    handleDeleteView(id);
  }, [closePanels, handleDeleteView]);

  const handleClearAllFilters = useCallback(() => {
    setSearchQuery('');
    setCurrentPage(1);
    setSelectedIds(new Set());
    handleResetDefaultAndClose();
  }, [handleResetDefaultAndClose]);

  // Stable id of the NPP system view — used to detect when NPP mode is active.
  const nppViewId = useMemo(
    () => savedViews.find(v => v.source === 'system' && /npp/i.test(v.name))?.id ?? null,
    [savedViews]
  );

  /** Clear every filter, then optionally apply a single task-status or enrol-status filter.
   * If the NPP view is currently active, reset to the default view first. */
  const applyWorklistFilter = useCallback((
    type: 'taskStatus' | 'enrolStatus',
    label: string,
  ) => {
    if (nppViewId && activeViewId === nppViewId) handleResetDefault();
    const is45Day = label === '_45DayLetter';
    setFilters({ verifiedCalc: false, unverifiedCalc: false, flagged: false, partnerships: false, fortyFiveDayLetter: is45Day, varianceAlert: false });
    setSearchQuery('');
    setTaskStatusFilter(type === 'taskStatus' ? new Set([label]) : new Set());
    setEnrolStatusFilter(type === 'enrolStatus' ? new Set([label]) : new Set());
    setYearFilter(new Set());
    setOwnerFilter(new Set());
    setAdvFilterNodes([]);
    setCurrentPage(1);
  }, [nppViewId, activeViewId, handleResetDefault]);

  // Sorting & filtering
  const { filteredRows, taskStatusOptions, enrolStatusOptions, yearOptions, ownerOptions } = useSortedAndFilteredRows(
    rows, sortKey, sortDir, filters,
    taskStatusFilter, enrolStatusFilter, yearFilter, ownerFilter, taskFilterOp, enrolFilterOp,
    advFilterNodes, advLogicOp, undefined,
  );

  const searchedRows = useMemo(() => {
    const term = searchQuery.trim().toLowerCase();
    if (!term) return filteredRows;

    return filteredRows.filter((row) => {
      const raw = row as unknown as Record<string, unknown>;
      const pin = row.vsi_name ?? '';
      const participant = (row.vsi_participantidname ?? raw['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue'] ?? '') as string;
      const farmCorp = (row.new_combinedfarmname ?? row.vsi_partnershipnames ?? '') as string;

      return [pin, participant, farmCorp].some((value) => String(value).toLowerCase().includes(term));
    });
  }, [filteredRows, searchQuery]);

  // Pagination
  const totalPages = Math.max(1, Math.ceil(searchedRows.length / PAGE_SIZE));
  const pagedRows = searchedRows.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);

  const allPageSelected = pagedRows.length > 0 && pagedRows.every(r => selectedIds.has(r.vsi_participantprogramyearid));
  const somePageSelected = pagedRows.some(r => selectedIds.has(r.vsi_participantprogramyearid));

  const toggleSelectAll = () => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (allPageSelected) {
        pagedRows.forEach(r => next.delete(r.vsi_participantprogramyearid));
      } else {
        pagedRows.forEach(r => next.add(r.vsi_participantprogramyearid));
      }
      return next;
    });
  };

  const toggleSelect = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const rangeSelect = (ids: string[], checked: boolean) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      ids.forEach(id => { if (checked) next.add(id); else next.delete(id); });
      return next;
    });
  };

  const toggleFilter = (key: keyof QuickFilterState) => {
    // When toggling the partnerships filter ON, switch to the matching system view.
    // When toggling it OFF, reset to the default view.
    if (key === 'partnerships') {
      const isCurrentlyOn = filters.partnerships;
      if (!isCurrentlyOn) {
        const partnerView = savedViews.find(
          v => v.source === 'system' && /partnership|combined/i.test(v.name)
        );
        if (partnerView) {
          // Capture current filter state before the view switch replaces it
          const prevFilters = filters;
          const prevTaskStatusFilter = taskStatusFilter;
          const prevEnrolStatusFilter = enrolStatusFilter;
          const prevTaskFilterOp = taskFilterOp;
          const prevEnrolFilterOp = enrolFilterOp;
          const prevAdvFilterNodes = advFilterNodes;
          const prevAdvLogicOp = advLogicOp;
          // Apply the view (columns, sort, widths), then restore existing filters + add partnerships
          handleSelectView(partnerView.id);
          setFilters({ ...prevFilters, partnerships: true });
          setTaskStatusFilter(prevTaskStatusFilter);
          setEnrolStatusFilter(prevEnrolStatusFilter);
          setTaskFilterOp(prevTaskFilterOp);
          setEnrolFilterOp(prevEnrolFilterOp);
          setAdvFilterNodes(prevAdvFilterNodes);
          setAdvLogicOp(prevAdvLogicOp);
          setCurrentPage(1);
          return;
        }
      } else {
        // Capture current filter state before the default view resets it
        const prevFilters = filters;
        const prevTaskStatusFilter = taskStatusFilter;
        const prevEnrolStatusFilter = enrolStatusFilter;
        const prevTaskFilterOp = taskFilterOp;
        const prevEnrolFilterOp = enrolFilterOp;
        const prevAdvFilterNodes = advFilterNodes;
        const prevAdvLogicOp = advLogicOp;
        // Restore default layout (columns, sort, widths), then reapply existing filters minus partnerships
        handleResetDefault();
        setFilters({ ...prevFilters, partnerships: false });
        setTaskStatusFilter(prevTaskStatusFilter);
        setEnrolStatusFilter(prevEnrolStatusFilter);
        setTaskFilterOp(prevTaskFilterOp);
        setEnrolFilterOp(prevEnrolFilterOp);
        setAdvFilterNodes(prevAdvFilterNodes);
        setAdvLogicOp(prevAdvLogicOp);
        setCurrentPage(1);
        return;
      }
    }
    // If the NPP view is active and a quick filter is toggled, reset to default view first.
    if (nppViewId && activeViewId === nppViewId) {
      handleResetDefault();
    }
    setFilters(current => ({ ...current, [key]: !current[key] }));
    setCurrentPage(1);
  };

  const setSort = (key: SortKey, dir: SortDir) => { setSortKey(key); setSortDir(dir); };

  const setColumnWidth = (key: SortKey) => (w: number | undefined) =>
    setColumnWidths(prev => {
      const next = { ...prev };
      if (w === undefined) delete next[key]; else next[key] = w;
      return next;
    });

  return (
    <>
    <div className="enrolment-wrapper">
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: 4 }}>
        <ViewsMenu
          views={savedViews}
          activeViewId={activeViewId}
          hasUnsavedChanges={hasUnsavedChanges}
          onSelectView={handleSelectViewAndClose}
          onSaveAsNew={handleSaveAsNew}
          onSaveCurrentView={handleSaveCurrentView}
          onResetDefault={handleResetDefaultAndClose}
          onDeleteView={handleDeleteViewAndClose}
          onRenameView={handleRenameView}
          viewsLoading={viewsLoading}
        />
      </div>
      {saveError && <p className="enrolment-error">{saveError}</p>}

      {loading && <p className="enrolment-loading">Loading…</p>}
      {error && <p className="enrolment-error">{error}</p>}

      {!loading && !error && (
        <>
          <div className="search-and-tools-row">
            <EnrollmentSearchBar
              value={searchQuery}
              onChange={(nextValue) => {
                setSearchQuery(nextValue);
                setCurrentPage(1);
                setSelectedIds(new Set());
              }}
            />
            <div style={{ display: 'flex', gap: '0.4rem', flexShrink: 0 }}>
              <button type="button" className="sa-filter-btn" onClick={() => setShowEditColumns(true)}>
                <Columns2 size={14} /> Edit columns
              </button>
              <button type="button" className="sa-filter-btn" onClick={() => setShowEditFilters(true)}>
                <Filter size={14} /> Edit filters
              </button>
              <button type="button" className="sa-filter-btn" onClick={handleClearAllFilters}>
                <FilterX size={14} /> Clear all filters
              </button>
              <button type="button" className="sa-filter-btn" onClick={handleRefresh} disabled={loading}>
                <RefreshCw size={14} />{loading ? 'Refreshing...' : 'Refresh'}
              </button>
            </div>
          </div>

          <div className="worklist-box">
            <div className="worklist-item">
              <Info size={14} className="worklist-icon" />
              <button className="worklist-link" onClick={() => {
                const nppView = savedViews.find(v => v.source === 'system' && /npp/i.test(v.name));
                if (nppView) {
                  handleSelectView(nppView.id);
                  setCurrentPage(1);
                }
              }}>
                New Participants: <strong>{rows.filter(r => r.vsi_isnewparticipant === true).length}</strong>
              </button>
            </div>
            <div className="worklist-item">
              <Info size={14} className="worklist-icon" />
              {activeRole === 'Verifier' ? (
                <button className="worklist-link" onClick={() => applyWorklistFilter('taskStatus', 'Supervisor')}>
                  Pending supervisor&rsquo;s approval: <strong>{rows.filter(r => r.vsi_taskstatus === 865520001).length}</strong>
                </button>
              ) : (
                <Link to="/supervisor-approval" className="worklist-link">
                  Pending supervisor&rsquo;s approval: <strong>{rows.filter(r => r.vsi_taskstatus === 865520001).length}</strong>
                </Link>
              )}
            </div>
          </div>

          <div className="sa-card">
            <div className="sa-card-header">
              <EnrolmentQuickFilters
                filters={filters}
                onToggleFilter={toggleFilter}
                activeAdvancedCount={countActiveNodes(advFilterNodes)}
              />
              <EnrolmentActionsBar
                hasSelection={selectedIds.size > 0}
                selectedCount={selectedIds.size}
                onOpenBulkNotices={() => setShowBulkModal(true)}
                onOpenAssign={() => setShowAssignModal(true)}
                onOpenReferToSupervisor={() => setShowSupervisorModal(true)}
                onOpenApproveCalculatedFees={() => setShowApproveFeesModal(true)}
              />
            </div>

          <EnrolmentDataTable
            allRowsCount={rows.length}
            pagedRows={pagedRows}
            visibleColumnKeys={visibleColumnKeys}
            allPageSelected={allPageSelected}
            somePageSelected={somePageSelected}
            onToggleSelectAll={toggleSelectAll}
            selectedIds={selectedIds}
            onToggleSelect={toggleSelect}
            onRangeSelect={rangeSelect}
            colDragIdx={colDragIdx}
            onColDragStart={handleColDragStart}
            onColDragOver={handleColDragOver}
            onColDragEnd={handleColDragEnd}
            taskStatusOptions={taskStatusOptions}
            taskStatusFilter={taskStatusFilter}
            taskFilterOp={taskFilterOp}
            onTaskStatusFilterChange={setTaskStatusFilterAndReset}
            onTaskFilterOperatorChange={setTaskFilterOpAndReset}
            enrolStatusOptions={enrolStatusOptions}
            enrolStatusFilter={enrolStatusFilter}
            enrolFilterOp={enrolFilterOp}
            onEnrolStatusFilterChange={setEnrolStatusFilterAndReset}
            onEnrolFilterOperatorChange={setEnrolFilterOpAndReset}
            yearOptions={yearOptions}
            yearFilter={yearFilter}
            onYearFilterChange={setYearFilterAndReset}
            ownerOptions={ownerOptions}
            ownerFilter={ownerFilter}
            onOwnerFilterChange={setOwnerFilterAndReset}
            ownerFilterShortcuts={currentUserDisplayName ? [{ label: 'Assigned to me', values: new Set([currentUserDisplayName]) }] : undefined}
            sortKey={sortKey}
            sortDir={sortDir}
            onSort={setSort}
            columnWidths={columnWidths}
            onColumnWidthChange={setColumnWidth}
            avatarUrls={avatarUrls}
            coreAppId={coreAppId}
            coreBaseUrl={coreBaseUrl}
          />

          <div className="dash-pagination">
            <span>
              {searchedRows.length === 0
                ? 'Showing 0 of 0 results'
                : `Showing ${Math.min((currentPage - 1) * PAGE_SIZE + 1, searchedRows.length)}\u2013${Math.min(currentPage * PAGE_SIZE, searchedRows.length)} of ${searchedRows.length} result${searchedRows.length !== 1 ? 's' : ''}`}
            </span>
            <div className="dash-pagination-controls">
              <button
                type="button"
                className="dash-page-btn"
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
              >
                &lsaquo; Previous
              </button>
              {(() => {
                const pages: (number | '...')[] = [];
                if (totalPages <= 5) {
                  for (let i = 1; i <= totalPages; i++) pages.push(i);
                } else {
                  // Always show first page
                  pages.push(1);
                  // Determine window of 3 middle pages centred on currentPage
                  let start = Math.max(2, currentPage - 1);
                  let end = Math.min(totalPages - 1, currentPage + 1);
                  // Shift window so it always shows 3 pages when possible
                  if (end - start < 2) {
                    if (start === 2) end = Math.min(totalPages - 1, start + 2);
                    else start = Math.max(2, end - 2);
                  }
                  if (start > 2) pages.push('...');
                  for (let i = start; i <= end; i++) pages.push(i);
                  if (end < totalPages - 1) pages.push('...');
                  // Always show last page
                  pages.push(totalPages);
                }
                return pages.map((p, idx) =>
                  p === '...'
                    ? <span key={`dots-${idx}`} className="dash-page-dots">&hellip;</span>
                    : <button
                        key={p}
                        type="button"
                        className={`dash-page-btn${p === currentPage ? ' active' : ''}`}
                        onClick={() => setCurrentPage(p)}
                      >{p}</button>
                );
              })()}
              <button
                type="button"
                className="dash-page-btn"
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
              >
                Next &rsaquo;
              </button>
            </div>
          </div>
          </div>{/* /sa-card */}

          {showApproveFeesModal && (
            <ApproveCalculatedFeesModal
              selectedIds={selectedIds}
              rows={rows}
              onClose={() => setShowApproveFeesModal(false)}
              onComplete={(updates) => {
                const updatesById = new Map(updates.map(update => [update.id, update]));
                setRows(prev => prev.map(r => {
                  const update = updatesById.get(r.vsi_participantprogramyearid);
                  if (!update) return r;
                  return {
                    ...r,
                    vsi_taskstatus: 865520003,
                    vsi_taskstatusapproveddate: update.approvedDate,
                  };
                }));
                setSelectedIds(new Set());
                addToast(`${updates.length} enrolment${updates.length === 1 ? '' : 's'} approved successfully.`);
              }}
              onError={(msg) => addToast(msg, 'error')}
            />
          )}
        </>
      )}

      {showEditColumns && (
        <EditColumnsPanel
          key={activeViewId ?? 'default'}
          visibleKeys={visibleColumnKeys}
          onApply={(keys) => { setVisibleColumnKeys(keys); setShowEditColumns(false); }}
          onCancel={() => setShowEditColumns(false)}
        />
      )}
      {showEditFilters && (
        <EditFiltersPanel
          key={activeViewId ?? 'default'}
          filterNodes={effectiveFilterNodes}
          logicOp={advLogicOp}
          onApply={(nodes, logic) => {
            setAdvFilterNodesAndReset(nodes);
            setAdvLogicOpAndReset(logic);
            // Consolidate: clear any quick filters + column filters that were merged in
            setFilters(f => ({ ...f, verifiedCalc: false, unverifiedCalc: false, fortyFiveDayLetter: false, partnerships: false }));
            setTaskStatusFilter(new Set());
            setEnrolStatusFilter(new Set());
            setShowEditFilters(false);
          }}
          onCancel={() => setShowEditFilters(false)}
        />
      )}
      {showBulkModal && (
        <BulkNoticesModal
          selectedIds={selectedIds}
          rows={rows}
          onClose={() => setShowBulkModal(false)}
          onSuccess={(message) => {
            setSelectedIds(new Set());
            addToast(message);
          }}
        />
      )}
      {/* Assign modal — logic to be implemented */}
      {showAssignModal && (
        <AssignOwnerModal
          selectedIds={selectedIds}
          rows={rows}
          onClose={() => setShowAssignModal(false)}
          onComplete={(assignedIds, ownerName) => {
            setRows(prev => prev.map(r =>
              assignedIds.includes(r.vsi_participantprogramyearid!)
                ? { ...r, owneridname: ownerName }
                : r
            ));
            setShowAssignModal(false);
            setSelectedIds(new Set());
            addToast(`${assignedIds.length} enrolment${assignedIds.length === 1 ? '' : 's'} assigned to ${ownerName}.`);
          }}
        />
      )}
      {showSupervisorModal && (
        <ReferToSupervisorModal
          selectedIds={selectedIds}
          rows={rows}
          onClose={() => setShowSupervisorModal(false)}
          onComplete={(updatedIds) => {
            setRows(prev => prev.map(r =>
              updatedIds.includes(r.vsi_participantprogramyearid)
                ? { ...r, vsi_taskstatus: 865520001 }
                : r
            ));
            setSelectedIds(new Set());
            clearSaCache();
            addToast(`${updatedIds.length} enrolment${updatedIds.length === 1 ? '' : 's'} referred to supervisor.`);
          }}
          onError={(msg) => addToast(msg, 'error')}
        />
      )}
    </div>
      <Toast toasts={toasts} onDismiss={dismissToast} />
    </>
  );
}
