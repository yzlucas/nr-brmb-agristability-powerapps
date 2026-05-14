import { useCallback, useMemo, useState, type DragEvent } from 'react';
import { Link } from 'react-router-dom';
import { Columns2, Filter, Info } from 'lucide-react';

import type { SortKey, SortDir, FilterOperator, AdvFilterNode, LogicOp, QuickFilterState } from '../types/enrollment';
import { DEFAULT_VISIBLE_KEYS } from '../constants/columns';
import { countActiveNodes } from '../utils/filterTree';
import { useEnrolmentData, useSortedAndFilteredRows } from '../hooks/useEnrolmentData';
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

export function DashboardHomePage() {
  const { rows, setRows, loading, error, avatarUrls, fetchEnrolments, coreAppId, coreBaseUrl, fetchCoreAppId } = useEnrolmentData();

  // Refresh handler is defined after useViews so reloadViews is available

  // Column & sort state
  const [visibleColumnKeys, setVisibleColumnKeys] = useState<SortKey[]>([...DEFAULT_VISIBLE_KEYS]);
  const [columnWidths, setColumnWidths] = useState<Partial<Record<SortKey, number>>>({});
  const [sortKey, setSortKey] = useState<SortKey | null>('modifiedOn');
  const [sortDir, setSortDir] = useState<SortDir>('desc');

  // Filter state
  const [filters, setFilters] = useState<QuickFilterState>({
    verifiedCalc: false,
    unverifiedCalc: false,
    flagged: false,
    partnerships: false,
    fortyFiveDayLetter: false,
  });
  const [searchQuery, setSearchQuery] = useState('');
  const [taskStatusFilter, setTaskStatusFilter] = useState<Set<string>>(new Set());
  const [enrolStatusFilter, setEnrolStatusFilter] = useState<Set<string>>(new Set());
  const [yearFilter, setYearFilter] = useState<Set<string>>(new Set());
  const [ownerFilter, setOwnerFilter] = useState<Set<string>>(new Set());
  const [taskFilterOp, setTaskFilterOp] = useState<FilterOperator>('equals');
  const [enrolFilterOp, setEnrolFilterOp] = useState<FilterOperator>('equals');
  const [advFilterNodes, setAdvFilterNodes] = useState<AdvFilterNode[]>([]);
  const [advLogicOp, setAdvLogicOp] = useState<LogicOp>('AND');

  // Pagination & selection
  const [currentPage, setCurrentPage] = useState(1);
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
  }, []);
  const setTaskStatusFilterAndReset = useCallback((next: Set<string>) => {
    setTaskStatusFilter(next);
    setCurrentPage(1);
  }, []);
  const setEnrolStatusFilterAndReset = useCallback((next: Set<string>) => {
    setEnrolStatusFilter(next);
    setCurrentPage(1);
  }, []);
  const setYearFilterAndReset = useCallback((next: Set<string>) => {
    setYearFilter(next);
    setCurrentPage(1);
  }, []);
  const setOwnerFilterAndReset = useCallback((next: Set<string>) => {
    setOwnerFilter(next);
    setCurrentPage(1);
  }, []);
  const setTaskFilterOpAndReset = useCallback((next: FilterOperator) => {
    setTaskFilterOp(next);
    setCurrentPage(1);
  }, []);
  const setEnrolFilterOpAndReset = useCallback((next: FilterOperator) => {
    setEnrolFilterOp(next);
    setCurrentPage(1);
  }, []);
  const setAdvFilterNodesAndReset = useCallback((next: AdvFilterNode[]) => {
    setAdvFilterNodes(next);
    setCurrentPage(1);
  }, []);
  const setAdvLogicOpAndReset = useCallback((next: LogicOp) => {
    setAdvLogicOp(next);
    setCurrentPage(1);
  }, []);

  /** Clear every filter, then optionally apply a single task-status or enrol-status filter. */
  const applyWorklistFilter = useCallback((
    type: 'taskStatus' | 'enrolStatus',
    label: string,
  ) => {
    const is45Day = label === '_45DayLetter';
    setFilters({ verifiedCalc: false, unverifiedCalc: false, flagged: false, partnerships: false, fortyFiveDayLetter: is45Day });
    setSearchQuery('');
    setTaskStatusFilter(type === 'taskStatus' ? new Set([label]) : new Set());
    setEnrolStatusFilter(type === 'enrolStatus' ? new Set([label]) : new Set());
    setYearFilter(new Set());
    setOwnerFilter(new Set());
    setAdvFilterNodes([]);
    setCurrentPage(1);
  }, []);

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
          handleSelectView(partnerView.id);
          setCurrentPage(1);
          return; // handleSelectView sets filters via applyView — don't also call setFilters
        }
      } else {
        handleResetDefault();
        setCurrentPage(1);
        return;
      }
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
              }}
            />
            <div style={{ display: 'flex', gap: '0.4rem', flexShrink: 0 }}>
              <button type="button" className="sa-filter-btn" onClick={() => setShowEditColumns(true)}>
                <Columns2 size={14} /> Edit columns
              </button>
              <button type="button" className="sa-filter-btn" onClick={() => setShowEditFilters(true)}>
                <Filter size={14} /> Edit filters
              </button>
              <button type="button" className="sa-filter-btn" onClick={handleRefresh} disabled={loading}>
                {loading ? 'Refreshing...' : 'Refresh'}
              </button>
            </div>
          </div>

          <div className="worklist-box">
            <div className="worklist-item">
              <Info size={15} className="worklist-icon" />
              <button className="worklist-link" onClick={() => applyWorklistFilter('enrolStatus', 'Initialized')}>
                New Participants: <strong>{rows.filter(r => r.vsi_enrolmentstatus === 865520004).length}</strong>
              </button>
            </div>
            <div className="worklist-item">
              <Info size={15} className="worklist-icon" />
              <Link to="/supervisor-approval" className="worklist-link">
                Pending supervisor&rsquo;s approval: <strong>{rows.filter(r => r.vsi_taskstatus === 865520001).length}</strong>
              </Link>
            </div>
            <div className="worklist-item">
              <Info size={15} className="worklist-icon" />
              <button className="worklist-link" onClick={() => applyWorklistFilter('enrolStatus', '_45DayLetter')}>
                45 day letters: <strong>{rows.filter(r => r.vsi_enrolmentstatus === 865520010).length}</strong>
              </button>
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
          filterNodes={advFilterNodes}
          logicOp={advLogicOp}
          onApply={(nodes, logic) => {
            setAdvFilterNodesAndReset(nodes);
            setAdvLogicOpAndReset(logic);
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
          onSuccess={(message) => addToast(message)}
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
