import { useCallback, useEffect, useMemo, useState } from 'react';
import type {
  PersonalView,
  ViewPayload,
  SortKey,
  SortDir,
  FilterOperator,
  AdvFilterNode,
  LogicOp,
  QuickFilterState,
} from '../types/enrollment';
import { DEFAULT_VIEW_SNAPSHOT, USERQUERY_ENTITY, USERQUERY_TYPE } from '../constants/columns';
import { UserqueriesService } from '../generated/services/UserqueriesService';
import { SavedqueriesService } from '../generated/services/SavedqueriesService';
import { generateLayoutXml, generateFetchXml, userqueryToView, savedqueryToView, loadActiveViewId, saveActiveViewId, resolveEntityObjectTypeCode, setEntityObjectTypeCode } from '../utils/viewSerializer';
import { serializeFilterNodes, deserializeFilterNodes } from '../utils/filterTree';

export interface ViewState {
  visibleColumnKeys: SortKey[];
  columnWidths: Partial<Record<SortKey, number>>;
  sortKey: SortKey | null;
  sortDir: SortDir;
  filters: QuickFilterState;
  taskStatusFilter: Set<string>;
  enrolStatusFilter: Set<string>;
  taskFilterOp: FilterOperator;
  enrolFilterOp: FilterOperator;
  advFilterNodes: AdvFilterNode[];
  advLogicOp: LogicOp;
}

export function useViews(state: ViewState, setters: {
  setVisibleColumnKeys: (keys: SortKey[]) => void;
  setColumnWidths: (w: Partial<Record<SortKey, number>>) => void;
  setSortKey: (k: SortKey | null) => void;
  setSortDir: (d: SortDir) => void;
  setFilters: (f: QuickFilterState) => void;
  setTaskStatusFilter: (s: Set<string>) => void;
  setEnrolStatusFilter: (s: Set<string>) => void;
  setTaskFilterOp: (op: FilterOperator) => void;
  setEnrolFilterOp: (op: FilterOperator) => void;
  setAdvFilterNodes: (n: AdvFilterNode[]) => void;
  setAdvLogicOp: (op: LogicOp) => void;
  setYearFilter: (s: Set<string>) => void;
  setOwnerFilter: (s: Set<string>) => void;
}) {
  const [savedViews, setSavedViews] = useState<PersonalView[]>([]);
  const [viewsLoading, setViewsLoading] = useState(true);
  const [activeViewId, setActiveViewId] = useState<string | null>(() => loadActiveViewId());
  const [saveError, setSaveError] = useState<string | null>(null);

  const ensureRequiredColumns = useCallback((keys: SortKey[]): SortKey[] => {
    // Ensure 'flagged' is always the first column
    const without = keys.filter((k): k is Exclude<SortKey, 'flagged'> => k !== 'flagged');
    return ['flagged', ...without];
  }, []);

  const applyView = useCallback((view: ViewPayload) => {
    setters.setVisibleColumnKeys(ensureRequiredColumns(view.visibleColumnKeys));
    setters.setColumnWidths({ ...view.columnWidths });
    setters.setSortKey(view.sortKey);
    setters.setSortDir(view.sortDir);
    setters.setFilters({ ...DEFAULT_VIEW_SNAPSHOT.filters, ...view.filters });
    setters.setTaskStatusFilter(new Set(view.taskStatusFilter));
    setters.setEnrolStatusFilter(new Set(view.enrolStatusFilter));
    setters.setTaskFilterOp(view.taskFilterOp ?? 'equals');
    setters.setEnrolFilterOp(view.enrolFilterOp ?? 'equals');
    setters.setAdvFilterNodes(deserializeFilterNodes(view.advFilterNodes as unknown[]));
    setters.setAdvLogicOp(view.advLogicOp ?? 'AND');
    // Reset year/owner filters since they are not part of the saved view payload
    setters.setYearFilter(new Set());
    setters.setOwnerFilter(new Set());
  }, [setters, ensureRequiredColumns]);

  const captureCurrentSnapshot = useCallback((): ViewPayload => ({
    visibleColumnKeys: ensureRequiredColumns(state.visibleColumnKeys),
    columnWidths: { ...state.columnWidths },
    sortKey: state.sortKey,
    sortDir: state.sortDir,
    filters: { ...state.filters },
    taskStatusFilter: [...state.taskStatusFilter],
    enrolStatusFilter: [...state.enrolStatusFilter],
    taskFilterOp: state.taskFilterOp,
    enrolFilterOp: state.enrolFilterOp,
    advFilterNodes: serializeFilterNodes(state.advFilterNodes),
    advLogicOp: state.advLogicOp,
  }), [state, ensureRequiredColumns]);

  const hasUnsavedChanges = useMemo(() => {
    const current = JSON.stringify(captureCurrentSnapshot());
    if (activeViewId) {
      const view = savedViews.find(v => v.id === activeViewId);
      if (!view) return true;
      const savedSnapshot: ViewPayload = {
        visibleColumnKeys: ensureRequiredColumns(view.visibleColumnKeys),
        columnWidths: { ...view.columnWidths },
        sortKey: view.sortKey,
        sortDir: view.sortDir,
        filters: { ...DEFAULT_VIEW_SNAPSHOT.filters, ...view.filters },
        taskStatusFilter: [...view.taskStatusFilter],
        enrolStatusFilter: [...view.enrolStatusFilter],
        taskFilterOp: view.taskFilterOp,
        enrolFilterOp: view.enrolFilterOp,
        advFilterNodes: serializeFilterNodes(deserializeFilterNodes(view.advFilterNodes as unknown[])),
        advLogicOp: view.advLogicOp,
      };
      return current !== JSON.stringify(savedSnapshot);
    }
    return current !== JSON.stringify(DEFAULT_VIEW_SNAPSHOT);
  }, [captureCurrentSnapshot, activeViewId, savedViews, ensureRequiredColumns]);

  const loadViews = useCallback(async (applyActiveView = false) => {
    setViewsLoading(true);
    try {
      let personal: PersonalView[] = [];
      let system: PersonalView[] = [];

      const [uqResult, sqResult] = await Promise.allSettled([
        UserqueriesService.getAll({
          select: ['userqueryid', 'name', 'layoutjson', 'layoutxml', 'returnedtypecode', 'querytype'],
          filter: `returnedtypecode eq '${USERQUERY_ENTITY}'`,
        }),
        SavedqueriesService.getAll({
          select: ['savedqueryid', 'name', 'layoutjson', 'layoutxml', 'fetchxml', 'returnedtypecode', 'querytype'],
          filter: `returnedtypecode eq '${USERQUERY_ENTITY}'`,
        }),
      ]);

      // Ensure the entity ObjectTypeCode is resolved before we might need it for
      // layoutxml generation. Does nothing if already resolved from a previous call.
      await resolveEntityObjectTypeCode();

      if (uqResult.status === 'fulfilled') {
        personal = (uqResult.value.data ?? []).map(uq => userqueryToView(uq));
      } else {
        console.error('[Views] Failed to load personal views:', uqResult.reason);
      }
      if (sqResult.status === 'fulfilled') {
        const allSq = sqResult.value.data ?? [];
        // Opportunistically extract the integer object type code from the first
        // record if the SDK returns the raw OData numeric value.
        if (allSq.length > 0) {
          const rawCode = (allSq[0] as unknown as Record<string, unknown>)['returnedtypecode'];
          const num = Number(rawCode);
          if (!isNaN(num) && num > 0) {
            setEntityObjectTypeCode(num);
            console.log('[Views] Extracted ObjectTypeCode from savedquery OData response:', num);
          }
        }
        const mainViews = allSq.filter(sq => String(sq.querytype) === '0');
        system = mainViews.map(savedqueryToView);
      } else {
        console.error('[Views] Failed to load system views:', sqResult.reason);
      }

      const allViews = [...personal, ...system];
      setSavedViews(allViews);

      if (applyActiveView) {
        const lastId = loadActiveViewId();
        if (lastId) {
          const match = allViews.find(v => v.id === lastId);
          if (match) applyView(match);
          else { setActiveViewId(null); saveActiveViewId(null); }
        }
      }

      return allViews;
    } catch (err) {
      console.error('[Views] Unexpected error loading views:', err);
      return [];
    } finally {
      setViewsLoading(false);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Load views on mount
  useEffect(() => {
    loadViews(true);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleSelectView = useCallback((id: string | null) => {
    setActiveViewId(id);
    saveActiveViewId(id);
    if (id) {
      const view = savedViews.find(v => v.id === id);
      if (view) applyView(view);
    } else {
      applyView(DEFAULT_VIEW_SNAPSHOT);
    }
  }, [savedViews, applyView]);

  const handleSaveAsNew = useCallback(async (name: string) => {
    setSaveError(null);
    const snap = captureCurrentSnapshot();
    // Only send fields that Dataverse accepts on create.
    // statecode/statuscode are managed by Dataverse and must NOT be included.
    const payload = {
      name,
      returnedtypecode: USERQUERY_ENTITY,
      querytype: USERQUERY_TYPE,
      fetchxml: generateFetchXml(snap.visibleColumnKeys, snap.advFilterNodes as unknown[]),
      layoutjson: JSON.stringify(snap),
      layoutxml: generateLayoutXml(snap.visibleColumnKeys, snap.columnWidths),
    };
    try {
      console.log('[Views] Creating new view:', name);
      const result = await UserqueriesService.create(payload as unknown as Parameters<typeof UserqueriesService.create>[0]);
      console.log('[Views] Create result — success:', result.success, 'data:', result.data, 'error:', result.error);

      // The SDK returns { success: false, error } without throwing on API errors.
      // Must check result.success explicitly since a failed create does not throw.
      if (!result.success) {
        const errMsg = result.error instanceof Error
          ? result.error.message
          : (result.error as { message?: string })?.message ?? JSON.stringify(result.error);
        console.error('[Views] Create failed (non-throwing):', result.error);
        setSaveError(`Failed to save view: ${errMsg}`);
        return;
      }

      const created = result.data;
      // Reload views from Dataverse to ensure the new view appears regardless
      // of whether result.data is populated (SDK typically returns null on create)
      const allViews = await loadViews(false);
      const newId = created?.userqueryid;
      if (newId) {
        const newView = allViews.find(v => v.id === newId) ?? { id: newId, name, source: 'personal' as const, ...snap };
        setSavedViews(prev => prev.some(v => v.id === newId) ? prev : [...prev, newView]);
        setActiveViewId(newId);
        saveActiveViewId(newId);
        applyView(newView);
      } else {
        // newId not in result — look for the newly named view in the reloaded list
        const match = allViews.find(v => v.name === name && v.source === 'personal');
        if (match) {
          setActiveViewId(match.id);
          saveActiveViewId(match.id);
          applyView(match);
        } else {
          console.warn('[Views] New view not found after reload — name:', name, 'all personal:', allViews.filter(v => v.source === 'personal').map(v => v.name));
        }
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      console.error('[Views] Failed to create view:', e);
      setSaveError(`Failed to save view: ${msg}`);
    }
  }, [captureCurrentSnapshot, loadViews, applyView]);

  const handleSaveCurrentView = useCallback(async () => {
    if (!activeViewId) return;
    const view = savedViews.find(v => v.id === activeViewId);
    if (!view || view.source !== 'personal') return;
    const snap = captureCurrentSnapshot();
    try {
      await UserqueriesService.update(activeViewId, {
        layoutjson: JSON.stringify(snap),
        layoutxml: generateLayoutXml(snap.visibleColumnKeys, snap.columnWidths),
        fetchxml: generateFetchXml(snap.visibleColumnKeys, snap.advFilterNodes as unknown[]),
      });
      setSavedViews(prev => prev.map(v => v.id === activeViewId ? { ...v, ...snap } : v));
      applyView({ ...view, ...snap });
    } catch (e) {
      console.error('Failed to update view:', e);
    }
  }, [activeViewId, savedViews, captureCurrentSnapshot, applyView]);

  const handleDeleteView = useCallback(async (id: string) => {
    const view = savedViews.find(v => v.id === id);
    if (!view || view.source !== 'personal') return;
    try {
      await UserqueriesService.delete(id);
      setSavedViews(prev => prev.filter(v => v.id !== id));
      if (activeViewId === id) {
        setActiveViewId(null);
        saveActiveViewId(null);
        applyView(DEFAULT_VIEW_SNAPSHOT);
      }
    } catch (e) {
      console.error('Failed to delete view:', e);
    }
  }, [savedViews, activeViewId, applyView]);

  const handleRenameView = useCallback(async (id: string, name: string) => {
    if (!name) return;
    const view = savedViews.find(v => v.id === id);
    if (!view || view.source !== 'personal') return;
    try {
      await UserqueriesService.update(id, { name });
      setSavedViews(prev => prev.map(v => v.id === id ? { ...v, name } : v));
    } catch (e) {
      console.error('Failed to rename view:', e);
    }
  }, [savedViews]);

  const handleResetDefault = useCallback(() => {
    setActiveViewId(null);
    saveActiveViewId(null);
    applyView(DEFAULT_VIEW_SNAPSHOT);
  }, [applyView]);

  return {
    savedViews,
    viewsLoading,
    activeViewId,
    hasUnsavedChanges,
    saveError,
    handleSelectView,
    handleSaveAsNew,
    handleSaveCurrentView,
    handleDeleteView,
    handleRenameView,
    handleResetDefault,
    reloadViews: loadViews,
  };
}
