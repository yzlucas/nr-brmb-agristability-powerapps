// In-memory cache for enrolment rows (persists while app is open)
let enrolmentRowsCache: Vsi_participantprogramyears[] | null = null;
let coreAppIdCache: string | null = null;
let coreBaseUrlCache: string | null = null;
let coreAppIdLoaded = false;
import { useCallback, useEffect, useMemo, useState } from 'react';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import {
  Vsi_participantprogramyearsvsi_enrolmentstatus,
  Vsi_participantprogramyearsvsi_taskstatus,
} from '../generated/models/Vsi_participantprogramyearsModel';
import { Vsi_armsconfigurationsService } from '../generated/services/Vsi_armsconfigurationsService';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { Office365UsersService } from '../generated/services/Office365UsersService';
import type {
  SortKey,
  SortDir,
  FilterOperator,
  AdvFilterNode,
  AdvFilterField,
  LogicOp,
  QuickFilterState,
} from '../types/enrollment';
import { ADV_FIELD_OPTIONS } from '../constants/columns';
import { getEnrolmentStatusLabel, getTaskStatusLabel, getSortValue } from '../utils/helpers';
import { isNodeActive } from '../utils/filterTree';

const FLAGGED_VARIANCE_THRESHOLD = 20;

function normalizeCoreBaseUrl(url: string | null | undefined) {
  const trimmed = url?.trim();
  if (!trimmed) return null;
  return /\/main\.aspx(?:$|[?#])/i.test(trimmed) ? trimmed : `${trimmed.replace(/\/$/, '')}/main.aspx`;
}


// Patch specific records in the in-memory cache by enrolment ID.
// Called by other pages (e.g. SupervisorApprovalPage) after mutating enrolment fields
// so the dashboard table reflects the change without a full reload.
export function patchEnrolmentCache(patches: Array<{ id: string; fields: Partial<Vsi_participantprogramyears> }>) {
  if (!enrolmentRowsCache) return;
  const patchMap = new Map(patches.map(p => [p.id.replace(/[{}]/g, '').toLowerCase(), p.fields]));
  enrolmentRowsCache = enrolmentRowsCache.map(row => {
    const rowId = row.vsi_participantprogramyearid?.replace(/[{}]/g, '').toLowerCase();
    if (!rowId) return row;
    const patch = patchMap.get(rowId);
    return patch ? { ...row, ...patch } : row;
  });
}

export function useEnrolmentData() {
  const [rows, setRows] = useState<Vsi_participantprogramyears[]>(() => enrolmentRowsCache || []);
  const [loading, setLoading] = useState(() => enrolmentRowsCache === null);
  const [error, setError] = useState<string | null>(null);
  const [avatarUrls, setAvatarUrls] = useState<Record<string, string>>({});
  const [coreAppId, setCoreAppId] = useState<string | null>(() => (coreAppIdLoaded ? coreAppIdCache : null));
  const [coreBaseUrl, setCoreBaseUrl] = useState<string | null>(() => (coreAppIdLoaded ? coreBaseUrlCache : null));

  const fetchCoreAppId = useCallback(async () => {
    try {
      const result = await Vsi_armsconfigurationsService.getAll({
        maxPageSize: 50,
        select: ['cr4dd_coreappid', 'vsi_coreenvironmenturl'],
      });
      const configRows = result.data ?? [];
      const nextCoreAppId = configRows
        .map(row => row.cr4dd_coreappid?.trim())
        .find((candidate): candidate is string => !!candidate) ?? null;
      const nextCoreBaseUrl = configRows
        .map(row => normalizeCoreBaseUrl(row.vsi_coreenvironmenturl))
        .find((candidate): candidate is string => !!candidate) ?? null;
      setCoreAppId(nextCoreAppId);
      setCoreBaseUrl(nextCoreBaseUrl);
      coreAppIdCache = nextCoreAppId;
      coreBaseUrlCache = nextCoreBaseUrl;
      coreAppIdLoaded = true;
    } catch {
      if (!coreAppIdLoaded) {
        setCoreAppId(null);
        setCoreBaseUrl(null);
        coreAppIdCache = null;
        coreBaseUrlCache = null;
        coreAppIdLoaded = true;
      }
    }
  }, []);

  // Fetch function (used for initial load and manual refresh)
  const fetchEnrolments = async () => {
    setLoading(true);
    setError(null);
    let cancelled = false;
    try {
      const allRows: Vsi_participantprogramyears[] = [];
      let skipToken: string | undefined;
      const baseOptions = {
        maxPageSize: 5000,
        select: [
          'vsi_name',
          '_vsi_participantid_value',
          '_vsi_programyearid_value',
          'vsi_enrolmentstatus',
          'vsi_taskstatus',
          'vsi_calculatedenfee',
          'vsi_previousyearcalculatedenfee',
          'vsi_enrolmentfeecalculated',
          'vsi_totalfeesowed',
          'vsi_totalfeespaid',
          'vsi_enrolmentfee',
          'vsi_latepaymentfee',
          'vsi_haspartners',
          'vsi_incombinedfarm',
          'vsi_sharepointdocumentfolder',
          'modifiedon',
          '_ownerid_value',
          'vsi_enrollmentregionaloffice',
          'vsi_farmingsector',
          'vsi_bringforward',
          'vsi_broughtforward',
          'vsi_manualreview',
          'vsi_enrolmentnoticesentdate',
          'vsi_filereceiveddate',
          'vsi_enrolmentfeespaiddate',
          'vsi_prevyearpartnotverified',
          'vsi_variancecalculation',
        ],
        orderBy: ['vsi_taskstatus desc'],
      };
      do {
        const result = await Vsi_participantprogramyearsService.getAll({
          ...baseOptions,
          ...(skipToken ? { skipToken } : {}),
        });
        if (cancelled) return;
        allRows.push(...(result.data ?? []));
        const raw = result as unknown as Record<string, unknown>;
        skipToken = (raw['skipToken'] ?? raw['@odata.nextLink']) as string | undefined;
      } while (skipToken);
      if (!cancelled) {
        setRows(allRows);
        enrolmentRowsCache = allRows;
      }
    } catch (e: unknown) {
      console.error('Error fetching enrolments:', e);
      if (!cancelled) setError(e instanceof Error ? e.message : 'Failed to load enrolments');
    } finally {
      if (!cancelled) setLoading(false);
    }
    return () => { cancelled = true; };
  };

  useEffect(() => {
    if (enrolmentRowsCache !== null) return;
    fetchEnrolments();
  }, []);

  useEffect(() => {
    if (coreAppIdLoaded) return;
    fetchCoreAppId();
  }, [fetchCoreAppId]);

  // Fetch avatar photos
  useEffect(() => {
    if (rows.length === 0) return;
    const ids = new Set<string>();
    for (const row of rows) {
      const raw = row as unknown as Record<string, unknown>;
      const ownerUid = raw['_ownerid_value'] as string | undefined;
      if (ownerUid) ids.add(ownerUid);
    }
    let cancelled = false;
    (async () => {
      const photos: Record<string, string> = {};
      await Promise.all([...ids].map(async (uid) => {
        try {
          const result = await Office365UsersService.UserPhoto_V2(uid);
          if (!cancelled && result.data) {
            photos[uid] = result.data;
          }
        } catch { /* no photo available */ }
      }));
      if (!cancelled) setAvatarUrls(photos);
    })();
    return () => { cancelled = true; };
  }, [rows]);

  return { rows, setRows, loading, error, avatarUrls, fetchEnrolments, coreAppId, coreBaseUrl, fetchCoreAppId };
}

export function useSortedAndFilteredRows(
  rows: Vsi_participantprogramyears[],
  sortKey: SortKey | null,
  sortDir: SortDir,
  filters: QuickFilterState,
  taskStatusFilter: Set<string>,
  enrolStatusFilter: Set<string>,
  yearFilter: Set<string>,
  ownerFilter: Set<string>,
  taskFilterOp: FilterOperator,
  enrolFilterOp: FilterOperator,
  advFilterNodes: AdvFilterNode[],
  advLogicOp: LogicOp,
  currentUserName?: string,
) {
  const taskStatusOptions = useMemo(() =>
    Object.values(Vsi_participantprogramyearsvsi_taskstatus) as string[],
  []);
  const enrolStatusOptions = useMemo(() =>
    Object.values(Vsi_participantprogramyearsvsi_enrolmentstatus) as string[],
  []);

  const yearOptions = useMemo(() => {
    const seen = new Set<string>();
    for (const row of rows) {
      const raw = row as unknown as Record<string, unknown>;
      const name = (row.vsi_programyearidname
        ?? raw['_vsi_programyearid_value@OData.Community.Display.V1.FormattedValue']
        ?? '') as string;
      if (name) seen.add(name);
    }
    return [...seen].sort();
  }, [rows]);

  const ownerOptions = useMemo(() => {
    const seen = new Set<string>();
    for (const row of rows) {
      const raw = row as unknown as Record<string, unknown>;
      const name = (row.owneridname
        ?? raw['_ownerid_value@OData.Community.Display.V1.FormattedValue']
        ?? '') as string;
      if (name) seen.add(name);
    }
    const sorted = [...seen].sort((a, b) => a.localeCompare(b));
    if (currentUserName && seen.has(currentUserName)) {
      return [currentUserName, ...sorted.filter(n => n !== currentUserName)];
    }
    return sorted;
  }, [rows, currentUserName]);

  const getRowFieldValue = useCallback((row: Vsi_participantprogramyears, field: AdvFilterField): string => {
    const raw = row as unknown as Record<string, unknown>;
    switch (field) {
      case 'taskStatus': return getTaskStatusLabel(row.vsi_taskstatus);
      case 'enrolStatus': return getEnrolmentStatusLabel(row.vsi_enrolmentstatus);
      case 'pin': return row.vsi_name ?? '';
      case 'producer':
        return (row.vsi_participantidname
          ?? raw['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue']
          ?? '') as string;
      case 'fee': return String(row.vsi_calculatedenfee ?? '');
      case 'hasPartners': return row.vsi_haspartners === true ? 'Yes' : 'No';
      case 'inCombinedFarm': return row.vsi_incombinedfarm === true ? 'Yes' : 'No';
    }
  }, []);

  const matchAdvRow = useCallback((row: Vsi_participantprogramyears, fr: { kind: 'row'; field: AdvFilterField; operator: string; values: Set<string>; textValue: string }): boolean => {
    const val = getRowFieldValue(row, fr.field);
    const fieldType = ADV_FIELD_OPTIONS[fr.field];
    if (fieldType === 'choice') {
      if (fr.values.size === 0) return true;
      const inSet = fr.values.has(val);
      return fr.operator === 'equals' ? inSet : !inSet;
    }
    if (!fr.textValue) return true;
    const lower = val.toLowerCase();
    const search = fr.textValue.toLowerCase();
    switch (fr.operator) {
      case 'equals': return lower === search;
      case 'notEquals': return lower !== search;
      case 'contains': return lower.includes(search);
      case 'notContains': return !lower.includes(search);
      case 'beginsWith': return lower.startsWith(search);
      case 'endsWith': return lower.endsWith(search);
      default: return true;
    }
  }, [getRowFieldValue]);

  const matchAdvNode = useCallback((row: Vsi_participantprogramyears, node: AdvFilterNode): boolean => {
    const evaluate = (currentNode: AdvFilterNode): boolean => {
      if (currentNode.kind === 'row') return matchAdvRow(row, currentNode);
      const activeChildren = currentNode.children.filter(isNodeActive);
      if (activeChildren.length === 0) return true;
      if (currentNode.logic === 'AND') return activeChildren.every(ch => evaluate(ch));
      return activeChildren.some(ch => evaluate(ch));
    };
    return evaluate(node);
  }, [matchAdvRow]);

  const sortedRows = useMemo(() => {
    if (!sortKey) return rows;
    return [...rows].sort((a, b) => {
      const va = getSortValue(a, sortKey);
      const vb = getSortValue(b, sortKey);
      let cmp: number;
      if (typeof va === 'number' && typeof vb === 'number') {
        cmp = va - vb;
      } else {
        cmp = String(va).localeCompare(String(vb));
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [rows, sortKey, sortDir]);

  const anyFilter = filters.verifiedCalc || filters.unverifiedCalc || filters.flagged || filters.partnerships || filters.fortyFiveDayLetter;

  const isYesValue = useCallback((value: unknown): boolean => {
    if (value === true || value === 1 || value === '1') return true;
    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      return normalized === 'yes' || normalized === 'true';
    }
    return false;
  }, []);

  const isReadyTaskStatus = useCallback((value: unknown): boolean => {
    if (typeof value === 'number') return value === 865520002;
    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      return normalized === '865520002' || normalized === 'ready';
    }
    return false;
  }, []);

  const isFlaggedByVariance = useCallback((row: Vsi_participantprogramyears): boolean => {
    if (row.vsi_prevyearpartnotverified === true) return true;
    if (row.vsi_calculatedenfee != null && row.vsi_previousyearcalculatedenfee == null) return true;
    const variance = row.vsi_variancecalculation != null ? row.vsi_variancecalculation * 100 : null;
    if (variance == null) return false;
    return Math.abs(variance) > FLAGGED_VARIANCE_THRESHOLD;
  }, []);

  const filteredRows = useMemo(() => {
    let result = sortedRows;

    if (anyFilter) {
      result = result.filter(row => {
        const isEnCalc = isYesValue(row.vsi_enrolmentfeecalculated);
        const isReady = isReadyTaskStatus(row.vsi_taskstatus);
        const matchesVerifiedCalc = isReady && isEnCalc;
        const matchesUnverifiedCalc = !isReady && isEnCalc;
        const matchesFlagged = isFlaggedByVariance(row);
        const matchesPartnerships = isYesValue(row.vsi_haspartners) || isYesValue(row.vsi_incombinedfarm);
        const matchesFortyFiveDayLetter = getEnrolmentStatusLabel(row.vsi_enrolmentstatus) === '_45DayLetter';

        if (filters.verifiedCalc && !matchesVerifiedCalc) return false;
        if (filters.unverifiedCalc && !matchesUnverifiedCalc) return false;
        if (filters.flagged && !matchesFlagged) return false;
        if (filters.partnerships && !matchesPartnerships) return false;
        if (filters.fortyFiveDayLetter && !matchesFortyFiveDayLetter) return false;
        return true;
      });
    }

    if (taskStatusFilter.size > 0) {
      result = result.filter(row => {
        const match = taskStatusFilter.has(getTaskStatusLabel(row.vsi_taskstatus));
        return taskFilterOp === 'equals' ? match : !match;
      });
    }
    if (enrolStatusFilter.size > 0) {
      result = result.filter(row => {
        const match = enrolStatusFilter.has(getEnrolmentStatusLabel(row.vsi_enrolmentstatus));
        return enrolFilterOp === 'equals' ? match : !match;
      });
    }

    if (yearFilter.size > 0) {
      result = result.filter(row => {
        const raw = row as unknown as Record<string, unknown>;
        const name = (row.vsi_programyearidname
          ?? raw['_vsi_programyearid_value@OData.Community.Display.V1.FormattedValue']
          ?? '') as string;
        return yearFilter.has(name);
      });
    }

    if (ownerFilter.size > 0) {
      result = result.filter(row => {
        const raw = row as unknown as Record<string, unknown>;
        const name = (row.owneridname
          ?? raw['_ownerid_value@OData.Community.Display.V1.FormattedValue']
          ?? '') as string;
        return ownerFilter.has(name);
      });
    }

    const activeAdvNodes = advFilterNodes.filter(isNodeActive);
    if (activeAdvNodes.length > 0) {
      result = result.filter(row => {
        if (advLogicOp === 'AND') return activeAdvNodes.every(n => matchAdvNode(row, n));
        return activeAdvNodes.some(n => matchAdvNode(row, n));
      });
    }

    return result;
  }, [sortedRows, filters, anyFilter, taskStatusFilter, enrolStatusFilter, yearFilter, ownerFilter, taskFilterOp, enrolFilterOp, advFilterNodes, advLogicOp, matchAdvNode, isYesValue, isFlaggedByVariance]);

  return { filteredRows, taskStatusOptions, enrolStatusOptions, yearOptions, ownerOptions };
}
