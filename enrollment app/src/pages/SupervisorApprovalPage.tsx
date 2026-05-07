import { resolveCurrentSystemUser } from '../utils/currentUser';
import { patchEnrolmentCache } from '../hooks/useEnrolmentData';
import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent } from 'react';
import { Filter, Calculator, ClipboardList, LogOut, UserPlus, CircleCheck, Wrench } from 'lucide-react';
import { Link } from 'react-router-dom';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import { Vsi_participantprogramyearsvsi_enrolmentstatus } from '../generated/models/Vsi_participantprogramyearsModel';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { QueueitemsService } from '../generated/services/QueueitemsService';
import { QueuesService } from '../generated/services/QueuesService';
import { Office365UsersService } from '../generated/services/Office365UsersService';
import { ColumnHeaderMenu } from '../components/ColumnHeaderMenu';

import { enrolmentStatusClass, formatCurrencyOr, formatVariancePercent, getEnrolmentStatusLabel, getInitials, getTaskStatusLabel, getVarianceClass, getAvatarColor } from '../utils/helpers';
import { AssignWorkerModal } from '../components/AssignWorkerModal';
import { ApprovalErrorModal } from '../components/ApprovalErrorModal';
import ManualErrorModal from '../components/ManualErrorModal';
import { ConfirmActionModal } from '../components/ConfirmActionModal';
import { Toast, nextToastId } from '../components/Toast';
import type { ToastMessage } from '../components/Toast';
import type { FilterOperator, SortDir } from '../types/enrollment';
import '../styles/supervisor-approval.css';

const CORE_APP_ID_FALLBACK = '88c024d9-9fd5-ec11-a7b5-002248ada475';
const CORE_BASE_URL_FALLBACK = 'https://aff-brmb-crm-dev.crm3.dynamics.com/main.aspx';

const PAGE_SIZE = 20;
const SUPERVISOR_QUEUE_NAME = 'Supervisor Approval Queue';
const ENROLMENT_STATUS_FILTER_OPTIONS = Object.values(Vsi_participantprogramyearsvsi_enrolmentstatus)
  .filter((label): label is (typeof Vsi_participantprogramyearsvsi_enrolmentstatus)[keyof typeof Vsi_participantprogramyearsvsi_enrolmentstatus] => typeof label === 'string' && label.length > 0)
  .sort((a, b) => a.localeCompare(b));

// Module-level cache so navigating away and back does not trigger a full reload.
// Cleared when the user explicitly clicks Refresh.
let saItemsCache: Vsi_participantprogramyears[] | null = null;
let saQueueWorkCache: Record<string, QueueWorkMeta> | null = null;
let saSupervisorQueueIdsCache: Set<string> | null = null;
let saWorkerAvatarUrlsCache: Record<string, string> | null = null;

type QueueWorkMeta = {
  workedBy: string;
  workedOn: string;
  workedOnRaw?: string;
  enteredQueue?: string;
  enteredQueueRaw?: string;
  workerId?: string;
  queueitemId?: string;
  queueId?: string;
  queueName?: string;
  isActive?: boolean;
};

type SupervisorColumnKey = 'enrolmentName' | 'participant' | 'taskStatus' | 'enrolmentStatus' | 'calculatedFee' | 'enteredQueue' | 'workedBy' | 'workedOn';

type SupervisorColumnDef = {
  key: SupervisorColumnKey;
  label: string;
};

const SUPERVISOR_COLUMNS: SupervisorColumnDef[] = [
  { key: 'enrolmentName', label: 'Enrolment Name' },
  { key: 'participant', label: 'Participant' },
  { key: 'taskStatus', label: 'Task Status' },
  { key: 'enrolmentStatus', label: 'Enrolment Status' },
  { key: 'calculatedFee', label: 'Calculated Fee' },
  { key: 'enteredQueue', label: 'Entered Queue' },
  { key: 'workedBy', label: 'Worked By' },
  { key: 'workedOn', label: 'Worked On' },
];

const DEFAULT_COLUMN_ORDER: SupervisorColumnKey[] = SUPERVISOR_COLUMNS.map(c => c.key);
const DEFAULT_FILTER_OPS: Record<SupervisorColumnKey, FilterOperator> = {
  enrolmentName: 'equals',
  participant: 'equals',
  taskStatus: 'equals',
  enrolmentStatus: 'equals',
  calculatedFee: 'equals',
  enteredQueue: 'equals',
  workedBy: 'equals',
  workedOn: 'equals',
};

const DEFAULT_COLUMN_WIDTHS: Partial<Record<SupervisorColumnKey, number>> = {
  workedBy: 120,
  workedOn: 110,
};

const createEmptyFilters = (): Record<SupervisorColumnKey, Set<string>> => ({
  enrolmentName: new Set(),
  participant: new Set(),
  taskStatus: new Set(),
  enrolmentStatus: new Set(),
  calculatedFee: new Set(),
  enteredQueue: new Set(),
  workedBy: new Set(),
  workedOn: new Set(),
});

type AssignTarget = {
  enrolmentId: string;
  enrolmentName: string;
  queueitemId: string | undefined;
  queueId?: string;
  queueName?: string;
  bulkRows?: Array<{ itemId: string; queueitemId: string | undefined }>;
};

function VariancePill({ variance }: { variance: number }) {
  const cls = getVarianceClass(variance);
  const text = formatVariancePercent(variance);
  return <span className={`variance-pill ${cls}`}>{text}</span>;
}

import { taskStatusIcon } from '../utils/helpers';

function StatusBadge({ status }: { status?: number }) {
  const label = getTaskStatusLabel(status) || 'Unknown';
  const clsByLabel: Record<string, string> = {
    Manual: 'pending',
    Supervisor: 'review',
    Ready: 'inprogress',
    Approved: 'approved',
  };
  const cls = clsByLabel[label] ?? 'pending';

  return (
    <span className={`sa-status-badge ${cls}`}>
      {taskStatusIcon(label)}
      <span style={{ marginLeft: 4 }}>{label}</span>
    </span>
  );
}

function getRowId(item: Vsi_participantprogramyears): string | null {
  return item.vsi_participantprogramyearid ?? null;
}

function formatWorkedOn(value?: string): string {
  if (!value) return '—';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toLocaleDateString();
}

function normalizeGuid(value?: string | null): string {
  return (value ?? '').replace(/[{}]/g, '').trim().toLowerCase();
}

function isSupervisorApprovalQueueName(name?: string): boolean {
  if (!name) return false;
  const n = name.toLowerCase();
  return n.includes('supervisor') && n.includes('approval');
}

export function SupervisorApprovalPage() {
  const [items, setItems] = useState<Vsi_participantprogramyears[]>(() => saItemsCache ?? []);
  const [queueWorkByEnrolmentId, setQueueWorkByEnrolmentId] = useState<Record<string, QueueWorkMeta>>(() => saQueueWorkCache ?? {});
  const [supervisorQueueIds, setSupervisorQueueIds] = useState<Set<string>>(() => saSupervisorQueueIdsCache ?? new Set());
  const [workerAvatarUrls, setWorkerAvatarUrls] = useState<Record<string, string>>(() => saWorkerAvatarUrlsCache ?? {});
  const fetchedQueueIds = useRef<Set<string>>(new Set(saItemsCache?.map(i => normalizeGuid(i.vsi_participantprogramyearid)) ?? []));
  const [loading, setLoading] = useState(saItemsCache === null);
  const [error, setError] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [page, setPage] = useState(1);
  const [assignTarget, setAssignTarget] = useState<AssignTarget | null>(null);
  const [currentUser, setCurrentUser] = useState<{ systemUserId: string; displayName: string } | null>(null);
  const [approvingBulk, setApprovingBulk] = useState(false);
  const [pickingBulk, setPickingBulk] = useState(false);
  const [releasingBulk, setReleasingBulk] = useState(false);
  const [actionError, setActionError] = useState<string | null>(null);
  const [approvalErrorModal, setApprovalErrorModal] = useState<string | null>(null);
  const [manualErrorModal, setManualErrorModal] = useState<string | null>(null);
  const [showApproveConfirm, setShowApproveConfirm] = useState(false);
  const [showManualConfirm, setShowManualConfirm] = useState(false);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const addToast = (message: string, type: ToastMessage['type'] = 'success') =>
    setToasts(prev => [...prev, { id: nextToastId(), message, type }]);
  const dismissToast = (id: number) => setToasts(prev => prev.filter(t => t.id !== id));
  const [columnOrder, setColumnOrder] = useState<SupervisorColumnKey[]>(DEFAULT_COLUMN_ORDER);
  const [colDragIdx, setColDragIdx] = useState<number | null>(null);
  const [sortKey, setSortKey] = useState<SupervisorColumnKey | null>(null);
  const [sortDir, setSortDir] = useState<SortDir>('asc');
  const [columnWidths, setColumnWidths] = useState<Partial<Record<SupervisorColumnKey, number>>>(DEFAULT_COLUMN_WIDTHS);
  const [columnFilters, setColumnFilters] = useState<Record<SupervisorColumnKey, Set<string>>>(() => createEmptyFilters());
  const [columnFilterOps, setColumnFilterOps] = useState<Record<SupervisorColumnKey, FilterOperator>>(DEFAULT_FILTER_OPS);
  const [refreshCounter, setRefreshCounter] = useState(0);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        if (currentUser) return;
        const resolved = await resolveCurrentSystemUser();
        if (!cancelled) {
          setCurrentUser({
            systemUserId: resolved.systemUserId,
            displayName: resolved.displayName,
          });
        }
      } catch {
        // silently fail
      }
    })();
    return () => { cancelled = true; };
  }, [currentUser]);

  useEffect(() => {
    let cancelled = false;

    // Use cached data when navigating back; only re-fetch when Refresh is clicked.
    if (saItemsCache !== null && refreshCounter === 0) return;

    (async () => {
      try {
        fetchedQueueIds.current.clear();
        setLoading(true);
        setError(null);
        const result = await Vsi_participantprogramyearsService.getAll({
          select: [
            'vsi_name',
            '_vsi_participantid_value',
            'vsi_enrolmentstatus',
            'vsi_calculatedenfee',
            'vsi_previousyearcalculatedenfee',
            'vsi_variancecalculation',
            'vsi_taskstatus',
            'modifiedon',
          ],
          filter: "vsi_taskstatus eq 865520001",
          orderBy: ['modifiedon desc'],
          maxPageSize: 5000,
        });

        const queuesResult = await QueuesService.getAll({
          select: ['queueid', 'name'],
          filter: `name eq '${SUPERVISOR_QUEUE_NAME}' and statecode eq 0`,
          maxPageSize: 1,
        });

        const queueFallbackResult = (!queuesResult.success || (queuesResult.data?.length ?? 0) === 0)
          ? await QueuesService.getAll({
              select: ['queueid', 'name'],
              filter: "contains(name,'Supervisor') and contains(name,'Approval') and statecode eq 0",
              maxPageSize: 20,
            })
          : null;

        const supervisorQueues = (queuesResult.success && (queuesResult.data?.length ?? 0) > 0)
          ? (queuesResult.data ?? [])
          : (queueFallbackResult?.success ? (queueFallbackResult.data ?? []) : []);

        const supervisorQueueIdSet = new Set(
          supervisorQueues
            .map(q => normalizeGuid(q.queueid))
            .filter((id): id is string => !!id)
        );

        const queueitems: Array<NonNullable<Awaited<ReturnType<typeof QueueitemsService.getAll>>['data']>[number]> = [];
        let queueitemsSuccess = true;
        let queueitemsErrorMessage: string | undefined;
        let queueSkipToken: string | undefined;

        do {
          const queueitemsResult = await QueueitemsService.getAll({
            select: [
              'queueitemid',
              '_objectid_value',
              '_workerid_value',
              '_queueid_value',
              'enteredon',
              'workeridmodifiedon',
              'statecode',
            ],
            filter: 'statecode eq 0',
            maxPageSize: 5000,
            ...(queueSkipToken ? { skipToken: queueSkipToken } : {}),
          });

          if (!queueitemsResult.success) {
            queueitemsSuccess = false;
            queueitemsErrorMessage = queueitemsResult.error?.message ?? 'Unable to load queue work metadata.';
            break;
          }

          queueitems.push(...(queueitemsResult.data ?? []));
          queueSkipToken = queueitemsResult.skipToken;
        } while (queueSkipToken);

        const queueMap: Record<string, QueueWorkMeta> = {};
        if (queueitemsSuccess) {
          for (const q of queueitems) {
            const enrolmentId = normalizeGuid(q._objectid_value);
            if (!enrolmentId) continue;
            const isActive = q.statecode === 0;
            if (!isActive) continue;

            const queueDisplayName = (q as unknown as Record<string, unknown>)['_queueid_value@OData.Community.Display.V1.FormattedValue'] as string | undefined;
            const queueId = normalizeGuid(q._queueid_value);
            const isSupervisorQueue =
              (queueId ? supervisorQueueIdSet.has(queueId) : false)
              || isSupervisorApprovalQueueName(queueDisplayName);
            if (!isSupervisorQueue) continue;

            const existing = queueMap[enrolmentId];
            if (existing) continue;

            const workerDisplayName = (q as unknown as Record<string, unknown>)['_workerid_value@OData.Community.Display.V1.FormattedValue'] as string | undefined;
            queueMap[enrolmentId] = {
              workedBy: workerDisplayName ?? '—',
              workedOn: formatWorkedOn(q.workeridmodifiedon),
              workedOnRaw: q.workeridmodifiedon,
              enteredQueue: formatWorkedOn(q.enteredon),
              enteredQueueRaw: q.enteredon,
              workerId: q._workerid_value,
              queueitemId: q.queueitemid,
              queueId,
              queueName: queueDisplayName,
              isActive,
            };
          }
        }

        if (Object.keys(queueMap).length > 0) {
          Object.values(queueMap).forEach(meta => {
            if (meta.queueId) supervisorQueueIdSet.add(meta.queueId);
          });
        }

        if (!cancelled) {
          if (!result.success) {
            setItems([]);
            setError(result.error?.message ?? 'Unable to load supervisor approval items.');
          } else {
            const allowedEnrolmentIds = new Set(Object.keys(queueMap));
            const filteredItems = (result.data ?? []).filter(item => {
              const id = normalizeGuid(item.vsi_participantprogramyearid);
              return !!id && allowedEnrolmentIds.has(id);
            });
            setItems(filteredItems);
            saItemsCache = filteredItems;
          }

          setQueueWorkByEnrolmentId(queueMap);
          setSupervisorQueueIds(new Set(supervisorQueueIdSet));
          saQueueWorkCache = queueMap;
          saSupervisorQueueIdsCache = new Set(supervisorQueueIdSet);
          if (!queueitemsSuccess) {
            setError(queueitemsErrorMessage ?? 'Unable to load queue work metadata.');
          }
        }
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [refreshCounter]);

  useEffect(() => {
    const workerIds = [...new Set(
      Object.values(queueWorkByEnrolmentId)
        .map(m => m.workerId)
        .filter((id): id is string => !!id)
    )];
    if (workerIds.length === 0) return;
    let cancelled = false;
    (async () => {
      const entries: [string, string][] = [];
      for (const id of workerIds) {
        if (cancelled) break;
        const result = await Office365UsersService.UserPhoto_V2(id);
        if (!cancelled && result.success && result.data) {
          entries.push([id, result.data]);
        }
      }
      if (!cancelled && entries.length > 0) {
        setWorkerAvatarUrls(prev => {
          const next = { ...prev, ...Object.fromEntries(entries) };
          saWorkerAvatarUrlsCache = next;
          return next;
        });
      }
    })();
    return () => { cancelled = true; };
  }, [queueWorkByEnrolmentId]);

  // Per-row fallback: for items not resolved by the bulk fetch, query individually
  useEffect(() => {
    if (items.length === 0) return;
    if (supervisorQueueIds.size === 0) return;
    let cancelled = false;

    const missingIds = items
      .map(item => normalizeGuid(item.vsi_participantprogramyearid))
      .filter((id): id is string => !!id && !fetchedQueueIds.current.has(id));

    if (missingIds.length === 0) return;
    missingIds.forEach(id => fetchedQueueIds.current.add(id));

    (async () => {
      const updates: Record<string, QueueWorkMeta> = {};
      for (const enrolmentId of missingIds) {
        if (cancelled) break;
        const queueConstraint = [...supervisorQueueIds]
          .map(id => `_queueid_value eq '${id}'`)
          .join(' or ');
        const result = await QueueitemsService.getAll({
          select: ['queueitemid', '_objectid_value', '_workerid_value', '_queueid_value', 'enteredon', 'workeridmodifiedon'],
          filter: `objectid_vsi_participantprogramyear/vsi_participantprogramyearid eq '${enrolmentId}' and statecode eq 0 and (${queueConstraint})`,
          maxPageSize: 1,
        });
        if (!cancelled && result.success && (result.data?.length ?? 0) > 0) {
          const q = result.data![0];
          const workerDisplayName = (q as unknown as Record<string, unknown>)['_workerid_value@OData.Community.Display.V1.FormattedValue'] as string | undefined;
          const queueDisplayName = (q as unknown as Record<string, unknown>)['_queueid_value@OData.Community.Display.V1.FormattedValue'] as string | undefined;
          updates[enrolmentId] = {
            workedBy: workerDisplayName ?? '—',
            workedOn: formatWorkedOn(q.workeridmodifiedon),
            workedOnRaw: q.workeridmodifiedon,
            enteredQueue: formatWorkedOn(q.enteredon),
            enteredQueueRaw: q.enteredon,
            workerId: q._workerid_value,
            queueitemId: q.queueitemid,
            queueId: q._queueid_value,
            queueName: queueDisplayName,
            isActive: true,
          };
        }
      }
      if (!cancelled && Object.keys(updates).length > 0) {
        setQueueWorkByEnrolmentId(prev => ({ ...prev, ...updates }));
      }
    })();

    return () => { cancelled = true; };
  }, [items, supervisorQueueIds]);

  const resolveCurrentUser = async (): Promise<{ systemUserId: string; displayName: string }> => {
    if (currentUser) return currentUser;
    const resolved = await resolveCurrentSystemUser();
    const nextUser = {
      systemUserId: resolved.systemUserId,
      displayName: resolved.displayName,
    };
    setCurrentUser(nextUser);
    return nextUser;
  };

  const canApproveRow = (row: SupervisorRowView): boolean => {
    // Must have a calculated fee (not null)
    if (row.calculatedFeeValue == null) return false;
    const workerId = normalizeGuid(row.workMeta?.workerId);
    if (!workerId) return true;
    if (!currentUser?.systemUserId) return false;
    return workerId === normalizeGuid(currentUser.systemUserId);
  };

  const getApprovalOwnershipError = (rows: SupervisorRowView[]): string | null => {
    const blockedRows = rows.filter(row => !canApproveRow(row));
    if (blockedRows.length === 0) return null;

    // Check if any are missing calculated fee
    const missingFee = blockedRows.find(row => row.calculatedFeeValue == null);
    if (missingFee) {
      return `${missingFee.enrolmentName} cannot be approved because it does not have a calculated fee.`;
    }

    if (blockedRows.length === 1) {
      return `${blockedRows[0].enrolmentName} is assigned to ${blockedRows[0].workedBy}. You can only approve items worked by you or with a blank Worked By value, and only if a calculated fee exists.`;
    }

    return 'Some selected enrolments are assigned to another worker or do not have a calculated fee. You can only approve items worked by you or with a blank Worked By value, and only if a calculated fee exists.';
  };

  const bulkApprovalBlockedTooltip = 'One or more selected approvals is being worked on by another user';

  const updateQueueItemWorker = async (queueItemId: string, systemUserId: string | null): Promise<void> => {
    const bindingValue = systemUserId ? `/systemusers(${systemUserId})` : null;

    let updateResult = await QueueitemsService.update(queueItemId, {
      'workerid_systemuser@odata.bind': bindingValue,
    } as unknown as Parameters<typeof QueueitemsService.update>[1]);

    if (!updateResult.success) {
      updateResult = await QueueitemsService.update(queueItemId, {
        'WorkerId@odata.bind': bindingValue,
      } as Parameters<typeof QueueitemsService.update>[1]);
    }

    if (!updateResult.success) {
      throw new Error(updateResult.error?.message ?? 'Failed to update queue item worker.');
    }
  };

  const removeApprovedRowsFromState = (rows: SupervisorRowView[]) => {
    const normalizedApproved = new Set(
      rows
        .map(r => normalizeGuid(r.itemId))
        .filter((id): id is string => !!id)
    );

    setItems(prev => prev.filter(item => !normalizedApproved.has(normalizeGuid(item.vsi_participantprogramyearid))));

    setQueueWorkByEnrolmentId(prev => {
      const next = { ...prev };
      for (const row of rows) {
        if (row.itemId) {
          delete next[row.itemId];
          delete next[normalizeGuid(row.itemId)];
        }
      }
      return next;
    });

    setSelectedIds(prev => {
      const next = new Set(prev);
      for (const row of rows) {
        if (row.itemId) next.delete(row.itemId);
      }
      return next;
    });
  };

  const approveRows = async (rows: SupervisorRowView[]) => {
   
    for (const row of rows) {
      if (!row.itemId) continue;

      const enrolmentId = row.itemId;
      const statusUpdateResult = await Vsi_participantprogramyearsService.update(enrolmentId, {
        vsi_taskstatus: 865520003,
      });
      if (!statusUpdateResult.success) {
        throw new Error(statusUpdateResult.error?.message ?? `Failed to set Approved status for ${enrolmentId}.`);
      }

      if (row.workMeta?.queueitemId) {
        await QueueitemsService.delete(row.workMeta.queueitemId);
      }
    }

    patchEnrolmentCache(rows
      .filter(r => r.itemId != null)
      .map(r => ({ id: r.itemId!, fields: { vsi_taskstatus: 865520003 as unknown as import('../generated/models/Vsi_participantprogramyearsModel').Vsi_participantprogramyearsvsi_taskstatus } }))
    );
    removeApprovedRowsFromState(rows);
  };

  const handleApproveSelected = async () => {
    if (selectedIds.size === 0) return;
    setApprovingBulk(true);
    setActionError(null);
    try {
      const rowsToApprove = allRows.filter(row => row.itemId != null && selectedIds.has(row.itemId));
      if (rowsToApprove.length === 0) return;
      const ownershipError = getApprovalOwnershipError(rowsToApprove);
      if (ownershipError) throw new Error(ownershipError);
      await approveRows(rowsToApprove);
      setSelectedIds(new Set());
      addToast(`${rowsToApprove.length} enrolment${rowsToApprove.length === 1 ? '' : 's'} approved successfully.`);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Approve selected failed';
      setActionError(msg);
      addToast(msg, 'error');
    } finally {
      setApprovingBulk(false);
    }
  };

  const handlePickSelected = async () => {
    if (selectedIds.size === 0) return;
    setPickingBulk(true);
    setActionError(null);
    try {
      const user = await resolveCurrentUser();
      const rowsToPick = allRows.filter(row => row.itemId != null && selectedIds.has(row.itemId) && row.workMeta?.queueitemId);
      for (const row of rowsToPick) {
        await updateQueueItemWorker(row.workMeta!.queueitemId!, user.systemUserId);
        setQueueWorkByEnrolmentId(prev => ({
          ...prev,
          [row.itemId!]: {
            ...prev[row.itemId!],
            workedBy: user.displayName,
            workedOn: new Date().toLocaleDateString(),
            workedOnRaw: new Date().toISOString(),
            workerId: user.systemUserId,
          },
        }));
      }
      addToast(`${rowsToPick.length} enrolment${rowsToPick.length === 1 ? '' : 's'} picked.`);
      setSelectedIds(new Set());
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Pick failed';
      setActionError(msg);
      addToast(msg, 'error');
    } finally {
      setPickingBulk(false);
    }
  };

  const handleReleaseSelected = async () => {
    if (selectedIds.size === 0) return;
    setReleasingBulk(true);
    setActionError(null);
    try {
      const rowsToRelease = allRows.filter(row => row.itemId != null && selectedIds.has(row.itemId) && row.workMeta?.queueitemId);
      for (const row of rowsToRelease) {
        await updateQueueItemWorker(row.workMeta!.queueitemId!, null);
        setQueueWorkByEnrolmentId(prev => ({
          ...prev,
          [row.itemId!]: {
            ...prev[row.itemId!],
            workedBy: '—',
            workedOn: '—',
            workedOnRaw: undefined,
            workerId: undefined,
          },
        }));
      }
      addToast(`${rowsToRelease.length} enrolment${rowsToRelease.length === 1 ? '' : 's'} released.`);
      setSelectedIds(new Set());
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Release failed';
      setActionError(msg);
      addToast(msg, 'error');
    } finally {
      setReleasingBulk(false);
    }
  };

  const handleAssignSelected = () => {
    if (selectedRows.length === 0) return;
    const first = selectedRows[0];
    if (!first.itemId) return;
    setAssignTarget({
      enrolmentId: first.itemId,
      enrolmentName: selectedRows.length === 1 ? (first.item.vsi_name ?? '—') : `${selectedRows.length} enrolments`,
      queueitemId: first.workMeta?.queueitemId,
      queueId: first.workMeta?.queueId,
      queueName: first.workMeta?.queueName,
      bulkRows: selectedRows.length > 1
        ? selectedRows.filter(r => r.itemId != null).map(r => ({ itemId: r.itemId!, queueitemId: r.workMeta?.queueitemId }))
        : undefined,
    });
  };

  const onSort = (key: SupervisorColumnKey, dir: SortDir) => {
    setSortKey(key);
    setSortDir(dir);
  };

  const onColumnWidthChange = (key: SupervisorColumnKey) => (width: number | undefined) => {
    setColumnWidths(prev => ({ ...prev, [key]: width }));
  };

  const onColDragStart = (index: number) => {
    setColDragIdx(index);
  };

  const onColDragOver = (event: DragEvent, index: number) => {
    event.preventDefault();
    if (colDragIdx === null || colDragIdx === index) return;

    setColumnOrder(prev => {
      const next = [...prev];
      const [moved] = next.splice(colDragIdx, 1);
      next.splice(index, 0, moved);
      return next;
    });
    setColDragIdx(index);
  };

  const onColDragEnd = () => {
    setColDragIdx(null);
  };

  type SupervisorRowView = {
    item: Vsi_participantprogramyears;
    itemId: string | null;
    workMeta: QueueWorkMeta | undefined;
    enrolmentName: string;
    participantName: string;
    taskStatusLabel: string;
    enrolmentStatusLabel: string;
    calculatedFeeValue: number | null;
    enteredQueue: string;
    enteredQueueRaw?: string;
    workedBy: string;
    workedOn: string;
    workedOnRaw?: string;
  };

  const allRows = useMemo<SupervisorRowView[]>(() => {
    return items.map(item => {
      const itemId = getRowId(item);
      const workMeta = itemId ? queueWorkByEnrolmentId[itemId] : undefined;
      const rawItem = item as unknown as Record<string, unknown>;
      const participantDisplayName = (rawItem['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue'] as string | undefined)
        ?? item.vsi_participantidname
        ?? '—';
      return {
        item,
        itemId,
        workMeta,
        enrolmentName: item.vsi_name ?? '—',
        participantName: participantDisplayName,
        taskStatusLabel: getTaskStatusLabel(item.vsi_taskstatus) || 'Unknown',
        enrolmentStatusLabel: getEnrolmentStatusLabel(item.vsi_enrolmentstatus) || '—',
        calculatedFeeValue: item.vsi_calculatedenfee ?? null,
        enteredQueue: workMeta?.enteredQueue ?? '—',
        enteredQueueRaw: workMeta?.enteredQueueRaw,
        workedBy: workMeta?.workedBy ?? '—',
        workedOn: workMeta?.workedOn ?? '—',
        workedOnRaw: workMeta?.workedOnRaw,
      };
    });
  }, [items, queueWorkByEnrolmentId]);

  const filterValue = useCallback((row: SupervisorRowView, key: SupervisorColumnKey): string => {
    switch (key) {
      case 'enrolmentName':
        return row.enrolmentName;
      case 'participant':
        return row.participantName;
      case 'taskStatus':
        return row.taskStatusLabel;
      case 'enrolmentStatus':
        return row.enrolmentStatusLabel;
      case 'calculatedFee':
        return formatCurrencyOr(row.calculatedFeeValue, '—');
      case 'enteredQueue':
        return row.enteredQueue;
      case 'workedBy':
        return row.workedBy;
      case 'workedOn':
        return row.workedOn;
      default:
        return '—';
    }
  }, []);

  const filterOptionsByColumn = useMemo<Record<SupervisorColumnKey, string[]>>(() => {
    const buckets: Record<SupervisorColumnKey, Set<string>> = {
      enrolmentName: new Set(),
      participant: new Set(),
      taskStatus: new Set(),
      enrolmentStatus: new Set(),
      calculatedFee: new Set(),
      enteredQueue: new Set(),
      workedBy: new Set(),
      workedOn: new Set(),
    };

    for (const row of allRows) {
      (Object.keys(buckets) as SupervisorColumnKey[]).forEach(key => {
        buckets[key].add(filterValue(row, key));
      });
    }

    return {
      enrolmentName: [...buckets.enrolmentName].sort((a, b) => a.localeCompare(b)),
      participant: [...buckets.participant].sort((a, b) => a.localeCompare(b)),
      taskStatus: [...buckets.taskStatus].sort((a, b) => a.localeCompare(b)),
      enrolmentStatus: ENROLMENT_STATUS_FILTER_OPTIONS,
      calculatedFee: [...buckets.calculatedFee].sort((a, b) => a.localeCompare(b)),
      enteredQueue: [...buckets.enteredQueue].sort((a, b) => a.localeCompare(b)),
      workedBy: [...buckets.workedBy].sort((a, b) => a.localeCompare(b)),
      workedOn: [...buckets.workedOn].sort((a, b) => a.localeCompare(b)),
    };
  }, [allRows, filterValue]);

  const filteredAndSortedRows = useMemo(() => {
    const filtered = allRows.filter(row => {
      for (const key of Object.keys(columnFilters) as SupervisorColumnKey[]) {
        const selected = columnFilters[key];
        if (selected.size === 0) continue;
        const value = filterValue(row, key);
        const hasValue = selected.has(value);
        const pass = columnFilterOps[key] === 'equals' ? hasValue : !hasValue;
        if (!pass) return false;
      }
      return true;
    });

    if (!sortKey) return filtered;

    const sorted = [...filtered].sort((a, b) => {
      let cmp = 0;
      switch (sortKey) {
        case 'calculatedFee': {
          const av = a.calculatedFeeValue ?? Number.NEGATIVE_INFINITY;
          const bv = b.calculatedFeeValue ?? Number.NEGATIVE_INFINITY;
          cmp = av - bv;
          break;
        }
        case 'workedOn': {
          const at = a.workedOnRaw ? new Date(a.workedOnRaw).getTime() : Number.NEGATIVE_INFINITY;
          const bt = b.workedOnRaw ? new Date(b.workedOnRaw).getTime() : Number.NEGATIVE_INFINITY;
          cmp = at - bt;
          break;
        }
        case 'enteredQueue': {
          const at = a.enteredQueueRaw ? new Date(a.enteredQueueRaw).getTime() : Number.NEGATIVE_INFINITY;
          const bt = b.enteredQueueRaw ? new Date(b.enteredQueueRaw).getTime() : Number.NEGATIVE_INFINITY;
          cmp = at - bt;
          break;
        }
        case 'enrolmentName':
          cmp = a.enrolmentName.localeCompare(b.enrolmentName);
          break;
        case 'participant':
          cmp = a.participantName.localeCompare(b.participantName);
          break;
        case 'taskStatus':
          cmp = a.taskStatusLabel.localeCompare(b.taskStatusLabel);
          break;
        case 'enrolmentStatus':
          cmp = a.enrolmentStatusLabel.localeCompare(b.enrolmentStatusLabel);
          break;
        case 'workedBy':
          cmp = a.workedBy.localeCompare(b.workedBy);
          break;
        default:
          cmp = 0;
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });

    return sorted;
  }, [allRows, columnFilterOps, columnFilters, filterValue, sortDir, sortKey]);

  const totalPages = Math.max(1, Math.ceil(filteredAndSortedRows.length / PAGE_SIZE));
  const pageRows = filteredAndSortedRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  useEffect(() => {
    setPage(prev => Math.min(prev, totalPages));
  }, [totalPages]);

  const allSelected =
    pageRows.length > 0 && pageRows.every(row => {
      return row.itemId != null && selectedIds.has(row.itemId);
    });

  const toggleSelectAll = () => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (allSelected) {
        pageRows.forEach(row => {
          const itemId = row.itemId;
          if (itemId != null) next.delete(itemId);
        });
      } else {
        pageRows.forEach(row => {
          const itemId = row.itemId;
          if (itemId != null) next.add(itemId);
        });
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

  const selectedCount = selectedIds.size;
  const selectedRows = allRows.filter(row => row.itemId != null && selectedIds.has(row.itemId));
  const hasBlockedSelectedRows = selectedRows.some(row => !canApproveRow(row));

  return (
    <div className="sa-wrapper">
      <div>
        <h1 className="sa-page-title">Supervisor&rsquo;s Approval Queue</h1>
        <p className="sa-page-subtitle">
          Review and manage pending approval items, update task statuses, and ensure timely processing.
        </p>
      </div>

      <div className="sa-filters-bar">
        <button type="button" className="sa-filter-btn">
          <Filter size={14} />
          Filters
        </button>
        <button
          type="button"
          className="sa-filter-btn"
          disabled={loading}
          onClick={() => {
            saItemsCache = null;
            saQueueWorkCache = null;
            saSupervisorQueueIdsCache = null;
            saWorkerAvatarUrlsCache = null;
            setRefreshCounter(prev => prev + 1);
          }}
        >
          {loading ? 'Refreshing...' : 'Refresh'}
        </button>
      </div>

      <div className="sa-card">
        <div className="sa-card-header">
          <div className="sa-card-title-block">
            <h2 className="sa-card-title">Pending Reviews</h2>
            <p className="sa-card-subtitle">Items requiring your attention, referred by admin(s)</p>
          </div>
          <div className="sa-bulk-actions">
            {selectedCount > 0 && (
              <span className="sa-selected-count">{selectedCount} selected</span>
            )}
            <button
              type="button"
              className="sa-btn-clear sa-bulk-btn"
              disabled={selectedCount === 0 || pickingBulk}
              onClick={() => void handlePickSelected()}
            >{pickingBulk ? 'Picking...' : <><ClipboardList size={15} /> Pick</>}</button>
            <button
              type="button"
              className="sa-btn-clear sa-bulk-btn"
              disabled={selectedCount === 0 || releasingBulk}
              onClick={() => void handleReleaseSelected()}
            >{releasingBulk ? 'Releasing...' : <><LogOut size={15} /> Release</>}</button>
            <button
              type="button"
              className="sa-btn-clear sa-bulk-btn"
              disabled={selectedCount === 0}
              onClick={handleAssignSelected}
            ><UserPlus size={15} /> Assign</button>
            <button
              type="button"
              className="sa-btn-primary sa-bulk-btn"
              disabled={selectedCount === 0}
              onClick={() => setShowManualConfirm(true)}
            >
              <Wrench size={15} style={{ marginRight: 4, marginBottom: -2 }} /> Manual
            </button>
            <span title={hasBlockedSelectedRows ? bulkApprovalBlockedTooltip : undefined}>
              <button
                type="button"
                className="sa-btn-primary sa-bulk-btn"
                disabled={selectedCount === 0 || approvingBulk}
                onClick={() => setShowApproveConfirm(true)}
              >
                {approvingBulk ? 'Approving...' : <><CircleCheck size={15} /> Approve</>}
              </button>
            </span>
          </div>
        </div>

        <div className="sa-table-container">
          {loading && <p className="sa-state-msg loading">Loading queue items…</p>}
          {error && <p className="sa-state-msg error">Error: {error}</p>}
          {actionError && (
            <p className="sa-state-msg error" style={{ cursor: 'pointer' }} onClick={() => setActionError(null)}>
              Action failed: {actionError} &times;
            </p>
          )}
          {approvalErrorModal && (
            <ApprovalErrorModal message={approvalErrorModal} onClose={() => setApprovalErrorModal(null)} />
          )}
          {manualErrorModal && (
            <ManualErrorModal message={manualErrorModal} onClose={() => setManualErrorModal(null)} />
          )}
          {showApproveConfirm && (
            <ConfirmActionModal
              title="Confirm Approve Enrolments"
              message={`Are you sure you want to approve the selected ${selectedCount} enrolment${selectedCount === 1 ? '' : 's'}?`}
              enrolments={selectedRows.map(row => ({ id: row.itemId!, name: row.enrolmentName }))}
              confirmLabel="Approve"
              cancelLabel="Cancel"
              loading={approvingBulk}
              onConfirm={async () => {
                setShowApproveConfirm(false);
                const rowsToApprove = allRows.filter(row => row.itemId != null && selectedIds.has(row.itemId));
                const blockedRows = rowsToApprove.filter(row => !canApproveRow(row));
                if (blockedRows.length > 0) {
                  let msg = '';
                  if (blockedRows.length === 1) {
                    msg = `${blockedRows[0].enrolmentName} cannot be approved. Only enrolments with status Ready and a calculated fee, assigned to you or with a blank Worked By value, can be approved.`;
                  } else {
                    msg = 'Only enrolments with status Ready and a calculated fee, assigned to you or with a blank Worked By value, can be approved. Please adjust your selection.';
                  }
                  setApprovalErrorModal(msg);
                  return;
                }
                await handleApproveSelected();
              }}
              onCancel={() => setShowApproveConfirm(false)}
            />
          )}
          {showManualConfirm && (
            <ConfirmActionModal
              title="Confirm Set to Manual"
              message={`Are you sure you want to set the selected ${selectedCount} enrolment${selectedCount === 1 ? '' : 's'} to Manual/To be reviewed?`}
              enrolments={selectedRows.map(row => ({ id: row.itemId!, name: row.enrolmentName }))}
              confirmLabel="Set to Manual"
              cancelLabel="Cancel"
              loading={false}
              onConfirm={async () => {
                setShowManualConfirm(false);
                if (!currentUser) {
                  addToast('User not resolved', 'error');
                  return;
                }
                const selectedRows = allRows.filter(row => row.itemId != null && selectedIds.has(row.itemId));
                const blockedRows = selectedRows.filter(row => {
                  const workedBy = row.workedBy?.trim();
                  return workedBy && workedBy !== '' && workedBy !== currentUser.displayName && workedBy !== '—';
                });
                if (blockedRows.length > 0) {
                  let msg = '';
                  if (blockedRows.length === 1) {
                    msg = `${blockedRows[0].enrolmentName} cannot be set to Manual. Only enrolments assigned to you or with a blank Worked By value can be set to Manual.`;
                  } else {
                    msg = 'Only enrolments assigned to you or with a blank Worked By value can be set to Manual. Please adjust your selection.';
                  }
                  setManualErrorModal(msg);
                  return;
                }
                const rowsToManual = selectedRows;
                try {
                  await Promise.all(rowsToManual.map(row =>
                    Promise.all([
                      Vsi_participantprogramyearsService.update(row.itemId!, {
                        vsi_taskstatus: 865520000, // Manual
                        vsi_enrolmentstatus: 865520009 // ToBeReviewed
                      }),
                      row.workMeta?.queueitemId ? QueueitemsService.delete(row.workMeta.queueitemId) : Promise.resolve()
                    ])
                  ));  
                  patchEnrolmentCache(rowsToManual
                    .filter(r => r.itemId != null)
                    .map(r => ({ id: r.itemId!, fields: {
                      vsi_taskstatus: 865520000 as unknown as import('../generated/models/Vsi_participantprogramyearsModel').Vsi_participantprogramyearsvsi_taskstatus,
                      vsi_enrolmentstatus: 865520009 as unknown as import('../generated/models/Vsi_participantprogramyearsModel').Vsi_participantprogramyearsvsi_enrolmentstatus,
                    } }))
                  );
                  removeApprovedRowsFromState(rowsToManual);
                  setSelectedIds(new Set());
                  addToast(`${rowsToManual.length} enrolment${rowsToManual.length === 1 ? '' : 's'} set to Manual/To be reviewed and removed from queue.`);
                } catch (err) {
                  const msg = err instanceof Error ? err.message : 'Manual action failed';
                  addToast(msg, 'error');
                }
              }}
              onCancel={() => setShowManualConfirm(false)}
            />
          )}
          {!loading && !error && items.length === 0 && (
            <p className="sa-state-msg empty">No pending items in the supervisor approval queue.</p>
          )}
          {!loading && !error && items.length > 0 && (
            <table className="sa-table">
              <thead>
                <tr>
                  <th className="sa-th-check">
                    <input
                      type="checkbox"
                      checked={allSelected}
                      onChange={toggleSelectAll}
                      aria-label="Select all"
                    />
                  </th>
                  {columnOrder.map((key, index) => {
                    const colDef = SUPERVISOR_COLUMNS.find(c => c.key === key)!;
                    return (
                      <ColumnHeaderMenu
                        key={key}
                        label={colDef.label}
                        sortKey={key}
                        currentSortKey={sortKey}
                        currentSortDir={sortDir}
                        onSort={(k, dir) => onSort(k as SupervisorColumnKey, dir)}
                        columnWidth={columnWidths[key]}
                        onColumnWidthChange={onColumnWidthChange(key)}
                        filterOptions={filterOptionsByColumn[key]}
                        selectedFilters={columnFilters[key]}
                        filterOperator={columnFilterOps[key]}
                        onFilterChange={next => setColumnFilters(prev => ({ ...prev, [key]: new Set(next) }))}
                        onFilterOperatorChange={op => setColumnFilterOps(prev => ({ ...prev, [key]: op }))}
                        dragProps={{
                          draggable: true,
                          onDragStart: () => onColDragStart(index),
                          onDragOver: (event: DragEvent) => onColDragOver(event, index),
                          onDragEnd: onColDragEnd,
                          className: colDragIdx === index ? 'col-dragging' : undefined,
                        }}
                      />
                    );
                  })}
                  <th className="sa-th-actions"></th>
                </tr>
              </thead>
              <tbody>
                {pageRows.map((row, index) => {
                  const item = row.item;
                  const itemId = row.itemId;
                  const hasCalculatedFee = item.vsi_calculatedenfee != null;
                  const variance = item.vsi_calculatedenfee != null && item.vsi_variancecalculation != null ? item.vsi_variancecalculation * 100 : null;
                  const workMeta = row.workMeta;

                  // ...existing code...
                  return (
                    <tr key={itemId ?? `${item.vsi_name ?? 'row'}-${index}`}> 
                      <td className="sa-td-check">
                        <input
                          type="checkbox"
                          checked={itemId != null && selectedIds.has(itemId)}
                          onChange={() => {
                            if (itemId != null) toggleSelect(itemId);
                          }}
                          aria-label={`Select ${item.vsi_name ?? itemId ?? 'row'}`}
                          disabled={itemId == null}
                        />
                      </td>
                      {columnOrder.map(key => {
                        if (key === 'enrolmentName') {
                          return (
                            <td key={key} className="sa-pin">
                              {itemId
                                ? <Link className="cell-pin-link" to={`/enrolment/supervisor/${itemId}`}>{item.vsi_name ?? '—'}</Link>
                                : item.vsi_name ?? '—'}
                            </td>
                          );
                        }
                        if (key === 'participant') {
                          const participantId = item._vsi_participantid_value;
                          const href = `${CORE_BASE_URL_FALLBACK}?appid=${encodeURIComponent(CORE_APP_ID_FALLBACK)}&pagetype=entityrecord&etn=account&id=${encodeURIComponent(participantId ?? '')}`;
                          return (
                            <td key={key}>
                              {participantId
                                ? <a className="cell-pin-link" href={href} target="_blank" rel="noopener noreferrer">{row.participantName}</a>
                                : row.participantName}
                            </td>
                          );
                        }
                        if (key === 'taskStatus') {
                          return (
                            <td key={key}>
                              <StatusBadge status={item.vsi_taskstatus} />
                            </td>
                          );
                        }
                        if (key === 'enrolmentStatus') {
                          const statusLabel = getEnrolmentStatusLabel(item.vsi_enrolmentstatus);
                          return (
                            <td key={key}>
                              <span className={`enrol-badge ${enrolmentStatusClass(statusLabel)}`}>{statusLabel || '—'}</span>
                            </td>
                          );
                        }
                        if (key === 'calculatedFee') {
                          return (
                            <td key={key}>
                              <span className="sa-fee-cell">
                                {itemId && hasCalculatedFee
                                  ? <Link className="sa-fee-amount sa-fee-link" to={`/calculation/supervisor/${itemId}`}>{formatCurrencyOr(item.vsi_calculatedenfee, '—')}</Link>
                                  : <span className="sa-fee-amount">{formatCurrencyOr(item.vsi_calculatedenfee, '—')}</span>}
                                {variance != null ? <VariancePill variance={variance} /> : null}
                              </span>
                            </td>
                          );
                        }
                        if (key === 'enteredQueue') {
                          return <td key={key}>{workMeta?.enteredQueue ?? '—'}</td>;
                        }
                        if (key === 'workedBy') {
                          return (
                            <td key={key} className="sa-worked-by-cell">
                              {!workMeta || workMeta.workedBy === '—'
                                ? '—'
                                : (
                                  <span className="sa-worked-by-content">
                                    {workMeta.workerId && workerAvatarUrls[workMeta.workerId]
                                      ? (
                                        <span className="avatar-circle" title={workMeta.workedBy} style={{ background: '#fff', padding: 0 }}>
                                          <img src={`data:image/jpeg;base64,${workerAvatarUrls[workMeta.workerId]}`} alt={workMeta.workedBy} style={{ width: '100%', height: '100%', borderRadius: '50%' }} />
                                        </span>
                                      )
                                      : (
                                        <span className="avatar-circle" title={workMeta.workedBy} style={{ background: getAvatarColor(workMeta.workedBy), color: '#fff', fontWeight: 600 }}>
                                          {getInitials(workMeta.workedBy)}
                                        </span>
                                      )}
                                    <span className="sa-worked-by-name">{workMeta.workedBy}</span>
                                  </span>
                                )}
                            </td>
                          );
                        }
                        return <td key={key}>{workMeta?.workedOn ?? '—'}</td>;
                      })}
                      <td className="sa-td-actions">
                        <div className="sa-row-actions">
                          {itemId
                            ? (
                              <Link
                                to={`/calculation/supervisor/${itemId}`}
                                aria-label="Go to calculation"
                                data-tooltip="Go to calculation"
                                className="sa-calc-link"
                              >
                                <Calculator size={16} />
                              </Link>
                            )
                            : (
                              <span
                                aria-label="Go to calculation"
                                data-tooltip="Go to calculation"
                                className="sa-calc-link sa-calc-link-disabled"
                              >
                                <Calculator size={16} className="sa-action-icon-disabled" />
                              </span>
                            )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
          {!loading && !error && items.length > 0 && filteredAndSortedRows.length === 0 && (
            <p className="sa-state-msg empty">No records match the current filters.</p>
          )}
        </div>

        {!loading && !error && items.length > 0 && (
          <div className="sa-pagination">
            <span>
              {filteredAndSortedRows.length === 0
                ? 'Showing 0 of 0 results'
                : `Showing ${Math.min((page - 1) * PAGE_SIZE + 1, filteredAndSortedRows.length)}–${Math.min(page * PAGE_SIZE, filteredAndSortedRows.length)} of ${filteredAndSortedRows.length} result${filteredAndSortedRows.length !== 1 ? 's' : ''}`}
            </span>
            <div className="sa-pagination-controls">
              <button
                type="button"
                className="sa-page-btn"
                onClick={() => setPage(p => Math.max(1, p - 1))}
                disabled={page === 1}
              >
                &lsaquo; Previous
              </button>
              {Array.from({ length: totalPages }, (_, i) => i + 1).map(p => (
                <button
                  key={p}
                  type="button"
                  className={`sa-page-btn${p === page ? ' active' : ''}`}
                  onClick={() => setPage(p)}
                >
                  {p}
                </button>
              ))}
              <button
                type="button"
                className="sa-page-btn"
                onClick={() => setPage(p => Math.min(totalPages, p + 1))}
                disabled={page === totalPages}
              >
                Next &rsaquo;
              </button>
            </div>
          </div>
        )}
      </div>

      {assignTarget && (
        <AssignWorkerModal
          enrolmentName={assignTarget.enrolmentName}
          queueitemId={assignTarget.queueitemId}
          queueId={assignTarget.queueId}
          queueName={assignTarget.queueName}
          onClose={() => setAssignTarget(null)}
          onAssigned={(workerId, workerName) => {
            const rowsToUpdate = assignTarget.bulkRows ?? [{ itemId: assignTarget.enrolmentId, queueitemId: assignTarget.queueitemId }];
            setQueueWorkByEnrolmentId(prev => {
              const next = { ...prev };
              for (const r of rowsToUpdate) {
                next[r.itemId] = {
                  ...next[r.itemId],
                  workedBy: workerName,
                  workedOn: new Date().toLocaleDateString(),
                  workerId,
                };
              }
              return next;
            });
            // Apply assignment to remaining queue items (first was handled by the modal)
            // and update ownerid on all enrolment records
            void (async () => {
              try {
                const remainingRows = assignTarget.bulkRows && assignTarget.bulkRows.length > 1
                  ? assignTarget.bulkRows.slice(1)
                  : [];
                for (const r of remainingRows) {
                  if (r.queueitemId) await QueueitemsService.delete(r.queueitemId);
                }
                for (const r of rowsToUpdate) {
                  await Vsi_participantprogramyearsService.update(r.itemId, {
                    'ownerid@odata.bind': `/systemusers(${workerId})`,
                    vsi_enrolmentstatus: 865520010,
                    vsi_taskstatus: 865520000,
                  } as unknown as Parameters<typeof Vsi_participantprogramyearsService.update>[1]);
                }
                patchEnrolmentCache(rowsToUpdate.map(r => ({ id: r.itemId, fields: {
                  vsi_taskstatus: 865520000 as unknown as import('../generated/models/Vsi_participantprogramyearsModel').Vsi_participantprogramyearsvsi_taskstatus,
                  vsi_enrolmentstatus: 865520010 as unknown as import('../generated/models/Vsi_participantprogramyearsModel').Vsi_participantprogramyearsvsi_enrolmentstatus,
                  '_ownerid_value': workerId,
                } })));
              } catch (err) {
                setActionError(err instanceof Error ? err.message : 'Assign partially failed');
              }
            })();
            setAssignTarget(null);
            // Remove assigned rows from the table, queue work state, and selection — same as Manual/Approve
            const assignedItemIds = new Set(rowsToUpdate.map(r => r.itemId));
            removeApprovedRowsFromState(allRows.filter(r => r.itemId != null && assignedItemIds.has(r.itemId)));
            const count = assignTarget.bulkRows?.length ?? 1;
            addToast(`${count} enrolment${count === 1 ? '' : 's'} assigned to ${workerName}.`);
          }}
        />
      )}
      <Toast toasts={toasts} onDismiss={dismissToast} />
    </div>
  );
}

