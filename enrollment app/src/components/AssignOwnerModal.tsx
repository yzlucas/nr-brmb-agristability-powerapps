import { useState, useEffect } from 'react';
import { SystemusersService } from '../generated/services/SystemusersService';
import { QueuemembershipsService } from '../generated/services/QueuemembershipsService';
import { QueuesService } from '../generated/services/QueuesService';
import { TeamsService } from '../generated/services/TeamsService';
import { TeammembershipsService } from '../generated/services/TeammembershipsService';
import { RolesService } from '../generated/services/RolesService';
import { SystemuserrolescollectionService } from '../generated/services/SystemuserrolescollectionService';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import { getInitials } from '../utils/helpers';

function buildUserIdOrFilter(ids: string[]): string {
  if (ids.length === 0) return 'false';
  return `(${ids.map(id => `systemuserid eq '${id}'`).join(' or ')})`;
}

interface AssignableUser {
  systemUserId: string;
  displayName: string;
  jobTitle?: string;
  mail?: string;
  group?: string;
}

const ASSIGNABLE_TASK_STATUSES = new Set([865520000, 865520002]); // Manual, Ready

export function AssignOwnerModal({
  selectedIds,
  rows,
  onClose,
  onComplete,
}: {
  selectedIds: Set<string>;
  rows: Vsi_participantprogramyears[];
  onClose: () => void;
  onComplete: (assignedIds: string[], ownerName: string) => void;
}) {
  const [searchTerm, setSearchTerm] = useState('');
  const [searching, setSearching] = useState(false);
  const [results, setResults] = useState<AssignableUser[]>([]);
  const [selected, setSelected] = useState<AssignableUser | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [hasSearched, setHasSearched] = useState(false);
  const [allowedMemberIds, setAllowedMemberIds] = useState<Set<string> | null>(null);
  const [memberIdsReady, setMemberIdsReady] = useState(false);
  const [groupIds, setGroupIds] = useState<{ sysAdmin: Set<string>; admin: Set<string>; queue: Set<string>; verifier: Set<string> }>(
    { sysAdmin: new Set(), admin: new Set(), queue: new Set(), verifier: new Set() }
  );

  const selectedRows = rows.filter(r => selectedIds.has(r.vsi_participantprogramyearid));
  const blockedRows = selectedRows.filter(r => !ASSIGNABLE_TASK_STATUSES.has(r.vsi_taskstatus as unknown as number));

  const applyGroupBadges = (users: AssignableUser[], g: typeof groupIds): AssignableUser[] =>
    users.map(u => ({
      ...u,
      group: g.sysAdmin.has(u.systemUserId) ? 'System Administrator'
        : g.admin.has(u.systemUserId) ? 'Enrolment Admin'
        : g.queue.has(u.systemUserId) ? 'Supervisor Approval Queue Member'
        : g.verifier.has(u.systemUserId) ? 'Verifier Team Member'
        : undefined,
    }));

  const searchUsers = async (term: string, memberIds: Set<string> | null, currentGroupIds: typeof groupIds) => {
    setSearching(true);
    setError(null);
    setHasSearched(true);
    try {
      const escaped = term.replace(/'/g, "''");
      const memberConstraint = memberIds && memberIds.size > 0
        ? ` and ${buildUserIdOrFilter([...memberIds])}`
        : '';
      const filter = term
        ? `contains(fullname,'${escaped}') and isdisabled eq false and not startswith(fullname,'#')${memberConstraint}`
        : `isdisabled eq false and not startswith(fullname,'#')${memberConstraint}`;

      const result = await SystemusersService.getAll({
        select: ['systemuserid', 'fullname', 'jobtitle', 'internalemailaddress'],
        filter,
        orderBy: ['fullname asc'],
        maxPageSize: 20,
      });

      if (!result.success) throw new Error(result.error?.message ?? 'Failed to search users');

      const mapped: AssignableUser[] = (result.data ?? [])
        .filter(u => u.systemuserid && !u.isdisabled)
        .map(u => ({
          systemUserId: u.systemuserid,
          displayName: u.fullname ?? u.internalemailaddress ?? u.systemuserid,
          jobTitle: u.jobtitle ?? undefined,
          mail: u.internalemailaddress ?? undefined,
        }));
      setResults(applyGroupBadges(mapped, currentGroupIds));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Search failed');
    } finally {
      setSearching(false);
    }
  };

  // On mount: resolve the same user pool as AssignWorkerModal on the SA page
  // (supervisor approval queue members + Enrolment Admin team + Verifier Team + System Administrators)
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const getTeamMemberIds = async (teamName: string): Promise<Set<string>> => {
          const teamsResp = await TeamsService.getAll({ filter: `name eq '${teamName}'`, maxPageSize: 1 });
          if (!teamsResp.success || !teamsResp.data?.length) return new Set();
          const teamId = teamsResp.data[0].teamid;
          const membersResp = await TeammembershipsService.getAll({ filter: `teamid eq '${teamId}'`, maxPageSize: 500 });
          if (!membersResp.success || !membersResp.data) return new Set();
          return new Set(membersResp.data.map(m => m.systemuserid).filter((id): id is string => !!id));
        };

        const getRoleUserIds = async (roleName: string): Promise<Set<string>> => {
          const rolesResp = await RolesService.getAll({ filter: `name eq '${roleName}'`, maxPageSize: 500 });
          if (!rolesResp.success || !rolesResp.data?.length) return new Set();
          const userIds = new Set<string>();
          await Promise.all(rolesResp.data.map(async role => {
            const userRolesResp = await SystemuserrolescollectionService.getAll({ filter: `roleid eq '${role.roleid}'`, maxPageSize: 500 });
            if (userRolesResp.success && userRolesResp.data) {
              userRolesResp.data.forEach(r => { if (r.systemuserid) userIds.add(r.systemuserid); });
            }
          }));
          return userIds;
        };

        // Find supervisor approval queues by name and get their members
        const queuesResp = await QueuesService.getAll({
          select: ['queueid', 'name'],
          filter: `contains(name,'Supervisor') and contains(name,'Approval')`,
          maxPageSize: 10,
        });
        const queueMemberIds = new Set<string>();
        if (queuesResp.success && queuesResp.data) {
          for (const q of queuesResp.data) {
            if (!q.queueid) continue;
            const memberships = await QueuemembershipsService.getAll({
              select: ['systemuserid'],
              filter: `queueid eq '${q.queueid}'`,
              maxPageSize: 500,
            });
            if (memberships.success && memberships.data) {
              memberships.data.forEach(m => { if (m.systemuserid) queueMemberIds.add(m.systemuserid); });
            }
          }
        }

        const [adminIds, verifierIds, sysAdminIds] = await Promise.all([
          getTeamMemberIds('Enrolment Admin'),
          getTeamMemberIds('Verifier Team Member'),
          getRoleUserIds('System Administrator'),
        ]);

        if (cancelled) return;
        const allIds = new Set([...queueMemberIds, ...adminIds, ...verifierIds, ...sysAdminIds]);
        setAllowedMemberIds(allIds.size > 0 ? allIds : null);
        setGroupIds({ sysAdmin: sysAdminIds, admin: adminIds, queue: queueMemberIds, verifier: verifierIds });
      } catch {
        // Non-fatal: fall back to all active users
      } finally {
        if (!cancelled) setMemberIdsReady(true);
      }
    })();
    return () => { cancelled = true; };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Debounced search — waits until allowed member IDs are resolved
  useEffect(() => {
    if (!memberIdsReady) return;
    const timer = setTimeout(() => { void searchUsers(searchTerm.trim(), allowedMemberIds, groupIds); }, searchTerm ? 300 : 0);
    return () => clearTimeout(timer);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchTerm, allowedMemberIds, memberIdsReady, groupIds]);

  const handleAssign = async () => {
    if (!selected) return;
    setSubmitting(true);
    setError(null);
    try {
      const eligibleRows = selectedRows.filter(r => ASSIGNABLE_TASK_STATUSES.has(r.vsi_taskstatus as unknown as number));
      for (const row of eligibleRows) {
        const result = await Vsi_participantprogramyearsService.update(row.vsi_participantprogramyearid, {
          'ownerid@odata.bind': `/systemusers(${selected.systemUserId})`,
        } as unknown as Parameters<typeof Vsi_participantprogramyearsService.update>[1]);
        if (!result.success) throw new Error(result.error?.message ?? `Failed to assign ${row.vsi_name ?? row.vsi_participantprogramyearid}`);
      }
      onComplete(eligibleRows.map(r => r.vsi_participantprogramyearid), selected.displayName);
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Assignment failed');
    } finally {
      setSubmitting(false);
    }
  };

  const eligibleCount = selectedRows.length - blockedRows.length;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box assign-modal-box" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Assign Owner</h3>
          <button type="button" className="modal-close" onClick={onClose}>&times;</button>
        </div>

        <div className="modal-body">
          <p className="assign-modal-subtitle">
            Assigning <strong>{eligibleCount}</strong> enrolment{eligibleCount !== 1 ? 's' : ''} to a new owner.
          </p>

          {blockedRows.length > 0 && (
            <p className="assign-no-queueitem">
              {blockedRows.length} selected enrolment{blockedRows.length !== 1 ? 's' : ''} will be skipped — task status must be Ready or Manual to assign.
            </p>
          )}

          {eligibleCount === 0 ? (
            <p className="assign-no-results">No eligible enrolments selected. Task status must be Ready or Manual.</p>
          ) : (
            <>
              <div className="assign-search-row">
                <input
                  className="assign-search-input"
                  type="text"
                  placeholder="Search users by name or email…"
                  value={searchTerm}
                  onChange={e => setSearchTerm(e.target.value)}
                  autoFocus
                />
              </div>

              {hasSearched && !searching && results.length === 0 && (
                <p className="assign-no-results">No users found.</p>
              )}

              {results.length > 0 && (
                <ul className="assign-results-list">
                  {results.map(u => (
                    <li
                      key={u.systemUserId}
                      className={`assign-result-item${selected?.systemUserId === u.systemUserId ? ' selected' : ''}`}
                      onClick={() => setSelected(u)}
                    >
                      <span className="avatar-circle">{getInitials(u.displayName)}</span>
                      <span className="assign-result-info">
                        <span className="assign-result-name">
                          {u.displayName}
                          {u.group && <span className="assign-group-badge">{u.group}</span>}
                        </span>
                        {(u.jobTitle || u.mail) && (
                          <span className="assign-result-sub">{u.jobTitle ?? u.mail}</span>
                        )}
                      </span>
                    </li>
                  ))}
                </ul>
              )}
            </>
          )}
        </div>

        <div className="modal-footer">
          {eligibleCount > 0 && (
            <button
              type="button"
              className="btn-ok"
              disabled={!selected || submitting}
              onClick={() => void handleAssign()}
            >
              {submitting ? 'Assigning…' : 'Assign'}
            </button>
          )}
          <button type="button" className="btn-cancel" disabled={submitting} onClick={onClose}>
            Cancel
          </button>
          {error && <span className="modal-error">{error}</span>}
        </div>
      </div>
    </div>
  );
}
