import { useState, useEffect } from 'react';
import { SystemusersService } from '../generated/services/SystemusersService';
import { QueuemembershipsService } from '../generated/services/QueuemembershipsService';
import { QueuesService } from '../generated/services/QueuesService';
import { TeamsService } from '../generated/services/TeamsService';
import { TeammembershipsService } from '../generated/services/TeammembershipsService';
import { RolesService } from '../generated/services/RolesService';
import { SystemuserrolescollectionService } from '../generated/services/SystemuserrolescollectionService';

export interface AssignableUser {
  systemUserId: string;
  displayName: string;
  jobTitle?: string;
  mail?: string;
  group?: string;
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function chunkArray<T>(items: T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < items.length; i += size) chunks.push(items.slice(i, i + size));
  return chunks;
}

function buildUserIdOrFilter(ids: string[]): string {
  if (ids.length === 0) return 'false';
  return `(${ids.map(id => `systemuserid eq '${id}'`).join(' or ')})`;
}

async function loadUsersByIds(ids: string[]): Promise<AssignableUser[]> {
  if (ids.length === 0) return [];
  const collected = new Map<string, AssignableUser>();
  for (const chunk of chunkArray(ids, 50)) {
    const resp = await SystemusersService.getAll({
      select: ['systemuserid', 'fullname', 'jobtitle', 'internalemailaddress', 'domainname', 'isdisabled'],
      filter: `${buildUserIdOrFilter(chunk)} and isdisabled eq false and not startswith(fullname,'#')`,
      orderBy: ['fullname asc'],
      maxPageSize: 500,
    });
    if (!resp.success) throw new Error(resp.error?.message ?? 'Failed to load users');
    for (const u of resp.data ?? []) {
      if (!u.systemuserid || u.isdisabled || collected.has(u.systemuserid)) continue;
      collected.set(u.systemuserid, {
        systemUserId: u.systemuserid,
        displayName: u.fullname ?? u.internalemailaddress ?? u.domainname ?? u.systemuserid,
        jobTitle: u.jobtitle ?? undefined,
        mail: u.internalemailaddress ?? undefined,
      });
    }
  }
  return [...collected.values()].sort((a, b) => a.displayName.localeCompare(b.displayName));
}

async function getTeamMemberIds(teamName: string): Promise<Set<string>> {
  const teamsResp = await TeamsService.getAll({ filter: `name eq '${teamName}'`, maxPageSize: 1 });
  if (!teamsResp.success || !teamsResp.data?.length) return new Set();
  const teamId = teamsResp.data[0].teamid;
  const membersResp = await TeammembershipsService.getAll({ filter: `teamid eq '${teamId}'`, maxPageSize: 500 });
  if (!membersResp.success || !membersResp.data) return new Set();
  return new Set(membersResp.data.map(m => m.systemuserid).filter((id): id is string => !!id));
}

async function getRoleUserIds(roleName: string): Promise<Set<string>> {
  const rolesResp = await RolesService.getAll({ filter: `name eq '${roleName}'`, maxPageSize: 500 });
  if (!rolesResp.success || !rolesResp.data?.length) return new Set();
  const userIds = new Set<string>();
  await Promise.all(rolesResp.data.map(async role => {
    const resp = await SystemuserrolescollectionService.getAll({ filter: `roleid eq '${role.roleid}'`, maxPageSize: 500 });
    if (resp.success && resp.data) {
      resp.data.forEach(r => { if (r.systemuserid) userIds.add(r.systemuserid); });
    }
  }));
  return userIds;
}

/** Load members of one specific queue by ID. Returns empty set on any error. */
async function getQueueMemberIds(queueId: string): Promise<Set<string>> {
  try {
    const resp = await QueuemembershipsService.getAll({
      select: ['systemuserid'],
      filter: `queueid eq '${queueId}'`,
      maxPageSize: 500,
    });
    if (!resp.success || !resp.data) return new Set();
    return new Set(resp.data.map(m => m.systemuserid).filter((id): id is string => !!id));
  } catch {
    return new Set();
  }
}

/** Discover supervisor approval queues by name and return the union of their members. */
async function getSupervisorApprovalQueueMemberIds(): Promise<Set<string>> {
  const queuesResp = await QueuesService.getAll({
    select: ['queueid'],
    filter: `contains(name,'Supervisor') and contains(name,'Approval')`,
    maxPageSize: 10,
  });
  const memberIds = new Set<string>();
  if (queuesResp.success && queuesResp.data) {
    for (const q of queuesResp.data) {
      if (!q.queueid) continue;
      const ids = await getQueueMemberIds(q.queueid);
      ids.forEach(id => memberIds.add(id));
    }
  }
  return memberIds;
}

function applyGroupBadges(
  users: AssignableUser[],
  sysAdmin: Set<string>,
  admin: Set<string>,
  queue: Set<string>,
  verifier: Set<string>,
): AssignableUser[] {
  return users.map(u => ({
    ...u,
    group: sysAdmin.has(u.systemUserId) ? 'System Administrator'
      : admin.has(u.systemUserId) ? 'Enrolment Admin'
      : queue.has(u.systemUserId) ? 'Supervisor Approval Queue Member'
      : verifier.has(u.systemUserId) ? 'Verifier Team Member'
      : undefined,
  }));
}

// ─── Hook ─────────────────────────────────────────────────────────────────────

export interface UseAssignableUsersResult {
  /** Filtered subset of loaded users matching the current search term. */
  results: AssignableUser[];
  searchTerm: string;
  setSearchTerm: (t: string) => void;
  /** True while the initial user list is loading. */
  searching: boolean;
  loadError: string | null;
  hasLoaded: boolean;
}

/**
 * Loads the set of users eligible for assignment.
 *
 * - If `queueId` is provided, members of that specific queue are loaded.
 * - If `queueId` is omitted, supervisor approval queues are discovered by name.
 * - Either way, Enrolment Admin team, Verifier Team Member team, and System
 *   Administrators are always added to the pool.
 *
 * The returned `results` are filtered client-side by `searchTerm` (name or email).
 *
 * @param queueId - Optional specific queue to scope to.
 * @param enabled - Set to false to skip all loading (e.g. no queue item). Defaults to true.
 */
export function useAssignableUsers(queueId?: string, enabled = true): UseAssignableUsersResult {
  const [allUsers, setAllUsers] = useState<AssignableUser[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [searching, setSearching] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [hasLoaded, setHasLoaded] = useState(false);

  useEffect(() => {
    if (!enabled) return;
    let cancelled = false;
    (async () => {
      setSearching(true);
      setLoadError(null);
      try {
        const [queueMemberIds, adminIds, verifierIds, sysAdminIds] = await Promise.all([
          queueId ? getQueueMemberIds(queueId) : getSupervisorApprovalQueueMemberIds(),
          getTeamMemberIds('Enrolment Admin').catch(() => new Set<string>()),
          getTeamMemberIds('Verifier Team Member').catch(() => new Set<string>()),
          getRoleUserIds('System Administrator').catch(() => new Set<string>()),
        ]);
        if (cancelled) return;
        const allIds = new Set([...queueMemberIds, ...adminIds, ...verifierIds, ...sysAdminIds]);
        const rawUsers = await loadUsersByIds([...allIds]);
        if (cancelled) return;
        setAllUsers(applyGroupBadges(rawUsers, sysAdminIds, adminIds, queueMemberIds, verifierIds));
      } catch (err) {
        if (!cancelled) setLoadError(err instanceof Error ? err.message : 'Failed to load users');
      } finally {
        if (!cancelled) { setSearching(false); setHasLoaded(true); }
      }
    })();
    return () => { cancelled = true; };
  // queueId and enabled are the only inputs that should re-trigger the load
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [queueId, enabled]);

  const results = searchTerm.trim()
    ? allUsers.filter(u => {
        const term = searchTerm.trim().toLowerCase();
        return u.displayName.toLowerCase().includes(term) || (u.mail?.toLowerCase().includes(term) ?? false);
      })
    : allUsers;

  return { results, searchTerm, setSearchTerm, searching, loadError, hasLoaded };
}
