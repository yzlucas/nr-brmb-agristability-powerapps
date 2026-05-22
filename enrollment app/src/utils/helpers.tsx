import {
  Vsi_participantprogramyearsvsi_enrolmentstatus,
  Vsi_participantprogramyearsvsi_taskstatus,
  Vsi_participantprogramyearsvsi_enrollmentregionaloffice,
  Vsi_participantprogramyearsvsi_farmingsector,
} from "../generated/models/Vsi_participantprogramyearsModel";
import type { Vsi_participantprogramyears } from "../generated/models/Vsi_participantprogramyearsModel";
import type { SortKey } from '../types/enrollment';

export function getEnrolmentStatusLabel(value: unknown): string {
  if (value == null) return '';
  return Vsi_participantprogramyearsvsi_enrolmentstatus[
    value as keyof typeof Vsi_participantprogramyearsvsi_enrolmentstatus
  ] ?? String(value);
}

const ENROLMENT_STATUS_DISPLAY: Record<string, string> = {
  Initialized: 'Initialized',
  UnverifiedENCalculated: 'Unverified EN Calculated',
  VerifiedENCalculalted: 'Verified EN Calculated',
  EnrolmentNoticeSent: 'Enrolment Notice Sent',
  Enrolled_NotPaid: 'Enrolled, Not Paid',
  Enrolled: 'Enrolled',
  LateEnrolled: 'Late Enrolled',
  Ineligible: 'Ineligible',
  OptedOut: 'Opted Out',
  ToBeReviewed: 'To Be Reviewed',
  _45DayLetter: '45 Day Letter',
  Dormant: 'Dormant',
};

export function formatEnrolmentStatusDisplay(label: string): string {
  return ENROLMENT_STATUS_DISPLAY[label] ?? label;
}

export function getTaskStatusLabel(value: unknown): string {
  if (value == null) return '';
  return Vsi_participantprogramyearsvsi_taskstatus[
    value as keyof typeof Vsi_participantprogramyearsvsi_taskstatus
  ] ?? String(value);
}

export function taskStatusIcon(label: string): React.ReactNode {
  switch (label) {
    case 'Manual':
      return <span className="ts-icon ts-manual" title="Manual">&#x26A0;</span>;
    case 'Supervisor':
      return <span className="ts-icon ts-supervisor" title="Supervisor">&#x1F50D;</span>;
    case 'Ready':
      return <span className="ts-icon ts-ready" title="Ready">&#x2714;</span>;
    case 'Approved':
      return <span className="ts-icon ts-approved" title="Approved">&#x2605;</span>;
    default:
      return null;
  }
}

export function formatCurrency(value: unknown): string {
  if (value == null) return '';
  const n = Number(value);
  return isNaN(n) ? String(value) : '$' + n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function formatCurrencyOr(value: unknown, fallback = ''): string {
  return formatCurrency(value) || fallback;
}

export function calculateVariance(currentFee: unknown, previousFee: unknown): number | null {
  const current = Number(currentFee);
  const previous = Number(previousFee);
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) return null;
  return ((current - previous) / previous) * 100;
}

export function getVarianceClass(variance: number | null): 'neutral' | 'alert' | 'positive' {
  if (variance == null) return 'neutral';
  if (Math.abs(variance) >= 20) return 'alert';
  return variance > 0 ? 'positive' : 'neutral';
}

export function formatVariancePercent(variance: number | null): string {
  if (variance == null) return '';
  return `${variance > 0 ? '+' : ''}${Math.round(variance)}%`;
}

export function getInitials(name: string): string {
  if (!name) return '?';
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

const AVATAR_COLORS = [
  '#4a6fa5', '#2e7d6b', '#7b4f9e', '#c05621', '#1a6b8a',
  '#8e3a59', '#3a7d44', '#6b4226', '#5b4fa5', '#a55b2e',
];

export function getAvatarColor(name: string): string {
  if (!name) return AVATAR_COLORS[0];
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
}

export function getChoiceOptions(field: string): string[] {
  if (field === 'taskStatus') return Object.values(Vsi_participantprogramyearsvsi_taskstatus) as string[];
  if (field === 'enrolStatus') return Object.values(Vsi_participantprogramyearsvsi_enrolmentstatus) as string[];
  if (field === 'hasPartners' || field === 'inCombinedFarm' || field === 'isNewParticipant') return ['Yes', 'No'];
  return [];
}

export function getSortValue(row: Vsi_participantprogramyears, key: SortKey): string | number {
  const raw = row as unknown as Record<string, unknown>;
  switch (key) {
    case 'pin': return row.vsi_name ?? '';
    case 'producer':
      return (row.vsi_participantidname
        ?? raw['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue']
        ?? '') as string;
    case 'year':
      return (row.vsi_programyearidname
        ?? raw['_vsi_programyearid_value@OData.Community.Display.V1.FormattedValue']
        ?? '') as string;
    case 'taskStatus': return getTaskStatusLabel(row.vsi_taskstatus);
    case 'enrolStatus': return getEnrolmentStatusLabel(row.vsi_enrolmentstatus);
    case 'fee': return Number(row.vsi_calculatedenfee) || 0;
    case 'owner': {
      return (row.owneridname
        ?? raw['_ownerid_value@OData.Community.Display.V1.FormattedValue']
        ?? '') as string;
    }
    case 'totalFeesOwedCalculated': return Number(row.vsi_totalfeesowedcalculated) || 0;
    case 'totalFeesPaid': return Number(row.vsi_totalfeespaid) || 0;
    case 'enrolmentFee': return Number(row.vsi_enrolmentfee) || 0;
    case 'latePay': return Number(row.vsi_latepaymentfee) || 0;
    case 'modifiedOn': return row.modifiedon ?? '';
    case 'flagged': {
      const v = row.vsi_variancecalculation != null ? row.vsi_variancecalculation * 100 : null;
      return v != null && Math.abs(v) > 20 ? 1 : 0;
    }
    case 'regionalOffice': return Vsi_participantprogramyearsvsi_enrollmentregionaloffice[row.vsi_enrollmentregionaloffice as keyof typeof Vsi_participantprogramyearsvsi_enrollmentregionaloffice] ?? '';
    case 'farmingSector': return Vsi_participantprogramyearsvsi_farmingsector[row.vsi_farmingsector as keyof typeof Vsi_participantprogramyearsvsi_farmingsector] ?? '';
    case 'isNewParticipant': return row.vsi_isnewparticipant === true ? 1 : 0;
    default: return '';
  }
}

