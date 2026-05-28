import { Fragment, useEffect, useMemo, useState } from 'react';
import { CircleCheck, ExternalLink, PanelRightClose, PanelRightOpen, Pin, RefreshCw, Send, User } from 'lucide-react';
import sharepointIconUrl from '/icons/sharepoint.svg?url';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { ApprovalErrorModal } from '../components/ApprovalErrorModal';
import { Send45DayLetterModal } from '../components/Send45DayLetterModal';
import { ConfirmActionModal } from '../components/ConfirmActionModal';
import { ReferToSupervisorModal } from '../components/ReferToSupervisorModal';
import { patchEnrolmentCache } from '../hooks/useEnrolmentData';
import { removeSaItemsFromCache } from './SupervisorApprovalPage';
import { useRole } from '../context/RoleContext';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import { MicrosoftDataverseService } from '../generated/services/MicrosoftDataverseService';
import { QueueitemsService } from '../generated/services/QueueitemsService';
import { Vsi_armsconfigurationsService } from '../generated/services/Vsi_armsconfigurationsService';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { Vsi_programyearsService } from '../generated/services/Vsi_programyearsService';
import { farmsApi } from '../services/farmsApi';
import { resolveCurrentSystemUser } from '../utils/currentUser';
import { formatCurrencyOr, getEnrolmentStatusLabel, getTaskStatusLabel } from '../utils/helpers';

const DATAVERSE_ORG_URL = 'https://aff-brmb-crm-dev.crm3.dynamics.com/';
const BENEFIT_MARGIN_COUNT = 5;
const APPROVABLE_STATUSES = new Set([865520005, 865520006]);
const APPROVABLE_TASK_STATUSES = new Set([865520000, 865520001, 865520002]);
const PARTNER_ENROLMENT_SELECT = [
  'vsi_name',
  '_vsi_participantid_value',
  '_vsi_programyearid_value',
  'vsi_participantidname',
  'vsi_participantprogramyearid',
  'vsi_enrolmentstatus',
  'vsi_enrolmentfee',
  'vsi_calculatedenfee',
  'vsi_sharepointdocumentfolder',
  'vsi_haspartners',
  'vsi_partnershipnames',
  'vsi_partnershippins',
  'vsi_partnershippercents',
  'vsi_incombinedfarm',
  'vsi_combinedfarmpins',
  'vsi_combinedfarmpercents',
  'new_combinedfarmname',
];

type CurrentUser = {
  systemUserId: string;
  displayName: string;
};

type CalculationTableKey = 'enrolmentFee' | 'benefit' | 'proxy' | 'manual';

type CalculationError = string | {
  message?: string;
  errorMessage?: string;
  [key: string]: unknown;
};

type EnwProductiveValue = {
  bpuMargin?: number | null;
  productiveValue?: number | null;
};

type EnwProductiveUnit = {
  code?: string | null;
  description?: string | null;
  productiveCapacity?: number | null;
  productiveValues?: EnwProductiveValue[] | null;
};

type EnwEnrolment = {
  enrolmentFee?: number | null;
  contributionMargin?: number | null;
  benefitMarginYears?: number[] | null;
  proxyMarginYears?: number[] | null;
  enwProductiveUnits?: EnwProductiveUnit[] | null;
  proxyMargins?: Array<number | null> | null;
  benefitContributionMargin?: number | null;
  benefitEnrolmentFee?: number | null;
  proxyContributionMargin?: number | null;
  proxyEnrolmentFee?: number | null;
  manualContributionMargin?: number | null;
  manualEnrolmentFee?: number | null;
  enrolmentCalculationTypeCode?: string | null;
  benefitMarginYearMinus2?: number | null;
  benefitMarginYearMinus3?: number | null;
  benefitMarginYearMinus4?: number | null;
  benefitMarginYearMinus5?: number | null;
  benefitMarginYearMinus6?: number | null;
  benefitMarginYearMinus2Used?: boolean | null;
  benefitMarginYearMinus3Used?: boolean | null;
  benefitMarginYearMinus4Used?: boolean | null;
  benefitMarginYearMinus5Used?: boolean | null;
  benefitMarginYearMinus6Used?: boolean | null;
  manualMarginYearMinus2?: number | null;
  manualMarginYearMinus3?: number | null;
  manualMarginYearMinus4?: number | null;
};

type EnrolmentWorkflowCalculation = {
  benefitCalculationErrors?: CalculationError[] | null;
  enwEnrolment?: EnwEnrolment | null;
};

type PartnerComparisonRow = {
  pin: string;
  name: string;
  percent: string;
  enrolment?: Vsi_participantprogramyears;
};

type XrmWebApiHost = {
  Xrm?: {
    WebApi?: {
      retrieveRecord?: (entityType: string, id: string, options?: string) => Promise<Record<string, unknown>>;
    };
  };
};

const getStringField = (record: unknown, field: string): string => {
  if (!record || typeof record !== 'object') return '';
  const raw = record as Record<string, unknown>;
  const value = raw[field];
  if (typeof value === 'string') return value.trim();

  const dynamicProperties = raw.dynamicProperties;
  if (dynamicProperties && typeof dynamicProperties === 'object') {
    const dynamicValue = (dynamicProperties as Record<string, unknown>)[field];
    if (typeof dynamicValue === 'string') return dynamicValue.trim();
  }

  return '';
};

function normalizeGuid(value?: string | null): string {
  return (value ?? '').replace(/[{}]/g, '').trim().toLowerCase();
}

function normalizeUrlBase(value: string): string {
  return value.replace(/\/+$/, '');
}

async function getFarmsLegacyBaseUrl(): Promise<string | null> {
  const result = await Vsi_armsconfigurationsService.getAll({
    maxPageSize: 50,
  });
  const configRows = result.data ?? [];
  const farmsUrl = configRows
    .map(row => getStringField(row, 'cr4dd_FARMSURLNEW') || getStringField(row, 'cr4dd_farmsurlnew'))
    .find((candidate): candidate is string => !!candidate);
  return farmsUrl ? normalizeUrlBase(farmsUrl) : null;
}

async function getAccountFromXrm(accountId: string): Promise<Record<string, unknown> | null> {
  const candidates = [window, window.parent, window.top];
  for (const candidate of candidates) {
    try {
      if (!candidate) continue;
      const retrieveRecord = (candidate as unknown as XrmWebApiHost).Xrm?.WebApi?.retrieveRecord;
      if (!retrieveRecord) continue;
      return await retrieveRecord('account', accountId, '?$select=vsi_pin,accountnumber,name');
    } catch {
      // Try the next window context.
    }
  }
  return null;
}

async function getAccountPinFromXrm(accountId: string): Promise<string> {
  const account = await getAccountFromXrm(accountId);
  let pin = getStringField(account, 'vsi_pin');
  if (!pin) pin = getStringField(account, 'accountnumber');
  return pin;
}

async function getAccountPin(accountId: string): Promise<string> {
  const genericAccount = await MicrosoftDataverseService.GetItemWithOrganization(
    '',
    'application/json',
    DATAVERSE_ORG_URL,
    'accounts',
    accountId,
    false,
    false,
    'vsi_pin,accountnumber,name',
  );
  let pin = getStringField(genericAccount.data, 'vsi_pin');
  if (!pin) pin = getStringField(genericAccount.data, 'accountnumber');
  if (pin) return pin;

  return getAccountPinFromXrm(accountId);
}

function getProgramYear(record: Vsi_participantprogramyears | null): number | null {
  const raw = record as unknown as Record<string, unknown> | null;
  const candidates = [
    record?.vsi_programyearidname,
    record?.vsi_name,
    raw?.['_vsi_programyearid_value@OData.Community.Display.V1.FormattedValue'],
  ];
  const match = candidates
    .filter((value): value is string => typeof value === 'string')
    .map(value => value.match(/\b(19|20)\d{2}\b/))
    .find((result): result is RegExpMatchArray => result != null);
  return match ? Number(match[0]) : null;
}

function getBooleanText(value: unknown): string {
  if (value === true || value === 1 || value === '1') return 'Yes';
  if (value === false || value === 0 || value === '0') return 'No';
  return '';
}

function formatNumberOrBlank(value: unknown, fractionDigits: number): string {
  if (value == null || value === '') return '';
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return String(value);
  return numberValue.toLocaleString(undefined, {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  });
}

function formatCurrencyBlank(value: unknown): string {
  return formatCurrencyOr(value, '');
}

function formatPercentText(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) return '';
  return trimmed.includes('%') ? trimmed : `${trimmed}%`;
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function getCalculationHref(id: string, routeSource: string | undefined): string {
  return `#/calculation/${routeSource || 'dashboard'}/${id}`;
}

function buildFarmsScenarioUrl(baseUrl: string, pinValue: string, scenarioProgramYear: number | null): string {
  if (!baseUrl || !pinValue || !scenarioProgramYear) return '';
  const params = new URLSearchParams({
    pin: pinValue,
    year: String(scenarioProgramYear),
    refresh: 'true',
  });
  return `${baseUrl}/farm800.do?${params.toString()}`;
}

function parsePinList(value?: string | null): string[] {
  return (value?.match(/\d+/g) ?? []).map(pinValue => pinValue.trim()).filter(Boolean);
}

function splitSummaryList(value: string | null | undefined, expectedCount: number): string[] {
  const trimmed = value?.trim();
  if (!trimmed) return [];

  const preferred = trimmed
    .split(/\r?\n|;|\|/)
    .map(item => item.trim())
    .filter(Boolean);
  if (preferred.length === expectedCount) return preferred;

  const commaSeparated = trimmed
    .split(',')
    .map(item => item.trim())
    .filter(Boolean);
  if (commaSeparated.length === expectedCount) return commaSeparated;

  return preferred.length > 1 ? preferred : [trimmed];
}

function parsePercentList(value: string | null | undefined, expectedCount: number): string[] {
  const matches = value?.match(/\d+(?:\.\d+)?%?/g) ?? [];
  if (matches.length) return matches.map(formatPercentText);
  return splitSummaryList(value, expectedCount).map(formatPercentText);
}

function getPartnerSummaries(record: Vsi_participantprogramyears | null): PartnerComparisonRow[] {
  const pins = parsePinList(record?.vsi_partnershippins);
  const names = splitSummaryList(record?.vsi_partnershipnames, pins.length);
  const percents = parsePercentList(record?.vsi_partnershippercents, pins.length);

  return pins.map((pinValue, index) => ({
    pin: pinValue,
    name: names[index] ?? '',
    percent: percents[index] ?? '',
  }));
}

async function findProgramYearId(year: number): Promise<string | null> {
  const safeYear = escapeODataString(String(year));
  const result = await Vsi_programyearsService.getAll({
    maxPageSize: 1,
    top: 1,
    select: ['vsi_programyearid', 'vsi_year'],
    filter: `vsi_year eq '${safeYear}'`,
  });
  return result.data?.[0]?.vsi_programyearid ?? null;
}

async function findPartnerEnrolment(pinValue: string, programYearId: string): Promise<Vsi_participantprogramyears | null> {
  const safePin = escapeODataString(pinValue);
  const safeProgramYearId = escapeODataString(normalizeGuid(programYearId));

  const exact = await Vsi_participantprogramyearsService.getAll({
    maxPageSize: 1,
    top: 1,
    select: PARTNER_ENROLMENT_SELECT,
    filter: `_vsi_programyearid_value eq '${safeProgramYearId}' and vsi_name eq '${safePin}'`,
  });
  if (exact.data?.[0]) return exact.data[0];

  const containsPin = await Vsi_participantprogramyearsService.getAll({
    maxPageSize: 1,
    top: 1,
    select: PARTNER_ENROLMENT_SELECT,
    filter: `_vsi_programyearid_value eq '${safeProgramYearId}' and contains(vsi_name,'${safePin}')`,
  });
  return containsPin.data?.[0] ?? null;
}

function getFarmsWorkflowErrorMessage(error: unknown): string {
  const fallback = 'Unable to load FARMS enrolment calculation.';
  const message = error instanceof Error ? error.message : String(error || fallback);
  return message.includes('404') ? 'No ENW scenario has been created in FARMS.' : message;
}

function getErrorText(error: CalculationError): string {
  if (typeof error === 'string') return error;
  return error.message ?? error.errorMessage ?? JSON.stringify(error);
}

function getErrorSearchText(error: CalculationError): string {
  if (typeof error === 'string') return error.toLowerCase();
  return Object.values(error)
    .filter(value => typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')
    .join(' ')
    .toLowerCase();
}

function matchesCalculationTable(error: CalculationError, table: CalculationTableKey): boolean {
  const text = getErrorSearchText(error);
  if (!text) return false;
  switch (table) {
    case 'enrolmentFee':
      return text.includes('enrolment fee') || text.includes('enrollment fee') || text.includes('enrolmentfee') || text.includes('enrollmentfee');
    case 'benefit':
      return text.includes('benefit');
    case 'proxy':
      return text.includes('proxy') || text.includes('productive') || text.includes('bpu');
    case 'manual':
      return text.includes('manual');
    default:
      return false;
  }
}

function getTableErrorMessages(calculation: EnrolmentWorkflowCalculation | null, table: CalculationTableKey): string[] {
  return (calculation?.benefitCalculationErrors ?? [])
    .filter(error => matchesCalculationTable(error, table))
    .map(getErrorText);
}

function getUnmatchedErrorMessages(calculation: EnrolmentWorkflowCalculation | null): string[] {
  return (calculation?.benefitCalculationErrors ?? [])
    .filter(error => !matchesCalculationTable(error, 'enrolmentFee')
      && !matchesCalculationTable(error, 'benefit')
      && !matchesCalculationTable(error, 'proxy')
      && !matchesCalculationTable(error, 'manual'))
    .map(getErrorText);
}

function CalculationErrorMessages({ messages }: { messages: string[] }) {
  if (!messages.length) return null;
  return (
    <div className="calc-legacy-error" role="alert">
      {messages.map((message, index) => (
        <div key={`${message}-${index}`}>{message}</div>
      ))}
    </div>
  );
}

function CalculationOption({ checked, label }: { checked: boolean; label: string }) {
  return (
    <label className="calc-benefit-option">
      <input
        className="calc-benefit-radio"
        type="radio"
        checked={checked}
        readOnly
        tabIndex={-1}
        onChange={() => undefined}
      />
      <span className={`calc-benefit-radio-visual${checked ? ' calc-benefit-radio-visual-checked' : ''}`} aria-hidden="true" />
      <span>{label}</span>
    </label>
  );
}

function PartnerViewPanel({
  rows,
  loading,
  error,
  farmsLegacyBaseUrl,
  farmsScenarioProgramYear,
  routeSource,
  open,
  pinned,
  onToggleOpen,
  onTogglePinned,
}: {
  rows: PartnerComparisonRow[];
  loading: boolean;
  error: string | null;
  farmsLegacyBaseUrl: string;
  farmsScenarioProgramYear: number | null;
  routeSource?: string;
  open: boolean;
  pinned: boolean;
  onToggleOpen: () => void;
  onTogglePinned: () => void;
}) {
  if (!open) {
    return (
      <aside className="calc-comparison-panel calc-comparison-panel-collapsed" aria-label="Partner comparison panel">
        <button className="calc-panel-tab" type="button" onClick={onToggleOpen} title="Expand partner panel" aria-label="Expand partner panel">
          <PanelRightOpen size={14} aria-hidden="true" />
          <span>Partner view</span>
        </button>
      </aside>
    );
  }

  return (
    <aside className="calc-comparison-panel" aria-label="Partner comparison panel">
      <div className="calc-comparison-header">
        <h2>Partner view</h2>
        <div className="calc-comparison-actions">
          <button
            className={`calc-panel-icon-btn${pinned ? ' calc-panel-icon-btn-active' : ''}`}
            type="button"
            onClick={onTogglePinned}
            title={pinned ? 'Unpin panel' : 'Pin panel'}
            aria-label={pinned ? 'Unpin partner panel' : 'Pin partner panel'}
          >
            <Pin size={14} aria-hidden="true" />
          </button>
          <button className="calc-panel-icon-btn calc-panel-icon-btn-square" type="button" onClick={onToggleOpen} title="Collapse partner panel" aria-label="Collapse partner panel">
            <PanelRightClose size={14} aria-hidden="true" />
          </button>
        </div>
      </div>

      {loading && <p className="calc-panel-state">Loading partners...</p>}
      {error && <p className="calc-panel-state calc-panel-state-error">{error}</p>}
      {!loading && !error && !rows.length && <p className="calc-panel-state">No partner data found.</p>}
      {!loading && !error && rows.length > 0 && (
        <div className="calc-partner-list">
          {rows.map(row => {
            const enrolment = row.enrolment;
            const name = row.name || enrolment?.vsi_participantidname || '';
            const status = enrolment ? getEnrolmentStatusLabel(enrolment.vsi_enrolmentstatus) : '';
            const farmsUrl = buildFarmsScenarioUrl(farmsLegacyBaseUrl, row.pin, farmsScenarioProgramYear);
            return (
              <div className="calc-partner-card" key={row.pin}>
                <div className="calc-partner-card-top">
                  <div>
                    <div className="calc-partner-name">{name || '-'}</div>
                    <div className="calc-partner-pin">
                      {enrolment?.vsi_participantprogramyearid ? (
                        <a href={getCalculationHref(enrolment.vsi_participantprogramyearid, routeSource)} target="_blank" rel="noopener noreferrer">
                          PIN {row.pin}
                        </a>
                      ) : (
                        <span>PIN {row.pin}</span>
                      )}
                    </div>
                  </div>
                  {farmsUrl ? (
                    <a className="calc-partner-farms-link" href={farmsUrl} target="_blank" rel="noopener noreferrer" title="Open FARMS scenario" aria-label={`Open FARMS scenario for PIN ${row.pin}`}>
                      <ExternalLink size={15} aria-hidden="true" />
                    </a>
                  ) : (
                    <span className="calc-partner-farms-link calc-partner-farms-link-disabled" aria-label="FARMS link unavailable">
                      <ExternalLink size={15} aria-hidden="true" />
                    </span>
                  )}
                </div>
                <dl className="calc-partner-details">
                  <div>
                    <dt>Fee</dt>
                    <dd>{formatCurrencyBlank(enrolment?.vsi_enrolmentfee)}</dd>
                  </div>
                  <div>
                    <dt>EN Status</dt>
                    <dd>{status}</dd>
                  </div>
                  {row.percent && (
                    <div>
                      <dt>Percent</dt>
                      <dd>{row.percent}</dd>
                    </div>
                  )}
                </dl>
              </div>
            );
          })}
        </div>
      )}
    </aside>
  );
}

function getApprovalError(
  record: Vsi_participantprogramyears,
): string | null {
  const enrolmentName = record.vsi_name ?? 'This enrolment';

  if (!APPROVABLE_STATUSES.has(record.vsi_enrolmentstatus as unknown as number)) {
    return `${enrolmentName} cannot be approved because its enrolment status is not Verified EN Calculated or Unverified EN Calculated.`;
  }

  if (!APPROVABLE_TASK_STATUSES.has(record.vsi_taskstatus as unknown as number)) {
    return `${enrolmentName} cannot be approved because its task status must be Manual, Supervisor, or Ready.`;
  }

  if (record.vsi_calculatedenfee == null) {
    return `${enrolmentName} cannot be approved because it does not have a calculated fee.`;
  }

  return null;
}

export function EnrolmentCalculationPage() {
  const { enrolmentId, source } = useParams<{ enrolmentId: string; source: string }>();
  const navigate = useNavigate();
  const { activeRole } = useRole();
  const backTo = source === 'supervisor' ? '/supervisor-approval' : '/dashboard-home';
  const backLabel = source === 'supervisor' ? 'Back to Supervisor Approval' : 'Back to Dashboard';
  const [record, setRecord] = useState<Vsi_participantprogramyears | null>(null);
  const [participantPin, setParticipantPin] = useState('');
  const [participantPinLoading, setParticipantPinLoading] = useState(false);
  const [farmsLegacyBaseUrl, setFarmsLegacyBaseUrl] = useState('');
  const [farmsLegacyBaseUrlLoading, setFarmsLegacyBaseUrlLoading] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [refreshKey, setRefreshKey] = useState(0);
  const [showSupervisorModal, setShowSupervisorModal] = useState(false);
  const [currentUser, setCurrentUser] = useState<CurrentUser | null>(null);
  const [approvalErrorModal, setApprovalErrorModal] = useState<string | null>(null);
  const [showApproveConfirm, setShowApproveConfirm] = useState(false);
  const [approving, setApproving] = useState(false);
  const [showCompleteConfirm, setShowCompleteConfirm] = useState(false);
  const [completing, setCompleting] = useState(false);
  const [farmsWorkflowCalculation, setFarmsWorkflowCalculation] = useState<EnrolmentWorkflowCalculation | null>(null);
  const [farmsWorkflowCalculationLoading, setFarmsWorkflowCalculationLoading] = useState(false);
  const [farmsWorkflowCalculationError, setFarmsWorkflowCalculationError] = useState<string | null>(null);
  const [show45DayModal, setShow45DayModal] = useState(false);
  const [letterSentMessage, setLetterSentMessage] = useState<string | null>(null);
  const [counterActionLoading, setCounterActionLoading] = useState(false);
  const [counterActionError, setCounterActionError] = useState<string | null>(null);
  const [partnerRows, setPartnerRows] = useState<PartnerComparisonRow[]>([]);
  const [partnerRowsLoading, setPartnerRowsLoading] = useState(false);
  const [partnerRowsError, setPartnerRowsError] = useState<string | null>(null);
  const [partnerPanelOpen, setPartnerPanelOpen] = useState(true);
  const [partnerPanelPinned, setPartnerPanelPinned] = useState(true);
  const [queueWorkerName, setQueueWorkerName] = useState<string | null>(null);
  const [queueWorkerId, setQueueWorkerId] = useState<string | null>(null);

  useEffect(() => {
    if (!enrolmentId) {
      setError('Missing enrolment id.');
      setLoading(false);
      return;
    }

    let cancelled = false;
    (async () => {
      try {
        setLoading(true);
        setError(null);
        let result = await Vsi_participantprogramyearsService.get(enrolmentId, {
          select: [
            'vsi_name',
            'vsi_taskstatus',
            'vsi_enrolmentstatus',
            '_ownerid_value',
            'owneridname',
            'vsi_calculatedenfee',
            'vsi_previousyearcalculatedenfee',
            'modifiedon',
            '_vsi_programyearid_value',
            'vsi_programyearidname',
            '_vsi_participantid_value',
            'vsi_participantidname',
            'vsi_sharepointdocumentfolder',
            'vsi_contributionmargin',
            'vsi_enrolmentfee',
            'vsi_programyearmargin1',
            'vsi_programyearmargin1used',
            'vsi_programyearmargin2',
            'vsi_programyearmargin2used',
            'vsi_programyearmargin3',
            'vsi_programyearmargin3used',
            'vsi_programyearmargin4',
            'vsi_programyearmargin4used',
            'vsi_programyearmargin5',
            'vsi_programyearmargin5used',
            'vsi_fortyfivedayletterstartdate',
            'vsi_fortyfivedaylettersent',
            'vsi_fortyfivedaycounterpaused',
            'vsi_fortyfivedaypausedate',
            'vsi_haspartners',
            'vsi_partnershipnames',
            'vsi_partnershippins',
            'vsi_partnershippercents',
          ],
        });

        // Match the details page behavior: some environments can return no data
        // for a selected retrieve even when the full record is readable.
        if (!result?.data) {
          result = await Vsi_participantprogramyearsService.get(enrolmentId);
        }

        if (cancelled) return;
        if (!result.data) {
          setError(result.error?.message ?? 'Unable to load enrolment calculation data.');
          setRecord(null);
          return;
        }

        setRecord(result.data);

        const participantId = result.data._vsi_participantid_value?.replace(/[{}]/g, '');
        setParticipantPin('');
        if (participantId) {
          setParticipantPinLoading(true);
          try {
            const pin = await getAccountPin(participantId);
            if (!cancelled) setParticipantPin(pin);
          } catch {
            if (!cancelled) setParticipantPin('');
          } finally {
            if (!cancelled) setParticipantPinLoading(false);
          }
        }
      } catch (err) {
        if (!cancelled) setError(String(err));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [enrolmentId, refreshKey]);

  // Fetch queue worker (who has "picked" this enrolment in the supervisor queue)
  useEffect(() => {
    if (!enrolmentId) {
      setQueueWorkerName(null);
      setQueueWorkerId(null);
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const [queueResult, user] = await Promise.all([
          QueueitemsService.getAll({
            filter: `objectid_vsi_participantprogramyear/vsi_participantprogramyearid eq '${enrolmentId}' and statecode eq 0`,
            select: ['queueitemid', '_workerid_value'],
            maxPageSize: 1,
          }),
          resolveCurrentSystemUser(),
        ]);
        if (cancelled) return;
        setCurrentUser({ systemUserId: user.systemUserId, displayName: user.displayName });
        const q = queueResult.data?.[0];
        if (q && q._workerid_value) {
          const raw = q as unknown as Record<string, unknown>;
          setQueueWorkerName((raw['_workerid_value@OData.Community.Display.V1.FormattedValue'] as string) ?? null);
          setQueueWorkerId(q._workerid_value);
        } else {
          setQueueWorkerName(null);
          setQueueWorkerId(null);
        }
      } catch {
        if (!cancelled) {
          setQueueWorkerName(null);
          setQueueWorkerId(null);
        }
      }
    })();
    return () => { cancelled = true; };
  }, [enrolmentId, refreshKey]);

  useEffect(() => {
    let cancelled = false;
    setFarmsLegacyBaseUrlLoading(true);
    getFarmsLegacyBaseUrl()
      .then((url) => {
        if (!cancelled) setFarmsLegacyBaseUrl(url ?? '');
      })
      .catch(() => {
        if (!cancelled) setFarmsLegacyBaseUrl('');
      })
      .finally(() => {
        if (!cancelled) setFarmsLegacyBaseUrlLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [refreshKey]);

  const sharePointFolderUrl = record?.vsi_sharepointdocumentfolder;
  const programYear = useMemo(() => getProgramYear(record), [record]);
  const farmsScenarioProgramYear = programYear ? programYear - 2 : null;
  const currentProgramYearId = record?._vsi_programyearid_value ?? null;
  const farmsScenarioUrl = useMemo(() => {
    return buildFarmsScenarioUrl(farmsLegacyBaseUrl, participantPin, farmsScenarioProgramYear);
  }, [farmsLegacyBaseUrl, participantPin, farmsScenarioProgramYear]);

  useEffect(() => {
    if (!participantPin || !farmsScenarioProgramYear) {
      setFarmsWorkflowCalculation(null);
      setFarmsWorkflowCalculationError(null);
      setFarmsWorkflowCalculationLoading(false);
      return;
    }

    let cancelled = false;
    setFarmsWorkflowCalculation(null);
    setFarmsWorkflowCalculationError(null);
    setFarmsWorkflowCalculationLoading(true);

    farmsApi.getEnrolmentNoticeWorkflowCalculation<EnrolmentWorkflowCalculation>(
      participantPin,
      farmsScenarioProgramYear,
    )
      .then((result) => {
        if (cancelled) return;
        if (!result.success) {
          throw new Error(result.error?.message ?? 'Unable to load FARMS enrolment calculation.');
        }
        setFarmsWorkflowCalculation(result.data ?? null);
      })
      .catch((err) => {
        if (cancelled) return;
        setFarmsWorkflowCalculation(null);
        setFarmsWorkflowCalculationError(getFarmsWorkflowErrorMessage(err));
      })
      .finally(() => {
        if (!cancelled) setFarmsWorkflowCalculationLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [participantPin, farmsScenarioProgramYear]);

  useEffect(() => {
    if (!participantPin || !programYear || !farmsScenarioProgramYear) {
      setPartnerRows([]);
      setPartnerRowsError(null);
      setPartnerRowsLoading(false);
      return;
    }

    let cancelled = false;
    setPartnerRows([]);
    setPartnerRowsError(null);
    setPartnerRowsLoading(true);

    (async () => {
      try {
        const lookupProgramYearId = await findProgramYearId(farmsScenarioProgramYear);
        if (cancelled) return;
        if (!lookupProgramYearId) {
          setPartnerRowsError(`Unable to load partner data because lookup program year ${farmsScenarioProgramYear} was not found.`);
          setPartnerRows([]);
          return;
        }

        const displayProgramYearId = currentProgramYearId
          ? normalizeGuid(currentProgramYearId)
          : await findProgramYearId(programYear);
        if (cancelled) return;
        if (!displayProgramYearId) {
          setPartnerRowsError(`Unable to load partner data because enrolment year ${programYear} was not found.`);
          setPartnerRows([]);
          return;
        }

        const sourceEnrolment = await findPartnerEnrolment(participantPin, lookupProgramYearId);
        if (cancelled) return;
        if (!sourceEnrolment) {
          setPartnerRows([]);
          return;
        }

        const summaries = getPartnerSummaries(sourceEnrolment);
        setPartnerRows(summaries);

        const rows = await Promise.all(summaries.map(async (summary) => ({
          ...summary,
          enrolment: await findPartnerEnrolment(summary.pin, displayProgramYearId) ?? undefined,
        })));
        if (!cancelled) setPartnerRows(rows);
      } catch (err) {
        if (cancelled) return;
        setPartnerRows([]);
        setPartnerRowsError(err instanceof Error ? err.message : 'Unable to load partner enrolment rows.');
      } finally {
        if (!cancelled) setPartnerRowsLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [participantPin, programYear, farmsScenarioProgramYear, currentProgramYearId]);

  const participantName = useMemo(() => {
    if (!record) return '';
    const raw = record as unknown as Record<string, unknown>;
    return (record.vsi_participantidname
      ?? raw['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue']
      ?? '') as string;
  }, [record]);
  const ownerName = record
    ? (record.owneridname || ((record as unknown as Record<string, unknown>)['_ownerid_value@OData.Community.Display.V1.FormattedValue'] as string | undefined) || null)
    : null;
  const taskStatusLabel = getTaskStatusLabel(record?.vsi_taskstatus) || null;
  const workedByMe = !!(queueWorkerId && currentUser && normalizeGuid(queueWorkerId) === normalizeGuid(currentUser.systemUserId));
  const fallbackBenefitYears = useMemo(() => {
    return Array.from({ length: BENEFIT_MARGIN_COUNT }, (_, index) => (
      programYear ? String(programYear - (BENEFIT_MARGIN_COUNT + 1) + index) : `Year ${index + 1}`
    ));
  }, [programYear]);
  const farmsEnrolment = farmsWorkflowCalculation?.enwEnrolment ?? null;
  const calculationTypeCode = (farmsEnrolment?.enrolmentCalculationTypeCode ?? '').toUpperCase();
  const benefitMarginRows = useMemo(() => {
    const apiYears = farmsEnrolment?.benefitMarginYears?.map(String) ?? [];
    const labels = apiYears.length ? apiYears : fallbackBenefitYears;
    const apiMargins = [
      farmsEnrolment?.benefitMarginYearMinus6,
      farmsEnrolment?.benefitMarginYearMinus5,
      farmsEnrolment?.benefitMarginYearMinus4,
      farmsEnrolment?.benefitMarginYearMinus3,
      farmsEnrolment?.benefitMarginYearMinus2,
    ];
    const apiUsed = [
      farmsEnrolment?.benefitMarginYearMinus6Used,
      farmsEnrolment?.benefitMarginYearMinus5Used,
      farmsEnrolment?.benefitMarginYearMinus4Used,
      farmsEnrolment?.benefitMarginYearMinus3Used,
      farmsEnrolment?.benefitMarginYearMinus2Used,
    ];

    return labels.map((label, index) => {
      const position = index + 1;
      const raw = (record ?? {}) as unknown as Record<string, unknown>;
      return {
        label,
        margin: apiYears.length ? apiMargins[index] : raw[`vsi_programyearmargin${position}`],
        used: apiYears.length ? apiUsed[index] : raw[`vsi_programyearmargin${position}used`],
      };
    });
  }, [fallbackBenefitYears, farmsEnrolment, record]);
  const proxyYears = useMemo(() => {
    if (farmsEnrolment?.proxyMarginYears?.length) return farmsEnrolment.proxyMarginYears.map(String);
    if (farmsScenarioProgramYear) {
      return [farmsScenarioProgramYear - 2, farmsScenarioProgramYear - 1, farmsScenarioProgramYear].map(String);
    }
    return ['Year 1', 'Year 2', 'Year 3'];
  }, [farmsEnrolment?.proxyMarginYears, farmsScenarioProgramYear]);
  const manualMargins = [
    farmsEnrolment?.manualMarginYearMinus4,
    farmsEnrolment?.manualMarginYearMinus3,
    farmsEnrolment?.manualMarginYearMinus2,
  ];
  const tableErrors = {
    enrolmentFee: getTableErrorMessages(farmsWorkflowCalculation, 'enrolmentFee'),
    benefit: getTableErrorMessages(farmsWorkflowCalculation, 'benefit'),
    proxy: getTableErrorMessages(farmsWorkflowCalculation, 'proxy'),
    manual: getTableErrorMessages(farmsWorkflowCalculation, 'manual'),
    unmatched: getUnmatchedErrorMessages(farmsWorkflowCalculation),
  };

  const handle45DayPause = async () => {
    if (!record || !enrolmentId) return;
    setCounterActionLoading(true);
    setCounterActionError(null);
    try {
      const today = new Date().toISOString();
      const patch: Partial<Vsi_participantprogramyears> = {
        vsi_fortyfivedaycounterpaused: true,
        vsi_fortyfivedaypausedate: today,
      };
      const result = await Vsi_participantprogramyearsService.update(enrolmentId, patch);
      if (!result.success) throw new Error(result.error?.message ?? 'Failed to pause counter.');
      setRecord(prev => prev ? { ...prev, ...patch } : prev);
      patchEnrolmentCache([{ id: enrolmentId, fields: patch }]);
    } catch (err) {
      setCounterActionError(err instanceof Error ? err.message : 'Failed to pause counter.');
    } finally {
      setCounterActionLoading(false);
    }
  };

  const handle45DayResume = async () => {
    if (!record || !enrolmentId) return;
    setCounterActionLoading(true);
    setCounterActionError(null);
    try {
      const pauseDate = record.vsi_fortyfivedaypausedate;
      const startDate = record.vsi_fortyfivedayletterstartdate;
      if (!pauseDate || !startDate) throw new Error('Cannot resume: pause date or start date is missing.');
      const pausedDays = Math.floor((Date.now() - new Date(pauseDate).getTime()) / (1000 * 60 * 60 * 24));
      const newStartDate = new Date(new Date(startDate).getTime() + pausedDays * 24 * 60 * 60 * 1000).toISOString();
      const resumeResult = await Vsi_participantprogramyearsService.update(enrolmentId, {
        vsi_fortyfivedaycounterpaused: false,
        vsi_fortyfivedayletterstartdate: newStartDate,
        vsi_fortyfivedaypausedate: null as unknown as string,
      });
      if (!resumeResult.success) throw new Error(resumeResult.error?.message ?? 'Failed to resume counter.');
      const patch: Partial<Vsi_participantprogramyears> = {
        vsi_fortyfivedaycounterpaused: false,
        vsi_fortyfivedaypausedate: undefined,
        vsi_fortyfivedayletterstartdate: newStartDate,
      };
      setRecord(prev => prev ? { ...prev, ...patch } : prev);
      patchEnrolmentCache([{ id: enrolmentId, fields: patch }]);
    } catch (err) {
      setCounterActionError(err instanceof Error ? err.message : 'Failed to resume counter.');
    } finally {
      setCounterActionLoading(false);
    }
  };

  const handleCompleteConfirm = async () => {
    if (!record || !enrolmentId) return;
    setCompleting(true);
    setError(null);
    try {
      const result = await Vsi_participantprogramyearsService.update(enrolmentId, {
        vsi_taskstatus: 865520002, // Ready
        vsi_enrolmentstatus: 865520006, // VerifiedENCalculalted
      });
      if (!result.success) {
        throw new Error(result.error?.message ?? `Failed to complete ${enrolmentId}.`);
      }
      const completedFields: Partial<Vsi_participantprogramyears> = {
        vsi_taskstatus: 865520002 as unknown as Vsi_participantprogramyears['vsi_taskstatus'],
        vsi_enrolmentstatus: 865520006 as unknown as Vsi_participantprogramyears['vsi_enrolmentstatus'],
      };
      patchEnrolmentCache([{ id: enrolmentId, fields: completedFields }]);
      setRecord(prev => prev ? { ...prev, ...completedFields } : prev);
      setShowCompleteConfirm(false);
      setRefreshKey(prev => prev + 1);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Complete failed.');
    } finally {
      setCompleting(false);
    }
  };

  const handleApproveClick = async () => {
    if (!record) return;

    try {
      const approvalError = getApprovalError(record);
      if (approvalError) {
        setApprovalErrorModal(approvalError);
        return;
      }

      setShowApproveConfirm(true);
    } catch (err) {
      setApprovalErrorModal(err instanceof Error ? err.message : 'Unable to validate approval.');
    }
  };

  const handleApproveConfirm = async () => {
    if (!record || !enrolmentId) return;

    setApproving(true);
    setError(null);
    try {
      const approvalError = getApprovalError(record);
      if (approvalError) {
        setShowApproveConfirm(false);
        setApprovalErrorModal(approvalError);
        return;
      }

      const statusUpdateResult = await Vsi_participantprogramyearsService.update(enrolmentId, {
        vsi_taskstatus: 865520003,
      });
      if (!statusUpdateResult.success) {
        throw new Error(statusUpdateResult.error?.message ?? `Failed to set Approved status for ${enrolmentId}.`);
      }

      // Remove all active queue items for this enrolment (including supervisor queue)
      try {
        const allQueueItems = await QueueitemsService.getAll({
          filter: `objectid_vsi_participantprogramyear/vsi_participantprogramyearid eq '${enrolmentId}' and statecode eq 0`,
          select: ['queueitemid'],
        });
        for (const qi of allQueueItems.data ?? []) {
          await QueueitemsService.delete(qi.queueitemid);
        }
      } catch {
        // ignore queue cleanup errors
      }

      const approvedFields: Partial<Vsi_participantprogramyears> = {
        vsi_taskstatus: 865520003 as unknown as Vsi_participantprogramyears['vsi_taskstatus'],
      };
      patchEnrolmentCache([{ id: enrolmentId, fields: approvedFields }]);
      if (source === 'supervisor') removeSaItemsFromCache([enrolmentId]);
      setRecord(prev => prev ? { ...prev, ...approvedFields } : prev);
      setShowApproveConfirm(false);
      setRefreshKey(prev => prev + 1);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Approve failed.');
    } finally {
      setApproving(false);
    }
  };


  return (
    <section className="page-card calc-page">
      <div className="calc-title-row">
        <button type="button" className="calc-back-btn" onClick={() => navigate(backTo)}>{backLabel}</button>
        <span className="calc-page-label">Enrolment App / FARMS Calculation</span>
        {(record || loading) && (
          <div className="calc-meta-strip">
            <span className="calc-meta-owner">
              <User size={13} aria-hidden="true" />
              {ownerName || (loading ? '...' : '—')}
            </span>
            {taskStatusLabel && (
              <span className={`calc-task-badge calc-task-badge--${taskStatusLabel.toLowerCase()}`}>
                {taskStatusLabel}
              </span>
            )}
            {queueWorkerName && (
              <>
                <span className="calc-meta-sep" aria-hidden="true">•</span>
                <span className="calc-meta-worked-by">
                  Worked by: {workedByMe ? `${queueWorkerName} (you)` : queueWorkerName}
                </span>
              </>
            )}
          </div>
        )}
      </div>

      <div className="calc-identity">
        <Link to={`/enrolment/${source}/${enrolmentId}`} className="calc-enrolment-name-link">
          {record?.vsi_name ?? (loading ? 'Loading...' : '-')}
        </Link>
        <h1 className="calc-participant-name">{participantName || (loading ? 'Loading...' : '-')}</h1>
      </div>

      <div className="calc-toolbar">
        <button className="calc-outline-btn" type="button" onClick={() => setShow45DayModal(true)}>Send 45-Day Letter</button>
        <button
          className="calc-outline-btn"
          type="button"
          onClick={() => setRefreshKey(prev => prev + 1)}
          disabled={loading}
        >
          <RefreshCw size={14} aria-hidden="true" />
          Refresh
        </button>
        <button
          className="calc-outline-btn"
          type="button"
          onClick={() => setShowSupervisorModal(true)}
          disabled={!record}
        >
          <Send size={14} aria-hidden="true" />
          Refer to Supervisor
        </button>
        {activeRole !== 'Verifier' && (
          <button
            className="calc-outline-btn"
            type="button"
            onClick={() => void handleApproveClick()}
            disabled={!record || approving}
          >
            <CircleCheck size={14} aria-hidden="true" />
            {approving ? 'Approving...' : 'Approve'}
          </button>
        )}
        {activeRole === 'Verifier' && (
          <button
            className="calc-outline-btn"
            type="button"
            onClick={() => setShowCompleteConfirm(true)}
            disabled={!record || completing}
          >
            <CircleCheck size={14} aria-hidden="true" />
            {completing ? 'Completing...' : 'Complete'}
          </button>
        )}
        <div className="calc-toolbar-gap" />
        {farmsScenarioUrl ? (
          <a
            className="calc-outline-btn calc-sharepoint-btn"
            href={farmsScenarioUrl}
            target="_blank"
            rel="noopener noreferrer"
          >
            <ExternalLink size={14} aria-hidden="true" />
            Open a Scenario in FARMS
          </a>
        ) : (
          <button
            className="calc-outline-btn calc-sharepoint-btn"
            type="button"
            disabled
            title={loading || participantPinLoading || farmsLegacyBaseUrlLoading ? 'Loading FARMS scenario link' : 'FARMS URL, PIN, or program year is missing for this enrolment'}
          >
            <ExternalLink size={14} aria-hidden="true" />
            Open a Scenario in FARMS
          </button>
        )}
        {sharePointFolderUrl ? (
          <a
            className="calc-outline-btn calc-sharepoint-btn"
            href={sharePointFolderUrl}
            target="_blank"
            rel="noopener noreferrer"
          >
            <img src={sharepointIconUrl} className="calc-sharepoint-icon" alt="" aria-hidden="true" />
            Go to SharePoint
          </a>
        ) : (
          <button
            className="calc-outline-btn calc-sharepoint-btn"
            type="button"
            disabled
            title={loading ? 'Loading SharePoint folder link' : 'No SharePoint folder link found for this enrolment'}
          >
            <img src={sharepointIconUrl} className="calc-sharepoint-icon" alt="" aria-hidden="true" />
            Go to SharePoint
          </button>
        )}
      </div>

      {record && record.vsi_enrolmentstatus === 865520010 && (() => {
        const startDate = record.vsi_fortyfivedayletterstartdate;
        const paused = !!record.vsi_fortyfivedaycounterpaused;
        const pauseDate = record.vsi_fortyfivedaypausedate;
        const referenceMs = paused && pauseDate ? new Date(pauseDate).getTime() : Date.now();
        const elapsedDays = startDate
          ? Math.floor((referenceMs - new Date(startDate).getTime()) / (1000 * 60 * 60 * 24))
          : null;
        const remainingDays = elapsedDays !== null ? 45 - elapsedDays : null;
        return (
          <div className="calc-fortyfiveday-card" aria-label="45-day letter counter">
            <div className="calc-fortyfiveday-title">45-Day Counter</div>
            <div className="calc-fortyfiveday-grid">
              <div>
                <div className="calc-fortyfiveday-label">Start Date</div>
                <div className="calc-fortyfiveday-value">{startDate ? new Date(startDate).toLocaleDateString() : '-'}</div>
              </div>
              <div>
                <div className="calc-fortyfiveday-label">Elapsed</div>
                <div className="calc-fortyfiveday-value">{elapsedDays !== null ? `${elapsedDays} / 45 days` : '-'}</div>
              </div>
              <div>
                <div className="calc-fortyfiveday-label">Remaining</div>
                <div className={`calc-fortyfiveday-value${remainingDays !== null && remainingDays <= 10 && !paused ? ' calc-fortyfiveday-warning' : ''}`}>
                  {remainingDays !== null ? `${remainingDays} days` : '-'}
                </div>
              </div>
              <div>
                <div className="calc-fortyfiveday-label">Status</div>
                <div className="calc-fortyfiveday-value">
                  {paused
                    ? <span className="fortyfiveday-badge fortyfiveday-badge-paused">⏸ Paused{pauseDate ? ` since ${new Date(pauseDate).toLocaleDateString()}` : ''}</span>
                    : <span className="fortyfiveday-badge fortyfiveday-badge-running">▶ Running</span>}
                </div>
              </div>
            </div>
            <div className="calc-fortyfiveday-actions">
              {paused ? (
                <button
                  className="calc-outline-btn"
                  type="button"
                  onClick={() => void handle45DayResume()}
                  disabled={counterActionLoading}
                >
                  {counterActionLoading ? 'Resuming...' : 'Resume Counter'}
                </button>
              ) : (
                <button
                  className="calc-outline-btn"
                  type="button"
                  onClick={() => void handle45DayPause()}
                  disabled={counterActionLoading}
                >
                  {counterActionLoading ? 'Pausing...' : 'Pause Counter'}
                </button>
              )}
            </div>
            {counterActionError && <p className="calc-fortyfiveday-error">{counterActionError}</p>}
          </div>
        );
      })()}

      {loading && <p className="calc-state">Loading summary...</p>}
      {error && <p className="calc-state calc-state-error">Error loading summary: {error}</p>}
      {farmsWorkflowCalculationLoading && <p className="calc-state">Loading FARMS enrolment calculation...</p>}
      {farmsWorkflowCalculationError && <p className="calc-state calc-state-error">{farmsWorkflowCalculationError}</p>}

      {!loading && !error && record && (
        <div className={`calc-workspace${partnerPanelPinned ? ' calc-workspace-panel-pinned' : ''}`}>
          <div className="calc-legacy-workflow" aria-label="FARMS enrolment calculation">
            <CalculationErrorMessages messages={tableErrors.unmatched} />

            <section className="calc-legacy-panel" aria-label="Enrolment fee">
              <h2 className="calc-legacy-title">Enrolment Fee</h2>
              <CalculationErrorMessages messages={tableErrors.enrolmentFee} />
              <div className="calc-legacy-table-wrap">
                <table className="calc-legacy-table calc-legacy-table-compact">
                  <thead>
                    <tr>
                      <th scope="col">Contribution Margin</th>
                      <th scope="col">Fee</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>{formatCurrencyBlank(farmsEnrolment?.contributionMargin ?? record.vsi_contributionmargin)}</td>
                      <td>{formatCurrencyBlank(farmsEnrolment?.enrolmentFee ?? record.vsi_enrolmentfee)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>

          <section className="calc-legacy-panel" aria-label="Benefit margin calculation">
            <CalculationOption checked={!calculationTypeCode || calculationTypeCode === 'BENEFIT'} label="Calculate using Benefit Margins" />
            <h2 className="calc-benefit-title">Production Margins after Structural Change</h2>
            <CalculationErrorMessages messages={tableErrors.benefit} />

            <div className="calc-legacy-table-wrap">
              <table className="calc-legacy-table calc-benefit-table">
                <thead>
                  <tr>
                    <th scope="col"></th>
                    {benefitMarginRows.map(item => (
                      <th key={item.label} scope="col">{item.label}</th>
                    ))}
                    <th scope="col">Contribution Margin</th>
                    <th scope="col">Enrolment Fee</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <th scope="row">Margin</th>
                    {benefitMarginRows.map(item => (
                      <td key={item.label}>{formatCurrencyBlank(item.margin)}</td>
                    ))}
                    <td>{formatCurrencyBlank(farmsEnrolment?.benefitContributionMargin ?? record.vsi_contributionmargin)}</td>
                    <td>{formatCurrencyBlank(farmsEnrolment?.benefitEnrolmentFee ?? record.vsi_enrolmentfee)}</td>
                  </tr>
                  <tr>
                    <th scope="row">Used In Calculation</th>
                    {benefitMarginRows.map(item => (
                      <td key={item.label}>{getBooleanText(item.used)}</td>
                    ))}
                    <td></td>
                    <td></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>

          <section className="calc-legacy-panel" aria-label="Proxy margin calculation">
            <CalculationOption checked={calculationTypeCode === 'PROXY'} label="Calculate using Proxy Margins" />
            <h2 className="calc-benefit-title">Program Year Productive Value</h2>
            <CalculationErrorMessages messages={tableErrors.proxy} />
            <div className="calc-legacy-table-wrap">
              <table className="calc-legacy-table calc-proxy-table">
                <thead>
                  <tr>
                    <th scope="col">Code</th>
                    <th scope="col">Description</th>
                    <th scope="col">Productive Capacity</th>
                    {proxyYears.map(year => (
                      <Fragment key={year}>
                        <th scope="col">{year} BPU</th>
                        <th scope="col">{year} Margin</th>
                      </Fragment>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {(farmsEnrolment?.enwProductiveUnits?.length ? farmsEnrolment.enwProductiveUnits : [null]).map((unit, rowIndex) => (
                    <tr key={unit?.code ?? `blank-productive-unit-${rowIndex}`}>
                      <th scope="row">{unit?.code ?? ''}</th>
                      <td className="calc-legacy-text-cell">{unit?.description ?? ''}</td>
                      <td>{formatNumberOrBlank(unit?.productiveCapacity, 3)}</td>
                      {proxyYears.map((year, yearIndex) => {
                        const productiveValue = unit?.productiveValues?.[yearIndex];
                        return (
                          <Fragment key={`${unit?.code ?? rowIndex}-${year}`}>
                            <td>{formatCurrencyBlank(productiveValue?.bpuMargin)}</td>
                            <td>{formatCurrencyBlank(productiveValue?.productiveValue)}</td>
                          </Fragment>
                        );
                      })}
                    </tr>
                  ))}
                  <tr>
                    <th scope="row">Total</th>
                    <td></td>
                    <td></td>
                    {proxyYears.map((year, yearIndex) => (
                      <Fragment key={`${year}-total`}>
                        <td></td>
                        <td>{formatCurrencyBlank(farmsEnrolment?.proxyMargins?.[yearIndex])}</td>
                      </Fragment>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="calc-legacy-table-wrap">
              <table className="calc-legacy-table calc-legacy-table-compact">
                <thead>
                  <tr>
                    {proxyYears.map(year => (
                      <th key={year} scope="col">{year}</th>
                    ))}
                    <th scope="col">Contribution Margin</th>
                    <th scope="col">Enrolment Fee</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    {proxyYears.map((year, index) => (
                      <td key={year}>{formatCurrencyBlank(farmsEnrolment?.proxyMargins?.[index])}</td>
                    ))}
                    <td>{formatCurrencyBlank(farmsEnrolment?.proxyContributionMargin)}</td>
                    <td>{formatCurrencyBlank(farmsEnrolment?.proxyEnrolmentFee)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>

          <section className="calc-legacy-panel" aria-label="Manual margin calculation">
            <CalculationOption checked={calculationTypeCode === 'MANUAL'} label="Calculate using Manually Entered Margins" />
            <CalculationErrorMessages messages={tableErrors.manual} />
            <div className="calc-legacy-table-wrap">
              <table className="calc-legacy-table calc-manual-table">
                <thead>
                  <tr>
                    {proxyYears.map(year => (
                      <th key={year} scope="col">{year}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    {proxyYears.map((year, index) => (
                      <td key={year}>
                        <span className="calc-manual-input">{formatCurrencyBlank(manualMargins[index])}</span>
                      </td>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="calc-legacy-table-wrap">
              <table className="calc-legacy-table calc-legacy-table-compact">
                <thead>
                  <tr>
                    <th scope="col">Contribution Margin</th>
                    <th scope="col">Enrolment Fee</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td>{formatCurrencyBlank(farmsEnrolment?.manualContributionMargin)}</td>
                    <td>{formatCurrencyBlank(farmsEnrolment?.manualEnrolmentFee)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
          </div>

          <PartnerViewPanel
            rows={partnerRows}
            loading={partnerRowsLoading}
            error={partnerRowsError}
            farmsLegacyBaseUrl={farmsLegacyBaseUrl}
            farmsScenarioProgramYear={farmsScenarioProgramYear}
            routeSource={source}
            open={partnerPanelOpen}
            pinned={partnerPanelPinned}
            onToggleOpen={() => setPartnerPanelOpen(prev => !prev)}
            onTogglePinned={() => setPartnerPanelPinned(prev => !prev)}
          />
        </div>
      )}

      {showSupervisorModal && record && (
        <ReferToSupervisorModal
          selectedIds={new Set([record.vsi_participantprogramyearid])}
          rows={[record]}
          onClose={() => setShowSupervisorModal(false)}
          onComplete={() => {
            setRecord(prev => prev ? { ...prev, vsi_taskstatus: 865520001 } : prev);
            setRefreshKey(prev => prev + 1);
          }}
          onError={(message) => setError(message)}
        />
      )}
      {approvalErrorModal && (
        <ApprovalErrorModal message={approvalErrorModal} onClose={() => setApprovalErrorModal(null)} />
      )}
      {letterSentMessage && (
        <p className="calc-state" style={{ color: '#16a34a' }}>{letterSentMessage}</p>
      )}
      {show45DayModal && (
        <Send45DayLetterModal
          enrolmentId={enrolmentId ?? ''}
          enrolmentName={record?.vsi_name ?? ''}
          programYear={String(getProgramYear(record) ?? '')}
          onClose={() => setShow45DayModal(false)}
          onSuccess={() => setLetterSentMessage('45-day letter sent successfully.')}
        />
      )}
      {showApproveConfirm && record && (
        <ConfirmActionModal
          title="Confirm Approve Enrolments"
          message="Are you sure you want to approve the selected 1 enrolment?"
          enrolments={[{ id: record.vsi_participantprogramyearid, name: record.vsi_name ?? '' }]}
          confirmLabel="Approve"
          cancelLabel="Cancel"
          loading={approving}
          onConfirm={() => void handleApproveConfirm()}
          onCancel={() => setShowApproveConfirm(false)}
        />
      )}
      {showCompleteConfirm && record && (
        <ConfirmActionModal
          title="Confirm Complete Enrolment"
          message="This will set the task status to Ready and the enrolment status to Verified EN Calculated. Continue?"
          enrolments={[{ id: record.vsi_participantprogramyearid, name: record.vsi_name ?? '' }]}
          confirmLabel="Complete"
          cancelLabel="Cancel"
          loading={completing}
          onConfirm={() => void handleCompleteConfirm()}
          onCancel={() => setShowCompleteConfirm(false)}
        />
      )}
    </section>
  );
}
