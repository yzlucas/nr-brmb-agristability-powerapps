import { useEffect, useMemo, useState } from 'react';
import { CircleCheck, ExternalLink, RefreshCw, Send } from 'lucide-react';
import sharepointIconUrl from '/icons/sharepoint.svg?url';
import { Link, useParams } from 'react-router-dom';
import { ApprovalErrorModal } from '../components/ApprovalErrorModal';
import { Send45DayLetterModal } from '../components/Send45DayLetterModal';
import { ConfirmActionModal } from '../components/ConfirmActionModal';
import { ReferToSupervisorModal } from '../components/ReferToSupervisorModal';
import { patchEnrolmentCache } from '../hooks/useEnrolmentData';
import { removeSaItemsFromCache } from './SupervisorApprovalPage';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import { MicrosoftDataverseService } from '../generated/services/MicrosoftDataverseService';
import { QueueitemsService } from '../generated/services/QueueitemsService';
import { Vsi_armsconfigurationsService } from '../generated/services/Vsi_armsconfigurationsService';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { resolveCurrentSystemUser } from '../utils/currentUser';
import { calculateVariance, formatCurrencyOr, formatVariancePercent, getTaskStatusLabel } from '../utils/helpers';

const DATAVERSE_ORG_URL = 'https://aff-brmb-crm-dev.crm3.dynamics.com/';
const BENEFIT_MARGIN_COUNT = 5;
const APPROVABLE_STATUSES = new Set([865520005, 865520006]);

type CurrentUser = {
  systemUserId: string;
  displayName: string;
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
  return '-';
}

function getApprovalError(
  record: Vsi_participantprogramyears,
  currentUser: CurrentUser,
): string | null {
  const enrolmentName = record.vsi_name ?? 'This enrolment';

  if (!APPROVABLE_STATUSES.has(record.vsi_enrolmentstatus as unknown as number)) {
    return `${enrolmentName} cannot be approved because its enrolment status is not Verified EN Calculated or Unverified EN Calculated.`;
  }

  if (record.vsi_calculatedenfee == null) {
    return `${enrolmentName} cannot be approved because it does not have a calculated fee.`;
  }

  const raw = record as unknown as Record<string, unknown>;
  if (normalizeGuid(raw['_ownerid_value'] as string) !== normalizeGuid(currentUser.systemUserId)) {
    return `${enrolmentName} cannot be approved because you are not the owner of this enrolment.`;
  }

  return null;
}

export function EnrolmentCalculationPage() {
  const { enrolmentId, source } = useParams<{ enrolmentId: string; source: string }>();
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
  const [farmsApiTestMessage,] = useState('');
  const [show45DayModal, setShow45DayModal] = useState(false);
  const [letterSentMessage, setLetterSentMessage] = useState<string | null>(null);

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
            'vsi_calculatedenfee',
            'vsi_previousyearcalculatedenfee',
            'modifiedon',
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
        try {
        } catch {
        }

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

  const variance = useMemo(() => {
    return calculateVariance(record?.vsi_calculatedenfee, record?.vsi_previousyearcalculatedenfee);
  }, [record?.vsi_calculatedenfee, record?.vsi_previousyearcalculatedenfee]);

  const sharePointFolderUrl = record?.vsi_sharepointdocumentfolder;
  const programYear = useMemo(() => getProgramYear(record), [record]);
  const farmsScenarioUrl = useMemo(() => {
    if (!farmsLegacyBaseUrl || !participantPin || !programYear) return '';
    const params = new URLSearchParams({
      pin: participantPin,
      year: String(programYear),
      refresh: 'true',
    });
    return `${farmsLegacyBaseUrl}/farm800.do?${params.toString()}`;
  }, [farmsLegacyBaseUrl, participantPin, programYear]);
  const participantName = useMemo(() => {
    if (!record) return '';
    const raw = record as unknown as Record<string, unknown>;
    return (record.vsi_participantidname
      ?? raw['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue']
      ?? '') as string;
  }, [record]);
  const pin = participantPin || (participantPinLoading ? 'Loading...' : '-');
  const benefitMarginYears = useMemo(() => {
    return Array.from({ length: BENEFIT_MARGIN_COUNT }, (_, index) => (
      programYear ? String(programYear - (BENEFIT_MARGIN_COUNT + 1) + index) : `Year ${index + 1}`
    ));
  }, [programYear]);
  const benefitMargins = useMemo(() => {
    if (!record) return [];
    return Array.from({ length: BENEFIT_MARGIN_COUNT }, (_, index) => {
      const position = index + 1;
      const raw = record as unknown as Record<string, unknown>;
      return {
        label: benefitMarginYears[index],
        margin: raw[`vsi_programyearmargin${position}`] as number | undefined,
        used: raw[`vsi_programyearmargin${position}used`],
      };
    });
  }, [benefitMarginYears, record]);

  const resolveUser = async (): Promise<CurrentUser> => {
    if (currentUser) return currentUser;
    const resolved = await resolveCurrentSystemUser();
    const nextUser = {
      systemUserId: resolved.systemUserId,
      displayName: resolved.displayName,
    };
    setCurrentUser(nextUser);
    return nextUser;
  };

  const handleApproveClick = async () => {
    if (!record) return;

    try {
      const user = await resolveUser();
      const approvalError = getApprovalError(record, user);
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
      const user = await resolveUser();
      const approvalError = getApprovalError(record, user);
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
      <div className="calc-record-header">
        <div className="calc-record-main">
          <div className="calc-pin-line">
            <span>PIN:</span> {pin}
          </div>
          <h1 className="calc-participant-name">{participantName || (loading ? 'Loading...' : '-')}</h1>
          <div className="calc-primary-actions">
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
            <button
              className="calc-outline-btn"
              type="button"
              onClick={() => void handleApproveClick()}
              disabled={!record || approving}
            >
              <CircleCheck size={14} aria-hidden="true" />
              {approving ? 'Approving...' : 'Approve'}
            </button>
            {/* <button
              className="calc-outline-btn"
              type="button"
              onClick={() => void handleFarmsApiTestClick()}
              disabled={farmsApiTestLoading}
            >
              <ExternalLink size={14} aria-hidden="true" />
              {farmsApiTestLoading ? 'Calling FARMS...' : 'Test FARMS Line Items'}
            </button> */}
            {/* <button
              className="calc-outline-btn"
              type="button"
              onClick={() => void handleFarmsWorkflowTestClick()}
              disabled={farmsApiTestLoading}
            >
              <ExternalLink size={14} aria-hidden="true" />
              Test FARMS Workflow
            </button> */}
            {farmsApiTestMessage && <span className="calc-inline-status">{farmsApiTestMessage}</span>}
          </div>
        </div>

        <div className="calc-sharepoint-action">
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
      </div>

      {loading && <p className="calc-state">Loading summary...</p>}
      {error && <p className="calc-state calc-state-error">Error loading summary: {error}</p>}

      {!loading && !error && record && (
        <>
          <div className="calc-benefit-section" aria-label="Benefit margin calculation">
            <div className="calc-benefit-option">
              <span className="calc-benefit-radio" aria-hidden="true" />
              <span>Calculate using Benefit Margins</span>
            </div>

            <h2 className="calc-benefit-title">Production Margins after Structural Change</h2>

            <div className="calc-benefit-table-wrap">
              <table className="calc-benefit-table">
                <thead>
                  <tr>
                    <th scope="col"></th>
                    {benefitMargins.map(item => (
                      <th key={item.label} scope="col">{item.label}</th>
                    ))}
                    <th scope="col">Contribution Margin</th>
                    <th scope="col">Enrolment Fee</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <th scope="row">Margin</th>
                    {benefitMargins.map(item => (
                      <td key={item.label}>{formatCurrencyOr(item.margin, '-')}</td>
                    ))}
                    <td>{formatCurrencyOr(record.vsi_contributionmargin, '-')}</td>
                    <td>{formatCurrencyOr(record.vsi_enrolmentfee, '-')}</td>
                  </tr>
                  <tr>
                    <th scope="row">Used In Calculation</th>
                    {benefitMargins.map(item => (
                      <td key={item.label}>{getBooleanText(item.used)}</td>
                    ))}
                    <td></td>
                    <td></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <div className="calc-summary-card" aria-label="Enrolment summary card">
            <h2 className="calc-summary-title">Summary</h2>
            <div className="calc-summary-grid">
              <div>
                <div className="calc-label">Enrolment Name</div>
                <div className="calc-value">{record.vsi_name || '-'}</div>
              </div>
              <div>
                <div className="calc-label">Task Status</div>
                <div className="calc-value">{getTaskStatusLabel(record.vsi_taskstatus) || '-'}</div>
              </div>
              <div>
                <div className="calc-label">Calculated Fee</div>
                <div className="calc-value">{formatCurrencyOr(record.vsi_calculatedenfee, '-')}</div>
              </div>
              <div>
                <div className="calc-label">Variance</div>
                <div className="calc-value">{formatVariancePercent(variance) || '-'}</div>
              </div>
            </div>
          </div>
        </>
      )}

      <div className="calc-links">
        <Link to={backTo}>{backLabel}</Link>
      </div>

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
    </section>
  );
}
