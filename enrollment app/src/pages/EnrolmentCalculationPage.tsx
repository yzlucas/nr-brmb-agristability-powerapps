import { useEffect, useMemo, useState } from 'react';
import { Send } from 'lucide-react';
import sharepointIconUrl from '/icons/sharepoint.svg?url';
import { Link, useParams } from 'react-router-dom';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import { MicrosoftDataverseService } from '../generated/services/MicrosoftDataverseService';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { calculateVariance, formatCurrencyOr, formatVariancePercent, getTaskStatusLabel } from '../utils/helpers';

const DATAVERSE_ORG_URL = 'https://aff-brmb-crm-dev.crm3.dynamics.com/';
const BENEFIT_MARGIN_COUNT = 5;

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

export function EnrolmentCalculationPage() {
  const { enrolmentId, source } = useParams<{ enrolmentId: string; source: string }>();
  const backTo = source === 'supervisor' ? '/supervisor-approval' : '/dashboard-home';
  const backLabel = source === 'supervisor' ? 'Back to Supervisor Approval' : 'Back to Dashboard';
  const [record, setRecord] = useState<Vsi_participantprogramyears | null>(null);
  const [participantPin, setParticipantPin] = useState('');
  const [participantPinLoading, setParticipantPinLoading] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

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
  }, [enrolmentId]);

  const variance = useMemo(() => {
    return calculateVariance(record?.vsi_calculatedenfee, record?.vsi_previousyearcalculatedenfee);
  }, [record?.vsi_calculatedenfee, record?.vsi_previousyearcalculatedenfee]);

  const sharePointFolderUrl = record?.vsi_sharepointdocumentfolder;
  const participantName = useMemo(() => {
    if (!record) return '';
    const raw = record as unknown as Record<string, unknown>;
    return (record.vsi_participantidname
      ?? raw['_vsi_participantid_value@OData.Community.Display.V1.FormattedValue']
      ?? '') as string;
  }, [record]);
  const pin = participantPin || (participantPinLoading ? 'Loading...' : '-');
  const benefitMarginYears = useMemo(() => {
    const programYear = getProgramYear(record);
    return Array.from({ length: BENEFIT_MARGIN_COUNT }, (_, index) => (
      programYear ? String(programYear - (BENEFIT_MARGIN_COUNT + 1) + index) : `Year ${index + 1}`
    ));
  }, [record]);
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

  return (
    <section className="page-card calc-page">
      <div className="calc-record-header">
        <div className="calc-record-main">
          <div className="calc-pin-line">
            <span>PIN:</span> {pin}
          </div>
          <h1 className="calc-participant-name">{participantName || (loading ? 'Loading...' : '-')}</h1>
          <div className="calc-primary-actions">
            <button className="calc-outline-btn" type="button">Send 45-Day Letter</button>
            <button className="calc-outline-btn" type="button">
              <Send size={14} aria-hidden="true" />
              Refer to Supervisor
            </button>
          </div>
        </div>

        <div className="calc-sharepoint-action">
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
    </section>
  );
}
