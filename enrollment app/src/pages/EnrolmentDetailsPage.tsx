import { useEffect, useMemo, useState, type ChangeEvent } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { Calculator, User } from 'lucide-react';
import sharepointIconUrl from '/icons/sharepoint.svg?url';
import {
  Vsi_participantprogramyearsvsi_enrolmentstatus,
  type Vsi_participantprogramyears,
  type Vsi_participantprogramyearsBase,
  type Vsi_participantprogramyearsvsi_enrolmentstatus as EnrolmentStatusValue,
} from '../generated/models/Vsi_participantprogramyearsModel';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { formatEnrolmentStatusDisplay, getTaskStatusLabel } from '../utils/helpers';

type DateField =
  | 'vsi_enrolmentnoticesentdate'
  | 'vsi_programyearoptoutdate'
  | 'vsi_lateenrolmentnoticesentdate'
  | 'vsi_enrolmentfeespaiddate'
  | 'vsi_enrolmentfeesnonpenaltyduedate'
  | 'vsi_enrolmentfeesfinaldeadlinedate';

type DetailFormState = {
  enrolmentStatus: EnrolmentStatusValue;
  vsi_fullyprovinciallyfunded: boolean;
  vsi_enrolmentnoticesentdate: string;
  vsi_programyearoptoutdate: string;
  vsi_lateenrolmentnoticesentdate: string;
  vsi_enrolmentfeespaiddate: string;
  vsi_enrolmentfeesnonpenaltyduedate: string;
  vsi_enrolmentfeesfinaldeadlinedate: string;
};

const DATE_FIELDS: DateField[] = [
  'vsi_enrolmentnoticesentdate',
  'vsi_programyearoptoutdate',
  'vsi_lateenrolmentnoticesentdate',
  'vsi_enrolmentfeespaiddate',
  'vsi_enrolmentfeesnonpenaltyduedate',
  'vsi_enrolmentfeesfinaldeadlinedate',
];

const DETAIL_SELECT = [
  'vsi_name',
  '_vsi_participantid_value',
  '_vsi_programyearid_value',
  'vsi_sharepointdocumentfolder',
  'vsi_enrolmentstatus',
  'vsi_taskstatus',
  'owneridname',
  'vsi_totalfeesowedcalculated',
  'vsi_totalfeespaid',
  'vsi_enrolmentnoticesentdate',
  'vsi_programyearoptoutdate',
  'vsi_lateenrolmentnoticesentdate',
  'vsi_manualreview',
  'vsi_fullyprovinciallyfunded',
  'vsi_enrolmentfee',
  'vsi_enrolmentfeespaiddate',
  'vsi_enrolmentfeesnonpenaltyduedate',
  'vsi_enrolmentfeesfinaldeadlinedate',
  'vsi_administrativecostsharingfee',
  'vsi_latepaymentfee',
  'vsi_adjustedlateenrolmentfee',
  '_vsi_feemodifiedby_value',
  'vsi_fortyfivedayletterstartdate',
  'vsi_fortyfivedaycounterpaused',
  'vsi_fortyfivedaypausedate',
  'vsi_isnewparticipant',
] as const;

const formatCad = (value: number | undefined): string => {
  if (value == null || Number.isNaN(Number(value))) return '---';
  return `CA$${Number(value).toLocaleString('en-CA', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
};

const yesNoText = (value: unknown): string => {
  if (value === true || value === 1 || value === '1') return 'Yes';
  if (value === false || value === 0 || value === '0') return 'No';
  return '---';
};

const toDateInputValue = (value: string | undefined): string => {
  if (!value) return '';
  const directMatch = value.match(/^(\d{4}-\d{2}-\d{2})/);
  if (directMatch) return directMatch[1];

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  const yyyy = parsed.getFullYear();
  const mm = String(parsed.getMonth() + 1).padStart(2, '0');
  const dd = String(parsed.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
};

const initialFormFromRecord = (record: Vsi_participantprogramyears): DetailFormState => ({
  enrolmentStatus: record.vsi_enrolmentstatus,
  vsi_fullyprovinciallyfunded: Boolean(record.vsi_fullyprovinciallyfunded),
  vsi_enrolmentnoticesentdate: toDateInputValue(record.vsi_enrolmentnoticesentdate),
  vsi_programyearoptoutdate: toDateInputValue(record.vsi_programyearoptoutdate),
  vsi_lateenrolmentnoticesentdate: toDateInputValue(record.vsi_lateenrolmentnoticesentdate),
  vsi_enrolmentfeespaiddate: toDateInputValue(record.vsi_enrolmentfeespaiddate),
  vsi_enrolmentfeesnonpenaltyduedate: toDateInputValue(record.vsi_enrolmentfeesnonpenaltyduedate),
  vsi_enrolmentfeesfinaldeadlinedate: toDateInputValue(record.vsi_enrolmentfeesfinaldeadlinedate),
});

const getFormattedLookup = (record: Vsi_participantprogramyears, key: string): string => {
  const raw = record as unknown as Record<string, unknown>;
  const value = raw[key];
  if (typeof value === 'string' && value.trim().length > 0) return value;
  return '';
};


export function EnrolmentDetailsPage() {
  // Read both source and enrolmentId from params
  const { source = 'dashboard', enrolmentId } = useParams<{ source?: string; enrolmentId: string }>();
  const navigate = useNavigate();

  const [record, setRecord] = useState<Vsi_participantprogramyears | null>(null);
  const [formState, setFormState] = useState<DetailFormState | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [saveNotice, setSaveNotice] = useState<string | null>(null);

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
          select: [...DETAIL_SELECT],
        });

        // Some environments are strict about select fields; retry without select to avoid hard-fail.
        if (!result?.data) {
          result = await Vsi_participantprogramyearsService.get(enrolmentId);
        }
        if (cancelled) return;
        const loaded = result.data;
        if (!loaded) {
          setError('Unable to load enrolment details.');
          return;
        }
        setRecord(loaded);
        setFormState(initialFormFromRecord(loaded));
      } catch (e: unknown) {
        if (!cancelled) {
          const message = e instanceof Error ? e.message : 'Unable to load enrolment details.';
          setError(`Unable to load enrolment details. ${message}`);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [enrolmentId]);

  const statusOptions = useMemo(
    () => Object.entries(Vsi_participantprogramyearsvsi_enrolmentstatus).map(([value, label]) => ({
      value: Number(value) as EnrolmentStatusValue,
      label: formatEnrolmentStatusDisplay(label),
    })),
    [],
  );

  const baseline = useMemo(() => (record ? initialFormFromRecord(record) : null), [record]);

  const hasChanges = useMemo(() => {
    if (!baseline || !formState) return false;
    return (
      baseline.enrolmentStatus !== formState.enrolmentStatus
      || baseline.vsi_fullyprovinciallyfunded !== formState.vsi_fullyprovinciallyfunded
      || DATE_FIELDS.some(field => baseline[field] !== formState[field])
    );
  }, [baseline, formState]);

  const participantName = useMemo(() => {
    if (!record) return '---';
    const label = record.vsi_participantidname
      ?? getFormattedLookup(record, '_vsi_participantid_value@OData.Community.Display.V1.FormattedValue');
    return label || '---';
  }, [record]);

  const programYear = useMemo(() => {
    if (!record) return '---';
    const label = record.vsi_programyearidname
      ?? getFormattedLookup(record, '_vsi_programyearid_value@OData.Community.Display.V1.FormattedValue');
    return label || '---';
  }, [record]);

  const feeModifiedBy = useMemo(() => {
    if (!record) return '---';
    const label = record.vsi_feemodifiedbyname
      ?? getFormattedLookup(record, '_vsi_feemodifiedby_value@OData.Community.Display.V1.FormattedValue');
    return label || '---';
  }, [record]);

  const updateDateField = (field: DateField) => (event: ChangeEvent<HTMLInputElement>) => {
    const { value } = event.target;
    setSaveNotice(null);
    setFormState(prev => (prev ? { ...prev, [field]: value } : prev));
  };

  const onStatusChange = (event: ChangeEvent<HTMLSelectElement>) => {
    const nextValue = Number(event.target.value) as EnrolmentStatusValue;
    setSaveNotice(null);
    setFormState(prev => (prev ? { ...prev, enrolmentStatus: nextValue } : prev));
  };

  const onLateParticipantChange = (event: ChangeEvent<HTMLInputElement>) => {
    const checked = event.target.checked;
    setSaveNotice(null);
    setFormState(prev => (prev ? { ...prev, vsi_fullyprovinciallyfunded: checked } : prev));
  };

  const handleSave = async () => {
    if (!record || !formState) return;

    const changedFields: Partial<Omit<Vsi_participantprogramyearsBase, 'vsi_participantprogramyearid'>> = {};

    if (formState.enrolmentStatus !== record.vsi_enrolmentstatus) {
      changedFields.vsi_enrolmentstatus = formState.enrolmentStatus;

      // Clear 45-day fields when moving away from the 45 Day Letter status
      if (record.vsi_enrolmentstatus === 865520010 && formState.enrolmentStatus !== 865520010) {
        changedFields.vsi_fortyfivedayletterstartdate = null as unknown as string;
        changedFields.vsi_fortyfivedaylettersent = null as unknown as string;
        changedFields.vsi_fortyfivedaycounterpaused = null as unknown as boolean;
        changedFields.vsi_fortyfivedaypausedate = null as unknown as string;
      }

      // Set start date immediately when changing TO 45 Day Letter so the counter shows right away
      if (formState.enrolmentStatus === 865520010 && !record.vsi_fortyfivedayletterstartdate) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        changedFields.vsi_fortyfivedayletterstartdate = today.toISOString();
      }
    }
    if (formState.vsi_fullyprovinciallyfunded !== Boolean(record.vsi_fullyprovinciallyfunded)) {
      changedFields.vsi_fullyprovinciallyfunded = formState.vsi_fullyprovinciallyfunded;
    }

    for (const field of DATE_FIELDS) {
      const existingValue = toDateInputValue(record[field]);
      const nextValue = formState[field];
      if (existingValue !== nextValue && nextValue) {
        changedFields[field] = nextValue;
      }
    }

    if (Object.keys(changedFields).length === 0) {
      setSaveNotice('No changes to save.');
      return;
    }

    try {
      setSaving(true);
      setError(null);
      setSaveNotice(null);
      await Vsi_participantprogramyearsService.update(record.vsi_participantprogramyearid, changedFields);
      let refreshed = await Vsi_participantprogramyearsService.get(record.vsi_participantprogramyearid, {
        select: [...DETAIL_SELECT],
      });
      if (!refreshed?.data) {
        refreshed = await Vsi_participantprogramyearsService.get(record.vsi_participantprogramyearid);
      }
      const updated = refreshed.data ?? record;
      setRecord(updated);
      setFormState(initialFormFromRecord(updated));
      setSaveNotice('Changes saved.');
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Unable to save changes.');
    } finally {
      setSaving(false);
    }
  };


  if (loading) {
    return <section className="details-wrapper"><p className="enrolment-loading">Loading details...</p></section>;
  }

  // Determine back link and label
  let backPath = '/dashboard-home';
  let backLabel = 'Back to Dashboard';
  if (source === 'supervisor') {
    backPath = '/supervisor-approval';
    backLabel = 'Back to Supervisor Approval';
  }

  if (error || !record || !formState) {
    return (
      <section className="details-wrapper">
        <p className="enrolment-error">{error ?? 'Enrolment record not found.'}</p>
        <button type="button" className="details-back-btn" onClick={() => navigate(backPath)}>{backLabel}</button>
      </section>
    );
  }


  return (
    <section className="details-wrapper">
      <div className="details-title-row">
        <button type="button" className="details-back-btn" onClick={() => navigate(backPath)}>{backLabel}</button>
        <h1 className="details-page-title">Enrolment App / Deadlines &amp; Fees</h1>
        <div className="details-meta-strip">
          <span className="details-meta-owner">
            <User size={13} aria-hidden="true" />
            {record.owneridname || getFormattedLookup(record, '_ownerid_value@OData.Community.Display.V1.FormattedValue') || '—'}
          </span>
          {getTaskStatusLabel(record.vsi_taskstatus) && (
            <span className={`details-task-badge details-task-badge--${(getTaskStatusLabel(record.vsi_taskstatus) ?? '').toLowerCase()}`}>
              {getTaskStatusLabel(record.vsi_taskstatus)}
            </span>
          )}
        </div>
      </div>

      <div className="details-composite">
        <div className="details-header-band">
          <div className="details-header-grid">
            <div className="details-field">
              <span className="details-label">Participant:</span>
              <strong className="details-value-strong">{participantName}</strong>
            </div>
            <div className="details-field">
              <span className="details-label">Program Year <span className="required-mark">*</span></span>
              <strong className="details-value-strong">{programYear}</strong>
            </div>
            <div className="details-field">
              <span className="details-label">Late Participant</span>
              <label htmlFor="late-participant" className="details-checkbox-control">
                <input
                  id="late-participant"
                  type="checkbox"
                  checked={formState.vsi_fullyprovinciallyfunded}
                  onChange={onLateParticipantChange}
                  disabled={saving}
                />
              </label>
            </div>
            <div className="details-field">
              <span className="details-label">NPP</span>
              <label className="details-checkbox-control">
                <input
                  type="checkbox"
                  checked={Boolean(record.vsi_isnewparticipant)}
                  readOnly
                  tabIndex={-1}
                  style={{ pointerEvents: 'none' }}
                />
              </label>
            </div>
            <div className="details-field details-link-field">
              {record.vsi_sharepointdocumentfolder ? (
                <a
                  className="calc-outline-btn calc-sharepoint-btn"
                  href={record.vsi_sharepointdocumentfolder}
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
                  title="No SharePoint folder link found for this enrolment"
                >
                  <img src={sharepointIconUrl} className="calc-sharepoint-icon" alt="" aria-hidden="true" />
                  Go to SharePoint
                </button>
              )}
              <button
                type="button"
                className="calc-outline-btn"
                onClick={() => navigate(`/calculation/${source}/${enrolmentId}`)}
              >
                <Calculator size={15} /> Go to Calculation
              </button>
            </div>
            <div className="details-field details-link-field details-fortyfiveday-cell">
              {record.vsi_enrolmentstatus === 865520010 && (() => {
                const startDate = record.vsi_fortyfivedayletterstartdate as string | undefined;
                const paused = !!(record as unknown as Record<string, unknown>)['vsi_fortyfivedaycounterpaused'];
                const pauseDate = (record as unknown as Record<string, unknown>)['vsi_fortyfivedaypausedate'] as string | undefined;
                const referenceMs = paused && pauseDate ? new Date(pauseDate).getTime() : Date.now();
                const elapsedDays = startDate
                  ? Math.floor((referenceMs - new Date(startDate).getTime()) / (1000 * 60 * 60 * 24))
                  : null;
                const remainingDays = elapsedDays !== null ? 45 - elapsedDays : null;
                return (
                  <div className="calc-fortyfiveday-card details-fortyfiveday-card" aria-label="45-day letter counter">
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
                  </div>
                );
              })()}
            </div>
          </div>
        </div>

        <div className="details-content-section details-content-main">
          <div className="details-main-grid">
            <div className="details-field">
              <label htmlFor="enrolment-status" className="details-label">Enrolment Status <span className="required-mark">*</span></label>
              <select
                id="enrolment-status"
                value={formState.enrolmentStatus}
                onChange={onStatusChange}
                className="details-select"
                disabled={saving}
              >
                {statusOptions.map(option => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </div>

            <div className="details-field">
              <span className="details-label">Total Fees Owed</span>
              <strong className="details-money">{formatCad(record.vsi_totalfeesowedcalculated)}</strong>
            </div>

            <div className="details-field">
              <span className="details-label">Total Fees Paid</span>
              <strong className="details-money">{formatCad(record.vsi_totalfeespaid)}</strong>
            </div>

            <div className="details-field">
              <label htmlFor="enrol-notice-date" className="details-label">Enrolment Notice Sent Date</label>
              <input
                id="enrol-notice-date"
                type="date"
                className="details-date"
                value={formState.vsi_enrolmentnoticesentdate}
                onChange={updateDateField('vsi_enrolmentnoticesentdate')}
                disabled={saving}
              />
            </div>

            <div className="details-field">
              <label htmlFor="opt-out-date" className="details-label">Program Year Opt-Out Date</label>
              <input
                id="opt-out-date"
                type="date"
                className="details-date"
                value={formState.vsi_programyearoptoutdate}
                onChange={updateDateField('vsi_programyearoptoutdate')}
                disabled={saving}
              />
            </div>

            <div className="details-field">
              <label htmlFor="late-notice-date" className="details-label">Late Enrolment Notice Sent Date</label>
              <input
                id="late-notice-date"
                type="date"
                className="details-date"
                value={formState.vsi_lateenrolmentnoticesentdate}
                onChange={updateDateField('vsi_lateenrolmentnoticesentdate')}
                disabled={saving}
              />
            </div>

            <div className="details-field">
              <span className="details-label">Manual Review</span>
              <strong className="details-value-strong">{yesNoText(record.vsi_manualreview)}</strong>
            </div>
          </div>
        </div>

        <div className="details-section-break" />

        <div className="details-content-section details-content-fees">
          <h2 className="details-section-title">Enrolment Deadlines and Fees</h2>

          <div className="details-fees-grid">
            <div className="details-field">
              <span className="details-label">Enrolment Fee</span>
              <strong className="details-money">{formatCad(record.vsi_enrolmentfee)}</strong>
            </div>

            <div className="details-field">
              <label htmlFor="enrol-fees-paid-date" className="details-label">Enrolment Fees Paid Date</label>
              <input
                id="enrol-fees-paid-date"
                type="date"
                className="details-date"
                value={formState.vsi_enrolmentfeespaiddate}
                onChange={updateDateField('vsi_enrolmentfeespaiddate')}
                disabled={saving}
              />
            </div>

            <div className="details-field">
              <label htmlFor="non-penalty-date" className="details-label">Enrolment-fees non-penalty due date</label>
              <input
                id="non-penalty-date"
                type="date"
                className="details-date"
                value={formState.vsi_enrolmentfeesnonpenaltyduedate}
                onChange={updateDateField('vsi_enrolmentfeesnonpenaltyduedate')}
                disabled={saving}
              />
            </div>

            <div className="details-field">
              <label htmlFor="final-deadline-date" className="details-label">Enrolment fees final deadline date</label>
              <input
                id="final-deadline-date"
                type="date"
                className="details-date"
                value={formState.vsi_enrolmentfeesfinaldeadlinedate}
                onChange={updateDateField('vsi_enrolmentfeesfinaldeadlinedate')}
                disabled={saving}
              />
            </div>

            <div className="details-field">
              <span className="details-label">Administrative cost Sharing fee</span>
              <strong className="details-money details-money-alert">{formatCad(record.vsi_administrativecostsharingfee)}</strong>
            </div>

            <div className="details-field">
              <span className="details-label">Late payment fee</span>
              <strong className="details-money">{formatCad(record.vsi_latepaymentfee)}</strong>
            </div>

            <div className="details-field">
              <span className="details-label">Adjusted late enrolment fee</span>
              <strong className="details-money">{formatCad(record.vsi_adjustedlateenrolmentfee)}</strong>
            </div>
          </div>

          <div className="details-field details-fee-modified">
            <span className="details-label">Fee modified by:</span>
            <strong className="details-value-strong">{feeModifiedBy}</strong>
          </div>
        </div>
      </div>

      <div className="details-actions">
        {saveNotice ? <span className="details-save-notice">{saveNotice}</span> : null}
        <button type="button" className="details-save-btn" onClick={handleSave} disabled={saving || !hasChanges}>
          {saving ? 'Saving...' : 'Save Changes'}
        </button>
      </div>
    </section>
  );
}

