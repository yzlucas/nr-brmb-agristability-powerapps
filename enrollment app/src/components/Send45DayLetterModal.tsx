import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Generate45DayLetterService } from '../generated/services/Generate45DayLetterService';
import { Vsi_programyearsService } from '../generated/services/Vsi_programyearsService';

type Props = {
  enrolmentId: string;
  enrolmentName: string;
  programYear: string;
  onClose: () => void;
  onSuccess: () => void;
};

function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
}

export function Send45DayLetterModal({ enrolmentId, enrolmentName, programYear, onClose, onSuccess }: Props) {
  const [letterDate, setLetterDate] = useState(todayIso());
  const [selectedYears, setSelectedYears] = useState<string[]>(programYear ? [programYear] : []);
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [yearOptions, setYearOptions] = useState<string[]>([]);
  const [yearsLoading, setYearsLoading] = useState(true);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [moreInfo, setMoreInfo] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    Vsi_programyearsService.getAll({ select: ['vsi_year'], filter: "statecode eq 0 and vsi_year ne '9999'", orderBy: ['vsi_year desc'] })
      .then(result => {
        const years = (result.data ?? [])
          .map(r => r.vsi_year)
          .filter((y): y is string => !!y);
        setYearOptions(years);
        if (programYear && !years.includes(programYear)) {
          setYearOptions(prev => [programYear, ...prev].sort((a, b) => Number(b) - Number(a)));
        }
      })
      .finally(() => setYearsLoading(false));
  }, []);

  const toggleYear = (y: string, checked: boolean) => {
    setSelectedYears(prev => checked ? [...prev, y] : prev.filter(v => v !== y));
  };
  const removeYear = (y: string) => setSelectedYears(prev => prev.filter(v => v !== y));
  const sortedSelected = [...selectedYears].sort((a, b) => Number(b) - Number(a));
  const unselected = yearOptions.filter(y => !selectedYears.includes(y));

  const handleSubmit = async () => {
    setSubmitting(true);
    setError(null);
    try {
      const result = await Generate45DayLetterService.Run({
        text: enrolmentId,
        date: letterDate,
        text_1: selectedYears.join(', '),
        text_2: moreInfo,
      });
      if (result.error) {
        const code = (result.error as { code?: number | string }).code;
        const inner = (result.error as { innerError?: { error?: { message?: string } } }).innerError?.error?.message;
        if (code === 502 || code === '502' || (result.error.message ?? '').toLowerCase().includes('badgateway')) {
          setError('The letter generation service did not respond (502 Bad Gateway). Please check the Power Automate flow run history for details.');
        } else {
          setError(inner ?? result.error.message ?? 'Failed to send 45-day letter.');
        }
      } else {
        onSuccess();
        onClose();
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to send 45-day letter.');
    } finally {
      setSubmitting(false);
    }
  };

  return createPortal(
    <div className="modal-overlay" onClick={submitting ? undefined : onClose}>
      <div className="modal-box" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h3>Send 45-Day Letter</h3>
            {enrolmentName && <p className="modal-subtitle">{enrolmentName}</p>}
          </div>
          <button type="button" className="modal-close" onClick={onClose} disabled={submitting}>&times;</button>
        </div>

        <div className="modal-body">
          <div className="modal-field">
            <label>Letter Date</label>
            <input
              type="date"
              value={letterDate}
              onChange={e => setLetterDate(e.target.value)}
              disabled={submitting}
            />
          </div>

          <div className="modal-field">
            <label>Program Year</label>
            <div className="modal-year-selector" ref={dropdownRef}>
              {/* Tags row */}
              <div className="modal-year-tags-row">
                {sortedSelected.map(y => (
                  <span key={y} className="modal-year-tag">
                    {y}
                    <button
                      type="button"
                      className="modal-year-tag-remove"
                      onClick={() => removeYear(y)}
                      disabled={submitting}
                      aria-label={`Remove ${y}`}
                    >&times;</button>
                  </span>
                ))}
                {unselected.length > 0 && (
                  <button
                    type="button"
                    className="modal-year-add-btn"
                    onClick={() => setDropdownOpen(o => !o)}
                    disabled={submitting || yearsLoading}
                  >
                    {yearsLoading ? 'Loading…' : '+ Add year'}
                  </button>
                )}
              </div>
              {/* Dropdown */}
              {dropdownOpen && (
                <div className="modal-year-dropdown">
                  {unselected.map(y => (
                    <button
                      key={y}
                      type="button"
                      className="modal-year-dropdown-item"
                      onClick={() => { toggleYear(y, true); setDropdownOpen(false); }}
                      disabled={submitting}
                    >
                      {y}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          <div className="modal-field">
            <label>More Info</label>
            <select
              value={moreInfo}
              onChange={e => setMoreInfo(e.target.value)}
              disabled={submitting}
              className="modal-text-input"
            >
              <option value="">-- Select one --</option>
              <option value="All Reference years">All Reference years</option>
              <option value="Statement A">Statement A</option>
              <option value="Production Information">Production Information</option>
            </select>
          </div>

          {error && <p className="modal-error">{error}</p>}
        </div>

        <div className="modal-footer">
          <button className="btn-ok" type="button" onClick={handleSubmit} disabled={submitting || yearsLoading || !letterDate || selectedYears.length === 0 || !moreInfo}>
            {submitting ? 'Sending…' : 'Send Letter'}
          </button>
          <button className="btn-cancel" type="button" onClick={onClose} disabled={submitting}>
            Cancel
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}
