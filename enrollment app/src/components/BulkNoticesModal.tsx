import { useState } from 'react';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';

import { GenerateBulkEnrolmentNoticesService } from '../generated/services/GenerateBulkEnrolmentNoticesService';

export function BulkNoticesModal({
  selectedIds,
  rows,
  onClose,
  onSuccess,
}: {
  selectedIds: Set<string>;
  rows: Vsi_participantprogramyears[];
  onClose: () => void;
  onSuccess?: (message: string) => void;
}) {
  const [bulkSentDate, setBulkSentDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [bulkFeeDate, setBulkFeeDate] = useState(() => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth() + 1, 0).toISOString().slice(0, 10);
  });
  const [bulkMergedPdf, setBulkMergedPdf] = useState(true);
  const [bulkSubmitting, setBulkSubmitting] = useState(false);
  const [bulkError, setBulkError] = useState<string | null>(null);
  const [bulkResult, setBulkResult] = useState<unknown>(null);

  const count = selectedIds.size;

  return (
    <div className="modal-overlay" onClick={bulkSubmitting ? undefined : onClose}>
      <div className="modal-box" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Generate Bulk Enrolment Notices</h3>
          {!bulkSubmitting && <button className="modal-close" onClick={onClose}>&times;</button>}
        </div>
        <div className="modal-body">
          {bulkSubmitting ? (
            <p className="bulk-submitting-label">Submitting {count} enrolment notice{count === 1 ? '' : 's'}…</p>
          ) : (
            <>
              <label className="modal-field">
                <span>Enrolment Sent Date:</span>
                <input type="date" value={bulkSentDate} onChange={e => setBulkSentDate(e.target.value)} />
              </label>
              <label className="modal-field">
                <span>Enrolment Fee Date:</span>
                <input type="date" value={bulkFeeDate} onChange={e => setBulkFeeDate(e.target.value)} />
              </label>
              <label className="modal-checkbox">
                <input type="checkbox" checked={bulkMergedPdf} onChange={e => setBulkMergedPdf(e.target.checked)} />
                Produce Merged PDF
              </label>
            </>
          )}
        </div>
        <div className="modal-footer">
          {!bulkSubmitting && !bulkResult && (
            <>
              <button
                className="btn-ok"
                disabled={selectedIds.size === 0}
                onClick={async () => {
                  setBulkSubmitting(true);
                  setBulkError(null);
                  try {
                    const allIds = Array.from(selectedIds);

                    const result = await GenerateBulkEnrolmentNoticesService.Run({
                      text: allIds.join(','),
                      date: bulkSentDate,
                      date_1: bulkFeeDate,
                      boolean: bulkMergedPdf,
                    });
                    if (!result.success) {
                      const msg = result.error instanceof Error
                        ? result.error.message
                        : (result.error as { message?: string } | undefined)?.message
                          ?? 'Workflow failed';
                      throw new Error(msg);
                    }

                    setBulkResult(true);
                    onSuccess?.(`${allIds.length} enrolment notice${allIds.length === 1 ? '' : 's'} submitted for generation. Processing will continue in the background.`);
                    onClose();
                  } catch (err) {
                    setBulkError(err instanceof Error ? err.message : 'Workflow failed');
                  } finally {
                    setBulkSubmitting(false);
                  }
                }}
              >
                OK
              </button>
              <button className="btn-cancel" onClick={onClose}>Cancel</button>
            </>
          )}
          {bulkError && <span className="modal-error">{bulkError}</span>}
          {bulkError && <button className="btn-cancel" onClick={onClose}>Close</button>}
        </div>
        {selectedIds.size > 0 && (
          <div className="modal-selected-list">
            <table className="selected-enrolments-table">
              <tbody>
                {rows
                  .filter(r => selectedIds.has(r.vsi_participantprogramyearid))
                  .map((r, i) => (
                    <tr key={r.vsi_participantprogramyearid}>
                      <td className="selected-row-num">{i + 1}</td>
                      <td>{r.vsi_name ?? ''}</td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
