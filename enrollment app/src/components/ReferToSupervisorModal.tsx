import { useState } from 'react';
import type { Vsi_participantprogramyears } from '../generated/models/Vsi_participantprogramyearsModel';
import { QueueitemsService } from '../generated/services/QueueitemsService';
import { Vsi_participantprogramyearsService } from '../generated/services/Vsi_participantprogramyearsService';
import { QueuesService } from '../generated/services/QueuesService';

const SUPERVISOR_QUEUE_NAME = 'Supervisor Approval Queue';

export function ReferToSupervisorModal({
  selectedIds,
  rows,
  onClose,
  onComplete,
  onError,
}: {
  selectedIds: Set<string>;
  rows: Vsi_participantprogramyears[];
  onClose: () => void;
  onComplete: (updatedIds: string[]) => void;
  onError?: (message: string) => void;
}) {

  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Utility to sanitize error messages (removes URLs and HelpLinks)
  function sanitizeError(msg: string | null): string | null {
    if (!msg) return null;
    // Remove URLs (http, https, www)
    let sanitized = msg.replace(/https?:\/\/\S+/gi, '').replace(/www\.[^\s]+/gi, '');
    // Remove Microsoft.PowerApps.CDS.HelpLink and similar patterns
    sanitized = sanitized.replace(/Microsoft\.PowerApps\.CDS\.HelpLink.*?(?=\s|$)/gi, '');
    // Remove any leftover HTML tags
    sanitized = sanitized.replace(/<[^>]+>/g, '');
    return sanitized.trim();
  }
  const selectedRows = rows.filter(r => selectedIds.has(r.vsi_participantprogramyearid));
  const noSelection = selectedRows.length === 0;
  const alreadySupervisorRows = selectedRows.filter(r => r.vsi_taskstatus === 865520001);

  const handleSubmit = async () => {
    setSubmitting(true);
    setError(null);
    try {
      // Look up the Supervisor Approval Queue
      const queuesResult = await QueuesService.getAll({
        filter: `name eq '${SUPERVISOR_QUEUE_NAME}'`,
        select: ['queueid', 'name'],
        maxPageSize: 1,
      });
      const queue = queuesResult.data?.[0];
      if (!queue) {
        throw new Error(`Queue "${SUPERVISOR_QUEUE_NAME}" not found`);
      }

      const rowsToProcess = selectedRows.filter(r => r.vsi_taskstatus !== 865520001);

      for (const row of rowsToProcess) {
        const enrolmentId = row.vsi_participantprogramyearid;

        // 1. Set task status to Supervisor and assign to AST Worker (owner)
        await Vsi_participantprogramyearsService.update(enrolmentId, {
          vsi_taskstatus: 865520001, // Vsi_participantprogramyearsvsi_taskstatus.Supervisor
        });

        // 2. Check for existing queueitem for this enrolment
        let existingQueueItem: { queueitemid: string } | null = null;
        try {
          const queueitemResult = await QueueitemsService.getAll({
            filter: `objectid_vsi_participantprogramyear/vsi_participantprogramyearid eq '${enrolmentId}' and statecode eq 0`,
            select: ['queueitemid', 'queueid'],
            maxPageSize: 1,
          });
          if (queueitemResult.data && queueitemResult.data.length > 0) {
            existingQueueItem = { queueitemid: queueitemResult.data[0].queueitemid };
          }
        } catch {
          // If lookup fails, continue to try to create
          existingQueueItem = null;
        }

        if (existingQueueItem) {
          // Update the queueid to Supervisor Approval Queue
          try {
            await QueueitemsService.update(existingQueueItem.queueitemid, {
              'QueueId@odata.bind': `/queues(${queue.queueid})`,
            });
          } catch (err) {
            setError('Failed to update existing queue item: ' + (err instanceof Error ? err.message : String(err)));
            setSubmitting(false);
            return;
          }
        } else {
          // Create a new queueitem
          let createResult;
          try {
            createResult = await QueueitemsService.create({
              'queueid@odata.bind': `/queues(${queue.queueid})`,
              'objectid_vsi_participantprogramyear@odata.bind': `/vsi_participantprogramyears(${enrolmentId})`,
            } as unknown as Parameters<typeof QueueitemsService.create>[0]);
            if (!createResult.success) {
              throw new Error(createResult.error?.message || 'Queue item creation failed');
            }
          } catch (err) {
            // Handle specific error for existing queue item gracefully
            const rawMsg = err instanceof Error ? err.message : String(err);
            if (rawMsg.startsWith('An active Queue Item already exists for the object')) {
              setError('This enrolment is already awaiting supervisor approval.');
            } else {
              setError('Failed to add to Supervisor Approval Queue: ' + rawMsg);
            }
            setSubmitting(false);
            return;
          }
        }
      }

      onComplete(rowsToProcess.map(r => r.vsi_participantprogramyearid));
      onClose();
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to refer to supervisor';
      setError(msg);
      onError?.(msg);
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Refer to Supervisor</h3>
          <button className="modal-close" onClick={onClose}>&times;</button>
        </div>
        <div className="modal-body">
          {noSelection ? (
            <div className="no-selection-message">No Enrolments Selected</div>
          ) : (
            <>
              <p>
                This will assign {selectedIds.size} enrolment{selectedIds.size !== 1 ? 's' : ''} to
                the <strong>Supervisor Approval Queue</strong> and set their task status
                to <strong>Supervisor</strong>.
              </p>
              {alreadySupervisorRows.length > 0 && (
                <p className="modal-warning">
                  {alreadySupervisorRows.length === selectedRows.length
                    ? 'All selected enrolments are already assigned to supervisor.'
                    : `${alreadySupervisorRows.length} of the selected enrolment${alreadySupervisorRows.length !== 1 ? 's are' : ' is'} already assigned to supervisor.`}
                </p>
              )}
              <div className="modal-selected-list">
                <table className="selected-enrolments-table">
                  <tbody>
                    {selectedRows.map((r, i) => {
                        const isSupervisor = r.vsi_taskstatus === 865520001;
                        return (
                          <tr key={r.vsi_participantprogramyearid} className={isSupervisor ? 'row-already-supervisor' : ''}>
                            <td className="selected-row-num">{i + 1}</td>
                            <td>{r.vsi_name ?? ''}</td>
                            <td>{isSupervisor && <span className="already-supervisor-badge">Already assigned</span>}</td>
                          </tr>
                        );
                      })}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </div>
        <div className="modal-footer">
          <button
            className="btn-ok"
            disabled={submitting || noSelection}
            onClick={handleSubmit}
          >
            {submitting ? 'Submitting...' : 'Confirm'}
          </button>
          <button className="btn-cancel" disabled={submitting} onClick={onClose}>Cancel</button>
          {error && <span className="modal-error">{sanitizeError(error)}</span>}
        </div>
      </div>
    </div>
  );
}
