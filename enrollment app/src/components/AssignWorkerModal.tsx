import { useState } from 'react';
import { QueueitemsService } from '../generated/services/QueueitemsService';
import { useAssignableUsers } from '../hooks/useAssignableUsers';
import type { AssignableUser } from '../hooks/useAssignableUsers';
import { getInitials } from '../utils/helpers';

export type { AssignableUser };

export function AssignWorkerModal({
  enrolmentName,
  queueitemId,
  queueId,
  onClose,
  onAssigned,
}: {
  enrolmentName: string;
  queueitemId: string | undefined;
  queueId?: string;
  queueName?: string;
  onClose: () => void;
  onAssigned: (workerId: string, workerName: string) => void;
}) {
  const { results, searchTerm, setSearchTerm, searching, loadError, hasLoaded } =
    useAssignableUsers(queueId, !!queueitemId);

  const [selected, setSelected] = useState<AssignableUser | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);

  const handleAssign = async () => {
    if (!selected || !queueitemId) return;
    setSubmitting(true);
    setSubmitError(null);
    try {
      await QueueitemsService.delete(queueitemId);
      onAssigned(selected.systemUserId, selected.displayName);
      onClose();
    } catch (err) {
      setSubmitError(err instanceof Error ? err.message : 'Assignment failed');
    } finally {
      setSubmitting(false);
    }
  };

  const error = submitError ?? loadError;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box assign-modal-box" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Assign Queue Item</h3>
          <button type="button" className="modal-close" onClick={onClose}>&times;</button>
        </div>

        <div className="modal-body">
          <p className="assign-modal-subtitle">
            Assigning: <strong>{enrolmentName}</strong>
          </p>

          {!queueitemId && (
            <p className="assign-no-queueitem">
              This enrolment has no active queue item. Refer it to the supervisor queue first.
            </p>
          )}

          {queueitemId && (
            <>
              <div className="assign-search-row">
                <input
                  className="assign-search-input"
                  type="text"
                  placeholder="Filter by name or email…"
                  value={searchTerm}
                  onChange={e => setSearchTerm(e.target.value)}
                  autoFocus
                />
              </div>

              {hasLoaded && !searching && results.length === 0 && (
                <p className="assign-no-results">No matching users.</p>
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
          {queueitemId && (
            <button
              type="button"
              className="btn-ok"
              disabled={!selected || submitting || searching}
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
