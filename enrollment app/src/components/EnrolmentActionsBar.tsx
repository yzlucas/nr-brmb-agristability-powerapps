import { FileText, UserCheck, CircleCheck, UserPlus } from 'lucide-react';

type Props = {
  hasSelection: boolean;
  onOpenBulkNotices: () => void;
  onOpenAssign: () => void;
  onOpenReferToSupervisor: () => void;
  onOpenApproveCalculatedFees: () => void;
};

export function EnrolmentActionsBar({ hasSelection, onOpenBulkNotices, onOpenAssign, onOpenReferToSupervisor, onOpenApproveCalculatedFees }: Props) {
  return (
    <div className="enrolment-actions">
      <button className="dash-btn-secondary" onClick={onOpenBulkNotices} disabled={!hasSelection}>
        <FileText size={15} /> Bulk EN Notices
      </button>
      <button className="dash-btn-secondary" onClick={onOpenAssign} disabled={!hasSelection}>
        <UserPlus size={15} /> Assign
      </button>
      <button className="dash-btn-secondary" onClick={onOpenReferToSupervisor} disabled={!hasSelection}>
        <UserCheck size={15} /> Refer to Supervisor
      </button>
      <button className="dash-btn-primary" onClick={onOpenApproveCalculatedFees} disabled={!hasSelection}>
        <CircleCheck size={15} /> Approve Calculated Fees
      </button>
    </div>
  );
}

