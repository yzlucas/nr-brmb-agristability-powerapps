import type { QuickFilterState } from '../types/enrollment';

type Props = {
  filters: QuickFilterState;
  onToggleFilter: (key: keyof QuickFilterState) => void;
  activeAdvancedCount: number;
};

export function EnrolmentQuickFilters({
  filters,
  onToggleFilter,
  activeAdvancedCount,
}: Props) {
  return (
    <div className="enrolment-filters">
      <strong>Apply Filters</strong>
      <label><input type="checkbox" checked={filters.verifiedCalc} onChange={() => onToggleFilter('verifiedCalc')} /> Verified, EN Calculated</label>
      <label><input type="checkbox" checked={filters.unverifiedCalc} onChange={() => onToggleFilter('unverifiedCalc')} /> Unverified, EN Calculated</label>
      <label><input type="checkbox" checked={filters.flagged} onChange={() => onToggleFilter('flagged')} /> Flagged files</label>
      <label><input type="checkbox" checked={filters.partnerships} onChange={() => onToggleFilter('partnerships')} /> Partnerships/Combined</label>
      <label><input type="checkbox" checked={filters.fortyFiveDayLetter} onChange={() => onToggleFilter('fortyFiveDayLetter')} /> 45 day Letter</label>
      {activeAdvancedCount > 0 && (
        <span className="ef-active-count">{activeAdvancedCount} advanced filter(s)</span>
      )}
    </div>
  );
}

