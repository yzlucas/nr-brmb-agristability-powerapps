import type { ColumnDef, SortKey, ViewPayload } from '../types/enrollment';

export const ALL_COLUMNS: ColumnDef[] = [
  { key: 'flagged', label: 'Flagged', icon: 'check', removable: true },
  { key: 'pin', label: 'Enrolment Name', icon: 'text', removable: false },
  { key: 'producer', label: 'Participant', icon: 'link', removable: true },
  { key: 'year', label: 'Year', icon: 'link', removable: true },
  { key: 'enrolStatus', label: 'Enrolment Status', icon: 'list', removable: true },
  { key: 'taskStatus', label: 'Task Status', icon: 'list', removable: true },
  { key: 'fee', label: 'Calculated fee', icon: 'number', removable: true },
  { key: 'totalFeesOwed', label: 'Total Fees Owed', icon: 'number', removable: true },
  { key: 'totalFeesPaid', label: 'Total Fees Paid', icon: 'number', removable: true },
  { key: 'enrolmentFee', label: 'Enrolment Fee', icon: 'number', removable: true },
  { key: 'latePay', label: 'Late Payment Fee', icon: 'number', removable: true },
  { key: 'sharepoint', label: '', icon: 'link', removable: true },
  { key: 'owner', label: 'Owner', icon: 'user', removable: true },
  { key: 'modifiedOn', label: 'Modified on', icon: 'date', removable: true },
  { key: 'regionalOffice', label: 'Regional Office', icon: 'list', removable: true },
  { key: 'farmingSector', label: 'Farming Sector', icon: 'list', removable: true },
  { key: 'bringForward', label: 'Bring Forward', icon: 'check', removable: true },
  { key: 'broughtForward', label: 'Brought Forward', icon: 'check', removable: true },
  { key: 'hasPartners', label: 'Has Partners', icon: 'check', removable: true },
  { key: 'inCombinedFarm', label: 'In Combined Farm', icon: 'check', removable: true },
  { key: 'manualReview', label: 'Manual Review', icon: 'check', removable: true },
  { key: 'enrolNoticeDate', label: 'EN Notice Sent Date', icon: 'date', removable: true },
  { key: 'fileReceivedDate', label: 'File Received Date', icon: 'date', removable: true },
  { key: 'feesPaidDate', label: 'Fees Paid Date', icon: 'date', removable: true },
];

export const DEFAULT_VISIBLE_KEYS: SortKey[] = [
  'flagged', 'pin', 'producer', 'year', 'taskStatus', 'enrolStatus', 'fee', 'sharepoint', 'owner', 'modifiedOn',
];

export const SORTKEY_TO_FIELD: Record<SortKey, string> = {
  pin: 'vsi_name',
  producer: 'vsi_participantid',
  year: 'vsi_programyearid',
  taskStatus: 'vsi_taskstatus',
  enrolStatus: 'vsi_enrolmentstatus',
  fee: 'vsi_calculatedenfee',
  totalFeesOwed: 'vsi_totalfeesowed',
  totalFeesPaid: 'vsi_totalfeespaid',
  enrolmentFee: 'vsi_enrolmentfee',
  latePay: 'vsi_latepaymentfee',
  sharepoint: 'vsi_sharepointdocumentfolder',
  owner: 'ownerid',
  modifiedOn: 'modifiedon',
  regionalOffice: 'vsi_enrollmentregionaloffice',
  farmingSector: 'vsi_farmingsector',
  bringForward: 'vsi_bringforward',
  broughtForward: 'vsi_broughtforward',
  hasPartners: 'vsi_haspartners',
  inCombinedFarm: 'vsi_incombinedfarm',
  manualReview: 'vsi_manualreview',
  enrolNoticeDate: 'vsi_enrolmentnoticesentdate',
  fileReceivedDate: 'vsi_filereceiveddate',
  feesPaidDate: 'vsi_enrolmentfeespaiddate',
  flagged: 'vsi_previousyearcalculatedenfee',
};

export const FIELD_TO_SORTKEY: Record<string, SortKey> = {
  ...Object.fromEntries(
    Object.entries(SORTKEY_TO_FIELD).map(([k, v]) => [v, k as SortKey])
  ),
  // Dataverse layoutxml stores lookup columns using the "name" suffix variant
  vsi_participantidname: 'producer',
  vsi_programyearidname: 'year',
  owneridname: 'owner',
  // OData _value prefix variants (less common in layoutxml but handle anyway)
  '_vsi_participantid_value': 'producer',
  '_vsi_programyearid_value': 'year',
  '_ownerid_value': 'owner',
} as Record<string, SortKey>;

export const ACTIVE_VIEW_KEY = 'enrolments-active-view';
export const USERQUERY_ENTITY = 'vsi_participantprogramyear';
export const USERQUERY_TYPE = 0;

export const DEFAULT_VIEW_SNAPSHOT: ViewPayload = {
  visibleColumnKeys: [...DEFAULT_VISIBLE_KEYS],
  columnWidths: {},
  sortKey: null,
  sortDir: 'asc',
  filters: { verifiedCalc: false, unverifiedCalc: false, flagged: false, partnerships: false, fortyFiveDayLetter: false },
  taskStatusFilter: [],
  enrolStatusFilter: [],
  taskFilterOp: 'equals',
  enrolFilterOp: 'equals',
  advFilterNodes: [],
  advLogicOp: 'AND',
};

export const ADV_FIELD_LABELS: Record<string, string> = {
  taskStatus: 'Task Status',
  enrolStatus: 'Enrol Status',
  pin: 'PIN',
  producer: 'Producer name',
  fee: 'Calculated fee',
  hasPartners: 'Has Partners',
  inCombinedFarm: 'In Combined Farm',
};

export const ADV_FIELD_OPTIONS: Record<string, 'choice' | 'text'> = {
  taskStatus: 'choice',
  enrolStatus: 'choice',
  pin: 'text',
  producer: 'text',
  fee: 'text',
  hasPartners: 'choice',
  inCombinedFarm: 'choice',
};

export const ADV_OP_LABELS: Record<string, string> = {
  equals: 'Equals',
  notEquals: 'Does not equal',
  contains: 'Contains',
  notContains: 'Does not contain',
  beginsWith: 'Begins with',
  endsWith: 'Ends with',
};
