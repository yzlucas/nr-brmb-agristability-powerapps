import type { SortKey, PersonalView, ViewPayload, QuickFilterState, AdvFilterNode, AdvFilterField } from '../types/enrollment';
import type { Userqueries } from '../generated/models/UserqueriesModel';
import type { Savedqueries } from '../generated/models/SavedqueriesModel';
import { SORTKEY_TO_FIELD, FIELD_TO_SORTKEY, DEFAULT_VIEW_SNAPSHOT, ACTIVE_VIEW_KEY } from '../constants/columns';
import { nextFilterId, serializeFilterNodes, deserializeFilterNodes } from './filterTree';
import { Vsi_participantprogramyearsvsi_taskstatus, Vsi_participantprogramyearsvsi_enrolmentstatus } from '../generated/models/Vsi_participantprogramyearsModel';

// Entity object type code for `vsi_participantprogramyear`.
// Populated once on mount via resolveEntityObjectTypeCode().
// Falls back to extracting from system view layoutxml (cached below).
let entityObjectTypeCode: string | null = null;

/** Called once on mount with the ObjectTypeCode from entity metadata. */
export function setEntityObjectTypeCode(code: number | string): void {
  entityObjectTypeCode = String(code);
}

type WinWithXrmWebApi = {
  Xrm?: {
    WebApi?: {
      retrieveMultipleRecords: (
        entityType: string,
        options?: string,
        maxPageSize?: number
      ) => Promise<{ entities: Array<Record<string, unknown>> }>;
    };
  };
};

/**
 * Resolves the ObjectTypeCode for vsi_participantprogramyear using Xrm.WebApi,
 * which is available in model-driven app PCF contexts and handles auth internally.
 * Falls back to cachedGridOpenTag (populated from system view layoutxml).
 */
export async function resolveEntityObjectTypeCode(): Promise<void> {
  if (entityObjectTypeCode) return; // already resolved
  const candidates = [window, window.parent, window.top];
  for (const candidate of candidates) {
    try {
      if (!candidate) continue;
      const webApi = (candidate as unknown as WinWithXrmWebApi).Xrm?.WebApi;
      if (!webApi?.retrieveMultipleRecords) continue;
      const result = await webApi.retrieveMultipleRecords(
        'savedquery',
        `?$select=returnedtypecode&$top=1&$filter=returnedtypecode eq ${JSON.stringify('vsi_participantprogramyear')}`,
        1
      );
      const record = result?.entities?.[0];
      if (record) {
        // returnedtypecode from raw Xrm.WebApi OData is the integer type code
        const raw = record['returnedtypecode'];
        const num = Number(raw);
        if (!isNaN(num) && num > 0) {
          entityObjectTypeCode = String(num);
          console.log('[viewSerializer] Resolved ObjectTypeCode via Xrm.WebApi:', entityObjectTypeCode);
          return;
        }
      }
    } catch (e) {
      console.warn('[viewSerializer] Xrm.WebApi resolveEntityObjectTypeCode error:', e);
    }
  }
  if (cachedGridOpenTag) {
    console.log('[viewSerializer] Using cachedGridOpenTag fallback for layoutxml object type code');
  } else {
    console.warn('[viewSerializer] Could not resolve entity ObjectTypeCode — layoutxml may be rejected by Dataverse');
  }
}

// Fallback: cache the full <grid ...> opening tag from the first system view
// we parse that contains an `object` attribute.
let cachedGridOpenTag: string | null = null;

// These keys are UI-only or not valid view columns in the model-driven app.
// They must be excluded from layoutxml and fetchxml to avoid blank/invalid columns.
// Note: 'pin' maps to vsi_name (the jump/link field) and must remain INCLUDED
// so Dataverse renders it as the named "Enrolment Name" link column rather than
// a floating navigation icon.
const LAYOUT_XML_EXCLUDED_KEYS = new Set<SortKey>(['flagged']);

/**
 * Generates a Dataverse fetchxml for a personal view.
 * Explicitly lists all visible columns as <attribute> elements — required by
 * model-driven app views to populate data in the grid.
 */
export function generateFetchXml(keys: SortKey[], advFilterNodes: unknown[] = []): string {
  const includedKeys = keys.filter(k => !LAYOUT_XML_EXCLUDED_KEYS.has(k));
  const fields = new Set<string>();
  for (const k of includedKeys) {
    fields.add(SORTKEY_TO_FIELD[k]);
  }
  const attrs = Array.from(fields)
    .map(f => `<attribute name="${f}"/>`)
    .join('');

  let filterXml = '';
  if (advFilterNodes.length > 0) {
    const nodes = deserializeFilterNodes(advFilterNodes);
    const parts = nodes.map(advNodeToFilterXml).filter(Boolean);
    if (parts.length === 1) filterXml = parts[0];
    else if (parts.length > 1) filterXml = `<filter type="and">${parts.join('')}</filter>`;
  }

  return `<fetch><entity name="vsi_participantprogramyear">${attrs}${filterXml}<order attribute="vsi_name" descending="false"/></entity></fetch>`;
}

// ── fetchxml filter helpers ──────────────────────────────────────────────────

const ADV_FIELD_TO_ATTR: Partial<Record<AdvFilterField, string>> = {
  taskStatus: 'vsi_taskstatus',
  enrolStatus: 'vsi_enrolmentstatus',
  pin: 'vsi_name',
  fee: 'vsi_calculatedenfee',
  hasPartners: 'vsi_haspartners',
  inCombinedFarm: 'vsi_incombinedfarm',
  // producer omitted — lookup display name requires linked-entity join
};

function labelToValue(field: AdvFilterField, label: string): string | null {
  if (field === 'taskStatus') {
    const entry = Object.entries(Vsi_participantprogramyearsvsi_taskstatus).find(([, v]) => v === label);
    return entry?.[0] ?? null;
  }
  if (field === 'enrolStatus') {
    const entry = Object.entries(Vsi_participantprogramyearsvsi_enrolmentstatus).find(([, v]) => v === label);
    return entry?.[0] ?? null;
  }
  return null;
}

function advRowToConditions(node: AdvFilterNode & { kind: 'row' }): string {
  const attr = ADV_FIELD_TO_ATTR[node.field];
  if (!attr || !node.field) return '';

  // Boolean fields (Yes/No)
  if (node.field === 'hasPartners' || node.field === 'inCombinedFarm') {
    if (node.values.size === 0) return '';
    const op = node.operator === 'notEquals' ? 'ne' : 'eq';
    const conds = [...node.values].map(v => `<condition attribute="${attr}" operator="${op}" value="${v === 'Yes' ? '1' : '0'}"/>`);
    if (conds.length === 1) return conds[0];
    const joinType = node.operator === 'notEquals' ? 'and' : 'or';
    return `<filter type="${joinType}">${conds.join('')}</filter>`;
  }

  // Choice fields (integer enum)
  if (node.field === 'taskStatus' || node.field === 'enrolStatus') {
    if (node.values.size === 0) return '';
    const op = node.operator === 'notEquals' ? 'ne' : 'eq';
    const conds = [...node.values]
      .map(label => labelToValue(node.field, label))
      .filter((v): v is string => v !== null)
      .map(v => `<condition attribute="${attr}" operator="${op}" value="${v}"/>`); 
    if (conds.length === 0) return '';
    if (conds.length === 1) return conds[0];
    const joinType = node.operator === 'notEquals' ? 'and' : 'or';
    return `<filter type="${joinType}">${conds.join('')}</filter>`;
  }

  // Text fields
  if (!node.textValue) return '';
  let fetchOp: string;
  let value = node.textValue;
  switch (node.operator) {
    case 'equals':     fetchOp = 'eq'; break;
    case 'notEquals':  fetchOp = 'ne'; break;
    case 'contains':   fetchOp = 'like'; value = `%${node.textValue}%`; break;
    case 'notContains':fetchOp = 'not-like'; value = `%${node.textValue}%`; break;
    case 'beginsWith': fetchOp = 'like'; value = `${node.textValue}%`; break;
    case 'endsWith':   fetchOp = 'like'; value = `%${node.textValue}`; break;
    default:           fetchOp = 'eq';
  }
  return `<condition attribute="${attr}" operator="${fetchOp}" value="${value}"/>`;
}

function advNodeToFilterXml(node: AdvFilterNode): string {
  if (node.kind === 'row') return advRowToConditions(node);
  const childXmls = node.children.map(advNodeToFilterXml).filter(Boolean);
  if (childXmls.length === 0) return '';
  if (childXmls.length === 1) return childXmls[0];
  return `<filter type="${node.logic.toLowerCase()}">${childXmls.join('')}</filter>`;
}

export function generateLayoutXml(keys: SortKey[], widths: Partial<Record<SortKey, number>>): string {
  const cells = keys
    .filter(k => !LAYOUT_XML_EXCLUDED_KEYS.has(k))
    .map(k => {
      const field = SORTKEY_TO_FIELD[k];
      const w = widths[k] ?? 125;
      return `<cell name="${field}" width="${w}" />`;
    })
    .join('');

  // Build the <grid> opening tag. Prefer the explicitly fetched object type code;
  // fall back to the tag template extracted from a system view's layoutxml.
  let gridOpen: string;
  if (entityObjectTypeCode) {
    gridOpen = `<grid name="resultset" object="${entityObjectTypeCode}" jump="vsi_name" select="1" icon="1" preview="1">`;
  } else if (cachedGridOpenTag) {
    gridOpen = cachedGridOpenTag;
  } else {
    // Last resort — no object attribute; Dataverse schema may reject this
    gridOpen = '<grid name="resultset" jump="vsi_name" select="1" icon="1" preview="1">';
    console.warn('[viewSerializer] generateLayoutXml: object type code not yet resolved — layoutxml may be rejected by Dataverse');
  }

  return `${gridOpen}<row name="result" id="vsi_participantprogramyearid">${cells}</row></grid>`;
}

export function parseLayoutXml(xml: string | undefined | null): SortKey[] | null {
  if (!xml) return null;
  try {
    // Cache the full <grid ...> opening tag from system view layoutxml so we
    // can use it as a template fallback if entity metadata isn't fetched yet.
    if (!cachedGridOpenTag) {
      const tagMatch = xml.match(/<grid[^>]+>/);
      if (tagMatch && /object="\d+"/.test(tagMatch[0])) {
        cachedGridOpenTag = tagMatch[0];
      }
    }
    const cellRegex = /<cell\s[^>]*name="([^"]+)"/g;
    const keys: SortKey[] = [];
    let match: RegExpExecArray | null;
    while ((match = cellRegex.exec(xml)) !== null) {
      const field = match[1];
      const sk = FIELD_TO_SORTKEY[field];
      if (sk) keys.push(sk);
    }
    return keys.length > 0 ? keys : null;
  } catch { return null; }
}

export function userqueryToView(uq: Userqueries): PersonalView {
  try {
    const payload: ViewPayload = JSON.parse(uq.layoutjson ?? '{}');
    if (payload.visibleColumnKeys) {
      return { id: uq.userqueryid, name: uq.name, source: 'personal', ...payload };
    }
  } catch { /* layoutjson not in our format */ }
  const xmlCols = parseLayoutXml(uq.layoutxml);
  const snapshot: ViewPayload = xmlCols
    ? { ...DEFAULT_VIEW_SNAPSHOT, visibleColumnKeys: xmlCols }
    : { ...DEFAULT_VIEW_SNAPSHOT };
  return { id: uq.userqueryid, name: uq.name, source: 'personal', ...snapshot };
}

// Columns that must always be present regardless of what the view definition says
const REQUIRED_COLUMNS: SortKey[] = ['flagged', 'pin'];

/** Merge required columns into a parsed column list without duplicating. */
function mergeRequiredColumns(keys: SortKey[]): SortKey[] {
  const required = REQUIRED_COLUMNS.filter(k => !keys.includes(k));
  return [...required, ...keys];
}

export function savedqueryToView(sq: Savedqueries): PersonalView {
  const fetchFilters = parseFetchXmlToFilters(sq.fetchxml);
  // Serialize nodes (Set→array) so deserializeFilterNodes in applyView works correctly
  const advFilterNodes = serializeFilterNodes(parseFetchXmlToAdvNodes(sq.fetchxml));

  // Parse column layout from layoutxml; fall back to defaults if unparseable
  const xmlCols = parseLayoutXml(sq.layoutxml);
  const visibleColumnKeys = xmlCols
    ? mergeRequiredColumns(xmlCols)
    : [...DEFAULT_VIEW_SNAPSHOT.visibleColumnKeys];

  const snapshot: ViewPayload = {
    ...DEFAULT_VIEW_SNAPSHOT,
    visibleColumnKeys,
    filters: { ...DEFAULT_VIEW_SNAPSHOT.filters, ...fetchFilters },
    advFilterNodes,
  };
  return { id: sq.savedqueryid, name: sq.name, source: 'system', ...snapshot };
}

/** Map Dataverse condition attributes to AdvFilterField for boolean (yes/no) fields. */
const BOOLEAN_CONDITION_FIELDS: Record<string, AdvFilterField> = {
  vsi_haspartners: 'hasPartners',
  vsi_incombinedfarm: 'inCombinedFarm',
};

/**
 * Parse a Dataverse fetchxml `<filter>` element and convert conditions for
 * known fields into AdvFilterNode[]. Returns an empty array if no known
 * conditions are found or the xml is invalid.
 */
export function parseFetchXmlToAdvNodes(fetchxml: string | undefined | null): AdvFilterNode[] {
  if (!fetchxml) return [];
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(fetchxml, 'text/xml');
    if (doc.querySelector('parsererror')) return [];

    function processFilter(filterEl: Element): AdvFilterNode | null {
      const ownLogic = (filterEl.getAttribute('type') ?? 'and').toUpperCase() as 'AND' | 'OR';
      const children: AdvFilterNode[] = [];
      // In Dataverse fetchxml, sibling <filter type="or"> elements indicate
      // the child should be OR-joined with its previous sibling (not AND).
      let hasOrSiblingFilter = false;

      for (const child of Array.from(filterEl.children)) {
        if (child.tagName === 'condition') {
          const attr = child.getAttribute('attribute') ?? '';
          const op = child.getAttribute('operator') ?? 'eq';
          const rawVal = child.getAttribute('value') ?? '1';
          const field = BOOLEAN_CONDITION_FIELDS[attr];
          if (!field) continue;

          const isTrue = rawVal === '1' || rawVal.toLowerCase() === 'true';
          const isNegated = op === 'ne' || op === 'neq' || op === 'not-eq';
          const choiceVal = isTrue ? 'Yes' : 'No';
          const operator = isNegated ? 'notEquals' : 'equals';

          children.push({
            kind: 'row',
            id: nextFilterId(),
            field,
            operator,
            values: new Set([choiceVal]),
            textValue: '',
          });
        } else if (child.tagName === 'filter') {
          const siblingType = (child.getAttribute('type') ?? 'and').toUpperCase();
          if (siblingType === 'OR') hasOrSiblingFilter = true;
          const nested = processFilter(child);
          if (nested) children.push(nested);
        }
      }

      if (children.length === 0) return null;
      if (children.length === 1 && children[0].kind === 'row') return children[0];
      // If any sibling filter declared OR, that's the effective join between siblings
      const logic = hasOrSiblingFilter ? 'OR' : ownLogic;
      return { kind: 'group', id: nextFilterId(), logic, children };
    }

    const entityEl = doc.querySelector('entity');
    const filterEl = entityEl?.querySelector(':scope > filter');
    if (!filterEl) return [];

    const result = processFilter(filterEl);
    return result ? [result] : [];
  } catch {
    return [];
  }
}

/** Parse a Dataverse fetchxml string and extract known QuickFilterState flags. */
export function parseFetchXmlToFilters(fetchxml: string | undefined | null): Partial<QuickFilterState> {
  if (!fetchxml) return {};
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(fetchxml, 'text/xml');
    const conditions = Array.from(doc.querySelectorAll('condition')).map(c => ({
      attr: c.getAttribute('attribute') ?? '',
      op: c.getAttribute('operator') ?? 'eq',
      val: (c.getAttribute('value') ?? '').toLowerCase(),
    }));

    const hasAttr = (attr: string) => conditions.some(c => c.attr === attr);
    const hasEq = (attr: string, val: string) =>
      conditions.some(c => c.attr === attr && c.op === 'eq' && c.val === val);
    const result: Partial<QuickFilterState> = {};

    // Partnerships/Combined: any condition on vsi_haspartners or vsi_incombinedfarm
    if (hasAttr('vsi_haspartners') || hasAttr('vsi_incombinedfarm')) {
      result.partnerships = true;
    }

    // Verified EN Calculated = enrolment status = VerifiedENCalculalted (865520006)
    if (hasEq('vsi_enrolmentstatus', '865520006')) {
      result.verifiedCalc = true;
    }

    // Unverified EN Calculated = enrolment status = UnverifiedENCalculated (865520005)
    if (hasEq('vsi_enrolmentstatus', '865520005')) {
      result.unverifiedCalc = true;
    }

    // 45-day letter: enrolment status = _45DayLetter (865520010)
    if (hasEq('vsi_enrolmentstatus', '865520010')) {
      result.fortyFiveDayLetter = true;
    }

    return result;
  } catch {
    return {};
  }
}

export function loadActiveViewId(): string | null {
  return localStorage.getItem(ACTIVE_VIEW_KEY);
}

export function saveActiveViewId(id: string | null) {
  if (id) localStorage.setItem(ACTIVE_VIEW_KEY, id);
  else localStorage.removeItem(ACTIVE_VIEW_KEY);
}
