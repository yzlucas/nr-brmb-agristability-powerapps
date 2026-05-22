import { useState, useRef, useEffect } from 'react';
import { createPortal } from 'react-dom';
import type { AdvFilterNode, AdvFilterRow, AdvFilterGroup, AdvFilterField, AdvFilterOp, LogicOp } from '../types/enrollment';
import { ADV_FIELD_LABELS, ADV_FIELD_OPTIONS, ADV_OP_LABELS } from '../constants/columns';
import { getChoiceOptions, formatEnrolmentStatusDisplay } from '../utils/helpers';
import {
  cloneNode, updateNodeInTree, wrapNodeInGroup, ungroupNode,
  addNodeToParent, emptyFilterRow, emptyFilterGroup, summarizeGroup,
} from '../utils/filterTree';

function NodeContextMenu({
  nodeKind,
  onDelete,
  onMakeGroup,
  onUngroup,
}: {
  nodeKind: 'row' | 'group';
  onDelete: () => void;
  onMakeGroup?: () => void;
  onUngroup?: () => void;
}) {
  const [open, setOpen] = useState(false);
  const btnRef = useRef<HTMLButtonElement>(null);
  const [menuStyle, setMenuStyle] = useState<React.CSSProperties>({});

  useEffect(() => {
    if (open && btnRef.current) {
      const r = btnRef.current.getBoundingClientRect();
      setMenuStyle({ position: 'fixed', top: r.bottom + 2, left: r.right - 150, minWidth: 150 });
    }
  }, [open]);

  return (
    <div className="ef-ctx-wrapper">
      <button ref={btnRef} className="ef-ctx-btn" onClick={() => setOpen(o => !o)} title="More options">⋯</button>
      {open && createPortal(
        <>
          <div className="ef-add-backdrop" onClick={() => setOpen(false)} />
          <div className="ef-ctx-menu" style={menuStyle}>
            <button className="ef-ctx-item" onClick={() => { onDelete(); setOpen(false); }}>
              <span className="ef-ctx-icon">🗑</span> Delete
            </button>
            {nodeKind === 'row' && onMakeGroup && (
              <button className="ef-ctx-item" onClick={() => { onMakeGroup(); setOpen(false); }}>
                <span className="ef-ctx-icon">⊞</span> Make group
              </button>
            )}
            {nodeKind === 'group' && onUngroup && (
              <button className="ef-ctx-item" onClick={() => { onUngroup(); setOpen(false); }}>
                <span className="ef-ctx-icon">⊟</span> Ungroup
              </button>
            )}
          </div>
        </>,
        document.body
      )}
    </div>
  );
}

function AddMenu({ onAddRow, onAddGroup }: { onAddRow: () => void; onAddGroup: () => void }) {
  const [open, setOpen] = useState(false);
  const btnRef = useRef<HTMLButtonElement>(null);
  const [menuStyle, setMenuStyle] = useState<React.CSSProperties>({});

  useEffect(() => {
    if (open && btnRef.current) {
      const r = btnRef.current.getBoundingClientRect();
      setMenuStyle({ position: 'fixed', top: r.bottom + 2, left: r.left, minWidth: 170 });
    }
  }, [open]);

  return (
    <div className="ef-add-wrapper">
      <button ref={btnRef} className="ef-add-btn" onClick={() => setOpen(o => !o)}>+ Add ▾</button>
      {open && createPortal(
        <>
          <div className="ef-add-backdrop" onClick={() => setOpen(false)} />
          <div className="ef-add-menu" style={menuStyle}>
            <button className="ef-add-menu-item" onClick={() => { onAddRow(); setOpen(false); }}>
              <span className="ef-add-icon">⊕</span> Add row
            </button>
            <button className="ef-add-menu-item" onClick={() => { onAddGroup(); setOpen(false); }}>
              <span className="ef-add-icon">≡</span> Add group
            </button>
          </div>
        </>,
        document.body
      )}
    </div>
  );
}

function FilterRowEditor({
  row,
  onChange,
  onDelete,
  onMakeGroup,
}: {
  row: AdvFilterRow;
  onChange: (patch: Partial<AdvFilterRow>) => void;
  onDelete: () => void;
  onMakeGroup: () => void;
}) {
  const fieldType = row.field ? ADV_FIELD_OPTIONS[row.field] : null;
  const choiceOpts = fieldType === 'choice' ? getChoiceOptions(row.field) : [];
  const formatChoiceLabel = (opt: string) => row.field === 'enrolStatus' ? formatEnrolmentStatusDisplay(opt) : opt;
  const [valDropdownOpen, setValDropdownOpen] = useState(false);
  const valBtnRef = useRef<HTMLButtonElement>(null);
  const [dropdownStyle, setDropdownStyle] = useState<React.CSSProperties>({});

  useEffect(() => {
    if (valDropdownOpen && valBtnRef.current) {
      const r = valBtnRef.current.getBoundingClientRect();
      setDropdownStyle({ position: 'fixed', top: r.bottom + 2, left: r.left, minWidth: r.width });
    }
  }, [valDropdownOpen]);

  const toggleValue = (val: string) => {
    const next = new Set(row.values);
    if (next.has(val)) next.delete(val); else next.add(val);
    onChange({ values: next });
  };

  return (
    <div className="ef-row">
      <input type="checkbox" className="ef-node-check" readOnly />
      <select
        className="ef-cell ef-cell-field ef-select"
        value={row.field}
        onChange={e => onChange({ field: e.target.value as AdvFilterField })}
      >
        <option value="" disabled>Select a field</option>
        {(Object.keys(ADV_FIELD_LABELS) as AdvFilterField[]).map(f => (
          <option key={f} value={f}>{ADV_FIELD_LABELS[f]}</option>
        ))}
      </select>
      <select
        className="ef-cell ef-cell-op ef-select"
        value={row.operator}
        onChange={e => onChange({ operator: e.target.value as AdvFilterOp })}
      >
        {fieldType === 'choice' ? (
          <>
            <option value="equals">Equals</option>
            <option value="notEquals">Does not equal</option>
          </>
        ) : (
          (Object.keys(ADV_OP_LABELS) as AdvFilterOp[]).map(op => (
            <option key={op} value={op}>{ADV_OP_LABELS[op]}</option>
          ))
        )}
      </select>
      {fieldType === 'choice' ? (
        <div className="ef-cell ef-cell-val ef-val-choice-wrapper">
          <button ref={valBtnRef} className="ef-val-choice-btn" onClick={() => setValDropdownOpen(o => !o)}>
            {row.values.size === 0 ? <span className="ef-val-placeholder">Value</span> : [...row.values].map(formatChoiceLabel).join(', ')}
            <span className="ef-val-chevron">▾</span>
          </button>
          {valDropdownOpen && createPortal(
            <>
              <div className="ef-add-backdrop" onClick={() => setValDropdownOpen(false)} />
              <div className="ef-val-dropdown" style={dropdownStyle}>
                {choiceOpts.map(opt => (
                  <label key={opt} className="ef-val-opt">
                    <input type="checkbox" checked={row.values.has(opt)} onChange={() => toggleValue(opt)} />
                    <span>{formatChoiceLabel(opt)}</span>
                  </label>
                ))}
              </div>
            </>,
            document.body
          )}
        </div>
      ) : (
        <input
          className="ef-cell ef-cell-val ef-text-input"
          type={row.field === 'fee' ? 'number' : 'text'}
          value={row.textValue}
          placeholder="Value"
          onChange={e => onChange({ textValue: e.target.value })}
        />
      )}
      <NodeContextMenu nodeKind="row" onDelete={onDelete} onMakeGroup={onMakeGroup} />
    </div>
  );
}

function FilterGroupEditor({
  group,
  onUpdate,
  onRemove,
  onMakeGroup,
  onUngroup,
  onAddRow,
  onAddGroup,
}: {
  group: AdvFilterGroup;
  onUpdate: (id: number, updater: (n: AdvFilterNode) => AdvFilterNode | null) => void;
  onRemove: (id: number) => void;
  onMakeGroup: (id: number) => void;
  onUngroup: (id: number) => void;
  onAddRow: (parentId: number | null) => void;
  onAddGroup: (parentId: number | null) => void;
}) {
  const [collapsed, setCollapsed] = useState(false);

  const setGroupLogic = (logic: LogicOp) =>
    onUpdate(group.id, n => ({ ...n as AdvFilterGroup, logic }));

  if (collapsed) {
    return (
      <div className="ef-group ef-group-collapsed">
        <div className="ef-group-collapsed-bar">
          <input type="checkbox" className="ef-node-check" readOnly />
          <select className="ef-logic-select ef-logic-select-sm" value={group.logic} onChange={e => setGroupLogic(e.target.value as LogicOp)}>
            <option value="AND">AND</option>
            <option value="OR">OR</option>
          </select>
          <span className="ef-col-hdr ef-col-field">Field</span>
          <span className="ef-col-hdr ef-col-op">Operator</span>
          <span className="ef-col-hdr ef-col-val">Value</span>
          <span className="ef-group-summary">{summarizeGroup(group)}</span>
          <button className="ef-collapse-btn" onClick={() => setCollapsed(false)} title="Expand">↙</button>
          <NodeContextMenu nodeKind="group" onDelete={() => onRemove(group.id)} onUngroup={() => onUngroup(group.id)} />
        </div>
      </div>
    );
  }

  return (
    <div className="ef-group">
      <div className="ef-group-top">
        <input type="checkbox" className="ef-node-check" readOnly />
        <select className="ef-logic-select ef-logic-select-sm" value={group.logic} onChange={e => setGroupLogic(e.target.value as LogicOp)}>
          <option value="AND">AND</option>
          <option value="OR">OR</option>
        </select>
        <span className="ef-col-hdr ef-col-field">Field</span>
        <span className="ef-col-hdr ef-col-op">Operator</span>
        <span className="ef-col-hdr ef-col-val">Value</span>
        <button className="ef-collapse-btn" onClick={() => setCollapsed(true)} title="Collapse">↗</button>
        <NodeContextMenu nodeKind="group" onDelete={() => onRemove(group.id)} onUngroup={() => onUngroup(group.id)} />
      </div>
      <div className="ef-tree-nodes ef-group-nodes">
        {group.children.map((child, i) => (
          <div key={child.id} className={`ef-tree-item${i === group.children.length - 1 ? ' ef-tree-item-last' : ''}`}>
            <div className="ef-tree-connector">
              <div className="ef-tree-vline" />
              <div className="ef-tree-hline" />
            </div>
            <div className="ef-tree-content">
              {child.kind === 'row' ? (
                <FilterRowEditor
                  row={child}
                  onChange={patch => onUpdate(child.id, n => {
                    const updated = { ...n as AdvFilterRow, ...patch };
                    if (patch.field && patch.field !== (n as AdvFilterRow).field) {
                      updated.values = new Set();
                      updated.textValue = '';
                      updated.operator = 'equals';
                    }
                    return updated;
                  })}
                  onDelete={() => onRemove(child.id)}
                  onMakeGroup={() => onMakeGroup(child.id)}
                />
              ) : (
                <FilterGroupEditor
                  group={child}
                  onUpdate={onUpdate}
                  onRemove={onRemove}
                  onMakeGroup={onMakeGroup}
                  onUngroup={onUngroup}
                  onAddRow={onAddRow}
                  onAddGroup={onAddGroup}
                />
              )}
            </div>
          </div>
        ))}
        <div className="ef-tree-item ef-tree-item-last">
          <div className="ef-tree-connector">
            <div className="ef-tree-vline" />
            <div className="ef-tree-hline" />
          </div>
          <div className="ef-tree-content">
            <AddMenu onAddRow={() => onAddRow(group.id)} onAddGroup={() => onAddGroup(group.id)} />
          </div>
        </div>
      </div>
    </div>
  );
}

export function EditFiltersPanel({
  filterNodes,
  logicOp,
  onApply,
  onCancel,
}: {
  filterNodes: AdvFilterNode[];
  logicOp: LogicOp;
  onApply: (nodes: AdvFilterNode[], logic: LogicOp) => void;
  onCancel: () => void;
}) {
  const [nodes, setNodes] = useState<AdvFilterNode[]>(() =>
    filterNodes.length > 0 ? filterNodes.map(cloneNode) : []
  );
  const [logic, setLogic] = useState<LogicOp>(logicOp);

  const handleUpdate = (id: number, updater: (n: AdvFilterNode) => AdvFilterNode | null) =>
    setNodes(prev => updateNodeInTree(prev, id, updater));
  const handleRemove = (id: number) => handleUpdate(id, () => null);
  const handleMakeGroup = (id: number) => setNodes(prev => wrapNodeInGroup(prev, id));
  const handleUngroup = (id: number) => setNodes(prev => ungroupNode(prev, id));
  const handleAddRow = (parentId: number | null) => setNodes(prev => addNodeToParent(prev, parentId, emptyFilterRow()));
  const handleAddGroup = (parentId: number | null) => setNodes(prev => addNodeToParent(prev, parentId, emptyFilterGroup()));
  const clearAll = () => { setNodes([]); };

  return (
    <div className="ef-overlay" onClick={onCancel}>
      <div className="ef-panel" onClick={e => e.stopPropagation()}>
        <div className="ef-header">
          <h3>Edit filters: Enrolments</h3>
          <button className="ef-close" onClick={onCancel}>✕</button>
        </div>
        <div className="ef-toolbar">
          <button className="ef-toolbar-btn" onClick={clearAll}>
            <span className="ef-tb-icon">&#x25BD;</span> Delete all filters
          </button>
        </div>
        <div className="ef-body">
          <div className="ef-tree">
            <div className="ef-tree-top">
              <select className="ef-logic-select" value={logic} onChange={e => setLogic(e.target.value as LogicOp)}>
                <option value="AND">AND</option>
                <option value="OR">OR</option>
              </select>
              <span className="ef-col-hdr ef-col-field">Field</span>
              <span className="ef-col-hdr ef-col-op">Operator</span>
              <span className="ef-col-hdr ef-col-val">Value</span>
              <span className="ef-col-hdr ef-col-actions" />
            </div>
            <div className="ef-tree-nodes">
              {nodes.map((node, i) => (
                <div key={node.id} className={`ef-tree-item${i === nodes.length - 1 ? ' ef-tree-item-last' : ''}`}>
                  <div className="ef-tree-connector">
                    <div className="ef-tree-vline" />
                    <div className="ef-tree-hline" />
                  </div>
                  <div className="ef-tree-content">
                    {node.kind === 'row' ? (
                      <FilterRowEditor
                        row={node}
                        onChange={patch => handleUpdate(node.id, n => {
                          const updated = { ...n as AdvFilterRow, ...patch };
                          if (patch.field && patch.field !== (n as AdvFilterRow).field) {
                            updated.values = new Set();
                            updated.textValue = '';
                            updated.operator = 'equals';
                          }
                          return updated;
                        })}
                        onDelete={() => handleRemove(node.id)}
                        onMakeGroup={() => handleMakeGroup(node.id)}
                      />
                    ) : (
                      <FilterGroupEditor
                        group={node}
                        onUpdate={handleUpdate}
                        onRemove={handleRemove}
                        onMakeGroup={handleMakeGroup}
                        onUngroup={handleUngroup}
                        onAddRow={handleAddRow}
                        onAddGroup={handleAddGroup}
                      />
                    )}
                  </div>
                </div>
              ))}
              <div className="ef-tree-item ef-tree-item-last">
                <div className="ef-tree-connector">
                  <div className="ef-tree-vline" />
                  <div className="ef-tree-hline" />
                </div>
                <div className="ef-tree-content">
                  <AddMenu onAddRow={() => handleAddRow(null)} onAddGroup={() => handleAddGroup(null)} />
                </div>
              </div>
            </div>
          </div>
        </div>
        <div className="ef-footer">
          <button className="ef-apply" onClick={() => onApply(nodes, logic)}>Apply</button>
          <button className="ef-cancel" onClick={onCancel}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
