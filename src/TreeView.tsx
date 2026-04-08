import type { TreeNode } from "./types";

type ExpandedMap = Record<string, boolean>;

export function TreeView(props: {
  nodes: TreeNode[];
  expanded: ExpandedMap;
  selectedRowId: string | null;
  onToggle: (id: string) => void;
  onSelectRow: (id: string) => void;
}) {
  const { nodes, expanded, selectedRowId, onToggle, onSelectRow } = props;
  return (
    <div className="tree-container">
      {nodes.map((node) => (
        <TreeNodeView
          key={node.id}
          node={node}
          expanded={expanded}
          selectedRowId={selectedRowId}
          onToggle={onToggle}
          onSelectRow={onSelectRow}
        />
      ))}
    </div>
  );
}

function TreeNodeView(props: {
  node: TreeNode;
  expanded: ExpandedMap;
  selectedRowId: string | null;
  onToggle: (id: string) => void;
  onSelectRow: (id: string) => void;
}) {
  const { node, expanded, selectedRowId, onToggle, onSelectRow } = props;

  if (node.type === "leaf") {
    return (
      <button
        type="button"
        data-row-id={node.id}
        className={selectedRowId === node.id ? "tree-leaf is-active" : "tree-leaf"}
        onClick={() => onSelectRow(node.id)}
      >
        <span className="tree-leaf-label">{node.label}</span>
      </button>
    );
  }

  const open =
    node.type === "group" ? expanded[node.id] === true : expanded[node.id] !== false;
  const roleLabel =
    node.type === "file" ? "Workbook" : node.type === "sheet" ? "Sheet" : "Group";
  return (
    <div className={`tree-node tree-node--${node.type}`}>
      <button
        type="button"
        className={`node-toggle node-toggle--${node.type}`}
        onClick={() => onToggle(node.id)}
        title={roleLabel}
      >
        <span className="node-chevron" aria-hidden>
          {open ? "▼" : "▶"}
        </span>
        <span className="node-label">{node.label}</span>
      </button>
      {open && (
        <div className="node-children">
          {node.children.map((child) => (
            <TreeNodeView
              key={child.id}
              node={child}
              expanded={expanded}
              selectedRowId={selectedRowId}
              onToggle={onToggle}
              onSelectRow={onSelectRow}
            />
          ))}
        </div>
      )}
    </div>
  );
}

export function defaultExpandedState(nodes: TreeNode[]): ExpandedMap {
  const map: ExpandedMap = {};
  function walk(n: TreeNode): void {
    if (n.type === "file" || n.type === "sheet") {
      map[n.id] = true;
    } else if (n.type === "group") {
      map[n.id] = false;
    }
    if (n.type !== "leaf") {
      n.children.forEach(walk);
    }
  }
  nodes.forEach(walk);
  return map;
}

function collectExpandableIds(nodes: TreeNode[], into: Set<string>): void {
  for (const n of nodes) {
    if (n.type === "leaf") continue;
    into.add(n.id);
    collectExpandableIds(n.children, into);
  }
}

/** Defaults for new nodes; keep prior open/closed for ids that still exist. */
export function mergeExpandedState(prev: ExpandedMap, nodes: TreeNode[]): ExpandedMap {
  const defaults = defaultExpandedState(nodes);
  const valid = new Set<string>();
  collectExpandableIds(nodes, valid);
  const out: ExpandedMap = { ...defaults };
  for (const k of Object.keys(prev)) {
    if (valid.has(k)) out[k] = prev[k];
  }
  return out;
}

export function filterTree(nodes: TreeNode[], query: string): TreeNode[] {
  const q = query.trim().toLowerCase();
  if (!q) return nodes;

  function filterNode(n: TreeNode): TreeNode | null {
    if (n.type === "leaf") {
      return n.label.toLowerCase().includes(q) ? n : null;
    }
    const kids = n.children.map(filterNode).filter((c): c is TreeNode => c !== null);
    if (kids.length === 0) return null;
    return { ...n, children: kids };
  }

  return nodes.map(filterNode).filter((c): c is TreeNode => c !== null);
}

export function expandedAllState(nodes: TreeNode[], open: boolean): ExpandedMap {
  const map: ExpandedMap = {};
  function walk(n: TreeNode): void {
    if (n.type === "leaf") return;
    map[n.id] = open;
    n.children.forEach(walk);
  }
  nodes.forEach(walk);
  return map;
}
