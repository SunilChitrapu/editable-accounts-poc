import { ChangeEvent, FormEvent, useEffect, useMemo, useState } from "react";
import {
  buildTreeForWorkbooks,
  exportAllWorkbooks,
  exportWorkbook,
  findDataRow,
  getSheetHeaders,
  parseInput,
  parseWorkbookFile,
  updateDataRowLabel,
} from "./excel";
import type { LoadedWorkbook, TreeNode } from "./types";
import {
  expandedAllState,
  filterTree,
  mergeExpandedState,
  TreeView,
} from "./TreeView";

type ExpandedMap = Record<string, boolean>;

export function App() {
  const [workbooks, setWorkbooks] = useState<LoadedWorkbook[]>([]);
  const [status, setStatus] = useState<string>(
    "Add your financial workbook, chart of accounts, or bank exports. Everything stays in your browser."
  );
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const [expanded, setExpanded] = useState<ExpandedMap>({});
  const [groupByActivity, setGroupByActivity] = useState(true);
  const [treeFilter, setTreeFilter] = useState("");

  const tree = useMemo(
    () => buildTreeForWorkbooks(workbooks, { groupByActivity }),
    [workbooks, groupByActivity]
  );

  const visibleTree = useMemo(() => filterTree(tree, treeFilter), [tree, treeFilter]);

  useEffect(() => {
    if (workbooks.length === 0) {
      setExpanded({});
      return;
    }
    setExpanded((prev) => mergeExpandedState(prev, tree));
  }, [tree, workbooks.length]);

  const selectedRow = useMemo(
    () => (selectedRowId ? findDataRow(workbooks, selectedRowId) : null),
    [selectedRowId, workbooks]
  );

  const selectedFileName = useMemo(
    () => workbooks.find((w) => w.id === selectedRow?.fileId)?.fileName,
    [workbooks, selectedRow]
  );

  useEffect(() => {
    if (!selectedRowId) return;
    const active = document.querySelector(".tree-leaf.is-active");
    if (active instanceof HTMLElement) {
      active.scrollIntoView({ block: "nearest", inline: "nearest" });
    }
  }, [selectedRowId]);

  async function onFilesSelected(event: ChangeEvent<HTMLInputElement>) {
    const files = event.target.files;
    if (!files?.length) return;

    setStatus("Reading files…");
    const next: LoadedWorkbook[] = [...workbooks];
    const errors: string[] = [];

    for (const file of Array.from(files)) {
      try {
        const parsed = await parseWorkbookFile(file);
        next.push(parsed);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        errors.push(`${file.name}: ${msg}`);
      }
    }

    setWorkbooks(next);
    setSelectedRowId(null);
    if (errors.length) {
      setStatus(`Loaded ${next.length} file(s). Some files had issues: ${errors.join(" · ")}`);
    } else {
      setStatus(
        `Ready — ${next.length} workbook(s). Open the tree on the left; salary, rent, and other expenses are grouped automatically when activity grouping is on.`
      );
    }
    event.target.value = "";
  }

  function toggleNode(id: string) {
    setExpanded((prev) => ({
      ...prev,
      [id]: !prev[id],
    }));
  }

  function onClearAll() {
    setWorkbooks([]);
    setSelectedRowId(null);
    setExpanded({});
    setTreeFilter("");
    setStatus("Cleared. Import files to continue.");
  }

  function onExportOne(wb: LoadedWorkbook) {
    exportWorkbook(wb);
    setStatus(`Downloaded ${wb.baseName}_updated.xlsx`);
  }

  function onExportAll() {
    if (workbooks.length === 0) return;
    exportAllWorkbooks(workbooks);
    setStatus(
      workbooks.length > 1
        ? "Started downloads for each workbook (allow multiple downloads if prompted)."
        : "Download started."
    );
  }

  function onSaveTransaction(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!selectedRow) return;

    const formData = new FormData(event.currentTarget);
    for (const [key, value] of formData.entries()) {
      selectedRow.row[key] = parseInput(String(value));
    }
    const headers = getSheetHeaders(workbooks, selectedRow.fileId, selectedRow.sheetName);
    updateDataRowLabel(selectedRow, headers);
    setWorkbooks([...workbooks]);
    setStatus(`Saved changes to this row. Use Export when you are ready to save the Excel file.`);
  }

  function onExpandTreeAll() {
    setExpanded((prev) => ({ ...prev, ...expandedAllState(visibleTree, true) }));
  }

  function onCollapseTreeGroups() {
    const next: ExpandedMap = { ...expandedAllState(visibleTree, false) };
    function openFilesAndSheets(nodes: TreeNode[]): void {
      for (const n of nodes) {
        if (n.type === "file" || n.type === "sheet") next[n.id] = true;
        if (n.type !== "leaf") openFilesAndSheets(n.children);
      }
    }
    openFilesAndSheets(visibleTree);
    setExpanded(next);
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="app-header-inner">
          <h1 className="app-title">Workbook explorer</h1>
          <p className="app-subtitle">
            Browse TB, P&amp;L, balance sheet, transactions, aging, chart of accounts, and bank tabs in one
            place. Turn on <strong>activity grouping</strong> to collect salary, payroll, rent, software, and
            other expense patterns under clear folders — then edit any row and export.
          </p>
        </div>
      </header>

      <main className="app-main">
        <section className="toolbar card">
          <div className="toolbar-row">
            <div className="file-drop">
              <label className="btn btn-primary file-input-label">
                <input
                  type="file"
                  accept=".xlsx,.xlsm,.xlsb,.xls"
                  multiple
                  onChange={onFilesSelected}
                  className="visually-hidden"
                />
                Add Excel files
              </label>
              <span className="hint">Hold Ctrl to select several files (e.g. 2023 workbook + BOA + SVB).</span>
            </div>
            <div className="toolbar-actions">
              <button type="button" className="btn btn-primary" onClick={onExportAll} disabled={workbooks.length === 0}>
                Export all
              </button>
              <button type="button" className="btn btn-muted" onClick={onClearAll} disabled={workbooks.length === 0}>
                Clear all
              </button>
            </div>
          </div>

          <label className="toggle">
            <input
              type="checkbox"
              checked={groupByActivity}
              onChange={(e) => setGroupByActivity(e.target.checked)}
            />
            <span>
              Group by activity (expenses: salary &amp; payroll, rent, software, marketing, … — plus revenue and
              bank buckets)
            </span>
          </label>

          {workbooks.length > 0 && (
            <div className="status-banner" role="status">
              {status}
            </div>
          )}
          {workbooks.length === 0 && <p className="status-inline">{status}</p>}
        </section>

        {workbooks.length > 0 && (
          <section className="card file-cards">
            <h2 className="section-title">Loaded files</h2>
            <ul className="file-cards-list">
              {workbooks.map((wb) => (
                <li key={wb.id} className="file-card">
                  <div className="file-card-body">
                    <span className="file-card-name">{wb.fileName}</span>
                    <span className={`pill pill--${wb.kind}`}>{humanKind(wb.kind)}</span>
                    <span className="file-card-meta">{wb.sheets.length} sheets</span>
                  </div>
                  <button type="button" className="btn btn-sm btn-outline" onClick={() => onExportOne(wb)}>
                    Export
                  </button>
                </li>
              ))}
            </ul>
          </section>
        )}

        <section className="workspace">
          <aside className="panel tree-panel card">
            <div className="panel-head">
              <h2 className="panel-title">Data tree</h2>
              <div className="panel-tools">
                <button type="button" className="btn btn-sm btn-outline" onClick={onExpandTreeAll}>
                  Expand all
                </button>
                <button type="button" className="btn btn-sm btn-outline" onClick={onCollapseTreeGroups}>
                  Collapse groups
                </button>
              </div>
            </div>
            <div className="tree-search">
              <label htmlFor="treeFilter" className="visually-hidden">
                Filter rows
              </label>
              <input
                id="treeFilter"
                type="search"
                className="input"
                placeholder="Filter by description, account, amount…"
                value={treeFilter}
                onChange={(e) => setTreeFilter(e.target.value)}
                disabled={workbooks.length === 0}
              />
            </div>
            <div className="tree-scroll">
              {workbooks.length === 0 ? (
                <p className="empty-hint">
                  No files yet. Use <strong>Add Excel files</strong> to load your Saas Inc 2023 workbook, chart of
                  accounts, or bank statement exports.
                </p>
              ) : visibleTree.length === 0 ? (
                <p className="empty-hint">No rows match your filter. Try a shorter search.</p>
              ) : (
                <TreeView
                  nodes={visibleTree}
                  expanded={expanded}
                  selectedRowId={selectedRowId}
                  onToggle={toggleNode}
                  onSelectRow={setSelectedRowId}
                />
              )}
            </div>
          </aside>

          <section className={`panel editor-panel card ${selectedRow ? "is-floating" : ""}`}>
            <div className="panel-head">
              <h2 className="panel-title">Edit row</h2>
            </div>
            <div className="editor-body">
              {!selectedRow ? (
                <div className="empty-state">
                  <p className="empty-state-title">Select a row</p>
                  <p className="empty-hint">
                    Click any line in the tree. Values you change here update the sheet data; export to write a new
                    Excel file.
                  </p>
                </div>
              ) : (
                <form key={selectedRow.id} className="editor-form" onSubmit={onSaveTransaction}>
                  <div className="editor-meta card-inset">
                    <div>
                      <span className="meta-label">File</span>
                      <span className="meta-value">{selectedFileName ?? "—"}</span>
                    </div>
                    <div>
                      <span className="meta-label">Sheet</span>
                      <span className="meta-value">{selectedRow.sheetName}</span>
                    </div>
                    <div>
                      <span className="meta-label">Excel row</span>
                      <span className="meta-value">{selectedRow.rowIndex + 1}</span>
                    </div>
                  </div>
                  <div className="fields-scroll">
                    {Object.entries(selectedRow.row).map(([key, value]) => (
                      <div className="field-row" key={key}>
                        <label htmlFor={`f-${key}`}>{key || "(empty column name)"}</label>
                        <input id={`f-${key}`} name={key} className="input" defaultValue={String(value ?? "")} />
                      </div>
                    ))}
                  </div>
                  <div className="editor-footer">
                    <button type="submit" className="btn btn-primary">
                      Save changes
                    </button>
                  </div>
                </form>
              )}
            </div>
          </section>
        </section>
      </main>
    </div>
  );
}

function humanKind(kind: LoadedWorkbook["kind"]): string {
  switch (kind) {
    case "financial":
      return "Financial";
    case "coa":
      return "Chart of accounts";
    case "bank":
      return "Bank";
    default:
      return "Workbook";
  }
}
