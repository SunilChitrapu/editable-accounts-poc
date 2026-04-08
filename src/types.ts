export type TransactionRow = Record<string, string | number>;

/** Stable id for a row within a loaded file */
export interface DataRow {
  id: string;
  fileId: string;
  sheetName: string;
  rowIndex: number;
  row: TransactionRow;
  /** Short label for tree leaf */
  label: string;
}

export interface SheetModel {
  sheetName: string;
  headers: string[];
  rows: TransactionRow[];
  dataRows: DataRow[];
}

export type WorkbookKind = "financial" | "coa" | "bank" | "generic";

export interface LoadedWorkbook {
  id: string;
  fileName: string;
  baseName: string;
  kind: WorkbookKind;
  sheets: SheetModel[];
}

/** Tree node for UI */
export type TreeNode =
  | {
      type: "file";
      id: string;
      label: string;
      children: TreeNode[];
    }
  | {
      type: "sheet";
      id: string;
      label: string;
      children: TreeNode[];
    }
  | {
      type: "group";
      id: string;
      label: string;
      children: TreeNode[];
    }
  | {
      type: "leaf";
      id: string;
      label: string;
      rowRef: DataRow;
    };
