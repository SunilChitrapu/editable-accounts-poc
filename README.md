# Accounting Workbook Drilldown Editor (React + TypeScript)

React app to:

- Import Excel workbook with required sheets: `TB`, `P&L`, `BS`
- Drill down in a tree: Sheet -> Account -> Transaction
- Edit transaction fields
- Export updated workbook to a new `.xlsx` file

## Run

```bash
npm install
npm run dev
```

Open the local Vite URL shown in terminal.

## Deploy to GitHub Pages (free)

1. Create a GitHub repo (example repo name: `accounting-drilldown-editor`)
2. In this project, edit `.github/workflows/deploy-pages.yml` and replace:
   - `VITE_BASE: "/REPO_NAME_HERE/"`
   with:
   - `VITE_BASE: "/<your-repo-name>/"`
   Example:
   - `VITE_BASE: "/accounting-drilldown-editor/"`
3. Push to GitHub on the `main` branch.
4. In GitHub: **Settings → Pages**
   - **Source**: `GitHub Actions`
5. After the workflow runs, your site will be at:
   - `https://<your-username>.github.io/<your-repo-name>/`

## Build

```bash
npm run build
npm run preview
```

## Usage

1. Click **Import Excel** and choose your workbook.
2. Expand the tree and select a transaction leaf.
3. Edit values in the transaction editor and click **Save Transaction**.
4. Click **Export Updated Excel** to download the updated workbook.

## Notes

- Column mapping auto-detects common names (`Account`, `Description`, `Amount`, etc.).
- Even with custom column names, rows remain fully editable and export correctly.
