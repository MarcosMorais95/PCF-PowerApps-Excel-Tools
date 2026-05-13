<img width="599" height="425" alt="image" src="https://github.com/user-attachments/assets/d934fe93-d7bc-46d5-8f95-91e923cf5fa4" /># Excel Tools PCF

> Two production-grade PowerApps Component Framework (PCF) controls that bring real Excel import/export to Canvas Apps — packaged in a single Dataverse solution.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Power Platform](https://img.shields.io/badge/Power%20Platform-PCF-742774?logo=powerplatform&logoColor=white)](https://learn.microsoft.com/power-apps/developer/component-framework/overview)
[![SheetJS](https://img.shields.io/badge/SheetJS-0.18-217346)](https://sheetjs.com/)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Marcos%20Morais-0A66C2?logo=linkedin&logoColor=white)](https://www.linkedin.com/in/marcosmorais95/)

---

## Why this exists

Canvas Apps cannot natively read or write `.xlsx` files. The usual workaround is a Power Automate flow with the Excel Online connector, which:
- Requires the file to live in OneDrive / SharePoint as a **table** before it can be queried.
- Adds latency, throttling, and an extra license footprint.
- Cannot trigger a real client-side download.

These two controls solve the problem entirely on the client: the user picks a file (or clicks a button) and the data flows directly into / out of the app — no flow, no connector, no round trip.

<img width="599" height="425" alt="image" src="https://github.com/user-attachments/assets/75502563-b4e8-433d-92ea-98c100a0df86" />

---

## Table of contents

- [The controls](#the-controls)
  - [PCFImportExcel — Excel/CSV → JSON](#pcfimportexcel--excelcsv--json)
  - [PCFExportToExcel — JSON → .xlsx](#pcfexporttoexcel--json--xlsx)
- [Architecture](#architecture)
- [Project structure](#project-structure)
- [Download](#download)
- [Getting started](#getting-started)
- [Build and deploy](#build-and-deploy)
- [License](#license)

---

## The controls

### PCFImportExcel — Excel/CSV → JSON

Lets the user pick one or more Excel/CSV files in the browser, parses them **automatically on selection**, and exposes the data as a JSON string the app can bind to a collection. The control fires `OnChange` as soon as parsing finishes — no extra button click needed.

**Two display modes** (selectable via the `displayMode` property):
- **Card** — rich UI with title, guidelines, drop-zone, drag-and-drop, status panel and file chip with parsed-row stats.
- **Button** — compact single-button variant for tight layouts. Includes loading spinner and transient success state.

**Highlights**
- **Auto-parse on selection** — no manual "Parse" click; `OnChange` fires when data is ready.
- **Drag-and-drop** in Card mode.
- **Two parsing modes:**
  - **Named table** — locates an Excel table by name across the workbook (works even if the table sits on a different sheet from the active one).
  - **Range** — auto-detects the header row using a scoring heuristic (uniqueness, non-numeric content, density of the next row).
- **Multi-file** support with a per-row `fileName` tag.
- **Themeable** — `accentColor`, `backgroundColor`, `borderColor` let the control fit any app theme.
- **Trace log** output for debugging — every step (file read, sheet selection, header detection, row extraction) is timestamped and labeled.
- **Strict outputs**: `jsonResult`, `meta`, `trace`, `isValid`, `errorMessage` — so the host app can react granularly.

**Inputs**

| Property | Type | Default | Description |
|---|---|---|---|
| `displayMode` | Enum | `Card` | `Card` (rich UI) or `Button` (compact) |
| `hasTable` | TwoOptions | `false` | Parse a named Excel table instead of a worksheet range |
| `tableName` | Text | — | Name of the Excel table (required when `hasTable = true`) |
| `sheetIndex` | Whole | `0` | Zero-based worksheet index (range mode only) |
| `allowMultipleFiles` | TwoOptions | `false` | Allow the user to select more than one file at a time |
| `maxRowsToScan` | Whole | `50` | Rows to scan when detecting the header (5–200) |
| `includeFileName` | TwoOptions | `false` | Include the source fileName on each parsed row |
| `enableTrace` | TwoOptions | `false` | Enable detailed trace logging |
| `resetSignal` | TwoOptions | — | Toggle this value to reset the control and clear all selected files |
| `showTitle` / `title` | TwoOptions / Text | `true` / — | Show and customize the component title (Card mode) |
| `showUserGuidelines` / `userGuidelinesText` | TwoOptions / Multiple | `true` / — | Show usage instructions (Card mode) |
| `showStatus` | TwoOptions | `true` | Show the status panel (Card mode) |
| `dropzonePrimaryText` / `dropzoneHintText` | Text | — | Localize the dropzone labels (Card mode) |
| `buttonText` | Text | `Import` | Label for the import button (Button mode) |
| `removeButtonText` | Text | `Remove` | Label for the remove-file button on the file chip (Card mode) |
| `accentColor` | Text | `#323130` | Interactive accent color (hex) |
| `backgroundColor` | Text | `#ffffff` | Card / button background (hex) |
| `borderColor` | Text | `#e1dfdd` | Card / button border (hex) |

**Outputs**

| Property | Type | Description |
|---|---|---|
| `jsonResult` | Multiple | Parsed data serialized as a JSON array |
| `meta` | Multiple | Metadata about the parsing run (mode, sheet, headers, totals) |
| `trace` | Multiple | Detailed trace log when `enableTrace = true` |
| `isValid` | TwoOptions | `true` when the last parsing attempt succeeded |
| `errorMessage` | Text | Human-readable error message when parsing fails |

**Sample Power Fx**

```powerfx
// In the parent screen's OnVisible (or a button OnSelect),
// react to the import result and load it into a collection:
If(
    ImportExcel1.isValid,
    ClearCollect(
        colImported,
        ForAll(
            Table(ParseJSON(ImportExcel1.jsonResult)),
            {
                Name:  Text(ThisRecord.Value.Name),
                Email: Text(ThisRecord.Value.Email),
                Age:   Value(ThisRecord.Value.Age)
            }
        )
    ),
    Notify(ImportExcel1.errorMessage, NotificationType.Error)
)
```

---

### PCFExportToExcel — JSON → .xlsx

A styled button that converts a JSON payload into a real `.xlsx` workbook and triggers the browser download — no connector, no Power Automate.

**Highlights**
- **Flexible input shapes** — array of objects, array of arrays, single object, primitive, or a mapped-columns payload (`{ columns: [...], rows: [...] }`) for friendly headers + explicit field keys.
- **Auto-filter** enabled on the exported sheet by default.
- **Fully styled** — width, height, font size, colors, border radius, icon visibility — all bindable from Canvas.
- **Smart fileNames** — auto-generates a timestamped name if none is provided; ensures the `.xlsx` extension.

**Inputs**

| Property | Type | Default | Description |
|---|---|---|---|
| `dataJson` | Multiple (required) | — | JSON payload to export |
| `useFieldMapping` | TwoOptions | `false` | Use mapped-columns mode (`columns` + `rows`) |
| `fileName` | Text | — | Output file name (auto-generates a timestamp if empty) |
| `sheetName` | Text | `Export` | Worksheet name |
| `buttonText` | Text | `Export Excel` | Button label |
| `showIcon` | TwoOptions | `true` | Show the inline Excel SVG icon |
| `isDisabled` | TwoOptions | `false` | Disable the button |
| `buttonWidthPx` / `buttonHeightPx` | Whole | `120` / `32` | Button dimensions |
| `fontSizePx` | Whole | `12` | Button font size |
| `backgroundColor` / `fontColor` | Text | `#0078d4` / `#ffffff` | Button colors (hex) |
| `borderRadiusPx` | Whole | `6` | Corner radius |

**Sample Power Fx — array of objects**

```powerfx
ExportToExcel1.dataJson  = JSON(colMyData, JSONFormat.IgnoreBinaryData);
ExportToExcel1.fileName  = "Customers";
ExportToExcel1.sheetName = "Customer list";
```

**Sample Power Fx — mapped columns (friendly headers)**

```powerfx
Set(
    varExportPayload,
    JSON(
        {
            columns: [
                { name: "Customer name", field: "CustomerName" },
                { name: "Total spend",   field: "Total" }
            ],
            rows: ForAll(
                colMyData,
                { CustomerName: Customer, Total: TotalSpend }
            )
        },
        JSONFormat.IgnoreBinaryData
    )
);
ExportToExcel1.useFieldMapping = true;
ExportToExcel1.dataJson        = varExportPayload;
```

---

## Architecture

```
┌─────────────────────┐                ┌──────────────────────┐
│   Canvas App        │   inputs  ─►   │                      │
│                     │                │   PCF control        │
│                     │   ◄─ outputs   │   (TypeScript)       │
└─────────────────────┘                └──────────┬───────────┘
                                                  │
                                                  ▼
                                       ┌──────────────────────┐
                                       │   SheetJS (xlsx)     │
                                       │   read / write       │
                                       └──────────┬───────────┘
                                                  │
                                                  ▼
                                       ┌──────────────────────┐
                                       │   Browser            │
                                       │   File picker /      │
                                       │   Blob download      │
                                       └──────────────────────┘
```

Both controls run **entirely client-side**. No data leaves the browser; no Power Automate flow or Excel Online connector is involved.

---

## Project structure

```
PCF-Impor&Export/
├── PCFImportExcel/                   # Import PCF project (.pcfproj)
│   ├── ImportExcel/
│   │   ├── ControlManifest.Input.xml
│   │   ├── index.ts
│   │   └── strings/
│   │       └── ImportExcel.1033.resx
│   ├── PCFImportExcel.pcfproj
│   └── package.json
│
├── PCFExportToExcel/                 # Export PCF project (.pcfproj)
│   ├── ExportToExcelComponent/
│   │   ├── ControlManifest.Input.xml
│   │   ├── index.ts
│   │   └── strings/
│   │       └── ExportToExcelComponent.1033.resx
│   ├── PCFExportToExcel.pcfproj
│   └── package.json
│
├── Solution/                         # Dataverse solution project (.cdsproj)
│   ├── Solution.cdsproj              # references both PCFs above
│   └── src/Other/Solution.xml        # Excel Tools PCF, publisher: marcosmorais
│
├── .gitignore
├── LICENSE                           # MIT
└── README.md
```

The `Solution/` project is a standard Dataverse solution that references the two `.pcfproj` files via relative paths. A single `dotnet build` produces one deployable `.zip` containing both controls.

| Solution attribute | Value |
|---|---|
| Unique name | `PCFExcelTools` |
| Display name | Excel Tools PCF |
| Publisher name | `marcosmorais` |
| Customization prefix | `mm` |
| Package type | Unmanaged (Dev) |
| Version | `1.1.0` |

---

## Download

The easiest way to get started is to download the ready-to-import solution from the [**Releases page**](https://github.com/MarcosMorais95/PCF-PowerApps-Excel-Tools/releases).

Each release ships a single `.zip` (`PCFExcelTools_<version>.zip`) that contains both controls. No build step required.

```powershell
pac solution import `
  --path PCFExcelTools_<version>.zip `
  --activate-plugins `
  --force-overwrite `
  --publish-changes
```

> See [Add a control to a Canvas App](#add-a-control-to-a-canvas-app) for the next step after importing.

---

## Getting started

### Prerequisites
- [Node.js](https://nodejs.org/) 16 or later
- [Power Platform CLI](https://aka.ms/PowerAppsCLI) (`pac`)
- [.NET SDK](https://dotnet.microsoft.com/download) (for the solution build)

### Install dependencies

```powershell
cd PCFImportExcel
npm install

cd ..\PCFExportToExcel
npm install
```

### Run a control in the test harness

```powershell
# Import control
cd PCFImportExcel
npm start

# Export control
cd PCFExportToExcel
npm start
```

The test harness opens at `http://localhost:8181/` with a sandboxed Canvas-like environment for binding inputs and inspecting outputs.

---

## Build and deploy

### Build the consolidated solution

```powershell
cd Solution
dotnet build --configuration Release
# Output: Solution\bin\Release\PCFExcelTools_<version>.zip
# (version is read from Solution\src\Other\Solution.xml at build time)
```

This single `.zip` contains **both** controls.

### Authenticate to your environment

```powershell
pac auth create --url https://<your-org>.crm.dynamics.com --name PCF-Dev
pac auth select --name PCF-Dev
```

### Import to Dataverse

```powershell
pac solution import `
  --path Solution\bin\Release\PCFExcelTools_<version>.zip `
  --activate-plugins `
  --force-overwrite `
  --publish-changes
```

> **Important:** `--publish-changes` is mandatory. Without it, Dataverse keeps the previous bundle published in Canvas, even after a successful import.

### Add a control to a Canvas App

1. In the maker portal, open your app → **Insert** → **Get more components** → **Code** tab.
2. Pick **Import Excel** and/or **Export to Excel** from the **Excel Tools PCF** solution.
3. Drop the control onto a screen and bind its properties via the formula bar.

### Test file

A sample file [`SampleData.xlsx`](SampleData.xlsx) is included in the repository. It contains **100 records** on sheet `tab01` with the same columns used in the Power Fx samples above (`Name`, `Email`, `Age`), so you can drop it straight into the import control and run the code examples without any changes.

### Version bumps before re-deploy

Every redeploy requires bumping **two** versions:

1. The control's `ControlManifest.Input.xml` → `<control version="...">`
2. The solution's `Solution/src/Other/Solution.xml` → `<Version>...</Version>`

If you skip this, Dataverse imports the package successfully but Canvas keeps loading the old bundle.

---

## License

[MIT](LICENSE) © 2026 [Marcos Morais](https://www.linkedin.com/in/marcosmorais95/)
