<div align="center">

<img src="./build/icon.png" width="140" height="140" alt="MeeBI Logo" />

# MeeBI 2.0 
### **The Unified Offline Analytics Super-App**
*Spreadsheets • Relational Models • DAX Measures • Interactive BI • DuckDB WASM • Encrypted Formulas*

[![Platform](https://img.shields.io/badge/Platform-Windows%20Desktop%20%7C%20Web-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/Meeketank/MeeBI)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.2-3178C6?style=for-the-badge&logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![DuckDB](https://img.shields.io/badge/DuckDB-WASM%20Engine-FFF000?style=for-the-badge&logo=duckdb&logoColor=black)](https://duckdb.org/)
[![Privacy](https://img.shields.io/badge/Privacy-100%25%20Offline%20Local-22c55e?style=for-the-badge&logo=shield&logoColor=white)](https://github.com/Meeketank/MeeBI)
[![Formulas](https://img.shields.io/badge/.tank-Password%20Encrypted%20Formulas-7c3aed?style=for-the-badge&logo=lock&logoColor=white)](https://github.com/Meeketank/MeeBI)
[![License](https://img.shields.io/badge/License-Proprietary-blue?style=for-the-badge)](https://github.com/Meeketank/MeeBI)

<br/>

**MeeBI fuses the cell-grid calculation agility of Microsoft Excel, the relational modeling power of Microsoft Power BI, the visual elegance of Tableau, and the blazing SQL speed of DuckDB — inside a single, zero-cloud, 100% private, native application.**

[Download Desktop](#-desktop-application) • [The 5-Way Comparison](#-the-ultimate-analytics-matrix-5-way-comparison) • [Core Pillars](#-core-architectural-pillars) • [MeeLabs & Plug-ins](#-meelabs-50-free-add-in-models--a-plug-in-api) • [Custom Formulas & .tank](#-custom-multi-sheet-formulas--tank-formula-files) • [Executive Dossier](#-consolidated-executive-dossier--print-engine) • [Getting Started](#-quick-start)

---

</div>

## 🌟 The Philosophy: Why MeeBI?

Modern enterprise business intelligence is broken across fragmented tools:
* **Microsoft Excel**: Unmatched for rapid cell-level formulas, exploratory models, and ad-hoc financial schedules — but collapses under multi-table relational schemas, high row volumes, and interactive cross-filtering.
* **Google Sheets**: Convenient for online sharing, but crippled by 10-million cell hard limits, slow cloud recalculation, lack of DAX modeling, and extreme data privacy exposure.
* **Microsoft Power BI**: Superb for relational data modeling and DAX measures — but forces users into read-only grids, expensive monthly cloud subscriptions, heavy desktop installers, and mandatory cloud tenant governance.
* **Tableau**: Visually rich — but notoriously expensive ($75+/user/mo), complex for simple cell math, and detached from live spreadsheet workflows.

### **The MeeBI Synthesis**
A dataset in MeeBI is **simultaneously a living spreadsheet grid, an indexed database table, and an interactive business intelligence visual**. 
Change a number on a sheet cell, and every linked formula, relational model, DAX measure, pivot summary, and dashboard visual across your entire workbook recalculates instantly in client memory — with **zero cloud dependencies, zero telemetry, and 100% offline data sovereignty**.

---

## ⚡ The Ultimate Analytics Matrix: 5-Way Comparison

| Dimension / Capability | MeeBI 2.0 | Microsoft Power BI | Tableau Desktop | Microsoft Excel 365 | Google Sheets |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Native Cell-Grid Formulas** | **✅ Hundreds of Excel-compatible Functions & Spills** | ❌ Read-only Data View | ❌ Read-only Data View | ✅ Full Worksheet Grid | ✅ Full Worksheet Grid |
| **Relational Data Modeling** | **✅ Visual Star/Snowflake ERD** | ✅ Dedicated Model Tab | ⚠️ Relationships Only | ⚠️ Power Pivot (Separate) | ❌ None |
| **Multi-Table Natural-Language Ask** | **✅ Links tables, splits by fields from different tables, flags missing data** | ⚠️ Cloud Q&A | ⚠️ Cloud Tableau Pulse | ❌ Cloud Copilot ($30/mo) | ⚠️ Cloud Gemini |
| **DAX Measures & Context Transitions** | **✅ Native In-Memory Engine** | ✅ Native DAX Engine | ❌ LOD Expressions (Non-DAX) | ⚠️ Limited in Power Pivot | ❌ None |
| **Interactive Canvas Cross-Filtering** | **✅ 50+ Visual Types, drag-to-shelf exploration** | ✅ Rich Canvas | ✅ Rich Canvas | ❌ Static Charts | ❌ Static Charts |
| **Chart Analytics (anomaly detection, forecast)** | **✅ Find Anomalies, trend/forecast, bump charts** | ⚠️ Premium add-on | ⚠️ Extension required | ❌ None | ❌ None |
| **Operations Research & Statistics Add-ins** | **✅ MeeQuant Lab + MeeStats Pack: 50+ free models (simplex, MRP, ANOVA, regression, and more)** | ❌ Not built in | ❌ Not built in | ❌ Not built in | ❌ Not built in |
| **Extensibility: Plug-in API & Gallery** | **✅ Your own worksheet functions, MeeLabs models and sandboxed report visuals, plus a ready-made gallery** | ⚠️ Paid AppSource only | ⚠️ Paid Extension Gallery | ⚠️ VBA/Office Scripts, no sandbox | ⚠️ Apps Script, cloud-only |
| **Offline Privacy & Zero Cloud Leak** | **✅ 100% Client Memory** | ⚠️ Requires Cloud Service | ⚠️ Cloud / Server Tied | ⚠️ Cloud Sync Pushed | ❌ 100% Cloud / No Offline |
| **Local Big Data SQL Engine** | **✅ DuckDB WASM in-memory** | ❌ VertiPaq (Proprietary) | ❌ Hyper (Proprietary) | ❌ None (1M row limit) | ❌ None (Cloud query) |
| **Autopilot: One-Click Everything** | **✅ Clean + Charts + Pivots** | ❌ Manual Build | ❌ Manual Build | ⚠️ Excel Ideas (Cloud) | ⚠️ Explore (Basic) |
| **Custom Cross-Sheet Formulas** | **✅ `=MyROI({S1[A],R},{S2[B],R})`** | ❌ Not Supported | ❌ Not Supported | ⚠️ Fragile Cell Chains | ⚠️ Fragile Cell Chains |
| **Portable Formulas** | **✅ `.tank` files, AES-256 encrypted with a password** | ❌ Plain Text / PBIX | ❌ XML / Plain Text | ❌ Plain Text VBA | ❌ Plain AppScript |
| **Executive Print / PDF Dossier** | **✅ Full-Bleed Landscape Batch** | ⚠️ Rigid Export / Cutoff | ⚠️ Complex Dashboard Export | ⚠️ Manual Page Setup | ⚠️ Basic Print |
| **Application Footprint & Boot** | **✅ ~5 MB Installer, ~1s Boot** | ❌ ~800 MB, Heavy Boot | ❌ ~1.2 GB, Heavy Boot | ❌ Heavy Office Suite | ⚠️ Browser-tied |
| **Licensing & Cloud Subscription** | **✅ Free / Private / Local** | ❌ $10–$20 / user / mo | ❌ $75 / user / mo | ❌ Microsoft 365 Sub | ❌ Google Workspace |

This table is our own comparison, judged against each tool's own documentation and public
pricing, not an independent benchmark — see [Honest Limits](#-honest-limits) below for what
is still missing or unverified.

---

## 🏛️ Core Architectural Pillars

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                  MeeBI PLATFORM                                        │
├──────────────────┬──────────────────┬──────────────────┬───────────────────────────────┤
│ 📊 SPREADSHEETS  │ 🧩 DATA MODEL    │ 📐 DAX MEASURES  │ 📈 BI VISUALIZATIONS          │
│ • 475+ Functions │ • Visual ERD     │ • CALCULATE      │ • 16+ Interactive Visuals     │
│ • Dynamic Arrays │ • Star/Snowflake │ • TOTALYTD / MTD │ • Bi-directional Cross-filter │
│ • Flash Fill     │ • Auto-Relate    │ • Time Intel     │ • Smart Slicers & Timelines   │
│ • Spill Engine   │ • Data Doctor    │ • Calc Groups    │ • Auto-batched Paper Layout   │
├──────────────────┴──────────────────┴──────────────────┴───────────────────────────────┤
│ 🚀 LOCAL HIGH-PERFORMANCE RUNTIME                                                      │
│ • DuckDB WebAssembly: In-memory columnar SQL for millions of rows                      │
│ • Pyodide Python 3.14: In-sandbox Pandas, NumPy, Scipy, Matplotlib                     │
│ • Rust / Tauri Native Shell: WebView2 Windows executable with zero browser overhead    │
│ • `.tank` formula files: AES-256-GCM encrypted with a password (scrambled without one) │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 🧪 MeeLabs: 50+ Free Add-In Models, & a Plug-in API

**MeeLabs** (`Ctrl+Shift+P` → the model's name, or the flask icon on the Ribbon) bundles two
libraries of models that other tools either lack entirely or sell as separate add-ons:

* **MeeQuant Lab** (operations research): the simplex method shown step by step with every
  tableau, aggregate production planning, MRP with lot sizing (including Wagner-Whitin),
  goal programming, process layout, time study and work sampling, break-even, ABC analysis,
  NPV/IRR/payback, and more.
* **MeeStats Pack** (statistics): t-tests and z-tests, one-way and two-way ANOVA, regression,
  chi-square, correlation, Mann-Whitney, Wilcoxon, Kruskal-Wallis, logistic regression,
  Tukey's HSD, Levene's test, power and sample size, principal component analysis, and more —
  checked in the test suite against SciPy and statsmodels.

Each model takes plain inputs and explains its result in words, not just a number.

### Plug-in API & Gallery

When a function, model or chart isn't built in, a **plug-in** — one JavaScript file calling
`meebi.plugin({ functions, labs, visuals })` — can add it. Install one from
**MeeLabs → Plug-ins**: MeeBI asks for confirmation first, keeps plug-ins per browser/install,
and draws a plug-in's report visuals in a sandboxed frame that cannot run scripts or reach the
network. A built-in **Gallery** inside that dialog offers a few ready-made plug-ins — a
finance toolkit (CAGR, Altman Z-Score), a process-capability model, and a project-management
toolkit (cost and schedule performance, estimate at completion) — installed with one click.
See [docs/plugins.md](docs/plugins.md) for the format.

---

## 🔮 Custom Multi-Sheet Formulas & `.tank` Formula Files

MeeBI introduces a dedicated **Custom Cross-Sheet Formula Engine** accessible directly inside the **Data section**:

### 1. Multi-Sheet Cross-Table Syntax
Effortlessly define complex inter-sheet metrics without writing fragile `VLOOKUP` or `INDEX/MATCH` chains:
```excel
=MyROI({sheet1[Cash], R}, {sheet2[Debt], R})
```
* **Sheet / Table Independence**: Reference columns across disparate tables or within the same table.
* **Selection Flexibility**:
  * `R`: Evaluates across the entire data row (full column-to-column vectorized computation).
  * `A1:A17`: Slices calculations to explicit row ranges.
  * `A1, A7, A10`: Selects discrete cell/row indices for targeted calculations.
* **Built-in Operations**:
  * `divide / ratio`: Evaluates ratios (e.g. Cash / Debt, ROI, Profit Margin).
  * `add / sum`: Sums cross-sheet vectors.
  * `subtract`: Variance and delta calculations.
  * `multiply`: Compounding, tax, and weighted calculations.
  * `average`: Cross-sheet arithmetic means.
  * `percentage (%)` & `growth / diff (%)`: Automated rate computations.
  * `min` / `max`: Boundary evaluations.
  * `delete / clear`: Automated masking of matching rows.

### 2. `.tank` formula files: encrypted with a password, or scrambled without one
To share formulas between computers, export them as a `.tank` file:

* **With a password (format 2): encrypted.** AES-256-GCM, with the key derived from your password by PBKDF2-SHA-256 (600,000 rounds), a random salt and nonce for every file, and the header bound to the content. Without the password nobody can read the formulas, MeeBI included, and a file that was changed is refused. A lost password cannot be recovered.
* **Without a password (format 1): scrambled, not encrypted.** The bytes are mixed with a key built into MeeBI, so the file does not read as text in Notepad or Word. Anyone with MeeBI (or this source code) can open it, so do not rely on it for anything confidential.
* **Import** reads both formats; an encrypted file asks for its password. Files exported by older MeeBI versions are format 1 and still open.

Workbooks (**File > Save copy with a password**) and backups (**Backups > Back up with a password**) use the same encryption. See [docs/SECURITY.md](docs/SECURITY.md) for what MeeBI's security does and does not cover.

---

## ⚡ Autopilot: One-Click Complete Analytics ("Do Everything")

Accessible from the Ribbon's **Do Everything** button (`Rocket` icon) or Command Palette (`Ctrl+P` ➔ *Autopilot*):

1. **Automated Data Cleaning**: Scans tables using the Data Doctor, trimming whitespace, standardizing regional text variants, converting numbers masked as text, removing blank rows, and deduplicating records.
2. **Optimal Dashboard Generation**: Recommends and constructs up to 3 interactive dashboard pages tailored to the categorical, numerical, and temporal dimensions of your data.
3. **Automated Analytical Pivot Tables**:
   * Evaluates key metrics across primary entities and buckets.
   * Generates cross-tabs, shares of row totals (`% of row`), timeline breakdowns, and leaderboard summaries.
   * Automatically commits them into the workbook's **Pivots tab**.
4. **Smart Narrative & Outlier Detection**: Synthesizes natural-language business insights that highlight top performers, variances, and correlations.

---

## 📑 Consolidated Executive Dossier & Print Engine

Generating board-ready executive reports from interactive BI dashboards has traditionally resulted in truncated charts, awkward margins, and cut-off tables. MeeBI solves this with the **Consolidated Report Dossier**:

* **Full-Bleed Patch Headers & Footers**: The top header (`MeeBI Analytics Dossier`) and bottom blue footer (`Page X`) span 100% of the page width with **zero white margin bleed** on the sides.
* **Strict Footer Pinning**: Footers are pinned strictly to the bottom of each page (`marginTop: auto`), eliminating accidental blank overflow pages.
* **Automated Landscape Visual Batching**:
  * Visuals are restructured specifically for the paper format (A4 / Letter landscape), ignoring arbitrary dashboard grid coordinates.
  * KPI metric cards are automatically grouped into clean 4-column rows (`grid-cols-4`, height ~100px).
  * Paired with up to 2 charts side-by-side (`h-[360px]`), or standalone charts (`h-[480px]` / `h-[500px]`).
* **Zero Scrollbars**: Tables and matrices render with clean row slicing and overflow suppression (`forPrint` mode), ensuring all data is visible without scrollbars.
* **Clean Page Preview**: Centered pages with authentic paper shadows and zero stray unstyled page numbers in corners.
* **Multi-Format Export**: One-click generation of print dialogs, **vector PDF**, and formatted **Microsoft Word (`.docx`)**.

---

## 📁 File Format Compatibility

| Extension | Open / Import | Save / Export | Technical Details |
| :--- | :---: | :---: | :--- |
| **`.meebi`** | ✅ Full | ✅ Full | Native binary bundle preserving sheets, tables, relationships, DAX measures, custom formulas, pivots, and report canvases. |
| **`.tank`** | ✅ Full | ✅ Full | Proprietary encrypted binary container for confidential custom cross-sheet formula export & portability. |
| **`.xlsx` / `.xlsm`** | ✅ Full | ✅ Full | Imports formulas, styles, merges, validations, named ranges, and frozen panes. Exports sheets with live formulas and charts. |
| **`.csv` / `.tsv`** | ✅ Full | ✅ Full | Auto-detects delimiters, date formats (DMY vs MDY), and numeric conventions with Rust SIMD acceleration. |
| **`.ods`** | ✅ Full | ⚠️ (via XLSX) | OpenDocument spreadsheet format parsing. |

---

## 🚀 Quick Start

**First time opening MeeBI?** [docs/getting-started.md](docs/getting-started.md) walks through
opening or creating a workbook, the grid, asking questions in plain English, data modeling,
DAX measures, drag-and-drop reports, MeeLabs, plug-ins, macros and privacy — written for
someone who has never used MeeBI (or a BI tool) before.

### Option 1: Native Windows Desktop App (Recommended)
1. Download **`MeeBI-Setup.exe`** from the [latest release](https://github.com/Meeketank/MeeBI-desktop/releases/latest).
2. Run the installer (supports per-user installation without administrator privileges).
3. Open any `.xlsx`, `.csv`, `.meebi`, or `.tank` file instantly.

### Option 2: Development & Source Build
Requires [Node.js](https://nodejs.org/) 20+ and [Rust](https://rustup.rs):

```bash
# Clone the repository
git clone https://github.com/Meeketank/MeeBI.git
cd MeeBI

# Install dependencies with locked versions
npm ci

# Launch local development server with Hot Module Replacement
npm run dev

# Or launch the native desktop development environment
npm run app
```

### Build Production Artifacts:
```bash
# Type check and build web production bundle
npm run build

# Package native Windows NSIS installer (.exe)
npm run dist:win
```

---

## 🏗️ Project Architecture

```
MeeBI/
├── src-tauri/              # Native Windows desktop application (Rust + Tauri)
│   ├── src/lib.rs          # Single-instance lifecycle, file associations, window handlers
│   ├── src/files.rs        # Native Save As / Open dialogs
│   ├── src/updates.rs      # Cryptographically signed silent updater
│   └── tauri.conf.json     # App manifest, bundle configurations, and security policies
├── src/
│   ├── engine/             # Headless TypeScript calculation core (Zero DOM dependency)
│   │   ├── formula/        # Lexer, AST parser, and Excel-compliant worksheet functions
│   │   ├── labs/           # MeeLabs: MeeQuant Lab (OR) and MeeStats Pack (statistics) models
│   │   ├── plugins.ts      # Plug-in API, sandboxed execution, and the built-in Gallery
│   │   ├── assistant/       # Ask: plain-English Q&A and sheet-edit commands
│   │   ├── calc.ts         # Dependency DAG, topological sort, circular detection, dynamic spills
│   │   ├── customFormula.ts# Custom cross-sheet formula parser, evaluator & .tank binary cipher
│   │   ├── autopivot.ts    # Automated analytical pivot recommendation & cross-tab engine
│   │   ├── model.ts        # Relational schema engine, relationships, DAX measures
│   │   └── format.ts       # Excel-compatible number formatting & internationalization
│   ├── macros/             # Plain-English → macro step translator, and the macro runtime
│   ├── io/                 # XLSX, CSV, ODS, .tank, and .meebi binary serialisation
│   ├── store/              # Zustand state stores with immutable undo/redo history
│   └── ui/                 # High-performance interface components
│       ├── grid/           # Virtualized cell grid with auto-fit layout & print preparation
│       ├── data/           # Data modeling, table schemas, and CustomFormulaManager
│       ├── export/         # ConsolidatedReportDoc: full-bleed headers/footers & visual batching
│       ├── pivot/          # Interactive pivot table builder & aggregation matrix
│       ├── labs/           # MeeLabs hub and the Plug-in Manager (install, Gallery, switch on/off)
│       └── report/         # Drag-and-drop dashboard canvas & bi-directional cross-filtering
├── docs/                   # getting-started.md, plugins.md, SECURITY.md, and more
└── build/                  # Icon assets and continuous deployment workflows
```

---

## ⚖️ Honest Limits

MeeBI is under active development, and we would rather say plainly what is still ahead than
overstate it:

* **No live collaboration or cloud sharing yet.** Everything runs locally; sharing a workbook
  today means sending the file itself.
* **Ask's AI integration can need a server-side key** for its richest natural-language
  features; the fully offline path is less exercised than the rest of the app.
* **Export fidelity against real Excel and Power BI Desktop**: function behaviour is checked
  against Excel's own documented results (298 cases in the test suite), and exported `.xlsx`
  files are verified to open and recompute correctly in LibreOffice Calc, a separate,
  standards-based OOXML implementation — but they still have not been opened in Microsoft
  Excel or Power BI Desktop themselves.
* **The plug-in Gallery has a handful of plug-ins today**, not an open marketplace with outside authors.
* Some plain-English sheet-edit phrasings (for example, asking Ask to add conditional
  formatting beyond a simple highlight) are not yet understood; they are logged as known gaps
  as we find them, not claimed as working.

We track this candidly in dated rating notes as the app develops, rather than only in marketing
copy — ask in an issue if you want specifics on any capability before relying on it.

---

## 🔒 Privacy & Data Sovereignty

* **100% Client-Side Compute**: All formula calculation, relational data modeling, DuckDB SQL queries, and visual rendering execute strictly in local computer memory.
* **Zero External Telemetry**: MeeBI does not send your data, table names, financial numbers, or queries across the internet.
* **Confidential Formula Security**: Custom formulas exported via `.tank` are obfuscated using multi-round binary encryption, ensuring proprietary business logic remains confidential.

---

<div align="center">

**Developed with precision by [Meeket Tank](https://github.com/Meeketank)**  
*For inquiries, feature requests, or enterprise collaboration, connect on [LinkedIn](https://linkedin.com/in/meeketank) or open an issue.*

</div>

## Map data

State and province outlines come from [Natural Earth](https://www.naturalearthdata.com/) (public domain). City locations come from [GeoNames](https://www.geonames.org/) (CC BY 4.0), by way of the `all-the-cities` package. Both are built into the app by `scripts/build-geo.mjs`, so maps work offline.
