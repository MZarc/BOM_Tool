<div align="center">

  <img src="public/icon.png" alt="BOM Tool" width="128" />

  # BOM Tool

  **Desktop BOM analysis, comparison and Pick-and-Place validation.**
  Local-first. No uploads. No telemetry.

  [![Version](https://img.shields.io/badge/version-1.2.5-22d3ee?style=flat-square)](package.json)
  [![Electron](https://img.shields.io/badge/Electron-30-47848F?style=flat-square&logo=electron&logoColor=white)](https://www.electronjs.org/)
  [![React](https://img.shields.io/badge/React-18-61DAFB?style=flat-square&logo=react&logoColor=black)](https://react.dev/)
  [![TypeScript](https://img.shields.io/badge/TypeScript-5-3178C6?style=flat-square&logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
  [![Vite](https://img.shields.io/badge/Vite-5-646CFF?style=flat-square&logo=vite&logoColor=white)](https://vitejs.dev/)
  [![License](https://img.shields.io/badge/license-MIT-blue?style=flat-square)](LICENSE)

  [![Windows](https://img.shields.io/badge/platform-Windows%20x64-0078D6?style=flat-square&logo=windows&logoColor=white)](#build--release)

</div>

---

## What it does

**BOM Tool** is an Electron desktop app for hardware and SMT engineers who need to
get a bill of materials from a spreadsheet into a trustworthy state — and keep it
that way between revisions.

Open a workbook, let the app find the header row and map the columns, then edit,
normalize and validate in place. Export writes back into the **original workbook**
with ExcelJS, so fonts, fills, merged cells, logos and template layout survive the
round trip.

| Workspace | What it answers |
| :-- | :-- |
| 🗂 **BOM Analyzer** | Is this single BOM internally consistent? |
| ⇄ **Comparator** | What changed between the reference and the new revision? |
| 📍 **PnP Validator** | Does this pick-and-place file match the BOM? |

Everything runs locally in the renderer. Files are read into memory and never leave
the machine.

---

## Workspaces

### 🗂 BOM Analyzer

Load one BOM (`.xlsx`, `.xls`, `.csv`) and clean it up before it goes anywhere else.

- **Header detection** — scans the first rows for header keywords so title blocks and
  logos above the table don't break the import.
- **Automatic column mapping** — recognizes common aliases for MPN, quantity,
  designators and description; falls back to an interactive mapping dialog, and lets
  you re-map at any time.
- **Multi-sheet workbooks** — pick the sheet to parse.
- **Editable grid** — edit MPN, qty, description and designators directly. Designator
  edits are split into a clean list and re-checked against the row quantity immediately.
- **History** — every committed edit is a step in the per-file undo/redo stack.
- **Tools**
  - **Normalize** — canonicalizes designator separators/spacing.
  - **Expand** — turns ranges such as `R1-R5` into `R1,R2,R3,R4,R5`.
  - **Validate** — recomputes duplicate designators and quantity mismatches.
- **Per-column filters**, **row search** (`Ctrl/Cmd+F`) and a **status ledger** that
  explains every flagged row.
- **Export** — writes the edited rows back into the original workbook and appends an
  **Analyzer Report** sheet for the last tool run.

### ⇄ Comparator

Drop a reference BOM and a target BOM. The diff recomputes live from the current
grids, so edits made after opening the report are reflected immediately.

Reported categories:

| Category | Meaning |
| :-- | :-- |
| **Missing** | MPN exists in the reference but not in the target |
| **Extra** | MPN exists in the target but not in the reference |
| **Qty mismatch** | Same MPN, different summed quantity across all rows |
| **Swapped** | Same designator is assigned to a different MPN in each file |
| **Duplicate (self)** | The same designator appears more than once within one file |
| **Qty error (self)** | A row's quantity does not equal its designator count |
| **Designator overlap** | Shared MPN whose designator sets differ between files |

Also included: row-level status dots with an explanatory tooltip, **Sync MPN from
reference** for swapped rows, and jump-to-row from the report.

### 📍 PnP Validator

Upload the placement file and the BOM it was generated from.

- **Placement input** — Altium-style `.txt` / `.csv` (tab, comma, semicolon or pipe
  delimited, quoted fields supported) or a spreadsheet (`.xlsx`, `.xls`).
- **BOM input** — the same spreadsheet pipeline as the Analyzer, including header
  detection and column mapping.
- Matching is by **designator** (trimmed, case-insensitive).
- Each designator gets one status:

| Status | Meaning |
| :-- | :-- |
| **Match** | Present in both, and descriptions agree |
| **Mismatch** | Present in both, descriptions differ |
| **Not in BOM** | Placement designator has no BOM entry |
| **Not in PnP** | BOM designator is absent from the placement file |

If the placement file has no description column (or all descriptions are empty), the
description check is skipped and presence alone is validated. Results can be filtered
to mismatches or missing items, and searched.

---

## Export fidelity

Most CSV/XLSX exporters rebuild the sheet from scratch. This one does not:

1. The original file buffer is kept in memory from the moment it is opened.
2. On export, ExcelJS loads that buffer and updates only the mapped cells.
3. Existing styles, merged cells, images and template structure are left untouched.
4. The report sheet (`Analyzer Report` or `Comparison Analysis`) is added or replaced.
5. The workbook is saved through a native save dialog.

If the original buffer is unavailable, the app refuses the export rather than
silently producing a stripped-down file.

---

## Tech stack

| Layer | Choice |
| :-- | :-- |
| Desktop shell | Electron 30 |
| UI | React 18 + TypeScript 5 |
| Bundler | Vite 5 + `vite-plugin-electron` |
| Styling | Tailwind CSS 4 |
| Animation | Framer Motion |
| Icons | Lucide React |
| Reading spreadsheets | SheetJS (`xlsx`) |
| Writing spreadsheets | ExcelJS |

---

## Project structure

```text
BOM_Tool/
├── electron/
│   ├── main.ts                 # window lifecycle, IPC handlers (save dialog, window, links)
│   └── preload.ts              # contextBridge IPC bridge
├── src/
│   ├── components/
│   │   ├── bom-dashboard.tsx   # Analyzer + Comparator workspaces, grid, export
│   │   ├── pnp-validator.tsx   # PnP validation workspace
│   │   ├── SplashScreen.tsx    # animated launch screen
│   │   └── ui/                 # reusable dialogs, search, filters, buttons
│   ├── lib/bom/
│   │   ├── bom-processor.ts    # parsing, header/column mapping, designator logic
│   │   ├── bom-compare.ts      # single-pass Map/Set comparison engine
│   │   └── pnp-parser.ts       # delimited placement-file parser
│   ├── App.tsx
│   └── main.tsx
├── samples/pnp.xlsx            # sample Altium-style placement export
├── public/                     # app + installer artwork
├── electron-builder.json5      # packaging configuration
└── vite.config.ts
```

---

## Getting started

**Requirements:** Node.js 18+ and npm 9+.

```bash
git clone https://github.com/MZarc/BOM_Tool.git
cd BOM_Tool
npm install
npm run dev
```

`npm run dev` starts the Vite dev server; `vite-plugin-electron` launches the Electron
window against it with hot reload.

### Scripts

| Command | Does |
| :-- | :-- |
| `npm run dev` | Vite dev server + Electron |
| `npm run lint` | ESLint, zero warnings allowed |
| `npm run build` | `tsc` → `vite build` → `electron-builder` |

### Build & release

```bash
npm run build
```

Output for the current version:

```text
release/1.2.5/
├── BOM Tool-Windows-1.2.5-Setup.exe   # NSIS installer, x64
└── win-unpacked/                      # unpacked application
```

`electron-builder.json5` also declares macOS (DMG) and Linux (AppImage) targets, but
only the Windows x64 NSIS target is built and exercised here.

---

## Keyboard

| Shortcut | Action |
| :-- | :-- |
| `Ctrl/Cmd + F` | Universal search for the active workspace |

Undo/redo are available from the toolbar buttons next to the file name; history is
kept per file.

---

## Notes and limitations

- **Unsigned build.** No code-signing certificate is configured, so Windows
  SmartScreen may warn on first launch.
- **Renderer security model.** The renderer has no Node integration. A preload script
  exposes a general IPC bridge (`send` / `invoke` / `on`) to the main process for the
  save dialog, window controls and opening external links.
- **Optional certificate bootstrap.** `electron/main.ts` checks for
  `Meet_Mistry_cert.pfx` and, only if the file is present, asks for elevation to add
  it to the Windows Root and TrustedPublisher stores. The certificate is not included
  in this repository (`.pfx` is git-ignored) nor in the published installer.
- **Fixed window.** The window opens maximized and resizing is disabled by design.
- Validation is deterministic string/number logic over the parsed rows; it does not
  consult external part databases.

---

## License

MIT © 2026 Meet Mistry — see [LICENSE](LICENSE).

<div align="center">
  <sub>Built for hardware teams that would rather catch the mistake before the board is fabricated.</sub>
</div>
