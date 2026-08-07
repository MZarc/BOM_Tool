<p align="center">
  <img src="public/icon.png" alt="BOM Tool Icon" width="120" />
  
</p>

# BOM Tool

> **Production-Ready Hardware BOM & Pick-and-Place (PNP) Validation Engine**

[![Electron](https://img.shields.io/badge/Electron-30.0.1-47A248?style=flat&logo=electron&logoColor=white)](https://www.electronjs.org/)
[![React](https://img.shields.io/badge/React-18.2.0-61DAFB?style=flat&logo=react&logoColor=black)](https://react.dev/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.2-3178C6?style=flat&logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Vite](https://img.shields.io/badge/Vite-5.1-646CFF?style=flat&logo=vite&logoColor=white)](https://vitejs.dev/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A high-performance desktop engineering utility engineered for hardware manufacturing, SMT assembly, and component procurement teams to manage, compare, validate, and export Bills of Materials (BOM) and Pick-and-Place (PNP) data with zero data loss.

---

## 📸 Overview

In PCB manufacturing, component mismatches, designator oversights, or modified MPNs between hardware revisions cause costly production delays and scrap boards. 

**BOM Tool** solves this by providing instant $O(N)$ differential analysis between reference and target BOMs, auto-validating component designators against Pick & Place files, and exporting publication-ready Excel reports that **preserve original corporate headers, logos, and custom template styles**.

---

## ✨ Key Features

### ⚡ 1. High-Performance Differential BOM Comparator
- **Instantaneous Comparison:** Rapid $O(N)$ differential engine capable of processing 10,000+ line items in milliseconds.
- **Mismatch Categorization:** Automatically categorizes differences into:
  - 🔴 **Missing Items:** Components in Reference but absent in Target.
  - 🟢 **Extra Items:** Components added in Target BOM.
  - 🟡 **Quantity Mismatches:** Discrepancies in total part counts.
  - 🔵 **MPN / Value Swaps:** Substituted or altered Manufacturer Part Numbers.

### 🎯 2. Pick & Place (PNP) File Validator
- Validates SMT placement files against master BOM lists.
- Checks designator counts (e.g., `C1, C2, C3`).
- Highlights unplaced BOM parts and unlisted PNP designators.

### 📊 3. High-Fidelity Template-Preserving Excel Export (ExcelJS)
- Unlike standard CSV or XLSX exporters that strip formatting, **BOM Tool** uses `ExcelJS` to inject differential status columns while preserving corporate title blocks, logos, merged cells, fonts, and custom color fills.

### 🤖 4. Intelligent Heuristic Column Mapping
- Automatically identifies and maps non-standard column headers across vendor spreadsheets (e.g., `MPN`, `Manf Part #`, `Qty`, `Quantity`, `RefDes`, `Designators`, `Description`).
- Includes interactive manual mapping overrides for complex non-standard templates.

### 🎨 5. Liquid Glass Modern UI
- Premium dark/light interface designed with modern glassmorphism aesthetic.
- Fluid micro-animations powered by **Framer Motion** and **Tailwind CSS**.
- Built-in universal instant search & custom status filtering.

---

## 🛠 Tech Stack

| Domain | Technology | Description |
| :--- | :--- | :--- |
| **Desktop Shell** | [Electron 30](https://www.electronjs.org/) | Cross-platform desktop runtime |
| **Frontend Framework** | [React 18](https://react.dev/) + [TypeScript 5](https://www.typescriptlang.org/) | Type-safe UI components |
| **Build Tool** | [Vite 5](https://vitejs.dev/) | Lightning-fast development & bundling |
| **Excel Parsing** | [SheetJS (XLSX)](https://sheetjs.com/) | Rapid memory extraction of `.xlsx`, `.xls`, `.csv` |
| **Excel Exporting** | [ExcelJS](https://github.com/exceljs/exceljs) | High-fidelity template preservation |
| **Styling & UI** | [Tailwind CSS v4](https://tailwindcss.com/) + [Framer Motion](https://www.framer.com/motion/) | Responsive design & smooth animations |
| **Icons** | [Lucide React](https://lucide.dev/) | Clean vector iconography |

---

## 📂 Project Structure

```
BOM_Tool/
├── electron/                 # Electron main & preload processes
│   ├── main.ts               # IPC handlers, window management, auto-updates
│   ├── preload.ts            # Secure contextBridge API bindings
│   └── electron-env.d.ts     # Process environment type definitions
├── src/
│   ├── components/           # React UI components
│   │   ├── bom-dashboard.tsx # Main BOM Comparison Workspace & Table Viewer
│   │   ├── pnp-validator.tsx # Pick & Place Validation Interface
│   │   ├── SplashScreen.tsx  # Animated app launcher screen
│   │   └── ui/               # Reusable UI primitives (Dialogs, Search, Buttons)
│   ├── lib/                  # Core domain logic
│   │   └── bom/
│   │       ├── bom-processor.ts # Excel parsing, column mapping & template engine
│   │       ├── bom-compare.ts   # O(N) differential comparison engine
│   │       └── pnp-parser.ts    # Pick & Place placement parser
│   ├── App.tsx               # Root application router & shell
│   └── main.tsx              # React entry point
├── samples/                  # Sample test spreadsheets (e.g., pnp.xlsx)
├── public/                   # Static application assets & icons
├── electron-builder.json5   # Desktop packager & release configuration
├── package.json              # Dependencies & build scripts
└── vite.config.ts            # Vite & Electron integration configuration
```

---

## 🚀 Getting Started

### Prerequisites
- **Node.js**: `v18.0.0` or higher
- **npm**: `v9.0.0` or higher

### 1. Clone the Repository
```bash
git clone https://github.com/MZarc/BOM_Tool.git
cd BOM_Tool
```

### 2. Install Dependencies
```bash
npm install
```

### 3. Run Development Server
Launches the Vite dev server and Electron app simultaneously with hot-reload enabled:
```bash
npm run dev
```

### 4. Build Desktop Executable
Compiles TypeScript, bundles frontend assets with Vite, and builds standalone Windows/macOS binaries via `electron-builder`:
```bash
npm run build
```
Output executables will be placed in the `release/` directory.

---

## 🔒 Security & Best Practices

- **Code Signing:** Certificates (`.pfx`) are excluded from source control. Provide code signing secrets via environment variables during CI/CD builds.
- **Local File Security:** Electron main process operates with restricted `nodeIntegration: false` and secure `contextBridge` IPC channels.

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
