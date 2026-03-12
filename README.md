# ReportTools for Frontline CA ERP

![Status](https://img.shields.io/badge/status-active%20development-blue)
![Version](https://img.shields.io/badge/version-0.6.0-informational)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Excel-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

A production-grade Excel add-in that normalizes exported reports from Frontline CA ERP systems into clean, analysis-ready datasets.

Many of Frontline CA ERP's reports do not export in a structured, Excel-friendly format. ReportTools converts those raw exports into normalized, tabular datasets suitable for pivot analysis, reconciliation, audit workflows, and downstream data validation — without modifying the source worksheet.

## Contents

- [ReportTools for Frontline CA ERP](#reporttools-for-frontline-ca-erp)
  - [Contents](#contents)
  - [Report Modules](#report-modules)
  - [Architecture](#architecture)
    - [Core Add-In — `ReportTools_Core.xlam`](#core-add-in--reporttools_corexlam)
    - [Launcher Add-In — `ReportTools_Launcher.xlam` *(planned)*](#launcher-add-in--reporttools_launcherxlam-planned)
  - [Design Philosophy](#design-philosophy)
  - [Build Process](#build-process)
  - [Installation](#installation)
    - [End Users](#end-users)
    - [Managed IT Deployment](#managed-it-deployment)
  - [Roadmap to v1.0.0](#roadmap-to-v100)
    - [Report Modules](#report-modules-1)
    - [Launcher Add-In](#launcher-add-in)
    - [Distribution \& Integrity](#distribution--integrity)
    - [Documentation](#documentation)
  - [Distribution \& Signing](#distribution--signing)
  - [Requirements](#requirements)
    - [End Users](#end-users-1)
    - [Developers](#developers)
  - [Versioning](#versioning)
  - [Disclaimer](#disclaimer)

## Report Modules

Each module targets a specific ERP export format and normalizes it into a flat, tabular dataset. Modules are written to tolerate header variations, footer noise, multi-line record blocks, and common ERP formatting inconsistencies.

| Module | Report | Output Description |
|--------|--------|--------------------|
| `Pay03` | Payroll Summary | Normalized payroll summary rows |
| `Pay14` | Net Pay / Deductions | One row per employee per deduction/contribution item |
| `Ben02` | Benefits | One row per employee per benefit provider and level |
| `Pos04` | Position Control | One row per employee per budget code allocation |
| `Fiscal05` | Fiscal Conversion | Normalized fiscal conversion report rows |

Each module is accessible from the ReportTools ribbon tab via a workbook and sheet picker — no manual sheet wiring required.

## Architecture

ReportTools is structured as a two-part system:

### Core Add-In — `ReportTools_Core.xlam`

Contains all report transformation logic and the RibbonX UI.

- **Picker architecture**: each report entry point prompts for a source workbook and sheet, then passes a `Worksheet` reference to a private worker procedure
- **Never operates on `ThisWorkbook`**: all logic is scoped to `wsSrc.Parent` to avoid side effects
- **Performance-first design**: array reads, bulk writes, calculation guards, and `ScreenUpdating` guards throughout
- **No external service calls**, no dynamic code execution, no obfuscation

### Launcher Add-In — `ReportTools_Launcher.xlam` *(planned)*

Responsible for update management, Core validation, and versioning. Designed for enterprise deployment scenarios where IT manages distribution centrally.

## Design Philosophy

This project is built for use in audited financial environments where behavior must be predictable and traceable.

- **Deterministic transformations** — the same input always produces the same output
- **No silent failures** — anomalies are flagged in output, not quietly discarded or corrected
- **No hidden behavior** — no external calls, no dynamic execution, no obfuscation
- **Reproducible builds** — the distributed `.xlam` is always generated from source via a documented build script

The goal is to make this tool acceptable in managed IT environments, county offices of education, public school districts, and audited financial operations.

## Build Process

The add-in is built from source using a PowerShell script and Excel COM automation. The distributed `.xlam` is never hand-edited.

```
build/build_core.ps1
```

**Steps:**

1. Copies `CoreTemplate.xlam` to `dist/`
2. Clears existing VBA components from the copy
3. Imports all modules from `src/vba/`
4. Injects RibbonX from `src/ribbon/CustomUI14.xml`
5. Saves the production-ready `.xlam`

**Development requirement:** Excel's Trust Center must have *"Trust access to the VBA project object model"* enabled.

See [`docs/DEVELOPMENT.md`](docs/DEVELOPMENT.md) for the full development and release workflow.

## Installation

### End Users

1. Download `ReportTools_Core.xlam` from the [latest release](../../releases/latest)
2. Open Excel → **File → Options → Add-ins**
3. At the bottom, select **Manage: Excel Add-ins** → **Go**
4. Click **Browse**, locate the downloaded `.xlam`, and click **OK**
5. Ensure **ReportTools** appears checked in the list
6. Macros must be enabled

The **ReportTools** tab will appear in the Excel ribbon.

### Managed IT Deployment

The add-in can be deployed via Group Policy or SCCM by placing the `.xlam` in a shared network location and configuring the `OPEN` registry key under the Excel startup path. See Microsoft's documentation on [deploying Excel add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview).

## Roadmap to v1.0.0

ReportTools is under active development. The following milestones define the path to a stable v1.0.0 release.

### Report Modules
- [ ] Additional HR/Payroll report modules (in progress)
- [ ] Additional Finance/Fiscal report modules (in progress)
- [ ] Full module test coverage with known-good reference exports

### Launcher Add-In
- [ ] `ReportTools_Launcher.xlam` — update management and Core validation
- [ ] Version check on startup
- [ ] Graceful handling of Core load failures

### Distribution & Integrity
- [ ] Self-signed code certificate applied to release artifacts
- [ ] SHA-256 checksums published with each release
- [ ] Build script extended to generate checksum automatically

### Documentation
- [ ] End-user guide per report module
- [ ] Managed IT deployment guide
- [ ] Contributor guide

v1.0.0 will be the first release recommended for production deployment in managed environments.

## Distribution & Signing

Production-signed releases with SHA-256 checksums are planned for v1.0.0. Pre-release builds are unsigned and intended for evaluation and testing only.

Organizations evaluating deployment in managed environments should wait for v1.0.0 or test pre-release artifacts in a non-production context.

## Requirements

### End Users
- Microsoft Excel for Windows (desktop)
- Macros enabled

### Developers
- Microsoft Excel for Windows
- PowerShell
- Excel Trust Center: *"Trust access to the VBA project object model"* enabled
- [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) (for ribbon changes only)

## Versioning

This project follows [Semantic Versioning](https://semver.org/): `MAJOR.MINOR.PATCH`

| Increment | When |
|-----------|------|
| `MAJOR` | Breaking architectural changes |
| `MINOR` | New report modules or significant enhancements |
| `PATCH` | Bug fixes or minor improvements |

See [CHANGELOG.md](CHANGELOG.md) for the full project history.

## Disclaimer

This project is an independent reporting utility. It is not affiliated with, endorsed by, or supported by any ERP vendor or educational organization.