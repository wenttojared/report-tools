# ReportTools – Developer Guide

## Purpose

This document describes the development and release workflow for the ReportTools Excel add-in. This project uses a source-controlled build system to produce reproducible `.xlam` release artifacts. This is irrelevant to end users as they **do not build the project**.

## Source of Truth

The authoritative source for the add-in is:
- `/src/vba`
- `/src/ribbon/CustomUI14.xml`

The `.xlam` file is always generated from source via the build script. Do not manually edit a built `.xlam` and commit it.

## Development Workflow

1. Make changes in development add-in.
2. Export updated modules into `/src/vba`.
3. Update `/src/ribbon/CustomUI14.xml` if ribbon changes were made.
4. Commit changes.
5. Run build script: `.\build\build_core.ps1`
6. Verify the generated file: `dist/ReportTools_Core.xlam`
7. Test locally by loading it in Excel.

## Release Workflows

Two workflows apply depending on the release stage. The pre-1.0.0 workflow is current. The v1.0.0+ workflow documents the intended process for production releases.

### Pre-1.0.0 Release Workflow (current)

1. Update version number in `modCoreMeta`.
2. Update `CHANGELOG.md`.
3. Commit changes with a version bump commit message (e.g. `chore: bump version to 0.7.0`).
4. Run build script: `.\build\build_core.ps1`
5. Verify the generated file: `dist/ReportTools_Core.xlam`
6. Create a GitHub Release:
   - Tag: `v0.x.x`
   - Mark as **Pre-release** in the GitHub UI
   - Paste the relevant `CHANGELOG.md` section as the release body
   - Include the pre-release notice in the release notes (see below)
   - Upload `ReportTools_Core.xlam` as the release artifact
7. Publish the release.

**Required pre-release notice in release notes:**

> ⚠️ **Pre-Release Build**
> This is an unsigned pre-release artifact intended for evaluation and testing only.
> SHA-256 checksums and code signing are planned for v1.0.0 production releases.
> To use, enable macros and install via Excel's Add-ins dialog. See the README for setup instructions.

### v1.0.0+ Release Workflow (planned)

1. Update version number in `modCoreMeta`.
2. Update `CHANGELOG.md`.
3. Commit changes.
4. Run build script: `.\build\build_core.ps1`
5. Verify the generated file: `dist/ReportTools_Core.xlam`
6. Digitally sign the artifact.
7. Generate SHA-256 checksum:
   ```powershell
   Get-FileHash dist\ReportTools_Core.xlam -Algorithm SHA256 | Select-Object Hash | Out-File dist\ReportTools_Core.xlam.sha256.txt
   ```
8. Create a GitHub Release:
   - Tag: `v1.x.x`
   - Do **not** mark as Pre-release
   - Paste the relevant `CHANGELOG.md` section as the release body
   - Upload the signed `.xlam` and the `.sha256.txt` as release artifacts
9. Publish the release.

## Requirements for Development

- Windows
- Microsoft Excel (desktop)
- PowerShell
- Excel setting enabled:
  - Trust Center → *"Trust access to the VBA project object model"*
- [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) (for ribbon changes only)

## Security Model

| Property | Pre-1.0.0 | v1.0.0+ |
|----------|-----------|---------|
| Code signing | None | Self-signed certificate |
| SHA-256 checksum | Not published | Published with each release |
| Obfuscation | None | None |
| External service calls in Core | None | None |
| Dynamic code execution | None | None |

Pre-1.0.0 artifacts are unsigned and intended for evaluation and testing only. Organizations deploying in managed environments should wait for v1.0.0.

## Versioning

Semantic Versioning: `MAJOR.MINOR.PATCH`

| Increment | When |
|-----------|------|
| `MAJOR` | Breaking architectural changes |
| `MINOR` | New report modules or significant enhancements |
| `PATCH` | Bug fixes or minor improvements |

Examples:
- `1.0.0` — First production release
- `1.1.0` — New report module added
- `1.1.1` — Bug fix

## Release Integrity

### Pre-1.0.0
Pre-release artifacts are unsigned and distributed without checksums. They are intended for evaluation and testing only.

### v1.0.0+
For every production release:
- The artifact must be digitally signed.
- A SHA-256 checksum must be published alongside the artifact.
- The checksum must be verified against the signed artifact before publishing.
