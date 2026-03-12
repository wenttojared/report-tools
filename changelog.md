# Changelog

All notable changes to ReportTools will be documented here. Versioning follows [Semantic Versioning](https://semver.org/): `MAJOR.MINOR.PATCH`

## [0.7.0] - 2026-03-12
### Added
- **repPay13** - New report module. Normalizes Pay13a Payroll Adjustments by Employee exports into a flat table with one row per code entry. Columns: `District, EmployeeName, EmployeeID, SSN_Last4, PayCycle, PayDate, T, Code, Description, Date, Position_Vendor_HRA, DeductionAmount, ContributionAmount, PayRate, Units, EarningsAmount, BudgetCode, RetirementSystem, PayPeriod, CC, PC, Wrk_Assgn, Rate, SourceSheet`. Handles D/C type entries (Line 1 only), earnings codes with no retirement system, optional budget code rows, and optional retirement system lines (PERS and STRS variants). District name backfilled via pre-scan block lookup; numeric summary and footnote rows skipped.
- **modEntryPoints** - `Run_Pay13_WithPicker` entry point.
- **CustomUI14.xml** - Pay13 button added to Payroll menu.
### Changed
- **modCoreMeta** - Version bumped to 0.7.0.
- **modRT_Parse** - `IsEmployeeHeader` and `ParseEmployeeHeader` promoted from private to public. Previously duplicated in repPay14; now shared across all modules requiring employee header parsing.
- **repPay14** - Private `IsEmployeeHeader` and `ParseEmployeeHeader` removed; replaced with calls to public versions in modRT_Parse.

## [0.6.0] - 2026-02-25
### Added
- **repPos04** - New report module. Normalizes Pos04 position control exports into a flat table with one row per employee per budget code. Columns: `OrgID, BU, AssignType, Employee, EmployeeID, Location, JobCategory, JobClass, CalendarDays, Placement, Rate, StartDate, EndDate, FTE_Authorized, FTE_Assigned, BudgetCode, AccountPct, Amount, SourceSheet`. Handles multi-Org files, mid-year assignment splits (multiple header rows sharing one budget code block), multi-account salary distributions, and multi-district exports.
- **modEntryPoints** - `Run_Pos04_WithPicker` entry point.
- **CustomUI14.xml** - Pos04 button added to Position Control menu.
### Changed
- **modCoreMeta** - Version bumped to 0.6.0.

## [0.5.1] - 2026-02-24
### Added
- **modRT_Parse** - New public function `TryParseVendor`. Parses vendor cells of the form `(NNNNNN/NNN) VendorType` into `VendorID` (6 digits), `VendorAddrID` (3 digits), and `VendorType` (text label). Includes `Debug.Print` diagnostics on each structural check.
- **repPay14** - `Vendor` column replaced by three columns: `VendorType`, `VendorID`, `VendorAddrID`.
- **repPay14** - Output column order revised: `EmployeeID, SSN_Last4, EmployeeName, PayDate, EffectiveDate, NetPay, VendorID, VendorAddrID, VendorType, DedContribName, DeductionAmount, ContributionAmount, SubjectGross_Ded, SubjectGross_Contrib, SourceSheet`.
### Changed
- **modRT_Parse** - `TryParseIdSsn4` and `TryParseOrg3` now emit `Debug.Print` diagnostics on non-trivial parse failures instead of exiting silently. Blank cells still skip without noise.
- **repPay14** - `COL_CC_FALLBACK` and `PAY_DATE_OFFSET` promoted to named constants. A warning is logged when CC column detection falls back to the default.
- **repPay14** - Removed unnecessary `outFinal` array copy prior to sheet write. Output buffer is now written directly, halving peak memory usage on large files.
- **repPay14** - Missing employee header and missing `Total Deductions` sentinel row now both log `Debug.Print` warnings with row and sheet context instead of skipping silently.
- **repFiscal05** - `IsDetailDataRow` documented with the Frontline account code format it detects (`Fd-Resc-Y-Goal-Func-Objt-SO-Sch-DD1-DD2`). Account-before-org evaluation order explained inline.
### Fixed
- **repPay14** - `ParseEmployeeHeader` now routes through `TryParseIdSsn4` and `MaskSsn4` (shared parse module) instead of duplicating extraction logic. EmployeeID is correctly zero-padded to 6 digits; SSN_Last4 is formatted as `XXX-XX-####`.
- **repPay14** - PayDate and EffectiveDate columns now render as `MM/DD/YYYY` instead of raw Excel date serials.
- **repPay14** - EmployeeID, VendorID, and VendorAddrID columns are forced to text format before the bulk write, preventing Excel from silently dropping leading zeroes.
- **repBen02** - Rows belonging to the final org section no longer go silently unstamped when the source data ends without a trailing `Total for Org` row. A warning is now logged to the Immediate window identifying the affected row range and sheet.
