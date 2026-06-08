# Version History

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While `MAJOR=0`, breaking changes can occur on `MINOR` bumps.

## [0.3.0] - 2026-06-08

### Added

- **Read `.ods` (OpenDocument Spreadsheet) files** in both the Desktop and Web apps, alongside `.xlsx`. A new format-agnostic front door (`SpreadsheetReader`) dispatches by extension.
- New `ODSReader` parses an `.ods` `content.xml` into the **same** `XLSXWorkbook` / `XLSXSheet` / `XLSXCell` model the XLSX reader uses, so tabs, listbox rendering, the format engine, the Memory/Disk zip backend, and the per-phase parse timer all work unchanged.
- ODS coverage: cell value types (string / float / percentage / currency / date / time / boolean), inline `<text:p>` text, formula cells (cached value), merged cells (`table:number-rows/columns-spanned` + `<table:covered-table-cell>`), and `number-columns/rows-repeated` compression.
- Converts OpenDocument `<number:date-style>` / `<number:number-style>` / `<number:currency-style>` / `<number:percentage-style>` element trees into the format-code strings the engine already renders (currency symbols emitted as `[$SYM-0]` tags; ISO dates converted to Excel serials).
- `.ods` files use the identical `XLSXZip` dual backend — the "Read in memory" checkbox and the `zip X + xml Y (Memory|Disk)` parse-time readout apply to ODS too.
- Reused-infrastructure additions: `XLSXSheet` empty constructor + `PutCell` / `AddMergedRange`; `XLSXCell.FormatCode`; `XLSXFormatter.DateTimeToExcelSerial`.
- ODS test fixtures: `ods-multi-sheet.ods`, `ods-types.ods`, `ods-sales.ods`, `ods-styled.ods`.

### Compatibility

- No breaking changes. `XLSXReader.Open` is unchanged; `SpreadsheetReader.Open` is the new recommended entry point, and `ODSReader.Open` is available to force ODS.

## [0.2.1] - 2026-05-11

### Added

- More Excel format-code support:
  - **Scientific notation**: `0.00E+00`, `0E+00`, `##0.0E+0` (built-in id 48).
  - **Number with parens for negatives**: built-in ids 37–40 (`#,##0 ;(#,##0)` and variants).
  - **Accounting**: built-in ids 41–44 — `$` prefix when present, parentheses for negatives, dash for zero. Detected via the `_(` … `_)` spacer pattern, so custom accounting codes work too.
  - **Currency tag**: codes like `[$$-409]#,##0.00` and `[$€-2]#,##0` — symbol extracted from `[$X-Y]` and used as prefix or suffix based on position.
  - **Text format**: `@` → cell value passes through unchanged.
  - **More dates**: `m/d/yy h:mm`, `m/d/yyyy h:mm`, `mm/dd/yyyy`.
- Excel built-in `numFmtId` 37–44, 48, 49 are now seeded in `XLSXStyles`. Workbooks that reference these IDs without an explicit `<numFmt>` element resolve correctly.

### Behavior

- The `microsoft-financial-sample.xlsx` Sale Price / Gross Sales columns now render as `$1,618.50` (accounting) instead of the raw `1618.5`.

## [0.2.0] - 2026-05-11

### Added

- Dual-backend `XLSXZip` chosen at Open time via a new `mode` parameter on both `XLSXReader.Open` and `XLSXZip.Open`:
  - **Auto** (default) — Memory on Xojo 2024r3+, else Disk.
  - **Memory** — pure in-memory zip parse + `MemoryBlock.Decompress` (no disk I/O, sandbox-friendly).
  - **Disk** — `FolderItem.Unzip` into `SpecialFolder.Temporary` (previous behaviour).
- Memory backend contributed by Andrew Lambert (@charonn0) — parses local file headers from a `MemoryBlock` and wraps raw deflate entries in a synthetic GZIP header so the framework `MemoryBlock.Decompress` (Xojo 2024r3+) can decompress them.
- `XLSXEnums.eOpenMode` enum with `Extends ToString` helper.
- Desktop: visible "Open…" button at the top-left of the main window, sharing the same handler as File → Open (Cmd-O).
- Desktop + Web: "Read in memory" checkbox (default checked) lets the user pick the backend at runtime.
- Desktop + Web: parse-time label with per-phase breakdown — e.g. "Parsed in 487 ms (zip 12 + xml 475, Memory)". Lets users see the ~10× zip-phase speedup of Memory vs Disk.
- `XLSXWorkbook` exposes `OpenMode`, `ZipMicroseconds`, `XmlMicroseconds` for diagnostic timing.

### Notes

- The Memory backend's real benefit is sandbox-friendliness — the speed parity for the *total* parse comes from the XML-parse phase dominating (identical work in both backends). The zip-extraction phase itself is roughly 10× faster in Memory mode.

## [0.1.0] - 2026-05-10

### Added

- Pure-Xojo XLSX reader in shared `Common/` folder, referenced from both `.xojo_project` files via `Module=`/`Class=` entries pointing at `../Common/`. Cross-platform: Mac / Windows / Linux desktop, plus Web 2.0 server runtime.
- Single-call public API: `XLSXReader.Open(file As FolderItem) As XLSXWorkbook` + a `MemoryBlock` overload for in-memory inputs.
- Workbook model: `XLSXWorkbook` (sheets / sharedStrings / styles), `XLSXSheet` (sparse cell dictionary + merged ranges), `XLSXCell` (typed value + lazy `DisplayText(styles)`), `XLSXCellRange`, `XLSXCellRef`.
- Format engine: `XLSXFormatter` covers `General` / `0` / `0.00` / `#,##0` / `#,##0.00` / `0%` / `0.00%` for numbers and `yyyy-mm-dd` / `dd/mm/yyyy` / `hh:mm` / `hh:mm:ss` / `yyyy-mm-dd hh:mm` for dates. Excel epoch 1899-12-30. Built-in numFmt ids 0..22 seeded in `XLSXStyles`; custom `<numFmts>` codes override.
- Zip part reader (`XLSXZip`) using the framework's own `FolderItem.Unzip`; magic-byte verification; per-instance temp-folder cleanup in `Destructor`. Both `FolderItem` and `MemoryBlock` entry points.
- Typed exception (`XLSXException`) with `eParseError` code (`NotAZip` / `MissingPart` / `MalformedXML` / `Encrypted` / `Unsupported`) plus an `XLSXHelpers` global module with `Extends`-based `eEnum.ToString`.
- Desktop app: `MainWindow` with `DesktopTabPanel TabPanelSheets` + `DesktopListBox ListboxData`; `MainMenuBar` File → Open (Cmd-O) opens a filtered `OpenFileDialog`; per-sheet tabs; column auto-sizing + horizontal scrollbar; user-drag column resizing.
- Web app: `MainPage` with `WebFileUploader UploaderXLSX` (auto-triggers `StartUpload` on `FileAdded`) + `WebTabPanel TabPanelSheets` + `WebListBox ListboxData`; per-sheet tabs; same UX as Desktop.
- Both fillers skip styled-but-empty rows so workbooks like `BUA 2024` (485 nominal rows, mostly empty) render only their real content.
- Localizable user-visible strings in the `strings` module; UI errors mapped to localizable kStrError… constants.
- Developer documentation: [`developper_doc.md`](developper_doc.md).
- Test fixtures: 4 public XLSX samples in [`test_files/`](test_files/).

### Verified manually

- Both projects analyze cleanly in Xojo 2026r1 with no API 2 warnings.
- Desktop opens a 33-sheet test workbook end-to-end: tabs populated, listbox repopulates on tab change, shared-string lookup resolves, merged-cell anchors render correctly, columns auto-sized + user-resizable + horizontally scrollable.
- Web opens the same workbook end-to-end via file upload.

### Known limitations

- Read-only; no formula evaluation (cached values shown); no images / charts / pivots / conditional formatting; no encrypted workbooks; cell colors / fonts / borders are not rendered.
- Parser is eager — large workbooks (≫ 10k rows × 50 cols) may load slowly.
- Format-code support is a pragmatic subset; unknown codes fall back to `Double.ToString` / `DateTime.SQLDateTime`.
