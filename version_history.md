# Version History

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While `MAJOR=0`, breaking changes can occur on `MINOR` bumps.

## [0.6.0] - 2026-08-16

### Added
- **Iterate a workbook's sheets with `For Each`.** `XLSXWorkbook` now implements Xojo's `Iterable`, so `For Each sheet As XLSXSheet In workbook` just works (backed by a new `XLSXSheetIterator`). `XLSXWorkbook.Sheets` also returns a plain `XLSXSheet()` array when you'd rather index or sort them.
  - **Rows, columns and cells** iterate through arrays rather than dedicated iterator classes — `XLSXSheet.RowCells(row)` and `ColumnCells(col)` return `XLSXCell()`, which Xojo's `For Each` consumes natively: `For Each c As XLSXCell In sheet.RowCells(3)`. Absent cells come back as the shared empty sentinel, so the arrays are always full-length and never `Nil`. Both have always-1-based `*Raw` twins.
- **Optional 0-based indexing — `XLSXHelpers.gZeroBasedSheetsRowsColumns`.** A global Boolean, **False by default**, that switches the public API's sheet/row/column index base from Excel-style 1-based to 0-based. Set it once at startup; with it `False` **nothing changes at all** — the translation is a literal identity, so existing code and files behave exactly as before.
  - **Applies to indexes going in** — `XLSXSheet.CellAt` / `Cell` (get + set) / all `Put*` / `PutCell` / `ColumnWidth` / `SetColumnWidth` / `RowHeight` / `SetRowHeight` / `AutoFitColumn` / `EffectiveStyle` / `IsCellMergedFollower` / `AddMergedRange`, and `XLSXWorkbook.SheetAt`.
  - **…and coordinates coming out** — `XLSXCellRange` and `XLSXTable` corners (`FirstRow` / `FirstCol` / `LastRow` / `LastCol`, now computed properties), `XLSXTable.Contains` / `IsHeaderRow` / `IsTotalsRow` / `FirstDataRow`, `XLSXSheet.TabIndex`, and `XLSXCellRef.ColLettersToIndex` / `IndexToColLetters` / `A1ToRowCol`.
  - **The model stays 1-based internally.** Every parser, serializer, filler and UI path was moved onto a new always-1-based `*Raw` family (`CellAtRaw`, `PutCellRaw`, `SheetAtRaw`, `ColumnWidthRaw`, `IndexToColLettersRaw`, `A1ToRowColRaw`, `XLSXCellRange.NewRaw`, `XLSXTable.NewRaw`, `FirstRowRaw` …), so file coordinates — which are always 1-based on disk — can never be shifted by the flag. Collection iterators (`MergedRangeAt(i)`, `TableAt(i)`) were already 0-based and are unaffected.
  - **Desktop test switch:** a "0-based indexes" checkbox flips the flag at runtime. Because the app itself runs on the `*Raw` API, toggling it should change *nothing* on screen — the grid shifting would mean a call site still goes through the translating API. What does visibly change is **Gen code**, whose emitted source switches between 1-based and 0-based indices to match the base the generated code will run under.
- **ODS cell styling — read.** `.ods` files now read back their visual cell styling (fonts, fills, alignment, borders), not just number formats — inverting what `ODSWriter` already emits. `ODSReader.ParseCellVisualStyle` parses each `table-cell` `<style:style>`'s `<style:table-cell-properties>` (fill, per-edge + `fo:border` shorthand, vertical align, wrap), `<style:paragraph-properties>` (horizontal align), and `<style:text-properties>` (bold/italic/underline, colour, size, family) into an `XLSXCellStyle` attached to the cell. So a styled `.ods` (or an `.xlsx` saved as `.ods`) now renders its styling on reopen. Fixture: `test_files/ods-visual-styles.ods`.
- **Auto-fit column widths.** New **Auto-fit** button (Desktop + Web) sizes the current sheet's columns to their content, and `XLSXSheet.AutoFitColumns` / `AutoFitColumn(col)` do the same from code. Measured in Excel's own character-width unit (longest line per cell, with a nudge for bold and larger fonts, clamped 4–60 chars); merged cells are skipped, as Excel's AutoFit does. Both take an optional `charWidthPx` (default `7.0`, Excel's Calibri-11 reference) so callers can match the font they actually render in — `XLSXHelpers.ColumnCharsToPoints` gained the same parameter. Widths land in the model, so they show on screen **and** are written to the saved file.
- **Excel Table — write.** Tables (ListObjects) are now re-emitted on `.xlsx` save, not just rendered: `XLSXWriter` writes `xl/tables/tableN.xml` (range, `<autoFilter>`, `<tableColumns>` named from the header row, `<tableStyleInfo>` with the banding flags), the worksheet `_rels` + `<tableParts>`, and the `[Content_Types]` overrides. So a workbook with an Excel Table (e.g. the Microsoft Financial Sample) saved as `.xlsx` keeps the table as a real table, header/banding intact.

### Fixed
- **Excel Table banding wiped out the cell's own font.** `XLSXSheet.EffectiveStyle` returned the synthetic table style *instead of* the cell's style, so a banded row lost the workbook's font while the plain row next to it kept it — visibly different text sizes on alternating rows of the same table. The table style is now **merged** onto the cell's: the table contributes the fill (and a header row's bold + light text), the cell keeps its own font name and size.
- **Web rendered fonts ~27% too small.** `WebStyle.FontSize` is in CSS **pixels**, but the model stores **points**, and the filler passed points straight through (an 11pt font became 11px instead of ~14.7px). Now converted via `XLSXHelpers.PointsToPixels` — the same points-vs-pixels mix-up already fixed for Desktop column widths.
- **XLSX write invented a font size / name for cells that didn't specify one.** `FontIdFor` filled in `11pt Calibri` whenever the model left them unset, so "unspecified" became an explicit override on reopen — an italic-only cell (common from ODS, which only records what's set) came back rendered smaller than the original. Size and name are now emitted only when actually specified; the complete default font stays at index 0, so the 0.4.0 Quick Look row-height fix is unaffected.
- **Desktop: the last column couldn't be scrolled to fully.** The listbox can't scroll past the sum of its column widths, so the rightmost column's edge sat flush against the frame (partly under the vertical scrollbar) — most visible right after an auto-fit. The filler now adds a slack pad to the last column, and `CaptureColumnWidths` subtracts the same amount so it never accumulates into the model or the saved file.
- **Desktop column widths were ~25% too narrow, and shrank on every save.** The model stores widths in points, but the Desktop filler passed them to the listbox as pixels, and `CaptureColumnWidths` read listbox pixels back as points. Both now convert via new `XLSXHelpers.PointsToPixels` / `PixelsToPoints` (the Web filler already converted correctly and now shares the helper). Columns render at their true width and no longer drift across open/save cycles.
- **`kMoneyFormat` comma truncation (regression in 0.5.0).** The `#,##0.00` constant had an unescaped comma; the Xojo IDE silently truncates a String constant at an unescaped comma, so `kMoneyFormat` resolved to `#` — breaking `MoneyCell` and the `.Money` mutator. Escaped as `\x2C`.

## [0.5.0] - 2026-07-28

### Added
- **Formulas — read & write (XLSX).** Formula text is now preserved instead of being flattened to the cached value: reading keeps each cell's `<f>` (A1) text, and saving re-emits it — so formulas survive an open → edit → save → reopen round-trip.
  - **Fluent authoring API.** `XLSXCell` gains Shared factories `TextCell` / `NumberCell` / `MoneyCell` / `DateCell` / `DateTimeCell` / `BoolCell` / `FormulaCell` / `FormulaCellA1`, plus chainable style mutators (`.Bold` / `.Italic` / `.Underline` / `.Money` / `.Format(code)` / `.Align(h)` / `.BackColor` / `.FontColor` / `.FontFace` / `.FontSize`, each returning the cell). `XLSXSheet` gains an indexed accessor (`ws.Cell(r,c) = cell`) and placers (`PutText` / `PutNumber` / `PutMoney` / `PutDate` / `PutDateTime` / `PutBool` / `PutFormula` / `PutFormulaA1`) that place a cell and return it for chaining. `XLSXWorkbook.AddSheet(name)` creates and returns an empty sheet.
  - **R1C1 authoring.** `FormulaCell` / `PutFormula` accept Excel **R1C1 relative** notation (e.g. `SUM(R[+3]C:R[+1000]C)`, `RC[-2]`) and convert it to A1 at write time, anchored at the cell — so `SUM(R[+3]C:R[+1000]C)` placed at (1,6) is written `SUM(F4:F1001)`. A leading `=` is optional. `FormulaCellA1` / `PutFormulaA1` take plain A1, written verbatim.
  - **"Show formulas" toggle (Desktop + Web).** A checkbox flips the grid between cached values (default) and the `=…` formula text, per sheet. `XLSXCell.HasFormula` / `FormulaText` back it.
  - Out of scope this pass (parked): **ODS** formula read/write (saving to `.ods` still writes the cached value only), formula **evaluation** (cached value shown; we don't compute).
- **Cell styling (whole-cell): fonts, fills, alignment, borders — read, render (Desktop) and write.** A new `XLSXCellStyle` value object carries font (bold/italic/underline/name/size/color), solid fill background, horizontal/vertical alignment + wrap, and per-edge borders (None/Thin/Medium/Thick + color).
  - **Read (XLSX):** `XLSXStyles` parses `<fonts>` / `<fills>` / `<borders>` and resolves each cellXf to an `XLSXCellStyle` (`CellStyleAt`). `XLSXCell.ResolvedStyle` / `CellStyle`.
  - **Theme & indexed colours:** `XLSXStyles` now reads `xl/theme/theme1.xml` and resolves `<color theme="N" tint="…"/>` (with the Excel dk/lt index swap and the HSL tint algorithm, verified against Excel's swatch values) plus the legacy `<color indexed="N"/>` palette — so theme-coloured fills/fonts/borders render instead of being dropped.
  - **Excel Tables (ListObjects):** the reader parses `xl/tables/*.xml` (via each worksheet's `<tableParts>` + rels) into `XLSXTable` objects on the sheet. A built-in table style name (e.g. `TableStyleLight9`) is resolved against the theme into a synthetic header fill (solid accent + white bold) and banded-row tint by `XLSXStyles.TableStyleCell`, and `XLSXSheet.EffectiveStyle` applies it to table cells that have no explicit style of their own — so a table's coloured header + row stripes now render (e.g. the Microsoft financial sample) on Desktop and Web. The accent is derived from the style-name number (groups of seven). *Rendering only — table styling is not yet re-written on Save (the saved file keeps explicit cell styles but drops the table overlay).*
  - **Render (Desktop):** the listbox paints cell fills + borders (`PaintCellBackground`) and font + horizontal alignment (`PaintCellText`).
  - **Styled first row stays a data row:** a plain first row is still promoted to the column-header bar, but a *styled* first row (a designed header) is shown as the first data row with A/B/C column letters, so its fill/font renders natively on **both** Desktop and Web instead of being lost in the unstyleable header bar (`XLSX*ListboxFiller.PromotesHeader`).
  - **Write (XLSX):** the writer emits real `<fonts>` / `<fills>` / `<borders>` / `<cellXfs>` — one deduped xf per distinct appearance — so styling survives a save/round-trip (previously dropped).
  - **Write (ODS):** `ODSWriter` emits each cell style as a `<style:style>` with `<style:table-cell-properties>` (fill, per-edge borders, vertical align, wrap), `<style:paragraph-properties>` (horizontal align), and `<style:text-properties>` (font weight/style/underline/colour/size/family), keyed by the same appearance signature.
  - **Render (Web):** `XLSXWebListboxFiller` applies a `WebListBoxStyleRenderer`/`WebStyle` per styled cell (fill, text colour, bold/italic/underline, font name/size, uniform border). WebStyle has no alignment, so numeric right-align isn't reproduced on the Web.
  - **Viewing mode now trims only *trailing* empty rows** (the styled-bloat case) and keeps *interior* empty rows, so the on-screen grid matches the file and a saved copy of it (`XLSX*ListboxFiller.LastContentRow`).
  - New enums `XLSXEnums.eAlignH` / `eAlignV` / `eBorderStyle`. Styled test fixtures: `styled-basics` / `styled-borders` / `styled-combined.xlsx`, `theme-colors.xlsx`.
  - Out of scope (parked): rich text (per-run formatting within a cell), ODS style **read**, Excel Table **write** (re-emit the table overlay on Save), Web cell alignment.
- **Column widths & row heights — read, write, and preserve.** Opening a workbook now captures each column's width and each custom row height; saving writes them back (XLSX `<cols>` + row `ht`; ODS per-column/row styles), so round-tripping no longer flattens the layout. Cross-format conversion carries them too.
  - Model: `XLSXSheet.ColumnWidth` / `SetColumnWidth` / `RowHeight` / `SetRowHeight` (+ `HasColumnWidths` / `HasRowHeights`). Canonical unit is **points**.
  - Unit conversion helpers in `XLSXHelpers`: `ColumnCharsToPoints` / `ColumnPointsToChars` (Excel character units) and `OdsLengthToPoints` / `PointsToOdsLength` (cm/mm/in/pt/pc/px).
  - Desktop: the listbox seeds real column widths (falling back to the content heuristic for columns with none), and **user column-resizes are captured back into the model** (on Save and when switching sheets) so the saved file matches what's on screen. Web: `WebListBox.ColumnWidths` set from the real widths (points → pixels); capturing a Web user's resize is a later follow-up.
  - Note: neither listbox control supports per-row heights, so row heights are **preserved on round-trip** (model + both writers) but the on-screen grid uses a uniform row height.
- **1904 date system.** Workbooks authored with the 1904 epoch (common from older Mac Excel) declared via `<workbookPr date1904="1"/>` now read correctly — their date serials were previously ~4 years off. Detected on open and normalized to the 1900 system the model uses, so dates display and re-save correctly. (ODS is unaffected — it stores ISO dates.)

### Fixed
- **Spurious borders on styled cells from LibreOffice-saved files.** A border edge written as `<left style="none"/>` (how LibreOffice emits "no border"; Excel omits the attribute) was misread as a thin border, so any custom-painted cell (bold/italic/filled) picked up phantom borders. `"none"` now resolves to no border.

## [0.4.0] - 2026-06-10

### Added

- **Write `.xlsx` and `.ods` files** — save any open workbook in either format, with full cross-format conversion. `SpreadsheetWriter.Save(wb, file)` picks the serializer from the destination extension; `XLSXWriter` / `ODSWriter` also expose `ToMemoryBlock(wb)` for in-memory use.
  - XLSX: OPC parts with inline strings (no sharedStrings table), a custom numFmt (id 164+) per distinct format code, `<mergeCells>`, and `<dimension>` + `<sheetFormatPr defaultRowHeight="15">` so minimal renderers size rows correctly.
  - ODS: `mimetype` first and stored per the ODF packaging spec, `manifest.xml`, and `content.xml` with format codes converted back into `<number:*-style>` element trees (dates/times with month-vs-minutes disambiguation, percentage, currency `[$SYM]`, grouped numbers); ISO dates, `PTnHnMnS` durations, merge spans + covered cells.
  - Both assemble through `SpreadsheetZipWriter`, a pure-Xojo in-memory zip builder (raw deflate from `MemoryBlock.Compress`, own CRC-32, ordered entries with per-entry stored control). No temp files.
- **In-place cell editing (Desktop + Web)** — click a cell to edit; the commit writes back into the workbook model. Numeric input keeps the cell's type and number format (dates/currency re-render formatted), text becomes a string cell, blank empties it.
- **Create a sheet from scratch** — a **New** button opens an empty editable grid (column letters as headers); **+ Row / − Row / + Col / − Col** buttons grow/shrink the sheet (they work on opened files too, switching the view to the full grid). Backed by new model API: `XLSXSheet.AppendRow/RemoveLastRow/AppendColumn/RemoveLastColumn`, `XLSXHelpers.NewWorkbook`.
- **Save UX** — Desktop: **Save…** + native save dialog (on macOS the two file types appear as a *Format* popup; the chosen extension picks the format). Web: format picker + **Save…** pushing a browser download with the proper MIME type.
- Grid display mode in both listbox fillers (`showAllCells`), gridlines always visible, and editing API additions: `XLSXCell.FormatCode` (direct format code, wins over `StyleIndex`), `XLSXHelpers.XmlEscape` / `EffectiveFormatCode` / `WriteFormatCode` / `IsNumericString`, `XLSXFormatter.DateTimeToExcelSerial`.

### Fixed

- macOS **Quick Look** rendered huge row heights for data rows in saved `.xlsx` files; saved sheets now declare the default row height and a complete font. Excel/LibreOffice were unaffected either way.

### Compatibility

- No breaking changes — all reader APIs unchanged; writers and editing API are pure additions.

### Known limitations (writing)

- Formula text is not preserved (the cached value is written); cell colors / fonts / borders are not written; ODS number formats outside the supported subset save as unstyled values.

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
