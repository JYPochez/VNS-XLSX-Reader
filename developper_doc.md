# VNS XLSX Reader — Developer Documentation

This document covers the public API of the `Common/` readers **and writers** shared by the **Desktop** and **Web 2.0** Xojo projects (`VNS-Desktop-XLSX_Reader` and `VNS-Web-XLSX-Reader`). Both `.xlsx` (Office Open XML) and `.ods` (OpenDocument Spreadsheet) are supported, for reading and writing, through one shared workbook model.

Everything is pure-Xojo (API 2.0 only — no plugins, no external libraries) and runs on macOS, Windows, Linux desktop, plus the Web 2.0 server runtime.

---

## Quick start

### Desktop — open

```xojo
Var dlg As New OpenFileDialog
Var t As New FileType
t.Name = "Spreadsheet"
t.Extensions = "xlsx;ods"
dlg.Filter = t
Var f As FolderItem = dlg.ShowModal(Self)
If f = Nil Then Return

Try
  ' SpreadsheetReader dispatches by extension: .ods -> ODSReader, else XLSXReader.
  Var wb As XLSXWorkbook = SpreadsheetReader.Open(f)
  For i As Integer = 1 To wb.SheetCount
    Var s As XLSXSheet = wb.SheetAt(i)
    System.DebugLog s.Name + " " + Str(s.RowCount) + "x" + Str(s.ColCount)
  Next
Catch ex As XLSXException
  ' ex.Code is XLSXEnums.eParseError; ex.Code.ToString gives a symbol name
  ' (Extends helper in XLSXHelpers).
  MessageBox "Could not read: " + ex.Code.ToString + " - " + ex.Detail
End Try
```

### Desktop — save

```xojo
Var dlg As New SaveFileDialog
Var tXlsx As New FileType
tXlsx.Name = "Excel Workbook"
tXlsx.Extensions = "xlsx"
Var tOds As New FileType
tOds.Name = "OpenDocument Spreadsheet"
tOds.Extensions = "ods"
dlg.Filter = tXlsx + tOds              ' macOS shows these as a Format popup
dlg.SuggestedFileName = "Workbook.xlsx"
Var f As FolderItem = dlg.ShowModal(Self)
If f = Nil Then Return

' The destination extension picks the serializer (.ods -> ODSWriter, else XLSXWriter).
SpreadsheetWriter.Save(wb, f)
```

A workbook opened from one format can be saved into the other — both writers consume the same `XLSXWorkbook` model. `XLSXHelpers.NewWorkbook(name, rows, cols)` builds an empty workbook from scratch that saves the same way.

### Web 2.0 — open

`WebFileUploader` does **not** auto-upload — you must call `StartUpload` to trigger the UploadFinished event. The single-file UX:

```xojo
' WebFileUploader.FileAdded
Sub FileAdded(filename As String, bytes As UInt64, mimeType As String)
  UploaderXLSX.StartUpload
End Sub

' WebFileUploader.UploadFinished
Sub UploadFinished(files() As WebUploadedFile)
  If files.Count = 0 Then Return
  Try
    ' pass the original filename so the dispatcher can pick xlsx vs ods
    Var wb As XLSXWorkbook = SpreadsheetReader.Open(files(0).File, files(0).Name)
    ' bind sheets to a WebTabPanel + WebListBox
  Catch ex As XLSXException
    Var d As New WebMessageDialog
    d.Title = "Cannot open file"
    d.Explanation = ex.Detail
    d.Show
  End Try
End Sub
```

`WebUploadedFile.File` returns a `FolderItem` already on disk, so the same `SpreadsheetReader.Open(file As FolderItem, …)` overload works on both platforms.

### Web 2.0 — save (browser download)

```xojo
' mDownloadFile must be a page/session property so the WebFile outlives the method.
Var data As MemoryBlock = XLSXWriter.ToMemoryBlock(wb)   ' or ODSWriter.ToMemoryBlock(wb)
mDownloadFile = New WebFile
mDownloadFile.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
mDownloadFile.ForceDownload = True
mDownloadFile.FileName = "Workbook.xlsx"
mDownloadFile.Data = data.StringValue(0, data.Size)
Self.GoToURL(mDownloadFile.URL)
```

---

## Architecture

```
                          READ                                          WRITE
FolderItem (Desktop) ─┐
                      ├─► SpreadsheetReader.Open(file, …, mode)   SpreadsheetWriter.Save(wb, file)
WebUploadedFile (Web)─┘        │ by extension                          │ by destination extension
                     ┌─────────┴─────────┐                   ┌─────────┴─────────┐
                     ▼                   ▼                    ▼                   ▼
              XLSXReader            ODSReader            XLSXWriter           ODSWriter
                     │                   │              (OPC parts:         (mimetype first
                     └───► XLSXZip ◄─────┘               workbook/styles/    + stored, manifest,
              (Memory: zip directory +                   sheetN.xml,         content.xml with
               MemoryBlock.Decompress;                   inline strings,     number:*-style
               Disk: FolderItem.Unzip                    numFmts 164+)       trees)
               into SpecialFolder.Temporary)                  │                   │
                     │                                        └────────┬──────────┘
                     ▼                                                 ▼
        XLSX: sharedStrings/styles/workbook/sheetN.xml    SpreadsheetZipWriter
        ODS:  content.xml (+ styles.xml)                  (raw deflate from MemoryBlock.Compress,
                     │                                     own CRC-32, ordered entries)
                     ▼
              ┌──────────────────────────────────────────┐
              │         XLSXWorkbook (shared model)      │
              │ SourceName / SharedStrings() / Styles    │
              │ SheetCount / SheetAt / SheetByName       │
              │   XLSXSheet (cells, merges, row/col ops) │
              │     XLSXCell (lazy DisplayText)          │
              └──────────────────────────────────────────┘
                                  │
                                  ▼
              ┌──────────────────────────────────────────┐
              │ XLSX*ListboxFiller.Fill(lb, sheet,       │
              │   wb.Styles [, showAllCells])            │
              │ one per project (DesktopListBox /        │
              │ WebListBox); grid mode for editing       │
              └──────────────────────────────────────────┘
```

Parsing is **eager** — every sheet is parsed during `Open`. `XLSXCell.DisplayText(styles)` is **lazy** and **cached** on first call. The writers serialize the same model back out, so read → edit → write round-trips, including cross-format conversion.

---

## Class & module reference

All shared parser code lives in `Common/`. Per-project UI fillers live in their respective projects.

### Common (shared between Desktop and Web)

| Symbol | Kind | Brief |
|---|---|---|
| `SpreadsheetReader.Open(file [, nameHint] [, mode])` | Module function | **Format-agnostic front door.** Dispatches `.ods` → `ODSReader`, else → `XLSXReader`. The UI calls this. |
| `XLSXReader.Open(file As FolderItem [, mode])` | Module function | XLSX entry; returns `XLSXWorkbook`; raises `XLSXException` on failure |
| `XLSXReader.Open(data As MemoryBlock, sourceName As String [, mode])` | Module function | Same, for in-memory input |
| `ODSReader.Open(file As FolderItem [, mode])` | Module function | OpenDocument (.ods) entry; returns the same `XLSXWorkbook` model |
| `ODSReader.Open(data As MemoryBlock, sourceName As String [, mode])` | Module function | Same, for in-memory input |
| `SpreadsheetWriter.Save(wb, file)` | Module function | **Format-agnostic save.** Destination extension picks the serializer: `.ods` → `ODSWriter`, else → `XLSXWriter` |
| `XLSXWriter.Save(wb, file)` / `.ToMemoryBlock(wb)` | Module functions | Serialize the workbook model to .xlsx (OPC zip; inline strings, custom numFmts 164+, mergeCells) |
| `ODSWriter.Save(wb, file)` / `.ToMemoryBlock(wb)` | Module functions | Serialize the workbook model to .ods (mimetype first + stored; format codes → `<number:*-style>` trees) |
| `SpreadsheetZipWriter` | Class | In-memory zip builder used by both writers: raw deflate from `MemoryBlock.Compress` (wrapper stripped), own CRC-32, ordered entries with per-entry stored control |
| `XLSXWorkbook` | Class | Owns sheets, shared strings, styles. The shared model for both formats, read and write. |
| `XLSXSheet` | Class | One sheet's cells + merged ranges, plus mutation: `PutCell`, `AddMergedRange`, `AppendRow/RemoveLastRow/AppendColumn/RemoveLastColumn` |
| `XLSXCell` | Class | One cell + lazy `DisplayText(styles)`; `FormatCode` (direct format, wins over `StyleIndex`) |
| `XLSXCellRange` | Class | Inclusive 1-based row/col range |
| `XLSXStyles` | Class | Parses `xl/styles.xml`; `NumberFormatCodeAt(xfIndex)` |
| `XLSXFormatter` | Module | `Format(rawValue, cellType, formatCode)`; `ExcelSerialToDateTime(d)` / `DateTimeToExcelSerial(dt)`; `IsDateFormatCode(s)` |
| `XLSXCellRef` | Module | A1 ↔ row/col helpers |
| `XLSXZip` | Class | Zip reading — Memory backend (zip directory + `MemoryBlock.Decompress`) or Disk (`FolderItem.Unzip` into a temp folder) |
| `XLSXEnums` | Module | `eCellType`, `eParseError`, `eOpenMode` (referenced with namespace prefix) |
| `XLSXHelpers` | Module | Enum `ToString`/`FromString` helpers; `NewWorkbook`, `XmlEscape`, `EffectiveFormatCode`, `WriteFormatCode`, `IsNumericString` |
| `XLSXException` | Class | Subclass of `RuntimeException`, carries `Code` + `Detail` |
| `strings` | Module | Localizable `kStr…` constants used by both UIs |

### Per-project (UI-only)

| Symbol | Project | Purpose |
|---|---|---|
| `XLSXDesktopListboxFiller.Fill(lb, sheet, styles [, showAllCells])` | Desktop | Pours a sheet into a `DesktopListBox`. Auto-sizes columns. |
| `XLSXWebListboxFiller.Fill(lb, sheet, styles [, showAllCells])` | Web | Pours a sheet into a `WebListBox`. |

Both fillers have two display modes:
- **Viewing** (default): first non-empty row (probed up to row 50) becomes the listbox header; rows where every cell renders empty are skipped; merged-cell followers render blank so values don't appear duplicated.
- **Grid** (`showAllCells = True`): column letters (A, B, C…) as headers and *every* row shown — used when editing or building a sheet from scratch. Each listbox row is tagged with its sheet row (`RowTagAt` on Desktop, `CellTagAt(row, 0)` on Web) so cell edits map back to the model.

---

## Public API details

### `SpreadsheetReader` / `XLSXReader` / `ODSReader` (Modules)

```xojo
' Format-agnostic front door — dispatches .ods -> ODSReader, else XLSXReader.
' nameHint lets Web pass the original uploaded filename when the on-disk
' FolderItem has a temp name. An XLSX parse failing with MissingPart is
' retried as ODS (wrong/missing extension).
Public Function SpreadsheetReader.Open(file As FolderItem, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook
Public Function SpreadsheetReader.Open(file As FolderItem, nameHint As String, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook
Public Function SpreadsheetReader.Open(data As MemoryBlock, sourceName As String, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook

' Per-format entries (same signatures on both):
Public Function XLSXReader.Open(file As FolderItem, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook
Public Function XLSXReader.Open(data As MemoryBlock, sourceName As String, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook
Public Function ODSReader.Open(...)   ' identical shape, parses content.xml into the same model
```

`mode` picks the zip-extraction backend (`XLSXEnums.eOpenMode`):
- `Auto` — Memory on Xojo 2024r3+, else Disk (default).
- `Memory` — pure in-memory parse + `MemoryBlock.Decompress`. No disk I/O, sandbox-friendly.
- `Disk` — `FolderItem.Unzip` into `SpecialFolder.Temporary`.

The returned workbook carries timing diagnostics: `wb.ZipMicroseconds`, `wb.XmlMicroseconds`, `wb.OpenMode`.

All raise `XLSXException` with `XLSXEnums.eParseError` set to:
- `NotAZip` — file's first 4 bytes are not `50 4B 03 04`
- `MissingPart` — file does not exist, or `xl/workbook.xml` (XLSX) / `content.xml` (ODS) not in the archive
- `MalformedXML` — `XmlDocument.LoadXml` failed on a required part
- `Encrypted` — reserved (we don't yet detect this; an encrypted OLE-wrapped XLSX would currently raise `NotAZip`)
- `Unsupported` — environment limitation (e.g. `SpecialFolder.Temporary` unavailable)

### `SpreadsheetWriter` / `XLSXWriter` / `ODSWriter` (Modules)

```xojo
' Format-agnostic save — destination extension picks the serializer:
' .ods -> ODSWriter, anything else -> XLSXWriter.
Public Sub SpreadsheetWriter.Save(wb As XLSXWorkbook, file As FolderItem)

' Per-format serializers; ToMemoryBlock feeds the Web download path.
Public Sub XLSXWriter.Save(wb As XLSXWorkbook, file As FolderItem)
Public Function XLSXWriter.ToMemoryBlock(wb As XLSXWorkbook) As MemoryBlock
Public Sub ODSWriter.Save(wb As XLSXWorkbook, file As FolderItem)
Public Function ODSWriter.ToMemoryBlock(wb As XLSXWorkbook) As MemoryBlock
```

Design notes:
- **XLSX**: emits `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml` (+ rels), `xl/styles.xml`, and one `xl/worksheets/sheetN.xml` per sheet. Strings are written **inline** (`t="inlineStr"`, no sharedStrings table). Every distinct format code in the workbook becomes a custom `numFmt` (id 164+) and a `cellXf`. Merged ranges emit `<mergeCells>`. Sheets declare `<dimension>` + `<sheetFormatPr defaultRowHeight="15">` so minimal renderers (macOS Quick Look) size rows correctly.
- **ODS**: emits `mimetype` **first and stored** (per the ODF packaging spec), `META-INF/manifest.xml`, and `content.xml`. Format codes are converted **back** into `<number:*-style>` element trees — the exact inverse of what `ODSReader` parses (date/time tokenizing with month-vs-minutes disambiguation, percentage, currency `[$SYM-…]` tags, plain numbers with decimals + grouping). Dates serialize as ISO strings, sub-day times as `PTnHnMnS` durations; merges as span attributes + `<table:covered-table-cell>`.
- **Zip container**: both writers assemble through `SpreadsheetZipWriter` — a pure-Xojo in-memory zip builder. Raw deflate is obtained by stripping the gzip/zlib wrapper off `MemoryBlock.Compress` output (entries store uncompressed when that fails or doesn't shrink); CRC-32 is computed in-class (the framework has none); entries are written in `AddPart` order with per-entry stored control. No disk staging, no zip64.
- **What writing preserves**: values, types, dates/times, booleans, number-format codes (the supported subset), merged ranges, sheet names. **Not preserved**: formula text (cached value is written), cell colors/fonts/borders, charts, images, conditional formatting — none of these exist in the model.

### `SpreadsheetZipWriter` (Class)

```xojo
Public Sub AddPart(name As String, content As String, forceStored As Boolean = False)
Public Function ToMemoryBlock() As MemoryBlock
Public Sub SaveTo(file As FolderItem)
```

Generic enough for any small zip-of-text-parts container. `forceStored` exists for the ODS `mimetype` entry.

### `XLSXWorkbook` (Class)

```xojo
Public Property SourceName As String                ' filename or "<memory>"
Public Property SharedStrings() As String           ' resolved shared-string table
Public Property Styles As XLSXStyles
Public Property OpenMode As XLSXEnums.eOpenMode     ' resolved backend
Public Property ZipMicroseconds As Double           ' time spent in XLSXZip.Open
Public Property XmlMicroseconds As Double           ' time spent on XML + sheet construction
Public Function SheetCount() As Integer
Public Function SheetAt(index As Integer) As XLSXSheet      ' 1-based; Nil out-of-range
Public Function SheetByName(name As String) As XLSXSheet    ' Nil if not found
Public Function SheetNames() As String()
```

### `XLSXSheet` (Class)

```xojo
Public Property Name As String
Public Property TabIndex As Integer                          ' 1-based
Public Function RowCount() As Integer
Public Function ColCount() As Integer
Public Function CellAt(row As Integer, col As Integer) As XLSXCell    ' 1-based; never Nil
Public Function MergedRangeCount() As Integer
Public Function MergedRangeAt(i As Integer) As XLSXCellRange
Public Function IsCellMergedFollower(row As Integer, col As Integer) As Boolean

' Mutation (editing / building):
Public Sub Constructor(name As String, tabIndex As Integer)   ' empty sheet, no XML parsing
Public Sub PutCell(row As Integer, col As Integer, cell As XLSXCell)  ' grows the extent
Public Sub AddMergedRange(firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
Public Sub AppendRow()                                        ' +1 empty row at the bottom
Public Sub RemoveLastRow()                                    ' drops cells + shrinks extent
Public Sub AppendColumn()
Public Sub RemoveLastColumn()
```

`CellAt` returns a shared empty sentinel for absent cells — no nil-check required at call sites. The empty constructor + `PutCell`/`AddMergedRange` are how `ODSReader` (and any non-XLSX source) populates the model; the row/col operations back the UI's structure buttons.

### `XLSXCell` (Class)

```xojo
Public Property eType As XLSXEnums.eCellType
Public Property RawString As String
Public Property StyleIndex As Integer                 ' -1 if absent
Public Property FormatCode As String                  ' direct format code; wins over StyleIndex
Public Function IsEmpty() As Boolean
Public Function NumberValue() As Double
Public Function DateValue() As DateTime               ' Nil unless eType=DateValue
Public Function BooleanValue() As Boolean
Public Function DisplayText(styles As XLSXStyles) As String   ' lazy + cached
```

`DisplayText` upgrades a numeric cell whose style carries a date format code to `DateValue` formatting (so `44621` with style `dd/mm/yyyy` displays the date, not the serial). `FormatCode` is how ODS-loaded and user-edited cells carry their format without an `XLSXStyles` entry; when non-empty it takes precedence over the `StyleIndex` lookup.

### `XLSXHelpers` (Module) — beyond the enum helpers

```xojo
Public Function NewWorkbook(sourceName As String, rows As Integer, cols As Integer) As XLSXWorkbook
Public Function XmlEscape(s As String) As String
Public Function EffectiveFormatCode(cell As XLSXCell, styles As XLSXStyles) As String
Public Function WriteFormatCode(cell As XLSXCell, styles As XLSXStyles) As String
Public Function IsNumericString(s As String) As Boolean
```

`NewWorkbook` builds an empty workbook (one sheet, rows×cols grid) for the "New" buttons. `EffectiveFormatCode` mirrors `DisplayText`'s resolution order; `WriteFormatCode` adds default ISO date codes for date cells with no explicit code (used by both writers). `IsNumericString` types user-edited cell text.

### `XLSXEnums.eCellType`

`Empty / Str / Number / DateValue / Bool / FormulaCached / ErrorVal`

Reference cases via the namespace: `XLSXEnums.eCellType.Number`. To get a symbolic string for logging:

```xojo
System.DebugLog "type: " + cell.eType.ToString    ' Extends method in XLSXHelpers
```

### `XLSXEnums.eParseError`

`NotAZip / MissingPart / MalformedXML / Encrypted / Unsupported`

### `XLSXEnums.eOpenMode`

`Auto / Memory / Disk` — the zip-extraction backend (see the reader section above).

All enums follow the same convention: `ToString` extension methods and `<EnumName>FromString` parsers live in `XLSXHelpers`.

---

## Format-code support

`XLSXFormatter` recognizes a pragmatic subset of Excel format codes. The numeric router walks several layers in order — simple-pattern match, scientific, accounting, currency-tag, then fallback — so adding new codes is mostly localized.

### Numeric formats

| Category | Examples | Notes |
|---|---|---|
| **Plain** | `0`, `0.00`, `#,##0`, `#,##0.00`, `General`, `""` (empty) | Direct dispatch in `FormatNumberValue`. |
| **Percent** | `0%`, `0.00%` | Multiplies by 100 then formats. |
| **Scientific** | `0.00E+00`, `0E+00`, `##0.0E+0` (Excel built-in id **48**) | Detected via `IsScientificFormat`; rendered via Xojo's framework `Format()`. |
| **Number with parens for negatives** | `#,##0 ;(#,##0)`, `#,##0.00;(#,##0.00)` (Excel built-in ids **37–40**) | The `[Red]` color hint inside these codes is stripped by `StripColorHints`. |
| **Accounting** | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` (Excel built-in ids **41–44**) | Detected via `IsAccountingFormat` (`_(` … `_)` spacer pattern). Renders as `$ 1,618.50` for positives, `($ 1,618.50)` for negatives, `$ -` (or `-` with no currency) for zero. |
| **Currency-tagged** | `[$$-409]#,##0.00`, `[$€-2]#,##0`, `#,##0.00\ [$€-2]` | `ExtractCurrencySymbol` parses the `[$X-Y]` tag. Position-aware — symbol is prefix if it appears before the numeric core, suffix otherwise. |
| **Text** | `@` | Cell value passes through unchanged. |

### Date / time formats

`yyyy-mm-dd`, `dd/mm/yyyy`, `mm/dd/yyyy`, `m/d/yy h:mm`, `m/d/yyyy h:mm`, `hh:mm`, `hh:mm:ss`, `yyyy-mm-dd hh:mm`.

The trailing `;@` text-section that Excel often appends (e.g. `m/d/yy h:mm;@`) is stripped before matching.

### Built-in numFmtId seeding

`XLSXStyles.SeedBuiltInNumFmts` populates the common Excel built-in IDs so workbooks that reference them without an explicit `<numFmt>` element still resolve:

- **0–22** — General, integer, decimal, thousands, percent, common date / time forms.
- **37–40** — `#,##0 ;(#,##0)` and variants (negative-in-parens).
- **41–44** — Accounting, with and without `$` prefix, with and without 2 decimals.
- **48** — Scientific (`##0.0E+0`).
- **49** — Text (`@`).

Custom codes declared in `<numFmts>` (id ≥ 164) override these.

### Adding a new format code

1. Decide whether it's date-shaped (contains y/m/d/h/s outside literals) or numeric.
2. **Date-shaped** → add a `Case` in `XLSXFormatter.FormatDateValue`. If detection needs more than y/m/d/h/s presence, extend `IsDateFormatCode`.
3. **Numeric** → pick a tier:
   - If it's a fixed string match (e.g. `0.0`, `#,###`), add a `Case` in `FormatNumberValue`'s primary `Select Case`.
   - If it's a family (e.g. another currency-tag form), add a detector + renderer alongside `IsScientificFormat` / `IsAccountingFormat` / `HasCurrencyTag`.
4. If the code is an Excel built-in id, seed it in `XLSXStyles.SeedBuiltInNumFmts` so workbooks that reference it without an explicit `<numFmt>` element resolve correctly.

Unknown codes fall back to `Double.ToString` for numbers or `DateTime.SQLDateTime` for dates — readable but not Excel-faithful.

---

## Limitations

- Cell **values** only — colors, fonts, borders, conditional formatting, theme/palette are ignored (reading and writing).
- Formulas show their **cached** value; we do not evaluate formulas, and saving writes the cached value (formula text is not kept in the model).
- No images, charts, pivots.
- No encrypted workbooks (CompoundDoc-wrapped XLSX would currently surface as `NotAZip`).
- No virtual / lazy listbox painting; suited to ~10k rows × ~50 cols per sheet. Larger workbooks may load slowly because every sheet is parsed eagerly during `Open`.
- `XLSXFormatter` covers a pragmatic format-code subset (see above); ODS writing converts that same subset back to `<number:*-style>` trees — codes outside it save as unstyled values.
- Frozen panes and column widths from the source workbook are **not** preserved in the listbox display.

---

## Test fixtures

| File | Notable for |
|---|---|
| `test_files/excelize-book1.xlsx` | 2 sheets, basic smoke |
| `test_files/excelize-sharedstrings.xlsx` | shared-strings table |
| `test_files/excelize-calcchain.xlsx` | formulas with cached values |
| `test_files/sheetjs-cdn-pres.xlsx` | small typical workbook |
| `test_files/microsoft-financial-sample.xlsx` | real-world workbook — accounting/currency formats, dates, ~700 rows |
| `test_files/ods-multi-sheet.ods` | ODS multi-sheet smoke |
| `test_files/ods-types.ods` | every ODS value type — string / float / date / time / date-time / boolean |
| `test_files/ods-sales.ods` | wider numeric grid |
| `test_files/ods-styled.ods` | `<number:*-style>` → format-code conversion (currency, number, percentage, date) + a merged cell |

When opening any of the test fixtures, you should see:
- The window/page title becomes `VNS XLSX Reader — <filename> [<sheet count>]`.
- One tab per sheet, listbox repopulates on tab change.
- Shared-string lookups produce text (not numeric indices).
- Merged-cell rectangles render the value once, in the top-left cell.

For writing: open a fixture, edit a cell, save as both `.xlsx` and `.ods`, and reopen each in the app and in Excel / LibreOffice — values, formats, and merges should survive (cross-format too).

---

## Versioning

See `version_history.md` for the per-commit log under `[Unreleased]` plus released versions. SemVer; `MAJOR=0` until V1 is feature-complete.

---

## Adding a new file to the Common pipeline

1. Author the `.xojo_code` file under `Common/` (text format — see existing files for the `#tag Module` / `#tag Class` skeleton).
2. Register it in **both** `.xojo_project` files: add a `Module=Name;../Common/Name.xojo_code;&hID;&hParentID;false` (or `Class=`) line under the `Folder=Common` entry, with a unique 16-hex-digit ItemID.
3. Open one project at a time in Xojo (text-format files cannot be edited concurrently from two project instances).

### Description and Note attributes

Xojo's text format expects `Description = …` attributes on `#tag Method` / `#tag Property` / `#tag Constant` / `#tag Enum` to be **hex-encoded UTF-8 with a trailing 0A byte**. Plain text in that slot corrupts the IDE display. **Never put `Description = …` on a `#tag Class` / `#tag Module` line** — the IDE then fails to parse the next `Inherits` line. Use `#tag Note, Name = …` blocks for class/module-level documentation instead.

If you want the same encoder/decoder/patcher tooling we used internally (drives a TSV/JSON spec into the right hex), open an issue and we can publish those helpers separately.
