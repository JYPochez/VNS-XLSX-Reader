# VNS XLSX Reader — Developer Documentation

This document covers the public API of the `Common/` readers **and writers** shared by the **Desktop** and **Web 2.0** Xojo projects (`VNS-Desktop-XLSX_Reader` and `VNS-Web-XLSX-Reader`). Both `.xlsx` (Office Open XML) and `.ods` (OpenDocument Spreadsheet) are supported, for reading and writing, through one shared workbook model.

Everything is pure-Xojo (API 2.0 only — no plugins, no external libraries) and runs on macOS, Windows, Linux desktop, plus the Web 2.0 server runtime.

---

## Contents

- [Quick start](#quick-start)
  - [Desktop — open](#desktop--open)
  - [Desktop — save](#desktop--save)
  - [Web 2.0 — open](#web-20--open)
  - [Web 2.0 — save (browser download)](#web-20--save-browser-download)
- [Architecture](#architecture)
- [Class & module reference](#class--module-reference)
  - [Common (shared between Desktop and Web)](#common-shared-between-desktop-and-web)
  - [Per-project (UI-only)](#per-project-ui-only)
- [Public API details](#public-api-details)
  - [`SpreadsheetReader` / `XLSXReader` / `ODSReader` (Modules)](#spreadsheetreader--xlsxreader--odsreader-modules)
  - [`SpreadsheetWriter` / `XLSXWriter` / `ODSWriter` (Modules)](#spreadsheetwriter--xlsxwriter--odswriter-modules)
  - [`SpreadsheetZipWriter` (Class)](#spreadsheetzipwriter-class)
  - [`XLSXWorkbook` (Class)](#xlsxworkbook-class)
  - [`XLSXSheet` (Class)](#xlsxsheet-class)
  - [`XLSXCell` (Class)](#xlsxcell-class)
  - [`XLSXHelpers` (Module) — beyond the enum helpers](#xlsxhelpers-module--beyond-the-enum-helpers)
  - [`XLSXCellStyle` (Class) — whole-cell visual style](#xlsxcellstyle-class--whole-cell-visual-style)
  - [`XLSXTable` (Class) — Excel Table overlay](#xlsxtable-class--excel-table-overlay)
  - [`XLSXEnums.eAlignH` / `eAlignV` / `eBorderStyle`](#xlsxenumsealignh--ealignv--eborderstyle)
  - [`XLSXEnums.eCellType`](#xlsxenumsecelltype)
  - [`XLSXEnums.eParseError`](#xlsxenumseparseerror)
  - [`XLSXEnums.eOpenMode`](#xlsxenumseopenmode)
- [Format-code support](#format-code-support)
  - [Numeric formats](#numeric-formats)
  - [Date / time formats](#date--time-formats)
  - [Built-in numFmtId seeding](#built-in-numfmtid-seeding)
  - [Adding a new format code](#adding-a-new-format-code)
- [Limitations](#limitations)
- [Test fixtures](#test-fixtures)
- [Versioning](#versioning)
- [Adding a new file to the Common pipeline](#adding-a-new-file-to-the-common-pipeline)
  - [Description and Note attributes](#description-and-note-attributes)

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
| `XLSXSheet` | Class | One sheet's cells + merged ranges + Excel Tables, plus mutation: `PutCell`, `AddMergedRange`, `AppendRow/…`. `EffectiveStyle(row, col)` folds table styling into a cell's resolved style |
| `XLSXCell` | Class | One cell + lazy `DisplayText(styles)`; `FormatCode` (direct format); `ResolvedStyle(styles)` / `CellStyle` (visual style) |
| `XLSXCellStyle` | Class | Whole-cell visual style: font (bold/italic/underline/name/size/color), fill background, alignment, per-edge borders |
| `XLSXCellRange` | Class | Inclusive 1-based row/col range |
| `XLSXTable` | Class | An Excel Table (ListObject) overlaid on a sheet: range, header/totals counts, built-in style name + banding flags |
| `XLSXStyles` | Class | Parses `xl/styles.xml` + `xl/theme/theme1.xml`; `NumberFormatCodeAt(xfIndex)` + `CellStyleAt(xfIndex)`; resolves theme/indexed colours; `TableStyleCell(name, isHeader, striped)` for table styling |
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
| `XLSXDesktopListboxFiller.Fill(lb, sheet, styles [, showAllCells] [, showFormulas])` | Desktop | Pours a sheet into a `DesktopListBox`. Auto-sizes columns. `showFormulas` renders `=…` for formula cells. |
| `XLSXWebListboxFiller.Fill(lb, sheet, styles [, showAllCells] [, showFormulas])` | Web | Pours a sheet into a `WebListBox`. `showFormulas` renders `=…` for formula cells. |

Both fillers have two display modes:
- **Viewing** (default): the first non-empty row (probed up to row 50) becomes the listbox header **only when it is plain** — a *styled* first row (a designed/table header) is instead shown as the first data row with A/B/C column-letter headers (`PromotesHeader`), so its fill/font renders natively on Desktop *and* Web (the Web header bar can't be styled). Only **trailing** empty rows are trimmed (`LastContentRow`); interior empty rows are kept so the grid matches the file. Merged-cell followers render blank so values don't appear duplicated.
- **Grid** (`showAllCells = True`): column letters (A, B, C…) as headers and *every* row shown — used when editing or building a sheet from scratch. Each listbox row is tagged with its sheet row (`RowTagAt` on Desktop, `CellTagAt(row, 0)` on Web) so cell edits map back to the model.

Both fillers resolve each cell's style through `XLSXSheet.EffectiveStyle(row, col)`, so explicit cell styles, theme colours, and Excel-Table styling all flow through the same Desktop paint events / Web `WebListBoxStyleRenderer`.

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
- **What writing preserves**: values, types, dates/times, booleans, number-format codes (the supported subset), merged ranges, sheet names, whole-cell styling (fonts/fills/alignment/borders), column widths/row heights, and **formula text** (XLSX `<f>`; ODS `table:formula`). **Not preserved**: charts, images, conditional formatting, rich text — none of these exist in the model.

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
Public Sub AddSheet(sheet As XLSXSheet)                     ' append a pre-built sheet
Public Function AddSheet(name As String) As XLSXSheet       ' create + append an empty sheet, return it
```

### `XLSXSheet` (Class)

```xojo
Public Property Name As String
Public Property TabIndex As Integer                          ' 1-based
Public Function RowCount() As Integer
Public Function ColCount() As Integer
Public Function CellAt(row As Integer, col As Integer) As XLSXCell    ' 1-based; never Nil

' Fluent authoring — indexed accessor + placers (each Put* returns the placed cell):
Public Function Cell(row As Integer, col As Integer) As XLSXCell             ' getter: ws.Cell(r,c)
Public Sub Cell(row As Integer, col As Integer, Assigns c As XLSXCell)       ' setter: ws.Cell(r,c) = cell
Public Function PutText(row, col As Integer, text As String) As XLSXCell
Public Function PutNumber(row, col As Integer, value As Double) As XLSXCell
Public Function PutMoney(row, col As Integer, value As Double) As XLSXCell
Public Function PutDate(row, col As Integer, dt As DateTime) As XLSXCell
Public Function PutDateTime(row, col As Integer, dt As DateTime) As XLSXCell
Public Function PutBool(row, col As Integer, value As Boolean) As XLSXCell
Public Function PutFormula(row, col As Integer, formula As String) As XLSXCell     ' R1C1
Public Function PutFormulaA1(row, col As Integer, formula As String) As XLSXCell   ' A1

Public Function MergedRangeCount() As Integer
Public Function MergedRangeAt(i As Integer) As XLSXCellRange
Public Function IsCellMergedFollower(row As Integer, col As Integer) As Boolean

' Excel Tables + the style used to render a cell:
Public Sub AddTable(t As XLSXTable)
Public Function TableCount() As Integer
Public Function TableAt(i As Integer) As XLSXTable               ' Nil if out of range
Public Function EffectiveStyle(row As Integer, col As Integer) As XLSXCellStyle   ' never Nil

' Mutation (editing / building):
Public Sub Constructor(name As String, tabIndex As Integer)   ' empty sheet, no XML parsing
Public Sub PutCell(row As Integer, col As Integer, cell As XLSXCell)  ' grows the extent
Public Sub AddMergedRange(firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
Public Sub AppendRow()                                        ' +1 empty row at the bottom
Public Sub RemoveLastRow()                                    ' drops cells + shrinks extent
Public Sub AppendColumn()
Public Sub RemoveLastColumn()

' Column widths / row heights (canonical unit = points; 0 = sheet default):
Public Function ColumnWidth(col As Integer) As Double
Public Sub SetColumnWidth(col As Integer, widthPoints As Double)
Public Function RowHeight(row As Integer) As Double
Public Sub SetRowHeight(row As Integer, heightPoints As Double)
Public Function HasColumnWidths() As Boolean
Public Function HasRowHeights() As Boolean
```

Column widths and row heights are read from the source (XLSX `<cols>` + row `ht`; ODS column/row styles), preserved in the model, and written back by both serializers — so layout survives a round-trip and cross-format conversion. Widths are stored in **points**; `XLSXHelpers` provides the conversions (`ColumnCharsToPoints`/`ColumnPointsToChars` for Excel's character units, `OdsLengthToPoints`/`PointsToOdsLength` for ODS cm/in/pt/…). The Desktop/Web listboxes apply column widths; per-row heights aren't supported by either control, so row heights are preserve-only (not rendered per-row).

`CellAt` returns a shared empty sentinel for absent cells — no nil-check required at call sites. The empty constructor + `PutCell`/`AddMergedRange` are how `ODSReader` (and any non-XLSX source) populates the model; the row/col operations back the UI's structure buttons.

### `XLSXCell` (Class)

```xojo
Public Property eType As XLSXEnums.eCellType
Public Property RawString As String
Public Property StyleIndex As Integer                 ' -1 if absent
Public Property FormatCode As String                  ' direct format code; wins over StyleIndex
Public Property Formula As String                     ' formula text without the leading "="
Public Property FormulaIsR1C1 As Boolean              ' True → convert to A1 on write (authored R1C1)
Public Function IsEmpty() As Boolean
Public Function HasFormula() As Boolean               ' FormulaCached with non-empty Formula
Public Function FormulaText() As String               ' "=<Formula>" for display, else ""
Public Function NumberValue() As Double
Public Function DateValue() As DateTime               ' Nil unless eType=DateValue
Public Function BooleanValue() As Boolean
Public Function DisplayText(styles As XLSXStyles) As String   ' lazy + cached

' Fluent authoring — Shared factories (each returns a new XLSXCell):
Shared Function TextCell(text As String) As XLSXCell
Shared Function NumberCell(value As Double) As XLSXCell
Shared Function MoneyCell(value As Double) As XLSXCell          ' number + "#,##0.00"
Shared Function DateCell(dt As DateTime) As XLSXCell            ' serial + "yyyy-mm-dd"
Shared Function DateTimeCell(dt As DateTime) As XLSXCell        ' serial + "yyyy-mm-dd hh:mm"
Shared Function BoolCell(value As Boolean) As XLSXCell
Shared Function FormulaCell(formula As String) As XLSXCell      ' R1C1 (converted to A1 on write)
Shared Function FormulaCellA1(formula As String) As XLSXCell    ' A1 verbatim

' Fluent style mutators — each sets an aspect and returns Me for chaining:
Function Bold(on As Boolean = True) As XLSXCell
Function Italic(on As Boolean = True) As XLSXCell
Function Underline(on As Boolean = True) As XLSXCell
Function Format(code As String) As XLSXCell            ' number format code
Function Money() As XLSXCell                           ' "#,##0.00"
Function Align(h As XLSXEnums.eAlignH) As XLSXCell
Function BackColor(c As Color) As XLSXCell
Function FontColor(c As Color) As XLSXCell
Function FontFace(name As String) As XLSXCell
Function FontSize(points As Double) As XLSXCell
```

`DisplayText` upgrades a numeric cell whose style carries a date format code to `DateValue` formatting (so `44621` with style `dd/mm/yyyy` displays the date, not the serial). `FormatCode` is how ODS-loaded and user-edited cells carry their format without an `XLSXStyles` entry; when non-empty it takes precedence over the `StyleIndex` lookup.

**Formulas.** A formula cell keeps its formula text in `Formula` (read from the file in A1 notation; authored via the factories). `DisplayText` still returns the cached **value**; the UIs show the formula via `FormulaText` only when their **Show formulas** toggle is on. On write the formula is re-emitted — so formulas survive a save/reopen. Authored formulas may use Excel **R1C1 relative** notation (`FormulaCell` / `FormulaIsR1C1 = True`), converted to A1 at write time anchored at the cell (see `XLSXHelpers.FormulaToA1`); `FormulaCellA1` stores A1 verbatim.

**Fluent authoring example** (mirrors a "totals" sheet — see also `XLSXSheet.Put*` and `XLSXWorkbook.AddSheet`):

```xojo
Var wb As New XLSXWorkbook("Devis.xlsx")
Var ws As XLSXSheet = wb.AddSheet("Devis")
ws.PutText(1, 1, "Produit").Bold
ws.PutText(1, 4, "Total HT").Bold
ws.PutNumber(2, 2, 10)                                   ' Qté
ws.PutMoney(2, 3, 2.5)                                   ' PU
ws.PutFormula(2, 4, "RC[-2]*RC[-1]").Money               ' = B2*C2
ws.PutFormula(7, 4, "SUM(R[-5]C:R[-1]C)").Money.Bold     ' = SUM(D2:D6)
SpreadsheetWriter.Save(wb, outFile)
```

### `XLSXHelpers` (Module) — beyond the enum helpers

```xojo
Public Function NewWorkbook(sourceName As String, rows As Integer, cols As Integer) As XLSXWorkbook
Public Function XmlEscape(s As String) As String
Public Function EffectiveFormatCode(cell As XLSXCell, styles As XLSXStyles) As String
Public Function WriteFormatCode(cell As XLSXCell, styles As XLSXStyles) As String
Public Function IsNumericString(s As String) As Boolean
```

`NewWorkbook` builds an empty workbook (one sheet, rows×cols grid) for the "New" buttons. `EffectiveFormatCode` mirrors `DisplayText`'s resolution order; `WriteFormatCode` adds default ISO date codes for date cells with no explicit code (used by both writers). `IsNumericString` types user-edited cell text.

### `XLSXCellStyle` (Class) — whole-cell visual style

```xojo
Public Property Bold, Italic, Underline As Boolean
Public Property FontName As String
Public Property FontSize As Double          ' points, 0 = default
Public Property FontColor As Color          ' valid when HasFontColor
Public Property HasFontColor As Boolean
Public Property BackgroundColor As Color    ' valid when HasBackground
Public Property HasBackground As Boolean
Public Property AlignH As XLSXEnums.eAlignH
Public Property AlignV As XLSXEnums.eAlignV
Public Property WrapText As Boolean
Public Property BorderLeft/Right/Top/Bottom As XLSXEnums.eBorderStyle
Public Property BorderColor As Color        ' single colour for the cell's edges
Public Property HasBorderColor As Boolean
Public Function IsDefault() As Boolean      ' nothing set
Public Function HasAnyBorder() As Boolean
```

- **Read:** `XLSXStyles` parses `<fonts>` / `<fills>` (solid `fgColor`) / `<borders>` and assembles one `XLSXCellStyle` per `<cellXfs><xf>` (font + fill + border by id + `<alignment>`). Colours resolve from `rgb` (`AARRGGBB`/`RRGGBB`), **theme** (`theme="N"` + optional `tint`, via the `theme1.xml` palette with the Excel dk/lt index swap and the HSL tint algorithm), and the legacy **indexed** palette. Get a cell's style with `cell.ResolvedStyle(wb.Styles)` — a directly-set `cell.CellStyle` wins (for ODS/edited cells), else the `StyleIndex` → `XLSXStyles.CellStyleAt` lookup. Never Nil. To include Excel-Table styling, go through `XLSXSheet.EffectiveStyle(row, col)` (explicit cell style wins; otherwise the enclosing table's header/banded-row style applies unless the cell brings its own background).
- **Render (Desktop):** the filler stashes each cell's `EffectiveStyle` in the cell tag; `MainWindow.PaintCellBackground` draws fill + borders and `PaintCellText` draws font + horizontal alignment. A styled first row is rendered as a normal data row (see Viewing mode), so there is no styled header bar to paint.
- **Render (Web):** the filler attaches a `WebListBoxStyleRenderer(WebStyle, text)` per styled cell (fill, text colour, bold/italic/underline, font name/size, uniform border). `WebStyle` has no alignment, so numeric right-align isn't reproduced on the Web.
- **Write (XLSX):** `XLSXWriter.BuildStyleTables` walks every cell, dedups fonts/fills/borders/numFmts, and emits one `<xf>` per distinct appearance (keyed by a style signature), so styling round-trips through a save.
- **Write (ODS):** `ODSWriter` emits each cell style as a `<style:style>` with `<style:table-cell-properties>` (fill, per-edge borders, vertical align, wrap), `<style:paragraph-properties>` (horizontal align) and `<style:text-properties>` (font weight/style/underline/colour/size/family), keyed by the same appearance signature.
- **Excel Tables:** `XLSXStyles.TableStyleCell(styleName, isHeader, striped)` turns a built-in table style name (e.g. `TableStyleLight9`) into a synthetic style — header = solid accent fill + white bold, striped body row = light accent tint — with the accent derived from the style-name number (groups of seven) resolved against the theme. `XLSXSheet.EffectiveStyle` applies it; the reader parses `xl/tables/*.xml` into `XLSXTable`s. *Render-only — table styling is not re-emitted on Save.*
- **Not yet:** rich text (per-run formatting within one cell), ODS style **read**, Excel Table **write**, Web cell alignment.

### `XLSXTable` (Class) — Excel Table overlay

```xojo
Public Sub Constructor(name As String, firstRow, firstCol, lastRow, lastCol As Integer)
Public Property Name As String
Public Property FirstRow / FirstCol / LastRow / LastCol As Integer   ' 1-based, inclusive
Public Property HeaderRowCount As Integer = 1
Public Property TotalsRowCount As Integer = 0
Public Property StyleName As String                                  ' e.g. "TableStyleLight9"
Public Property ShowRowStripes / ShowColumnStripes As Boolean
Public Property ShowFirstColumn / ShowLastColumn As Boolean
Public Function Contains(row, col As Integer) As Boolean
Public Function IsHeaderRow(row As Integer) As Boolean
Public Function IsTotalsRow(row As Integer) As Boolean
Public Function FirstDataRow() As Integer
```

`XLSXReader` parses each worksheet's `<tableParts>` (resolved through the worksheet rels → `xl/tables/tableN.xml`) into `XLSXTable`s attached to the sheet via `AddTable`. The visual styling a table implies isn't stored per cell — it's generated from `StyleName` + the workbook theme at render time by `XLSXStyles.TableStyleCell` and applied through `XLSXSheet.EffectiveStyle`. Parsing never fails an open (a malformed table part is silently skipped). Render-only: tables aren't re-emitted on Save.

### `XLSXEnums.eAlignH` / `eAlignV` / `eBorderStyle`

`eAlignH`: General / Left / Center / Right. `eAlignV`: Bottom / Top / Middle. `eBorderStyle`: None / Thin / Medium / Thick (Excel's finer styles collapse to the nearest).

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

- **Whole-cell styling** (font / fill / alignment / border) is read (RGB + theme/tint + indexed colours), rendered on Desktop **and** Web, and written for XLSX **and** ODS. **Excel Tables** (ListObjects) are parsed and their built-in style is rendered (header fill + banded rows) on both platforms. Still out: **rich text** (per-run formatting within one cell), **ODS** style **read**, Excel Table **write** (re-emit the table overlay on Save), **Web** cell alignment, and conditional formatting.
- Formula **text** is preserved (read + written, for `.xlsx` and `.ods`) and shown on demand via the **Show formulas** toggle; the grid otherwise shows the **cached** value. We do **not** evaluate formulas — a freshly authored formula has no cached value until Excel/LibreOffice recomputes it on open.
- No images, charts, pivots.
- No encrypted workbooks (CompoundDoc-wrapped XLSX would currently surface as `NotAZip`).
- No virtual / lazy listbox painting; suited to ~10k rows × ~50 cols per sheet. Larger workbooks may load slowly because every sheet is parsed eagerly during `Open`.
- `XLSXFormatter` covers a pragmatic format-code subset (see above); ODS writing converts that same subset back to `<number:*-style>` trees — codes outside it save as unstyled values.
- Column widths from the source workbook **are** preserved (model + both serializers) and applied to the listbox; per-row heights are preserve-only (neither listbox renders per-row heights). Frozen panes are **not** preserved.

---

## Test fixtures

| File | Notable for |
|---|---|
| `test_files/excelize-book1.xlsx` | 2 sheets, basic smoke |
| `test_files/excelize-sharedstrings.xlsx` | shared-strings table |
| `test_files/excelize-calcchain.xlsx` | formulas with cached values (single row) |
| `test_files/formulas-sample.xlsx` | multi-row "Devis" with headers + 12 formulas (`*`, `SUM`, `AVERAGE`, `MAX`, `COUNT`) — exercises formula read/write and the Show-formulas toggle |
| `test_files/sheetjs-cdn-pres.xlsx` | small typical workbook |
| `test_files/microsoft-financial-sample.xlsx` | real-world workbook — accounting/currency formats, dates, ~700 rows |
| `test_files/ods-multi-sheet.ods` | ODS multi-sheet smoke |
| `test_files/ods-types.ods` | every ODS value type — string / float / date / time / date-time / boolean |
| `test_files/ods-sales.ods` | wider numeric grid |
| `test_files/ods-styled.ods` | `<number:*-style>` → format-code conversion (currency, number, percentage, date) + a merged cell |
| `test_files/formulas-sample.ods` | ODS `table:formula="of:=…"` (SUM/AVERAGE/MAX/COUNT + arithmetic) — exercises ODS formula read |

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
