# VNS XLSX Reader — Developer Documentation

This document covers the public API of the `Common/` parser shared by the **Desktop** and **Web 2.0** Xojo projects (`VNS-Desktop-XLSX_Reader` and `VNS-Web-XLSX-Reader`).

The parser is pure-Xojo (API 2.0 only — no plugins, no external libraries) and runs on macOS, Windows, Linux desktop, plus the Web 2.0 server runtime.

---

## Quick start

### Desktop

```xojo
Var dlg As New OpenFileDialog
Var t As New FileType
t.Name = "Excel Workbook"
t.Extensions = "xlsx"
dlg.Filter = t
Var f As FolderItem = dlg.ShowModal(Self)
If f = Nil Then Return

Try
  Var wb As XLSXWorkbook = XLSXReader.Open(f)
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

### Web 2.0

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
    Var wb As XLSXWorkbook = XLSXReader.Open(files(0).File)
    ' bind sheets to a WebTabPanel + WebListBox
  Catch ex As XLSXException
    Var d As New WebMessageDialog
    d.Title = "Cannot open file"
    d.Explanation = ex.Detail
    d.Show
  End Try
End Sub
```

`WebUploadedFile.File` returns a `FolderItem` already on disk, so the same `XLSXReader.Open(file As FolderItem)` overload works on both platforms.

---

## Architecture

```
FolderItem (Desktop dialog) ─┐
                             │
WebUploadedFile.File (Web) ──┴──► XLSXReader.Open(file, mode) ──► XLSXZip
                                                                   │
                                                       (Memory: parse zip directory + MemoryBlock.Decompress)
                                                       (Disk:   FolderItem.Unzip into SpecialFolder.Temporary)
                                                                   ▼
                                                              shared parts
                                  ┌─────────────┬──────┴──────┬──────────────┐
                                  ▼             ▼             ▼              ▼
                          sharedStrings   styles.xml    workbook.xml   sheetN.xml*
                          .xml String[]   XLSXStyles    sheet name +     XLSXSheet
                                                        rId targets         │
                                                                            ▼
                                                                       XLSXCell
                                                                  (lazy DisplayText)
                                  ┌──────────────────────────────────────────┐
                                  │         XLSXWorkbook                     │
                                  │ SourceName / SharedStrings() / Styles    │
                                  │ SheetCount / SheetAt / SheetByName       │
                                  └──────────────────────────────────────────┘
                                                    │
                                                    ▼
                              ┌──────────────────────────────────────────┐
                              │   XLSX*ListboxFiller.Fill(lb, sheet,     │
                              │                          wb.Styles)       │
                              │   one per project (DesktopListBox /       │
                              │                    WebListBox)            │
                              └──────────────────────────────────────────┘
```

Parsing is **eager** — every sheet is parsed during `XLSXReader.Open`. `XLSXCell.DisplayText(styles)` is **lazy** and **cached** on first call.

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
| `XLSXWorkbook` | Class | Owns sheets, shared strings, styles. The shared model for both formats. |
| `XLSXSheet` | Class | One sheet's cells + merged ranges |
| `XLSXCell` | Class | One cell + lazy `DisplayText(styles)` |
| `XLSXCellRange` | Class | Inclusive 1-based row/col range |
| `XLSXStyles` | Class | Parses `xl/styles.xml`; `NumberFormatCodeAt(xfIndex)` |
| `XLSXFormatter` | Module | `Format(rawValue, cellType, formatCode)`; `ExcelSerialToDateTime(d)`; `IsDateFormatCode(s)` |
| `XLSXCellRef` | Module | A1 ↔ row/col helpers |
| `XLSXZip` | Class | Cracks an .xlsx via `FolderItem.Unzip` into a temp folder |
| `XLSXEnums` | Module | `eCellType`, `eParseError` (referenced with namespace prefix) |
| `XLSXHelpers` | Module | Extension methods on the enums (e.g. `myEnum.ToString`) + `<EnumName>FromString` helpers |
| `XLSXException` | Class | Subclass of `RuntimeException`, carries `Code` + `Detail` |
| `strings` | Module | Localizable `kStr…` constants used by both UIs |

### Per-project (UI-only)

| Symbol | Project | Purpose |
|---|---|---|
| `XLSXDesktopListboxFiller.Fill(lb As DesktopListBox, sheet, styles)` | Desktop | Pours a sheet into a `DesktopListBox`. Auto-sizes columns. |
| `XLSXWebListboxFiller.Fill(lb As WebListBox, sheet, styles)` | Web | Pours a sheet into a `WebListBox`. |

Both fillers:
- Use the first non-empty row (probed up to row 50) as the listbox header.
- Skip rows where every cell renders empty.
- Render merged-cell followers as blank so values don't appear duplicated.

---

## Public API details

### `XLSXReader` (Module)

```xojo
Public Function Open(file As FolderItem, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook
Public Function Open(data As MemoryBlock, sourceName As String, mode As XLSXEnums.eOpenMode = Auto) As XLSXWorkbook
```

`mode` picks the zip-extraction backend:
- `Auto` — Memory on Xojo 2024r3+, else Disk (default).
- `Memory` — pure in-memory parse + `MemoryBlock.Decompress`. No disk I/O, sandbox-friendly.
- `Disk` — `FolderItem.Unzip` into `SpecialFolder.Temporary`.

The returned workbook carries timing diagnostics: `wb.ZipMicroseconds`, `wb.XmlMicroseconds`, `wb.OpenMode`.

Both raise `XLSXException` with `XLSXEnums.eParseError` set to:
- `NotAZip` — file's first 4 bytes are not `50 4B 03 04`
- `MissingPart` — file does not exist, or `xl/workbook.xml` not in the archive
- `MalformedXML` — `XmlDocument.LoadXml` failed on `workbook.xml`, `sharedStrings.xml`, `styles.xml`, or a `sheetN.xml`
- `Encrypted` — reserved (we don't yet detect this; an encrypted OLE-wrapped XLSX would currently raise `NotAZip`)
- `Unsupported` — environment limitation (e.g. `SpecialFolder.Temporary` unavailable)

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
```

`CellAt` returns a shared empty sentinel for absent cells — no nil-check required at call sites.

### `XLSXCell` (Class)

```xojo
Public Property eType As XLSXEnums.eCellType
Public Property RawString As String
Public Property StyleIndex As Integer                 ' -1 if absent
Public Function IsEmpty() As Boolean
Public Function NumberValue() As Double
Public Function DateValue() As DateTime               ' Nil unless eType=DateValue
Public Function BooleanValue() As Boolean
Public Function DisplayText(styles As XLSXStyles) As String   ' lazy + cached
```

`DisplayText` upgrades a numeric cell whose style carries a date format code to `DateValue` formatting (so `44621` with style `dd/mm/yyyy` displays the date, not the serial).

### `XLSXEnums.eCellType`

`Empty / Str / Number / DateValue / Bool / FormulaCached / ErrorVal`

Reference cases via the namespace: `XLSXEnums.eCellType.Number`. To get a symbolic string for logging:

```xojo
System.DebugLog "type: " + cell.eType.ToString    ' Extends method in XLSXHelpers
```

### `XLSXEnums.eParseError`

`NotAZip / MissingPart / MalformedXML / Encrypted / Unsupported`

Same ToString convention via `XLSXHelpers`.

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

## Limitations (V1 / 0.1.0)

- Read-only (no writing / saving).
- Cell **values** only — colors, fonts, borders, conditional formatting, theme/palette are ignored.
- Formulas show their **cached** value; we do not evaluate formulas.
- No images, charts, pivots.
- No encrypted workbooks (CompoundDoc-wrapped XLSX would currently surface as `NotAZip`).
- No virtual / lazy listbox painting; suited to ~10k rows × ~50 cols per sheet. Larger workbooks may load slowly because every sheet is parsed eagerly during `Open`.
- `XLSXFormatter` covers a small format-code subset (see above).
- Frozen panes and column widths from the source workbook are **not** preserved in the listbox display.

---

## Test fixtures

| File | Where | Notable for |
|---|---|---|
| `test_files/excelize-book1.xlsx` | tests | 2 sheets, basic smoke |
| `test_files/excelize-sharedstrings.xlsx` | tests | shared-strings table |
| `test_files/excelize-calcchain.xlsx` | tests | formulas with cached values |
| `test_files/sheetjs-cdn-pres.xlsx` | tests | small typical workbook |

When opening any of the test fixtures, you should see:
- The window/page title becomes `VNS XLSX Reader — <filename> [<sheet count>]`.
- One tab per sheet, listbox repopulates on tab change.
- Shared-string lookups produce text (not numeric indices).
- Merged-cell rectangles render the value once, in the top-left cell.

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
