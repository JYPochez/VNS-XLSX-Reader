# VNS XLSX Reader

[![Xojo](https://img.shields.io/badge/Xojo-2026r1-blue)](https://www.xojo.com)
[![Version](https://img.shields.io/badge/version-0.3.0-green)](version_history.md)
[![Platforms](https://img.shields.io/badge/platforms-macOS%20%7C%20Windows%20%7C%20Linux%20%7C%20Web-lightgrey)]()
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

> 💡 **Idea source:** this project grew out of the Xojo forum post [*Extracting xls/Excel file natively*](https://forum.xojo.com/t/extracting-xls-excel-file-natively/88210/1) — reading spreadsheet files with no plugins or third-party libraries.

A cross-platform `.xlsx` **and** `.ods` viewer written in **pure Xojo** (API 2.0, no plugins). Ships as a **Desktop** app for macOS / Windows / Linux and a **Web 2.0** browser app from the same shared parser.

Open any Excel or OpenDocument workbook → one tab per sheet → cell values rendered in a Listbox.

## Features

- 📂 **Open `.xlsx` and `.ods` files** via a native file dialog (Desktop) or browser upload (Web). One format-agnostic reader (`SpreadsheetReader`) picks the parser by extension.
- 📗 **OpenDocument (`.ods`) support** — the ODS reader parses `content.xml` into the same workbook model, converting OpenDocument `<number:*-style>` trees (date / number / currency / percentage) into the format codes the engine already understands.
- 📑 **One tab per sheet**, picked from the workbook's sheet order.
- 🔢 **Resolves cell types**: shared strings, numbers, booleans, errors, inline strings, formulas (cached values).
- 📅 **Excel format codes**: a pragmatic subset for numbers (`0`, `0.00`, `#,##0`, `#,##0.00`, `0%`, `0.00%`), **scientific** (`0.00E+00`), **accounting** (`_("$"* #,##0.00_)…` with parens for negatives), **currency tags** (`[$X-Y]`), and dates (`dd/mm/yyyy`, `yyyy-mm-dd`, `m/d/yy h:mm`, `hh:mm`, …); custom `numFmtId ≥ 164` from `styles.xml` honored.
- 🔁 **Merged cells**: top-left anchor renders the value, follower cells stay blank.
- 📏 **Auto-sized columns** with user-resizable dividers and horizontal scroll on the Desktop.
- 🌍 **Localizable strings** via Xojo Dynamic constants (the `strings` module).
- ⚠️ **Typed errors** (`XLSXException` with an `eParseError` code) so UI code can show friendly messages.
- 🔌 **Zero external dependencies** — uses only Xojo framework classes (`FolderItem.Unzip`, `XmlDocument`, `DateTime`).

## Screenshot

![VNS XLSX Reader desktop — Microsoft Financial Sample workbook open with the per-phase parse-time readout](screenshot.png)

The Desktop window above shows the "Read in memory" checkbox and the per-phase parse-time label (`Parsed in 220 ms (zip 4 + xml 216, Memory)`). Web has the same controls.

## Quick start

`SpreadsheetReader.Open` is the format-agnostic entry point — it dispatches `.xlsx` → `XLSXReader` and `.ods` → `ODSReader` and returns the same `XLSXWorkbook` either way. (Call `XLSXReader.Open` / `ODSReader.Open` directly if you want to force a format.)

```xojo
' Desktop — open .xlsx or .ods
Var f As FolderItem = ... ' from OpenFileDialog (filter: xlsx;ods)
Try
  Var wb As XLSXWorkbook = SpreadsheetReader.Open(f)
  System.DebugLog "Sheets: " + Str(wb.SheetCount)
  For i As Integer = 1 To wb.SheetCount
    Var s As XLSXSheet = wb.SheetAt(i)
    System.DebugLog s.Name + " — " + Str(s.RowCount) + " rows × " + Str(s.ColCount) + " cols"
  Next
Catch ex As XLSXException
  MessageBox "Could not read: " + ex.Code.ToString + " — " + ex.Detail
End Try
```

```xojo
' Web 2.0 — wire the uploader to auto-start, then handle the result
Sub FileAdded(filename As String, bytes As UInt64, mimeType As String)
  UploaderXLSX.StartUpload
End Sub

Sub UploadFinished(files() As WebUploadedFile)
  ' pass the original filename so the dispatcher can pick xlsx vs ods
  Var wb As XLSXWorkbook = SpreadsheetReader.Open(files(0).File, files(0).Name)
  ' bind sheets to a WebTabPanel + WebListBox
End Sub
```

Full API reference: [`developper_doc.md`](developper_doc.md).

## Repository layout

```
VNS-XLSX-Reader/
├── Common/                          ← shared parser (UI-free)
│   ├── SpreadsheetReader.xojo_code  ← format-agnostic front door (.xlsx / .ods)
│   ├── XLSXReader.xojo_code         ← .xlsx entry point
│   ├── ODSReader.xojo_code          ← .ods entry point (same workbook model)
│   ├── XLSXWorkbook.xojo_code       ← workbook aggregate
│   ├── XLSXSheet.xojo_code          ← sheet model + parser
│   ├── XLSXCell.xojo_code           ← cell + lazy DisplayText
│   ├── XLSXStyles.xojo_code         ← styles.xml parser
│   ├── XLSXFormatter.xojo_code      ← number/date format codes
│   ├── XLSXZip.xojo_code            ← framework FolderItem.Unzip wrapper
│   ├── XLSXEnums.xojo_code          ← eCellType / eParseError
│   ├── XLSXHelpers.xojo_code        ← Extends-based ToString helpers
│   ├── XLSXException.xojo_code      ← typed exception
│   ├── XLSXCellRange.xojo_code      ← merged-range value type
│   ├── XLSXCellRef.xojo_code        ← A1 ↔ row/col helpers
│   └── strings.xojo_code            ← localizable kStr… constants
├── VNS-Desktop-XLSX_Reader/         ← Desktop app (DesktopTabPanel + DesktopListBox)
├── VNS-Web-XLSX-Reader/             ← Web 2.0 app (WebFileUploader + WebTabPanel + WebListBox)
├── test_files/                      ← .xlsx and .ods samples for testing
├── developper_doc.md                ← developer API reference
├── version_history.md               ← per-release changelog
├── README.md
└── LICENSE
```

The two `.xojo_project` files reference every `Common/*.xojo_code` via `Module=` / `Class=` lines pointing at `../Common/`. Edit the source once, both projects pick it up.

## Building from source

**Requirements:** Xojo IDE **2026r1** or later.

1. Clone the repo:

   ```bash
   git clone https://github.com/<your-org>/VNS-XLSX-Reader.git
   cd VNS-XLSX-Reader
   ```

2. Open one of the `.xojo_project` files in Xojo:
   - Desktop: `VNS-Desktop-XLSX_Reader/VNS-Desktop-XLSX_Reader.xojo_project`
   - Web: `VNS-Web-XLSX-Reader/VNS-Web-XLSX-Reader.xojo_project`

3. **Run** (⌘R) to debug, or **Build** for release.

> ⚠️ **Don't open both projects in Xojo at the same time** — Xojo's text-format files cannot be safely co-edited from two IDE instances; saving one would clobber the shared `Common/` files. Open them serially.

## Limitations (V1 / 0.1.0)

| Out of scope | Workaround / status |
|---|---|
| Writing / saving an XLSX | Read-only |
| Formula evaluation | Cached values are shown as-is |
| Cell colors / fonts / borders | Values only — no styling fidelity |
| Images, charts, pivots | Ignored |
| Conditional formatting | Ignored |
| Encrypted (OLE-wrapped) workbooks | Surface as `NotAZip` |
| Virtual / lazy listbox painting | Suited to ~10k rows × 50 cols per sheet |
| Format codes outside the V1 subset | Fall back to default `Double.ToString` / `DateTime.SQLDateTime` |

See [`developper_doc.md`](developper_doc.md) for the full feature matrix and how to extend the format-code subset.

## Test fixtures

[`test_files/`](test_files/) ships small `.xlsx` samples (from the test data of popular open-source XLSX libraries) plus generated `.ods` fixtures, covering:
- multi-sheet smoke testing, shared-strings resolution, formulas with cached values;
- ODS value types (string / number / date / time / boolean) and `<number:*-style>` → format-code conversion (currency, number, percentage, date) with a merged cell.

See [`test_files/README.md`](test_files/README.md) for the per-fixture mapping.

## Versioning

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). See [`version_history.md`](version_history.md).

## Contributing

Pull requests welcome. A few non-obvious Xojo gotchas worth knowing before you edit `.xojo_code` files by hand:

- ❌ **Never put `Description = …` on `#tag Class` / `#tag Module`** — it breaks the IDE's `Inherits` parser. Use a `#tag Note` instead.
- ❌ **Never call `.ToString` on a parenthesized intrinsic expression** (e.g. `(a + b).ToString` fails to compile in API 2). Use `Str(...)` or extract to a typed local first.
- ❌ **Don't end a `Module=…` manifest line with `;true`** — the IDE silently drops it on load. Use `;false`. The `Extends` mechanism makes extension methods globally callable regardless of the "Global module" flag.
- ✅ **Edit only one `.xojo_project` at a time** — text-format files cannot be safely co-edited from two IDE instances.

## License

[MIT](LICENSE) — Copyright © 2026 VeryNiceSW.

## Credits

Built by **VeryNiceSW** (`fr.verynicesw.vns…`).

Test fixtures are public samples from popular open-source XLSX libraries — see [`test_files/README.md`](test_files/README.md) for sources.
