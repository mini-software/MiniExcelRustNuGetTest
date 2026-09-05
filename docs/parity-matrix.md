# MiniExcel Rust parity matrix

Baseline captured on 2026-09-05:

- C# source: `D:\git\MiniExcel`, commit `5de51f7fc0edd99388faeb72184e3f5af80d1374`.
- Rust source: `D:\git\MiniExcel-Rust`, commit `905ad6189da92f91221798996ad2c837fb083147`.
- Package-test base: `D:\git\MiniExcelRust`, commit `12890be18a837e8dbe5babee3a040bae9ab347ae`.
- NuGet oracle used by package tests: MiniExcel `2.0.0-preview.4`.

Run against the checked-out source without modifying it:

```powershell
.\build\Test-Package.ps1 -Rid win-x64 -MiniExcelSourceRoot D:\git\MiniExcel
```

Status meanings: **Verified** has a managed-vs-Rust package test; **Partial** works with documented
limits; **Missing** has no production implementation yet.

| Area | API/capability | Status | Evidence or remaining work |
| --- | --- | --- | --- |
| XLSX read | Dynamic `Query`, path | Verified | Header/headerless, sheet, start cell, scalar values, Unicode |
| XLSX read | Dynamic `Query`, stream | Verified | `leaveOpen` and early disposal; currently stages to a temp file |
| XLSX read | `QueryRange`, path/stream | Verified | Inclusive A1 end cell |
| XLSX read | `ReadOptions` | Partial | Ignore missing rows and header behavior tested; merged-fill semantics differ |
| XLSX read | `QueryTable`, path/stream | Verified | Named table and case-insensitive table name |
| Metadata | Sheet names | Verified | Path, stream, sync and task-based async |
| Metadata | Column names | Verified | Header/headerless, sheet and start cell |
| Metadata | Sheet dimensions | Verified | Path/stream and normal declared OOXML dimensions |
| Metadata | Sheet information | Verified | ID, index, name, hidden state, active state and sheet type |
| Adapters | `QueryAsDataTable` | Verified | Single selected sheet; materialized managed adapter |
| Adapters | `GetReader` | Partial | Single selected sheet and materialized rows; no `NextResult` yet |
| Async | Metadata tasks | Partial | Runs Rust operation on a worker; no in-flight native cancellation |
| Async | `IAsyncEnumerable<T>` query | Missing | Requires cancellable native iterator and netstandard async interfaces |
| Typed read | POCO/attribute mapping | Partial | Properties, name aliases, GUID, enum, integer and nullable conversions verified; full attributes/culture/errors remain |
| Comments | Notes/threaded comments | Verified | Path/stream, authors, timestamps, replies, resolved state and legacy notes |
| CSV read | Dynamic query, path/stream | Verified | Header, delimiter, BOM, Unicode, quoted text and empty string |
| CSV metadata | Column names | Verified | Path, stream, sync and task-based async |
| CSV adapters | DataTable/Reader | Verified | Materialized managed adapters |
| CSV write | Dynamic save/append, path | Partial | Delimiter, BOM, header, overwrite and append verified; stream, typed and async remain |
| XLSX write | Dynamic single-sheet `SaveAs`, path/stream | Partial | Basic scalars and overwrite behavior verified; temporal values, schema, styles and multi-sheet remain |
| XLSX write | Typed/async/multi-sheet export | Missing | Rust core exists; input schema/callback ABI required |
| Workbook edits | Rename/reorder/visibility | Verified | Atomic path operations checked through C# and Rust metadata readers |
| Workbook edits | Insert/copy-and-add | Missing | Row/schema ABI and package-preservation tests required |
| Templates | Fill/merge | Missing | Rust core is partial; formula and relationship parity required |
| Pictures | AddPicture | Missing | Rust core implementation required |
| Fluent mapping | Read/write/template | Missing | Managed mapping plan plus Rust execution required |
| Legacy facade | `MiniExcelLibs.MiniExcel` | Missing | Must be added after behavior-level APIs stabilize |

## Confirmed differences

1. `FillMergedCells`: C# fills only explicitly represented empty cells in a merged range; Rust 0.4
   synthesizes cells that are absent from worksheet XML.
2. Self-closing empty row: C# currently emits `<row/>` even with `IgnoreEmptyRows=true`; Rust skips it.
3. Missing `<dimension>`: C# reports an empty range; Rust scans cells and returns the actual range.
4. CSV empty-as-null: the checked-out C# source applies the setting, while the pinned
   `2.0.0-preview.4` NuGet oracle returned an empty string in the exercised dynamic query.

These rows remain open until the intended C# source contract is selected and encoded as explicit
compatibility behavior. They must not be hidden by loosening equality assertions.