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
| XLSX read | `ReadOptions` | Verified | Header/trim, missing and self-closing rows, physical merged-cell presence and shared-cache options covered |
| XLSX read | `QueryTable`, path/stream | Verified | Named table and case-insensitive table name |
| Metadata | Sheet names | Verified | Path, stream, sync and task-based async |
| Metadata | Column names | Verified | Header/headerless, sheet and start cell |
| Metadata | Sheet dimensions | Verified | Path/stream, declared ranges and missing-dimension parity |
| Metadata | Sheet information | Verified | ID, index, name, hidden state, active state and sheet type |
| Adapters | `QueryAsDataTable` | Verified | Single selected sheet; materialized managed adapter |
| Adapters | `GetReader` | Verified | Selected/all sheets, path/stream, materialized rows and `NextResult` verified |
| Async | Metadata tasks | Partial | Runs Rust operation on a worker; no in-flight native cancellation |
| Async | `IAsyncEnumerable<T>` query | Partial | XLSX/CSV path/stream dynamic/typed/table enumeration plus pre-cancellation verified; in-flight native batch cancellation remains |
| Typed read/write | POCO/attribute mapping | Verified | Properties/fields, index, readonly export, rename/alias/ignore, GUID, enum, nullable, culture, localization, dynamic overrides/formatters, ExcelFormat, errors and formula columns verified |
| Comments | Notes/threaded comments | Verified | Path/stream, authors, timestamps, replies, resolved state and legacy notes |
| CSV read | Dynamic query, path/stream | Verified | Header, delimiter, BOM, Unicode, quoted text and empty string |
| CSV metadata | Column names | Verified | Path, stream, sync and task-based async |
| CSV adapters | DataTable/Reader | Verified | Materialized managed adapters |
| Conversion | CSV/XLSX, path/stream | Verified | Both directions, header behavior and task wrappers verified |
| CSV write | Dynamic/typed save/append, path/stream | Verified | Delimiter, BOM, header, overwrite, append, bounded native spool and in-flight cancellation verified |
| XLSX write | Dynamic single-sheet `SaveAs`, path/stream | Partial | Scalars, temporal values, schema and full exposed styles verified; per-cell completion progress works, in-flight native progress remains |
| XLSX write | Dynamic multi-sheet, path/stream | Verified | Ordered sheets, per-sheet row counts, content and overwrite verified |
| XLSX write | Typed/async export | Verified | POCO attributes/scalars, bounded spool, native one-pass consumption, progress and in-flight cancellation verified |
| Workbook edits | Rename/reorder/visibility | Verified | Atomic path operations checked through C# and Rust metadata readers |
| Workbook edits | Dynamic insert/copy-and-add, path | Partial | Add/reject/replace/source preservation verified; stream and complex relationship policies remain |
| Templates | Path/stream/byte[] fill | Partial | All source/destination combinations, scalars, list expansion, strict missing variables and overwrite verified; advanced parity remains |
| Templates | `MergeSameCells`, path/stream/byte[] | Verified | Merge refs, marker removal, source preservation and overwrite verified |
| Pictures | AddPicture | Verified | PNG path/stream, OneCell/Absolute/TwoCell anchors, repeated drawing updates and package preservation verified |
| Fluent mapping | Exact-cell reads | Verified | Rust `CellMap` path/stream dynamic and typed adapters verified |
| Fluent mapping | Collections/write/template | Verified | Properties, formats, per-cell formulas, vertical spacing, nested collections, path/stream/async export and read, and style-preserving mapped templates verified; defined layouts are compared with the source oracle |
| Legacy facade | `MiniExcelLibs.MiniExcel` | Partial | Sync/async path/stream groups, ExcelType/configuration routing, object/POCO/dictionary/DataTable/DataSet writes and conversions delegate only to Rust; exact model/overload parity remains |

## Confirmed differences

1. CSV empty-as-null: the checked-out C# source applies the setting, while the pinned
   `2.0.0-preview.4` NuGet oracle returned an empty string in the exercised dynamic query.

These rows remain open until the intended C# source contract is selected and encoded as explicit
compatibility behavior. They must not be hidden by loosening equality assertions.