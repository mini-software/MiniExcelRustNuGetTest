using System.Diagnostics;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using MiniExcelLib;
using MiniExcelLib.Csv;
using MiniExcelLib.OpenXml;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using ManagedMiniExcel = MiniExcelLib.MiniExcel;

if (args.Length == 0)
  return RunSuite(lifecycleIterations: 1_000, maxPrivateGrowthMb: 32);

return args[0].ToLowerInvariant() switch
{
  "suite" => RunSuite(
    args.Length >= 2 ? int.Parse(args[1], CultureInfo.InvariantCulture) : 1_000,
    args.Length >= 3 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 32),
  "comments" => VerifyCommentsParity(args),
  "merge" => VerifyMergeSameCells(args),
  "verify" => VerifyFileParity(args),
  "generate" => GenerateBenchmarkWorkbook(args),
  "managed" => Benchmark(args, useRust: false),
  "rust" => Benchmark(args, useRust: true),
  _ => Usage()
};

static int VerifyMergeSameCells(string[] arguments)
{
  if (arguments.Length != 2)
    return Usage();

  var sourcePath = Path.GetFullPath(arguments[1]);
  var destinationPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-merged-{Guid.NewGuid():N}.xlsx");
  var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(sourcePath)));
  try
  {
    MiniExcelRust.MergeSameCells(destinationPath, sourcePath);
    Require(
      ReadMergeReferences(destinationPath).SequenceEqual(new[] { "A2:A4", "C3:C4", "A7:A8" }, StringComparer.Ordinal),
      "merge-same-cells: generated ranges differ.");
    Require(
      MiniExcelRust.Query(destinationPath).SelectMany(row => row.Values).All(value => value is not "@merge" and not "@endmerge"),
      "merge-same-cells: marker values remain in output.");
    var sourceHashAfter = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(sourcePath)));
    Require(sourceHash == sourceHashAfter, "merge-same-cells: source workbook changed.");

    var rejectedOverwrite = false;
    try
    {
      MiniExcelRust.MergeSameCells(destinationPath, sourcePath);
    }
    catch (InvalidOperationException)
    {
      rejectedOverwrite = true;
    }
    Require(rejectedOverwrite, "merge-same-cells: overwrite=false should reject an existing destination.");
    MiniExcelRust.MergeSameCells(destinationPath, sourcePath, overwriteFile: true);

    using var outputStream = new MemoryStream();
    MiniExcelRust.MergeSameCells(outputStream, File.ReadAllBytes(sourcePath), leaveOpen: true);
    Require(outputStream.CanWrite, "merge-same-cells-stream: leaveOpen should preserve the stream.");
    outputStream.Position = 0;
    Require(
      ReadMergeReferencesFromStream(outputStream).SequenceEqual(new[] { "A2:A4", "C3:C4", "A7:A8" }, StringComparer.Ordinal),
      "merge-same-cells-stream: generated ranges differ.");
    Console.WriteLine("Verified merge-same-cells output and source preservation.");
    return 0;
  }
  finally
  {
    if (File.Exists(destinationPath))
      File.Delete(destinationPath);
  }
}

static List<string> ReadMergeReferences(string path)
{
  using var stream = File.OpenRead(path);
  return ReadMergeReferencesFromStream(stream);
}

static List<string> ReadMergeReferencesFromStream(Stream stream)
{
  using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
  var entry = archive.GetEntry("xl/worksheets/sheet1.xml")
    ?? throw new InvalidDataException("The workbook has no first worksheet.");
  using var entryStream = entry.Open();
  var document = XDocument.Load(entryStream);
  XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  return document.Descendants(spreadsheet + "mergeCell")
    .Select(element => (string?)element.Attribute("ref") ?? string.Empty)
    .ToList();
}

static int VerifyCommentsParity(string[] arguments)
{
  if (arguments.Length is < 2 or > 3)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var sheetName = arguments.Length == 3 ? arguments[2] : null;
  var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
  var managed = importer.RetrieveComments(path, sheetName);
  var rust = MiniExcelRust.RetrieveComments(path, sheetName);
  Require(string.Equals(managed.SheetName, rust.SheetName, StringComparison.OrdinalIgnoreCase), "comments: sheet name differs.");
  Require(managed.Comments.Count == rust.Comments.Count, "comments: threaded comment count differs.");
  Require(managed.Notes.Count == rust.Notes.Count, "comments: note count differs.");

  for (var index = 0; index < managed.Comments.Count; index++)
  {
    var expected = managed.Comments[index];
    var actual = rust.Comments[index];
    Require(expected.Id == actual.Id, $"comments: id differs at {index}.");
    Require(expected.ReferenceCell == actual.ReferenceCell, $"comments: cell differs at {index}.");
    Require(expected.Resolved == actual.Resolved, $"comments: resolved differs at {index}.");
    Require(expected.Text == actual.Text, $"comments: text differs at {index}.");
    Require(expected.CreatedAt == actual.CreatedAt, $"comments: timestamp differs at {index}.");
    CompareAuthors(expected.Author, actual.Author, $"comments[{index}].author");
    Require(expected.Replies.Count == actual.Replies.Count, $"comments: reply count differs at {index}.");
    for (var replyIndex = 0; replyIndex < expected.Replies.Count; replyIndex++)
    {
      var expectedReply = expected.Replies[replyIndex];
      var actualReply = actual.Replies[replyIndex];
      Require(expectedReply.Id == actualReply.Id, $"comments: reply id differs at {index},{replyIndex}.");
      Require(expectedReply.ParentId == actualReply.ParentId, $"comments: parent id differs at {index},{replyIndex}.");
      Require(expectedReply.Text == actualReply.Text, $"comments: reply text differs at {index},{replyIndex}.");
      Require(expectedReply.CreatedAt == actualReply.CreatedAt, $"comments: reply timestamp differs at {index},{replyIndex}.");
      CompareAuthors(expectedReply.Author, actualReply.Author, $"comments[{index}].replies[{replyIndex}].author");
    }
  }

  var missingNoteIds = 0;
  for (var index = 0; index < managed.Notes.Count; index++)
  {
    var expected = managed.Notes[index];
    var actual = rust.Notes[index];
    Require(expected.ReferenceCell == actual.ReferenceCell, $"comments: note cell differs at {index}.");
    Require(expected.Author == actual.Author, $"comments: note author differs at {index}.");
    Require(expected.Text == actual.Text, $"comments: note text differs at {index}.");
    if (actual.Id is null)
      missingNoteIds++;
    else
      Require(expected.Id == actual.Id, $"comments: note id differs at {index}.");
  }

  using var stream = File.OpenRead(path);
  var streamResult = MiniExcelRust.RetrieveComments(stream, sheetName, leaveOpen: true);
  Require(streamResult.Comments.Count == rust.Comments.Count, "comments-stream: comment count differs.");
  Require(stream.CanRead, "comments-stream: leaveOpen should preserve the stream.");
  Console.WriteLine($"Verified comments for {rust.SheetName}; missing Rust legacy-note IDs: {missingNoteIds}.");
  return 0;
}

static void CompareAuthors(
  MiniExcelLib.OpenXml.Models.Author? expected,
  MiniExcelRustCommentAuthor? actual,
  string scenario)
{
  Require((expected is null) == (actual is null), $"{scenario}: presence differs.");
  if (expected is null || actual is null)
    return;
  Require(expected.Id == actual.Id, $"{scenario}: id differs.");
  Require(expected.DisplayName == actual.DisplayName, $"{scenario}: display name differs.");
  Require(expected.ProviderId == actual.ProviderId, $"{scenario}: provider id differs.");
}

static int RunSuite(int lifecycleIterations, int maxPrivateGrowthMb)
{
  var workbookPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-{Guid.NewGuid():N}.xlsx");
  var csvPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-{Guid.NewGuid():N}.csv");
  try
  {
    CreateWorkbook(workbookPath);
    File.WriteAllText(csvPath, "Name;Note\r\nalpha;\"Taiwan 台灣\"\r\nbeta;\r\n", new UTF8Encoding(true));
    VerifyParity(workbookPath);
    VerifyCsvParity(csvPath);
    VerifySaveAs();
    VerifyMultiSheetSaveAs();
    VerifyCsvWrite();
    VerifyTypedConversions();
    VerifyInsertAndCopy();
    VerifyTemplateFill();
    VerifyWorkbookMutations(workbookPath);
    VerifyLifecycle(workbookPath, lifecycleIterations, maxPrivateGrowthMb);
    Console.WriteLine("MiniExcelRust parity and lifecycle suite passed.");
    return 0;
  }
  finally
  {
    if (File.Exists(workbookPath))
      File.Delete(workbookPath);
    if (File.Exists(csvPath))
      File.Delete(csvPath);
  }
}

static void VerifySaveAs()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-write-{Guid.NewGuid():N}.xlsx");
  var rows = new List<IDictionary<string, object?>>
  {
    new Dictionary<string, object?>
    {
      ["Name"] = "alpha",
      ["Value"] = 42d,
      ["Enabled"] = true,
      ["When"] = new DateTime(2026, 9, 6, 12, 34, 56, 789),
      ["Duration"] = TimeSpan.FromMilliseconds(3723004),
      ["Date"] = new DateOnly(2026, 9, 6),
      ["Time"] = new TimeOnly(12, 34, 56, 789)
    },
    new Dictionary<string, object?>
    {
      ["Name"] = "beta",
      ["Value"] = null,
      ["Enabled"] = false,
      ["When"] = null,
      ["Duration"] = null,
      ["Date"] = null,
      ["Time"] = null
    }
  };
  try
  {
    var written = MiniExcelRust.SaveAs(path, rows, sheetName: "Exported");
    Require(written == rows.Count, $"save-as: expected {rows.Count} written rows, received {written}.");
    var managedRows = QueryManaged(path, true, "Exported").ToList();
    var rustRows = MiniExcelRust.Query(path, true, "Exported").ToList();
    CompareRows(managedRows, rustRows, "save-as-roundtrip");

    var rejectedExistingFile = false;
    try
    {
      MiniExcelRust.SaveAs(path, rows);
    }
    catch (InvalidOperationException)
    {
      rejectedExistingFile = true;
    }
    Require(rejectedExistingFile, "save-as: overwrite=false should reject an existing file.");

    written = MiniExcelRust.SaveAs(path, rows, sheetName: "Exported", overwriteFile: true);
    Require(written == rows.Count, "save-as: overwrite=true did not rewrite the workbook.");

    using (var stream = new MemoryStream())
    {
      written = MiniExcelRust.SaveAs(stream, rows, sheetName: "Streamed", leaveOpen: true);
      Require(written == rows.Count, "save-as-stream: row count differs.");
      Require(stream.CanWrite, "save-as-stream: leaveOpen should preserve the stream.");
      stream.Position = 0;
      var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
      var managedStreamRows = importer.Query(stream, true, "Streamed", leaveOpen: true)
        .Cast<IDictionary<string, object?>>()
        .ToList();
      CompareRows(managedRows, managedStreamRows, "save-as-stream");
    }

    var closingStream = new MemoryStream();
    MiniExcelRust.SaveAs(closingStream, rows);
    Require(!closingStream.CanWrite, "save-as-stream: the default should close the stream.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifyMultiSheetSaveAs()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-multisheet-{Guid.NewGuid():N}.xlsx");
  var firstRows = new List<IDictionary<string, object?>>
  {
    new Dictionary<string, object?> { ["Name"] = "one", ["Value"] = 1d },
    new Dictionary<string, object?> { ["Name"] = "two", ["Value"] = 2d }
  };
  var secondRows = new List<IDictionary<string, object?>>
  {
    new Dictionary<string, object?> { ["Code"] = "A", ["Enabled"] = true }
  };
  var sheets = new[]
  {
    new KeyValuePair<string, IEnumerable<IDictionary<string, object?>>>("First", firstRows),
    new KeyValuePair<string, IEnumerable<IDictionary<string, object?>>>("Second", secondRows)
  };
  try
  {
    var counts = MiniExcelRust.SaveAs(path, sheets);
    Require(counts.SequenceEqual(new[] { 2, 1 }), "multi-sheet-save: row counts differ.");
    Require(
      MiniExcelRust.GetSheetNames(path).SequenceEqual(new[] { "First", "Second" }, StringComparer.Ordinal),
      "multi-sheet-save: sheet order differs.");
    CompareRows(firstRows, QueryManaged(path, true, "First").ToList(), "multi-sheet-first");
    CompareRows(secondRows, QueryManaged(path, true, "Second").ToList(), "multi-sheet-second");

    var rejectedOverwrite = false;
    try
    {
      MiniExcelRust.SaveAs(path, sheets);
    }
    catch (InvalidOperationException)
    {
      rejectedOverwrite = true;
    }
    Require(rejectedOverwrite, "multi-sheet-save: overwrite=false should reject an existing file.");
    counts = MiniExcelRust.SaveAs(path, sheets, overwriteFile: true);
    Require(counts.SequenceEqual(new[] { 2, 1 }), "multi-sheet-save: overwrite counts differ.");

    using var stream = new MemoryStream();
    counts = MiniExcelRust.SaveAs(stream, sheets, leaveOpen: true);
    Require(counts.SequenceEqual(new[] { 2, 1 }) && stream.CanWrite, "multi-sheet-stream: write failed.");
    stream.Position = 0;
    Require(
      MiniExcelRust.GetSheetNames(stream, leaveOpen: true).SequenceEqual(new[] { "First", "Second" }, StringComparer.Ordinal),
      "multi-sheet-stream: sheet order differs.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifyCsvWrite()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-write-{Guid.NewGuid():N}.csv");
  var initialRows = new List<IDictionary<string, object?>>
  {
    new Dictionary<string, object?> { ["Name"] = "alpha", ["Value"] = 42d },
    new Dictionary<string, object?> { ["Name"] = "台灣", ["Value"] = string.Empty }
  };
  var appendedRows = new List<IDictionary<string, object?>>
  {
    new Dictionary<string, object?> { ["Name"] = "omega", ["Value"] = -7.5d }
  };
  var writeOptions = new MiniExcelRustCsvWriteOptions { Delimiter = ';' };
  var readOptions = new MiniExcelRustCsvReadOptions { Delimiter = ';' };
  try
  {
    var written = MiniExcelRust.SaveAsCsv(path, initialRows, writeOptions);
    Require(written == initialRows.Count, "csv-write: initial row count differs.");

    var rejectedExistingFile = false;
    try
    {
      MiniExcelRust.SaveAsCsv(path, initialRows, writeOptions);
    }
    catch (InvalidOperationException)
    {
      rejectedExistingFile = true;
    }
    Require(rejectedExistingFile, "csv-write: overwrite=false should reject an existing file.");

    written = MiniExcelRust.AppendCsv(path, appendedRows, writeOptions);
    Require(written == appendedRows.Count, "csv-append: appended row count differs.");

    var importer = ManagedMiniExcel.Importers.GetCsvImporter();
    var managedRows = importer.Query(path, true, new CsvConfiguration { Seperator = ';' })
      .Cast<IDictionary<string, object?>>()
      .ToList();
    var rustRows = MiniExcelRust.QueryCsv(path, true, readOptions).ToList();
    CompareRows(managedRows, rustRows, "csv-write-roundtrip");
    Require(rustRows.Count == 3, $"csv-write-roundtrip: expected 3 rows, received {rustRows.Count}.");

    using var stream = new MemoryStream();
    written = MiniExcelRust.SaveAsCsv(stream, initialRows, writeOptions, leaveOpen: true);
    Require(written == initialRows.Count && stream.CanWrite, "csv-write-stream: initial write failed.");
    written = MiniExcelRust.AppendCsv(stream, appendedRows, writeOptions, leaveOpen: true);
    Require(written == appendedRows.Count && stream.CanWrite, "csv-append-stream: append failed.");
    stream.Position = 0;
    var streamRows = MiniExcelRust.QueryCsv(stream, true, readOptions, leaveOpen: true).ToList();
    CompareRows(rustRows, streamRows, "csv-write-stream-roundtrip");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifyTypedConversions()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-typed-{Guid.NewGuid():N}.csv");
  var identifier = Guid.Parse("1ad46df8-08df-4ca6-8528-f79068fc23ea");
  try
  {
    File.WriteAllText(
      path,
      $"Identifier,State,Count,Optional\r\n{identifier},Ready,12,\r\n",
      new UTF8Encoding(true));
    var managedConfiguration = new CsvConfiguration { ReadEmptyStringAsNull = true };
    var rustConfiguration = new MiniExcelRustCsvReadOptions { ReadEmptyStringAsNull = true };
    var importer = ManagedMiniExcel.Importers.GetCsvImporter();
    var managed = importer.Query<TypedConversionRow>(path, configuration: managedConfiguration).Single();
    var rust = MiniExcelRust.QueryCsv<TypedConversionRow>(path, configuration: rustConfiguration).Single();
    Require(managed.Identifier == rust.Identifier && rust.Identifier == identifier, "typed-conversion: GUID differs.");
    Require(managed.State == rust.State && rust.State == RowState.Ready, "typed-conversion: enum differs.");
    Require(managed.Count == rust.Count && rust.Count == 12, "typed-conversion: integer differs.");
    Require(managed.Optional == rust.Optional && rust.Optional is null, "typed-conversion: nullable differs.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifyWorkbookMutations(string sourcePath)
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-mutate-{Guid.NewGuid():N}.xlsx");
  File.Copy(sourcePath, path);
  try
  {
    MiniExcelRust.RenameSheet(path, "Data", "Archive");
    MiniExcelRust.ReorderSheet(path, "Archive", 0);
    MiniExcelRust.SetSheetVisibility(path, "Sheet1", MiniExcelRustSheetState.VeryHidden);

    var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
    var managedNames = importer.GetSheetNames(path);
    var rustNames = MiniExcelRust.GetSheetNames(path);
    Require(managedNames.SequenceEqual(rustNames, StringComparer.Ordinal), "sheet-mutation: names differ.");
    Require(rustNames.SequenceEqual(new[] { "Archive", "Sheet1", "Options" }, StringComparer.Ordinal), "sheet-mutation: order differs.");

    var managedInfo = importer.GetSheetInformations(path);
    var rustInfo = MiniExcelRust.GetSheetInformations(path);
    Require(managedInfo.Count == rustInfo.Count, "sheet-mutation: info count differs.");
    for (var index = 0; index < managedInfo.Count; index++)
    {
      Require(managedInfo[index].Name == rustInfo[index].Name, $"sheet-mutation: name differs at {index}.");
      Require(managedInfo[index].State.ToString() == rustInfo[index].State.ToString(), $"sheet-mutation: state differs at {index}.");
    }
    Require(rustInfo.Single(sheet => sheet.Name == "Sheet1").State == MiniExcelRustSheetState.VeryHidden, "sheet-mutation: visibility was not updated.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifyInsertAndCopy()
{
  var insertPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-insert-{Guid.NewGuid():N}.xlsx");
  var destinationPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-copy-{Guid.NewGuid():N}.xlsx");
  var rows = new List<IDictionary<string, object?>>
  {
    new Dictionary<string, object?> { ["Key"] = "A", ["Amount"] = 10d },
    new Dictionary<string, object?> { ["Key"] = "B", ["Amount"] = 20d }
  };
  try
  {
    MiniExcelRust.SaveAs(
      insertPath,
      new[] { new Dictionary<string, object?> { ["Seed"] = "value" } },
      sheetName: "Base");
    var inserted = MiniExcelRust.InsertSheet(insertPath, rows, "Inserted");
    Require(inserted == rows.Count, "insert-sheet: row count differs.");
    CompareRows(rows, MiniExcelRust.Query(insertPath, true, "Inserted").ToList(), "insert-sheet");

    var rejectedDuplicate = false;
    try
    {
      MiniExcelRust.InsertSheet(insertPath, rows, "Inserted");
    }
    catch (InvalidOperationException)
    {
      rejectedDuplicate = true;
    }
    Require(rejectedDuplicate, "insert-sheet: duplicate sheet should be rejected by default.");

    var replaced = MiniExcelRust.InsertSheet(
      insertPath,
      rows,
      "Base",
      new MiniExcelRustInsertOptions { ReplaceExistingSheet = true });
    Require(replaced == rows.Count, "insert-sheet: replacement row count differs.");
    CompareRows(rows, QueryManaged(insertPath, true, "Base").ToList(), "replace-sheet");

    var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(insertPath)));
    var copied = MiniExcelRust.CopyAndAddSheet(insertPath, destinationPath, rows, "Copied");
    Require(copied == rows.Count, "copy-add-sheet: row count differs.");
    CompareRows(rows, QueryManaged(destinationPath, true, "Copied").ToList(), "copy-add-sheet");
    var sourceHashAfter = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(insertPath)));
    Require(sourceHash == sourceHashAfter, "copy-add-sheet: source workbook changed.");
  }
  finally
  {
    if (File.Exists(insertPath))
      File.Delete(insertPath);
    if (File.Exists(destinationPath))
      File.Delete(destinationPath);
  }
}

static void VerifyTemplateFill()
{
  var templatePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-template-{Guid.NewGuid():N}.xlsx");
  var managedPath = Path.Combine(Path.GetTempPath(), $"miniexcel-managed-template-{Guid.NewGuid():N}.xlsx");
  var rustPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-template-output-{Guid.NewGuid():N}.xlsx");
  var bytesPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-template-bytes-{Guid.NewGuid():N}.xlsx");
  var streamTemplatePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-template-stream-{Guid.NewGuid():N}.xlsx");
  var strictPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-template-strict-{Guid.NewGuid():N}.xlsx");
  try
  {
    MiniExcelRust.SaveAs(
      templatePath,
      new[]
      {
        new Dictionary<string, object?> { ["A"] = "{{title}}", ["B"] = "{{active}}" },
        new Dictionary<string, object?> { ["A"] = "{{items.name}}", ["B"] = "{{items.score}}" }
      },
      printHeader: false);
    var value = new
    {
      title = "Quarterly <Report>",
      active = true,
      items = new[] { new { name = "Ada", score = 10 }, new { name = "Linus", score = 20 } }
    };

    ManagedMiniExcel.Templaters.GetOpenXmlTemplater().FillTemplate(managedPath, templatePath, value);
    MiniExcelRust.FillTemplate(rustPath, templatePath, value);
    var expectedRows = QueryManaged(managedPath, false).ToList();
    CompareRows(expectedRows, MiniExcelRust.Query(rustPath).ToList(), "template-fill");

    var templateBytes = File.ReadAllBytes(templatePath);
    MiniExcelRust.FillTemplate(bytesPath, templateBytes, value);
    CompareRows(expectedRows, MiniExcelRust.Query(bytesPath).ToList(), "template-fill-bytes");

    using (var templateStream = new MemoryStream(templateBytes))
    {
      MiniExcelRust.FillTemplate(streamTemplatePath, templateStream, value, leaveTemplateOpen: true);
      Require(templateStream.CanRead, "template-stream: leaveTemplateOpen should preserve the stream.");
      CompareRows(expectedRows, MiniExcelRust.Query(streamTemplatePath).ToList(), "template-fill-template-stream");
    }

    using (var outputStream = new MemoryStream())
    {
      MiniExcelRust.FillTemplate(outputStream, templatePath, value, leaveOpen: true);
      Require(outputStream.CanWrite, "template-output-stream: leaveOpen should preserve the stream.");
      outputStream.Position = 0;
      CompareRows(expectedRows, MiniExcelRust.Query(outputStream).ToList(), "template-fill-output-stream");
    }

    using (var outputStream = new MemoryStream())
    using (var templateStream = new MemoryStream(templateBytes))
    {
      MiniExcelRust.FillTemplate(
        outputStream,
        templateStream,
        value,
        leaveOpen: true,
        leaveTemplateOpen: true);
      Require(outputStream.CanWrite && templateStream.CanRead, "template-streams: leave-open contract failed.");
      outputStream.Position = 0;
      CompareRows(expectedRows, MiniExcelRust.Query(outputStream).ToList(), "template-fill-streams");
    }

    using (var outputStream = new MemoryStream())
    {
      MiniExcelRust.FillTemplate(outputStream, templateBytes, value, leaveOpen: true);
      outputStream.Position = 0;
      CompareRows(expectedRows, MiniExcelRust.Query(outputStream).ToList(), "template-fill-byte-stream");
    }

    var rejectedOverwrite = false;
    try
    {
      MiniExcelRust.FillTemplate(rustPath, templatePath, value);
    }
    catch (InvalidOperationException)
    {
      rejectedOverwrite = true;
    }
    Require(rejectedOverwrite, "template-fill: overwrite=false should reject an existing file.");

    var rejectedMissing = false;
    try
    {
      MiniExcelRust.FillTemplate(strictPath, templatePath, new { title = "Missing items" }, ignoreMissingVariables: false);
    }
    catch (InvalidOperationException)
    {
      rejectedMissing = true;
    }
    Require(rejectedMissing, "template-fill: strict missing variables should fail.");
  }
  finally
  {
    foreach (var path in new[] { templatePath, managedPath, rustPath, bytesPath, streamTemplatePath, strictPath })
    {
      if (File.Exists(path))
        File.Delete(path);
    }
  }
}

static void VerifyCsvParity(string path)
{
  var managedConfiguration = new CsvConfiguration
  {
    Seperator = ';'
  };
  var rustConfiguration = new MiniExcelRustCsvReadOptions
  {
    Delimiter = ';'
  };
  var importer = ManagedMiniExcel.Importers.GetCsvImporter();
  var managedRows = importer.Query(path, true, managedConfiguration)
    .Cast<IDictionary<string, object?>>()
    .ToList();
  var rustRows = MiniExcelRust.QueryCsv(path, true, rustConfiguration).ToList();
  CompareRows(managedRows, rustRows, "csv");
  Require(rustRows.Count == 2, $"csv: expected 2 rows, received {rustRows.Count}.");
  Require(Equals(rustRows[1]["Note"], string.Empty), "csv: empty field should remain an empty string.");

  var managedTypedRows = importer.Query<TypedCsvRow>(path, configuration: managedConfiguration).ToList();
  var rustTypedRows = MiniExcelRust.QueryCsv<TypedCsvRow>(path, configuration: rustConfiguration).ToList();
  Require(managedTypedRows.Count == rustTypedRows.Count, "csv-typed: row count differs.");
  for (var index = 0; index < managedTypedRows.Count; index++)
  {
    Require(managedTypedRows[index].Name == rustTypedRows[index].Name, $"csv-typed: name differs at {index}.");
    Require(managedTypedRows[index].Note == rustTypedRows[index].Note, $"csv-typed: note differs at {index}.");
  }
  var asyncRows = CollectAsync(MiniExcelRust.QueryCsvAsync(path, true, rustConfiguration)).GetAwaiter().GetResult();
  CompareRows(managedRows, asyncRows, "csv-async");

  var managedColumns = importer.GetColumnNames(path, true, managedConfiguration);
  var rustColumns = MiniExcelRust.GetCsvColumnNames(path, true, rustConfiguration);
  Require(managedColumns.SequenceEqual(rustColumns, StringComparer.Ordinal), "csv-columns: values differ.");
  var asyncRustColumns = MiniExcelRust.GetCsvColumnNamesAsync(path, true, rustConfiguration).GetAwaiter().GetResult();
  Require(managedColumns.SequenceEqual(asyncRustColumns, StringComparer.Ordinal), "csv-columns-async: values differ.");

  var managedTable = importer.QueryAsDataTable(path, true, managedConfiguration);
  var rustTable = MiniExcelRust.QueryCsvAsDataTable(path, true, rustConfiguration);
  CompareDataTables(managedTable, rustTable, "csv-data-table");

  using var reader = MiniExcelRust.GetCsvReader(path, true, rustConfiguration);
  var readerRows = 0;
  while (reader.Read())
    readerRows++;
  Require(readerRows == 2, $"csv-data-reader: expected 2 rows, received {readerRows}.");

  using var stream = new MemoryStream(File.ReadAllBytes(path));
  var streamRows = MiniExcelRust.QueryCsv(stream, true, rustConfiguration, leaveOpen: true).ToList();
  CompareRows(managedRows, streamRows, "csv-stream");
  Require(stream.CanRead, "csv-stream: leaveOpen should preserve the stream.");

  stream.Position = 0;
  var streamColumns = MiniExcelRust.GetCsvColumnNames(stream, true, rustConfiguration, leaveOpen: true);
  Require(managedColumns.SequenceEqual(streamColumns, StringComparer.Ordinal), "csv-columns-stream: values differ.");
  Require(stream.CanRead, "csv-columns-stream: leaveOpen should preserve the stream.");
}

static void VerifyParity(string path)
{
  var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
  var managedSheetNames = importer.GetSheetNames(path);
  var rustSheetNames = MiniExcelRust.GetSheetNames(path);
  Require(
    managedSheetNames.SequenceEqual(rustSheetNames, StringComparer.Ordinal),
    $"sheet-names: managed={string.Join(",", managedSheetNames)}, rust={string.Join(",", rustSheetNames)}.");

  var asyncRustSheetNames = MiniExcelRust.GetSheetNamesAsync(path).GetAwaiter().GetResult();
  Require(
    managedSheetNames.SequenceEqual(asyncRustSheetNames, StringComparer.Ordinal),
    "sheet-names-async: the Rust result differs from the managed baseline.");

  VerifyColumnNames(importer, path, true, "Sheet1", "A1");
  VerifyColumnNames(importer, path, false, "Data", "C2");
  var asyncRustColumns = MiniExcelRust.GetColumnNamesAsync(path, true, "Sheet1").GetAwaiter().GetResult();
  Require(
    asyncRustColumns.SequenceEqual(new[] { "Name", "Value", "Note" }, StringComparer.Ordinal),
    "column-names-async: the Rust result did not match the expected headers.");

  var managedDimensions = importer.GetSheetDimensions(path);
  var rustDimensions = MiniExcelRust.GetSheetDimensions(path);
  Require(managedDimensions.Count == rustDimensions.Count, "sheet-dimensions: sheet count differs.");
  for (var index = 0; index < managedDimensions.Count; index++)
  {
    Require(
      managedDimensions[index].StartCell == rustDimensions[index].StartCell &&
      managedDimensions[index].EndCell == rustDimensions[index].EndCell,
      $"sheet-dimensions: range differs at sheet index {index}: " +
      $"managed={managedDimensions[index].StartCell}:{managedDimensions[index].EndCell}, " +
      $"rust={rustDimensions[index].StartCell}:{rustDimensions[index].EndCell}.");
  }
  var asyncRustDimensions = MiniExcelRust.GetSheetDimensionsAsync(path).GetAwaiter().GetResult();
  Require(asyncRustDimensions.Count == managedDimensions.Count, "sheet-dimensions-async: sheet count differs.");

  var managedSheetInfo = importer.GetSheetInformations(path);
  var rustSheetInfo = MiniExcelRust.GetSheetInformations(path);
  Require(managedSheetInfo.Count == rustSheetInfo.Count, "sheet-info: sheet count differs.");
  for (var index = 0; index < managedSheetInfo.Count; index++)
  {
    Require(managedSheetInfo[index].Id == rustSheetInfo[index].Id, $"sheet-info: id differs at index {index}.");
    Require(managedSheetInfo[index].Index == rustSheetInfo[index].Index, $"sheet-info: index differs at index {index}.");
    Require(managedSheetInfo[index].Name == rustSheetInfo[index].Name, $"sheet-info: name differs at index {index}.");
    Require(managedSheetInfo[index].State.ToString() == rustSheetInfo[index].State.ToString(), $"sheet-info: state differs at index {index}.");
    Require(managedSheetInfo[index].Active == rustSheetInfo[index].Active, $"sheet-info: active state differs at index {index}.");
    Require(rustSheetInfo[index].SheetType == MiniExcelRustSheetType.Worksheet, $"sheet-info: unexpected type at index {index}.");
  }
  var asyncRustSheetInfo = MiniExcelRust.GetSheetInformationsAsync(path).GetAwaiter().GetResult();
  Require(asyncRustSheetInfo.Count == managedSheetInfo.Count, "sheet-info-async: sheet count differs.");

  var managedRangeRows = QueryManagedRange(path, true, "Data", "C2", "D3").ToList();
  var rustRangeRows = MiniExcelRust.QueryRange(path, true, "Data", "C2", "D3").ToList();
  CompareRows(managedRangeRows, rustRangeRows, "bounded-range");
  Require(rustRangeRows.Count == 1, $"bounded-range: expected 1 row, received {rustRangeRows.Count}.");

  var managedTableRows = QueryManagedTable(path, "Data", "DataTable").ToList();
  var rustTableRows = MiniExcelRust.QueryTable(path, "Data", "datatable").ToList();
  CompareRows(managedTableRows, rustTableRows, "named-table");
  Require(rustTableRows.Count == 2, $"named-table: expected 2 rows, received {rustTableRows.Count}.");

  var managedConfiguration = new OpenXmlConfiguration
  {
    IgnoreEmptyRows = true,
    TrimColumnNames = true
  };
  var rustConfiguration = new MiniExcelRustReadOptions
  {
    IgnoreEmptyRows = true,
    TrimColumnNames = true
  };
  var managedConfiguredRows = QueryManaged(
    path,
    true,
    "Options",
    "A1",
    managedConfiguration).ToList();
  var rustConfiguredRows = MiniExcelRust.Query(
    path,
    true,
    "Options",
    "A1",
    rustConfiguration).ToList();
  CompareRows(managedConfiguredRows, rustConfiguredRows, "configured-query");
  Require(rustConfiguredRows.Count == 2, $"configured-query: expected 2 rows, received {rustConfiguredRows.Count}.");

  var scenarios = new[]
  {
    new QueryScenario("header", true, "Sheet1", "A1"),
    new QueryScenario("headerless", false, "Sheet1", "A1"),
    new QueryScenario("sheet-and-start-cell", true, "Data", "C2")
  };

  foreach (var scenario in scenarios)
  {
    var managedRows = QueryManaged(path, scenario.UseHeaderRow, scenario.SheetName, scenario.StartCell).ToList();
    var rustRows = MiniExcelRust.Query(path, scenario.UseHeaderRow, scenario.SheetName, scenario.StartCell).ToList();
    CompareRows(managedRows, rustRows, scenario.Name);
  }

  var rows = MiniExcelRust.Query(path, useHeaderRow: true, sheetName: "Sheet1").ToList();
  Require(rows.Count == 3, $"Expected 3 rows, received {rows.Count}.");
  Require(Equals(rows[0]["Name"], "alpha"), "The first string value did not match.");
  Require(Equals(rows[0]["Value"], 42d), "The first numeric value did not match.");
  Require(Equals(rows[0]["Note"], "Taiwan 台灣"), "The Unicode value did not match.");
  Require(Equals(rows[1]["Name"], "beta"), "The second string value did not match.");
  Require(Equals(rows[1]["Value"], true), "The boolean value did not match.");
  Require(rows[1]["Note"] is null, "The empty value should be null.");

  var managedTypedRows = importer.Query<TypedSheetRow>(path, "Sheet1").ToList();
  var rustTypedRows = MiniExcelRust.Query<TypedSheetRow>(path, "Sheet1").ToList();
  CompareTypedRows(managedTypedRows, rustTypedRows, "typed-query");
  var managedAliases = importer.Query<TypedAliasRow>(path, "Sheet1").ToList();
  var rustAliases = MiniExcelRust.Query<TypedAliasRow>(path, "Sheet1").ToList();
  Require(
    managedAliases.Select(row => row.Label).SequenceEqual(rustAliases.Select(row => row.Label)),
    "typed-alias: values differ.");

  var typedTableRows = MiniExcelRust.QueryTable<TypedTableRow>(path, "Data", "DataTable").ToList();
  Require(
    typedTableRows.Count == 2 && typedTableRows[0].Code == "x" && typedTableRows[0].Amount == 3.5d,
    "typed-table: values differ.");

  var asyncRows = CollectAsync(MiniExcelRust.QueryAsync(path, true, "Sheet1")).GetAwaiter().GetResult();
  CompareRows(rows, asyncRows, "async-query");
  var asyncTypedRows = CollectAsync(MiniExcelRust.QueryAsync<TypedSheetRow>(path, "Sheet1")).GetAwaiter().GetResult();
  CompareTypedRows(managedTypedRows, asyncTypedRows, "async-typed-query");
  var asyncTableRows = CollectAsync(MiniExcelRust.QueryTableAsync(path, "Data", "DataTable")).GetAwaiter().GetResult();
  CompareRows(managedTableRows, asyncTableRows, "async-table-query");

  using var cancellation = new CancellationTokenSource();
  cancellation.Cancel();
  var cancelled = false;
  try
  {
    _ = CollectAsync(MiniExcelRust.QueryAsync(path, cancellationToken: cancellation.Token)).GetAwaiter().GetResult();
  }
  catch (OperationCanceledException)
  {
    cancelled = true;
  }
  Require(cancelled, "async-query: a pre-cancelled token should cancel enumeration.");

  var managedTable = importer.QueryAsDataTable(path, true, "Sheet1");
  var rustTable = MiniExcelRust.QueryAsDataTable(path, true, "Sheet1");
  CompareDataTables(managedTable, rustTable, "data-table");

  using var reader = MiniExcelRust.GetReader(path, true, "Sheet1");
  Require(reader.FieldCount == 3, $"data-reader: expected 3 fields, received {reader.FieldCount}.");
  var readerRows = 0;
  while (reader.Read())
    readerRows++;
  Require(readerRows == 3, $"data-reader: expected 3 rows, received {readerRows}.");

  VerifyStreamParity(path);
}

static void CompareDataTables(DataTable expected, DataTable actual, string scenario)
{
  Require(expected.Columns.Count == actual.Columns.Count, $"{scenario}: column count differs.");
  Require(expected.Rows.Count == actual.Rows.Count, $"{scenario}: row count differs.");
  for (var column = 0; column < expected.Columns.Count; column++)
    Require(expected.Columns[column].ColumnName == actual.Columns[column].ColumnName, $"{scenario}: column name differs at {column}.");
  for (var row = 0; row < expected.Rows.Count; row++)
  {
    for (var column = 0; column < expected.Columns.Count; column++)
      Require(Equals(expected.Rows[row][column], actual.Rows[row][column]), $"{scenario}: value differs at {row},{column}.");
  }
}

static void CompareTypedRows(
  IReadOnlyList<TypedSheetRow> expected,
  IReadOnlyList<TypedSheetRow> actual,
  string scenario)
{
  Require(expected.Count == actual.Count, $"{scenario}: row count differs.");
  for (var index = 0; index < expected.Count; index++)
  {
    Require(expected[index].Name == actual[index].Name, $"{scenario}: name differs at {index}.");
    Require(Equals(expected[index].Value, actual[index].Value), $"{scenario}: value differs at {index}.");
    Require(expected[index].Note == actual[index].Note, $"{scenario}: note differs at {index}.");
  }
}

static async Task<List<T>> CollectAsync<T>(IAsyncEnumerable<T> values)
{
  var result = new List<T>();
  await foreach (var value in values)
    result.Add(value);
  return result;
}

static void VerifyStreamParity(string path)
{
  var bytes = File.ReadAllBytes(path);
  using (var stream = new MemoryStream(bytes))
  {
    var names = MiniExcelRust.GetSheetNames(stream, leaveOpen: true);
    Require(names.SequenceEqual(new[] { "Sheet1", "Data", "Options" }, StringComparer.Ordinal), "stream-sheet-names: values differ.");
    Require(stream.CanRead, "stream-sheet-names: leaveOpen should preserve the stream.");
  }

  using (var stream = new MemoryStream(bytes))
  {
    var dimensions = MiniExcelRust.GetSheetDimensions(stream, leaveOpen: true);
    Require(dimensions.Count == 3, "stream-sheet-dimensions: expected one range per sheet.");
    Require(stream.CanRead, "stream-sheet-dimensions: leaveOpen should preserve the stream.");
  }

  using (var stream = new MemoryStream(bytes))
  {
    var sheetInfo = MiniExcelRust.GetSheetInformations(stream, leaveOpen: true);
    Require(sheetInfo.Count == 3, "stream-sheet-info: expected one record per sheet.");
    Require(stream.CanRead, "stream-sheet-info: leaveOpen should preserve the stream.");
  }

  using (var stream = new MemoryStream(bytes))
  {
    var expected = QueryManaged(path, true, "Data", "C2").ToList();
    var actual = MiniExcelRust.Query(stream, true, "Data", "C2", leaveOpen: true).ToList();
    CompareRows(expected, actual, "stream-query");
    Require(stream.CanRead, "stream-query: leaveOpen should preserve the stream.");
  }

  using (var stream = new MemoryStream(bytes))
  {
    var tableRows = MiniExcelRust.QueryTable(stream, "Data", "DataTable", leaveOpen: true).ToList();
    Require(tableRows.Count == 2, "stream-query-table: expected 2 rows.");
    Require(stream.CanRead, "stream-query-table: leaveOpen should preserve the stream.");
  }

  using (var stream = new MemoryStream(bytes))
  {
    using var reader = MiniExcelRust.GetReader(stream, true, "Sheet1", leaveOpen: true);
    Require(reader.Read(), "stream-data-reader: expected a row.");
    Require(stream.CanRead, "stream-data-reader: leaveOpen should preserve the stream.");
  }

  var closingStream = new MemoryStream(bytes);
  _ = MiniExcelRust.GetColumnNames(closingStream, true, "Sheet1");
  Require(!closingStream.CanRead, "stream-column-names: the default should close the stream.");

  var earlyDisposalStream = new MemoryStream(bytes);
  var enumerator = MiniExcelRust.Query(earlyDisposalStream).GetEnumerator();
  try
  {
    Require(enumerator.MoveNext(), "stream-query: expected a row before early disposal.");
  }
  finally
  {
    enumerator.Dispose();
  }
  Require(!earlyDisposalStream.CanRead, "stream-query: early disposal should close the stream.");
}

static void VerifyColumnNames(
  OpenXmlImporter importer,
  string path,
  bool useHeaderRow,
  string sheetName,
  string startCell)
{
  var managedColumns = importer.GetColumnNames(path, useHeaderRow, sheetName, startCell);
  var rustColumns = MiniExcelRust.GetColumnNames(path, useHeaderRow, sheetName, startCell);
  Require(
    managedColumns.SequenceEqual(rustColumns, StringComparer.Ordinal),
    $"column-names: managed={string.Join(",", managedColumns)}, rust={string.Join(",", rustColumns)}.");
}

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

static IEnumerable<IDictionary<string, object?>> QueryManaged(
  string path,
  bool useHeaderRow,
  string? sheetName = null,
  string startCell = "A1",
  OpenXmlConfiguration? configuration = null)
{
  var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
  foreach (IDictionary<string, object?> row in importer.Query(
    path,
    hasHeaderRow: useHeaderRow,
    sheetName: sheetName,
    startCell: startCell,
    configuration: configuration))
    yield return row;
}

static IEnumerable<IDictionary<string, object?>> QueryManagedRange(
  string path,
  bool useHeaderRow,
  string? sheetName,
  string startCell,
  string? endCell)
{
  var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
  foreach (IDictionary<string, object?> row in importer.QueryRange(
    path,
    hasHeaderRow: useHeaderRow,
    sheetName: sheetName,
    startCell: startCell,
    endCell: endCell))
    yield return row;
}

static IEnumerable<IDictionary<string, object?>> QueryManagedTable(
  string path,
  string? sheetName,
  string tableName)
{
  var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
  foreach (IDictionary<string, object?> row in importer.QueryTable(path, sheetName, tableName))
    yield return row;
}

static void CompareRows(
  IReadOnlyList<IDictionary<string, object?>> expected,
  IReadOnlyList<IDictionary<string, object?>> actual,
  string scenario)
{
  Require(expected.Count == actual.Count, $"{scenario}: row count differs: managed={expected.Count}, rust={actual.Count}.");
  for (var rowIndex = 0; rowIndex < expected.Count; rowIndex++)
    CompareRow(expected[rowIndex], actual[rowIndex], scenario, rowIndex);
}

static void CompareRow(
  IDictionary<string, object?> expected,
  IDictionary<string, object?> actual,
  string scenario,
  long rowIndex)
{
  Require(
    expected.Keys.SequenceEqual(actual.Keys, StringComparer.Ordinal),
    $"{scenario}: column order differs at row {rowIndex}.");
  foreach (var key in expected.Keys)
  {
    Require(
      Equals(expected[key], actual[key]),
      $"{scenario}: value differs at row {rowIndex}, column {key}: managed={expected[key] ?? "<null>"}, rust={actual[key] ?? "<null>"}.");
  }
}

static int VerifyFileParity(string[] arguments)
{
  if (arguments.Length is < 2 or > 3)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var useHeaderRow = arguments.Length == 3 && bool.Parse(arguments[2]);
  using var managed = QueryManaged(path, useHeaderRow).GetEnumerator();
  using var rust = MiniExcelRust.Query(path, useHeaderRow).GetEnumerator();
  long rowIndex = 0;

  while (true)
  {
    var hasManaged = managed.MoveNext();
    var hasRust = rust.MoveNext();
    Require(hasManaged == hasRust, $"benchmark-parity: row count differs after row {rowIndex}.");
    if (!hasManaged)
      break;
    CompareRow(managed.Current, rust.Current, "benchmark-parity", rowIndex);
    rowIndex++;
  }

  Console.WriteLine($"Verified {rowIndex} rows against the configured MiniExcel baseline.");
  return 0;
}

static void VerifyLifecycle(string path, int iterations, int maxPrivateGrowthMb)
{
  if (iterations < 100)
    throw new ArgumentOutOfRangeException(nameof(iterations), "Lifecycle iterations must be at least 100.");

  for (var iteration = 0; iteration < 50; iteration++)
    QueryAndDisposeEarly(path);

  ForceCollection();
  var baseline = CaptureResources();
  var samples = new List<ResourceSample> { baseline };
  var sampleInterval = Math.Max(100, iterations / 10);

  for (var iteration = 1; iteration <= iterations; iteration++)
  {
    QueryAndDisposeEarly(path);
    if (iteration % 25 == 0)
      _ = MiniExcelRust.Query(path, useHeaderRow: true).Count();
    if (iteration % sampleInterval == 0)
    {
      ForceCollection();
      samples.Add(CaptureResources());
    }
  }

  using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }

  var final = samples[^1];
  var privateGrowth = final.PrivateBytes - baseline.PrivateBytes;
  var workingSetGrowth = final.WorkingSetBytes - baseline.WorkingSetBytes;
  var maxPrivateGrowth = maxPrivateGrowthMb * 1024L * 1024L;
  Require(
    privateGrowth <= maxPrivateGrowth,
    $"Private memory grew by {ToMb(privateGrowth):F2} MB after {iterations} early-disposal cycles; limit is {maxPrivateGrowthMb} MB.");

  if (baseline.ResourceCount is not null && final.ResourceCount is not null)
  {
    Require(
      final.ResourceCount - baseline.ResourceCount <= 4,
      $"Native handle/file-descriptor count grew from {baseline.ResourceCount} to {final.ResourceCount}.");
  }

  Console.WriteLine(
    $"Lifecycle: iterations={iterations}, private-growth={ToMb(privateGrowth):F2} MB, " +
    $"working-set-growth={ToMb(workingSetGrowth):F2} MB, resources={baseline.ResourceCount?.ToString() ?? "n/a"}->{final.ResourceCount?.ToString() ?? "n/a"}.");
}

static void QueryAndDisposeEarly(string path)
{
  using var rows = MiniExcelRust.Query(path).GetEnumerator();
  Require(rows.MoveNext(), "The lifecycle fixture did not contain a row.");
}

static ResourceSample CaptureResources()
{
  using var process = Process.GetCurrentProcess();
  process.Refresh();
  return new ResourceSample(process.PrivateMemorySize64, process.WorkingSet64, GetResourceCount(process));
}

static int? GetResourceCount(Process process)
{
  if (OperatingSystem.IsLinux() && Directory.Exists("/proc/self/fd"))
    return Directory.EnumerateFileSystemEntries("/proc/self/fd").Count();

  try
  {
    return process.HandleCount;
  }
  catch (PlatformNotSupportedException)
  {
    return null;
  }
}

static void ForceCollection()
{
  GC.Collect();
  GC.WaitForPendingFinalizers();
  GC.Collect();
}

static double ToMb(long bytes) => bytes / 1024d / 1024d;

static int GenerateBenchmarkWorkbook(string[] arguments)
{
  if (arguments.Length is < 2 or > 4)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var rows = arguments.Length >= 3 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 100_000;
  var columns = arguments.Length >= 4 ? int.Parse(arguments[3], CultureInfo.InvariantCulture) : 10;
  CreateBenchmarkWorkbook(path, rows, columns);
  Console.WriteLine($"Generated {rows}x{columns} benchmark workbook at {path}.");
  return 0;
}

static int Benchmark(string[] arguments, bool useRust)
{
  if (arguments.Length is < 2 or > 4)
    return Usage();

  var path = Path.GetFullPath(arguments[1]);
  var passes = arguments.Length >= 3 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 1;
  var warmupPasses = arguments.Length >= 4 ? int.Parse(arguments[3], CultureInfo.InvariantCulture) : 0;

  for (var pass = 0; pass < warmupPasses; pass++)
    Consume(path, useRust);

  ForceCollection();
  var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
  var stopwatch = Stopwatch.StartNew();
  var firstRowMilliseconds = 0d;
  long rowCount = 0;
  long cellCount = 0;

  for (var pass = 0; pass < passes; pass++)
  {
    var rows = useRust ? MiniExcelRust.Query(path) : QueryManaged(path, useHeaderRow: false);
    foreach (var row in rows)
    {
      if (rowCount == 0)
        firstRowMilliseconds = stopwatch.Elapsed.TotalMilliseconds;
      rowCount++;
      cellCount += row.Count;
    }
  }

  stopwatch.Stop();
  Console.WriteLine(JsonSerializer.Serialize(new BenchmarkResult(
    useRust ? "MiniExcelRust" : "MiniExcelV2",
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    passes,
    rowCount,
    cellCount,
    stopwatch.Elapsed.TotalMilliseconds,
    firstRowMilliseconds,
    GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore)));
  return 0;
}

static void Consume(string path, bool useRust)
{
  var rows = useRust ? MiniExcelRust.Query(path) : QueryManaged(path, useHeaderRow: false);
  foreach (var row in rows)
    _ = row.Count;
}

static void CreateWorkbook(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddEntry(archive, "[Content_Types].xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
        </Types>
        """);
    AddEntry(archive, "_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/workbook.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <bookViews><workbookView activeTab="2"/></bookViews>
          <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" state="hidden" r:id="rId2"/><sheet name="Options" sheetId="3" r:id="rId3"/></sheets>
        </workbook>
        """);
    AddEntry(archive, "xl/_rels/workbook.xml.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/worksheets/sheet1.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:C4"/>
          <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c><c r="C1" t="inlineStr"><is><t>Note</t></is></c></row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>42</v></c><c r="C2" t="inlineStr"><is><t>Taiwan 台灣</t></is></c></row>
            <row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3" t="b"><v>1</v></c><c r="C3"/></row>
            <row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c><c r="B4"><v>-7.5</v></c><c r="C4" t="inlineStr"><is><t>last</t></is></c></row>
          </sheetData>
        </worksheet>
        """);
    AddEntry(archive, "xl/worksheets/sheet2.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <dimension ref="C2:D4"/>
          <sheetData>
            <row r="2"><c r="C2" t="inlineStr"><is><t>Code</t></is></c><c r="D2" t="inlineStr"><is><t>Amount</t></is></c></row>
            <row r="3"><c r="C3" t="inlineStr"><is><t>x</t></is></c><c r="D3"><v>3.5</v></c></row>
            <row r="4"><c r="C4" t="inlineStr"><is><t>y</t></is></c><c r="D4"><v>9</v></c></row>
          </sheetData>
          <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
        </worksheet>
        """);
    AddEntry(archive, "xl/worksheets/_rels/sheet2.xml.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/tables/table1.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="DataTable" displayName="DataTable" ref="C2:D4" totalsRowShown="0">
          <autoFilter ref="C2:D4"/>
          <tableColumns count="2"><tableColumn id="1" name="Code"/><tableColumn id="2" name="Amount"/></tableColumns>
          <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
        </table>
        """);
    AddEntry(archive, "xl/worksheets/sheet3.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:B4"/>
          <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> Left </t></is></c><c r="B1" t="inlineStr"><is><t>Right</t></is></c></row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>merged</t></is></c></row>
            <row r="4"><c r="A4" t="inlineStr"><is><t>x</t></is></c><c r="B4" t="inlineStr"><is><t>y</t></is></c></row>
          </sheetData>
          <mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>
        </worksheet>
        """);
}

static void CreateBenchmarkWorkbook(string path, int rows, int columns)
{
    if (rows < 1 || columns is < 1 or > 26)
        throw new ArgumentOutOfRangeException(nameof(rows), "Rows must be positive and columns must be between 1 and 26.");

    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    if (File.Exists(path))
      File.Delete(path);
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddEntry(archive, "[Content_Types].xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """);
    AddEntry(archive, "_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """);
    AddEntry(archive, "xl/workbook.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """);
    AddEntry(archive, "xl/_rels/workbook.xml.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        """);

    var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
    for (var row = 1; row <= rows; row++)
    {
        writer.Write($"<row r=\"{row}\">");
        for (var column = 1; column <= columns; column++)
        {
            var reference = $"{(char)('A' + column - 1)}{row}";
            var value = (long)(row - 1) * columns + column;
            writer.Write($"<c r=\"{reference}\"><v>{value}</v></c>");
        }
        writer.Write("</row>");
    }
    writer.Write("</sheetData></worksheet>");
}

static void AddEntry(ZipArchive archive, string name, string contents)
{
    var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
    writer.Write(contents);
}

  static int Usage()
  {
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  PublicNuGetSmoke suite [lifecycle-iterations] [max-private-growth-mb]");
    Console.Error.WriteLine("  PublicNuGetSmoke comments <xlsx-path> [sheet-name]");
    Console.Error.WriteLine("  PublicNuGetSmoke merge <xlsx-path>");
    Console.Error.WriteLine("  PublicNuGetSmoke verify <xlsx-path> [use-header-row]");
    Console.Error.WriteLine("  PublicNuGetSmoke generate <xlsx-path> [rows] [columns]");
    Console.Error.WriteLine("  PublicNuGetSmoke <managed|rust> <xlsx-path> [passes] [warmup-passes]");
    return 2;
  }

  internal sealed record QueryScenario(string Name, bool UseHeaderRow, string SheetName, string StartCell);

  internal sealed record ResourceSample(long PrivateBytes, long WorkingSetBytes, int? ResourceCount);

  internal sealed record BenchmarkResult(
    string Runtime,
    string DotNetRuntime,
    int Passes,
    long Rows,
    long Cells,
    double ElapsedMilliseconds,
    double FirstRowMilliseconds,
    long AllocatedBytes);

  internal sealed class TypedSheetRow
  {
    public string? Name { get; set; }
    public object? Value { get; set; }
    public string? Note { get; set; }
  }

  internal sealed class TypedAliasRow
  {
    [ExcelColumnName("Name")]
    public string? Label { get; set; }
  }

  internal sealed class TypedTableRow
  {
    public string? Code { get; set; }
    public double Amount { get; set; }
  }

  internal sealed class TypedCsvRow
  {
    public string? Name { get; set; }
    public string? Note { get; set; }
  }

  internal sealed class TypedConversionRow
  {
    public Guid Identifier { get; set; }
    public RowState State { get; set; }
    public int Count { get; set; }
    public int? Optional { get; set; }
  }

  internal enum RowState
  {
    Unknown,
    Ready
  }