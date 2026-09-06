using System.Diagnostics;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Resources;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using MiniExcelLib;
using MiniExcelLib.Csv;
using MiniExcelLib.OpenXml;
using MiniExcelLib.OpenXml.FluentMapping;
using MiniExcelLib.OpenXml.FluentMapping.Api;
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
    VerifyMissingDimensionParity(workbookPath);
    VerifySelfClosingEmptyRowParity(workbookPath);
    VerifyCompatibilityFacade(workbookPath);
    VerifyCsvParity(csvPath);
    VerifyConversions();
    VerifySaveAs();
    VerifyMultiSheetSaveAs();
    VerifyConfiguredWrite();
    VerifyCsvWrite();
    VerifyTypedConversions();
    VerifyTypedExports();
    VerifyFluentMapping();
    VerifyInsertAndCopy();
    VerifyTemplateFill();
    VerifyPictures();
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

static void VerifyFluentMapping()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-{Guid.NewGuid():N}.xlsx");
  var formulaPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-formula-{Guid.NewGuid():N}.xlsx");
  var formulaTemplatePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-formula-template-{Guid.NewGuid():N}.xlsx");
  var managedPath = Path.Combine(Path.GetTempPath(), $"miniexcel-managed-fluent-{Guid.NewGuid():N}.xlsx");
  var oraclePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-oracle-{Guid.NewGuid():N}.xlsx");
  var templatePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-template-{Guid.NewGuid():N}.xlsx");
  var mappedTemplatePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-template-output-{Guid.NewGuid():N}.xlsx");
  var invalidPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-fluent-invalid-{Guid.NewGuid():N}.xlsx");
  try
  {
    var mapping = new MiniExcelRustMapping<MappedDepartment>().ToWorksheet("Mapped");
    mapping.Property(value => value.Name).ToCell("A1");
    mapping.Collection(value => value.PhoneNumbers).StartAt("A3").WithSpacing(1);
    mapping.Collection(value => value.Projects)
      .StartAt("C3")
      .WithItemMapping<MappedProject>(project =>
      {
        project.Property(value => value.Code).ToCell("C3");
        project.Collection(value => value.Tasks)
          .StartAt("D3")
          .WithItemMapping<MappedTask>(task =>
          {
            task.Property(value => value.Name).ToCell("D3");
            task.Property(value => value.Hours).ToCell("E3");
          });
      });

    var source = new[]
    {
      new MappedDepartment
      {
        Name = "Engineering",
        PhoneNumbers = ["555-1000", "555-2000"],
        Projects =
        [
          new MappedProject
          {
            Code = "P1",
            Tasks = [new MappedTask { Name = "Design", Hours = 2 }, new MappedTask { Name = "Build", Hours = 5 }]
          },
          new MappedProject
          {
            Code = "P2",
            Tasks = [new MappedTask { Name = "Test", Hours = 3 }]
          }
        ]
      }
    };
    var written = MiniExcelRustMappingExtensions.ExportMapped(path, source, mapping);
    Require(written == 5, $"fluent-mapping: expected 5 physical rows, received {written}.");
    var rows = MiniExcelRust.Query(path, false, "Mapped").ToList();
    Require(rows.Count == 5, $"fluent-mapping: expected 5 query rows, received {rows.Count}.");
    Require(Equals(rows[0]["A"], "Engineering"), "fluent-mapping: scalar cell differs.");
    Require(Equals(rows[2]["A"], "555-1000") && Equals(rows[4]["A"], "555-2000"), "fluent-mapping: collection spacing differs.");
    Require(Equals(rows[2]["C"], "P1") && Equals(rows[2]["D"], "Design"), "fluent-mapping: first nested item differs.");
    Require(Equals(rows[3]["D"], "Build") && Convert.ToInt32(rows[3]["E"], CultureInfo.InvariantCulture) == 5, "fluent-mapping: nested collection expansion differs.");
    Require(Equals(rows[4]["C"], "P2") && Equals(rows[4]["D"], "Test"), "fluent-mapping: second parent item overlapped the first.");
    var managedRegistry = new MappingRegistry();
    managedRegistry.Configure<MappedDepartment>(configuration =>
    {
      configuration.ToWorksheet("Mapped");
      configuration.Property(value => value.Name).ToCell("A1");
      configuration.Collection(value => value.PhoneNumbers).StartAt("A3").WithSpacing(1);
      configuration.Collection(value => value.Projects)
        .StartAt("C3")
        .WithItemMapping<MappedProject>(project =>
        {
          project.Property(value => value.Code).ToCell("C3");
          project.Collection(value => value.Tasks)
            .StartAt("D3")
            .WithItemMapping<MappedTask>(task =>
            {
              task.Property(value => value.Name).ToCell("D3");
              task.Property(value => value.Hours).ToCell("E3");
            });
        });
    });
    var oracleSource = new[]
    {
      new MappedDepartment
      {
        Name = source[0].Name,
        PhoneNumbers = source[0].PhoneNumbers,
        Projects = [source[0].Projects[0]]
      }
    };
    ManagedMiniExcel.Exporters.GetMappingExporter(managedRegistry).Export(managedPath, oracleSource, overwriteFile: true);
    MiniExcelRustMappingExtensions.ExportMapped(oraclePath, oracleSource, mapping);
    CompareRows(
      QueryManaged(managedPath, useHeaderRow: false, sheetName: "Mapped").ToList(),
      MiniExcelRust.Query(oraclePath, false, "Mapped").ToList(),
      "fluent-mapping-source-oracle");
    var roundTrip = MiniExcelRustMappingExtensions.ReadMapped(path, mapping);
    Require(roundTrip.Name == "Engineering", "fluent-mapping-read: scalar property differs.");
    Require(roundTrip.PhoneNumbers.SequenceEqual(new[] { "555-1000", "555-2000" }), "fluent-mapping-read: simple collection differs.");
    Require(roundTrip.Projects.Count == 2, "fluent-mapping-read: complex collection count differs.");
    Require(roundTrip.Projects[0].Tasks.Count == 2 && roundTrip.Projects[0].Tasks[1].Name == "Build", "fluent-mapping-read: first nested collection differs.");
    Require(roundTrip.Projects[1].Code == "P2" && roundTrip.Projects[1].Tasks.Single().Hours == 3, "fluent-mapping-read: second nested collection differs.");
    using (var mappedStream = File.OpenRead(path))
    {
      var streamRoundTrip = MiniExcelRustMappingExtensions.ReadMapped(mappedStream, mapping, leaveOpen: true);
      Require(streamRoundTrip.Projects.Count == 2 && mappedStream.CanRead, "fluent-mapping-read-stream: values or ownership differ.");
    }
    var asyncRoundTrip = MiniExcelRustMappingExtensions.ReadMappedAsync(path, mapping).GetAwaiter().GetResult();
    Require(asyncRoundTrip.Projects[1].Code == "P2", "fluent-mapping-read-async: complex collection differs.");
    using (var exportStream = new MemoryStream())
    {
      var streamCount = MiniExcelRustMappingExtensions.ExportMappedAsync(exportStream, source, mapping, leaveOpen: true).GetAwaiter().GetResult();
      exportStream.Position = 0;
      var streamRows = MiniExcelRust.Query(exportStream, false, "Mapped", leaveOpen: true).ToList();
      Require(streamCount == 5 && Equals(streamRows[4]["C"], "P2") && exportStream.CanRead, "fluent-mapping-export-stream: values or ownership differ.");
    }

    var formulaMapping = new MiniExcelRustMapping<MappedFormula>();
    formulaMapping.Property(value => value.Name).ToCell("A1");
    formulaMapping.Property(value => value.Amount).ToCell("B1").WithFormat("#,##0.00");
    formulaMapping.Property(value => value.Total).ToCell("C1").WithFormula("=B1*2");
    formulaMapping.Property(value => value.Name).ToCell("C2");
    MiniExcelRustMappingExtensions.ExportMapped(
      formulaPath,
      new[] { new MappedFormula { Name = "Line", Amount = 12.5 } },
      formulaMapping);
    var formulaXml = ReadZipEntryText(formulaPath, "xl/worksheets/sheet1.xml");
    var stylesXml = ReadZipEntryText(formulaPath, "xl/styles.xml");
    Require(formulaXml.Contains("<f>B1*2</f>", StringComparison.Ordinal), "fluent-mapping: formula cell was not written.");
    Require(MiniExcelRust.Query(formulaPath, false).ElementAt(1)["C"]?.ToString() == "Line", "fluent-mapping: a regular cell in the formula column was changed.");
    Require(stylesXml.Contains("#,##0.00", StringComparison.Ordinal), "fluent-mapping: number format was not written.");
    MiniExcelRustMappingExtensions.FillMappedTemplate(
      formulaTemplatePath,
      formulaPath,
      new[] { new MappedFormula { Name = "Updated", Amount = 7.5 } },
      formulaMapping);
    var formulaTemplateXml = ReadZipEntryText(formulaTemplatePath, "xl/worksheets/sheet1.xml");
    Require(formulaTemplateXml.Contains("<f>B1*2</f>", StringComparison.Ordinal), "fluent-template: formula cell was not overlaid.");
    var formulaTemplateRows = MiniExcelRust.Query(formulaTemplatePath, false).ToList();
    Require(Equals(formulaTemplateRows[0]["A"], "Updated") && Convert.ToDouble(formulaTemplateRows[0]["B"], CultureInfo.InvariantCulture) == 7.5, "fluent-template: formula source cells differ.");

    var templateRows = Enumerable.Range(1, 5)
      .Select(index => (IDictionary<string, object?>)new Dictionary<string, object?>
      {
        ["A"] = index == 1 ? "Original" : null,
        ["F"] = $"Keep-{index}"
      })
      .ToList();
    var templateOptions = new MiniExcelRustWriteOptions { SheetName = "Mapped", PrintHeader = false };
    templateOptions.ColumnFormats["A"] = "@";
    MiniExcelRust.SaveAsWithSchema(templatePath, new[] { "A", "B", "C", "D", "E", "F" }, templateRows, templateOptions);
    var templateDocument = XDocument.Parse(ReadZipEntryText(templatePath, "xl/worksheets/sheet1.xml"));
    var templateStyle = templateDocument.Descendants().Single(element => element.Name.LocalName == "c" && (string?)element.Attribute("r") == "A1").Attribute("s")?.Value;
    MiniExcelRustMappingExtensions.FillMappedTemplate(mappedTemplatePath, templatePath, source, mapping);
    var mappedRows = MiniExcelRust.Query(mappedTemplatePath, false, "Mapped").ToList();
    Require(Equals(mappedRows[0]["A"], "Engineering") && Equals(mappedRows[0]["F"], "Keep-1"), "fluent-template: mapped or unrelated scalar cell differs.");
    Require(Equals(mappedRows[3]["D"], "Build") && Equals(mappedRows[4]["C"], "P2"), "fluent-template: nested collection rows differ.");
    Require(Equals(mappedRows[4]["F"], "Keep-5"), "fluent-template: unrelated template cells were not preserved.");
    var mappedDocument = XDocument.Parse(ReadZipEntryText(mappedTemplatePath, "xl/worksheets/sheet1.xml"));
    var mappedStyle = mappedDocument.Descendants().Single(element => element.Name.LocalName == "c" && (string?)element.Attribute("r") == "A1").Attribute("s")?.Value;
    Require(templateStyle is not null && mappedStyle == templateStyle, "fluent-template: existing target style was not preserved.");

    using (var outputStream = new MemoryStream())
    {
      MiniExcelRustMappingExtensions.FillMappedTemplateAsync(outputStream, File.ReadAllBytes(templatePath), source, mapping, leaveOpen: true).GetAwaiter().GetResult();
      outputStream.Position = 0;
      var streamRows = MiniExcelRust.Query(outputStream, false, "Mapped", leaveOpen: true).ToList();
      Require(Equals(streamRows[4]["C"], "P2") && outputStream.CanRead, "fluent-template-stream: values or ownership differ.");
    }

    var missingCell = new MiniExcelRustMapping<MappedFormula>();
    missingCell.Property(value => value.Name);
    var rejectedMissingCell = false;
    try
    {
      MiniExcelRustMappingExtensions.ExportMapped(invalidPath, new[] { new MappedFormula() }, missingCell);
    }
    catch (InvalidOperationException)
    {
      rejectedMissingCell = true;
    }
    Require(rejectedMissingCell, "fluent-mapping: missing ToCell should fail.");

    var missingStart = new MiniExcelRustMapping<MappedDepartment>();
    missingStart.Collection(value => value.PhoneNumbers);
    var rejectedMissingStart = false;
    try
    {
      MiniExcelRustMappingExtensions.ExportMapped(invalidPath, new[] { new MappedDepartment() }, missingStart);
    }
    catch (InvalidOperationException)
    {
      rejectedMissingStart = true;
    }
    Require(rejectedMissingStart, "fluent-mapping: missing StartAt should fail.");
  }
  finally
  {
    foreach (var candidate in new[] { path, formulaPath, formulaTemplatePath, managedPath, oraclePath, templatePath, mappedTemplatePath, invalidPath })
      if (File.Exists(candidate))
        File.Delete(candidate);
  }
}

static void VerifyMissingDimensionParity(string sourcePath)
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-no-dimension-{Guid.NewGuid():N}.xlsx");
  File.Copy(sourcePath, path);
  try
  {
    using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
    {
      foreach (var entryName in new[]
      {
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet3.xml"
      })
      {
        var entry = archive.GetEntry(entryName) ?? throw new InvalidDataException($"Missing {entryName}.");
        XDocument document;
        using (var input = entry.Open())
          document = XDocument.Load(input);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        document.Root?.Element(spreadsheet + "dimension")?.Remove();
        entry.Delete();
        var replacement = archive.CreateEntry(entryName, CompressionLevel.Fastest);
        using var output = replacement.Open();
        document.Save(output);
      }
    }

    var importer = ManagedMiniExcel.Importers.GetOpenXmlImporter();
    var managed = importer.GetSheetDimensions(path);
    var rust = MiniExcelRust.GetSheetDimensions(path);
    Require(managed.Count == rust.Count, "missing-dimension: sheet count differs.");
    for (var index = 0; index < managed.Count; index++)
    {
      Require(managed[index].StartCell == rust[index].StartCell, $"missing-dimension: start differs at {index}.");
      Require(managed[index].EndCell == rust[index].EndCell, $"missing-dimension: end differs at {index}.");
    }
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifySelfClosingEmptyRowParity(string sourcePath)
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-self-closing-{Guid.NewGuid():N}.xlsx");
  File.Copy(sourcePath, path);
  try
  {
    using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
    {
      const string entryName = "xl/worksheets/sheet3.xml";
      var entry = archive.GetEntry(entryName) ?? throw new InvalidDataException($"Missing {entryName}.");
      string xml;
      using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
        xml = reader.ReadToEnd();
      xml = xml.Replace("<row r=\"4\">", "<row r=\"3\"/><row r=\"4\">", StringComparison.Ordinal);
      entry.Delete();
      var replacement = archive.CreateEntry(entryName, CompressionLevel.Fastest);
      using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
      writer.Write(xml);
    }

    var managedConfiguration = new OpenXmlConfiguration { IgnoreEmptyRows = true };
    var rustConfiguration = new MiniExcelRustReadOptions { IgnoreEmptyRows = true };
    var managed = QueryManaged(path, true, "Options", "A1", managedConfiguration).ToList();
    var rust = MiniExcelRust.Query(path, true, "Options", "A1", rustConfiguration).ToList();
    CompareRows(managed, rust, "self-closing-empty-row");
    Require(rust.Count == 2, $"self-closing-empty-row: expected 2 rows, received {rust.Count}.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static void VerifyConversions()
{
  var csvPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-convert-{Guid.NewGuid():N}.csv");
  var xlsxPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-convert-{Guid.NewGuid():N}.xlsx");
  var roundtripPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-roundtrip-{Guid.NewGuid():N}.csv");
  try
  {
    File.WriteAllText(csvPath, "Name,Value\r\nalpha,1\r\nbeta,2\r\n", new UTF8Encoding(true));
    var expected = MiniExcelRust.QueryCsv(csvPath, true).ToList();
    MiniExcelRust.ConvertCsvToXlsxAsync(csvPath, xlsxPath, true).GetAwaiter().GetResult();
    CompareRows(expected, MiniExcelRust.Query(xlsxPath, true).ToList(), "convert-csv-xlsx");
    MiniExcelRust.ConvertXlsxToCsvAsync(xlsxPath, roundtripPath, true).GetAwaiter().GetResult();
    CompareRows(expected, MiniExcelRust.QueryCsv(roundtripPath, true).ToList(), "convert-xlsx-csv");

    using var csvInput = new MemoryStream(File.ReadAllBytes(csvPath));
    using var xlsxOutput = new MemoryStream();
    MiniExcelRust.ConvertCsvToXlsx(csvInput, xlsxOutput, true);
    Require(csvInput.CanRead && xlsxOutput.CanWrite, "convert-stream: stream ownership changed.");
    xlsxOutput.Position = 0;
    using var csvOutput = new MemoryStream();
    MiniExcelRust.ConvertXlsxToCsv(xlsxOutput, csvOutput, true);
    csvOutput.Position = 0;
    CompareRows(expected, MiniExcelRust.QueryCsv(csvOutput, true, leaveOpen: true).ToList(), "convert-stream-roundtrip");
  }
  finally
  {
    foreach (var path in new[] { csvPath, xlsxPath, roundtripPath })
    {
      if (File.Exists(path))
        File.Delete(path);
    }
  }
}

static void VerifyCompatibilityFacade(string path)
{
  var facade = typeof(MiniExcelRust).Assembly.GetType("MiniExcelLibs.MiniExcel")
    ?? throw new InvalidOperationException("compatibility-facade: type was not found.");
  var assembly = typeof(MiniExcelRust).Assembly;
  var excelType = assembly.GetType("MiniExcelLibs.ExcelType")
    ?? throw new InvalidOperationException("compatibility-facade: ExcelType was not found.");
  var configuration = assembly.GetType("MiniExcelLibs.IConfiguration")
    ?? throw new InvalidOperationException("compatibility-facade: IConfiguration was not found.");
  Require(assembly.GetType("MiniExcelLibs.OpenXml.OpenXmlConfiguration") is not null, "compatibility-facade: OpenXmlConfiguration was not found.");
  Require(assembly.GetType("MiniExcelLibs.Csv.CsvConfiguration") is not null, "compatibility-facade: CsvConfiguration was not found.");
  Require(
    facade.GetMethods().Any(candidate =>
      candidate.Name == "Query" &&
      candidate.GetParameters().Any(parameter => parameter.ParameterType == excelType) &&
      candidate.GetParameters().Any(parameter => parameter.ParameterType == configuration)),
    "compatibility-facade: configured Query overload was not found.");
  var methodNames = facade.GetMethods().Select(method => method.Name).ToHashSet(StringComparer.Ordinal);
  foreach (var required in new[]
  {
    "AddPicture", "AddPictureAsync", "Query", "QueryAsync", "QueryRange", "QueryRangeAsync", "SaveAs", "SaveAsAsync",
    "Insert", "SaveAsByTemplate", "MergeSameCells", "MergeSameCellsAsync", "GetReader",
    "QueryAsDataTable", "QueryAsDataTableAsync", "GetSheetNames", "GetSheetInformations",
    "GetSheetDimensions", "GetColumns", "GetColumnsAsync",
    "ConvertCsvToXlsx", "ConvertCsvToXlsxAsync", "ConvertXlsxToCsv", "ConvertXlsxToCsvAsync"
  })
    Require(methodNames.Contains(required), $"compatibility-facade: {required} was not found.");
  var expectedOverloadCounts = new Dictionary<string, int>(StringComparer.Ordinal)
  {
    ["AddPicture"] = 2,
    ["AddPictureAsync"] = 2,
    ["ConvertCsvToXlsx"] = 2,
    ["ConvertCsvToXlsxAsync"] = 2,
    ["ConvertXlsxToCsv"] = 2,
    ["ConvertXlsxToCsvAsync"] = 2,
    ["GetColumns"] = 2,
    ["GetColumnsAsync"] = 2,
    ["GetReader"] = 2,
    ["GetSheetDimensions"] = 2,
    ["GetSheetDimensionsAsync"] = 2,
    ["GetSheetInformations"] = 2,
    ["GetSheetInformationsAsync"] = 2,
    ["GetSheetNames"] = 2,
    ["GetSheetNamesAsync"] = 2,
    ["Insert"] = 2,
    ["InsertAsync"] = 2,
    ["MergeSameCells"] = 3,
    ["MergeSameCellsAsync"] = 3,
    ["Query"] = 4,
    ["QueryAsync"] = 4,
    ["QueryRange"] = 4,
    ["QueryRangeAsync"] = 4,
    ["QueryAsDataTable"] = 2,
    ["QueryAsDataTableAsync"] = 2,
    ["SaveAs"] = 2,
    ["SaveAsAsync"] = 2,
    ["SaveAsByTemplate"] = 6,
    ["SaveAsByTemplateAsync"] = 6
  };
  foreach (var expected in expectedOverloadCounts)
  {
    var actual = facade.GetMethods().Count(candidate => candidate.Name == expected.Key);
    Require(
      actual >= expected.Value,
      $"compatibility-facade: {expected.Key} has {actual} overloads; baseline requires {expected.Value}.");
  }
  var method = facade.GetMethod("GetSheetNames", new[] { typeof(string) })
    ?? throw new InvalidOperationException("compatibility-facade: GetSheetNames was not found.");
  var names = (List<string>?)method.Invoke(null, new object[] { path })
    ?? throw new InvalidOperationException("compatibility-facade: GetSheetNames returned null.");
  Require(
    names.SequenceEqual(new[] { "Sheet1", "Data", "Options" }, StringComparer.Ordinal),
    "compatibility-facade: sheet names differ.");

  var outputPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-facade-{Guid.NewGuid():N}.xlsx");
  try
  {
    var dataSet = new DataSet();
    var first = new DataTable("First");
    first.Columns.Add("Name", typeof(string));
    first.Rows.Add("one");
    var second = new DataTable("Second");
    second.Columns.Add("Value", typeof(double));
    second.Rows.Add(2d);
    dataSet.Tables.Add(first);
    dataSet.Tables.Add(second);
    var saveAs = facade.GetMethods().Single(candidate =>
      candidate.Name == "SaveAs" &&
      !candidate.IsGenericMethod &&
      candidate.GetParameters().Length == 7 &&
      candidate.GetParameters()[0].ParameterType == typeof(string) &&
      candidate.GetParameters()[1].ParameterType == typeof(object));
    var unknown = Enum.Parse(excelType, "UNKNOWN");
    var counts = (int[]?)saveAs.Invoke(
      null,
      new object?[] { outputPath, dataSet, true, "Sheet1", unknown, null, false });
    Require(counts?.SequenceEqual(new[] { 1, 1 }) is true, "compatibility-facade: DataSet counts differ.");
    Require(
      MiniExcelRust.GetSheetNames(outputPath).SequenceEqual(new[] { "First", "Second" }, StringComparer.Ordinal),
      "compatibility-facade: DataSet sheet names differ.");
  }
  finally
  {
    if (File.Exists(outputPath))
      File.Delete(outputPath);
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
    var progress = new CountingProgress();
    var written = MiniExcelRust.SaveAs(path, rows, sheetName: "Exported", progress: progress);
    Require(written == rows.Count, $"save-as: expected {rows.Count} written rows, received {written}.");
    Require(progress.Count == rows.Sum(row => row.Count), "save-as: progress count differs.");
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
    Require(progress.Count == rows.Sum(row => row.Count), "save-as: failed write changed progress.");

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
    var counts = MiniExcelRust.SaveAsSheets(path, sheets);
    Require(counts.SequenceEqual(new[] { 2, 1 }), "multi-sheet-save: row counts differ.");
    Require(
      MiniExcelRust.GetSheetNames(path).SequenceEqual(new[] { "First", "Second" }, StringComparer.Ordinal),
      "multi-sheet-save: sheet order differs.");
    CompareRows(firstRows, QueryManaged(path, true, "First").ToList(), "multi-sheet-first");
    CompareRows(secondRows, QueryManaged(path, true, "Second").ToList(), "multi-sheet-second");

    var rejectedOverwrite = false;
    try
    {
      MiniExcelRust.SaveAsSheets(path, sheets);
    }
    catch (InvalidOperationException)
    {
      rejectedOverwrite = true;
    }
    Require(rejectedOverwrite, "multi-sheet-save: overwrite=false should reject an existing file.");
    counts = MiniExcelRust.SaveAsSheets(path, sheets, overwriteFile: true);
    Require(counts.SequenceEqual(new[] { 2, 1 }), "multi-sheet-save: overwrite counts differ.");

    using var stream = new MemoryStream();
    counts = MiniExcelRust.SaveAsSheets(stream, sheets, leaveOpen: true);
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

static void VerifyConfiguredWrite()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-configured-{Guid.NewGuid():N}.xlsx");
  var schema = new[] { "Name", "Amount", "Secret" };
  var rows = new[]
  {
    new Dictionary<string, object?> { ["Amount"] = 12.5d, ["Name"] = "alpha", ["Secret"] = "hidden" }
  };
  var options = new MiniExcelRustWriteOptions
  {
    SheetName = "Styled",
    AutoFilter = true,
    RightToLeft = true,
    WrapCellContents = true,
    HorizontalAlignment = MiniExcelRustHorizontalAlignment.Center,
    VerticalAlignment = MiniExcelRustVerticalAlignment.Top,
    TableStyle = MiniExcelRustTableStyle.Default,
    HeaderWrapText = true,
    HeaderBackgroundColor = "2F5597",
    HeaderHorizontalAlignment = MiniExcelRustHorizontalAlignment.Center,
    HeaderVerticalAlignment = MiniExcelRustVerticalAlignment.Top,
    FreezeRowCount = 2,
    FreezeColumnCount = 1
  };
  options.ColumnFormats["Amount"] = "0.00";
  options.ColumnWidths["Amount"] = 22;
  options.HiddenColumns["Secret"] = true;
  try
  {
    var written = MiniExcelRust.SaveAsWithSchema(path, schema, rows, options);
    Require(written == 1, "configured-write: row count differs.");
    var managedRows = QueryManaged(path, true, "Styled").ToList();
    Require(managedRows[0].Keys.SequenceEqual(schema, StringComparer.Ordinal), "configured-write: schema order differs.");
    Require(Equals(managedRows[0]["Name"], "alpha"), "configured-write: data differs.");

    var worksheetXml = ReadZipEntryText(path, "xl/worksheets/sheet1.xml");
    Require(worksheetXml.Contains("rightToLeft=\"1\"", StringComparison.Ordinal), "configured-write: RTL missing.");
    Require(worksheetXml.Contains("xSplit=\"1\"", StringComparison.Ordinal), "configured-write: frozen column missing.");
    Require(worksheetXml.Contains("ySplit=\"2\"", StringComparison.Ordinal), "configured-write: frozen rows missing.");
    Require(worksheetXml.Contains("<autoFilter", StringComparison.Ordinal), "configured-write: auto filter missing.");
    Require(worksheetXml.Contains("hidden=\"1\"", StringComparison.Ordinal), "configured-write: hidden column missing.");
    Require(worksheetXml.Contains("width=\"22", StringComparison.Ordinal), "configured-write: column width missing.");
    var stylesXml = ReadZipEntryText(path, "xl/styles.xml");
    Require(stylesXml.Contains("formatCode=\"0.00\"", StringComparison.Ordinal), "configured-write: number format missing.");
    Require(stylesXml.Contains("FF2F5597", StringComparison.Ordinal), "configured-write: header color missing.");
    Require(stylesXml.Contains("horizontal=\"center\"", StringComparison.Ordinal), "configured-write: horizontal alignment missing.");
    Require(stylesXml.Contains("vertical=\"top\"", StringComparison.Ordinal), "configured-write: vertical alignment missing.");
    Require(stylesXml.Contains("wrapText=\"1\"", StringComparison.Ordinal), "configured-write: wrapping missing.");

    options.TableStyle = MiniExcelRustTableStyle.None;
    options.OverwriteFile = true;
    MiniExcelRust.SaveAsWithSchema(path, schema, rows, options);
    stylesXml = ReadZipEntryText(path, "xl/styles.xml");
    Require(!stylesXml.Contains("FF2F5597", StringComparison.Ordinal), "configured-write: TableStyle.None retained header fill.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
  }
}

static string ReadZipEntryText(string path, string entryName)
{
  using var archive = ZipFile.OpenRead(path);
  var entry = archive.GetEntry(entryName) ?? throw new InvalidDataException($"Missing ZIP entry {entryName}.");
  using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
  return reader.ReadToEnd();
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
  var culturePath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-culture-{Guid.NewGuid():N}.csv");
  var localizedPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-localized-{Guid.NewGuid():N}.csv");
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

    File.WriteAllText(culturePath, "Amount;Date\r\n12,5;06.09.2026\r\n", new UTF8Encoding(true));
    var culture = CultureInfo.GetCultureInfo("de-DE");
    var managedCulture = new CsvConfiguration { Seperator = ';', Culture = culture };
    var rustCulture = new MiniExcelRustCsvReadOptions { Delimiter = ';', Culture = culture };
    var managedFormatted = importer.Query<TypedFormattedRow>(culturePath, configuration: managedCulture).Single();
    var rustFormatted = MiniExcelRust.QueryCsv<TypedFormattedRow>(culturePath, configuration: rustCulture).Single();
    Require(managedFormatted.Amount == rustFormatted.Amount && rustFormatted.Amount == 12.5d, "typed-culture: amount differs.");
    Require(managedFormatted.Date == rustFormatted.Date && rustFormatted.Date == new DateTime(2026, 9, 6), "typed-format: date differs.");

    File.WriteAllText(localizedPath, "名稱\r\nAda\r\n", new UTF8Encoding(true));
    var localizedCulture = CultureInfo.GetCultureInfo("zh-TW");
    var localizedOptions = new MiniExcelRustCsvReadOptions { Culture = localizedCulture };
    var rustLocalized = MiniExcelRust.QueryCsv<TypedLocalizedRow>(localizedPath, configuration: localizedOptions).Single();
    Require(rustLocalized.Name == "Ada", $"typed-localization: rust={rustLocalized.Name ?? "<null>"}.");

    var dynamicOptions = new MiniExcelRustCsvReadOptions();
    dynamicOptions.DynamicColumns["Label"] = new MiniExcelRustDynamicColumn { Name = "Identifier" };
    var dynamicMapped = MiniExcelRust.QueryCsv<TypedDynamicRow>(path, configuration: dynamicOptions).Single();
    Require(dynamicMapped.Label == identifier, "typed-dynamic-column: value differs.");

    MiniExcelRustColumnNotFoundException? missingColumn = null;
    try
    {
      _ = MiniExcelRust.QueryCsv<TypedMissingColumnRow>(path).ToList();
    }
    catch (MiniExcelRustColumnNotFoundException error)
    {
      missingColumn = error;
    }
    Require(
      missingColumn?.ColumnName == "Missing" && missingColumn.Row == 1,
      "typed-mapping: missing-column metadata differs.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
    if (File.Exists(culturePath))
      File.Delete(culturePath);
    if (File.Exists(localizedPath))
      File.Delete(localizedPath);
  }
}

static void VerifyTypedExports()
{
  var xlsxPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-typed-export-{Guid.NewGuid():N}.xlsx");
  var asyncPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-async-export-{Guid.NewGuid():N}.xlsx");
  var cancelledPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-cancelled-export-{Guid.NewGuid():N}.xlsx");
  var nativeCancelledPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-native-cancelled-{Guid.NewGuid():N}.xlsx");
  var csvPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-typed-export-{Guid.NewGuid():N}.csv");
  var asyncCsvPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-async-export-{Guid.NewGuid():N}.csv");
  var cancelledCsvPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-cancelled-export-{Guid.NewGuid():N}.csv");
  var attributesPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-attribute-export-{Guid.NewGuid():N}.xlsx");
  var formatterPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-formatter-export-{Guid.NewGuid():N}.xlsx");
  var formulaPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-formula-export-{Guid.NewGuid():N}.xlsx");
  var identifier = Guid.Parse("32a2fac7-2ce2-4683-a735-118fe7c4949b");
  var rows = new[]
  {
    new TypedExportRow
    {
      Name = "Ada",
      Count = 7,
      State = RowState.Ready,
      Identifier = identifier,
      When = new DateTime(2026, 9, 6, 8, 30, 0)
    }
  };
  try
  {
    var written = MiniExcelRust.SaveAs(xlsxPath, rows, sheetName: "Typed");
    Require(written == 1, "typed-export: XLSX row count differs.");
    var managed = ManagedMiniExcel.Importers.GetOpenXmlImporter().Query<TypedExportRow>(xlsxPath, "Typed").Single();
    var rust = MiniExcelRust.Query<TypedExportRow>(xlsxPath, "Typed").Single();
    CompareTypedExportRows(managed, rust, "typed-export");

    written = MiniExcelRust.SaveAsCsv(csvPath, rows);
    Require(written == 1, "typed-export: CSV row count differs.");
    var rustCsv = MiniExcelRust.QueryCsv<TypedExportRow>(csvPath).Single();
    Require(rustCsv.Name == rows[0].Name && rustCsv.Identifier == identifier, "typed-export: CSV values differ.");

    using (var csvStream = new MemoryStream())
    {
      written = MiniExcelRust.SaveAsCsv(csvStream, rows, leaveOpen: true);
      Require(written == 1 && csvStream.CanWrite, "typed-export: CSV stream write failed.");
      written = MiniExcelRust.AppendCsv(csvStream, rows, leaveOpen: true);
      Require(written == 1, "typed-export: CSV stream append failed.");
      csvStream.Position = 0;
      Require(
        MiniExcelRust.QueryCsv<TypedExportRow>(csvStream, leaveOpen: true).Count() == 2,
        "typed-export: CSV stream row count differs.");
    }

    written = MiniExcelRust.SaveAsCsvAsync(asyncCsvPath, ProduceTypedRows(rows)).GetAwaiter().GetResult();
    Require(written == 1, "typed-export: async CSV row count differs.");
    Require(MiniExcelRust.QueryCsv<TypedExportRow>(asyncCsvPath).Single().Identifier == identifier, "typed-export: async CSV value differs.");

    using var csvCancellation = new CancellationTokenSource();
    var cancelled = false;
    try
    {
      _ = MiniExcelRust.SaveAsCsvAsync(
          cancelledCsvPath,
          ProduceRowsAndCancel(csvCancellation),
          cancellationToken: csvCancellation.Token)
        .GetAwaiter()
        .GetResult();
    }
    catch (OperationCanceledException)
    {
      cancelled = true;
    }
    Require(cancelled && !File.Exists(cancelledCsvPath), "typed-export: native CSV cancellation published a destination.");

    written = MiniExcelRust.SaveAsAsync(asyncPath, ProduceTypedRows(rows), sheetName: "Async")
      .GetAwaiter()
      .GetResult();
    Require(written == 1, "typed-export: async row count differs.");
    CompareTypedExportRows(rows[0], MiniExcelRust.Query<TypedExportRow>(asyncPath, "Async").Single(), "typed-async-export");

    using var cancellation = new CancellationTokenSource();
    cancellation.Cancel();
    cancelled = false;
    try
    {
      _ = MiniExcelRust.SaveAsAsync(
          cancelledPath,
          ProduceTypedRows(rows),
          cancellationToken: cancellation.Token)
        .GetAwaiter()
        .GetResult();
    }
    catch (OperationCanceledException)
    {
      cancelled = true;
    }
    Require(cancelled && !File.Exists(cancelledPath), "typed-export: cancellation published a destination.");

    using var nativeCancellation = new CancellationTokenSource();
    cancelled = false;
    try
    {
      _ = MiniExcelRust.SaveAsAsync(
          nativeCancelledPath,
          ProduceRowsAndCancel(nativeCancellation),
          cancellationToken: nativeCancellation.Token)
        .GetAwaiter()
        .GetResult();
    }
    catch (OperationCanceledException)
    {
      cancelled = true;
    }
    Require(cancelled && !File.Exists(nativeCancelledPath), "typed-export: native cancellation published a destination.");

    written = MiniExcelRust.SaveAs(
      attributesPath,
      new[] { new TypedAttributeExportRow("visible", "ignored", 9) },
      sheetName: "Attributes");
    Require(written == 1, "typed-attributes: row count differs.");
    var attributeRows = MiniExcelRust.Query(attributesPath, true, "Attributes").ToList();
    Require(
      attributeRows[0].Keys.SequenceEqual(new[] { "Renamed", "Readonly" }, StringComparer.Ordinal),
      "typed-attributes: exported columns differ.");
    Require(Equals(attributeRows[0]["Readonly"], 9d), "typed-attributes: readonly value differs.");

    var writeOptions = new MiniExcelRustWriteOptions { SheetName = "Formatted" };
    writeOptions.DynamicColumns["Count"] = new MiniExcelRustDynamicColumn
    {
      Name = "Formatted",
      CustomFormatter = value => $"#{value}"
    };
    written = MiniExcelRust.SaveAs(
      formatterPath,
      new[] { new TypedAttributeExportRow("visible", "ignored", 9) },
      writeOptions);
    Require(written == 1, "typed-formatter: row count differs.");
    var formatterRows = MiniExcelRust.Query(formatterPath, true, "Formatted").ToList();
    Require(
      formatterRows[0].ContainsKey("Formatted") && Equals(formatterRows[0]["Formatted"], "#9"),
      "typed-formatter: formatted value differs.");

    var formulaOptions = new MiniExcelRustWriteOptions { SheetName = "Formula" };
    formulaOptions.DynamicColumns["Formula"] = new MiniExcelRustDynamicColumn { IsFormula = true };
    written = MiniExcelRust.SaveAs(
      formulaPath,
      new[] { new TypedFormulaRow { Left = 2, Right = 3, Formula = "=A2+B2" } },
      formulaOptions);
    Require(written == 1, "typed-formula: row count differs.");
    var formulaXml = ReadZipEntryText(formulaPath, "xl/worksheets/sheet1.xml");
    Require(formulaXml.Contains("<f>A2+B2</f>", StringComparison.Ordinal), "typed-formula: formula cell was not written.");
    Require(!formulaXml.Contains("$=A2+B2", StringComparison.Ordinal), "typed-formula: template marker leaked.");
  }
  finally
  {
    foreach (var path in new[] { xlsxPath, asyncPath, cancelledPath, nativeCancelledPath, csvPath, asyncCsvPath, cancelledCsvPath, attributesPath, formatterPath, formulaPath })
    {
      if (File.Exists(path))
        File.Delete(path);
    }
  }
}

static async IAsyncEnumerable<TypedExportRow> ProduceTypedRows(IEnumerable<TypedExportRow> rows)
{
  foreach (var row in rows)
  {
    await Task.Yield();
    yield return row;
  }
}

static async IAsyncEnumerable<TypedExportRow> ProduceRowsAndCancel(CancellationTokenSource cancellation)
{
  for (var index = 0; index < 20_000; index++)
  {
    yield return new TypedExportRow
    {
      Name = $"row-{index}",
      Count = index,
      State = RowState.Ready,
      Identifier = Guid.Empty,
      When = new DateTime(2026, 9, 6)
    };
  }
  cancellation.CancelAfter(1);
  await Task.Yield();
}

static void CompareTypedExportRows(TypedExportRow expected, TypedExportRow actual, string scenario)
{
  Require(expected.Name == actual.Name, $"{scenario}: name differs.");
  Require(expected.Count == actual.Count, $"{scenario}: count differs.");
  Require(expected.State == actual.State, $"{scenario}: state differs.");
  Require(expected.Identifier == actual.Identifier, $"{scenario}: identifier differs.");
  Require(expected.When == actual.When, $"{scenario}: timestamp differs.");
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

static void VerifyPictures()
{
  var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-picture-{Guid.NewGuid():N}.xlsx");
  var png = Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2n0kAAAAASUVORK5CYII=");
  try
  {
    MiniExcelRust.SaveAs(
      path,
      new[] { new Dictionary<string, object?> { ["Name"] = "preserved" } });
    MiniExcelRust.AddPicture(
      path,
      new MiniExcelRustPicture
      {
        ImageBytes = png,
        CellAddress = "B2",
        WidthPx = 40,
        HeightPx = 40,
        Anchor = MiniExcelRustPictureAnchor.OneCell
      },
      new MiniExcelRustPicture
      {
        ImageBytes = png,
        CellAddress = "C3",
        WidthPx = 45,
        HeightPx = 45,
        Anchor = MiniExcelRustPictureAnchor.Absolute,
        LocationX = 20,
        LocationY = 30
      },
      new MiniExcelRustPicture
      {
        ImageBytes = png,
        CellAddress = "D4",
        WidthPx = 50,
        HeightPx = 50,
        Anchor = MiniExcelRustPictureAnchor.TwoCell
      });

    using (var archive = ZipFile.OpenRead(path))
    {
      Require(archive.Entries.Count(entry => entry.FullName.StartsWith("xl/media/image", StringComparison.Ordinal)) == 3, "pictures: media count differs.");
      Require(archive.GetEntry("xl/drawings/drawing1.xml") is not null, "pictures: drawing part missing.");
      Require(archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels") is not null, "pictures: drawing relationships missing.");
    }
    var drawing = ReadZipEntryText(path, "xl/drawings/drawing1.xml");
    Require(drawing.Contains("oneCellAnchor", StringComparison.Ordinal), "pictures: one-cell anchor missing.");
    Require(drawing.Contains("absoluteAnchor", StringComparison.Ordinal), "pictures: absolute anchor missing.");
    Require(drawing.Contains("twoCellAnchor", StringComparison.Ordinal), "pictures: two-cell anchor missing.");
    Require(Equals(MiniExcelRust.Query(path, true).Single()["Name"], "preserved"), "pictures: workbook data changed.");

    using var stream = new MemoryStream();
    stream.Write(File.ReadAllBytes(path));
    stream.Position = 0;
    MiniExcelRust.AddPicture(
      stream,
      leaveOpen: true,
      new MiniExcelRustPicture { ImageBytes = png, CellAddress = "E5" });
    Require(stream.CanRead, "pictures-stream: leaveOpen should preserve the stream.");
    stream.Position = 0;
    using var streamArchive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
    Require(streamArchive.Entries.Count(entry => entry.FullName.StartsWith("xl/media/image", StringComparison.Ordinal)) == 4, "pictures-stream: media count differs.");
  }
  finally
  {
    if (File.Exists(path))
      File.Delete(path);
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
  var asyncTypedCsv = CollectAsync(MiniExcelRust.QueryCsvAsync<TypedCsvRow>(path, configuration: rustConfiguration))
    .GetAwaiter()
    .GetResult();
  Require(asyncTypedCsv.Count == managedTypedRows.Count, "csv-typed-async: row count differs.");

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

  var cellMap = new Dictionary<string, string>
  {
    ["Name"] = "A2",
    ["Value"] = "B2",
    ["Note"] = "C2"
  };
  var mapped = MiniExcelRust.ReadMapped<TypedSheetRow>(path, cellMap, "Sheet1");
  Require(mapped.Name == "alpha" && Equals(mapped.Value, 42d) && mapped.Note == "Taiwan 台灣", "cell-map: values differ.");
  using (var mappedStream = File.OpenRead(path))
  {
    var streamMapped = MiniExcelRust.ReadMapped<TypedSheetRow>(mappedStream, cellMap, "Sheet1", leaveOpen: true);
    Require(streamMapped.Name == mapped.Name && mappedStream.CanRead, "cell-map-stream: values or ownership differ.");
  }

  var managedRangeRows = QueryManagedRange(path, true, "Data", "C2", "D3").ToList();
  var rustRangeRows = MiniExcelRust.QueryRange(path, true, "Data", "C2", "D3").ToList();
  CompareRows(managedRangeRows, rustRangeRows, "bounded-range");
  Require(rustRangeRows.Count == 1, $"bounded-range: expected 1 row, received {rustRangeRows.Count}.");

  var managedNumericRange = importer.QueryRange(
      path,
      hasHeaderRow: true,
      sheetName: "Data",
      startRowIndex: 2,
      startColumnIndex: 3,
      endRowIndex: 3,
      endColumnIndex: 4)
    .Cast<IDictionary<string, object?>>()
    .ToList();
  var rustNumericRange = MiniExcelRust.QueryRange(path, true, "Data", 2, 3, 3, 4).ToList();
  CompareRows(managedNumericRange, rustNumericRange, "numeric-range");

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

  managedConfiguration.FillMergedCells = true;
  rustConfiguration.FillMergedCells = true;
  var managedMergedRows = QueryManaged(path, true, "Options", "A1", managedConfiguration).ToList();
  var rustMergedRows = MiniExcelRust.Query(path, true, "Options", "A1", rustConfiguration).ToList();
  CompareRows(managedMergedRows, rustMergedRows, "merged-fill");
  Require(rustMergedRows[0]["Right"] is null, "merged-fill: physically absent merged cell should be null.");

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
  var mappedFields = MiniExcelRust.Query<TypedFieldRow>(path, "Sheet1").ToList();
  Require(
    mappedFields[0].Label == "alpha" && Equals(mappedFields[0].Value, 42d),
    "typed-field-index: values differ.");

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
  var asyncTypedTableRows = CollectAsync(MiniExcelRust.QueryTableAsync<TypedTableRow>(path, "Data", "DataTable"))
    .GetAwaiter()
    .GetResult();
  Require(asyncTypedTableRows.Count == 2, "async-typed-table: row count differs.");

  using (var asyncStream = new MemoryStream(File.ReadAllBytes(path)))
  {
    var streamAsyncRows = CollectAsync(
        MiniExcelRust.QueryAsync(asyncStream, true, "Sheet1", leaveOpen: true))
      .GetAwaiter()
      .GetResult();
    CompareRows(rows, streamAsyncRows, "async-stream-query");
    Require(asyncStream.CanRead, "async-stream-query: leaveOpen should preserve the stream.");
  }

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

  using var multiReader = MiniExcelRust.GetReader(path, true);
  var resultSets = 0;
  do
  {
    resultSets++;
    while (multiReader.Read()) { }
  }
  while (multiReader.NextResult());
  Require(resultSets == 3, $"data-reader: expected 3 result sets, received {resultSets}.");

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

  using (var stream = new MemoryStream(bytes))
  {
    using var reader = MiniExcelRust.GetReader(stream, true, leaveOpen: true);
    var resultSets = 1;
    while (reader.NextResult())
      resultSets++;
    Require(resultSets == 3, "stream-data-reader: result set count differs.");
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

  internal sealed class TypedFieldRow
  {
    [ExcelColumnName("Name")]
    public string? Label = null;

    [ExcelColumnIndex(1)]
    public object? Value = null;
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

  internal sealed class TypedExportRow
  {
    [ExcelColumnName("Display Name")]
    public string? Name { get; set; }
    public int Count { get; set; }
    public RowState State { get; set; }
    public Guid Identifier { get; set; }
    public DateTime When { get; set; }
  }

  internal sealed class TypedAttributeExportRow(string name, string ignored, int count)
  {
    [ExcelColumnName("Renamed")]
    public string Name { get; set; } = name;

    [ExcelIgnore]
    public string Ignored { get; set; } = ignored;

    [System.ComponentModel.DisplayName("Readonly")]
    public int Count { get; } = count;
  }

  internal sealed class TypedFormulaRow
  {
    public int Left { get; set; }
    public int Right { get; set; }
    public string? Formula { get; set; }
  }

  internal sealed class MappedDepartment
  {
    public string? Name { get; set; }
    public List<string> PhoneNumbers { get; set; } = [];
    public List<MappedProject> Projects { get; set; } = [];
  }

  internal sealed class MappedProject
  {
    public string? Code { get; set; }
    public List<MappedTask> Tasks { get; set; } = [];
  }

  internal sealed class MappedTask
  {
    public string? Name { get; set; }
    public int Hours { get; set; }
  }

  internal sealed class MappedFormula
  {
    public string? Name { get; set; }
    public double Amount { get; set; }
    public double Total { get; set; }
  }

  internal enum RowState
  {
    Unknown,
    Ready
  }

  internal sealed class CountingProgress : IProgress<int>
  {
    public int Count { get; private set; }

    public void Report(int value) => Count += value;
  }

  internal sealed class TypedFormattedRow
  {
    public double Amount { get; set; }

    [ExcelFormat("dd.MM.yyyy")]
    public DateTime Date { get; set; }
  }

  internal sealed class TypedMissingColumnRow
  {
    public string? Missing { get; set; }
  }

  internal sealed class TypedLocalizedRow
  {
    [ExcelColumnName("NameKey", ResourceType = typeof(TestLocalization))]
    public string? Name { get; set; }
  }

  internal sealed class TypedDynamicRow
  {
    public Guid Label { get; set; }
  }

  internal static class TestLocalization
  {
    public static ResourceManager ResourceManager { get; } = new TestResourceManager();
  }

  internal sealed class TestResourceManager : ResourceManager
  {
    public override string? GetString(string name, CultureInfo? culture) =>
      name == "NameKey" && culture?.Name == "zh-TW" ? "名稱" : name;
  }