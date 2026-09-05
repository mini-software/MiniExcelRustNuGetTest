using System.IO.Compression;
using System.Text;
using MiniExcelLibs;

var workbookPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-{Guid.NewGuid():N}.xlsx");

try
{
    CreateWorkbook(workbookPath);

    var rows = MiniExcelRust.Query(workbookPath, useHeaderRow: true).ToList();
    Require(rows.Count == 2, $"Expected 2 rows, received {rows.Count}.");
    Require(Equals(rows[0]["Name"], "alpha"), "The first string value did not match.");
    Require(Equals(rows[0]["Value"], 42d), "The first numeric value did not match.");
    Require(Equals(rows[1]["Name"], "beta"), "The second string value did not match.");
    Require(Equals(rows[1]["Value"], true), "The boolean value did not match.");

    using (var enumerator = MiniExcelRust.Query(workbookPath).GetEnumerator())
        Require(enumerator.MoveNext(), "The early-disposal query returned no rows.");

    File.Delete(workbookPath);
    Console.WriteLine("MiniExcelRust package smoke test passed.");
}
finally
{
    if (File.Exists(workbookPath))
        File.Delete(workbookPath);
}

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
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
    AddEntry(archive, "xl/worksheets/sheet1.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>42</v></c></row>
            <row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3" t="b"><v>1</v></c></row>
          </sheetData>
        </worksheet>
        """);
}

static void AddEntry(ZipArchive archive, string name, string contents)
{
    var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
    writer.Write(contents);
}