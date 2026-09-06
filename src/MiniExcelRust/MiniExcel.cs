using System.Data;

namespace MiniExcelLibs;

/// <summary>
/// Compatibility facade whose workbook operations are backed exclusively by MiniExcel Rust.
/// </summary>
public static class MiniExcel
{
    public static IEnumerable<IDictionary<string, object?>> Query(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null) =>
        IsCsv(path, excelType)
            ? MiniExcelRust.QueryCsv(path, useHeaderRow, CsvOptions(configuration))
            : MiniExcelRust.Query(path, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration));

    public static IEnumerable<T> Query<T>(
        string path,
        string? sheetName = null,
        string startCell = "A1",
        bool treatHeaderAsData = false,
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null)
        where T : class, new() =>
        IsCsv(path, excelType)
            ? MiniExcelRust.QueryCsv<T>(path, treatHeaderAsData, CsvOptions(configuration))
            : MiniExcelRust.Query<T>(path, sheetName, startCell, treatHeaderAsData, OpenXmlOptions(configuration));

    public static IEnumerable<IDictionary<string, object?>> Query(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.XLSX,
        IConfiguration? configuration = null,
        bool leaveOpen = false) =>
        excelType == ExcelType.CSV
            ? MiniExcelRust.QueryCsv(stream, useHeaderRow, CsvOptions(configuration), leaveOpen)
            : MiniExcelRust.Query(stream, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration), leaveOpen);

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.QueryAsync(path, useHeaderRow, sheetName, startCell, cancellationToken: cancellationToken);

    public static IAsyncEnumerable<T> QueryAsync<T>(
        string path,
        string? sheetName = null,
        string startCell = "A1",
        bool treatHeaderAsData = false,
        CancellationToken cancellationToken = default)
        where T : class, new() =>
        MiniExcelRust.QueryAsync<T>(path, sheetName, startCell, treatHeaderAsData, cancellationToken: cancellationToken);

    public static IEnumerable<IDictionary<string, object?>> QueryRange(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        string? endCell = null,
        IConfiguration? configuration = null) =>
        MiniExcelRust.QueryRange(path, useHeaderRow, sheetName, startCell, endCell, OpenXmlOptions(configuration));

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        string? endCell = null,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.QueryRangeAsync(path, useHeaderRow, sheetName, startCell, endCell, cancellationToken: cancellationToken);

    public static int SaveAs(
        string path,
        IEnumerable<IDictionary<string, object?>> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool overwriteFile = false) =>
        MiniExcelRust.SaveAs(path, rows, printHeader, sheetName, overwriteFile);

    public static int SaveAs<T>(
        string path,
        IEnumerable<T> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool overwriteFile = false) =>
        MiniExcelRust.SaveAs(path, rows, printHeader, sheetName, overwriteFile);

    public static int SaveAs(
        Stream stream,
        IEnumerable<IDictionary<string, object?>> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool leaveOpen = false) =>
        MiniExcelRust.SaveAs(stream, rows, printHeader, sheetName, leaveOpen);

    public static Task<int> SaveAsAsync<T>(
        string path,
        IAsyncEnumerable<T> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool overwriteFile = false,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.SaveAsAsync(
            path,
            rows,
            printHeader,
            sheetName,
            overwriteFile,
            cancellationToken: cancellationToken);

    public static int Insert(
        string path,
        IEnumerable<IDictionary<string, object?>> rows,
        string sheetName,
        MiniExcelRustInsertOptions? options = null) =>
        MiniExcelRust.InsertSheet(path, rows, sheetName, options);

    public static Task<int> InsertAsync(
        string path,
        IEnumerable<IDictionary<string, object?>> rows,
        string sheetName = "Sheet1",
        MiniExcelRustInsertOptions? options = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.InsertSheet(path, rows, sheetName, options), cancellationToken);

    public static void SaveAsByTemplate(
        string path,
        string templatePath,
        object value,
        bool overwriteFile = false) =>
        MiniExcelRust.FillTemplate(path, templatePath, value, overwriteFile);

    public static Task SaveAsByTemplateAsync(
        string path,
        string templatePath,
        object value,
        bool overwriteFile = false,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.FillTemplate(path, templatePath, value, overwriteFile), cancellationToken);

    public static void SaveAsByTemplate(
        Stream stream,
        byte[] templateBytes,
        object value,
        bool leaveOpen = false) =>
        MiniExcelRust.FillTemplate(stream, templateBytes, value, leaveOpen: leaveOpen);

    public static void MergeSameCells(
        string destinationPath,
        string sourcePath,
        bool overwriteFile = false) =>
        MiniExcelRust.MergeSameCells(destinationPath, sourcePath, overwriteFile);

    public static Task MergeSameCellsAsync(
        string destinationPath,
        string sourcePath,
        bool overwriteFile = false,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => MiniExcelRust.MergeSameCells(destinationPath, sourcePath, overwriteFile),
            cancellationToken);

    public static IDataReader GetReader(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null) =>
        IsCsv(path, excelType)
            ? MiniExcelRust.GetCsvReader(path, useHeaderRow, CsvOptions(configuration))
            : MiniExcelRust.GetReader(path, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration));

    public static IDataReader GetReader(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        ExcelType excelType = ExcelType.XLSX,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool leaveOpen = false) =>
        excelType == ExcelType.CSV
            ? MiniExcelRust.GetCsvReader(stream, useHeaderRow, CsvOptions(configuration), leaveOpen)
            : MiniExcelRust.GetReader(stream, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration), leaveOpen);

    public static DataTable QueryAsDataTable(
        string path,
        bool useHeaderRow = true,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null) =>
        IsCsv(path, excelType)
            ? MiniExcelRust.QueryCsvAsDataTable(path, useHeaderRow, CsvOptions(configuration))
            : MiniExcelRust.QueryAsDataTable(path, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration));

    public static DataTable QueryAsDataTable(
        Stream stream,
        bool useHeaderRow = true,
        string? sheetName = null,
        ExcelType excelType = ExcelType.XLSX,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool leaveOpen = false) =>
        excelType == ExcelType.CSV
            ? MiniExcelRust.QueryCsvAsDataTable(stream, useHeaderRow, CsvOptions(configuration), leaveOpen)
            : MiniExcelRust.QueryAsDataTable(stream, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration), leaveOpen);

    public static Task<DataTable> QueryAsDataTableAsync(
        string path,
        bool useHeaderRow = true,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => QueryAsDataTable(path, useHeaderRow, sheetName, startCell, excelType, configuration),
            cancellationToken);

    public static List<string> GetSheetNames(string path) => MiniExcelRust.GetSheetNames(path);

    public static List<string> GetSheetNames(Stream stream, bool leaveOpen = false) =>
        MiniExcelRust.GetSheetNames(stream, leaveOpen);

    public static Task<List<string>> GetSheetNamesAsync(
        string path,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.GetSheetNamesAsync(path, cancellationToken);

    public static Task<List<string>> GetSheetNamesAsync(
        Stream stream,
        bool leaveOpen = false,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.GetSheetNames(stream, leaveOpen), cancellationToken);

    public static List<MiniExcelRustSheetInfo> GetSheetInformations(string path) =>
        MiniExcelRust.GetSheetInformations(path);

    public static List<MiniExcelRustSheetInfo> GetSheetInformations(Stream stream, bool leaveOpen = false) =>
        MiniExcelRust.GetSheetInformations(stream, leaveOpen);

    public static Task<List<MiniExcelRustSheetInfo>> GetSheetInformationsAsync(
        string path,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.GetSheetInformationsAsync(path, cancellationToken);

    public static List<MiniExcelRustRange> GetSheetDimensions(string path) =>
        MiniExcelRust.GetSheetDimensions(path);

    public static List<MiniExcelRustRange> GetSheetDimensions(Stream stream, bool leaveOpen = false) =>
        MiniExcelRust.GetSheetDimensions(stream, leaveOpen);

    public static Task<List<MiniExcelRustRange>> GetSheetDimensionsAsync(
        string path,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.GetSheetDimensionsAsync(path, cancellationToken);

    public static List<string> GetColumns(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null) =>
        IsCsv(path, excelType)
            ? MiniExcelRust.GetCsvColumnNames(path, useHeaderRow, CsvOptions(configuration))
            : MiniExcelRust.GetColumnNames(path, useHeaderRow, sheetName, startCell);

    public static Task<ICollection<string>> GetColumnsAsync(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null,
        CancellationToken cancellationToken = default) =>
        Task.Run<ICollection<string>>(
            () => GetColumns(path, useHeaderRow, sheetName, startCell, excelType, configuration),
            cancellationToken);

    public static void ConvertCsvToXlsx(string csvPath, string xlsxPath, bool csvHasHeader = false) =>
        MiniExcelRust.ConvertCsvToXlsx(csvPath, xlsxPath, csvHasHeader);

    public static void ConvertCsvToXlsx(Stream csvStream, Stream xlsxStream, bool csvHasHeader = false) =>
        MiniExcelRust.ConvertCsvToXlsx(csvStream, xlsxStream, csvHasHeader);

    public static Task ConvertCsvToXlsxAsync(
        string csvPath,
        string xlsxPath,
        bool csvHasHeader = false,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.ConvertCsvToXlsxAsync(csvPath, xlsxPath, csvHasHeader, cancellationToken);

    public static void ConvertXlsxToCsv(string xlsxPath, string csvPath, bool xlsxHasHeader = true) =>
        MiniExcelRust.ConvertXlsxToCsv(xlsxPath, csvPath, xlsxHasHeader);

    public static void ConvertXlsxToCsv(Stream xlsxStream, Stream csvStream, bool xlsxHasHeader = true) =>
        MiniExcelRust.ConvertXlsxToCsv(xlsxStream, csvStream, xlsxHasHeader);

    public static Task ConvertXlsxToCsvAsync(
        string xlsxPath,
        string csvPath,
        bool xlsxHasHeader = true,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.ConvertXlsxToCsvAsync(xlsxPath, csvPath, xlsxHasHeader, cancellationToken);

    private static bool IsCsv(string path, ExcelType excelType) =>
        excelType == ExcelType.CSV ||
        excelType == ExcelType.UNKNOWN && string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);

    private static MiniExcelRustReadOptions? OpenXmlOptions(IConfiguration? configuration) =>
        (configuration as OpenXml.OpenXmlConfiguration)?.ToReadOptions();

    private static MiniExcelRustCsvReadOptions? CsvOptions(IConfiguration? configuration) =>
        (configuration as Csv.CsvConfiguration)?.ToReadOptions();
}