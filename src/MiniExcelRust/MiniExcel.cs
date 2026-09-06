using System.Data;
using System.Collections;
using System.Runtime.CompilerServices;

namespace MiniExcelLibs;

/// <summary>
/// Compatibility facade whose workbook operations are backed exclusively by MiniExcel Rust.
/// </summary>
public static class MiniExcel
{
    public static void AddPicture(string path, params MiniExcelRustPicture[] pictures) =>
        MiniExcelRust.AddPicture(path, pictures);

    public static void AddPicture(
        Stream stream,
        bool leaveOpen = false,
        params MiniExcelRustPicture[] pictures) =>
        MiniExcelRust.AddPicture(stream, leaveOpen, pictures);

    public static Task AddPictureAsync(
        string path,
        CancellationToken cancellationToken = default,
        params MiniExcelRustPicture[] pictures) =>
        Task.Run(() => MiniExcelRust.AddPicture(path, pictures), cancellationToken);

    public static Task AddPictureAsync(
        Stream stream,
        CancellationToken cancellationToken = default,
        params MiniExcelRustPicture[] pictures) =>
        Task.Run(() => MiniExcelRust.AddPicture(stream, leaveOpen: true, pictures), cancellationToken);

    public static IEnumerable<IDictionary<string, object?>> Query(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null) =>
        IsCsv(path, excelType)
            ? MiniExcelRust.QueryCsv(path, useHeaderRow, CsvOptions(configuration))
            : MiniExcelRust.Query(path, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration));

    public static IEnumerable<T> Query<T>(
        string path,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool hasHeader = true)
        where T : class, new() =>
        IsCsv(path, excelType)
            ? MiniExcelRust.QueryCsv<T>(path, !hasHeader, CsvOptions(configuration))
            : MiniExcelRust.Query<T>(path, sheetName, startCell, !hasHeader, OpenXmlOptions(configuration));

    public static IEnumerable<IDictionary<string, object?>> Query(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null) =>
        excelType == ExcelType.CSV
            ? MiniExcelRust.QueryCsv(stream, useHeaderRow, CsvOptions(configuration), leaveOpen: true)
            : MiniExcelRust.Query(stream, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration), leaveOpen: true);

    public static IEnumerable<T> Query<T>(
        Stream stream,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool hasHeader = true)
        where T : class, new() =>
        excelType == ExcelType.CSV
            ? MiniExcelRust.QueryCsv<T>(stream, !hasHeader, CsvOptions(configuration), leaveOpen: true)
            : MiniExcelRust.Query<T>(
                stream,
                sheetName,
                startCell,
                !hasHeader,
                OpenXmlOptions(configuration),
                leaveOpen: true);

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null,
        CancellationToken cancellationToken = default) =>
        IsCsv(path, excelType)
            ? MiniExcelRust.QueryCsvAsync(path, useHeaderRow, CsvOptions(configuration), cancellationToken)
            : MiniExcelRust.QueryAsync(path, useHeaderRow, sheetName, startCell, OpenXmlOptions(configuration), cancellationToken);

    public static IAsyncEnumerable<T> QueryAsync<T>(
        string path,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool hasHeader = true,
        CancellationToken cancellationToken = default)
        where T : class, new() =>
        excelType == ExcelType.CSV
            ? MiniExcelRust.QueryCsvAsync<T>(path, !hasHeader, CsvOptions(configuration), cancellationToken)
            : MiniExcelRust.QueryAsync<T>(path, sheetName, startCell, !hasHeader, OpenXmlOptions(configuration), cancellationToken);

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null,
        CancellationToken cancellationToken = default) =>
        excelType == ExcelType.CSV
            ? ToFacadeAsync(MiniExcelRust.QueryCsv(stream, useHeaderRow, CsvOptions(configuration), leaveOpen: true), cancellationToken)
            : MiniExcelRust.QueryAsync(
                stream,
                useHeaderRow,
                sheetName,
                startCell,
                OpenXmlOptions(configuration),
                leaveOpen: true,
                cancellationToken: cancellationToken);

    public static IAsyncEnumerable<T> QueryAsync<T>(
        Stream stream,
        string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool hasHeader = true,
        CancellationToken cancellationToken = default)
        where T : class, new() =>
        excelType == ExcelType.CSV
            ? ToFacadeAsync(MiniExcelRust.QueryCsv<T>(stream, !hasHeader, CsvOptions(configuration), leaveOpen: true), cancellationToken)
            : MiniExcelRust.QueryAsync<T>(
                stream,
                sheetName,
                startCell,
                !hasHeader,
                OpenXmlOptions(configuration),
                leaveOpen: true,
                cancellationToken: cancellationToken);

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

    public static IEnumerable<IDictionary<string, object?>> QueryRange(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        string? endCell = null,
        IConfiguration? configuration = null,
        bool leaveOpen = true) =>
        MiniExcelRust.QueryRange(
            stream,
            useHeaderRow,
            sheetName,
            startCell,
            endCell,
            OpenXmlOptions(configuration),
            leaveOpen);

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        string? endCell = null,
        IConfiguration? configuration = null,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.QueryRangeAsync(
            stream,
            useHeaderRow,
            sheetName,
            startCell,
            endCell,
            OpenXmlOptions(configuration),
            leaveOpen,
            cancellationToken);

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(
        Stream stream,
        bool useHeaderRow,
        string? sheetName,
        int startRowIndex,
        int startColumnIndex,
        int? endRowIndex = null,
        int? endColumnIndex = null,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.QueryRangeAsync(
            stream,
            useHeaderRow,
            sheetName,
            startRowIndex,
            startColumnIndex,
            endRowIndex,
            endColumnIndex,
            leaveOpen: leaveOpen,
            cancellationToken: cancellationToken);

    public static IEnumerable<IDictionary<string, object?>> QueryRange(
        string path,
        bool useHeaderRow,
        string? sheetName,
        int startRowIndex,
        int startColumnIndex,
        int? endRowIndex = null,
        int? endColumnIndex = null,
        IConfiguration? configuration = null) =>
        MiniExcelRust.QueryRange(
            path,
            useHeaderRow,
            sheetName,
            startRowIndex,
            startColumnIndex,
            endRowIndex,
            endColumnIndex,
            OpenXmlOptions(configuration));

    public static IEnumerable<IDictionary<string, object?>> QueryRange(
        Stream stream,
        bool useHeaderRow,
        string? sheetName,
        int startRowIndex,
        int startColumnIndex,
        int? endRowIndex = null,
        int? endColumnIndex = null,
        IConfiguration? configuration = null) =>
        MiniExcelRust.QueryRange(
            stream,
            useHeaderRow,
            sheetName,
            startRowIndex,
            startColumnIndex,
            endRowIndex,
            endColumnIndex,
            OpenXmlOptions(configuration),
            leaveOpen: true);

    public static IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(
        string path,
        bool useHeaderRow,
        string? sheetName,
        int startRowIndex,
        int startColumnIndex,
        int? endRowIndex = null,
        int? endColumnIndex = null,
        IConfiguration? configuration = null,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.QueryRangeAsync(
            path,
            useHeaderRow,
            sheetName,
            startRowIndex,
            startColumnIndex,
            endRowIndex,
            endColumnIndex,
            OpenXmlOptions(configuration),
            cancellationToken);

    public static int[] SaveAs(
        string path,
        IEnumerable<IDictionary<string, object?>> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool overwriteFile = false) =>
        [MiniExcelRust.SaveAs(path, rows, printHeader, sheetName, overwriteFile)];

    public static int[] SaveAs<T>(
        string path,
        IEnumerable<T> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool overwriteFile = false) =>
        [MiniExcelRust.SaveAs(path, rows, printHeader, sheetName, overwriteFile)];

    public static int[] SaveAs(
        string path,
        object value,
        bool printHeader = true,
        string sheetName = "Sheet1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null,
        bool overwriteFile = false)
    {
        if (value is DataSet dataSet)
        {
            var sheets = dataSet.Tables.Cast<DataTable>().Select(table =>
                new KeyValuePair<string, IEnumerable<IDictionary<string, object?>>>(
                    string.IsNullOrWhiteSpace(table.TableName) ? sheetName : table.TableName,
                    DataTableRows(table)));
            return MiniExcelRust.SaveAsSheets(path, sheets, printHeader, overwriteFile);
        }
        if (excelType == ExcelType.CSV || IsCsv(path, excelType))
            return [MiniExcelRust.SaveAsCsv(path, ObjectRows(value))];
        return [MiniExcelRust.SaveAs(path, ObjectRows(value), printHeader, sheetName, overwriteFile)];
    }

    public static int[] SaveAs(
        Stream stream,
        IEnumerable<IDictionary<string, object?>> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool leaveOpen = false) =>
        [MiniExcelRust.SaveAs(stream, rows, printHeader, sheetName, leaveOpen)];

    public static int[] SaveAs(
        Stream stream,
        object value,
        bool printHeader = true,
        string sheetName = "Sheet1",
        ExcelType excelType = ExcelType.XLSX,
        IConfiguration? configuration = null)
    {
        if (value is DataSet dataSet)
        {
            var sheets = dataSet.Tables.Cast<DataTable>().Select(table =>
                new KeyValuePair<string, IEnumerable<IDictionary<string, object?>>>(
                    string.IsNullOrWhiteSpace(table.TableName) ? sheetName : table.TableName,
                    DataTableRows(table)));
            return MiniExcelRust.SaveAsSheets(stream, sheets, printHeader, leaveOpen: true);
        }
        if (excelType == ExcelType.CSV)
            return [MiniExcelRust.SaveAsCsv(stream, ObjectRows(value), leaveOpen: true)];
        return [MiniExcelRust.SaveAs(stream, ObjectRows(value), printHeader, sheetName, leaveOpen: true)];
    }

    public static Task<int[]> SaveAsAsync(
        string path,
        object value,
        bool printHeader = true,
        string sheetName = "Sheet1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null,
        bool overwriteFile = false,
        IProgress<int>? progress = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => SaveAs(path, value, printHeader, sheetName, excelType, configuration, overwriteFile),
            cancellationToken);

    public static Task<int[]> SaveAsAsync(
        Stream stream,
        object value,
        bool printHeader = true,
        string sheetName = "Sheet1",
        ExcelType excelType = ExcelType.XLSX,
        IConfiguration? configuration = null,
        IProgress<int>? progress = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => SaveAs(stream, value, printHeader, sheetName, excelType, configuration), cancellationToken);

    public static async Task<int[]> SaveAsAsync<T>(
        string path,
        IAsyncEnumerable<T> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool overwriteFile = false,
        CancellationToken cancellationToken = default) =>
        [await MiniExcelRust.SaveAsAsync(
            path,
            rows,
            printHeader,
            sheetName,
            overwriteFile,
            cancellationToken: cancellationToken).ConfigureAwait(false)];

    public static async Task<int[]> SaveAsAsync<T>(
        Stream stream,
        IAsyncEnumerable<T> rows,
        bool printHeader = true,
        string sheetName = "Sheet1",
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        [await MiniExcelRust.SaveAsAsync(
            stream,
            rows,
            printHeader,
            sheetName,
            leaveOpen,
            cancellationToken: cancellationToken).ConfigureAwait(false)];

    public static int Insert(
        string path,
        IEnumerable<IDictionary<string, object?>> rows,
        string sheetName,
        MiniExcelRustInsertOptions? options = null) =>
        MiniExcelRust.InsertSheet(path, rows, sheetName, options);

    public static int Insert(
        string path,
        object value,
        string sheetName = "Sheet1",
        ExcelType excelType = ExcelType.UNKNOWN,
        IConfiguration? configuration = null,
        bool printHeader = true,
        bool overwriteSheet = false) =>
        MiniExcelRust.InsertSheet(
            path,
            ObjectRows(value),
            sheetName,
            new MiniExcelRustInsertOptions
            {
                PrintHeader = printHeader,
                ReplaceExistingSheet = overwriteSheet
            });

    public static Task<int> InsertAsync(
        string path,
        IEnumerable<IDictionary<string, object?>> rows,
        string sheetName = "Sheet1",
        MiniExcelRustInsertOptions? options = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.InsertSheet(path, rows, sheetName, options), cancellationToken);

    public static int Insert(
        Stream stream,
        IEnumerable<IDictionary<string, object?>> rows,
        string sheetName = "Sheet1",
        MiniExcelRustInsertOptions? options = null,
        bool leaveOpen = true) =>
        MiniExcelRust.InsertSheet(stream, rows, sheetName, options, leaveOpen);

    public static Task<int> InsertAsync(
        Stream stream,
        IEnumerable<IDictionary<string, object?>> rows,
        string sheetName = "Sheet1",
        MiniExcelRustInsertOptions? options = null,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => MiniExcelRust.InsertSheet(stream, rows, sheetName, options, leaveOpen),
            cancellationToken);

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

    public static void SaveAsByTemplate(
        string path,
        byte[] templateBytes,
        object value,
        bool overwriteFile = false) =>
        MiniExcelRust.FillTemplate(path, templateBytes, value, overwriteFile);

    public static void SaveAsByTemplate(
        string path,
        Stream templateStream,
        object value,
        bool overwriteFile = false,
        bool leaveTemplateOpen = true) =>
        MiniExcelRust.FillTemplate(path, templateStream, value, overwriteFile, leaveTemplateOpen: leaveTemplateOpen);

    public static void SaveAsByTemplate(
        Stream stream,
        string templatePath,
        object value,
        bool leaveOpen = true) =>
        MiniExcelRust.FillTemplate(stream, templatePath, value, leaveOpen: leaveOpen);

    public static void SaveAsByTemplate(
        Stream stream,
        Stream templateStream,
        object value,
        bool leaveOpen = true,
        bool leaveTemplateOpen = true) =>
        MiniExcelRust.FillTemplate(
            stream,
            templateStream,
            value,
            leaveOpen: leaveOpen,
            leaveTemplateOpen: leaveTemplateOpen);

    public static Task SaveAsByTemplateAsync(
        string path,
        byte[] templateBytes,
        object value,
        bool overwriteFile = false,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.FillTemplate(path, templateBytes, value, overwriteFile), cancellationToken);

    public static Task SaveAsByTemplateAsync(
        string path,
        Stream templateStream,
        object value,
        bool overwriteFile = false,
        bool leaveTemplateOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => MiniExcelRust.FillTemplate(
                path,
                templateStream,
                value,
                overwriteFile,
                leaveTemplateOpen: leaveTemplateOpen),
            cancellationToken);

    public static Task SaveAsByTemplateAsync(
        Stream stream,
        string templatePath,
        object value,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.FillTemplate(stream, templatePath, value, leaveOpen: leaveOpen), cancellationToken);

    public static Task SaveAsByTemplateAsync(
        Stream stream,
        byte[] templateBytes,
        object value,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.FillTemplate(stream, templateBytes, value, leaveOpen: leaveOpen), cancellationToken);

    public static Task SaveAsByTemplateAsync(
        Stream stream,
        Stream templateStream,
        object value,
        bool leaveOpen = true,
        bool leaveTemplateOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => MiniExcelRust.FillTemplate(
                stream,
                templateStream,
                value,
                leaveOpen: leaveOpen,
                leaveTemplateOpen: leaveTemplateOpen),
            cancellationToken);

    public static void MergeSameCells(
        string destinationPath,
        string sourcePath,
        bool overwriteFile = false) =>
        MiniExcelRust.MergeSameCells(destinationPath, sourcePath, overwriteFile);

    public static void MergeSameCells(Stream stream, string sourcePath) =>
        MiniExcelRust.MergeSameCells(stream, sourcePath, leaveOpen: true);

    public static void MergeSameCells(Stream stream, byte[] sourceBytes) =>
        MiniExcelRust.MergeSameCells(stream, sourceBytes, leaveOpen: true);

    public static Task MergeSameCellsAsync(
        string destinationPath,
        string sourcePath,
        bool overwriteFile = false,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => MiniExcelRust.MergeSameCells(destinationPath, sourcePath, overwriteFile),
            cancellationToken);

    public static Task MergeSameCellsAsync(
        Stream stream,
        string sourcePath,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.MergeSameCells(stream, sourcePath, leaveOpen), cancellationToken);

    public static Task MergeSameCellsAsync(
        Stream stream,
        byte[] sourceBytes,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.MergeSameCells(stream, sourceBytes, leaveOpen), cancellationToken);

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

    public static Task<DataTable> QueryAsDataTableAsync(
        Stream stream,
        bool useHeaderRow = true,
        string? sheetName = null,
        ExcelType excelType = ExcelType.XLSX,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(
            () => QueryAsDataTable(
                stream,
                useHeaderRow,
                sheetName,
                excelType,
                startCell,
                configuration,
                leaveOpen),
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

    public static Task<List<MiniExcelRustSheetInfo>> GetSheetInformationsAsync(
        Stream stream,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.GetSheetInformations(stream, leaveOpen), cancellationToken);

    public static List<MiniExcelRustRange> GetSheetDimensions(string path) =>
        MiniExcelRust.GetSheetDimensions(path);

    public static List<MiniExcelRustRange> GetSheetDimensions(Stream stream, bool leaveOpen = false) =>
        MiniExcelRust.GetSheetDimensions(stream, leaveOpen);

    public static Task<List<MiniExcelRustRange>> GetSheetDimensionsAsync(
        string path,
        CancellationToken cancellationToken = default) =>
        MiniExcelRust.GetSheetDimensionsAsync(path, cancellationToken);

    public static Task<List<MiniExcelRustRange>> GetSheetDimensionsAsync(
        Stream stream,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.GetSheetDimensions(stream, leaveOpen), cancellationToken);

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

    public static List<string> GetColumns(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool leaveOpen = true) =>
        MiniExcelRust.GetColumnNames(
            stream,
            useHeaderRow,
            sheetName,
            startCell,
            leaveOpen);

    public static Task<ICollection<string>> GetColumnsAsync(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        IConfiguration? configuration = null,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default) =>
        Task.Run<ICollection<string>>(
            () => GetColumns(stream, useHeaderRow, sheetName, startCell, configuration, leaveOpen),
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

    public static Task ConvertCsvToXlsxAsync(
        Stream csvStream,
        Stream xlsxStream,
        bool csvHasHeader = false,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.ConvertCsvToXlsx(csvStream, xlsxStream, csvHasHeader), cancellationToken);

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

    public static Task ConvertXlsxToCsvAsync(
        Stream xlsxStream,
        Stream csvStream,
        bool xlsxHasHeader = true,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => MiniExcelRust.ConvertXlsxToCsv(xlsxStream, csvStream, xlsxHasHeader), cancellationToken);

    private static bool IsCsv(string path, ExcelType excelType) =>
        excelType == ExcelType.CSV ||
        excelType == ExcelType.UNKNOWN && string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);

    private static MiniExcelRustReadOptions? OpenXmlOptions(IConfiguration? configuration) =>
        (configuration as OpenXml.OpenXmlConfiguration)?.ToReadOptions();

    private static MiniExcelRustCsvReadOptions? CsvOptions(IConfiguration? configuration) =>
        (configuration as Csv.CsvConfiguration)?.ToReadOptions();

    private static IEnumerable<IDictionary<string, object?>> ObjectRows(object value)
    {
        if (value is null)
            throw new ArgumentNullException(nameof(value));
        if (value is DataTable table)
            return DataTableRows(table);
        if (value is IDictionary<string, object?> row)
            return new[] { row };
        if (value is IEnumerable values && value is not string)
            return MiniExcelRustMapper.ToRows(values);
        return MiniExcelRustMapper.ToRows(new[] { value });
    }

    private static IEnumerable<IDictionary<string, object?>> DataTableRows(DataTable table)
    {
        var columns = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
        foreach (DataRow dataRow in table.Rows)
        {
            IDictionary<string, object?> row = new Dictionary<string, object?>(StringComparer.Ordinal);
            foreach (var column in columns)
                row[column] = dataRow[column] is DBNull ? null : dataRow[column];
            yield return row;
        }
    }

    private static async IAsyncEnumerable<T> ToFacadeAsync<T>(
        IEnumerable<T> values,
        [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        foreach (var value in values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return value;
            await Task.Yield();
        }
    }
}