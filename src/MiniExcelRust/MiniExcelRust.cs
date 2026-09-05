using System.Collections;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace MiniExcelLibs;

/// <summary>
/// Provides XLSX queries backed by the native MiniExcel Rust library.
/// </summary>
public static class MiniExcelRust
{
    private const int BatchSize = 64;

    /// <summary>
    /// Returns worksheet names in workbook order.
    /// </summary>
    public static List<string> GetSheetNames(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));

        EnsureAbiVersion();

        using var nativePath = new Utf8String(Path.GetFullPath(path));
        var result = NativeMethods.GetSheetNames(
            nativePath.Pointer,
            out var rawHandle,
            out var data,
            out var length);
        if (result < 0)
            throw CreateNativeException(result);

        using var handle = new NativeBufferHandle(rawHandle);
        var byteLength = checked((int)length.ToUInt64());
        var frame = new byte[byteLength];
        Marshal.Copy(data, frame, 0, byteLength);
        return DecodeStrings(frame);
    }

    /// <summary>
    /// Returns worksheet names from a stream in workbook order.
    /// </summary>
    public static List<string> GetSheetNames(Stream stream, bool leaveOpen = false)
    {
        return UseStagedStream(stream, leaveOpen, GetSheetNames);
    }

    /// <summary>
    /// Asynchronously returns worksheet names in workbook order.
    /// </summary>
    public static Task<List<string>> GetSheetNamesAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetSheetNames(path), cancellationToken);
    }

    /// <summary>
    /// Returns selected column names from an XLSX worksheet.
    /// </summary>
    public static List<string> GetColumnNames(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1")
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        if (string.IsNullOrWhiteSpace(startCell))
            throw new ArgumentException("The start cell is required.", nameof(startCell));

        EnsureAbiVersion();

        using var nativePath = new Utf8String(Path.GetFullPath(path));
        using var nativeSheetName = new Utf8String(sheetName);
        using var nativeStartCell = new Utf8String(startCell);
        var result = NativeMethods.GetColumns(
            nativePath.Pointer,
            useHeaderRow ? (byte)1 : (byte)0,
            nativeSheetName.Pointer,
            nativeStartCell.Pointer,
            out var rawHandle,
            out var data,
            out var length);
        if (result < 0)
            throw CreateNativeException(result);

        using var handle = new NativeBufferHandle(rawHandle);
        var byteLength = checked((int)length.ToUInt64());
        var frame = new byte[byteLength];
        Marshal.Copy(data, frame, 0, byteLength);
        return DecodeStrings(frame);
    }

    /// <summary>
    /// Returns selected column names from an XLSX stream.
    /// </summary>
    public static List<string> GetColumnNames(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        bool leaveOpen = false)
    {
        return UseStagedStream(
            stream,
            leaveOpen,
            path => GetColumnNames(path, useHeaderRow, sheetName, startCell));
    }

    /// <summary>
    /// Asynchronously returns selected column names from an XLSX worksheet.
    /// </summary>
    public static Task<List<string>> GetColumnNamesAsync(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        CancellationToken cancellationToken = default)
    {
        return Task.Run(
            () => GetColumnNames(path, useHeaderRow, sheetName, startCell),
            cancellationToken);
    }

    /// <summary>
    /// Returns the used range of every worksheet in workbook order.
    /// </summary>
    public static List<MiniExcelRustRange> GetSheetDimensions(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));

        EnsureAbiVersion();

        using var nativePath = new Utf8String(Path.GetFullPath(path));
        var result = NativeMethods.GetSheetDimensions(
            nativePath.Pointer,
            out var rawHandle,
            out var data,
            out var length);
        if (result < 0)
            throw CreateNativeException(result);

        using var handle = new NativeBufferHandle(rawHandle);
        var byteLength = checked((int)length.ToUInt64());
        var frame = new byte[byteLength];
        Marshal.Copy(data, frame, 0, byteLength);
        return DecodeRanges(frame);
    }

    /// <summary>
    /// Returns the used range of every worksheet in an XLSX stream.
    /// </summary>
    public static List<MiniExcelRustRange> GetSheetDimensions(Stream stream, bool leaveOpen = false)
    {
        return UseStagedStream(stream, leaveOpen, GetSheetDimensions);
    }

    /// <summary>
    /// Asynchronously returns the used range of every worksheet in workbook order.
    /// </summary>
    public static Task<List<MiniExcelRustRange>> GetSheetDimensionsAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetSheetDimensions(path), cancellationToken);
    }

    /// <summary>
    /// Returns detailed information for every sheet in workbook order.
    /// </summary>
    public static List<MiniExcelRustSheetInfo> GetSheetInformations(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));

        EnsureAbiVersion();

        using var nativePath = new Utf8String(Path.GetFullPath(path));
        var result = NativeMethods.GetSheetInfo(
            nativePath.Pointer,
            out var rawHandle,
            out var data,
            out var length);
        if (result < 0)
            throw CreateNativeException(result);

        using var handle = new NativeBufferHandle(rawHandle);
        var byteLength = checked((int)length.ToUInt64());
        var frame = new byte[byteLength];
        Marshal.Copy(data, frame, 0, byteLength);
        return DecodeSheetInfo(frame);
    }

    /// <summary>
    /// Returns detailed information for every sheet in an XLSX stream.
    /// </summary>
    public static List<MiniExcelRustSheetInfo> GetSheetInformations(
        Stream stream,
        bool leaveOpen = false)
    {
        return UseStagedStream(stream, leaveOpen, GetSheetInformations);
    }

    /// <summary>
    /// Asynchronously returns detailed information for every sheet in workbook order.
    /// </summary>
    public static Task<List<MiniExcelRustSheetInfo>> GetSheetInformationsAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetSheetInformations(path), cancellationToken);
    }

    /// <summary>
    /// Materializes an XLSX query as a DataTable.
    /// </summary>
    public static DataTable QueryAsDataTable(
        string path,
        bool hasHeaderRow = true,
        string? sheetName = null,
        string startCell = "A1",
        MiniExcelRustReadOptions? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));

        var fullPath = Path.GetFullPath(path);
        var columns = GetColumnNames(fullPath, hasHeaderRow, sheetName, startCell);
        var rows = Query(fullPath, hasHeaderRow, sheetName, startCell, configuration);
        return CreateDataTable(columns, rows);
    }

    /// <summary>
    /// Materializes an XLSX stream query as a DataTable.
    /// </summary>
    public static DataTable QueryAsDataTable(
        Stream stream,
        bool hasHeaderRow = true,
        string? sheetName = null,
        string startCell = "A1",
        MiniExcelRustReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        return UseStagedStream(
            stream,
            leaveOpen,
            path => QueryAsDataTable(path, hasHeaderRow, sheetName, startCell, configuration));
    }

    /// <summary>
    /// Returns a DataReader over a materialized Rust-backed XLSX query.
    /// </summary>
    public static IDataReader GetReader(
        string path,
        bool hasHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        MiniExcelRustReadOptions? configuration = null)
    {
        return QueryAsDataTable(path, hasHeaderRow, sheetName, startCell, configuration).CreateDataReader();
    }

    /// <summary>
    /// Returns a DataReader over a materialized Rust-backed XLSX stream query.
    /// </summary>
    public static IDataReader GetReader(
        Stream stream,
        bool hasHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        MiniExcelRustReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        var table = QueryAsDataTable(stream, hasHeaderRow, sheetName, startCell, configuration, leaveOpen);
        return table.CreateDataReader();
    }

    /// <summary>
    /// Streams rows from an XLSX file through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> Query(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        MiniExcelRustReadOptions? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        if (string.IsNullOrWhiteSpace(startCell))
            throw new ArgumentException("The start cell is required.", nameof(startCell));

        return QueryIterator(Path.GetFullPath(path), useHeaderRow, sheetName, startCell, null, configuration);
    }

    /// <summary>
    /// Streams rows from an XLSX stream through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> Query(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        MiniExcelRustReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        ValidateReadableStream(stream);
        if (string.IsNullOrWhiteSpace(startCell))
            throw new ArgumentException("The start cell is required.", nameof(startCell));

        return QueryStreamIterator(stream, useHeaderRow, sheetName, startCell, null, configuration, leaveOpen);
    }

    /// <summary>
    /// Streams rows from an inclusive XLSX cell range through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> QueryRange(
        string path,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        string? endCell = null,
        MiniExcelRustReadOptions? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        if (string.IsNullOrWhiteSpace(startCell))
            throw new ArgumentException("The start cell is required.", nameof(startCell));
        if (endCell is not null && string.IsNullOrWhiteSpace(endCell))
            throw new ArgumentException("The end cell cannot be empty.", nameof(endCell));

        return QueryIterator(Path.GetFullPath(path), useHeaderRow, sheetName, startCell, endCell, configuration);
    }

    /// <summary>
    /// Streams rows from an inclusive XLSX range in a stream through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> QueryRange(
        Stream stream,
        bool useHeaderRow = false,
        string? sheetName = null,
        string startCell = "A1",
        string? endCell = null,
        MiniExcelRustReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        ValidateReadableStream(stream);
        if (string.IsNullOrWhiteSpace(startCell))
            throw new ArgumentException("The start cell is required.", nameof(startCell));
        if (endCell is not null && string.IsNullOrWhiteSpace(endCell))
            throw new ArgumentException("The end cell cannot be empty.", nameof(endCell));

        return QueryStreamIterator(stream, useHeaderRow, sheetName, startCell, endCell, configuration, leaveOpen);
    }

    /// <summary>
    /// Streams rows from a named OpenXML table.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> QueryTable(
        string path,
        string? sheetName = null,
        string tableName = "Table1")
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        if (string.IsNullOrWhiteSpace(tableName))
            throw new ArgumentException("The table name is required.", nameof(tableName));

        return QueryTableIterator(Path.GetFullPath(path), sheetName, tableName);
    }

    /// <summary>
    /// Streams rows from a named OpenXML table in a stream.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> QueryTable(
        Stream stream,
        string? sheetName = null,
        string tableName = "Table1",
        bool leaveOpen = false)
    {
        ValidateReadableStream(stream);
        if (string.IsNullOrWhiteSpace(tableName))
            throw new ArgumentException("The table name is required.", nameof(tableName));

        return QueryTableStreamIterator(stream, sheetName, tableName, leaveOpen);
    }

    /// <summary>
    /// Streams rows from a CSV file through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> QueryCsv(
        string path,
        bool useHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        ValidateCsvConfiguration(configuration);

        return QueryCsvIterator(Path.GetFullPath(path), useHeaderRow, configuration);
    }

    /// <summary>
    /// Streams rows from a CSV stream through the native Rust query engine.
    /// </summary>
    public static IEnumerable<IDictionary<string, object?>> QueryCsv(
        Stream stream,
        bool useHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        ValidateReadableStream(stream);
        ValidateCsvConfiguration(configuration);

        return QueryCsvStreamIterator(stream, useHeaderRow, configuration, leaveOpen);
    }

    /// <summary>
    /// Returns selected column names from a CSV file.
    /// </summary>
    public static List<string> GetCsvColumnNames(
        string path,
        bool useHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));
        ValidateCsvConfiguration(configuration);
        EnsureAbiVersion();
        configuration ??= new MiniExcelRustCsvReadOptions();

        using var nativePath = new Utf8String(Path.GetFullPath(path));
        var result = NativeMethods.GetCsvColumns(
            nativePath.Pointer,
            useHeaderRow ? (byte)1 : (byte)0,
            (byte)configuration.Delimiter,
            (byte)configuration.Encoding,
            configuration.ReadEmptyStringAsNull ? (byte)1 : (byte)0,
            configuration.TrimColumnNames ? (byte)1 : (byte)0,
            out var rawHandle,
            out var data,
            out var length);
        if (result < 0)
            throw CreateNativeException(result);

        using var handle = new NativeBufferHandle(rawHandle);
        var byteLength = checked((int)length.ToUInt64());
        var frame = new byte[byteLength];
        Marshal.Copy(data, frame, 0, byteLength);
        return DecodeStrings(frame);
    }

    /// <summary>
    /// Returns selected column names from a CSV stream.
    /// </summary>
    public static List<string> GetCsvColumnNames(
        Stream stream,
        bool useHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        return UseStagedStream(
            stream,
            leaveOpen,
            path => GetCsvColumnNames(path, useHeaderRow, configuration));
    }

    /// <summary>
    /// Asynchronously returns selected column names from a CSV file.
    /// </summary>
    public static Task<List<string>> GetCsvColumnNamesAsync(
        string path,
        bool useHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(
            () => GetCsvColumnNames(path, useHeaderRow, configuration),
            cancellationToken);
    }

    /// <summary>
    /// Materializes a CSV query as a DataTable.
    /// </summary>
    public static DataTable QueryCsvAsDataTable(
        string path,
        bool hasHeaderRow = true,
        MiniExcelRustCsvReadOptions? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path is required.", nameof(path));

        var fullPath = Path.GetFullPath(path);
        var columns = GetCsvColumnNames(fullPath, hasHeaderRow, configuration);
        var rows = QueryCsv(fullPath, hasHeaderRow, configuration);
        return CreateDataTable(columns, rows);
    }

    /// <summary>
    /// Materializes a CSV stream query as a DataTable.
    /// </summary>
    public static DataTable QueryCsvAsDataTable(
        Stream stream,
        bool hasHeaderRow = true,
        MiniExcelRustCsvReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        return UseStagedStream(
            stream,
            leaveOpen,
            path => QueryCsvAsDataTable(path, hasHeaderRow, configuration));
    }

    /// <summary>
    /// Returns a DataReader over a materialized Rust-backed CSV query.
    /// </summary>
    public static IDataReader GetCsvReader(
        string path,
        bool hasHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null)
    {
        return QueryCsvAsDataTable(path, hasHeaderRow, configuration).CreateDataReader();
    }

    /// <summary>
    /// Returns a DataReader over a materialized Rust-backed CSV stream query.
    /// </summary>
    public static IDataReader GetCsvReader(
        Stream stream,
        bool hasHeaderRow = false,
        MiniExcelRustCsvReadOptions? configuration = null,
        bool leaveOpen = false)
    {
        return QueryCsvAsDataTable(stream, hasHeaderRow, configuration, leaveOpen).CreateDataReader();
    }

    private static IEnumerable<IDictionary<string, object?>> QueryStreamIterator(
        Stream stream,
        bool useHeaderRow,
        string? sheetName,
        string startCell,
        string? endCell,
        MiniExcelRustReadOptions? configuration,
        bool leaveOpen)
    {
        string? temporaryPath = null;
        try
        {
            temporaryPath = StageStream(stream);
            foreach (var row in QueryIterator(temporaryPath, useHeaderRow, sheetName, startCell, endCell, configuration))
                yield return row;
        }
        finally
        {
            if (!leaveOpen)
                stream.Dispose();
            DeleteTemporaryFile(temporaryPath);
        }
    }

    private static IEnumerable<IDictionary<string, object?>> QueryTableStreamIterator(
        Stream stream,
        string? sheetName,
        string tableName,
        bool leaveOpen)
    {
        string? temporaryPath = null;
        try
        {
            temporaryPath = StageStream(stream);
            foreach (var row in QueryTableIterator(temporaryPath, sheetName, tableName))
                yield return row;
        }
        finally
        {
            if (!leaveOpen)
                stream.Dispose();
            DeleteTemporaryFile(temporaryPath);
        }
    }

    private static IEnumerable<IDictionary<string, object?>> QueryTableIterator(
        string path,
        string? sheetName,
        string tableName)
    {
        EnsureAbiVersion();

        using var nativePath = new Utf8String(path);
        using var nativeSheetName = new Utf8String(sheetName);
        using var nativeTableName = new Utf8String(tableName);
        var result = NativeMethods.QueryTableOpen(
            nativePath.Pointer,
            nativeSheetName.Pointer,
            nativeTableName.Pointer,
            out var rawHandle);
        if (result < 0)
            throw CreateNativeException(result);

        foreach (var row in ReadRows(rawHandle))
            yield return row;
    }

    private static IEnumerable<IDictionary<string, object?>> QueryCsvStreamIterator(
        Stream stream,
        bool useHeaderRow,
        MiniExcelRustCsvReadOptions? configuration,
        bool leaveOpen)
    {
        string? temporaryPath = null;
        try
        {
            temporaryPath = StageStream(stream);
            foreach (var row in QueryCsvIterator(temporaryPath, useHeaderRow, configuration))
                yield return row;
        }
        finally
        {
            if (!leaveOpen)
                stream.Dispose();
            DeleteTemporaryFile(temporaryPath);
        }
    }

    private static IEnumerable<IDictionary<string, object?>> QueryCsvIterator(
        string path,
        bool useHeaderRow,
        MiniExcelRustCsvReadOptions? configuration)
    {
        EnsureAbiVersion();
        configuration ??= new MiniExcelRustCsvReadOptions();

        using var nativePath = new Utf8String(path);
        var result = NativeMethods.QueryCsvOpen(
            nativePath.Pointer,
            useHeaderRow ? (byte)1 : (byte)0,
            (byte)configuration.Delimiter,
            (byte)configuration.Encoding,
            configuration.ReadEmptyStringAsNull ? (byte)1 : (byte)0,
            configuration.TrimColumnNames ? (byte)1 : (byte)0,
            out var rawHandle);
        if (result < 0)
            throw CreateNativeException(result);

        foreach (var row in ReadRows(rawHandle))
            yield return row;
    }

    private static IEnumerable<IDictionary<string, object?>> QueryIterator(
        string path,
        bool useHeaderRow,
        string? sheetName,
        string startCell,
        string? endCell,
        MiniExcelRustReadOptions? configuration)
    {
        EnsureAbiVersion();

        using var nativePath = new Utf8String(path);
        using var nativeSheetName = new Utf8String(sheetName);
        using var nativeStartCell = new Utf8String(startCell);
        using var nativeEndCell = new Utf8String(endCell);
        using var nativeCachePath = new Utf8String(configuration?.SharedStringCachePath);
        int result;
        IntPtr rawHandle;
        if (configuration is not null)
        {
            result = NativeMethods.QueryOptionsOpen(
                nativePath.Pointer,
                useHeaderRow ? (byte)1 : (byte)0,
                nativeSheetName.Pointer,
                nativeStartCell.Pointer,
                nativeEndCell.Pointer,
                configuration.IgnoreEmptyRows ? (byte)1 : (byte)0,
                configuration.FillMergedCells ? (byte)1 : (byte)0,
                configuration.TrimColumnNames ? (byte)1 : (byte)0,
                configuration.EnableSharedStringCache ? (byte)1 : (byte)0,
                configuration.SharedStringCacheSize,
                nativeCachePath.Pointer,
                out rawHandle);
        }
        else if (endCell is not null)
        {
            result = NativeMethods.QueryRangeOpen(
                nativePath.Pointer,
                useHeaderRow ? (byte)1 : (byte)0,
                nativeSheetName.Pointer,
                nativeStartCell.Pointer,
                nativeEndCell.Pointer,
                out rawHandle);
        }
        else
        {
            result = NativeMethods.QueryOpen(
                nativePath.Pointer,
                useHeaderRow ? (byte)1 : (byte)0,
                nativeSheetName.Pointer,
                nativeStartCell.Pointer,
                out rawHandle);
        }
        if (result < 0)
            throw CreateNativeException(result);

        foreach (var row in ReadRows(rawHandle))
            yield return row;
    }

    private static IEnumerable<IDictionary<string, object?>> ReadRows(IntPtr rawHandle)
    {
        using var handle = new NativeQueryHandle(rawHandle);
        while (true)
        {
            var result = NativeMethods.QueryNextBatch(handle, BatchSize, out var data, out var length);
            if (result == 0)
                yield break;
            if (result < 0)
                throw CreateNativeException(result);

            var byteLength = checked((int)length.ToUInt64());
            var frame = new byte[byteLength];
            Marshal.Copy(data, frame, 0, byteLength);
            foreach (var row in DecodeBatch(frame))
                yield return row;
        }
    }

    private static IEnumerable<IDictionary<string, object?>> DecodeBatch(byte[] frame)
    {
        var reader = new FrameReader(frame);
        var rowCount = reader.ReadLength();
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var cellCount = reader.ReadLength();
            IDictionary<string, object?> row = new Dictionary<string, object?>(cellCount, StringComparer.Ordinal);
            for (var cellIndex = 0; cellIndex < cellCount; cellIndex++)
                row.Add(reader.ReadString(), reader.ReadValue());
            yield return row;
        }

        reader.EnsureComplete();
    }

    private static List<string> DecodeStrings(byte[] frame)
    {
        var reader = new FrameReader(frame);
        var count = reader.ReadLength();
        var values = new List<string>(count);
        for (var index = 0; index < count; index++)
            values.Add(reader.ReadString());
        reader.EnsureComplete();
        return values;
    }

    private static List<MiniExcelRustRange> DecodeRanges(byte[] frame)
    {
        var reader = new FrameReader(frame);
        var count = reader.ReadLength();
        var ranges = new List<MiniExcelRustRange>(count);
        for (var index = 0; index < count; index++)
        {
            var startCell = reader.ReadString();
            var endCell = reader.ReadString();
            ranges.Add(new MiniExcelRustRange(
                startCell.Length == 0 ? null : startCell,
                endCell.Length == 0 ? null : endCell));
        }
        reader.EnsureComplete();
        return ranges;
    }

    private static List<MiniExcelRustSheetInfo> DecodeSheetInfo(byte[] frame)
    {
        var reader = new FrameReader(frame);
        var count = reader.ReadLength();
        var sheets = new List<MiniExcelRustSheetInfo>(count);
        for (var index = 0; index < count; index++)
        {
            sheets.Add(new MiniExcelRustSheetInfo(
                reader.ReadUInt32(),
                reader.ReadUInt32(),
                reader.ReadString(),
                (MiniExcelRustSheetType)reader.ReadByte(),
                (MiniExcelRustSheetState)reader.ReadByte(),
                reader.ReadByte() != 0));
        }
        reader.EnsureComplete();
        return sheets;
    }

    private static DataTable CreateDataTable(
        IReadOnlyList<string> columns,
        IEnumerable<IDictionary<string, object?>> rows)
    {
        var table = new DataTable();
        foreach (var column in columns)
            table.Columns.Add(column, typeof(object));

        foreach (var row in rows)
        {
            var values = new object?[columns.Count];
            for (var index = 0; index < columns.Count; index++)
                values[index] = row.TryGetValue(columns[index], out var value) ? value ?? DBNull.Value : DBNull.Value;
            table.Rows.Add(values);
        }

        return table;
    }

    private static TResult UseStagedStream<TResult>(
        Stream stream,
        bool leaveOpen,
        Func<string, TResult> operation)
    {
        ValidateReadableStream(stream);
        string? temporaryPath = null;
        try
        {
            temporaryPath = StageStream(stream);
            return operation(temporaryPath);
        }
        finally
        {
            if (!leaveOpen)
                stream.Dispose();
            DeleteTemporaryFile(temporaryPath);
        }
    }

    private static string StageStream(Stream stream)
    {
        var path = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-{Guid.NewGuid():N}.xlsx");
        try
        {
            using var output = File.Create(path);
            stream.CopyTo(output);
            return path;
        }
        catch
        {
            DeleteTemporaryFile(path);
            throw;
        }
    }

    private static void ValidateReadableStream(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead)
            throw new ArgumentException("The stream must be readable.", nameof(stream));
    }

    private static void ValidateCsvConfiguration(MiniExcelRustCsvReadOptions? configuration)
    {
        if (configuration is not null && (configuration.Delimiter == '\0' || configuration.Delimiter > 0x7f))
            throw new ArgumentException("The CSV delimiter must be a single-byte ASCII character.", nameof(configuration));
    }

    private static void DeleteTemporaryFile(string? path)
    {
        if (path is not null && File.Exists(path))
            File.Delete(path);
    }

    private static void EnsureAbiVersion()
    {
        var version = NativeMethods.GetAbiVersion();
        if (version != 1)
            throw new NotSupportedException($"MiniExcel Rust ABI version {version} is not supported.");
    }

    private static Exception CreateNativeException(int result)
    {
        var data = NativeMethods.GetLastError(out var length);
        var byteLength = checked((int)length.ToUInt64());
        if (data == IntPtr.Zero || byteLength == 0)
            return new InvalidOperationException($"MiniExcel Rust query failed with native error {result}.");

        var bytes = new byte[byteLength];
        Marshal.Copy(data, bytes, 0, byteLength);
        return new InvalidOperationException(Encoding.UTF8.GetString(bytes));
    }

    private sealed class FrameReader(byte[] frame)
    {
        private int _offset;

        public int ReadLength()
        {
            var value = ReadUInt32();
            if (value > int.MaxValue)
                throw new InvalidDataException("The native MiniExcel frame contains an unsupported length.");
            return (int)value;
        }

        public string ReadString()
        {
            var length = ReadLength();
            EnsureAvailable(length);
            var value = Encoding.UTF8.GetString(frame, _offset, length);
            _offset += length;
            return value;
        }

        public object? ReadValue()
        {
            EnsureAvailable(1);
            return frame[_offset++] switch
            {
                0 => null,
                1 => ReadBoolean(),
                2 => Convert.ToDouble(ReadInt64(), CultureInfo.InvariantCulture),
                3 => BitConverter.Int64BitsToDouble(ReadInt64()),
                4 => ReadString(),
                5 => DateTime.ParseExact(ReadString(), "yyyy-MM-dd", CultureInfo.InvariantCulture),
                6 => TimeSpan.Parse(ReadString(), CultureInfo.InvariantCulture),
                7 => DateTime.Parse(ReadString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
                8 => TimeSpan.FromMilliseconds(ReadInt64()),
                9 => ReadString(),
                var tag => throw new InvalidDataException($"The native MiniExcel frame contains unknown value tag {tag}.")
            };
        }

        public void EnsureComplete()
        {
            if (_offset != frame.Length)
                throw new InvalidDataException("The native MiniExcel frame contains trailing data.");
        }

        private bool ReadBoolean()
        {
            EnsureAvailable(1);
            return frame[_offset++] != 0;
        }

        public byte ReadByte()
        {
            EnsureAvailable(1);
            return frame[_offset++];
        }

        public uint ReadUInt32()
        {
            EnsureAvailable(sizeof(uint));
            var value = (uint)(frame[_offset]
                | frame[_offset + 1] << 8
                | frame[_offset + 2] << 16
                | frame[_offset + 3] << 24);
            _offset += sizeof(uint);
            return value;
        }

        private long ReadInt64()
        {
            EnsureAvailable(sizeof(long));
            ulong value = 0;
            for (var index = 0; index < sizeof(long); index++)
                value |= (ulong)frame[_offset + index] << (index * 8);
            _offset += sizeof(long);
            return unchecked((long)value);
        }

        private void EnsureAvailable(int length)
        {
            if (length < 0 || _offset > frame.Length - length)
                throw new InvalidDataException("The native MiniExcel frame is truncated.");
        }
    }

    private sealed class Utf8String : IDisposable
    {
        public Utf8String(string? value)
        {
            if (value is null)
                return;

            var bytes = Encoding.UTF8.GetBytes(value);
            Pointer = Marshal.AllocHGlobal(bytes.Length + 1);
            Marshal.Copy(bytes, 0, Pointer, bytes.Length);
            Marshal.WriteByte(Pointer, bytes.Length, 0);
        }

        public IntPtr Pointer { get; private set; }

        public void Dispose()
        {
            if (Pointer == IntPtr.Zero)
                return;

            Marshal.FreeHGlobal(Pointer);
            Pointer = IntPtr.Zero;
        }
    }

    private sealed class NativeQueryHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public NativeQueryHandle() : base(true) { }

        public NativeQueryHandle(IntPtr value) : this()
        {
            SetHandle(value);
        }

        protected override bool ReleaseHandle()
        {
            NativeMethods.QueryClose(handle);
            return true;
        }
    }

    private sealed class NativeBufferHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public NativeBufferHandle() : base(true) { }

        public NativeBufferHandle(IntPtr value) : this()
        {
            SetHandle(value);
        }

        protected override bool ReleaseHandle()
        {
            NativeMethods.BufferClose(handle);
            return true;
        }
    }

    private static class NativeMethods
    {
        private const string LibraryName = "miniexcel_ffi";

        [DllImport(LibraryName, EntryPoint = "miniexcel_abi_version", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern uint GetAbiVersion();

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_open", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryOpen(
            IntPtr path,
            byte useHeaderRow,
            IntPtr sheetName,
            IntPtr startCell,
            out IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_range_open", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryRangeOpen(
            IntPtr path,
            byte useHeaderRow,
            IntPtr sheetName,
            IntPtr startCell,
            IntPtr endCell,
            out IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_options_open", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryOptionsOpen(
            IntPtr path,
            byte useHeaderRow,
            IntPtr sheetName,
            IntPtr startCell,
            IntPtr endCell,
            byte ignoreEmptyRows,
            byte fillMergedCells,
            byte trimHeaders,
            byte enableSharedStringCache,
            ulong sharedStringCacheSize,
            IntPtr sharedStringCachePath,
            out IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_table_open", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryTableOpen(
            IntPtr path,
            IntPtr sheetName,
            IntPtr tableName,
            out IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_csv_open", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryCsvOpen(
            IntPtr path,
            byte useHeaderRow,
            byte delimiter,
            byte encoding,
            byte readEmptyAsNull,
            byte trimHeaders,
            out IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_get_csv_columns", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int GetCsvColumns(
            IntPtr path,
            byte useHeaderRow,
            byte delimiter,
            byte encoding,
            byte readEmptyAsNull,
            byte trimHeaders,
            out IntPtr handle,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_next_batch", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int QueryNextBatch(
            NativeQueryHandle handle,
            uint maxRows,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_query_close", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void QueryClose(IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_get_sheet_names", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int GetSheetNames(
            IntPtr path,
            out IntPtr handle,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_get_columns", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int GetColumns(
            IntPtr path,
            byte useHeaderRow,
            IntPtr sheetName,
            IntPtr startCell,
            out IntPtr handle,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_get_sheet_dimensions", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int GetSheetDimensions(
            IntPtr path,
            out IntPtr handle,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_get_sheet_info", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern int GetSheetInfo(
            IntPtr path,
            out IntPtr handle,
            out IntPtr data,
            out UIntPtr length);

        [DllImport(LibraryName, EntryPoint = "miniexcel_buffer_close", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern void BufferClose(IntPtr handle);

        [DllImport(LibraryName, EntryPoint = "miniexcel_last_error", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        internal static extern IntPtr GetLastError(out UIntPtr length);
    }
}