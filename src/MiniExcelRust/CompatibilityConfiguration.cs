using System.Globalization;

namespace MiniExcelLibs
{
    public interface IConfiguration
    {
        CultureInfo Culture { get; set; }
    }

    public enum ExcelType
    {
        XLSX,
        CSV,
        UNKNOWN
    }
}

namespace MiniExcelLibs.OpenXml
{
    public sealed class OpenXmlConfiguration : IConfiguration
    {
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        public bool FillMergedCells { get; set; }
        public bool TrimColumnNames { get; set; } = true;
        public bool IgnoreEmptyRows { get; set; }
        public bool EnableSharedStringCache { get; set; } = true;
        public ulong SharedStringCacheSize { get; set; } = 5 * 1024 * 1024;
        public string? SharedStringCachePath { get; set; } = Path.GetTempPath();

        internal MiniExcelRustReadOptions ToReadOptions() => new()
        {
            Culture = Culture,
            FillMergedCells = FillMergedCells,
            TrimColumnNames = TrimColumnNames,
            IgnoreEmptyRows = IgnoreEmptyRows,
            EnableSharedStringCache = EnableSharedStringCache,
            SharedStringCacheSize = SharedStringCacheSize,
            SharedStringCachePath = SharedStringCachePath
        };
    }
}

namespace MiniExcelLibs.Csv
{
    public sealed class CsvConfiguration : IConfiguration
    {
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        public char Seperator { get; set; } = ',';
        public bool ReadEmptyStringAsNull { get; set; }

        internal MiniExcelRustCsvReadOptions ToReadOptions() => new()
        {
            Culture = Culture,
            Delimiter = Seperator,
            ReadEmptyStringAsNull = ReadEmptyStringAsNull
        };
    }
}