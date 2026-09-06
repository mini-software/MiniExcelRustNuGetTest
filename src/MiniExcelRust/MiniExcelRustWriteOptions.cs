namespace MiniExcelLibs;

/// <summary>
/// Configures Rust-backed XLSX writes.
/// </summary>
public sealed class MiniExcelRustWriteOptions
{
    public string SheetName { get; set; } = "Sheet1";
    public bool OverwriteFile { get; set; }
    public bool PrintHeader { get; set; } = true;
    public bool AutoFilter { get; set; } = true;
    public bool RightToLeft { get; set; }
    public bool AutoWidth { get; set; }
    public bool WrapCellContents { get; set; }
    public double MinWidth { get; set; } = 8.42857143;
    public double MaxWidth { get; set; } = 200;
    public uint FreezeRowCount { get; set; } = 1;
    public ushort FreezeColumnCount { get; set; }
    public string DateFormat { get; set; } = "yyyy-mm-dd";
    public string TimeFormat { get; set; } = "hh:mm:ss";
    public string DateTimeFormat { get; set; } = "yyyy-mm-dd hh:mm:ss";
    public string DurationFormat { get; set; } = "[h]:mm:ss";
    public IDictionary<string, string> ColumnFormats { get; } = new Dictionary<string, string>();
    public IDictionary<string, double> ColumnWidths { get; } = new Dictionary<string, double>();
    public IDictionary<string, bool> HiddenColumns { get; } = new Dictionary<string, bool>();
}