namespace MiniExcelLibs;

public sealed class MiniExcelRustMappingException : InvalidOperationException
{
    internal MiniExcelRustMappingException(
        string columnName,
        int row,
        object? value,
        Type targetType,
        Exception innerException)
        : base($"The value {value} in column {columnName} at row {row} cannot be assigned to {targetType.Name}.", innerException)
    {
        ColumnName = columnName;
        Row = row;
        Value = value;
        TargetType = targetType;
    }

    public string ColumnName { get; }
    public int Row { get; }
    public object? Value { get; }
    public Type TargetType { get; }
}