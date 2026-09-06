using System.ComponentModel;
using System.Globalization;
using System.Reflection;

namespace MiniExcelLibs;

internal static class MiniExcelRustMapper
{
    public static IEnumerable<IDictionary<string, object?>> ToRows<T>(IEnumerable<T> values)
    {
        var mappings = CreateMappings(typeof(T))
            .OrderBy(mapping => mapping.Index ?? int.MaxValue)
            .ToList();
        foreach (var value in values)
        {
            if (value is null)
                throw new ArgumentException("Typed export rows cannot contain null values.", nameof(values));
            IDictionary<string, object?> row = new Dictionary<string, object?>(StringComparer.Ordinal);
            foreach (var mapping in mappings)
                row.Add(mapping.Names[0], NormalizeWriteValue(mapping.GetValue(value)));
            yield return row;
        }
    }

    public static IEnumerable<T> Map<T>(IEnumerable<IDictionary<string, object?>> rows)
        where T : class, new()
    {
        var mappings = CreateMappings(typeof(T));
        var rowIndex = 1;
        foreach (var row in rows)
        {
            var instance = new T();
            var values = row.Values.ToList();
            foreach (var mapping in mappings)
            {
                object? value = null;
                var found = mapping.Index is int index
                    ? index >= 0 && index < values.Count && Assign(values[index], out value)
                    : TryGetValue(row, mapping.Names, out value);
                if (!found)
                    continue;

                try
                {
                    mapping.SetValue(instance, ConvertValue(value, mapping.ValueType));
                }
                catch (Exception error) when (error is InvalidCastException or FormatException or OverflowException or ArgumentException)
                {
                    throw new MiniExcelRustMappingException(
                        mapping.Names[0],
                        rowIndex,
                        value,
                        mapping.ValueType,
                        error);
                }
            }
            yield return instance;
            rowIndex++;
        }
    }

    private static IReadOnlyList<MemberMapping> CreateMappings(Type type)
    {
        const BindingFlags flags = BindingFlags.Instance | BindingFlags.Public;
        var members = type.GetProperties(flags)
            .Where(property => property.SetMethod is not null && property.GetIndexParameters().Length == 0)
            .Cast<MemberInfo>()
            .Concat(type.GetFields(flags).Where(HasMiniExcelAttribute));
        return members
            .Where(member => !IsIgnored(member))
            .Select(CreateMapping)
            .ToList();
    }

    private static MemberMapping CreateMapping(MemberInfo member)
    {
        var names = new List<string> { member.Name };
        int? index = null;
        foreach (var attribute in member.CustomAttributes)
        {
            var name = attribute.AttributeType.Name;
            if (name is "ExcelColumnNameAttribute" or "MiniExcelColumnNameAttribute")
            {
                AddConstructorName(attribute, names);
                AddNamedString(attribute, "Name", names);
                AddAliases(attribute, names);
            }
            else if (name is "ExcelColumnIndexAttribute" or "MiniExcelColumnIndexAttribute")
            {
                index = ReadIndex(attribute);
            }
            else if (name is "ExcelColumnAttribute" or "MiniExcelColumnAttribute")
            {
                AddNamedString(attribute, "Name", names);
                AddAliases(attribute, names);
                index = ReadNamedInt(attribute, "Index") ?? index;
            }
        }

        var valueType = member is PropertyInfo property ? property.PropertyType : ((FieldInfo)member).FieldType;
        return new MemberMapping(member, names.Distinct(StringComparer.OrdinalIgnoreCase).ToArray(), index, valueType);
    }

    private static bool TryGetValue(
        IDictionary<string, object?> row,
        IReadOnlyList<string> names,
        out object? value)
    {
        foreach (var name in names)
        {
            if (row.TryGetValue(name, out value))
                return true;
            var match = row.FirstOrDefault(cell => string.Equals(cell.Key, name, StringComparison.OrdinalIgnoreCase));
            if (match.Key is not null)
            {
                value = match.Value;
                return true;
            }
        }
        value = null;
        return false;
    }

    private static bool Assign(object? source, out object? value)
    {
        value = source;
        return true;
    }

    private static object? ConvertValue(object? value, Type targetType)
    {
        if (value is null || value is DBNull)
        {
            if (!targetType.IsValueType || Nullable.GetUnderlyingType(targetType) is not null)
                return null;
            return Activator.CreateInstance(targetType);
        }

        var effectiveType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (effectiveType.IsInstanceOfType(value))
            return value;
        if (effectiveType == typeof(string))
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        if (effectiveType == typeof(Guid))
            return Guid.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!);
        if (effectiveType == typeof(Uri))
            return new Uri(Convert.ToString(value, CultureInfo.InvariantCulture)!, UriKind.RelativeOrAbsolute);
        if (effectiveType == typeof(DateTime))
            return value is double serial ? DateTime.FromOADate(serial) : Convert.ToDateTime(value, CultureInfo.InvariantCulture);
        if (effectiveType == typeof(DateTimeOffset))
            return DateTimeOffset.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture);
        if (effectiveType == typeof(TimeSpan))
            return value is double milliseconds
                ? TimeSpan.FromMilliseconds(milliseconds)
                : TimeSpan.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture);
        if (effectiveType == typeof(bool))
        {
            var text = Convert.ToString(value, CultureInfo.InvariantCulture);
            return text switch { "1" => true, "0" => false, _ => bool.Parse(text!) };
        }
        if (effectiveType.IsEnum)
        {
            var text = Convert.ToString(value, CultureInfo.InvariantCulture)!;
            var described = effectiveType.GetFields()
                .FirstOrDefault(field => field.GetCustomAttribute<DescriptionAttribute>()?.Description == text);
            return Enum.Parse(effectiveType, described?.Name ?? text, ignoreCase: true);
        }
        return Convert.ChangeType(value, effectiveType, CultureInfo.InvariantCulture);
    }

    private static object? NormalizeWriteValue(object? value)
    {
        if (value is null)
            return null;
        var type = value.GetType();
        if (type.IsEnum)
        {
            var field = type.GetField(value.ToString()!);
            return field?.GetCustomAttribute<DescriptionAttribute>()?.Description ?? value.ToString();
        }
        if (value is Guid or Uri)
            return value.ToString();
        return value;
    }

    private static bool HasMiniExcelAttribute(MemberInfo member) =>
        member.CustomAttributes.Any(attribute => attribute.AttributeType.Name.IndexOf("Excel", StringComparison.Ordinal) >= 0);

    private static bool IsIgnored(MemberInfo member) => member.CustomAttributes.Any(attribute =>
        attribute.AttributeType.Name is "ExcelIgnoreAttribute" or "MiniExcelIgnoreAttribute" &&
        (attribute.ConstructorArguments.Count == 0 || attribute.ConstructorArguments[0].Value is not false));

    private static void AddConstructorName(CustomAttributeData attribute, IList<string> names)
    {
        if (attribute.ConstructorArguments.Count > 0 && attribute.ConstructorArguments[0].Value is string value && value.Length > 0)
            names.Insert(0, value);
    }

    private static void AddNamedString(CustomAttributeData attribute, string propertyName, IList<string> names)
    {
        var argument = attribute.NamedArguments.FirstOrDefault(item => item.MemberName == propertyName);
        if (argument.TypedValue.Value is string value && value.Length > 0)
            names.Insert(0, value);
    }

    private static void AddAliases(CustomAttributeData attribute, ICollection<string> names)
    {
        var argument = attribute.NamedArguments.FirstOrDefault(item => item.MemberName == "Aliases");
        if (argument.TypedValue.Value is IEnumerable<CustomAttributeTypedArgument> aliases)
        {
            foreach (var alias in aliases)
            {
                if (alias.Value is string value && value.Length > 0)
                    names.Add(value);
            }
        }
    }

    private static int? ReadIndex(CustomAttributeData attribute)
    {
        if (attribute.ConstructorArguments.Count == 0)
            return null;
        var value = attribute.ConstructorArguments[0].Value;
        if (value is int index)
            return index;
        if (value is string columnName)
            return ColumnNameToIndex(columnName);
        return null;
    }

    private static int? ReadNamedInt(CustomAttributeData attribute, string propertyName)
    {
        var argument = attribute.NamedArguments.FirstOrDefault(item => item.MemberName == propertyName);
        return argument.TypedValue.Value is int value && value >= 0 ? value : null;
    }

    private static int ColumnNameToIndex(string columnName)
    {
        var index = 0;
        foreach (var character in columnName.ToUpperInvariant())
        {
            if (character is < 'A' or > 'Z')
                throw new ArgumentException($"Invalid Excel column name '{columnName}'.");
            index = checked(index * 26 + character - 'A' + 1);
        }
        return index - 1;
    }

    private sealed class MemberMapping(
        MemberInfo member,
        string[] names,
        int? index,
        Type valueType)
    {
        public string[] Names { get; } = names;
        public int? Index { get; } = index;
        public Type ValueType { get; } = valueType;

        public void SetValue(object target, object? value)
        {
            if (member is PropertyInfo property)
                property.SetValue(target, value);
            else
                ((FieldInfo)member).SetValue(target, value);
        }

        public object? GetValue(object target)
        {
            return member is PropertyInfo property
                ? property.GetValue(target)
                : ((FieldInfo)member).GetValue(target);
        }
    }
}