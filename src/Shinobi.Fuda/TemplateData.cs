using System.Collections;
using System.Globalization;
using System.Reflection;

namespace Shinobi.Fuda;

/// <summary>
/// Marks a property to be excluded from template data flattening.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class TemplateIgnoreAttribute : Attribute
{
}

/// <summary>
/// Overrides the placeholder path segment used for a property.
/// Defaults to the property name.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class TemplateNameAttribute : Attribute
{
    public string Name { get; }

    public TemplateNameAttribute(string name) => Name = name;
}

/// <summary>
/// Applies a .NET format string when converting the property value to text,
/// e.g. "C2" for currency, "dd.MM.yyyy" for dates.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class TemplateFormatAttribute : Attribute
{
    public string Format { get; }

    public TemplateFormatAttribute(string format) => Format = format;
}

/// <summary>
/// Flattened data ready to feed into <see cref="DocxTemplateRenderer"/>.
///
/// Three buckets:
///  - Scalars: single values, dotted path (Customer.Name).
///  - Collections: flat row data for [[Name.Property]] placeholders inside a single
///    table row that gets cloned once per row-item (Records, Totals).
///  - Blocks: repeatable multi-element regions delimited by [[Repeat:X]]...[[EndRepeat:X]]
///    in the template. Each block item is itself a full nested TemplateData, because
///    (unlike a Collection row) a block item can contain its own collections
///    (e.g. a card has its own list of transactions).
///
/// A list property is auto-classified as a Block instead of a Collection if its
/// item type itself has any list-typed property — that's the signal that the item
/// needs its own scalars+collections rather than a single flat row.
/// </summary>
public sealed class TemplateData
{
    public Dictionary<string, string> Scalars { get; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, List<Dictionary<string, string>>> Collections { get; } =
        new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, List<TemplateData>> Blocks { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Reflects over a DTO graph and builds a TemplateData instance. Safe to call
    /// recursively (used internally to build each block item's nested TemplateData).
    /// </summary>
    public static TemplateData FromDto(object dto, CultureInfo? culture = null)
    {
        culture ??= CultureInfo.InvariantCulture;
        var data = new TemplateData();
        FlattenScalars(dto, prefix: null, data, culture);
        return data;
    }

    private static void FlattenScalars(object? obj, string? prefix, TemplateData data, CultureInfo culture)
    {
        if (obj is null) return;

        foreach (var prop in obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.GetIndexParameters().Length > 0) continue;
            if (prop.GetCustomAttribute<TemplateIgnoreAttribute>() is not null) continue;

            var value = prop.GetValue(obj);
            var name = prop.GetCustomAttribute<TemplateNameAttribute>()?.Name ?? prop.Name;
            var path = prefix is null ? name : $"{prefix}.{name}";

            if (value is null)
            {
                if (IsCollectionType(prop.PropertyType))
                    continue; // null list == empty, not an error

                data.Scalars[path] = string.Empty;
                continue;
            }

            if (value is string s)
            {
                data.Scalars[path] = s;
                continue;
            }

            if (value is IEnumerable enumerable)
            {
                var elementType = GetEnumerableElementType(prop.PropertyType);

                if (elementType is not null && HasNestedCollectionProperty(elementType))
                {
                    // Block: each item needs its own scalars + collections, not a flat row.
                    var blockItems = new List<TemplateData>();
                    foreach (var item in enumerable)
                    {
                        if (item is null) continue;
                        blockItems.Add(FromDto(item, culture));
                    }
                    data.Blocks[name] = blockItems;
                }
                else
                {
                    // Flat collection row (existing Records/Totals-style behavior).
                    var items = new List<Dictionary<string, string>>();
                    foreach (var item in enumerable)
                    {
                        var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        FlattenCollectionItem(item, row, culture);
                        items.Add(row);
                    }
                    data.Collections[name] = items;
                }

                continue;
            }

            if (IsScalarType(prop.PropertyType))
            {
                var format = prop.GetCustomAttribute<TemplateFormatAttribute>()?.Format;
                data.Scalars[path] = FormatScalar(value, format, culture);
                continue;
            }

            // Nested complex object: recurse to build the dotted path.
            FlattenScalars(value, path, data, culture);
        }
    }

    private static void FlattenCollectionItem(object? item, Dictionary<string, string> row, CultureInfo culture)
    {
        if (item is null) return;

        if (item is string || IsScalarType(item.GetType()))
        {
            row["Value"] = FormatScalar(item, null, culture);
            return;
        }

        foreach (var prop in item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.GetIndexParameters().Length > 0) continue;
            if (prop.GetCustomAttribute<TemplateIgnoreAttribute>() is not null) continue;

            var value = prop.GetValue(item);
            var name = prop.GetCustomAttribute<TemplateNameAttribute>()?.Name ?? prop.Name;
            var format = prop.GetCustomAttribute<TemplateFormatAttribute>()?.Format;

            row[name] = value is null ? string.Empty : FormatScalar(value, format, culture);
        }
    }

    private static bool IsCollectionType(Type type) =>
        type != typeof(string) && typeof(IEnumerable).IsAssignableFrom(type);

    private static Type? GetEnumerableElementType(Type type)
    {
        if (type.IsArray) return type.GetElementType();

        foreach (var candidate in type.GetInterfaces().Prepend(type))
        {
            if (candidate.IsGenericType && candidate.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                return candidate.GetGenericArguments()[0];
        }

        return null;
    }

    private static bool HasNestedCollectionProperty(Type itemType)
    {
        if (itemType == typeof(string) || IsScalarType(itemType)) return false;

        return itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Any(p => IsCollectionType(p.PropertyType));
    }

    private static bool IsScalarType(Type type)
    {
        var t = Nullable.GetUnderlyingType(type) ?? type;
        return t.IsPrimitive || t.IsEnum || t == typeof(decimal) || t == typeof(DateTime)
            || t == typeof(DateTimeOffset) || t == typeof(Guid) || t == typeof(TimeSpan);
    }

    private static string FormatScalar(object value, string? format, CultureInfo culture)
    {
        if (!string.IsNullOrEmpty(format) && value is IFormattable withFormat)
            return withFormat.ToString(format, culture);

        if (value is IFormattable formattable)
            return formattable.ToString(null, culture);

        return value.ToString() ?? string.Empty;
    }
}
