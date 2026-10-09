using System.Globalization;
using System.Text.Json;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>Typed accessors that record errors with paths.</summary>
internal sealed class JsonSpecReader(List<string> errors, List<string> warnings)
{
    public List<string> Errors { get; } = errors;

    public List<string> Warnings { get; } = warnings;

    public bool IsObject(JsonElement element, string path)
    {
        if (element.ValueKind == JsonValueKind.Object)
            return true;
        Errors.Add($"{path}: must be an object.");
        return false;
    }

    public void Unknown(JsonElement element, string path, params string[] known)
    {
        foreach (var property in element.EnumerateObject().Where(property => System.Array.IndexOf(known, property.Name) < 0))
            Warnings.Add($"{path}.{property.Name}: unknown property ignored.");
    }

    public string? String(JsonElement element, string path, string name, bool required, bool allowEmpty = false)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
        {
            if (required)
                Errors.Add($"{path}.{name}: required.");
            return null;
        }
        if (value.ValueKind != JsonValueKind.String)
        {
            Errors.Add($"{path}.{name}: must be a string (got {value.ValueKind.ToString().ToLowerInvariant()}).");
            return null;
        }
        var text = value.GetString()!;
        if (!allowEmpty && string.IsNullOrWhiteSpace(text))
        {
            if (required)
                Errors.Add($"{path}.{name}: must not be empty.");
            return null;
        }
        return text;
    }

    public string? Enum(JsonElement element, string path, string name, string[] allowed)
    {
        var value = String(element, path, name, required: false);
        if (value is not null && System.Array.IndexOf(allowed, value) < 0)
        {
            Errors.Add($"{path}.{name}: '{value}' must be one of {string.Join(", ", allowed)}.");
            return null;
        }
        return value;
    }

    public int? Int(JsonElement element, string path, string name, int min, int max)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var number))
        {
            Errors.Add($"{path}.{name}: must be a whole number.");
            return null;
        }
        if (number < min || number > max)
        {
            Errors.Add($"{path}.{name}: {number} is outside {min}-{max}.");
            return null;
        }
        return number;
    }

    public float? Float(JsonElement element, string path, string name, float min, float max)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Number)
        {
            Errors.Add($"{path}.{name}: must be a number.");
            return null;
        }
        var number = value.GetSingle();
        if (number < min || number > max)
        {
            Errors.Add($"{path}.{name}: {number.ToString(CultureInfo.InvariantCulture)} is outside {min.ToString(CultureInfo.InvariantCulture)}-{max.ToString(CultureInfo.InvariantCulture)}.");
            return null;
        }
        return number;
    }

    public bool? Bool(JsonElement element, string path, string name)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return value.GetBoolean();
        Errors.Add($"{path}.{name}: must be true or false.");
        return null;
    }

    public List<string>? StringArray(JsonElement element, string path, string name)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Array)
        {
            Errors.Add($"{path}.{name}: must be an array of strings.");
            return null;
        }
        var result = new List<string>();
        int index = 0;
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.String)
                result.Add(item.GetString()!);
            else if (item.ValueKind == JsonValueKind.Number)
                result.Add(item.GetRawText());
            else
                Errors.Add($"{path}.{name}[{index}]: must be a string.");
            index++;
        }
        return result;
    }

    public List<int>? IntArray(JsonElement element, string path, string name)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Array)
        {
            Errors.Add($"{path}.{name}: must be an array of whole numbers.");
            return null;
        }
        var result = new List<int>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Number && item.TryGetInt32(out var number))
                result.Add(number);
            else
                Errors.Add($"{path}.{name}: must contain whole numbers only.");
        }
        return result;
    }

    public List<float>? FloatArray(JsonElement element, string path, string name)
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Array)
        {
            Errors.Add($"{path}.{name}: must be an array of numbers.");
            return null;
        }
        var result = new List<float>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Number)
                result.Add(item.GetSingle());
            else
                Errors.Add($"{path}.{name}: must contain numbers only.");
        }
        return result;
    }

    public List<T>? Array<T>(JsonElement element, string path, string name, Func<JsonElement, string, T?> read) where T : class
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Array)
        {
            Errors.Add($"{path}.{name}: must be an array.");
            return null;
        }
        var result = new List<T>();
        int index = 0;
        foreach (var item in value.EnumerateArray())
        {
            if (read(item, $"{path}.{name}[{index}]") is { } parsed)
                result.Add(parsed);
            index++;
        }
        return result;
    }

    public T? Object<T>(JsonElement element, string path, string name, Func<JsonElement, string, T?> read) where T : class
    {
        if (!element.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
        {
            Errors.Add($"{path}.{name}: must be an object.");
            return null;
        }
        return read(value, $"{path}.{name}");
    }
}
