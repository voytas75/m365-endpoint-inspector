using System.Text.Json;

namespace M365.EndpointInspector.Domain;

internal static class Schema
{
    public static JsonElement Required(JsonElement obj, string name, JsonValueKind kind)
    {
        if (!obj.TryGetProperty(name, out var value))
            throw new FormatException($"Missing field '{name}'.");
        if (value.ValueKind != kind)
            throw new FormatException($"{name} must be a JSON {kind}.");
        return value;
    }

    public static T EnumName<T>(JsonElement obj, string name) where T : struct, Enum
    {
        string text = Required(obj, name, JsonValueKind.String).GetString()!;
        if (!Enum.TryParse<T>(text, out var value) || Enum.GetName(value) != text)
            throw new FormatException($"Unknown {name} '{text}'.");
        return value;
    }

    public static string? OptionalString(JsonElement obj, string name)
    {
        if (!obj.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.String)
            throw new FormatException($"{name} must be a string or null.");
        return value.GetString();
    }

    public static bool? OptionalBoolean(JsonElement obj, string name)
    {
        if (!obj.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new FormatException($"{name} must be a boolean or null.");
        return value.GetBoolean();
    }

    public static Dictionary<int, Endpoint> ReadEntries(JsonElement array, bool worldwide = false)
    {
        var entries = new Dictionary<int, Endpoint>();
        foreach (var item in array.EnumerateArray())
        {
            try
            {
                if (item.ValueKind != JsonValueKind.Object)
                    throw new FormatException("Expected an endpoint object.");
                var entry = new Endpoint(item, worldwide);
                if (!entries.TryAdd(entry.Id, entry))
                    throw new FormatException($"Duplicate id {entry.Id}.");
            }
            catch (FormatException ex)
            {
                throw new FormatException($"endpoints[{entries.Count}]: {ex.Message}", ex);
            }
        }
        return entries;
    }

    public static void ValidateProperties(JsonElement root, bool validateStrings = false)
    {
        var path = new List<string>();
        Visit(root);

        void Visit(JsonElement element)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                var names = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in element.EnumerateObject())
                {
                    string name = property.Name;
                    if (!names.Add(name))
                        throw new FormatException($"Duplicate property '{name}' at ${string.Concat(path)}.");
                    path.Add("." + name);
                    Visit(property.Value);
                    path.RemoveAt(path.Count - 1);
                }
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                int index = 0;
                foreach (var item in element.EnumerateArray())
                {
                    path.Add($"[{index++}]");
                    Visit(item);
                    path.RemoveAt(path.Count - 1);
                }
            }
            else if (validateStrings && element.ValueKind == JsonValueKind.String)
            {
                // JsonDocument decodes UTF-8 and escaped surrogate pairs lazily.
                _ = element.GetString();
            }
        }
    }
}
