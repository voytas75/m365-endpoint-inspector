using System.Globalization;
using System.Text.Json;

namespace M365.EndpointInspector.Domain;

// A JSON Pointer relative to exactly one entry. Equality is typed JSON equality.
public sealed class FieldMatch
{
    private readonly string[] segments;
    private readonly JsonElement expected;

    private FieldMatch(string path, JsonElement expected)
    {
        if (!path.StartsWith('/'))
            throw new FormatException("A field path must start with '/'.");
        segments = path[1..].Split('/').Select(Decode).ToArray();
        this.expected = expected;
    }

    public static FieldMatch Equal(string path, JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Undefined)
            throw new ArgumentException("Supply a JSON value, including null, or use Missing.", nameof(value));
        return new FieldMatch(path, value.Clone());
    }

    public static FieldMatch Missing(string path) => new(path, default);

    internal bool Matches(JsonElement entry)
    {
        var value = entry;
        foreach (string segment in segments)
        {
            if (value.ValueKind == JsonValueKind.Object && value.TryGetProperty(segment, out var child))
                value = child;
            else if (value.ValueKind == JsonValueKind.Array && IsIndex(segment, out int index) &&
                     index < value.GetArrayLength())
                value = value[index];
            else
                return expected.ValueKind == JsonValueKind.Undefined;
        }
        return expected.ValueKind != JsonValueKind.Undefined && JsonElement.DeepEquals(value, expected);
    }

    private static bool IsIndex(string segment, out int index)
    {
        index = 0;
        return segment.Length > 0 && (segment.Length == 1 || segment[0] != '0') &&
            segment.All(char.IsAsciiDigit) &&
            int.TryParse(segment, NumberStyles.None, CultureInfo.InvariantCulture, out index);
    }

    private static string Decode(string token)
    {
        for (int i = 0; i < token.Length; i++)
        {
            if (token[i] != '~')
                continue;
            i++;
            if (i == token.Length || token[i] is not ('0' or '1'))
                throw new FormatException($"Invalid JSON Pointer escape in '{token}'.");
        }
        return token.Replace("~1", "/").Replace("~0", "~");
    }
}
