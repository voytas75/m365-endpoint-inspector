using System.Text.Json;

namespace M365.EndpointInspector.Domain;

public static class SyntheticImporter
{
    public static EndpointSnapshot Import(string json)
    {
        using var document = Parse(json);
        return Import(document.RootElement);
    }

    public static EndpointSnapshot Import(JsonElement envelope)
    {
        JsonElement root = envelope.Clone();
        if (root.ValueKind != JsonValueKind.Object)
            throw new FormatException("Expected a synthetic snapshot object.");
        Schema.ValidateProperties(root, validateStrings: true);
        if (Schema.Required(root, "provenance", JsonValueKind.String).GetString() != "Synthetic")
            throw new FormatException("provenance must be Synthetic.");

        string version = Schema.Required(root, "version", JsonValueKind.String).GetString()!;
        if (string.IsNullOrWhiteSpace(version))
            throw new FormatException("version must contain text.");
        var timestamp = Schema.Required(root, "importedAt", JsonValueKind.String);
        string iso = timestamp.GetString()!;
        bool hasOffset = iso.EndsWith('Z') || (iso.Length >= 6 && (iso[^6] is '+' or '-') && iso[^3] == ':');
        if (!hasOffset || !iso.Contains('T') || !timestamp.TryGetDateTimeOffset(out var importedAt))
            throw new FormatException("importedAt must be an ISO timestamp with an explicit time zone.");

        var entries = Schema.ReadEntries(Schema.Required(root, "endpoints", JsonValueKind.Array));
        return new EndpointSnapshot(root, version, importedAt, entries, "Synthetic");
    }

    private static JsonDocument Parse(string json)
    {
        try
        {
            return JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            throw new FormatException($"Malformed snapshot JSON: {ex.Message}", ex);
        }
    }
}
