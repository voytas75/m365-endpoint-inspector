using System.Globalization;
using System.Text.Json;

namespace M365.EndpointInspector.Domain;

public enum WorldwideSource { MicrosoftWorldwide, SimulatedWorldwide }

// The version resource and the endpoint resource are two independent observations.
public sealed record WorldwideObservation(WorldwideSource Source, string ObservedVersion,
    DateTimeOffset VersionObservedAt, DateTimeOffset RetrievedAt);

public static class WorldwideImporter
{
    private static void ValidateVersion(string version)
    {
        if (version is not { Length: 10 } || !version.All(char.IsAsciiDigit) ||
            !DateOnly.TryParseExact(version[..8], "yyyyMMdd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _))
            throw new FormatException("Expected a published version in YYYYMMDDnn format.");
    }

    public static EndpointSnapshot Import(JsonElement response, WorldwideObservation observation)
    {
        return ImportRows(response, observation.RetrievedAt, observation.Source.ToString(), observation);
    }

    // An observation embedded in a file is a declaration from its author.
    public static EndpointSnapshot ImportFile(JsonElement response, DateTimeOffset importedAt,
        WorldwideObservation? declared = null) => ImportRows(response, importedAt, "File", declared);

    private static EndpointSnapshot ImportRows(JsonElement response, DateTimeOffset importedAt, string provenance,
        WorldwideObservation? observation)
    {
        if (importedAt == default) throw new FormatException("Missing import time.");
        if (observation != null)
        {
            ValidateVersion(observation.ObservedVersion);
            if (!Enum.IsDefined(observation.Source) || observation.VersionObservedAt == default ||
                observation.RetrievedAt < observation.VersionObservedAt)
                throw new FormatException("Invalid Worldwide observation metadata.");
        }
        if (response.ValueKind != JsonValueKind.Array || response.GetArrayLength() == 0)
            throw new FormatException("Expected a nonempty Worldwide endpoint array.");

        var source = response.Clone();
        ValidateJson(source);
        return new EndpointSnapshot(source, observation?.ObservedVersion, importedAt,
            Schema.ReadEntries(source, worldwide: true), provenance, observation);
    }

    public static string ReadPublishedVersion(JsonElement row)
    {
        ValidateJson(row);
        if (row.ValueKind == JsonValueKind.Array && row.GetArrayLength() == 1)
            row = row[0];
        if (row.ValueKind != JsonValueKind.Object ||
            !row.TryGetProperty("instance", out var instance) || instance.ValueKind != JsonValueKind.String ||
            instance.GetString() != "Worldwide" ||
            !row.TryGetProperty("latest", out var latest) || latest.ValueKind != JsonValueKind.String)
            throw new FormatException("Expected one Worldwide version record with a latest version string.");
        string version = latest.GetString()!;
        ValidateVersion(version);
        return version;
    }

    private static void ValidateJson(JsonElement value)
    {
        try
        {
            Schema.ValidateProperties(value, validateStrings: true);
        }
        catch (InvalidOperationException ex)
        {
            throw new FormatException("Invalid Unicode in Worldwide JSON.", ex);
        }
    }
}
