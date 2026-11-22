using System.Collections.ObjectModel;
using System.Text.Json;

namespace M365.EndpointInspector.Domain;

public enum ServiceArea { Common, Exchange, SharePoint, Skype }
public enum Category { Optimize, Allow, Default }

public sealed class EndpointSnapshot
{
    // Synthetic envelope or Worldwide response array. Entries share the owned JSON.
    public JsonElement Source { get; }
    public string Provenance { get; }
    public WorldwideObservation? Observation { get; }
    public string? Version { get; }
    public DateTimeOffset ImportedAt { get; }
    public IReadOnlyDictionary<int, Endpoint> Entries { get; }

    internal EndpointSnapshot(JsonElement source, string? version, DateTimeOffset importedAt,
        Dictionary<int, Endpoint> entries, string provenance, WorldwideObservation? observation = null)
    {
        Source = source;
        Provenance = provenance;
        Observation = observation;
        Version = version;
        ImportedAt = importedAt;
        Entries = new ReadOnlyDictionary<int, Endpoint>(entries);
    }
}

public sealed class Endpoint
{
    public JsonElement Source { get; }
    public int Id { get; }
    public ServiceArea ServiceArea { get; }
    public Category Category { get; }
    public string? ServiceAreaDisplayName { get; }
    public bool? Required { get; }
    public bool? ExpressRoute { get; }
    public string? Notes { get; }
    public IReadOnlyList<string> DnsPatterns { get; }
    public IReadOnlyList<IpPrefix> Networks { get; }
    public IReadOnlyList<PortRange> TcpPorts { get; }
    public IReadOnlyList<PortRange> UdpPorts { get; }

    internal Endpoint(JsonElement source, bool worldwide)
    {
        Source = source;
        var id = Schema.Required(source, "id", JsonValueKind.Number);
        if (!id.TryGetInt32(out int number) || number < 1)
            throw new FormatException("id must be a positive 32-bit integer.");
        Id = number;
        ServiceArea = Schema.EnumName<ServiceArea>(source, "serviceArea");
        Category = Schema.EnumName<Category>(source, "category");
        ServiceAreaDisplayName = Schema.OptionalString(source, "serviceAreaDisplayName");
        Notes = Schema.OptionalString(source, "notes");
        Required = Schema.OptionalBoolean(source, "required");
        ExpressRoute = Schema.OptionalBoolean(source, "expressRoute");
        DnsPatterns = ParseTargets(source, "urls", name => DnsPattern.Normalize(name, allowWildcard: true, worldwide: worldwide));
        Networks = ParseTargets(source, "ips", IpPrefix.Parse);
        if (DnsPatterns.Count == 0 && Networks.Count == 0)
            throw new FormatException("Expected at least one target in urls or ips.");
        TcpPorts = ParsePorts(source, "tcpPorts");
        UdpPorts = ParsePorts(source, "udpPorts");
    }

    private static IReadOnlyList<T> ParseTargets<T>(JsonElement source, string field, Func<string, T> parse)
    {
        if (!source.TryGetProperty(field, out var list) || list.ValueKind == JsonValueKind.Null)
            return Array.Empty<T>();
        if (list.ValueKind != JsonValueKind.Array)
            throw new FormatException($"{field} must be an array or null.");

        var targets = new List<T>();
        foreach (var item in list.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
                throw new FormatException($"{field}[{targets.Count}] must be a string.");
            try
            {
                targets.Add(parse(item.GetString()!));
            }
            catch (FormatException ex)
            {
                throw new FormatException($"{field}[{targets.Count}]: {ex.Message}", ex);
            }
        }
        return targets.AsReadOnly();
    }

    private static IReadOnlyList<PortRange> ParsePorts(JsonElement source, string field)
    {
        string? text = Schema.OptionalString(source, field);
        if (text is null)
            return Array.Empty<PortRange>();
        try
        {
            return PortRange.ParseList(text);
        }
        catch (FormatException ex)
        {
            throw new FormatException($"{field}: {ex.Message}", ex);
        }
    }
}
