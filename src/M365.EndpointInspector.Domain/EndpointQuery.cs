using System.Text.Json;

namespace M365.EndpointInspector.Domain;

public enum Transport { Tcp, Udp }
public enum ChangeKind { Added, Modified, Removed, Unchanged }
public enum ResultState { Unmeasured, Success, Failure, Inconclusive, NotTestable, Skipped, Cancelled }

public sealed record SyntheticResult(int EndpointId, ResultState State);

public sealed class EndpointFilter
{
    public int? Id { get; init; }
    public ServiceArea? ServiceArea { get; init; }
    public Category? Category { get; init; }
    public bool? Required { get; init; }
    public bool? ExpressRoute { get; init; }
    public string? ServiceAreaDisplayName { get; init; }
    public string? Notes { get; init; }
    public Transport? Protocol { get; init; }
    public string? Ports { get; init; }
    public string? Domain { get; init; }
    public string? AddressOrCidr { get; init; }
    public IReadOnlyList<FieldMatch> Fields { get; init; } = Array.Empty<FieldMatch>();
    public ChangeKind? Change { get; init; }
    public ResultState? Result { get; init; }
}

public sealed record EndpointMatch(Endpoint Entry, ChangeKind? Change, IReadOnlyList<ResultState> Results);

public sealed record Selection(IReadOnlyList<EndpointMatch> Hits, IReadOnlyList<EndpointMatch> CommonDependencies)
{
    public Category? HighestPriorityCategory => Hits.Count == 0 ? null : Hits[0].Entry.Category;
}

public static class EndpointQuery
{
    public static Selection Select(EndpointSnapshot current, EndpointFilter filter,
        EndpointSnapshot? previous = null, IEnumerable<SyntheticResult>? samples = null)
    {
        ValidateChoice(filter.ServiceArea, nameof(filter.ServiceArea));
        ValidateChoice(filter.Category, nameof(filter.Category));
        ValidateChoice(filter.Protocol, nameof(filter.Protocol));
        ValidateChoice(filter.Change, nameof(filter.Change));
        ValidateChoice(filter.Result, nameof(filter.Result));
        if (filter.Id is <= 0)
            throw new ArgumentOutOfRangeException(nameof(filter.Id), "An endpoint id must be positive.");
        if (filter.Change.HasValue && previous is null)
            throw new ArgumentException("A change filter requires a previous snapshot.", nameof(previous));
        if (filter.Ports is not null && filter.Protocol is null)
            throw new ArgumentException("A port query requires a protocol.", nameof(filter));

        string? host = filter.Domain is null ? null : DnsPattern.Normalize(filter.Domain, allowWildcard: false);
        var network = filter.AddressOrCidr is null ? null : IpPrefix.Parse(filter.AddressOrCidr);
        var ports = filter.Ports is null ? null : PortRange.ParseList(filter.Ports);
        var states = (samples ?? Array.Empty<SyntheticResult>()).ToLookup(sample => sample.EndpointId, sample =>
        {
            if (!Enum.IsDefined(sample.State) || sample.State == ResultState.Unmeasured)
                throw new ArgumentException("A sample must contain an observed or assigned state.", nameof(samples));
            if (!current.Entries.ContainsKey(sample.EndpointId) && previous?.Entries.ContainsKey(sample.EndpointId) != true)
                throw new ArgumentException($"Unknown sample endpoint id {sample.EndpointId}.", nameof(samples));
            return sample.State;
        });

        var hits = new List<EndpointMatch>();
        var dependencies = new List<EndpointMatch>();
        foreach (var (entry, change) in Compare(current, previous).OrderBy(row => row.Entry.Category)
                     .ThenBy(row => row.Entry.Id))
        {
            var results = states[entry.Id].Distinct().Order().DefaultIfEmpty(ResultState.Unmeasured).ToArray();
            var row = new EndpointMatch(entry, change, Array.AsReadOnly(results));
            if (filter.ServiceArea.HasValue && entry.ServiceArea != filter.ServiceArea)
            {
                if (entry.ServiceArea == ServiceArea.Common && change != ChangeKind.Removed)
                    dependencies.Add(row);
                continue;
            }
            if (filter.Change.HasValue && filter.Change != change ||
                filter.Result.HasValue && !results.Contains(filter.Result.Value))
                continue;
            if (Matches(entry, filter, host, network, ports))
                hits.Add(row);
        }
        return new Selection(hits.AsReadOnly(), dependencies.AsReadOnly());
    }

    private static void ValidateChoice<T>(T? choice, string name) where T : struct, Enum
    {
        if (choice.HasValue && !Enum.IsDefined(choice.Value))
            throw new ArgumentOutOfRangeException(name, choice.Value, $"Unknown {name}.");
    }

    private static IEnumerable<(Endpoint Entry, ChangeKind? Change)> Compare(
        EndpointSnapshot current, EndpointSnapshot? previous)
    {
        foreach (var entry in current.Entries.Values)
        {
            ChangeKind? change = null;
            if (previous is not null)
            {
                if (!previous.Entries.TryGetValue(entry.Id, out var old))
                    change = ChangeKind.Added;
                else
                    change = JsonElement.DeepEquals(entry.Source, old.Source)
                        ? ChangeKind.Unchanged : ChangeKind.Modified;
            }
            yield return (entry, change);
        }
        if (previous is not null)
        {
            foreach (var old in previous.Entries.Values)
                if (!current.Entries.ContainsKey(old.Id))
                    yield return (old, ChangeKind.Removed);
        }
    }

    private static bool Matches(Endpoint entry, EndpointFilter filter, string? host, IpPrefix? network,
        IReadOnlyList<PortRange>? ports)
    {
        if (filter.Id.HasValue && entry.Id != filter.Id ||
            filter.Category.HasValue && entry.Category != filter.Category ||
            filter.Required.HasValue && entry.Required != filter.Required ||
            filter.ExpressRoute.HasValue && entry.ExpressRoute != filter.ExpressRoute ||
            filter.ServiceAreaDisplayName is not null && entry.ServiceAreaDisplayName != filter.ServiceAreaDisplayName ||
            filter.Notes is not null && entry.Notes != filter.Notes)
            return false;

        if (filter.Protocol.HasValue)
        {
            var available = filter.Protocol == Transport.Tcp ? entry.TcpPorts : entry.UdpPorts;
            if (available.Count == 0 || ports is not null && !ports.Any(p => available.Any(p.Overlaps)))
                return false;
        }

        return (host is null || entry.DnsPatterns.Any(p => DnsPattern.Matches(p, host))) &&
            (network is null || entry.Networks.Any(network.Overlaps)) &&
            filter.Fields.All(field => field.Matches(entry.Source));
    }
}
