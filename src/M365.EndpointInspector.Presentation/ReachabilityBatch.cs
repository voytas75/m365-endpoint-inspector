using System.Text.Json;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Presentation;

public sealed record SkippedReachabilityTarget(string Value, string Reason);

public sealed class ReachabilityBatch
{
    public const int Parallelism = 4;
    public IReadOnlyList<ReachabilityTarget> Targets { get; }
    public IReadOnlyList<SkippedReachabilityTarget> SkippedTargets { get; }
    public int Skipped => SkippedTargets.Count;
    public string Summary => $"{Targets.Count} target(s), {Skipped} skipped, {Duplicates} duplicate(s).";
    public int Duplicates { get; }

    private ReachabilityBatch(IReadOnlyList<ReachabilityTarget> targets, IReadOnlyList<SkippedReachabilityTarget> skipped, int duplicates) =>
        (Targets, SkippedTargets, Duplicates) = (targets, skipped, duplicates);

    public static ReachabilityBatch FromValues(IEnumerable<string> values)
    {
        var targets = new List<ReachabilityTarget>();
        var seen = new HashSet<ReachabilityTarget>();
        var skipped = new List<SkippedReachabilityTarget>();
        int duplicates = 0;
        foreach (var value in values)
        {
            var target = ReachabilityTarget.Parse(value);
            if (target == null)
            {
                var reason = value.Contains('*') ? "Wildcard not expanded"
                    : IsCidr(value) ? "CIDR range not expanded"
                    : "Not a supported concrete target";
                skipped.Add(new(value, reason));
            }
            else if (!seen.Add(target)) duplicates++;
            else targets.Add(target);
        }
        return new(targets.AsReadOnly(), skipped.AsReadOnly(), duplicates);
    }

    private static bool IsCidr(string value)
    {
        if (!value.Contains('/') || value.Contains("://", StringComparison.Ordinal)) return false;
        try { IpPrefix.Parse(value.Trim()); return true; }
        catch (FormatException) { return false; }
    }

    internal static ReachabilityBatch FromTargets(IEnumerable<ReachabilityTarget> targets) => new(targets.ToArray(), [], 0);

    public static IEnumerable<string> SelectedValues(IEnumerable<BrowserRow> hits, IEnumerable<BrowserRow> common) =>
        hits.Concat(common).SelectMany(row => row.SearchableTargets).Select(target => target.Value);

    public static ReachabilityBatch FromArea(EndpointSnapshot snapshot, ServiceArea area) =>
        FromValues(snapshot.Entries.Values.Where(entry => entry.ServiceArea == area)
            .SelectMany(entry => entry.DnsPatterns.Concat(Addresses(entry))));

    private static IEnumerable<string> Addresses(Endpoint entry) =>
        entry.Source.TryGetProperty("ips", out var ips) && ips.ValueKind == JsonValueKind.Array
            ? ips.EnumerateArray().Select(ip => ip.GetString()!) : [];

    public Task RunAsync(ReachabilityProbe probe, Action<ProbeResult> publish, CancellationToken cancellation,
        Action<ReachabilityTarget>? admitted = null, Action<ReachabilityTarget, ProbeStage?>? activity = null,
        Action<ReachabilityTarget, bool>? finished = null) =>
        Parallel.ForEachAsync(Targets, new ParallelOptions
        {
            MaxDegreeOfParallelism = Parallelism, CancellationToken = cancellation
        }, async (target, token) =>
        {
            if (token.IsCancellationRequested) return;
            admitted?.Invoke(target);
            var completed = false;
            try
            {
                await probe.RunAsync(target, row =>
                {
                    publish(row);
                    completed = row.Stage == ProbeStage.HttpsSystemProxy && row.Status != ProbeStatus.Cancelled;
                }, token, activity == null ? null : stage => activity(target, stage)).ConfigureAwait(false);
            }
            finally
            {
                finished?.Invoke(target, completed);
            }
        });
}
