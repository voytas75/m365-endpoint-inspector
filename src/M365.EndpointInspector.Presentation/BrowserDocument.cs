using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed;
using M365.EndpointInspector.Storage;

namespace M365.EndpointInspector.Presentation;

public enum BrowserSource { Synthetic, File, Cache, Internet }

public sealed record SourceInfo(BrowserSource Kind, DateTimeOffset OpenedAt, string? Path,
    CacheGeneration? Generation = null, RefreshStatus? Status = null, DateTimeOffset? LastCheck = null,
    DateTimeOffset? NextCheck = null, string? Error = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)] bool VersionCheckedBeforeDeferral = false,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)] bool Csv = false);

internal sealed record SavedDocument(SourceInfo Origin, WorldwideObservation? Observation, JsonElement Payload);

public sealed class BrowserDocument
{
    internal const int FileLimit = 4 * 1024 * 1024;
    public EndpointSnapshot? Snapshot { get; }
    public SourceInfo? Origin { get; }
    public string SourceLabel { get; }
    public string Metadata { get; }
    public string? Notice { get; }
    public string? Details { get; }
    public bool HasDetails => !string.IsNullOrEmpty(Details);
    public bool IsSynthetic => Snapshot?.Provenance == "Synthetic";

    private BrowserDocument(EndpointSnapshot? snapshot, SourceInfo? origin, string label, string metadata, string? notice = null, string? details = null)
        => (Snapshot, Origin, SourceLabel, Metadata, Notice, Details) = (snapshot, origin, label, metadata, notice, details);

    public static BrowserDocument Empty { get; } = new(null, null, "No snapshot open",
        "Download Worldwide requirements or open a local JSON/CSV file.");

    public static BrowserDocument OpenFile(string path, TimeProvider clock, CancellationToken token) =>
        ReadFile(path, clock, token, syntheticOnly: false);

    public static BrowserDocument OpenSynthetic(string path, TimeProvider clock, CancellationToken token) =>
        ReadFile(path, clock, token, syntheticOnly: true);

    private static BrowserDocument ReadFile(string path, TimeProvider clock, CancellationToken token, bool syntheticOnly)
    {
        token.ThrowIfCancellationRequested();
        string fullPath = LocalFile.CheckedPath(Path.Combine(Environment.CurrentDirectory, path));
        var bytes = LocalFile.ReadBytes(fullPath, syntheticOnly ? FileLimit : WorldwideFeed.SnapshotFileLimit);
        ReadOnlyMemory<byte> body = bytes;
        if (body.Span.StartsWith(Encoding.UTF8.Preamble)) body = body[Encoding.UTF8.Preamble.Length..];
        token.ThrowIfCancellationRequested();
        var now = clock.GetUtcNow();
        EndpointSnapshot snapshot;
        string extension = System.IO.Path.GetExtension(fullPath);
        string leading = Encoding.UTF8.GetString(body.Span).TrimStart();
        bool csv = !syntheticOnly && (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase) ||
            (!extension.Equals(".json", StringComparison.OrdinalIgnoreCase) &&
             !leading.StartsWith("[") && !leading.StartsWith("{") &&
             WorldwideCsvImporter.Header.Split(',').All(name =>
                 leading.Split('\n')[0].TrimEnd('\r').Split(',').Select(cell => cell.Trim('"')).Contains(name))));
        if (!csv) new UTF8Encoding(false, true).GetCharCount(body.Span);
        if (csv) snapshot = WorldwideCsvImporter.Import(bytes, now);
        else if (bytes.Length > FileLimit)
        {
            // Only a checked cache generation may exceed the wire limit after JSON escaping.
            try { snapshot = WorldwideFeed.ImportSnapshotFile(body.Span, now); }
            catch (JsonException ex) { throw new InvalidDataException("JSON file exceeds the 4 MiB limit for an array or Synthetic envelope.", ex); }
        }
        else
        {
            try
            {
                using var json = JsonDocument.Parse(body);
                var root = json.RootElement;
                if (!syntheticOnly && root.ValueKind == JsonValueKind.Array)
                    snapshot = WorldwideImporter.ImportFile(root, now);
                else if (!syntheticOnly && root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("Sha256", out _) && !root.TryGetProperty("provenance", out _))
                    snapshot = WorldwideFeed.ImportSnapshotFile(body.Span, now);
                else
                    snapshot = SyntheticImporter.Import(root);
            }
            catch (JsonException ex) when (syntheticOnly) { throw new FormatException($"Invalid Synthetic JSON. {ex.Message}", ex); }
        }
        token.ThrowIfCancellationRequested();
        var kind = snapshot.Provenance == "Synthetic" ? BrowserSource.Synthetic : BrowserSource.File;
        return Describe(snapshot, new(kind, now, fullPath, Csv: csv), now);
    }

    public static async Task<BrowserWorkspace> BootstrapAsync(WorkspaceLoad loaded, BrowserWorkspace current,
        WorldwideFeed feed, TimeProvider clock, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!loaded.CanSave || current.Document.Snapshot != null)
            return current;
        var cache = feed.ReadCache();
        token.ThrowIfCancellationRequested();
        if (cache.Snapshot is { } snapshot)
            return current.Open(DescribeCache(snapshot, cache, feed.CacheDirectory, clock.GetUtcNow()));
        if (!cache.GenerationsMissing) return current;
        return current.Open(await DownloadAsync(feed, clock, token).ConfigureAwait(false));
    }

    public static async Task<BrowserDocument> DownloadAsync(WorldwideFeed feed, TimeProvider clock, CancellationToken token)
    {
        var result = await feed.RefreshAsync(token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        var now = clock.GetUtcNow();
        string? details = result.Error == "Further requests deferred by Retry-After."
            ? "The server deferred the next request." : result.Error;
        if (result.Cache.Error is { } cacheError)
            details = string.Join("\n", new[] { details, $"Cache read: {cacheError}" }.OfType<string>());
        var origin = new SourceInfo(BrowserSource.Internet, now, null, result.Cache.Generation, result.Status,
            result.LastVersionCheck, result.NextCheck, details, result.VersionCheckedBeforeDeferral);
        if (result.Cache.Snapshot is not { } snapshot)
            throw new InvalidDataException($"{RefreshNotice(origin)} No valid Worldwide cache copy.", Diagnostic(RefreshDetails(origin)));
        return Describe(snapshot, origin, now);
    }

    public static BrowserDocument OpenCache(string folder, TimeProvider clock, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var directory = new DirectoryInfo(folder);
        string? root = directory.Parent?.FullName;
        if (root is null || !Enum.TryParse<WorldwideSource>(directory.Name, true, out var source) ||
            !string.Equals(Enum.GetName(source), directory.Name, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Select the MicrosoftWorldwide or SimulatedWorldwide folder inside the machine data directory.");
        using var offlineClient = new HttpClient(new OfflineHandler());
        var feed = source == WorldwideSource.MicrosoftWorldwide
            ? WorldwideFeed.CreateMicrosoft(root, clock)
            : new WorldwideFeed(root, offlineClient, clock);
        var cache = feed.ReadCache();
        token.ThrowIfCancellationRequested();
        if (cache.Snapshot is not { } snapshot)
            return new(null, null, $"{TransportLabel(source)} · no valid cache copy", "No data to display.",
                "Could not read the local snapshot.", cache.Error);
        return DescribeCache(snapshot, cache, directory.FullName, clock.GetUtcNow());
    }

    private static BrowserDocument DescribeCache(EndpointSnapshot snapshot, CacheView cache, string path, DateTimeOffset now) =>
        Describe(snapshot, new(BrowserSource.Cache, now, path, cache.Generation, Error: cache.Error), now);

    internal SavedDocument? Pack() => Snapshot == null ? null : new(Origin!, Snapshot.Observation, Snapshot.Source);

    internal static BrowserDocument Restore(SavedDocument saved, DateTimeOffset now)
    {
        var origin = saved.Origin;
        bool cached = origin.Kind is BrowserSource.Cache or BrowserSource.Internet;
        if (origin.Csv && (origin.Kind != BrowserSource.File || saved.Observation != null) ||
            origin.OpenedAt == default || !Enum.IsDefined(origin.Kind) ||
            (origin.Kind == BrowserSource.Internet ? origin.Path != null : string.IsNullOrWhiteSpace(origin.Path)) ||
            origin.Path?.Length > 32767 || origin.Error?.Length > 65536 ||
            (cached ? origin.Generation is not (CacheGeneration.Active or CacheGeneration.Previous) : origin.Generation != null) ||
            (origin.Kind == BrowserSource.Internet) != origin.Status.HasValue ||
            (origin.Status is { } status && !Enum.IsDefined(status)) ||
            (origin.Kind != BrowserSource.Internet && (origin.NextCheck != null || origin.LastCheck != null)) ||
            (origin.VersionCheckedBeforeDeferral && origin.Status != RefreshStatus.Deferred) ||
            origin.NextCheck < origin.LastCheck || (!cached && origin.Error != null))
            throw new InvalidDataException("Invalid snapshot provenance metadata.");
        var snapshot = origin.Kind switch
        {
            BrowserSource.Synthetic when saved.Observation == null => SyntheticImporter.Import(saved.Payload),
            BrowserSource.File => WorldwideImporter.ImportFile(saved.Payload, origin.OpenedAt, saved.Observation),
            BrowserSource.Cache or BrowserSource.Internet when saved.Observation != null =>
                WorldwideImporter.Import(saved.Payload, saved.Observation),
            _ => throw new InvalidDataException("Inconsistent snapshot provenance.")
        };
        return Describe(snapshot, origin, now);
    }

    private static BrowserDocument Describe(EndpointSnapshot snapshot, SourceInfo origin, DateTimeOffset now)
    {
        var observation = snapshot.Observation;
        if (origin.Kind == BrowserSource.Synthetic)
            return new(snapshot, origin, "Synthetic data · Synthetic",
                $"Version {snapshot.Version}   ·   Import declared by file {Utc(snapshot.ImportedAt)}   ·   Age at opening {Age(now - snapshot.ImportedAt)}\n" +
                $"File opened {Utc(origin.OpenedAt)} · {origin.Path}");
        if (origin.Kind == BrowserSource.File)
        {
            string metadata = observation == null ? "Version and content retrieval time unknown." :
                $"Declared by file ({observation.Source}): {ObservationText(observation, now)}";
            return new(snapshot, origin, origin.Csv ? "File · CSV import" : "File · JSON import", metadata +
                $"\nImport {Utc(origin.OpenedAt)} · Since import {Age(now - origin.OpenedAt)} · {origin.Path}");
        }
        string generation = origin.Generation == CacheGeneration.Active ? "active copy" : "previous copy";
        string label = TransportLabel(observation!.Source);
        if (origin.Kind == BrowserSource.Internet) label = $"Internet · {label}";
        return new(snapshot, origin, $"{label} · {generation}", ObservationText(observation, now),
            origin.Kind == BrowserSource.Internet ? RefreshNotice(origin) :
                origin.Error == null ? null : "Could not read the active cache copy.", RefreshDetails(origin));
    }

    private static string TransportLabel(WorldwideSource source) => source == WorldwideSource.MicrosoftWorldwide
        ? "Production cache · MicrosoftWorldwide" : "Simulated cache · SimulatedWorldwide";

    private static string ObservationText(WorldwideObservation observation, DateTimeOffset now) =>
        $"Observed version {observation.ObservedVersion} ({Utc(observation.VersionObservedAt)})   ·   " +
        $"Retrieved {Utc(observation.RetrievedAt)}   ·   Age at opening {Age(now - observation.RetrievedAt)}\n" +
        "The version was checked before content retrieval. The release may have changed in between.";

    private static string RefreshNotice(SourceInfo origin)
    {
        string status = origin.Status switch
        {
            RefreshStatus.Updated => "Content downloaded.",
            RefreshStatus.Unchanged => "Version unchanged.",
            RefreshStatus.Deferred when origin.VersionCheckedBeforeDeferral =>
                "Version checked. The server deferred content retrieval.",
            // Older workspaces used the same status and reason for both kinds of deferral.
            RefreshStatus.Deferred when origin.Error?.StartsWith("Further requests deferred by Retry-After.", StringComparison.Ordinal) == true =>
                "Content retrieval deferred by the server.",
            RefreshStatus.Deferred => "Check deferred. No request sent.",
            RefreshStatus.Throttled => "The server limited requests (HTTP 429).",
            RefreshStatus.TimedOut => "Download timed out.",
            _ when origin.Error?.StartsWith("endpoints[", StringComparison.Ordinal) == true =>
                "Response rejected. Endpoint data failed validation.",
            _ => "Could not download data."
        };
        string next = origin.NextCheck is { } date ? Utc(date) : "unknown";
        string attempt = $"Attempt: {Utc(origin.OpenedAt)}.";
        if (origin.Status == RefreshStatus.Deferred && !origin.VersionCheckedBeforeDeferral && origin.LastCheck is { } previous)
        {
            string outcome = origin.Error switch
            {
                { } e when e.StartsWith("HTTP 429.", StringComparison.Ordinal) => " — request throttling (HTTP 429)",
                { } e when e.StartsWith("endpoints[", StringComparison.Ordinal) => " — data validation error",
                _ => ""
            };
            attempt = $"Previous attempt: {Utc(previous)}{outcome}.";
        }
        return $"{status} Earliest next check: {next}. {attempt}";
    }

    private static string? RefreshDetails(SourceInfo origin)
    {
        string? details = origin.Error;
        if (origin.Status == RefreshStatus.Deferred && !origin.VersionCheckedBeforeDeferral && origin.LastCheck is { } previous &&
            details != null && !details.StartsWith("Cache read:", StringComparison.Ordinal) &&
            // Compatibility with diagnostic text persisted by the older Polish UI.
            !details.StartsWith("Odczyt cache:", StringComparison.Ordinal))
            return $"Previous attempt ({Utc(previous)}): {details}";
        return details;
    }

    internal static Exception? Diagnostic(string? details) => details == null ? null : new IOException(details);

    private static string Utc(DateTimeOffset time) => time.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss 'UTC'");
    private static string Age(TimeSpan age)
    {
        if (age < TimeSpan.Zero) age = TimeSpan.Zero;
        return $"{age.Days} d {age.Hours} h {age.Minutes} min";
    }

    private sealed class OfflineHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            => throw new InvalidOperationException("Cache reads are local.");
    }
}
