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

    public static BrowserDocument Empty { get; } = new(null, null, "Brak otwartej migawki",
        "Pobierz wymagania Worldwide lub otwórz lokalny plik JSON/CSV.");

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
            catch (JsonException ex) { throw new InvalidDataException("Plik JSON przekracza limit 4 MiB dla listy lub koperty Synthetic.", ex); }
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
            catch (JsonException ex) when (syntheticOnly) { throw new FormatException($"Błędny JSON Synthetic. {ex.Message}", ex); }
        }
        token.ThrowIfCancellationRequested();
        var kind = snapshot.Provenance == "Synthetic" ? BrowserSource.Synthetic : BrowserSource.File;
        return Describe(snapshot, new(kind, now, fullPath, Csv: csv), now);
    }

    public static async Task<BrowserDocument> DownloadAsync(WorldwideFeed feed, TimeProvider clock, CancellationToken token)
    {
        var result = await feed.RefreshAsync(token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        var now = clock.GetUtcNow();
        string? details = result.Error == "Further requests deferred by Retry-After."
            ? "Serwer odroczył kolejne żądanie." : result.Error;
        if (result.Cache.Error is { } cacheError)
            details = string.Join("\n", new[] { details, $"Odczyt cache: {cacheError}" }.OfType<string>());
        var origin = new SourceInfo(BrowserSource.Internet, now, null, result.Cache.Generation, result.Status,
            result.LastVersionCheck, result.NextCheck, details, result.VersionCheckedBeforeDeferral);
        if (result.Cache.Snapshot is not { } snapshot)
            throw new InvalidDataException($"{RefreshNotice(origin)} Brak poprawnej kopii Worldwide.", Diagnostic(RefreshDetails(origin)));
        return Describe(snapshot, origin, now);
    }

    public static BrowserDocument OpenCache(string folder, TimeProvider clock, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var directory = new DirectoryInfo(folder);
        string? root = directory.Parent?.FullName;
        if (root is null || !Enum.TryParse<WorldwideSource>(directory.Name, true, out var source) ||
            !string.Equals(Enum.GetName(source), directory.Name, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Wybierz katalog MicrosoftWorldwide lub SimulatedWorldwide wewnątrz katalogu danych maszyny.");
        using var offlineClient = new HttpClient(new OfflineHandler());
        var feed = source == WorldwideSource.MicrosoftWorldwide
            ? WorldwideFeed.CreateMicrosoft(root, clock)
            : new WorldwideFeed(root, offlineClient, clock);
        var cache = feed.ReadCache();
        token.ThrowIfCancellationRequested();
        if (cache.Snapshot is not { } snapshot)
            return new(null, null, $"{TransportLabel(source)} · brak poprawnej kopii", "Brak danych do wyświetlenia.",
                "Nie udało się odczytać lokalnej migawki.", cache.Error);
        var now = clock.GetUtcNow();
        return Describe(snapshot, new(BrowserSource.Cache, now, directory.FullName, cache.Generation, Error: cache.Error), now);
    }

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
            throw new InvalidDataException("Niepoprawne metadane pochodzenia migawki.");
        var snapshot = origin.Kind switch
        {
            BrowserSource.Synthetic when saved.Observation == null => SyntheticImporter.Import(saved.Payload),
            BrowserSource.File => WorldwideImporter.ImportFile(saved.Payload, origin.OpenedAt, saved.Observation),
            BrowserSource.Cache or BrowserSource.Internet when saved.Observation != null =>
                WorldwideImporter.Import(saved.Payload, saved.Observation),
            _ => throw new InvalidDataException("Niespójne pochodzenie migawki.")
        };
        return Describe(snapshot, origin, now);
    }

    private static BrowserDocument Describe(EndpointSnapshot snapshot, SourceInfo origin, DateTimeOffset now)
    {
        var observation = snapshot.Observation;
        if (origin.Kind == BrowserSource.Synthetic)
            return new(snapshot, origin, "Dane syntetyczne · Synthetic",
                $"Wersja {snapshot.Version}   ·   Import wg pliku {Utc(snapshot.ImportedAt)}   ·   Wiek przy otwarciu {Age(now - snapshot.ImportedAt)}\n" +
                $"Otwarcie pliku {Utc(origin.OpenedAt)} · {origin.Path}");
        if (origin.Kind == BrowserSource.File)
        {
            string metadata = observation == null ? "Wersja i czas pobrania treści nieznane." :
                $"Według pliku ({observation.Source}): {ObservationText(observation, now)}";
            return new(snapshot, origin, origin.Csv ? "Plik · import CSV" : "Plik · import JSON", metadata +
                $"\nImport {Utc(origin.OpenedAt)} · Od importu {Age(now - origin.OpenedAt)} · {origin.Path}");
        }
        string generation = origin.Generation == CacheGeneration.Active ? "kopia aktywna" : "kopia zapasowa";
        string label = TransportLabel(observation!.Source);
        if (origin.Kind == BrowserSource.Internet) label = $"Internet · {label}";
        return new(snapshot, origin, $"{label} · {generation}", ObservationText(observation, now),
            origin.Kind == BrowserSource.Internet ? RefreshNotice(origin) :
                origin.Error == null ? null : "Nie udało się odczytać kopii aktywnej.", RefreshDetails(origin));
    }

    private static string TransportLabel(WorldwideSource source) => source == WorldwideSource.MicrosoftWorldwide
        ? "Cache produkcyjny · MicrosoftWorldwide" : "Cache symulowany · SimulatedWorldwide";

    private static string ObservationText(WorldwideObservation observation, DateTimeOffset now) =>
        $"Wersja zaobserwowana {observation.ObservedVersion} ({Utc(observation.VersionObservedAt)})   ·   " +
        $"Pobranie {Utc(observation.RetrievedAt)}   ·   Wiek przy otwarciu {Age(now - observation.RetrievedAt)}\n" +
        "Wersję sprawdzono przed pobraniem treści. W międzyczasie mogło zmienić się wydanie.";

    private static string RefreshNotice(SourceInfo origin)
    {
        string status = origin.Status switch
        {
            RefreshStatus.Updated => "Pobrano treść.",
            RefreshStatus.Unchanged => "Wersja bez zmian.",
            RefreshStatus.Deferred when origin.VersionCheckedBeforeDeferral =>
                "Sprawdzono wersję. Serwer odroczył pobranie treści.",
            // Older workspaces used the same status and reason for both kinds of deferral.
            RefreshStatus.Deferred when origin.Error?.StartsWith("Further requests deferred by Retry-After.", StringComparison.Ordinal) == true =>
                "Pobranie treści odroczone przez serwer.",
            RefreshStatus.Deferred => "Kontrola odroczona. Nie wysłano żądania.",
            RefreshStatus.Throttled => "Serwer ograniczył żądania (HTTP 429).",
            RefreshStatus.TimedOut => "Upłynął czas pobierania.",
            _ when origin.Error?.StartsWith("endpoints[", StringComparison.Ordinal) == true =>
                "Odrzucono odpowiedź. Dane endpointu nie przeszły walidacji.",
            _ => "Nie udało się pobrać danych."
        };
        string next = origin.NextCheck is { } date ? Utc(date) : "nieustalony";
        string attempt = $"Próba: {Utc(origin.OpenedAt)}.";
        if (origin.Status == RefreshStatus.Deferred && !origin.VersionCheckedBeforeDeferral && origin.LastCheck is { } previous)
        {
            string outcome = origin.Error switch
            {
                { } e when e.StartsWith("HTTP 429.", StringComparison.Ordinal) => " — ograniczenie żądań (HTTP 429)",
                { } e when e.StartsWith("endpoints[", StringComparison.Ordinal) => " — błąd walidacji danych",
                _ => ""
            };
            attempt = $"Poprzednia próba: {Utc(previous)}{outcome}.";
        }
        return $"{status} Najwcześniejsza kolejna kontrola: {next}. {attempt}";
    }

    private static string? RefreshDetails(SourceInfo origin)
    {
        string? details = origin.Error;
        if (origin.Status == RefreshStatus.Deferred && !origin.VersionCheckedBeforeDeferral && origin.LastCheck is { } previous &&
            details != null && !details.StartsWith("Odczyt cache:", StringComparison.Ordinal))
            return $"Poprzednia próba ({Utc(previous)}): {details}";
        return details;
    }

    internal static Exception? Diagnostic(string? details) => details == null ? null : new IOException(details);

    private static string Utc(DateTimeOffset time) => time.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss 'UTC'");
    private static string Age(TimeSpan age)
    {
        if (age < TimeSpan.Zero) age = TimeSpan.Zero;
        return $"{age.Days} d {age.Hours} godz. {age.Minutes} min";
    }

    private sealed class OfflineHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            => throw new InvalidOperationException("Odczyt cache jest lokalny.");
    }
}
