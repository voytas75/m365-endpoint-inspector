using M365.EndpointInspector.Storage;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Feed;

public enum CacheGeneration { Unavailable, Active, Previous }

public sealed record CacheView(CacheGeneration Generation, EndpointSnapshot? Snapshot,
    TimeSpan? Age, string? Error)
{
    public bool GenerationsMissing { get; internal init; }

    internal CacheView At(DateTimeOffset now)
    {
        if (Snapshot?.Observation is not { } observation)
            return this;
        var age = now - observation.RetrievedAt;
        return this with { Age = age < TimeSpan.Zero ? TimeSpan.Zero : age };
    }
}

internal sealed record RefreshState(int Format, Guid ClientRequestId, DateTimeOffset LastVersionCheck,
    DateTimeOffset NextCheck, int Failures, string? LastError)
{
    // The format-1 payload and checksum remain readable by older Release clients.
    [JsonIgnore] public DateTimeOffset? ServerNotBefore { get; init; }
    [JsonIgnore] public int Throttles { get; init; }

    internal DateTimeOffset Earliest(RefreshPolicy policy) => policy == RefreshPolicy.ManualDevelopment
        ? ServerNotBefore ?? LastVersionCheck : NextCheck;

    internal RefreshState WithThrottle(DateTimeOffset now, DateTimeOffset? retryAfter)
    {
        int count = Math.Min(Throttles + 1, 6);
        var until = retryAfter > now ? retryAfter.Value : now.AddHours(1 << (count - 1));
        if (ServerNotBefore > until) until = ServerNotBefore.Value;
        return this with { Throttles = count, ServerNotBefore = until, NextCheck = until > NextCheck ? until : NextCheck };
    }
}

internal sealed class CacheStore
{
    private sealed record CheckedDocument<T>(string Sha256, T Value);
    private sealed record ServerSchedule(string StateSha256, DateTimeOffset? NotBefore, int Throttles);
    private sealed record MachineDocument(string Sha256, RefreshState Value, CheckedDocument<ServerSchedule>? Server = null);
    private sealed record Generation(int Format, WorldwideObservation Observation, JsonElement Endpoints);

    private const int StateLimit = 16 * 1024;
    internal const int GenerationLimit = 6 * WorldwideFeed.EndpointBodyLimit + 65536;
    private static readonly JsonSerializerOptions Json = new()
    {
        AllowDuplicateProperties = false,
        RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true,
        Converters = { new JsonStringEnumConverter<WorldwideSource>(allowIntegerValues: false) }
    };

    private static readonly JsonSerializerOptions FileImportJson = new(Json)
    {
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    private readonly string directory;
    internal string DirectoryPath => directory;
    private readonly WorldwideSource source;
    private readonly TimeProvider clock;

    internal CacheStore(string machineAppDataRoot, WorldwideSource source, TimeProvider clock)
    {
        if (!Path.IsPathFullyQualified(machineAppDataRoot) ||
            machineAppDataRoot.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Contains(".."))
            throw new ArgumentException("Supply an absolute machine app-data directory without '..'.", nameof(machineAppDataRoot));
        directory = Path.Combine(Path.GetFullPath(machineAppDataRoot), source.ToString());
        this.source = source;
        this.clock = clock;
    }

    // Fixed leaf names only. Reject links, including dangling links, before opening or replacing files.
    private string Leaf(string name)
    {
        return LocalFile.CheckedPath(Path.Combine(directory, name));
    }

    internal async Task<FileStream> LockAsync(CancellationToken token)
    {
        string path = Leaf("refresh.lock");
        Directory.CreateDirectory(directory);
        var elapsed = Stopwatch.StartNew();
        while (true)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException ex) when ((ex.HResult & 0xffff) is 11 or 32 or 33 &&
                                         elapsed.Elapsed < TimeSpan.FromMinutes(2))
            {
                await Task.Delay(30, token).ConfigureAwait(false);
            }
        }
    }

    private static string Digest<T>(T value) =>
        Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(value, Json)));

    private static void WriteJson<T>(Stream stream, T value) =>
        JsonSerializer.Serialize(stream, new CheckedDocument<T>(Digest(value), value), Json);

    private static T ReadJson<T>(string path, int limit)
    {
        return ReadChecked<T>(LocalFile.ReadBytes(path, limit));
    }

    private static T ReadChecked<T>(ReadOnlySpan<byte> bytes, JsonSerializerOptions? options = null)
    {
        try
        {
            var document = JsonSerializer.Deserialize<CheckedDocument<T>>(bytes, options ?? Json);
            if (document == null || document.Value == null || document.Sha256 != Digest(document.Value))
                throw new FormatException("Cache document failed its integrity check.");
            return document.Value;
        }
        catch (InvalidOperationException ex)
        {
            throw new FormatException("Invalid JSON value in the cache document.", ex);
        }
    }

    private static RefreshState ReadState(string path)
    {
        var document = JsonSerializer.Deserialize<MachineDocument>(LocalFile.ReadBytes(path, StateLimit), Json)
            ?? throw new InvalidDataException("Missing machine state.");
        var state = document.Value;
        if (document.Sha256 != Digest(state) || state.Format != 1 || state.ClientRequestId == Guid.Empty ||
            state.Failures is < 0 or > 6 || state.LastVersionCheck == default ||
            state.NextCheck - state.LastVersionCheck < TimeSpan.FromHours(1))
            throw new InvalidDataException("Invalid machine identity or refresh schedule.");

        if (document.Server is { } server)
        {
            var schedule = server.Value;
            if (server.Sha256 != Digest(schedule) || schedule.StateSha256 != document.Sha256 ||
                schedule.Throttles is < 0 or > 6 || schedule.NotBefore > state.NextCheck)
                throw new InvalidDataException("Invalid server refresh schedule.");
            return state with { ServerNotBefore = schedule.NotBefore, Throttles = schedule.Throttles };
        }

        // Format 1 took the maximum of the failure timer (from completion) and all headers.
        // Only recognize ordinary local failures inside that timer plus its 30s request budget.
        // Unknown outcomes, throttling, Retry-After and longer deadlines stay in force.
        var error = state.LastError ?? "";
        bool parserOrTransport = error.StartsWith("endpoints[", StringComparison.Ordinal) ||
            error.StartsWith("Expected ", StringComparison.Ordinal) || error.StartsWith("Duplicate property ", StringComparison.Ordinal) ||
            error.StartsWith("Invalid Unicode in Worldwide JSON", StringComparison.Ordinal) ||
            error.Contains("BytePositionInLine:", StringComparison.Ordinal) ||
            error is "HTTP transport failed." or "Feed request timed out.";
        bool throttled = error.Contains("429", StringComparison.OrdinalIgnoreCase) ||
            error.Contains("Too Many Requests", StringComparison.OrdinalIgnoreCase);
        var oldDelay = TimeSpan.FromHours(1 << Math.Max(0, state.Failures - 1));
        var interval = state.NextCheck - state.LastVersionCheck;
        bool local = state.Failures > 0 && parserOrTransport && !throttled &&
            !error.Contains("Retry-After", StringComparison.OrdinalIgnoreCase) &&
            interval >= oldDelay && interval <= oldDelay + TimeSpan.FromSeconds(30);
        return state with { ServerNotBefore = local ? null : state.NextCheck, Throttles = throttled ? state.Failures : 0 };
    }

    internal RefreshState? LoadState(FileStream lease)
    {
        try
        {
            return ReadState(Leaf("machine-state.json"));
        }
        catch (FileNotFoundException) when (lease.Length == 0 &&
            !File.Exists(Leaf("active.json")) && !File.Exists(Leaf("previous.json")))
        {
            return null;
        }
    }

    internal void SaveState(RefreshState state, bool create = false)
    {
        string path = Leaf("machine-state.json");
        string digest = Digest(state);
        var schedule = new ServerSchedule(digest, state.ServerNotBefore, state.Throttles);
        var document = new MachineDocument(digest, state, new(Digest(schedule), schedule));
        AtomicFile.Write(path, stream => JsonSerializer.Serialize(stream, document, Json), temporary =>
        {
            if (ReadState(temporary) != state)
                throw new InvalidDataException("Refresh schedule failed read-back.");
        }, replace: !create, backup: null, CancellationToken.None);
    }

    private EndpointSnapshot ReadGeneration(string path)
    {
        var data = ReadJson<Generation>(path, GenerationLimit);
        if (data.Format != 1 || data.Observation.Source != source)
            throw new InvalidDataException("Unexpected cache format or transport provenance.");
        return WorldwideImporter.Import(data.Endpoints, data.Observation);
    }

    internal static EndpointSnapshot ImportSnapshotFile(ReadOnlySpan<byte> bytes, DateTimeOffset importedAt)
    {
        var data = ReadChecked<Generation>(bytes, FileImportJson);
        if (data.Format != 1) throw new InvalidDataException("Unsupported Feed snapshot format.");
        return WorldwideImporter.ImportFile(data.Endpoints, importedAt, data.Observation);
    }

    internal CacheView Read(bool previousOnly = false)
    {
        var errors = new List<string>();
        bool missing = !previousOnly;
        foreach (string name in previousOnly ? new[] { "previous" } : new[] { "active", "previous" })
        {
            try
            {
                var snapshot = ReadGeneration(Leaf(name + ".json"));
                return new CacheView(name == "active" ? CacheGeneration.Active : CacheGeneration.Previous,
                    snapshot, null,
                    errors.Count == 0 ? null : string.Join(" ", errors)).At(clock.GetUtcNow());
            }
            catch (Exception ex) when (IsStorageOrDataError(ex))
            {
                missing &= ex is FileNotFoundException or DirectoryNotFoundException;
                errors.Add($"{name}: {ex.Message}");
            }
        }
        return new(CacheGeneration.Unavailable, null, null, string.Join(" ", errors)) { GenerationsMissing = missing };
    }

    internal CacheView Publish(EndpointSnapshot snapshot, CacheView before, CancellationToken token)
    {
        string path = Leaf("active.json");
        var expected = before.Generation == CacheGeneration.Active ? before.Snapshot : null;
        string? backup = expected != null ? Leaf("previous.json") : null;
        var data = new Generation(1, snapshot.Observation!, snapshot.Source);
        AtomicFile.Write(path, stream => WriteJson(stream, data), temporary =>
        {
            ReadGeneration(temporary);
            VerifyActive(expected);
        }, replace: File.Exists(path), backup, token);
        return new CacheView(CacheGeneration.Active, snapshot, null, null).At(clock.GetUtcNow());
    }

    private void VerifyActive(EndpointSnapshot? expected)
    {
        EndpointSnapshot? active = null;
        try
        {
            active = ReadGeneration(Leaf("active.json"));
        }
        catch (Exception ex) when (ex is FileNotFoundException or FormatException or JsonException)
        {
            // A damaged active generation may be replaced, but must never rotate over the backup.
        }
        if (active?.Observation != expected?.Observation ||
            active != null && !JsonElement.DeepEquals(active.Source, expected!.Source))
            throw new IOException("Active generation changed during refresh.");
    }

    internal static bool IsStorageOrDataError(Exception ex) =>
        ex is IOException or UnauthorizedAccessException or JsonException or FormatException or InvalidDataException;
}
