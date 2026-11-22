using System.Net;
using System.Text.Json;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Feed;

public enum RefreshPolicy { Release, ManualDevelopment }

public enum RefreshStatus { Updated, Unchanged, Deferred, Throttled, TimedOut, Failed }

public sealed record RefreshResult(RefreshStatus Status, CacheView Cache,
    DateTimeOffset? LastVersionCheck, DateTimeOffset? NextCheck, string? Error)
{
    public bool VersionCheckedBeforeDeferral { get; init; }
}

public sealed class WorldwideFeed
{
    internal const int EndpointBodyLimit = 4 * 1024 * 1024;
    private static readonly TimeSpan MinimumInterval = TimeSpan.FromHours(1);
    private static readonly HttpClient MicrosoftClient = new(new SocketsHttpHandler
    {
        AllowAutoRedirect = false,
        UseCookies = false
    });
    private readonly HttpClient client;
    private readonly TimeProvider clock;
    private readonly WorldwideSource source;
    private readonly CacheStore store;
    private readonly RefreshPolicy policy;

    // An injected transport has its own provenance and on-disk namespace.
    public WorldwideFeed(string machineAppDataRoot, HttpClient simulatedClient, TimeProvider clock)
        : this(machineAppDataRoot, simulatedClient, clock, RefreshPolicy.Release) { }

    public WorldwideFeed(string machineAppDataRoot, HttpClient simulatedClient, TimeProvider clock, RefreshPolicy policy)
        : this(machineAppDataRoot, simulatedClient, clock, WorldwideSource.SimulatedWorldwide, policy) { }

    public static WorldwideFeed CreateMicrosoft(string machineAppDataRoot, TimeProvider? clock = null) =>
        CreateMicrosoft(machineAppDataRoot, clock, RefreshPolicy.Release);

    public static WorldwideFeed CreateMicrosoft(string machineAppDataRoot, TimeProvider? clock, RefreshPolicy policy) =>
        new(machineAppDataRoot, MicrosoftClient, clock ?? TimeProvider.System, WorldwideSource.MicrosoftWorldwide, policy);

    private WorldwideFeed(string root, HttpClient client, TimeProvider clock, WorldwideSource source, RefreshPolicy policy)
    {
        this.client = client ?? throw new ArgumentNullException(nameof(client));
        this.clock = clock ?? throw new ArgumentNullException(nameof(clock));
        if (!Enum.IsDefined(policy)) throw new ArgumentOutOfRangeException(nameof(policy));
        this.policy = policy;
        this.source = source;
        store = new CacheStore(root, source, clock);
    }

    public const int SnapshotFileLimit = CacheStore.GenerationLimit;

    public static EndpointSnapshot ImportSnapshotFile(ReadOnlySpan<byte> bytes, DateTimeOffset importedAt)
    {
        if (bytes.Length > SnapshotFileLimit) throw new InvalidDataException("Migawka Feed przekracza limit.");
        return CacheStore.ImportSnapshotFile(bytes, importedAt);
    }

    public CacheView ReadCache(bool previousOnly = false) => store.Read(previousOnly);

    public async Task<RefreshResult> RefreshAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            using var lease = await store.LockAsync(cancellationToken).ConfigureAwait(false);
            var saved = store.LoadState(lease);
            var cache = store.Read();
            var now = clock.GetUtcNow();
            if (saved != null && now < saved.Earliest(policy))
                return new(RefreshStatus.Deferred, cache, saved.LastVersionCheck, saved.Earliest(policy), saved.LastError);

            cancellationToken.ThrowIfCancellationRequested();
            var releaseNext = now + MinimumInterval;
            if (saved?.NextCheck > releaseNext) releaseNext = saved.NextCheck;
            int failures = saved?.Failures ?? 0;
            if (policy == RefreshPolicy.Release) failures = Math.Min(failures + 1, 6);
            var state = new RefreshState(1, saved?.ClientRequestId ?? Guid.NewGuid(), now, releaseNext,
                failures, "Refresh interrupted before completion.") { Throttles = saved?.Throttles ?? 0 };
            store.SaveState(state, create: saved == null);
            if (lease.Length == 0)
            {
                // If the schedule is subsequently lost, a restart must not create a new identity.
                lease.WriteByte(1);
                lease.Flush(flushToDisk: true);
            }
            return await FetchAsync(state, cache, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (CacheStore.IsStorageOrDataError(ex))
        {
            return new(RefreshStatus.Failed, store.Read(), null, null, ex.Message);
        }
    }

    private async Task<RefreshResult> FetchAsync(RefreshState state, CacheView cache, CancellationToken callerToken)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30), clock);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(callerToken, timeout.Token);
        CancellationToken token = linked.Token;
        DateTimeOffset? retryAfter = null;
        bool received429 = false;
        try
        {
            using var version = await GetAsync("version", 64 * 1024).ConfigureAwait(false);
            string observed = WorldwideImporter.ReadPublishedVersion(version.RootElement);
            var observedAt = clock.GetUtcNow();
            var status = RefreshStatus.Unchanged;
            if (cache.Snapshot?.Observation?.ObservedVersion != observed)
            {
                if (retryAfter > clock.GetUtcNow())
                {
                    var deferred = Completed("Further requests deferred by Retry-After.");
                    store.SaveState(deferred);
                    return new(RefreshStatus.Deferred, cache.At(clock.GetUtcNow()), deferred.LastVersionCheck,
                        deferred.Earliest(policy), deferred.LastError) { VersionCheckedBeforeDeferral = true };
                }
                using var endpoints = await GetAsync("endpoints", EndpointBodyLimit).ConfigureAwait(false);
                var observation = new WorldwideObservation(source, observed, observedAt, clock.GetUtcNow());
                var candidate = WorldwideImporter.Import(endpoints.RootElement, observation);
                cache = store.Publish(candidate, cache, token);
                status = RefreshStatus.Updated;
            }
            else
            {
                token.ThrowIfCancellationRequested();
            }

            var complete = Completed(null);
            store.SaveState(complete);
            return new(status, cache.At(clock.GetUtcNow()), complete.LastVersionCheck, complete.Earliest(policy), null);
        }
        catch (OperationCanceledException) when (callerToken.IsCancellationRequested)
        {
            // The reservation is durable. Never publish or retry a cancelled candidate.
            throw;
        }
        catch (Exception ex) when (CacheStore.IsStorageOrDataError(ex) || ex is HttpRequestException or OperationCanceledException)
        {
            var status = ex switch
            {
                HttpRequestException { StatusCode: HttpStatusCode.TooManyRequests } => RefreshStatus.Throttled,
                OperationCanceledException => RefreshStatus.TimedOut,
                _ => RefreshStatus.Failed
            };
            // Transport exception messages may include the request URI and its machine identifier.
            string reason = ex switch
            {
                HttpRequestException http when http.StatusCode is { } code => $"HTTP {(int)code}.",
                HttpRequestException => "HTTP transport failed.",
                OperationCanceledException => "Feed request timed out.",
                _ => ex.Message
            };
            if (status == RefreshStatus.Throttled && !received429)
                state = state.WithThrottle(clock.GetUtcNow(), null);
            var next = state.NextCheck;
            if (policy == RefreshPolicy.Release)
            {
                var localRetry = clock.GetUtcNow().AddHours(1 << (state.Failures - 1));
                if (localRetry > next) next = localRetry;
            }
            var failed = state with
            {
                NextCheck = next,
                LastError = reason.Length > 512 ? reason[..512] : reason
            };
            store.SaveState(failed);
            return new(status, store.Read(), failed.LastVersionCheck, failed.Earliest(policy), failed.LastError);
        }

        RefreshState Completed(string? error) => state with
        {
            Failures = policy == RefreshPolicy.Release ? 0 : state.Failures,
            Throttles = 0, LastError = error
        };

        async Task<JsonDocument> GetAsync(string resource, int limit)
        {
            token.ThrowIfCancellationRequested();
            var uri = new Uri($"https://endpoints.office.com/{resource}/Worldwide?ClientRequestId={state.ClientRequestId:D}");
            using var request = new HttpRequestMessage(HttpMethod.Get, uri);
            using var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token)
                .ConfigureAwait(false);
            var received = clock.GetUtcNow();
            retryAfter = FeedResponse.RetryAfter(response, received);
            var server = FeedResponse.ServerDeadline(response, received, retryAfter);
            if (state.ServerNotBefore is { } prior && (server == null || prior > server)) server = prior;
            var next = received + MinimumInterval;
            if (state.NextCheck > next) next = state.NextCheck;
            if (server > next) next = server.Value;
            var updated = state with { NextCheck = next, ServerNotBefore = server };
            received429 = response.StatusCode == HttpStatusCode.TooManyRequests;
            if (received429) updated = updated.WithThrottle(received, retryAfter);
            if (updated != state)
            {
                state = updated;
                store.SaveState(state);
            }

            return await FeedResponse.ReadAsync(response, limit, token).ConfigureAwait(false);
        }
    }
}
