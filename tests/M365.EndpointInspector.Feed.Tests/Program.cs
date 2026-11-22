using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using M365.EndpointInspector.Domain;
using static M365.EndpointInspector.Feed.Tests.SimulatedTransport;

namespace M365.EndpointInspector.Feed.Tests;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        if (args is ["--worker", var root, var mode])
            return await Worker(root, mode);

        (string Name, Func<Task> Run)[] cases =
        [
            ("first retrieval, observation times, filters and provenance isolation", FirstRetrieval),
            ("explicit DEV checks, validation retry, identity and shared Release reservation", PolicyCases.Manual),
            ("both policies honor short server deadlines and storage restrictions", PolicyCases.ServerLimits),
            ("repeated DEV throttling without Retry-After", PolicyCases.RepeatedThrottle),
            ("concurrent mixed clients, cancelled wait and durable server freshness", PolicyCases.ConcurrentCancellation),
            ("checked original state, conservative legacy holds and Release readback", PolicyCases.Legacy),
            ("hourly boundary, unchanged and changed publication", Schedule),
            ("publication between version and body", PublicationRace),
            ("concurrent clients and process restart", Processes),
            ("429 dates, deltas, fallback and persisted backoff", FailureCases.Throttling),
            ("freshness headers and storage restrictions", FailureCases.Headers),
            ("malformed versions, endpoint schemas and synthetic envelopes", FailureCases.Schema),
            ("bounded, partial and timed-out HTTP responses", FailureCases.Body),
            ("cancellation and durable reservation", FailureCases.Cancellation),
            ("corrupt active, previous generation and unavailable cache", StorageCases.Generations),
            ("interrupted candidate writes and conflicting publications", StorageCases.Atomicity),
            ("paths, links, lost state and identity", StorageCases.PathsAndIdentity)
        ];
        int failed = 0;
        foreach (var (name, run) in cases)
        {
            try
            {
                await run();
                Console.WriteLine($"PASS {name}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"FAIL {name}: {ex}");
                failed++;
            }
        }
        Console.WriteLine($"{cases.Length - failed}/{cases.Length} simulated transport groups passed.");
        return failed == 0 ? 0 : 1;
    }

    private static async Task FirstRetrieval()
    {
        using var f = new Fixture();
        var now = f.Clock.GetUtcNow();
        Check.Equal(CacheGeneration.Unavailable, f.Feed.ReadCache().Generation);
        f.Transport.Reply("version", Version(First));
        f.Transport.Queue("endpoints", _ =>
        {
            f.Clock.Advance(TimeSpan.FromSeconds(2));
            return Task.FromResult(Json(Rows));
        });
        var fetched = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Updated, fetched.Status);
        var snapshot = fetched.Cache.Snapshot!;
        Check.Equal("SimulatedWorldwide", snapshot.Provenance);
        Check.Equal(First, snapshot.Version);
        Check.Equal(now, snapshot.Observation!.VersionObservedAt);
        Check.Equal(now.AddSeconds(2), snapshot.Observation.RetrievedAt);
        Check.Equal(snapshot.Observation.RetrievedAt, snapshot.ImportedAt);
        Check.Equal(2, snapshot.Entries.Count);
        var selected = EndpointQuery.Select(snapshot, new()
        {
            Domain = "edgecdn.example.test", Protocol = Transport.Tcp, Ports = "443",
            Fields = [FieldMatch.Equal("/extra/simulation/1", JsonSerializer.SerializeToElement(27))]
        });
        Check.Equal(1, selected.Hits.Single().Entry.Id);
        Check.True(snapshot.Entries[2].ExpressRoute is null);
        Check.Equal(1, f.Transport.Ids.Distinct().Count());
        f.Clock.Advance(TimeSpan.FromMinutes(5));
        var disk = f.Feed.ReadCache();
        Check.Equal(TimeSpan.FromMinutes(5), disk.Age);
        Check.Equal(snapshot.Observation, disk.Snapshot!.Observation);
        Check.True(JsonElement.DeepEquals(snapshot.Source, disk.Snapshot.Source));
        Check.Equal(CacheGeneration.Unavailable, WorldwideFeed.CreateMicrosoft(f.Root).ReadCache().Generation);
        await Check.Throws<FormatException>(() => Task.FromResult(SyntheticImporter.Import(Rows)));
        f.Transport.Drained();
    }

    private static async Task Schedule()
    {
        using var f = new Fixture();
        await f.Seed();
        f.Clock.Advance(TimeSpan.FromHours(1) - TimeSpan.FromTicks(1));
        Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
        f.Clock.Advance(TimeSpan.FromTicks(1));
        f.Transport.Reply("version", Json("""[{"instance":"Worldwide","latest":"2000010101"}]"""));
        var unchanged = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Unchanged, unchanged.Status);
        Check.Equal(f.Clock.GetUtcNow(), unchanged.LastVersionCheck);
        Check.Equal(TimeSpan.FromHours(1), unchanged.Cache.Age);
        Check.True(unchanged.Cache.Snapshot!.Observation!.VersionObservedAt < unchanged.LastVersionCheck);
        f.Hour();
        await f.Seed(Second);
        Check.Equal(Second, f.Feed.ReadCache().Snapshot!.Version);
        Check.Equal(First, f.Feed.ReadCache(previousOnly: true).Snapshot!.Version);
        f.Hour();
        await f.Seed(Third);
        Check.Equal(Second, f.Feed.ReadCache(previousOnly: true).Snapshot!.Version);
        Check.Equal(7, f.Transport.Ids.Count);
        Check.Equal(1, f.Transport.Ids.Distinct().Count());
        f.Transport.Drained();
    }

    private static async Task PublicationRace()
    {
        using var f = new Fixture();
        // The unversioned body already contains a new record when the version call observed First.
        var laterRows = Rows.Replace("\"id\":2", "\"id\":3");
        f.Transport.Pair(First, laterRows);
        var result = await f.Feed.RefreshAsync();
        Check.Equal(First, result.Cache.Snapshot!.Observation!.ObservedVersion);
        Check.True(result.Cache.Snapshot.Entries.ContainsKey(3));
        Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
        f.Hour();
        f.Transport.Pair(Second, laterRows);
        Check.Equal(Second, (await f.Feed.RefreshAsync()).Cache.Snapshot!.Observation!.ObservedVersion);
        Check.Equal(4, f.Transport.Ids.Count);
    }

    private static async Task Processes()
    {
        using var f = new Fixture();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        f.Transport.Queue("version", async token =>
        {
            entered.SetResult();
            await release.Task.WaitAsync(token);
            return Version(First);
        });
        f.Transport.Reply("endpoints", Json(Rows));
        var leader = f.Feed.RefreshAsync();
        await entered.Task;
        var follower = f.Feed.RefreshAsync();
        using var worker = StartWorker(f.Root, "deferred");
        Check.Equal("ready", await worker.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10)));
        Check.Equal("waiting", await worker.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10)));
        release.SetResult();
        Check.Equal(RefreshStatus.Updated, (await leader).Status);
        Check.Equal(RefreshStatus.Deferred, (await follower).Status);
        await Finish(worker);
        Check.Equal(2, f.Transport.Ids.Count);

        // A new OS process advances to the boundary, reads the first generation and uses the same GUID.
        using var restart = StartWorker(f.Root, "unchanged");
        Check.Equal("ready", await restart.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10)));
        string? persistedId = await restart.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10));
        Check.True(Guid.TryParse(persistedId, out var id) && id == f.Transport.Ids[0]);
        await Finish(restart);
        f.Hour();
        Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
        f.Transport.Drained();
    }

    internal static async Task CheckThrottleInNewProcess(string root)
    {
        using var worker = StartWorker(root, "throttled");
        Check.Equal("ready", await worker.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10)));
        await Finish(worker);
    }

    private static Process StartWorker(string root, string mode)
    {
        var start = new ProcessStartInfo(Environment.ProcessPath!) { RedirectStandardOutput = true };
        if (Path.GetFileNameWithoutExtension(start.FileName) == "dotnet")
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        foreach (string arg in new[] { "--worker", root, mode })
            start.ArgumentList.Add(arg);
        return Process.Start(start)!;
    }

    private static async Task Finish(Process worker)
    {
        await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
        Check.Equal(0, worker.ExitCode);
    }

    private static async Task<int> Worker(string root, string mode)
    {
        using var f = new Fixture(root);
        Console.WriteLine("ready");
        if (mode == "unchanged")
        {
            f.Hour();
            f.Transport.Reply("version", Version(First));
        }
        if (mode == "throttled")
            f.Clock.Advance(TimeSpan.FromHours(3)); // Past the hourly floor, inside the server's deadline.
        var refresh = f.Feed.RefreshAsync();
        if (mode == "deferred")
        {
            Check.True(!refresh.IsCompleted);
            Console.WriteLine("waiting");
        }
        var result = await refresh;
        if (mode == "throttled")
        {
            Check.Equal("HTTP 429.", result.Error);
            Check.Equal(f.Clock.GetUtcNow().AddHours(1), result.NextCheck);
        }
        Check.Equal(mode == "unchanged" ? RefreshStatus.Unchanged : RefreshStatus.Deferred, result.Status);
        Check.Equal(First, result.Cache.Snapshot!.Version);
        if (mode == "unchanged")
            Console.WriteLine(f.Transport.Ids.Single());
        return 0;
    }
}
