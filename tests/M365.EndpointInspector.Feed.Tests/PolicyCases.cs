using System.Net;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using static M365.EndpointInspector.Feed.Tests.SimulatedTransport;

namespace M365.EndpointInspector.Feed.Tests;

internal static class PolicyCases
{
    private static WorldwideFeed Dev(Fixture f) => new(f.Root, f.Client, f.Clock, RefreshPolicy.ManualDevelopment);

    internal static async Task Manual()
    {
        using var f = new Fixture();
        await f.Seed();
        var now = f.Clock.GetUtcNow();
        for (int click = 0; click < 2; click++)
        {
            f.Transport.Reply("version", Version(First));
            var result = await Dev(f).RefreshAsync();
            Check.Equal(RefreshStatus.Unchanged, result.Status);
            Check.Equal(now, result.NextCheck);
        }
        f.Transport.Pair(Second, Rows.Replace("80,443", "70000"));
        var invalid = await Dev(f).RefreshAsync();
        Check.Equal(RefreshStatus.Failed, invalid.Status);
        Check.Equal(First, invalid.Cache.Snapshot!.Version);
        f.Transport.Pair(Second);
        Check.Equal(RefreshStatus.Updated, (await Dev(f).RefreshAsync()).Status);
        Check.Equal(First, f.Feed.ReadCache(previousOnly: true).Snapshot!.Version);
        var release = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Deferred, release.Status);
        Check.Equal(now.AddHours(1), release.NextCheck);
        Check.Equal(8, f.Transport.Ids.Count);
        Check.Equal(1, f.Transport.Ids.Distinct().Count());
        f.Transport.Drained();
    }

    internal static async Task ServerLimits()
    {
        foreach (var policy in Enum.GetValues<RefreshPolicy>())
        foreach (var directive in new[] { "date", "delta", "absent", "max-age", "expires", "no-store" })
        {
            using var f = new Fixture();
            var now = f.Clock.GetUtcNow();
            var until = now.AddMinutes(directive == "absent" ? 60 : 7);
            bool limited = directive is "date" or "delta" or "absent";
            var reply = limited ? Json("", HttpStatusCode.TooManyRequests) : Json(Rows);
            switch (directive)
            {
                case "date": reply.Headers.RetryAfter = new(until); break;
                case "delta": reply.Headers.RetryAfter = new(until - now); break;
                case "max-age":
                    reply.Headers.CacheControl = new() { MaxAge = TimeSpan.FromMinutes(10) };
                    reply.Headers.Age = TimeSpan.FromMinutes(3);
                    break;
                case "no-store":
                    reply.Headers.CacheControl = new() { NoStore = true };
                    goto case "expires";
                case "expires": reply.Content.Headers.Expires = until; break;
            }
            f.Transport.Reply("version", Version(First));
            f.Transport.Reply("endpoints", reply);
            var result = await new WorldwideFeed(f.Root, f.Client, f.Clock, policy).RefreshAsync();
            Check.Equal(limited ? RefreshStatus.Throttled : directive == "no-store" ? RefreshStatus.Failed : RefreshStatus.Updated,
                result.Status);
            Check.Equal(limited || directive == "no-store" ? CacheGeneration.Unavailable : CacheGeneration.Active,
                result.Cache.Generation);
            Check.Equal(policy == RefreshPolicy.Release ? now.AddHours(1) : until, result.NextCheck);
            f.Clock.Advance(until - now - TimeSpan.FromTicks(1));
            var dev = await Dev(f).RefreshAsync();
            Check.Equal(RefreshStatus.Deferred, dev.Status);
            Check.Equal(until, dev.NextCheck);
            var release = await f.Feed.RefreshAsync();
            Check.Equal(RefreshStatus.Deferred, release.Status);
            Check.Equal(now.AddHours(1), release.NextCheck);
            f.Clock.Advance(TimeSpan.FromTicks(1));
            if (result.Cache.Snapshot == null) f.Transport.Pair();
            else f.Transport.Reply("version", Version(First));
            Check.Equal(result.Cache.Snapshot == null ? RefreshStatus.Updated : RefreshStatus.Unchanged,
                (await Dev(f).RefreshAsync()).Status);
            Check.Equal(1, f.Transport.Ids.Distinct().Count());
            f.Transport.Drained();
        }
    }

    internal static async Task RepeatedThrottle()
    {
        using var f = new Fixture();
        for (int attempt = 0; attempt < 2; attempt++)
        {
            // Some injected handlers surface the status directly, without a response object.
            if (attempt == 0) f.Transport.Queue("version", _ => throw new HttpRequestException("limited", null, HttpStatusCode.TooManyRequests));
            else f.Transport.Reply("version", Json("", HttpStatusCode.TooManyRequests));
            var result = await Dev(f).RefreshAsync();
            Check.Equal(RefreshStatus.Throttled, result.Status);
            var wait = TimeSpan.FromHours(attempt + 1);
            Check.Equal(f.Clock.GetUtcNow() + wait, result.NextCheck);
            Check.Equal(RefreshStatus.Deferred, (await Dev(f).RefreshAsync()).Status);
            Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
            f.Clock.Advance(wait);
        }
        f.Transport.Pair();
        Check.Equal(RefreshStatus.Updated, (await Dev(f).RefreshAsync()).Status);
        Check.Equal(4, f.Transport.Ids.Count);
    }

    internal static async Task ConcurrentCancellation()
    {
        using var f = new Fixture();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancel = new CancellationTokenSource();
        var version = Version(First);
        version.Headers.CacheControl = new() { MaxAge = TimeSpan.FromMinutes(6) };
        f.Transport.Reply("version", version);
        f.Transport.Queue("endpoints", async token =>
        {
            started.SetResult();
            await Task.Delay(Timeout.Infinite, token);
            throw new InvalidOperationException("Expected cancellation");
        });
        var leader = Dev(f).RefreshAsync(cancel.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        var devWaiter = Dev(f).RefreshAsync();
        var releaseWaiter = f.Feed.RefreshAsync();
        Check.True(!devWaiter.IsCompleted && !releaseWaiter.IsCompleted);
        using var cancelWaiter = new CancellationTokenSource();
        var waiting = Dev(f).RefreshAsync(cancelWaiter.Token);
        cancelWaiter.Cancel();
        await Check.Throws<OperationCanceledException>(() => waiting);
        cancel.Cancel();
        await Check.Throws<OperationCanceledException>(() => leader);
        foreach (var waiter in new[] { devWaiter, releaseWaiter })
            Check.Equal(RefreshStatus.Deferred, (await waiter).Status);
        Check.Equal(2, f.Transport.Ids.Count);
        Check.Equal(CacheGeneration.Unavailable, f.Feed.ReadCache().Generation);
        f.Clock.Advance(TimeSpan.FromMinutes(6));
        f.Transport.Pair();
        Check.Equal(RefreshStatus.Updated, (await Dev(f).RefreshAsync()).Status);
        Check.Equal(1, f.Transport.Ids.Distinct().Count());
    }

    // Construct exactly the original checked envelope, independently of the current state writer.
    private sealed record OriginalState(int Format, Guid ClientRequestId, DateTimeOffset LastVersionCheck,
        DateTimeOffset NextCheck, int Failures, string? LastError);

    internal static async Task Legacy()
    {
        const string parser = "endpoints[2]: urls[1]: Invalid DNS name or pattern.";
        foreach (var (error, hours, retest) in new (string?, int, bool)[]
        {
            (parser, 2, true), ("HTTP transport failed.", 2, true),
            ("HTTP 429.", 2, false), ("Too Many Requests", 2, false),
            ("Further requests deferred by Retry-After.", 2, false), (null, 2, false),
            (parser, 3, false), ("unrecognized outcome", 2, false)
        })
        {
            using var f = new Fixture();
            await f.Seed();
            var id = f.Transport.Ids[0];
            var now = f.Clock.GetUtcNow();
            var until = now.AddHours(hours).AddSeconds(2);
            var value = new OriginalState(1, id, now, until, error == null ? 0 : 2, error);
            var path = Path.Combine(f.DirectoryPath, "machine-state.json");
            string Digest(OriginalState s) => Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(s)));
            string original = JsonSerializer.Serialize(new { Sha256 = Digest(value), Value = value });
            File.WriteAllText(path, original);
            Check.Equal(until, (await f.Feed.RefreshAsync()).NextCheck);
            if (retest) f.Transport.Reply("version", Version(First));
            var result = await Dev(f).RefreshAsync();
            Check.Equal(retest ? RefreshStatus.Unchanged : RefreshStatus.Deferred, result.Status);
            Check.Equal(First, result.Cache.Snapshot!.Version);
            Check.Equal(until, (await f.Feed.RefreshAsync()).NextCheck);
            if (retest)
            {
                var envelope = JsonNode.Parse(File.ReadAllText(path))!;
                var legacyRead = envelope["Value"].Deserialize<OriginalState>()!;
                Check.Equal(id, legacyRead.ClientRequestId);
                Check.Equal(until, legacyRead.NextCheck);
                Check.Equal(Digest(legacyRead), envelope["Sha256"]!.GetValue<string>());
                envelope["Server"]!["Sha256"] = "damaged";
                File.WriteAllText(path, envelope.ToJsonString());
                Check.Equal(RefreshStatus.Failed, (await Dev(f).RefreshAsync()).Status);
            }
            else Check.Equal(original, File.ReadAllText(path));
            Check.Equal(retest ? 3 : 2, f.Transport.Ids.Count);
            Check.Equal(1, f.Transport.Ids.Distinct().Count());
            f.Transport.Drained();
        }
    }
}
