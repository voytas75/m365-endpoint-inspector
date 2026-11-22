using System.Net;
using System.Net.Http.Headers;
using System.Text;
using static M365.EndpointInspector.Feed.Tests.SimulatedTransport;

namespace M365.EndpointInspector.Feed.Tests;

internal static class FailureCases
{
    internal static async Task Throttling()
    {
        foreach (string header in new[] { "date", "delta", "absent" })
        {
            using var f = new Fixture();
            await f.Seed();
            f.Hour();
            var now = f.Clock.GetUtcNow();
            var limited = Json("", HttpStatusCode.TooManyRequests);
            var delay = TimeSpan.FromHours(header == "absent" ? 1 : 3);
            if (header == "date") limited.Headers.RetryAfter = new(now + delay);
            if (header == "delta") limited.Headers.RetryAfter = new(delay);
            if (header == "delta")
            {
                f.Transport.Reply("version", Version(Second));
                f.Transport.Reply("endpoints", limited);
            }
            else f.Transport.Reply("version", limited);
            var result = await f.Feed.RefreshAsync();
            Check.Equal(RefreshStatus.Throttled, result.Status);
            Check.Equal(now + delay, result.NextCheck);
            Check.Equal(First, result.Cache.Snapshot!.Version);
            if (header == "date")
                await Program.CheckThrottleInNewProcess(f.Root);
            f.Clock.Advance(delay - TimeSpan.FromTicks(1));
            var deferred = await f.Feed.RefreshAsync();
            Check.Equal(RefreshStatus.Deferred, deferred.Status);
            Check.Equal("HTTP 429.", deferred.Error);
            f.Clock.Advance(TimeSpan.FromTicks(1));
            f.Transport.Reply("version", Json("", HttpStatusCode.TooManyRequests));
            result = await f.Feed.RefreshAsync();
            Check.Equal(f.Clock.GetUtcNow().AddHours(2), result.NextCheck);
            f.Hour();
            Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
            f.Hour();
            await f.Seed(Second);
            Check.Equal(1, f.Transport.Ids.Distinct().Count());
            f.Transport.Drained();
        }
    }

    internal static async Task Headers()
    {
        using var f = new Fixture();
        var now = f.Clock.GetUtcNow();
        var version = Version(First);
        version.Headers.CacheControl = new CacheControlHeaderValue { MaxAge = TimeSpan.FromHours(8) };
        version.Headers.Age = TimeSpan.FromHours(2);
        f.Transport.Reply("version", version);
        var endpoints = Json(Rows);
        endpoints.Content.Headers.Expires = now.AddHours(4);
        f.Transport.Reply("endpoints", endpoints);
        Check.Equal(now.AddHours(6), (await f.Feed.RefreshAsync()).NextCheck);
        f.Clock.Advance(TimeSpan.FromHours(6));
        var blocked = Json(Rows);
        blocked.Headers.CacheControl = new CacheControlHeaderValue { NoStore = true };
        blocked.Content.Headers.Expires = now.AddHours(9);
        f.Transport.Reply("version", Version(Second));
        f.Transport.Reply("endpoints", blocked);
        var result = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Failed, result.Status);
        Check.Equal(now.AddHours(9), result.NextCheck);
        Check.Equal(First, result.Cache.Snapshot!.Version);
        f.Hour();
        Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
        f.Clock.Advance(TimeSpan.FromHours(2));
        var pause = Version(Second);
        pause.Headers.RetryAfter = new(TimeSpan.FromHours(5));
        f.Transport.Reply("version", pause);
        var held = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Deferred, held.Status);
        Check.Equal(f.Clock.GetUtcNow().AddHours(5), held.NextCheck);
        Check.Equal(First, held.Cache.Snapshot!.Version);
        f.Transport.Drained();
    }

    internal static async Task Schema()
    {
        foreach (string bad in new[] { "{", "null", "[]", "{}", """{"instance":"Other","latest":"2000010101"}""",
                     """{"instance":"Worldwide","latest":"\uD800"}""", """{"instance":"Worldwide","latest":2000010101}""", """{"instance":"Worldwide","latest":"2000999900"}""",
                     """{"instance":"Worldwide","latest":"2000010101","latest":"2000010201"}""" })
        {
            using var f = new Fixture();
            f.Transport.Reply("version", Json(bad));
            var result = await f.Feed.RefreshAsync();
            Check.Equal(RefreshStatus.Failed, result.Status);
            Check.Equal(CacheGeneration.Unavailable, result.Cache.Generation);
            Check.True(!string.IsNullOrWhiteSpace(result.Error));
            Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
        }
        foreach (string bad in new[] { "{", "[]", "null", "[{}]", Rows.Replace("\"id\":2", "\"id\":1"),
                     Rows.Replace("\"id\":2", "\"id\":0"), Rows.Replace("\"category\":\"Allow\"", "\"category\":42"),
                     Rows.Replace("\"tcpPorts\":\"80,443\"", "\"required\":\"true\""),
                     Rows.Replace("192.0.2.0/24", "192.0.2.0/99"), Rows.Replace("27]", """27,"\uDC00"]"""), Rows.Replace("true,27", """{"x":1,"x":2}"""),
                     """{"provenance":"Synthetic","version":"synthetic-test","endpoints":[]}""" })
        {
            using var f = new Fixture();
            await f.Seed();
            f.Hour();
            f.Transport.Pair(Second, bad);
            var result = await f.Feed.RefreshAsync();
            Check.Equal(RefreshStatus.Failed, result.Status);
            Check.Equal(First, result.Cache.Snapshot!.Version);
            Check.Equal(TimeSpan.FromHours(1), result.Cache.Age);
            Check.True(!string.IsNullOrWhiteSpace(result.Error));
            Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
            f.Transport.Drained();
        }
    }

    private sealed class ForwardOnlyStream(byte[] bytes) : MemoryStream(bytes)
    {
        public override bool CanSeek => false;
    }

    private sealed class StalledBody(TestClock clock) : MemoryStream
    {
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
        {
            clock.Advance(TimeSpan.FromSeconds(31));
            Check.True(token.IsCancellationRequested);
            return ValueTask.FromCanceled<int>(token);
        }
    }

    internal static async Task Body()
    {
        foreach (string problem in new[] { "declared", "chunked", "partial", "length", "timeout", "transport", "utf8" })
        {
            using var f = new Fixture();
            await f.Seed();
            f.Hour();
            f.Transport.Reply("version", Version(Second));
            f.Transport.Queue("endpoints", _ =>
            {
                if (problem == "transport") throw new HttpRequestException("Simulated connection loss");
                var response = Json(Rows);
                switch (problem)
                {
                    case "timeout": response.Content = new StreamContent(new StalledBody(f.Clock)); break;
                    case "utf8":
                        var bytes = Encoding.UTF8.GetBytes(Rows);
                        bytes[Array.IndexOf(bytes, (byte)'*')] = 0xff;
                        response.Content = new ByteArrayContent(bytes);
                        break;
                    case "declared": response.Content.Headers.ContentLength = 50_000_000; break;
                    case "chunked":
                        response.Content = new StreamContent(new ForwardOnlyStream(new byte[WorldwideFeed.EndpointBodyLimit + 1]));
                        break;
                    case "partial": response.StatusCode = HttpStatusCode.PartialContent; break;
                    case "length": response.Content.Headers.ContentLength = Encoding.UTF8.GetByteCount(Rows) + 1; break;
                }
                return Task.FromResult(response);
            });
            var result = await f.Feed.RefreshAsync();
            Check.Equal(problem == "timeout" ? RefreshStatus.TimedOut : RefreshStatus.Failed, result.Status);
            Check.Equal(First, result.Cache.Snapshot!.Version);
            Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
            Check.Equal(4, f.Transport.Ids.Count);
            f.Transport.Drained();
        }
    }

    internal static async Task Cancellation()
    {
        using var f = new Fixture();
        await f.Seed();
        f.Hour();
        using var cancel = new CancellationTokenSource();
        await Check.Throws<OperationCanceledException>(() => f.Feed.RefreshAsync(new CancellationToken(true)));
        f.Transport.Reply("version", Version(Second));
        f.Transport.Queue("endpoints", token =>
        {
            cancel.Cancel();
            return Task.FromCanceled<HttpResponseMessage>(token);
        });
        await Check.Throws<OperationCanceledException>(() => f.Feed.RefreshAsync(cancel.Token));
        var result = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Deferred, result.Status);
        Check.True(!string.IsNullOrEmpty(result.Error));
        Check.Equal(First, result.Cache.Snapshot!.Version);
        Check.Equal(4, f.Transport.Ids.Count);
        f.Transport.Drained();
    }
}
