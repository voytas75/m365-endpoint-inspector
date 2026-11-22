using System.Net;
using System.Text;

namespace M365.EndpointInspector.Feed.Tests;

// All messages, endpoint records and times in this harness are simulated. No socket handler is installed.
internal sealed class SimulatedTransport : HttpMessageHandler
{
    internal const string First = "2000010101";
    internal const string Second = "2000010201";
    internal const string Third = "2000010301";
    internal const string Rows = """
        [{"id":1,"serviceArea":"Common","category":"Allow","urls":["*cdn.example.test"],
          "tcpPorts":"80,443","extra":{"simulation":[true,27],"text":"zażółć + <tag> & ©"}},
         {"id":2,"serviceArea":"Exchange","category":"Optimize","ips":["192.0.2.0/24"]}]
        """;

    private readonly Queue<(string Resource, Func<CancellationToken, Task<HttpResponseMessage>> Reply)> replies = new();
    internal List<Guid> Ids { get; } = [];

    internal static HttpResponseMessage Json(string body, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };

    internal static HttpResponseMessage Version(string value) =>
        Json($$"""{"instance":"Worldwide","latest":"{{value}}","unknown":"simulated"}""");

    internal void Queue(string resource, Func<CancellationToken, Task<HttpResponseMessage>> reply) =>
        replies.Enqueue((resource, reply));

    internal void Reply(string resource, HttpResponseMessage response) =>
        Queue(resource, _ => Task.FromResult(response));

    internal void Pair(string value = First, string rows = Rows)
    {
        Reply("version", Version(value));
        Reply("endpoints", Json(rows));
    }

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
    {
        (string resource, Func<CancellationToken, Task<HttpResponseMessage>> reply) = replies.Dequeue();
        var uri = request.RequestUri!;
        const string prefix = "?ClientRequestId=";
        Check.Equal(HttpMethod.Get, request.Method);
        Check.Equal("https://endpoints.office.com", uri.GetLeftPart(UriPartial.Authority));
        Check.Equal("/" + resource + "/Worldwide", uri.AbsolutePath);
        Check.True(uri.Query.StartsWith(prefix, StringComparison.Ordinal));
        var id = Guid.ParseExact(uri.Query[prefix.Length..], "D");
        Check.True(id != Guid.Empty);
        Ids.Add(id);
        return reply(token);
    }

    internal void Drained() => Check.Equal(0, replies.Count);
}

internal sealed class TestClock : TimeProvider
{
    private DateTimeOffset now = DateTimeOffset.Parse("2001-01-01T00:00:00Z");
    public override DateTimeOffset GetUtcNow() => now;
    private readonly List<Deadline> deadlines = [];
    public override ITimer CreateTimer(TimerCallback callback, object? state, TimeSpan dueTime, TimeSpan period)
    {
        var deadline = new Deadline(this, () => callback(state));
        deadline.Change(dueTime, period);
        return deadline;
    }

    internal void Advance(TimeSpan elapsed)
    {
        now += elapsed;
        foreach (var deadline in deadlines.ToArray())
            deadline.FireIfDue(now);
    }

    private sealed class Deadline(TestClock clock, Action fire) : ITimer
    {
        private DateTimeOffset due;
        public bool Change(TimeSpan delay, TimeSpan period)
        {
            Check.Equal(Timeout.InfiniteTimeSpan, period);
            clock.deadlines.Remove(this);
            if (delay != Timeout.InfiniteTimeSpan)
            {
                due = clock.now + delay;
                clock.deadlines.Add(this);
            }
            return true;
        }

        internal void FireIfDue(DateTimeOffset time)
        {
            if (time < due) return;
            Dispose();
            fire();
        }

        public void Dispose() => clock.deadlines.Remove(this);
        public ValueTask DisposeAsync() { Dispose(); return ValueTask.CompletedTask; }
    }
}

internal sealed class Fixture : IDisposable
{
    internal string Root { get; }
    internal string DirectoryPath => Path.Combine(Root, "SimulatedWorldwide");
    internal TestClock Clock { get; } = new();
    internal SimulatedTransport Transport { get; } = new();
    internal HttpClient Client { get; }
    internal WorldwideFeed Feed => new(Root, Client, Clock);

    internal Fixture(string? root = null)
    {
        var repo = new DirectoryInfo(AppContext.BaseDirectory);
        while (repo != null && !File.Exists(Path.Combine(repo.FullName, "M365.EndpointInspector.slnx")))
            repo = repo.Parent;
        Root = root ?? Path.Combine(repo!.FullName, "artifacts", "feed-tests", Guid.NewGuid().ToString("N"));
        Client = new HttpClient(Transport);
    }

    public void Dispose() => Client.Dispose();
    internal void Hour() => Clock.Advance(TimeSpan.FromHours(1));
    internal async Task Seed(string version = SimulatedTransport.First)
    {
        Transport.Pair(version);
        Check.Equal(RefreshStatus.Updated, (await Feed.RefreshAsync()).Status);
    }
}

internal static class Check
{
    public static void True(bool value)
    {
        if (!value) throw new InvalidOperationException("Assertion failed.");
    }

    public static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException($"Expected {expected}, got {actual}.");
    }

    public static async Task Throws<T>(Func<Task> action) where T : Exception
    {
        try { await action(); }
        catch (T) { return; }
        throw new InvalidOperationException($"Expected {typeof(T).Name}.");
    }
}
