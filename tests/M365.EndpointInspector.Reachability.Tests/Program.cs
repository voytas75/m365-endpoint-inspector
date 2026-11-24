using M365.EndpointInspector.Presentation;
using M365.EndpointInspector.Domain;

static void Check(bool condition)
{
    if (!condition) throw new InvalidOperationException("Reachability assertion failed.");
}

Check(ReachabilityTarget.Parse(" Example.TEST. ")?.Host == "example.test");
Check(ReachabilityTarget.Parse("https://example.test:8443/probe%20path") is { Host: "example.test", Port: 8443, HttpsUri.AbsolutePath: "/probe%20path", HttpsUri.Query: "" });
Check(ReachabilityTarget.Parse("https://[::1]:8443/probe") is { Authority: "[::1]:8443", IsAddress: true });
Check(ReachabilityTarget.Parse("example.test") is { Port: 443, HttpsUri.AbsolutePath: "/" });
foreach (var invalid in new[] { "*.test", "http://example.test", "https://user@example.test", "https://example.test/?token=secret", "https://example.test/#secret", "https://127.1", "a:443", "10.0.0.0/24", "127.1", "a@b" })
    Check(ReachabilityTarget.Parse(invalid) == null);
Check(ReachabilityTarget.Parse("2001:0db8::1")?.Host == "2001:db8::1");
Console.WriteLine("PASS concrete target validation");

var target = ReachabilityTarget.Parse("localhost")!;
var rows = new List<ProbeResult>();
var fake = new ReachabilityProbe(async (_, stage, token) =>
{
    if (stage == ProbeStage.Dns) throw new System.Net.Sockets.SocketException(11001);
    if (stage == ProbeStage.DirectTcp) await Task.Delay(Timeout.Infinite, token);
    return new(ProbeStatus.Restricted, ProbeCode.HttpResponse, ProbeRoute.Proxy, 403);
}, TimeSpan.FromMilliseconds(100));
await fake.RunAsync(target, rows.Add, default);
Check(rows.Select(r => r.Status).SequenceEqual(new[] { ProbeStatus.Failed, ProbeStatus.TimedOut, ProbeStatus.Restricted }));
Console.WriteLine("PASS independent failure, deadline and HTTP response outcomes");

var lifecycle = new List<string>();
var lifecycleProbe = new ReachabilityProbe((_, stage, _) =>
{
    lifecycle.Add($"execute {stage}");
    return Task.FromResult(new ProbeObservation(ProbeStatus.Passed, ProbeCode.Connected));
}, TimeSpan.FromSeconds(1));
await ReachabilityBatch.FromValues(["127.0.0.1"]).RunAsync(lifecycleProbe,
    row => lifecycle.Add($"result {row.Stage}"), default,
    admitted: _ => lifecycle.Add("admitted"),
    activity: (_, stage) => lifecycle.Add(stage is null ? "inactive" : $"active {stage}"),
    finished: (_, done) => lifecycle.Add($"done {done}"));
Check(lifecycle.SequenceEqual(new[] { "admitted", "result Dns", "active DirectTcp", "execute DirectTcp",
    "inactive", "result DirectTcp", "active HttpsSystemProxy", "execute HttpsSystemProxy", "inactive",
    "result HttpsSystemProxy", "done True" }));
Console.WriteLine("PASS actual stage lifecycle, literal DNS skip and target completion");

var batch = ReachabilityBatch.FromValues(new[] { "EXAMPLE.test.", "example.test", "*.test", "192.0.2.0/24", "192.0.2.1" });
Check(batch.Targets.Count == 2 && batch.Skipped == 2 && batch.Duplicates == 1);

using var fixture = System.Text.Json.JsonDocument.Parse("""
[
  {"id":1,"serviceArea":"Exchange","category":"Allow","urls":["One.test","*.test"]},
  {"id":2,"serviceArea":"Exchange","category":"Default","urls":["one.test","two.test"],"ips":["192.0.2.1","192.0.2.0/24"]},
  {"id":3,"serviceArea":"Common","category":"Allow","urls":["other.test"]}
]
""");
var snapshot = M365.EndpointInspector.Domain.WorldwideImporter.ImportFile(fixture.RootElement, DateTimeOffset.UnixEpoch);
var areaBatch = ReachabilityBatch.FromArea(snapshot, M365.EndpointInspector.Domain.ServiceArea.Exchange);
Check(areaBatch.Targets.Select(t => t.Host).SequenceEqual(new[] { "one.test", "two.test", "192.0.2.1" }));
Check(areaBatch.Skipped == 2 && areaBatch.Duplicates == 1);
Check(areaBatch.Summary == "3 target(s), 2 skipped, 1 duplicate(s).");
Check(areaBatch.SkippedTargets.SequenceEqual(new SkippedReachabilityTarget[] {
    new("*.test", "Wildcard not expanded"), new("192.0.2.0/24", "CIDR range not expanded")
}));
BrowserRow BrowserEntry(int id) => new(new EndpointMatch(snapshot.Entries.Values.Single(entry => entry.Id == id), null, []));
var selectedValues = ReachabilityBatch.SelectedValues([BrowserEntry(1)], [BrowserEntry(3)]).ToArray();
Check(selectedValues.SequenceEqual(new[] { "One.test", "*.test", "other.test" }));
Check(ReachabilityBatch.FromValues(selectedValues.TakeLast(1)).Targets.Single().Host == "other.test");
Check(!ReachabilityBatch.SelectedValues([], []).Any());
Console.WriteLine("PASS scope preview, skip reasons and selected-row projection/refinement");
using var stop = new CancellationTokenSource(TimeSpan.FromSeconds(5));
int started = 0;
var reports = new System.Collections.Concurrent.ConcurrentBag<ProbeResult>();
var ready = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
var blocking = new ReachabilityProbe(async (_, stage, token) =>
{
    if (stage == ProbeStage.Dns) return new(ProbeStatus.Passed, ProbeCode.Resolved);
    if (Interlocked.Increment(ref started) == ReachabilityBatch.Parallelism) ready.SetResult();
    await Task.Delay(Timeout.Infinite, token);
    throw new InvalidOperationException("Expected cancellation");
}, TimeSpan.FromSeconds(10));
var ended = new System.Collections.Concurrent.ConcurrentBag<bool>();
var running = ReachabilityBatch.FromValues(Enumerable.Range(0, 12).Select(i => $"host{i}.test"))
    .RunAsync(blocking, reports.Add, stop.Token, finished: (_, done) => ended.Add(done));
await ready.Task.WaitAsync(stop.Token);
Check(!running.IsCompleted && reports.Count == ReachabilityBatch.Parallelism);
stop.Cancel();
try { await running; throw new InvalidOperationException("Expected cancellation"); }
catch (OperationCanceledException) when (stop.IsCancellationRequested) { }
Check(started == ReachabilityBatch.Parallelism && reports.Count(r => r.Status == ProbeStatus.Cancelled) == started);
Check(ended.Count == ReachabilityBatch.Parallelism && ended.All(done => !done));
Console.WriteLine("PASS loaded service area scope, skips, deduplication, live results and bounded cancellation");

var mixed = rows.Concat(new[] {
    new ProbeResult(ReachabilityTarget.Parse("untested.test")!, ProbeStage.Dns, new(ProbeStatus.Cancelled, ProbeCode.UserCancelled), TimeSpan.Zero),
    new ProbeResult(ReachabilityTarget.Parse("good.test")!, ProbeStage.Dns, new(ProbeStatus.Passed, ProbeCode.Resolved), TimeSpan.Zero)
}).ToArray();
Check(ReachabilityReport.Failures(mixed).Targets.SequenceEqual(new[] { target }));
var text = ReachabilityReport.CopyText(mixed);
Check(text.Contains("localhost:443\tDNS\tFailed") && text.Contains("good.test:443\tDNS\tPassed") && !text.Contains("secret"));
var interruptedRetry = mixed.Append(rows[0] with { Attempt = 2, Observation = new(ProbeStatus.Cancelled, ProbeCode.UserCancelled) }).ToArray();
Check(ReachabilityReport.Failures(interruptedRetry).Targets.Single() == target);
var recovered = interruptedRetry.Concat(rows.Take(2).Select(row => row with { Attempt = 3, Observation = new(ProbeStatus.Passed, ProbeCode.Connected) }));
Check(ReachabilityReport.Failures(recovered).Targets.Count == 0);
var exported = ReachabilityReport.CopyText(interruptedRetry.Reverse()).Split(Environment.NewLine);
Check(exported[1].Contains("\tFailed\t1\t") || exported[1].Contains("\tTimedOut\t1\t"));
Check(exported.Any(line => line.Contains("\tCancelled\t2\t") && line.EndsWith("\tUserCancelled")));
Check(exported.Any(line => line.Contains("\tProxy\t") && line.EndsWith("\tHttpResponse HTTP 403")));
Console.WriteLine("PASS retry selection and sanitized clipboard projection");

using var stopAfterDns = new CancellationTokenSource();
var nativeRows = new List<ProbeResult>();
await new ReachabilityProbe().RunAsync(target, row =>
{
    nativeRows.Add(row);
    stopAfterDns.Cancel();
}, stopAfterDns.Token);
Check(nativeRows is [{ Stage: ProbeStage.Dns, Status: ProbeStatus.Passed }]);
Console.WriteLine("PASS native localhost DNS; cancelled before TCP or HTTPS");

await NativeCases.RunAsync();
