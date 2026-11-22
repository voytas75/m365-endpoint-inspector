using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed;
using M365.EndpointInspector.Feed.Tests;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class WorkspaceCases
{
    private static readonly CancellationToken Token = CancellationToken.None;

    internal static async Task<int> Run(string root)
    {
        (string Name, Func<string, Task> Test)[] groups =
        [
            ("explicit simulated download, cutoff, failure, fallback and restored age", Download),
            ("manual DEV retest and honest no-request notice after restart", ManualNotice),
            ("cancel in flight, close and reopen offline", Cancel),
            ("raw import, checked snapshot and Synthetic isolation", Imports),
            ("reported wildcard in simulated download and raw file", WildcardSources),
            ("workspace retains applied/draft, selected membership, reference and source files", RoundTrip),
            ("debounce, closing flush and conflicting workspace writers", WriteOrder),
            ("corrupt or unsupported workspace, fallback and failed save", Recovery),
            ("path links, access failure and bounded input", Limits)
        ];
        foreach (var (name, test) in groups)
        {
            string directory = Directory.CreateDirectory(Path.Combine(root, name.Split(',')[0])).FullName;
            await test(directory);
            Console.WriteLine($"PASS {name}");
        }
        return groups.Length;
    }

    private static WorkspaceStore Store(string root, TimeProvider clock) => new(Path.Combine(root, "state"), clock);
    private static string Active(string root) => Path.Combine(root, "state", "workspace.json");
    private static string Previous(string root) => Path.Combine(root, "state", "workspace.previous.json");

    private static async Task Download(string root)
    {
        using var f = new Fixture(Path.Combine(root, "machine"));
        Task<BrowserDocument> Click() => BrowserDocument.DownloadAsync(f.Feed, f.Clock, Token);
        f.Transport.Pair();
        var first = await Click();
        Check.Equal(RefreshStatus.Updated, first.Origin!.Status);
        Check.True(first.SourceLabel.Contains("Internet") && first.SourceLabel.Contains("SimulatedWorldwide"));
        Check.True(first.Metadata.Contains("The version was checked before content retrieval"));
        Check.Equal(f.Clock.GetUtcNow().AddHours(1), first.Origin.NextCheck);
        f.Clock.Advance(TimeSpan.FromMinutes(59));
        var held = await Click();
        Check.Equal(RefreshStatus.Deferred, held.Origin!.Status);
        Check.Equal(2, f.Transport.Ids.Count);
        Check.True(held.Notice!.Contains("Check deferred") && held.Notice.Contains("01:00:00 UTC"));
        f.Clock.Advance(TimeSpan.FromMinutes(1));
        f.Transport.Reply("version", SimulatedTransport.Version(SimulatedTransport.First));
        var unchanged = await Click();
        Check.Equal(RefreshStatus.Unchanged, unchanged.Origin!.Status);
        Check.Equal(first.Snapshot!.ImportedAt, unchanged.Snapshot!.ImportedAt);
        f.Hour();
        f.Transport.Pair(SimulatedTransport.Second);
        await Click();
        File.WriteAllText(Path.Combine(f.DirectoryPath, "active.json"), "{");
        f.Hour();
        var throttle = SimulatedTransport.Json("", HttpStatusCode.TooManyRequests);
        throttle.Headers.RetryAfter = new(TimeSpan.FromHours(3));
        f.Transport.Reply("version", throttle);
        var fallback = await Click();
        Check.Equal(RefreshStatus.Throttled, fallback.Origin!.Status);
        Check.Equal(CacheGeneration.Previous, fallback.Origin.Generation);
        Check.True(fallback.Notice!.Contains("429") && !fallback.Notice.Contains("active:") && fallback.Details!.Contains("active:"));
        Check.True(fallback.Metadata.Contains("3 h") && fallback.SourceLabel.Contains("previous copy"));
        Check.Equal(f.Clock.GetUtcNow().AddHours(3), fallback.Origin.NextCheck);
        var store = Store(root, f.Clock);
        store.Load();
        store.Save(BrowserWorkspace.Empty.Open(fallback));
        string persisted = File.ReadAllText(Active(root));
        Check.True(!persisted.Contains(f.Transport.Ids[0].ToString(), StringComparison.OrdinalIgnoreCase)
            && !persisted.Contains("ClientRequestId") && !persisted.Contains("Retry-After"));
        f.Clock.Advance(TimeSpan.FromDays(2));
        var restored = Store(root, f.Clock).Load().Workspace.Document;
        Check.Equal(fallback.Origin, restored.Origin);
        Check.True(restored.Metadata.Contains("2 d 3 h") && restored.Notice == fallback.Notice);
        Check.Equal(6, f.Transport.Ids.Count);
        Check.Equal(1, f.Transport.Ids.Distinct().Count());
        f.Transport.Drained();
    }

    private static async Task ManualNotice(string root)
    {
        using var f = new Fixture(root);
        var dev = new WorldwideFeed(root, f.Client, f.Clock, RefreshPolicy.ManualDevelopment);
        Task<BrowserDocument> Click() => BrowserDocument.DownloadAsync(dev, f.Clock, Token);
        var store = Store(root, f.Clock);
        store.Load();
        f.Transport.Pair(rows: SimulatedTransport.Rows.Replace("80,443", "70000"));
        try { await Click(); throw new Exception("Expected validation failure"); }
        catch (InvalidDataException ex)
        {
            Check.True(ex.Message.Contains("validation") && !ex.Message.Contains("active:"));
            Check.True(ex.InnerException!.Message.Contains("active:"));
        }
        try { await BrowserDocument.DownloadAsync(f.Feed, f.Clock, Token); throw new Exception("Expected Release deferral"); }
        catch (InvalidDataException ex)
        {
            Check.True(ex.Message.Contains("No request sent.") && ex.Message.Contains("— data validation error"));
            Check.True(ex.InnerException!.Message.Contains("Previous attempt") && !ex.Message.Contains("active:"));
        }
        var throttle = SimulatedTransport.Json("", HttpStatusCode.TooManyRequests);
        throttle.Headers.RetryAfter = new(TimeSpan.FromMinutes(5));
        f.Transport.Reply("version", throttle);
        await Check.Throws<InvalidDataException>(Click);
        f.Clock.Advance(TimeSpan.FromMinutes(1));
        try { await Click(); throw new Exception("Expected deferral without cache"); }
        catch (InvalidDataException ex)
        {
            Check.True(ex.Message.Contains("No request sent.") && ex.Message.Contains("00:05:00 UTC"));
            Check.True(ex.Message.Contains("Previous attempt: 2001-01-01 00:00:00 UTC — request throttling (HTTP 429)"));
            Check.True(!ex.Message.Contains("active:") && ex.InnerException!.Message.Contains("HTTP 429."));
        }
        Check.Equal(3, f.Transport.Ids.Count);
        f.Clock.Advance(TimeSpan.FromMinutes(4));
        f.Transport.Pair();
        var first = await Click();
        store.Save(BrowserWorkspace.Empty.Open(first));
        var reopened = Store(root, f.Clock).Load();
        Check.Equal(first.Origin, reopened.Workspace.Document.Origin);
        Check.Equal(5, f.Transport.Ids.Count);
        var version = SimulatedTransport.Version(SimulatedTransport.Second);
        version.Headers.RetryAfter = new(TimeSpan.FromMinutes(2));
        f.Transport.Reply("version", version);
        var partial = await Click();
        Check.True(partial.Notice!.Contains("Version checked.") && !partial.Notice.Contains("No request sent"));
        var held = await Click();
        Check.True(held.Notice!.Contains("No request sent.") && held.Details!.Contains("Previous attempt"));
        Check.Equal(6, f.Transport.Ids.Count);
        f.Transport.Drained();
    }

    private static async Task Cancel(string root)
    {
        using var f = new Fixture(Path.Combine(root, "machine"));
        var store = Store(root, f.Clock);
        var previous = store.Load().Workspace.Open(OpenFixture("synthetic-2", f.Clock)).Present().Workspace;
        var writer = new WorkspaceWriter(store, TimeSpan.FromHours(1));
        var queued = writer.ScheduleAsync(previous);
        using var cancel = new CancellationTokenSource();
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        f.Transport.Queue("version", async token =>
        {
            started.SetResult();
            await Task.Delay(Timeout.Infinite, token);
            throw new InvalidOperationException("Unexpected completion.");
        });
        var inFlight = BrowserDocument.DownloadAsync(f.Feed, f.Clock, cancel.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancel.Cancel();
        await Check.Throws<OperationCanceledException>(() => inFlight);
        await writer.FlushAsync();
        await queued;
        var reopened = Store(root, f.Clock).Load().Workspace;
        Check.Equal(previous.SelectedId, reopened.SelectedId);
        Check.True(reopened.Document.IsSynthetic && reopened.Document.Origin == previous.Document.Origin);
        Check.Equal(1, f.Transport.Ids.Count);
        // Even a new explicit click must obey the cancelled attempt's reservation.
        await Check.Throws<InvalidDataException>(() => BrowserDocument.DownloadAsync(f.Feed, f.Clock, Token));
        Check.Equal(1, f.Transport.Ids.Count);
        f.Transport.Drained();
    }

    private static BrowserDocument OpenFixture(string name, TimeProvider clock) =>
        BrowserDocument.OpenFile(Path.Combine(AppContext.BaseDirectory, "Fixtures", name + ".json"), clock, Token);

    private static async Task Imports(string root)
    {
        using var f = new Fixture(Path.Combine(root, "machine"));
        string path = Path.Combine(root, "synthetic-official-shaped.json");
        File.WriteAllText(path, SimulatedTransport.Rows, new UTF8Encoding(true));
        var file = BrowserDocument.OpenFile(path, f.Clock, Token);
        Check.Equal("File", file.Snapshot!.Provenance);
        Check.True(file.Snapshot.Version == null && file.Snapshot.Observation == null && !file.IsSynthetic);
        Check.True(file.Metadata.Contains("unknown") && file.Origin!.Path == path);
        Check.True(file.Snapshot.Entries[1].Source.GetProperty("extra").GetProperty("simulation")[0].GetBoolean());
        var original = BrowserWorkspace.Empty.Open(file).Present().Workspace;
        var displayed = original;
        foreach (string invalid in new[]
        {
            "[]", """{"provenance":"MicrosoftWorldwide","endpoints":[]}""",
            SimulatedTransport.Rows.Replace("\"id\":1", "\"id\":\"1\""),
            SimulatedTransport.Rows.Replace("\"text\":", "\"simulation\":"),
            """{"Format":1,"Current":null}"""
        })
        {
            File.WriteAllText(path, invalid);
            try { displayed = displayed.Open(BrowserDocument.OpenFile(path, f.Clock, Token)); }
            catch (Exception ex) when (WorkspaceStore.IsDataError(ex)) { }
            Check.True(ReferenceEquals(original, displayed) && displayed.SelectedId == 2);
        }
        var fixture = OpenFixture("synthetic-1", f.Clock);
        Check.True(fixture.IsSynthetic && fixture.SourceLabel.Contains("Synthetic"));
        File.Delete(path);
        var store = Store(root, f.Clock);
        store.Load();
        store.Save(original);
        f.Clock.Advance(TimeSpan.FromDays(3));
        var copy = Store(root, f.Clock).Load().Workspace;
        Check.True(copy.Document.Origin == file.Origin && copy.Document.Metadata.Contains("Since import 3 d"));
        Check.Equal(file.Snapshot.Source.GetRawText(), copy.Document.Snapshot!.Source.GetRawText());

        foreach (var row in BrowserResults.Select(copy.Document, new()).Hits)
            Program.AssertReadablePreview(file.Snapshot.Entries[row.Entry.Id].Source, row.RawJson);
        Check.Equal(file.Snapshot.Source.GetRawText(), copy.Document.Snapshot.Source.GetRawText());

        await f.Seed();
        f.Hour();
        string generation = Path.Combine(f.DirectoryPath, "active.json");
        var declared = BrowserDocument.OpenFile(generation, f.Clock, Token);
        Check.Equal("File", declared.Snapshot!.Provenance);
        Check.True(!declared.IsSynthetic && declared.Metadata.Contains("Declared by file (SimulatedWorldwide)") &&
            declared.Metadata.Contains("1 h") && declared.Snapshot.Observation!.RetrievedAt < declared.Origin!.OpenedAt);
        store.Save(BrowserWorkspace.Empty.Open(declared));
        var restored = Store(root, f.Clock).Load().Workspace.Document;
        Check.Equal(declared.Snapshot.Observation, restored.Snapshot!.Observation);
        Check.Equal(BrowserSource.File, restored.Origin!.Kind);
        File.WriteAllText(generation, File.ReadAllText(generation).Replace("SimulatedWorldwide", "MicrosoftWorldwide"));
        await Check.Throws<FormatException>(() => Task.FromResult(BrowserDocument.OpenFile(generation, f.Clock, Token)));
        // Escaping non-ASCII values can make the persisted generation larger than the wire limit.
        string escapedRows = SimulatedTransport.Rows.Replace("zażółć", new string('ó', 720000));
        f.Transport.Pair(SimulatedTransport.Second, escapedRows);
        var large = await BrowserDocument.DownloadAsync(f.Feed, f.Clock, Token);
        Check.True(large.Origin!.Status == RefreshStatus.Updated && new FileInfo(generation).Length > 4 * 1024 * 1024);
        Check.Equal(SimulatedTransport.Second, BrowserDocument.OpenFile(generation, f.Clock, Token).Snapshot!.Version);
        f.Transport.Drained();
    }

    private static async Task WildcardSources(string root)
    {
        using var f = new Fixture(Path.Combine(root, "machine"));
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "simulated-worldwide-wildcard.json");
        string rows = File.ReadAllText(path);
        f.Transport.Pair(rows: rows);
        var download = await BrowserDocument.DownloadAsync(f.Feed, f.Clock, Token);
        Check.Equal(RefreshStatus.Updated, download.Origin!.Status);
        Check.Equal(BrowserSource.Internet, download.Origin.Kind);
        Check.Equal("SimulatedWorldwide", download.Snapshot!.Provenance);
        Check.Equal(CacheGeneration.Unavailable, WorldwideFeed.CreateMicrosoft(f.Root).ReadCache().Generation);

        var file = BrowserDocument.OpenFile(path, f.Clock, Token);
        Check.Equal(BrowserSource.File, file.Origin!.Kind);
        Check.Equal("File", file.Snapshot!.Provenance);
        Check.True(file.Snapshot.Observation == null && file.Snapshot.Version == null);
        using var json = JsonDocument.Parse(rows);
        foreach (var document in new[] { file, download })
        {
            Check.True(JsonElement.DeepEquals(json.RootElement, document.Snapshot!.Source));
            var result = BrowserResults.Select(document, new() { Domain = "autodiscover.tenant.onmicrosoft.com" });
            Check.Equal(3, result.Hits.Single().Entry.Id);
        }

        var malformed = JsonNode.Parse(rows)!;
        malformed[2]!["urls"]!.AsArray().Add("bad**.example.test");
        f.Hour();
        f.Transport.Pair(SimulatedTransport.Second, malformed.ToJsonString());
        var stale = await BrowserDocument.DownloadAsync(f.Feed, f.Clock, Token);
        Check.Equal(RefreshStatus.Failed, stale.Origin!.Status);
        Check.True(stale.Origin.Error!.Contains("endpoints[2]: urls[2]: Invalid DNS"));
        Check.Equal(download.Snapshot.Observation, stale.Snapshot!.Observation);
        Check.True(JsonElement.DeepEquals(json.RootElement, stale.Snapshot.Source));
        Check.Equal(RefreshStatus.Deferred, (await f.Feed.RefreshAsync()).Status);
        Check.Equal(4, f.Transport.Ids.Count);
        f.Transport.Drained();
    }

    private static async Task RoundTrip(string root)
    {
        var clock = new TestClock();
        BrowserDocument Local(string name)
        {
            string path = Path.Combine(root, name + ".json");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", name + ".json"), path);
            var doc = BrowserDocument.OpenFile(path, clock, Token);
            File.Delete(path);
            return doc;
        }
        var applied = new BrowserFilter { Area = ServiceArea.Exchange, Protocol = Transport.Tcp, Ports = "5002-5004" };
        var current = BrowserWorkspace.Empty.Open(Local("synthetic-2")) with
        {
            Draft = applied, Reference = Local("synthetic-1"), SelectedId = 103,
            Window = new(-8000, 35, 1500, 610, true)
        };
        var good = current.Present(applyDraft: true).Workspace;
        var invalid = (good with { Draft = applied with { Ports = "70000" } }).Present(applyDraft: true);
        Check.True(invalid.DraftError!.Contains("70000") && invalid.Workspace.Applied == applied);
        Check.Equal(103, invalid.Workspace.SelectedId);
        var store = Store(root, clock);
        store.Load();
        var writer = new WorkspaceWriter(store, TimeSpan.FromHours(1));
        var queued = writer.ScheduleAsync(invalid.Workspace);
        await writer.FlushAsync();
        await queued;
        var restored = Store(root, clock).Load().Workspace.Present();
        Check.Equal(invalid.DraftError, restored.DraftError);
        Check.Equal(string.Join("|", invalid.FieldErrors), string.Join("|", restored.FieldErrors));
        Check.True(restored.FieldErrors.Single().Field == FilterField.Ports);
        Check.Equal(applied, restored.Workspace.Applied);
        Check.Equal(invalid.Workspace.Draft, restored.Workspace.Draft);
        Check.Equal(good.Window, restored.Workspace.Window);
        Check.Equal(new WindowPlacement(0, 35, 1200, 620, true), good.Window!.Fit(0, 0, 1200, 700, 940, 620));
        Check.Equal(good.Document.Origin, restored.Workspace.Document.Origin);
        Check.Equal(good.Reference!.Origin, restored.Workspace.Reference!.Origin);
        Check.Equal(good.Reference.Snapshot!.Source.GetRawText(), restored.Workspace.Reference.Snapshot!.Source.GetRawText());
        Check.True(restored.Results.Hits.Single().Entry.Id == 101 && restored.Workspace.SelectedId == 103 &&
            restored.Results.Common.Last().Entry.Id == 103);
        var removed = (good with { Draft = new() { Change = ChangeKind.Removed } }).Present(true).Workspace;
        store.Save(removed);
        Check.Equal(105, Store(root, clock).Load().Workspace.Present().Results.Hits.Single().Entry.Id);
    }

    private static async Task WriteOrder(string root)
    {
        var clock = new TestClock();
        var store = Store(root, clock);
        var initial = store.Load().Workspace.Open(OpenFixture("synthetic-2", clock)).Present().Workspace;
        var other = Store(root, clock);
        other.Load();
        var writer = new WorkspaceWriter(store, TimeSpan.FromMilliseconds(80));
        var first = writer.ScheduleAsync(initial);
        var draft = initial with { Draft = new() { Notes = "draft" } };
        var latest = writer.ScheduleAsync(draft);
        await Task.WhenAll(first, latest);
        Check.True(!File.Exists(Previous(root)));
        Check.Equal(draft.Draft, Store(root, clock).Load().Workspace.Draft);

        var pending = Enumerable.Range(0, 12).Select(i =>
            writer.ScheduleAsync(initial with { Draft = new() { Notes = i.ToString() } })).ToArray();
        await writer.FlushAsync();
        await Task.WhenAll(pending);
        Check.True(File.Exists(Previous(root)));
        Check.Equal("11", Store(root, clock).Load().Workspace.Draft.Notes);
        await Check.Throws<IOException>(() => Task.Run(() => other.Save(initial)));
        Check.Equal("11", Store(root, clock).Load().Workspace.Draft.Notes);
    }

    private static async Task Recovery(string root)
    {
        var clock = new TestClock();
        var store = Store(root, clock);
        var first = store.Load().Workspace.Open(OpenFixture("synthetic-1", clock)).Present().Workspace;
        store.Save(first);
        var next = first.Open(OpenFixture("synthetic-2", clock)).Present().Workspace;
        store.Save(next);
        string active = Active(root);
        string previous = Previous(root);
        string valid = File.ReadAllText(active);
        string backup = File.ReadAllText(previous);
        await Check.Throws<InvalidDataException>(() => Task.Run(() => store.Save(next with { SelectedId = -1 })));
        Check.Equal(valid, File.ReadAllText(active));
        Check.Equal(backup, File.ReadAllText(previous));
        static string WithChecksum(string json)
        {
            using var envelope = JsonDocument.Parse(json);
            var root = envelope.RootElement;
            string sha = Convert.ToHexString(SHA256.HashData(JsonMarshal.GetRawUtf8Value(root.GetProperty("Workspace"))));
            return json.Replace(root.GetProperty("Sha256").GetString()!, sha);
        }
        foreach (string damaged in new[]
        {
            "{", """{"Format":9,"Future":"keep"}""", valid.Replace("synthetic-2", "synthetic-3"),
            WithChecksum(valid.Replace("\"Kind\":\"Synthetic\"", "\"Kind\":\"Internet\"")),
            WithChecksum(valid.Replace("\"Applied\":{", "\"Applied\":{\"$type\":\"anything\",")),
            valid.Replace("\"Draft\":{", "\"Draft\":{\"Ports\":null,"),
            WithChecksum(valid.Replace("\"SelectedId\":101", "\"SelectedId\":0"))
        })
        {
            Check.True(damaged != valid);
            File.WriteAllText(active, damaged);
            var reader = Store(root, clock);
            var read = reader.Load();
            Check.True(!read.CanSave && read.Warning!.Contains("Previous"));
            Check.Equal(first.Document.Snapshot!.Version, read.Workspace.Document.Snapshot!.Version);
            await Check.Throws<InvalidOperationException>(() => Task.Run(() => reader.Save(next)));
            Check.Equal(damaged, File.ReadAllText(active));
            Check.Equal(backup, File.ReadAllText(previous));
        }
        File.WriteAllText(previous, "[]");
        var unusable = Store(root, clock).Load();
        Check.True(!unusable.CanSave && unusable.Warning != null && unusable.Workspace.Document.Snapshot == null);
        File.WriteAllText(active, valid);
        var badBackup = Store(root, clock);
        badBackup.Load();
        await Check.Throws<InvalidDataException>(() => Task.Run(() => badBackup.Save(first)));
        Check.Equal(valid, File.ReadAllText(active));
        Check.Equal("[]", File.ReadAllText(previous));
    }

    private static async Task Limits(string root)
    {
        var clock = new TestClock();
        string fixture = Path.Combine(root, "synthetic-oversize.json");
        using (var file = File.Create(fixture)) file.SetLength(WorldwideFeed.SnapshotFileLimit + 1);
        await Check.Throws<InvalidDataException>(() => Task.FromResult(BrowserDocument.OpenFile(fixture, clock, Token)));
        string active = Active(root);
        Directory.CreateDirectory(Path.GetDirectoryName(active)!);
        using (var file = File.Create(active)) file.SetLength(64 * 1024 * 1024);
        Check.True(!Store(root, clock).Load().CanSave);
        File.Delete(active);
        var store = Store(root, clock);
        var state = store.Load().Workspace;
        using (var lease = new FileStream(Path.Combine(root, "state", "workspace.lock"),
            FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            await Check.Throws<IOException>(() => Task.Run(() => store.Save(state)));
        Check.True(!File.Exists(active));
        store.Save(state);
        var oversizeDraft = state with { Draft = new() { Notes = new string('x', 16385) } };
        await Check.Throws<InvalidDataException>(() => Task.Run(() => store.Save(oversizeDraft)));
        Check.Equal("", Store(root, clock).Load().Workspace.Draft.Notes);
        await Check.Throws<ArgumentException>(() => Task.FromResult(new WorkspaceStore(Path.Combine(root, "..", "state"), clock)));
        if (OperatingSystem.IsWindows()) return; // Link creation can require a privilege on Windows.
        string linked = Path.Combine(root, "linked");
        Directory.CreateSymbolicLink(linked, root);
        Check.True(!Store(linked, clock).Load().CanSave);
        string sentinel = Path.Combine(root, "sentinel.json");
        File.WriteAllText(sentinel, "untouched");
        File.CreateSymbolicLink(Previous(root), sentinel);
        await Check.Throws<IOException>(() => Task.Run(() => store.Save(state)));
        File.Delete(active);
        File.CreateSymbolicLink(active, sentinel);
        Check.True(!Store(root, clock).Load().CanSave);
        await Check.Throws<IOException>(() => Task.FromResult(BrowserDocument.OpenFile(active, clock, Token)));
        Check.Equal("untouched", File.ReadAllText(sentinel));
        string restricted = Directory.CreateDirectory(Path.Combine(root, "no-access")).FullName;
        var mode = File.GetUnixFileMode(restricted);
        try
        {
            File.SetUnixFileMode(restricted, UnixFileMode.None);
            var denied = Store(restricted, clock).Load();
            // Privileged accounts can traverse mode 000; the exclusive-file test above works for them too.
            if (Environment.UserName != "root") Check.True(!denied.CanSave && denied.Warning != null);
        }
        finally { File.SetUnixFileMode(restricted, mode); }
        string dangling = Path.Combine(root, "dangling");
        File.CreateSymbolicLink(dangling, Path.Combine(root, "missing"));
        await Check.Throws<IOException>(() => Task.FromResult(BrowserDocument.OpenFile(dangling, clock, Token)));
    }
}
