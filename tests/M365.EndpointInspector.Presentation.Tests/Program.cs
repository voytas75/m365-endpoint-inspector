using System.Reflection;
using System.Xml.Linq;
using M365.EndpointInspector.Feed;
using System.Text;
using System.Text.Json;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed.Tests;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class Program
{
    private static readonly TimeProvider Clock = new FixedClock();
    private static readonly CancellationToken Token = CancellationToken.None;

    private sealed class FixedClock : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(2026, 1, 4, 13, 15, 0, TimeSpan.Zero);
    }

    private static BrowserDocument Open(string name) => BrowserDocument.OpenSynthetic(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name + ".json"), Clock, Token);

    private static async Task<int> Main()
    {
        var repo = new DirectoryInfo(AppContext.BaseDirectory);
        while (repo != null && !File.Exists(Path.Combine(repo.FullName, "M365.EndpointInspector.slnx")))
            repo = repo.Parent;
        string root = Path.Combine(Environment.GetEnvironmentVariable("M365_TEST_ARTIFACT_ROOT") ?? Path.Combine(repo!.FullName, "artifacts"), "presentation-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        (string Name, Action Run)[] groups =
        [
            ("bounded independent grid layout and legacy restore", () => GridLayoutCases.Run(repo!.FullName, root)),
            ("CSV file, details, restore, cancel and failure continuity", () => CsvImportCases.Run(root)),
            ("bounded browser CSV capture and safe save", () => CsvExportCases.Run(repo!.FullName, root)),
            ("field-local validation and WPF wiring", () => FilterValidationCases.Run(repo!.FullName)),
            ("embedded product version, policy and both WPF labels", () => BuildIdentity(repo!.FullName)),
            ("empty startup, empty snapshot and synthetic metadata", () => SourceStates(root)),
            ("import errors, strict UTF-8, size limit and cancellation", () => ImportErrors(root)),
            ("filter mappings, port ranges, Common and reset", FilterTable),
            ("shared targets, all IDs, details and priority", () => Memberships(root)),
            ("typed JSON pointer, missing fields and filter errors", AdvancedFilters),
            ("synthetic comparison and removed row details", Comparison)
        ];
        try
        {
            foreach (var group in groups)
            {
                group.Run();
                Console.WriteLine($"PASS {group.Name}");
            }
            await CacheStates(root);
            Console.WriteLine("PASS simulated active, fallback, age, errors and source isolation");
            int passed = groups.Length + 1 + await WorkspaceCases.Run(root);
            Console.WriteLine($"{passed}/{passed} offline presentation groups passed.");
            Console.WriteLine($"Simulated cache fixtures: {root}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static void BuildIdentity(string root)
    {
        var assembly = typeof(BuildInfo).Assembly;
        var info = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion.Split('+');
        Check.Equal("0.3.4-dev", info[0]);
        Check.True(info[1].Length == 64 && info[1].All(char.IsAsciiHexDigit));
#if DEBUG
        var build = new BuildInfo(assembly, true);
        Check.Equal(RefreshPolicy.ManualDevelopment, build.Policy);
        string mode = "DEV";
#else
        var build = new BuildInfo(assembly, false);
        Check.Equal(RefreshPolicy.Release, new BuildInfo(assembly, true).Policy);
        string mode = "RELEASE";
#endif
        Check.Equal(RefreshPolicy.Release, new BuildInfo(assembly).Policy);
        Check.Equal($"v{info[0]} · {mode} · build {info[1][..12]}", build.Label);
        Check.True(build.WindowTitle.EndsWith(build.Label, StringComparison.Ordinal));
        var window = XDocument.Load(Path.Combine(root, "src/M365.EndpointInspector.Wpf/MainWindow.xaml")).Root!;
        Check.Equal("{Binding WindowTitle}", (string?)window.Attribute("Title"));
        Check.True(window.Descendants().Any(e => (string?)e.Attribute("Text") == "{Binding Label}"));
        Console.WriteLine(build.WindowTitle);
    }

    private static void SourceStates(string root)
    {
        var empty = BrowserDocument.Empty;
        Check.True(empty.Snapshot == null);
        Check.Equal("Brak danych", BrowserResults.Select(empty, new()).Summary);
        string path = Path.Combine(root, "empty-synthetic.json");
        File.WriteAllText(path, """{"provenance":"Synthetic","version":"empty","importedAt":"2026-01-01T12:00:00Z","endpoints":[]}""",
            new UTF8Encoding(true));
        var document = BrowserDocument.OpenSynthetic(path, Clock, Token);
        Check.True(document.IsSynthetic && document.SourceLabel.Contains("Synthetic"));
        Check.Equal("Pusta migawka (0 wpisów)", BrowserResults.Select(document, new()).Summary);
        var demo = Open("synthetic-2");
        Check.True(demo.Metadata.Contains("synthetic-2") && demo.Metadata.Contains("2026-01-02 12:00:00 UTC") &&
            demo.Metadata.Contains("2 d 1 godz. 15 min"));
    }

    private static void ImportErrors(string root)
    {
        string path = Path.Combine(root, "invalid-synthetic.json");
        File.WriteAllText(path, """{"provenance":"MicrosoftWorldwide","endpoints":[]}""");
        Throws<FormatException>(() => BrowserDocument.OpenSynthetic(path, Clock, Token));
        File.WriteAllText(path, "[]");
        Throws<FormatException>(() => BrowserDocument.OpenSynthetic(path, Clock, Token));
        File.WriteAllBytes(path, [0xff, 0xff]);
        Throws<DecoderFallbackException>(() => BrowserDocument.OpenSynthetic(path, Clock, Token));
        using (var stream = File.Create(path)) stream.SetLength(4 * 1024 * 1024 + 1);
        Throws<InvalidDataException>(() => BrowserDocument.OpenSynthetic(path, Clock, Token));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Throws<OperationCanceledException>(() => BrowserDocument.OpenSynthetic(path, Clock, cancelled.Token));
        Throws<OperationCanceledException>(() => BrowserDocument.OpenCache(root, Clock, cancelled.Token));
    }

    private static void FilterTable()
    {
        var document = Open("synthetic-2");
        var filter = new BrowserFilter
        {
            Area = ServiceArea.Exchange, Category = Category.Optimize, Required = true,
            Protocol = Transport.Tcp, Ports = "5002-5008", Domain = "X.MAIL.EXAMPLE.TEST.",
            Address = "192.0.2.40", DisplayName = "Synthetic Mail", Notes = "Synthetic range"
        };
        var result = BrowserResults.Select(document, filter);
        Ids(result.Hits, 101);
        Ids(result.Common, 104, 103);
        Check.True(result.ShowCommon && result.Summary.Contains("(1)") && result.CommonHeading.Contains("(2)"));
        Check.True(result.Summary.Contains("Optimize"));
        foreach (var miss in new[]
        {
            filter with { Area = ServiceArea.SharePoint }, filter with { Category = Category.Allow },
            filter with { Required = false }, filter with { Protocol = Transport.Udp },
            filter with { Ports = "5005" }, filter with { Domain = "x.notmail.example.test" },
            filter with { Address = "203.0.113.0/25" }, filter with { DisplayName = "Synthetic" },
            filter with { Notes = "Synthetic range " }
        })
        {
            var noHits = BrowserResults.Select(document, miss);
            Ids(noHits.Hits);
            Ids(noHits.Common, 104, 103);
            Check.Equal("Brak trafień (0)", noHits.Summary);
            Check.True(noHits.FilterError == null && noHits.ShowCommon);
        }
        var reset = BrowserResults.Select(document, new());
        Ids(reset.Hits, 101, 102, 104, 103, 106);
        Check.True(!reset.ShowCommon && reset.Common.Count == 0);
        var common = BrowserResults.Select(document, new() { Area = ServiceArea.Common });
        Ids(common.Hits, 104, 103);
        Check.True(!common.ShowCommon);
    }

    private static void Memberships(string root)
    {
        var document = Open("synthetic-2");
        var exchange = BrowserResults.Select(document, new() { Area = ServiceArea.Exchange });
        var row = exchange.Hits.Single();
        Check.Equal("Tak", row.Required);
        Check.Equal("Tak", row.ExpressRoute);
        Check.Equal("443, 5000-5004", row.Tcp);
        Check.Equal("3478-3481", row.Udp);
        Check.Equal("*.mail.example.test, shared.example.test", row.Domains);
        Check.True(row.Addresses.Contains("2001:db8:1::20/123"));
        AssertReadablePreview(document.Snapshot!.Entries[101].Source, row.RawJson);
        Check.True(row.RawJson.Contains("\n    \"*.mail.example.test\",\n    \"shared.example.test\""));
        using var json = JsonDocument.Parse(row.RawJson);
        Check.Equal(2, json.RootElement.GetProperty("extra").GetProperty("routing")[0].GetProperty("weight").GetInt32());
        Check.Equal("—", exchange.Common.First().ExpressRoute);
        Check.True(exchange.Common.Last().RawJson.Contains("\"notes\": null"));
        var targets = row.SearchableTargets.ToArray();
        var pattern = targets.Single(t => t.Value == "*.mail.example.test").AllMemberships();
        Ids(BrowserResults.Select(document, pattern).Hits, 101);
        string sharedPattern = Path.Combine(root, "shared-pattern-synthetic.json");
        File.WriteAllText(sharedPattern, document.Snapshot.Source.GetRawText()
            .Replace("new.example.test", "*.MAIL.example.test."));
        var repeated = BrowserDocument.OpenSynthetic(sharedPattern, Clock, Token);
        Ids(BrowserResults.Select(repeated, pattern).Hits, 101, 106);
        var skypeOnly = BrowserResults.Select(repeated, pattern with { Area = ServiceArea.Skype });
        Ids(skypeOnly.Hits, 106);
        Ids(skypeOnly.Common, 104, 103);
        Check.True(skypeOnly.Summary.Contains("Default"));
        var shared = BrowserResults.Select(document, targets.Single(t => t.Value == "shared.example.test").AllMemberships());
        Ids(shared.Hits, 101, 102, 103);
        Check.True(shared.Hits.All(hit => hit.Title.Contains(hit.Entry.Id.ToString())));
        var cidr = targets.Single(t => t.Value == "2001:db8:1::20/123");
        Ids(BrowserResults.Select(document, cidr.AllMemberships()).Hits, 101, 102);
    }

    private static void AdvancedFilters()
    {
        var document = Open("synthetic-2");
        BrowserResults Run(BrowserFilter f) => BrowserResults.Select(document, f);
        var pointer = new BrowserFilter { Pointer = "/extra/routing/0/weight", JsonValue = "2" };
        Ids(Run(pointer).Hits, 101);
        Ids(Run(pointer with { JsonValue = "\"2\"" }).Hits);
        Ids(Run(new() { Pointer = "/extra/a~1b/~0name", JsonValue = "\"scoped\"" }).Hits, 101);
        Ids(Run(new() { Pointer = "/expressRoute", JsonValue = "null" }).Hits, 104);
        Ids(Run(new() { Pointer = "/expressRoute" }).Hits, 103, 106);
        foreach (var bad in new BrowserFilter[]
        {
            new() { Ports = "443" }, new() { Protocol = Transport.Tcp, Ports = "65536" },
            new() { Address = "192.0.2.1/24" }, new() { Domain = "*.example.test" },
            new() { Pointer = "/extra~2" }, new() { JsonValue = "null" }, new() { Pointer = "/id", JsonValue = "[" }
        })
        {
            var invalid = Run(bad);
            Check.Equal("Błąd filtra", invalid.Summary);
            Check.True(!string.IsNullOrEmpty(invalid.FilterError) && invalid.Hits.Count == 0);
        }
        Ids(Run(new() { Protocol = Transport.Udp, Ports = "65535" }).Hits, 106);
    }

    internal static void AssertReadablePreview(JsonElement source, string preview)
    {
        string original = source.GetRawText();
        using var formatted = JsonDocument.Parse(preview);
        Check.True(JsonElement.DeepEquals(source, formatted.RootElement));
        Check.True(preview.StartsWith("{\n  \""));
        Check.True(preview.Split('\n').Length > source.EnumerateObject().Count());
        Check.Equal(original, source.GetRawText());
        Check.Equal(preview, JsonSerializer.Serialize(source, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static void Comparison()
    {
        var document = Open("synthetic-2");
        var baseline = Open("synthetic-1");
        BrowserResults Changes(ChangeKind change) => BrowserResults.Select(document, new() { Change = change }, baseline);
        Ids(Changes(ChangeKind.Added).Hits, 106);
        Ids(Changes(ChangeKind.Modified).Hits, 101, 104);
        Ids(Changes(ChangeKind.Unchanged).Hits, 102, 103);
        var removed = Changes(ChangeKind.Removed).Hits.Single();
        Check.Equal("Usunięty", removed.Change);
        AssertReadablePreview(baseline.Snapshot!.Entries[105].Source, removed.RawJson);
        Check.True(BrowserResults.Select(document, new() { Change = ChangeKind.Removed }).FilterError != null);
        Check.True(BrowserResults.Select(document, new(), BrowserDocument.Empty).FilterError != null);
    }

    private static async Task CacheStates(string root)
    {
        using var f = new Fixture(Path.Combine(root, "active"));
        await f.Seed();
        f.Hour();
        await f.Seed(SimulatedTransport.Second);
        f.Clock.Advance(TimeSpan.FromDays(12));
        var active = BrowserDocument.OpenCache(f.DirectoryPath, f.Clock, Token);
        Check.True(active.SourceLabel.Contains("SimulatedWorldwide") && active.SourceLabel.Contains("kopia aktywna"));
        Check.True(active.Metadata.Contains("12 d 0 godz.") && active.Metadata.Contains(SimulatedTransport.Second) &&
            active.Metadata.Contains("Wersję sprawdzono przed pobraniem"));
        Check.True(active.Notice == null && !active.IsSynthetic);
        Ids(BrowserResults.Select(active, new()).Hits, 2, 1);
        Check.True(BrowserResults.Select(active, new(), Open("synthetic-1")).FilterError != null);

        string fallbackDir = Path.Combine(root, "fallback", "SimulatedWorldwide");
        Directory.CreateDirectory(fallbackDir);
        File.Copy(Path.Combine(f.DirectoryPath, "previous.json"), Path.Combine(fallbackDir, "previous.json"));
        string broken = Path.Combine(fallbackDir, "active.json");
        File.WriteAllText(broken, "{");
        var fallback = BrowserDocument.OpenCache(fallbackDir, f.Clock, Token);
        Check.True(fallback.SourceLabel.Contains("kopia zapasowa") && fallback.SourceLabel.Contains("symulowany"));
        Check.True(fallback.Metadata.Contains(SimulatedTransport.First) && fallback.Metadata.Contains("12 d 1 godz."));
        Check.True(fallback.Details!.Contains("active:"));
        Check.Equal("{", File.ReadAllText(broken));
        Ids(BrowserResults.Select(fallback, new()).Hits, 2, 1);

        var missingDir = Path.Combine(f.Root, "MicrosoftWorldwide");
        Directory.CreateDirectory(missingDir);
        var missing = BrowserDocument.OpenCache(missingDir, f.Clock, Token);
        Check.True(missing.SourceLabel.Contains("produkcyjny") && missing.SourceLabel.Contains("brak poprawnej kopii"));
        Check.True(missing.Snapshot == null && missing.Details!.Contains("previous:"));
        Check.True(!Directory.EnumerateFileSystemEntries(missingDir).Any());
        Throws<ArgumentException>(() => BrowserDocument.OpenCache(f.Root, f.Clock, Token));
        Throws<ArgumentException>(() => BrowserDocument.OpenCache(Path.Combine(f.Root, "2"), f.Clock, Token));
        f.Transport.Drained();
    }

    private static void Ids(IReadOnlyList<BrowserRow> rows, params int[] expected) =>
        Check.Equal(string.Join(",", expected), string.Join(",", rows.Select(row => row.Entry.Id)));

    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException($"Expected {typeof(T).Name}.");
    }
}
