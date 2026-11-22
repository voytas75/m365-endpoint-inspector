using System.Text;
using System.Xml.Linq;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed.Tests;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class CsvExportCases
{
    internal static void Run(string repo, string root)
    {
        using var fixture = new Fixture(Path.Combine(root, "csv-simulated"));
        BrowserDocument Open(string name) => BrowserDocument.OpenSynthetic(
            Path.Combine(AppContext.BaseDirectory, "Fixtures", name + ".json"), fixture.Clock, default);
        var document = Open("synthetic-2");
        var workspace = BrowserWorkspace.Empty.Open(document) with
        {
            Applied = new() { Area = ServiceArea.Exchange }, Draft = new() { Ports = "70000" }
        };
        var display = workspace.Present();
        BrowserCsvExport Capture(BrowserWorkspace w, BrowserResults r, bool? common = null) =>
            BrowserCsvExport.Capture(w, r.Hits, r.Common, common ?? r.ShowCommon, w.Draft);
        var captured = Capture(workspace, display.Results);
        var rows = Parse(captured.CopyBytes());
        Check.Equal("101,104,103", string.Join(",", rows.Skip(1).Select(r => r[1])));
        Check.Equal("Hits,Common,Common", string.Join(",", rows.Skip(1).Select(r => r[0])));
        Check.True(captured.Message.Contains("błędnego filtra") && workspace.Applied != workspace.Draft);
        Check.Equal(2, Parse(Capture(workspace, display.Results, false).CopyBytes()).Count);
        var reversed = BrowserCsvExport.Capture(workspace, display.Results.Hits.Reverse(),
            display.Results.Common.Reverse(), true, workspace.Draft);
        Check.Equal("101,103,104", string.Join(",", Parse(reversed.CopyBytes()).Skip(1).Select(r => r[1])));
        byte[] immutable = captured.CopyBytes();
        _ = workspace.Open(Open("synthetic-1")).Present();
        Check.True(immutable.SequenceEqual(captured.CopyBytes()));
        var before = captured.CopyBytes(); before[0] = 0;
        Check.True(captured.CopyBytes()[0] == 0xef);
        var removedWorkspace = workspace with { Reference = Open("synthetic-1"), Applied = new() { Change = ChangeKind.Removed } };
        var removed = Parse(Capture(removedWorkspace, removedWorkspace.Present().Results).CopyBytes())[1];
        Check.Equal("Reference", removed[12]); Check.Equal("Synthetic", removed[14]);
        Check.Equal("synthetic-1", removed[16]); Check.Equal("SyntheticDeclaration", removed[17]);
        Fail<InvalidDataException>(() => Capture(BrowserWorkspace.Empty, BrowserResults.Select(BrowserDocument.Empty, new())));
        var none = workspace with { Applied = new() { Domain = "absent.example.test" } };
        Fail<InvalidDataException>(() => Capture(none, none.Present().Results));
        var onlyCommon = workspace with { Applied = new() { Area = ServiceArea.Exchange, Domain = "absent.example.test" } };
        Check.Equal(3, Parse(Capture(onlyCommon, onlyCommon.Present().Results).CopyBytes()).Count);
        Fail<InvalidDataException>(() => BrowserCsvExport.Capture(workspace,
            Enumerable.Repeat(display.Results.Hits[0], BrowserCsvExport.RowLimit + 1), [], false, new()));
        string path = Path.Combine(root, "csv-untrusted-file.json");
        string text = "Żółć, \"cytat\"\r\nDruga linia";
        foreach (string value in new[] { text, "=1+1", " +1", "\t-2", "\r\n@SUM(1)", "\u200b=1", "bez formuły" })
        {
            File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new[] { new
            {
                id = 1, serviceArea = "Common", category = "Default", urls = new[] { "synthetic.example.test" },
                serviceAreaDisplayName = value
            } }));
            var fileWorkspace = BrowserWorkspace.Empty.Open(BrowserDocument.OpenFile(path, fixture.Clock, default));
            var cells = Parse(Capture(fileWorkspace, fileWorkspace.Present().Results).CopyBytes())[1];
            Check.Equal(value == text || value == "bez formuły" ? value : "'" + value, cells[9]);
            Check.Equal("File", cells[13]); Check.Equal("File", cells[14]);
            Check.Equal("Unknown", cells[16]); Check.Equal("Unknown", cells[17]);
            Check.Equal("Unknown", cells[21]); Check.Equal("Unknown", cells[22]);
            Check.True(!Encoding.UTF8.GetString(captured.CopyBytes()).Contains(document.Origin!.Path!));
        }
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new[] { new
        {
            id = 1, serviceArea = "Common", category = "Default", urls = new[] { "synthetic.example.test" },
            serviceAreaDisplayName = new string('Ż', 10000)
        } }));
        var large = BrowserWorkspace.Empty.Open(BrowserDocument.OpenFile(path, fixture.Clock, default));
        Fail<InvalidDataException>(() => BrowserCsvExport.Capture(large,
            Enumerable.Repeat(large.Present().Results.Hits[0], 1000), [], false, new()));
        fixture.Seed().GetAwaiter().GetResult();
        var declared = BrowserWorkspace.Empty.Open(BrowserDocument.OpenFile(Path.Combine(fixture.DirectoryPath, "active.json"), fixture.Clock, default));
        var declaration = Parse(Capture(declared, declared.Present().Results).CopyBytes())[1];
        Check.Equal("FileDeclaredObservation", declaration[17]);
        Check.Equal("SimulatedWorldwide", declaration[20]); Check.Equal(SimulatedTransport.First, declaration[16]);
        var cache = BrowserWorkspace.Empty.Open(BrowserDocument.OpenCache(fixture.DirectoryPath, fixture.Clock, default));
        var cacheRow = Parse(Capture(cache, cache.Present().Results).CopyBytes())[1];
        Check.Equal("Cache", cacheRow[13]); Check.Equal("SimulatedWorldwide", cacheRow[14]); Check.Equal("Active", cacheRow[15]);

        string target = Path.Combine(root, "export.csv");
        var newTarget = CsvExportFile.Inspect(target);
        File.WriteAllText(target, "concurrent");
        Fail<IOException>(() => newTarget.Save(captured, false)); Check.Equal("concurrent", File.ReadAllText(target));
        var approved = CsvExportFile.Inspect(target);
        Fail<InvalidOperationException>(() => approved.Save(captured, false));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Fail<OperationCanceledException>(() => approved.Save(captured, true, cancelled.Token));
        Check.Equal("concurrent", File.ReadAllText(target));
        File.WriteAllText(target, "changed after approval");
        Fail<IOException>(() => approved.Save(captured, true)); Check.Equal("changed after approval", File.ReadAllText(target));
        CsvExportFile.Inspect(target).Save(captured, true);
        Check.True(File.ReadAllBytes(target).SequenceEqual(captured.CopyBytes()));
        File.Delete(target); CsvExportFile.Inspect(target).Save(captured, false);
        Check.True(File.ReadAllBytes(target).SequenceEqual(captured.CopyBytes()));
        string denied = Path.Combine(root, "csv-directory-target"); Directory.CreateDirectory(denied);
        Fail<IOException>(() => CsvExportFile.Inspect(denied).Save(captured, false));
        Check.True(Directory.Exists(denied));
        string absent = Path.Combine(root, "csv-cancelled-new.csv");
        Fail<OperationCanceledException>(() => CsvExportFile.Inspect(absent).Save(captured, false, cancelled.Token));
        Check.True(!File.Exists(absent));
        Check.True(!Directory.EnumerateFiles(root).Any(p => p.EndsWith(".original")));

        var window = XDocument.Load(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/MainWindow.xaml"));
        Check.True(window.Descendants().Any(e => (string?)e.Attribute("Click") == "ExportCsv_Click"));
        string code = File.ReadAllText(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/MainWindow.xaml.cs"));
        string handler = code[code.IndexOf("private void ExportCsv_Click", StringComparison.Ordinal)..code.IndexOf("private async void Download_Click", StringComparison.Ordinal)];
        Check.True(handler.IndexOf("BrowserCsvExport.Capture", StringComparison.Ordinal) < handler.IndexOf("ShowDialog", StringComparison.Ordinal));
        Check.True(handler.Contains("OverwritePrompt = true") && handler.Contains("HitsGrid.Items.Cast<BrowserRow>()") && handler.Contains("CommonGrid.Items.Cast<BrowserRow>()"));
        Check.True(handler.Contains("if (picker.ShowDialog(this) != true) return;") && !handler.Contains("CaptureChanges") && !handler.Contains("ShowView") && !handler.Contains("ApplyAsync"));
    }

    // Independent strict reader: every exported cell is quoted; only CRLF separates records.
    private static List<string[]> Parse(byte[] bytes)
    {
        Check.True(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble));
        string csv = new UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3);
        var records = new List<string[]>(); var cells = new List<string>(); int i = 0;
        while (i < csv.Length)
        {
            Check.True(csv[i++] == '"'); var cell = new StringBuilder(); bool closed = false;
            while (i < csv.Length)
            {
                char c = csv[i++];
                if (c != '"') { cell.Append(c); continue; }
                if (i < csv.Length && csv[i] == '"') { cell.Append('"'); i++; continue; }
                closed = true; break;
            }
            Check.True(closed); cells.Add(cell.ToString());
            if (csv[i] == ',') { i++; continue; }
            Check.True(csv[i++] == '\r' && csv[i++] == '\n');
            Check.Equal(23, cells.Count); records.Add(cells.ToArray()); cells.Clear();
        }
        return records;
    }
    private static void Fail<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new InvalidOperationException($"Expected {typeof(T).Name}");
    }
}
