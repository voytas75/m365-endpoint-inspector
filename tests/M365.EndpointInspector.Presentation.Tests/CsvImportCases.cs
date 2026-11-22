using System.Text;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Presentation;
using M365.EndpointInspector.Feed.Tests;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class CsvImportCases
{
    internal static void Run(string root)
    {
        // Synthetic observed-schema simulation.
        string csv = WorldwideCsvImporter.Header + ",extra\n1,Common,Symulacja,\"*.example.test,other.example.test\",,443,,False,Allow,True,,zażółć";
        string path = Path.Combine(root, "simulated.csv");
        File.WriteAllText(path, csv, new UTF8Encoding(false));
        var doc = BrowserDocument.OpenFile(path, TimeProvider.System, CancellationToken.None);
        Check.True(doc.Origin!.Csv && doc.SourceLabel.Contains("CSV"));
        string rawSource = doc.Snapshot!.Source.GetRawText();
        byte[] originalFile = File.ReadAllBytes(path);
        const string prefix = "Reprezentacja JSON utworzona z CSV (nie oryginalny JSON).\n";
        string preview = BrowserResults.Select(doc, new()).Hits.Single().RawJson;
        Check.True(preview.StartsWith(prefix));
        Program.AssertReadablePreview(doc.Snapshot.Entries[1].Source, preview[prefix.Length..]);
        Check.Equal("zażółć", doc.Snapshot.Entries[1].Source.GetProperty("extra").GetString());
        Check.True(originalFile.SequenceEqual(File.ReadAllBytes(path)));
        Check.Equal(rawSource, doc.Snapshot.Source.GetRawText());
        var workspace = BrowserWorkspace.Empty.Open(doc).Present().Workspace;
        Check.True(BrowserResults.Select(doc, new()).Hits.Single().RawJson.Contains("utworzona z CSV"));
        var store = new WorkspaceStore(Path.Combine(root, "csv-workspace"), TimeProvider.System);
        store.Load(); store.Save(workspace);
        File.Delete(path);
        var restored = store.Load();
        Check.True(restored.Warning == null && restored.Workspace.Document.Origin!.Csv);
        Check.Equal(1, restored.Workspace.SelectedId);
        Check.Equal(rawSource, restored.Workspace.Document.Snapshot!.Source.GetRawText());
        string restoredPreview = BrowserResults.Select(restored.Workspace.Document, new()).Hits.Single().RawJson;
        Check.Equal(preview, restoredPreview);
        Program.AssertReadablePreview(restored.Workspace.Document.Snapshot.Entries[1].Source, restoredPreview[prefix.Length..]);
        var export = BrowserCsvExport.Capture(workspace, BrowserResults.Select(doc, new()).Hits, [], false, new());
        Check.True(Encoding.UTF8.GetString(export.CopyBytes()).Contains("FileCsv"));
        File.WriteAllBytes(path, [0xff]);
        try { BrowserDocument.OpenFile(path, TimeProvider.System, CancellationToken.None); throw new Exception("invalid UTF8 accepted"); }
        catch (FormatException ex) { Check.True(ex.Message.Contains("wiersz 1, kolumna 1")); }
        Check.True(BrowserResults.Select(restored.Workspace.Document, new()).Hits.Single().RepresentationLabel.Contains("CSV"));
        File.WriteAllBytes(path, export.CopyBytes());
        try { BrowserDocument.OpenFile(path, TimeProvider.System, CancellationToken.None); throw new Exception("own export accepted"); }
        catch (FormatException ex) { Check.True(ex.Message.Contains("BrowserCSV")); }
        Check.True(ReferenceEquals(doc, workspace.Document));
        try { BrowserDocument.OpenFile(path, TimeProvider.System, new CancellationToken(true)); throw new Exception("cancel accepted"); }
        catch (OperationCanceledException) { }
        Check.True(ReferenceEquals(doc, workspace.Document));
        string jsonPath = Path.Combine(root, "malformed.json"); File.WriteAllText(jsonPath, csv);
        try { BrowserDocument.OpenFile(jsonPath, TimeProvider.System, CancellationToken.None); throw new Exception("JSON treated as CSV"); }
        catch (System.Text.Json.JsonException) { }
        string unknownPath = Path.Combine(root, "simulated.data"); File.WriteAllText(unknownPath, csv);
        Check.True(BrowserDocument.OpenFile(unknownPath, TimeProvider.System, CancellationToken.None).Origin!.Csv);
    }
}
