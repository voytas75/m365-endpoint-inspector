using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using M365.EndpointInspector.Feed.Tests;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class ExpanderCases
{
    internal static void Run(string repo, string root)
    {
        var document = BrowserDocument.OpenSynthetic(Path.Combine(AppContext.BaseDirectory, "Fixtures", "synthetic-2.json"), TimeProvider.System, CancellationToken.None);
        foreach (bool source in new[] { false, true })
        foreach (bool comparison in new[] { false, true })
        {
            string path = Path.Combine(root, $"synthetic-expanders-{source}-{comparison}");
            var state = BrowserWorkspace.Empty.Open(document) with
            {
                SourceFieldsExpanded = source, ComparisonExpanded = comparison,
                Draft = new BrowserFilter { Pointer = "bad-pointer" }, SelectedId = 103
            };
            var store = new WorkspaceStore(path, TimeProvider.System);
            store.Load(); store.Save(state);
            string file = Path.Combine(path, "workspace.json");
            string original = File.ReadAllText(file);
            var loaded = new WorkspaceStore(path, TimeProvider.System).Load();
            Check.True(loaded.CanSave && loaded.Warning == null);
            Check.Equal(source, loaded.Workspace.SourceFieldsExpanded);
            Check.Equal(comparison, loaded.Workspace.ComparisonExpanded);
            Check.Equal(state.Draft, loaded.Workspace.Draft);
            Check.Equal(103, loaded.Workspace.SelectedId!.Value);
            Check.True(loaded.Workspace.Present().FieldErrors.Any(e => e.Field == FilterField.Pointer));
            Check.True(JsonElement.DeepEquals(document.Snapshot!.Source, loaded.Workspace.Document.Snapshot!.Source));
            Check.Equal(source, loaded.Workspace.Open(document).SourceFieldsExpanded);
            Check.Equal(comparison, loaded.Workspace.Open(document).ComparisonExpanded);
            Check.Equal(original, File.ReadAllText(file));

            void Rewrite(Action<JsonObject> change)
            {
                var payload = JsonNode.Parse(original)!["Workspace"]!.AsObject();
                change(payload);
                string text = payload.ToJsonString();
                string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
                File.WriteAllText(file, "{\"Format\":1,\"Sha256\":\"" + hash + "\",\"Workspace\":" + text + "}");
            }
            Rewrite(p => { p.Remove("SourceFieldsExpanded"); p.Remove("ComparisonExpanded"); });
            var legacy = new WorkspaceStore(path, TimeProvider.System).Load();
            Check.True(legacy.CanSave && legacy.Warning == null);
            Check.True(!legacy.Workspace.SourceFieldsExpanded && !legacy.Workspace.ComparisonExpanded);
            Check.Equal(state.Draft, legacy.Workspace.Draft);
            foreach (string bad in new[] { "null", "42", "\"true\"", "{}", "[]" })
            {
                Rewrite(p => { p["SourceFieldsExpanded"] = JsonNode.Parse(bad); p["ComparisonExpanded"] = true; });
                var fallback = new WorkspaceStore(path, TimeProvider.System).Load();
                Check.True(fallback.CanSave && fallback.Warning == null);
                Check.True(!fallback.Workspace.SourceFieldsExpanded && fallback.Workspace.ComparisonExpanded);
                Check.Equal(103, fallback.Workspace.SelectedId!.Value);
                Rewrite(p => { p["SourceFieldsExpanded"] = true; p["ComparisonExpanded"] = JsonNode.Parse(bad); });
                fallback = new WorkspaceStore(path, TimeProvider.System).Load();
                Check.True(fallback.CanSave && fallback.Workspace.SourceFieldsExpanded && !fallback.Workspace.ComparisonExpanded);
            }
        }
        string code = File.ReadAllText(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/MainWindow.xaml.cs"));
        foreach (string control in new[] { "SourceFields", "Comparison" })
        {
            Check.True(code.Contains(control + ".Expanded += (_, _) => CaptureChanges();"));
            Check.True(code.Contains(control + ".Collapsed += (_, _) => CaptureChanges();"));
            Check.True(code.Contains(control + "Expanded = " + control + ".IsExpanded"));
            Check.True(code.Contains(control + ".IsExpanded = true;"));
        }
        Check.True(code.IndexOf("SourceFields.IsExpanded = view.SourceFieldsExpanded;", StringComparison.Ordinal) <
            code.IndexOf("ShowFilterErrors(display.FieldErrors);", StringComparison.Ordinal));
        Check.True(code.Contains("if (!canSave || rendering || closing) return;"));
    }
}
