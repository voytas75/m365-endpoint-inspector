using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed.Tests;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class GridLayoutCases
{
    internal static void Run(string repo, string root)
    {
        var clock = TimeProvider.System;
        var document = BrowserDocument.OpenSynthetic(Path.Combine(AppContext.BaseDirectory, "Fixtures", "synthetic-2.json"), clock, CancellationToken.None);
        var hits = new GridLayout(GridLayout.Default.Columns.Select(c => c with
            { Order = 10 - c.Order, Width = c.Width + 17 }).ToArray(), [new("id", true), new("category", false)]);
        var common = new GridLayout(GridLayout.Default.Columns.Select(c => c with { Width = 200 }).ToArray(), [new("domains", false)]);
        Check.True(hits.IsValid() && common.IsValid());
        var filter = new BrowserFilter { Area = ServiceArea.Exchange };
        var workspace = BrowserWorkspace.Empty.Open(document) with
        {
            Applied = filter, Draft = filter with { Ports = "70000" }, SelectedId = 103,
            HitsLayout = hits, CommonLayout = common
        };
        var path = Path.Combine(root, "synthetic-layout-state");
        var store = new WorkspaceStore(path, clock);
        store.Load(); store.Save(workspace);
        var original = File.ReadAllText(Path.Combine(path, "workspace.json"));
        var restored = new WorkspaceStore(path, clock).Load();
        Check.True(restored.CanSave && restored.Warning == null);
        AssertLayout(hits, restored.Workspace.HitsLayout!);
        AssertLayout(common, restored.Workspace.CommonLayout!);
        Check.Equal(workspace.Applied, restored.Workspace.Applied);
        Check.Equal(workspace.Draft, restored.Workspace.Draft);
        Check.Equal(103, restored.Workspace.Present().Workspace.SelectedId!.Value);
        Check.True(JsonElement.DeepEquals(document.Snapshot!.Source, restored.Workspace.Document.Snapshot!.Source));
        AssertLayout(hits, restored.Workspace.Open(document).HitsLayout!);
        Check.Equal(original, File.ReadAllText(Path.Combine(path, "workspace.json")));

        void Rewrite(Action<JsonObject> change)
        {
            var envelope = JsonNode.Parse(original)!.AsObject();
            var payload = envelope["Workspace"]!.AsObject(); change(payload);
            string text = payload.ToJsonString();
            envelope["Sha256"] = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
            // Retain exact payload bytes used for the checksum.
            File.WriteAllText(Path.Combine(path, "workspace.json"), "{\"Format\":1,\"Sha256\":" + envelope["Sha256"]!.ToJsonString() + ",\"Workspace\":" + text + "}");
        }
        Rewrite(p => { p.Remove("HitsLayout"); p.Remove("CommonLayout"); });
        var legacy = new WorkspaceStore(path, clock).Load();
        Check.True(legacy.CanSave && legacy.Workspace.Document.Snapshot != null);
        AssertLayout(GridLayout.Default, legacy.Workspace.HitsLayout!);
        foreach (string bad in new[] { "null", "[]", "42", "\"bad\"", "{}", "{\"Columns\":null,\"Sorts\":[]}",
            "{\"Columns\":[],\"Sorts\":[]}", JsonSerializer.Serialize(hits with { Sorts = [new("unknown", false)] }),
            JsonSerializer.Serialize(hits with { Columns = hits.Columns.Select(c => c with { Order = 0 }).ToArray() }),
            JsonSerializer.Serialize(hits with { Columns = hits.Columns.Select(c => c with { Width = 2001 }).ToArray() }),
            JsonSerializer.Serialize(hits with { Sorts = [new("id", false), new("id", true)] }),
            JsonSerializer.Serialize(hits).Replace("\"Descending\":true", "\"Descending\":\"bad\"") })
        {
            Rewrite(p => p["HitsLayout"] = JsonNode.Parse(bad));
            var fallback = new WorkspaceStore(path, clock).Load();
            Check.True(fallback.CanSave && fallback.Workspace.Document.Snapshot != null);
            Check.Equal(103, fallback.Workspace.SelectedId!.Value);
            AssertLayout(GridLayout.Default, fallback.Workspace.HitsLayout!);
            AssertLayout(common, fallback.Workspace.CommonLayout!);
        }
        Check.True(!(hits with { Columns = hits.Columns.Select(c => c with { Width = double.NaN }).ToArray() }).IsValid());
        Check.True(!(hits with { Columns = hits.Columns.Select(c => c with { Width = 19 }).ToArray() }).IsValid());
        var grid = XDocument.Load(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/EndpointGrid.xaml")).Root!;
        Check.Equal("True", (string?)grid.Attribute("CanUserSortColumns"));
        Check.Equal("True", (string?)grid.Attribute("CanUserReorderColumns"));
        Check.Equal(11, grid.Descendants().Count(c => c.Name.LocalName == "DataGridTextColumn" && c.Attribute("SortMemberPath") != null));
        string code = File.ReadAllText(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/EndpointGrid.xaml.cs"));
        Check.True(code.Contains("new ListCollectionView") && code.Contains("rows.SortDescriptions.Add") && code.Contains("SelectedItem = selected"));
        string window = File.ReadAllText(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/MainWindow.xaml.cs"));
        Check.True(window.Contains("grid.UpdatingLayout") && window.Contains("HitsGrid.Items.Cast<BrowserRow>()") && window.Contains("CommonGrid.Items.Cast<BrowserRow>()"));
    }

    private static void AssertLayout(GridLayout expected, GridLayout actual)
    {
        Check.True(expected.Columns.SequenceEqual(actual.Columns));
        Check.True(expected.Sorts.SequenceEqual(actual.Sorts));
    }
}
