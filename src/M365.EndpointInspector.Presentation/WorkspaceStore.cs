using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using M365.EndpointInspector.Feed;
using M365.EndpointInspector.Storage;

namespace M365.EndpointInspector.Presentation;

public sealed record WorkspaceLoad(BrowserWorkspace Workspace, string? Warning, bool CanSave);

public sealed class WorkspaceStore
{
    private sealed record SavedWorkspace(SavedDocument? Current, SavedDocument? Reference,
        BrowserFilter Applied, BrowserFilter Draft, int? SelectedId, WindowPlacement? Window, JsonElement? HitsLayout = null, JsonElement? CommonLayout = null,
        JsonElement? SourceFieldsExpanded = null, JsonElement? ComparisonExpanded = null);

    // Keep the source text shown in the detail pane across a save/reopen cycle.
    private sealed class SourceJsonConverter : JsonConverter<JsonElement>
    {
        public override JsonElement Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options) =>
            JsonElement.ParseValue(ref reader);
        public override void Write(Utf8JsonWriter writer, JsonElement value, JsonSerializerOptions options) =>
            writer.WriteRawValue(JsonMarshal.GetRawUtf8Value(value), skipInputValidation: true);
    }

    // Two escaped 4 MiB snapshots, plus bounded path, filter and status metadata.
    private const int Limit = 2 * WorldwideFeed.SnapshotFileLimit + 1024 * 1024;
    private static readonly JsonSerializerOptions Json = new()
    {
        AllowDuplicateProperties = false, RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        MaxDepth = 80, Converters = { new SourceJsonConverter(), new JsonStringEnumConverter(allowIntegerValues: false) }
    };

    private readonly string root;
    private readonly TimeProvider clock;
    private readonly object gate = new();
    private string? expectedHash;
    private bool canSave;

    public WorkspaceStore(string root, TimeProvider clock)
    {
        if (!Path.IsPathFullyQualified(root) ||
            root.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Contains(".."))
            throw new ArgumentException("The state directory requires a full path without parent segments.");
        this.root = root;
        this.clock = clock;
    }

    private string Leaf(string name) => LocalFile.CheckedPath(Path.Combine(root, name));

    private static byte[]? ReadOptional(string path)
    {
        try { return LocalFile.ReadBytes(path, Limit); }
        catch (FileNotFoundException) { return null; }
        catch (DirectoryNotFoundException) { return null; }
    }

    private BrowserWorkspace Decode(byte[] bytes)
    {
        using var document = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = Json.MaxDepth, AllowDuplicateProperties = false });
        var rootElement = document.RootElement;
        if (rootElement.ValueKind != JsonValueKind.Object || !rootElement.TryGetProperty("Format", out var format) ||
            format.ValueKind != JsonValueKind.Number || !format.TryGetInt32(out int version) || version != 1)
            throw new InvalidDataException("Unsupported state version.");
        if (rootElement.GetPropertyCount() != 3 || !rootElement.TryGetProperty("Workspace", out var data) ||
            data.ValueKind != JsonValueKind.Object || !rootElement.TryGetProperty("Sha256", out var digest) ||
            digest.ValueKind != JsonValueKind.String ||
            digest.GetString() != Convert.ToHexString(SHA256.HashData(JsonMarshal.GetRawUtf8Value(data))))
            throw new InvalidDataException("Invalid state checksum or envelope.");
        var saved = data.Deserialize<SavedWorkspace>(Json)!;
        var now = clock.GetUtcNow();
        var doc = saved.Current == null ? BrowserDocument.Empty : BrowserDocument.Restore(saved.Current, now);
        var reference = saved.Reference == null ? null : BrowserDocument.Restore(saved.Reference, now);
        if (reference != null && !(doc.IsSynthetic && reference.IsSynthetic))
            throw new InvalidDataException("A reference requires two Synthetic snapshots.");
        ValidateFilter(saved.Applied);
        ValidateFilter(saved.Draft);
        saved.Window?.Validate();
        var results = BrowserResults.Select(doc, saved.Applied, reference);
        if (results.FilterError != null || saved.SelectedId is int id &&
            !results.Hits.Concat(results.Common).Any(row => row.Entry.Id == id))
            throw new InvalidDataException("Invalid applied filter or selected ID.");
        return new(doc, saved.Applied, saved.Draft, reference, saved.SelectedId, saved.Window,
            GridLayout.Read(saved.HitsLayout), GridLayout.Read(saved.CommonLayout),
            saved.SourceFieldsExpanded?.ValueKind == JsonValueKind.True,
            saved.ComparisonExpanded?.ValueKind == JsonValueKind.True);
    }

    private static void ValidateFilter(BrowserFilter filter)
    {
        if (!Defined(filter.Area) || !Defined(filter.Category) || !Defined(filter.Protocol) || !Defined(filter.Change) ||
            new[] { filter.Ports, filter.Domain, filter.Address, filter.DisplayName, filter.Notes, filter.Pointer,
                filter.JsonValue }.Any(value => value == null || value.Length > 16384))
            throw new InvalidDataException("Invalid filter form fields.");
        static bool Defined<T>(T? value) where T : struct, Enum => value == null || Enum.IsDefined(value.Value);
    }

    private static string? Hash(byte[]? bytes) => bytes == null ? null : Convert.ToHexString(SHA256.HashData(bytes));

    public WorkspaceLoad Load()
    {
        lock (gate)
        {
            canSave = false;
            expectedHash = null;
            string? error = null;
            try
            {
                var bytes = ReadOptional(Leaf("workspace.json"));
                if (bytes != null)
                {
                    var restored = Decode(bytes);
                    expectedHash = Hash(bytes);
                    canSave = true;
                    return new(restored, null, true);
                }
            }
            catch (Exception ex) when (IsDataError(ex)) { error = ex.Message; }

            try
            {
                var previous = ReadOptional(Leaf("workspace.previous.json"));
                if (previous != null)
                    return new(Decode(previous), $"Previous state restored. Saving suspended. {error ?? "Current file missing."}", false);
            }
            catch (Exception ex) when (IsDataError(ex)) { error = $"{error} {ex.Message}".Trim(); }
            canSave = error == null;
            return new(BrowserWorkspace.Empty, canSave ? null : $"Could not restore state. Saving suspended. {error}", canSave);
        }
    }

    public void Save(BrowserWorkspace workspace)
    {
        lock (gate)
        {
            if (!canSave) throw new InvalidOperationException("State saving is suspended.");
            ValidateFilter(workspace.Applied);
            ValidateFilter(workspace.Draft);
            workspace.Window?.Validate();
            string active = Leaf("workspace.json");
            string previous = Leaf("workspace.previous.json");
            string leasePath = Leaf("workspace.lock");
            Directory.CreateDirectory(root);
            using var lease = new FileStream(leasePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            if (Hash(ReadOptional(active)) != expectedHash)
                throw new IOException("On-disk state changed after the window opened. File preserved.");
            if (ReadOptional(previous) is { } backup) Decode(backup);
            var data = new SavedWorkspace(workspace.Document.Pack(), workspace.Reference?.Pack(),
                workspace.Applied, workspace.Draft, workspace.SelectedId, workspace.Window,
                (workspace.HitsLayout ?? GridLayout.Default).Pack(), (workspace.CommonLayout ?? GridLayout.Default).Pack(),
                JsonSerializer.SerializeToElement(workspace.SourceFieldsExpanded),
                JsonSerializer.SerializeToElement(workspace.ComparisonExpanded));
            string? nextHash = null;
            AtomicFile.Write(active, stream =>
            {
                byte[] payload = JsonSerializer.SerializeToUtf8Bytes(data, Json);
                using var writer = new Utf8JsonWriter(stream);
                writer.WriteStartObject();
                writer.WriteNumber("Format", 1);
                writer.WriteString("Sha256", Hash(payload));
                writer.WritePropertyName("Workspace");
                writer.WriteRawValue(payload, skipInputValidation: true);
                writer.WriteEndObject();
            }, candidate =>
            {
                byte[] bytes = LocalFile.ReadBytes(candidate, Limit);
                Decode(bytes);
                nextHash = Hash(bytes);
            }, replace: expectedHash != null, backup: expectedHash != null ? previous : null, CancellationToken.None);
            expectedHash = nextHash;
        }
    }

    public static bool IsDataError(Exception ex) => ex is IOException or UnauthorizedAccessException
        or InvalidDataException or FormatException or ArgumentException or InvalidOperationException or JsonException
        or System.Security.SecurityException or NotSupportedException;
}
