using System.Text.Json;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Presentation;

public sealed class BrowserResults
{
    public IReadOnlyList<BrowserRow> Hits { get; }
    public IReadOnlyList<BrowserRow> Common { get; }
    public bool ShowCommon { get; }
    public string Summary { get; }
    public string? FilterError { get; }
    public IReadOnlyList<FilterFieldError> FieldErrors { get; private init; } = [];
    public string CommonHeading => $"Common dependencies ({Common.Count})";

    private BrowserResults(string summary, BrowserRow[]? hits = null, BrowserRow[]? common = null,
        bool showCommon = false, string? error = null)
        => (Summary, Hits, Common, ShowCommon, FilterError) = (summary, Array.AsReadOnly(hits ?? []), Array.AsReadOnly(common ?? []), showCommon, error);

    public static BrowserResults Select(BrowserDocument document, BrowserFilter filter,
        BrowserDocument? reference = null)
    {
        if (document.Snapshot is not { } snapshot) return new("No data");
        var errors = filter.Validate(document, reference);
        if (errors.Count > 0)
            return new("Filter error", error: string.Join(" ", errors.Select(error => error.Message))) { FieldErrors = errors };
        try
        {
            if (reference != null && !(document.IsSynthetic && reference.IsSynthetic))
                throw new ArgumentException("Comparison requires two Synthetic snapshots.");
            if (filter.Change != null && reference == null)
                throw new ArgumentException("Open a Synthetic reference snapshot to filter changes.");

            var selected = EndpointQuery.Select(snapshot, filter.ToQuery(), reference?.Snapshot);
            IEnumerable<EndpointMatch> matches = selected.Hits;
            if (filter.UsesPublishedPattern)
            {
                string pattern = filter.Domain.Trim();
                if (pattern.EndsWith('.')) pattern = pattern[..^1];
                bool HasPattern(Endpoint entry) => entry.DnsPatterns.Contains(pattern, StringComparer.OrdinalIgnoreCase);
                if (!snapshot.Entries.Values.Any(HasPattern) && reference?.Snapshot?.Entries.Values.Any(HasPattern) != true)
                    throw new FormatException("Enter a concrete domain or a complete pattern from the open data.");
                matches = matches.Where(match => HasPattern(match.Entry));
            }
            var hits = matches.Select(match => new BrowserRow(match, (match.Change == ChangeKind.Removed ? reference : document)?.Origin?.Csv == true)).ToArray();
            var common = selected.CommonDependencies.Select(match => new BrowserRow(match, (match.Change == ChangeKind.Removed ? reference : document)?.Origin?.Csv == true)).ToArray();
            string summary = hits.Length > 0
                ? $"Hits ({hits.Length})   ·   Highest category: {hits[0].Entry.Category}"
                : snapshot.Entries.Count == 0 && reference?.Snapshot?.Entries.Count is null or 0
                    ? "Empty snapshot (0 entries)" : "No hits (0)";
            return new(summary, hits, common, filter.Area is not (null or ServiceArea.Common));
        }
        catch (Exception ex) when (ex is FormatException or ArgumentException or JsonException or InvalidOperationException)
        {
            return new("Filter error", error: ex.Message);
        }
    }
}

public sealed class BrowserRow(EndpointMatch match, bool convertedCsv = false)
{
    private static readonly JsonSerializerOptions PreviewOptions = new() { WriteIndented = true };

    public Endpoint Entry => match.Entry;
    public bool IsRemoved => match.Change == ChangeKind.Removed;
    public string Required => YesNo(Entry.Required);
    public string ExpressRoute => YesNo(Entry.ExpressRoute);
    public string Domains => string.Join(", ", Values("urls"));
    public string Addresses => string.Join(", ", Values("ips"));
    public string? Tcp => Text("tcpPorts");
    public string? Udp => Text("udpPorts");
    public string Change => ChangeLabel(match.Change);
    // Preserve the established CSV values independently of English UI labels.
    // Keep Change-column ordering stable across the UI language change.
    public string ChangeSortKey => CsvChange;
    internal string CsvRequired => CsvYesNo(Entry.Required);
    internal string CsvExpressRoute => CsvYesNo(Entry.ExpressRoute);
    internal string CsvChange => match.Change switch
    {
        ChangeKind.Added => "Dodany", ChangeKind.Modified => "Zmieniony",
        ChangeKind.Removed => "Usunięty", ChangeKind.Unchanged => "Bez zmian", _ => ""
    };
    private static string CsvYesNo(bool? value) => value switch { true => "Tak", false => "Nie", null => "—" };
    public string RepresentationLabel => convertedCsv ? "JSON converted from CSV (not original JSON)" : "Selected entry source JSON";
    public string RawJson => (convertedCsv ? "JSON representation converted from CSV (not original JSON).\n" : "") + JsonSerializer.Serialize(Entry.Source, PreviewOptions);
    public string Title => $"ID {Entry.Id}   ·   {Entry.ServiceArea}   ·   {Entry.Category}" +
        (match.Change == null ? "" : $"   ·   {Change}");
    public static string ChangeLabel(ChangeKind? change) => change switch
    {
        ChangeKind.Added => "Added", ChangeKind.Modified => "Modified",
        ChangeKind.Removed => "Removed", ChangeKind.Unchanged => "Unchanged", _ => ""
    };

    public IEnumerable<TargetChoice> SearchableTargets =>
        Values("urls").Select(name => new TargetChoice(name, false))
            .Concat(Values("ips").Select(ip => new TargetChoice(ip, true)));

    private string? Text(string key) => Entry.Source.TryGetProperty(key, out var value) ? value.GetString() : null;
    private IEnumerable<string> Values(string key) =>
        Entry.Source.TryGetProperty(key, out var list) && list.ValueKind == JsonValueKind.Array
            ? list.EnumerateArray().Select(item => item.GetString()!) : [];
    private static string YesNo(bool? value) => value switch { true => "Yes", false => "No", null => "—" };
}

public sealed record TargetChoice(string Value, bool IsAddress)
{
    public BrowserFilter AllMemberships() => IsAddress ? new() { Address = Value } : new() { Domain = Value };
}
