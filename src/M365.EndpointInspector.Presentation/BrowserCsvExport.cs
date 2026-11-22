using System.Globalization;
using System.Text;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Presentation;

// Owns serialized bytes; neither dialog callbacks nor later workspace changes can alter them.
public sealed class BrowserCsvExport
{
    public const int RowLimit = 10000;
    public const int ByteLimit = 16 * 1024 * 1024;
    private readonly byte[] bytes;
    public byte[] CopyBytes() => (byte[])bytes.Clone();
    public string Message { get; }
    private BrowserCsvExport(byte[] bytes, string message) => (this.bytes, Message) = (bytes, message);

    public static BrowserCsvExport Capture(BrowserWorkspace workspace, IEnumerable<BrowserRow> hits,
        IEnumerable<BrowserRow> common, bool showCommon, BrowserFilter draft)
    {
        if (workspace.Document.Snapshot == null) throw new InvalidDataException("Brak danych do eksportu.");
        using var output = new MemoryStream();
        output.Write(Encoding.UTF8.Preamble);
        void Record(IEnumerable<string?> cells)
        {
            var values = cells.ToArray();
            if (values.Any(value => value?.Length > ByteLimit))
                throw new InvalidDataException("Eksport przekracza limit 16 MiB.");
            string line = string.Join(",", values.Select(Escape)) + "\r\n";
            int size = Encoding.UTF8.GetByteCount(line);
            if (output.Length + size > ByteLimit) throw new InvalidDataException("Eksport przekracza limit 16 MiB.");
            output.Write(Encoding.UTF8.GetBytes(line));
        }
        Record(["Section", "ID", "Obszar", "Kategoria", "Wymagany", "Domeny", "Adresy IP / CIDR",
            "Porty TCP", "Porty UDP", "Nazwa usługi", "ExpressRoute", "Zmiana", "RowOrigin", "SourceKind",
            "Provenance", "Generation", "Version", "TimeBasis", "ImportedAt", "OpenedAt",
            "ObservationSource", "VersionObservedAt", "RetrievedAt"]);
        int count = 0;
        foreach (var (rows, section) in new[] { (hits, "Hits"), (showCommon ? common : [], "Common") })
        foreach (var row in rows)
        {
            if (++count > RowLimit) throw new InvalidDataException("Eksport przekracza limit 10000 wierszy.");
            bool removed = row.IsRemoved;
            var document = removed ? workspace.Reference : workspace.Document;
            var snapshot = document?.Snapshot ?? throw new InvalidDataException("Brak pochodzenia wiersza odniesienia.");
            var origin = document!.Origin;
            var observation = snapshot.Observation;
            string basis = document.IsSynthetic ? "SyntheticDeclaration" : observation == null ? "Unknown" :
                origin?.Kind == BrowserSource.File ? "FileDeclaredObservation" : "ObservedBeforeRetrieval";
            Record([section, row.Entry.Id.ToString(CultureInfo.InvariantCulture), row.Entry.ServiceArea.ToString(),
                row.Entry.Category.ToString(), row.Required, row.Domains, row.Addresses, row.Tcp, row.Udp,
                row.Entry.ServiceAreaDisplayName, row.ExpressRoute, row.Change, removed ? "Reference" : "Current",
                origin?.Kind.ToString() ?? "Unknown", (origin?.Csv == true ? "FileCsv" : snapshot.Provenance), origin?.Generation?.ToString() ?? "Unknown",
                snapshot.Version ?? "Unknown", basis, Time(snapshot.ImportedAt), Time(origin?.OpenedAt),
                observation?.Source.ToString() ?? "Unknown", Time(observation?.VersionObservedAt), Time(observation?.RetrievedAt)]);
        }
        if (count == 0) throw new InvalidDataException("Brak wyświetlonych wierszy do eksportu.");
        bool invalid = BrowserResults.Select(workspace.Document, draft, workspace.Reference).FilterError != null;
        return new(output.ToArray(), invalid
            ? "Nie zastosowano błędnego filtra roboczego. Wyeksportowano ostatni zastosowany widok."
            : "Wyeksportowano ostatni zastosowany widok (Trafienia i widoczne Common).");
    }

    private static string Time(DateTimeOffset? time) => time?.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture) ?? "Unknown";
    private static string Escape(string? value)
    {
        value ??= "";
        // Quote escaping alone does not prevent spreadsheet evaluation. Preserve original text after prefix.
        int first = 0;
        while (first < value.Length && (char.IsWhiteSpace(value[first]) || char.IsControl(value[first]) ||
            char.GetUnicodeCategory(value[first]) == UnicodeCategory.Format)) first++;
        if (first < value.Length && "=+-@".Contains(value[first])) value = "'" + value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
