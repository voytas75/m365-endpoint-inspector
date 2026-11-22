using System.Globalization;
using System.Text;
using System.Text.Json;

namespace M365.EndpointInspector.Domain;

// Observed Worldwide endpoint CSV schema; input is data, including notes and extra columns.
public static class WorldwideCsvImporter
{
    public const int ByteLimit = 4 * 1024 * 1024, RowLimit = 10000, ColumnLimit = 64, CellLimit = 65536;
    public const string Header = "id,serviceArea,serviceAreaDisplayName,urls,ips,tcpPorts,udpPorts,expressRoute,category,required,notes";
    private sealed record Cell(string Text, int Line, int Column);
    private static FormatException Error(Cell cell, string message) =>
        new($"CSV: row {cell.Line}, column {cell.Column}: {message}");

    public static EndpointSnapshot Import(ReadOnlySpan<byte> bytes, DateTimeOffset importedAt)
    {
        if (bytes.Length > ByteLimit) throw Error(new("", 1, 1), "The 4 MiB limit was exceeded.");
        if (bytes.StartsWith(Encoding.UTF8.Preamble)) bytes = bytes[3..];
        string text;
        try { text = new UTF8Encoding(false, true).GetString(bytes); }
        catch (DecoderFallbackException ex) { throw Error(new("", 1, 1), $"Invalid UTF-8 (byte {ex.Index + 1})."); }
        var rows = Read(text);
        if (rows.Count < 2) throw Error(new("", 1, 1), "A header and at least one record are required.");
        var header = rows[0];
        if (header.Any(c => c.Text == "Section") && header.Any(c => c.Text == "RowOrigin"))
            throw Error(header[0], "The application's BrowserCSV export (23 columns) is a different, unsupported import format.");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var cell in header)
            if (string.IsNullOrWhiteSpace(cell.Text) || !names.Add(cell.Text))
                throw Error(cell, "Empty or duplicate header.");
        foreach (string name in Header.Split(','))
            if (!names.Contains(name)) throw Error(header[0], $"Missing header '{name}'.");
        var payload = new List<Dictionary<string, object>>();
        var ids = new HashSet<int>();
        foreach (var row in rows.Skip(1))
        {
            if (row.Count != header.Count) throw Error(row[0], $"Expected {header.Count} columns, received {row.Count}.");
            var obj = new Dictionary<string, object>(StringComparer.Ordinal);
            for (int i = 0; i < header.Count; i++)
            {
                string name = header[i].Text, value = row[i].Text;
                bool required = name is "id" or "serviceArea" or "category" or "required";
                if (required && string.IsNullOrWhiteSpace(value)) throw Error(row[i], $"Field '{name}' is required.");
                if (Header.Split(',').Contains(name) && !required && value.Length == 0) continue;
                switch (name)
                {
                    case "id":
                        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int id) || id < 1)
                            throw Error(row[i], "ID must be a positive 32-bit integer.");
                        if (!ids.Add(id)) throw Error(row[i], $"Duplicate ID {id}.");
                        obj[name] = id; break;
                    case "required": case "expressRoute":
                        if (!bool.TryParse(value, out bool flag) || value != value.Trim())
                            throw Error(row[i], $"Field '{name}': only True/False allowed (case-insensitive).");
                        obj[name] = flag; break;
                    case "urls": case "ips": obj[name] = value.Split(',').Select(v => v.Trim()).ToArray(); break;
                    case "tcpPorts": case "udpPorts": obj[name] = string.Join(",", value.Split(',').Select(v => v.Trim())); break;
                    default: obj[name] = value; break;
                }
            }
            var element = JsonSerializer.SerializeToElement(obj);
            try { _ = new Endpoint(element, worldwide: true); }
            catch (FormatException ex)
            {
                int column = header.FindIndex(c => ex.Message.StartsWith(c.Text, StringComparison.Ordinal));
                throw Error(row[column < 0 ? 0 : column], "Invalid endpoint data: " + ex.Message);
            }
            payload.Add(obj);
        }
        return WorldwideImporter.ImportFile(JsonSerializer.SerializeToElement(payload), importedAt);
    }

    private static List<List<Cell>> Read(string text)
    {
        var rows = new List<List<Cell>>();
        var row = new List<Cell>();
        int pos = 0, line = 1, column = 1;
        char Take()
        {
            char c = text[pos++];
            if (c == '\n') { line++; column = 1; } else column++;
            return c;
        }
        while (pos < text.Length)
        {
            var start = new Cell("", line, column);
            var value = new StringBuilder();
            bool quoted = text[pos] == '"', closed = !quoted;
            if (quoted) Take();
            while (pos < text.Length)
            {
                char c = text[pos];
                if (quoted)
                {
                    if (c == '"')
                    {
                        Take();
                        if (pos < text.Length && text[pos] == '"') { Take(); value.Append('"'); }
                        else { closed = true; break; }
                    }
                    else value.Append(Take());
                }
                else
                {
                    if (c is ',' or '\r' or '\n') break;
                    if (c == '"') throw Error(new("", line, column), "Quote inside an unquoted field.");
                    value.Append(Take());
                }
                if (value.Length > CellLimit) throw Error(start, "The 65536-character field limit was exceeded.");
            }
            if (!closed) throw Error(start, "Unclosed quoted field.");
            row.Add(start with { Text = value.ToString() });
            if (row.Count > ColumnLimit) throw Error(start, "The 64-column limit was exceeded.");
            if (pos < text.Length && text[pos] == ',')
            {
                Take();
                if (pos == text.Length)
                {
                    if (row.Count == ColumnLimit) throw Error(start, "The 64-column limit was exceeded.");
                    row.Add(new("", line, column));
                }
                else continue;
            }
            else if (pos < text.Length)
            {
                if (text[pos] == '\r') { Take(); if (pos == text.Length || text[pos] != '\n') throw Error(new("", line, column), "Bare CR; CRLF or LF required."); }
                if (text[pos] != '\n') throw Error(new("", line, column), "Invalid character after closing a field.");
                Take();
            }
            rows.Add(row); row = [];
            if (rows.Count > RowLimit + 1) throw Error(start, "The 10000-record limit was exceeded.");
        }
        return rows;
    }
}
