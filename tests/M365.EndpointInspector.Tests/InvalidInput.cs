using System.Text.Json;
using System.Text.Json.Nodes;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Tests;

internal static class InvalidInput
{
    private const string Entry = """{"id":1,"serviceArea":"Common","category":"Allow","urls":["a.example.test"]}""";
    private const string Metadata = """{"provenance":"Synthetic","version":"synthetic-test","importedAt":"2026-01-01T00:00:00Z","endpoints":""";
    private static string Wrap(string entries) => Metadata + "[" + entries + "]}";

    private static void Reject<T>(Action action, string context) where T : Exception
    {
        try
        {
            action();
        }
        catch (T ex)
        {
            if (!ex.Message.Contains(context, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Expected '{context}' in the error: {ex.Message}", ex);
            return;
        }
        throw new InvalidOperationException($"Expected {typeof(T).Name} containing '{context}'.");
    }

    private static void Invalid(string json, string context) =>
        Reject<FormatException>(() => SyntheticImporter.Import(json), context);

    public static void Check(EndpointSnapshot current)
    {
        string valid = Wrap(Entry);
        var minimal = SyntheticImporter.Import(valid);
        Program.Equal(1, minimal.Entries.Count);
        Program.Equal<bool?>(null, minimal.Entries[1].Required);
        Program.Equal(0, SyntheticImporter.Import(Wrap("")).Entries.Count);
        var addressOnly = JsonNode.Parse(Entry)!.AsObject();
        addressOnly["urls"] = null;
        addressOnly["ips"] = new JsonArray("203.0.113.7");
        addressOnly["udpPorts"] = null;
        var one = SyntheticImporter.Import(Wrap(addressOnly.ToJsonString())).Entries[1];
        Program.Equal(0, one.DnsPatterns.Count);
        Program.Equal(1, one.Networks.Count);
        Program.Equal(0, one.UdpPorts.Count);
        Invalid("{", "JSON");
        Invalid("[]", "object");
        Invalid(Wrap("null"), "endpoints[0]");
        Invalid(Wrap(Entry + "," + Entry), "Duplicate id");
        Invalid(valid.Replace("\"id\":1", "\"id\":1,\"id\":2"), "Duplicate property");
        Invalid(valid.Replace("\"id\":1", "\"extra\":[{\"x\":1,\"x\":2}],\"id\":1"), "Duplicate property");

        foreach (string field in new[] { "provenance", "version", "importedAt", "endpoints" })
        {
            var envelope = JsonNode.Parse(valid)!.AsObject();
            envelope.Remove(field);
            Invalid(envelope.ToJsonString(), field);
        }
        Invalid(valid.Replace("\"Synthetic\"", "\"Official\""), "provenance");
        Invalid(valid.Replace("\"synthetic-test\"", "null"), "version");
        Invalid(valid.Replace("\"synthetic-test\"", "\" \""), "version");
        Invalid(valid.Replace("T00:00:00Z", "T00:00:00"), "importedAt");
        Invalid(valid.Replace("2026-01-01", "2026-99-99"), "importedAt");
        Invalid(Metadata + "{}}", "endpoints");

        foreach (string field in new[] { "id", "serviceArea", "category", "urls" })
        {
            var row = JsonNode.Parse(Entry)!.AsObject();
            row.Remove(field);
            Invalid(Wrap(row.ToJsonString()), field);
        }

        void BadField(string field, string jsonValue)
        {
            var row = JsonNode.Parse(Entry)!.AsObject();
            row[field] = JsonNode.Parse(jsonValue);
            Invalid(Wrap(row.ToJsonString()), field);
        }

        foreach (string id in new[] { "0", "-1", "1.5", "2147483648", "\"1\"", "null" })
            BadField("id", id);
        foreach (string category in new[] { "\"Other\"", "\"allow\"", "\"1\"", "1", "null" })
            BadField("category", category);
        BadField("serviceArea", "\"Example\"");
        BadField("serviceAreaDisplayName", "false");
        BadField("required", "\"true\"");
        BadField("expressRoute", "0");
        BadField("notes", "[]");
        BadField("urls", "\"a.example.test\"");
        BadField("urls", "[null]");
        BadField("urls", "[]");
        BadField("ips", "{}");
        BadField("ips", "[42]");
        BadField("tcpPorts", "443");
        BadField("udpPorts", "[53]");

        foreach (string value in new[] { "", "0", "65536", "2-1", "1-65536", "2-", "-2", "1-2-3",
                     "443,", ",443", "443,,80", "4 43", "+443", "0x50", "1.5", "０１" })
        {
            BadField("tcpPorts", JsonSerializer.Serialize(value));
            Reject<FormatException>(() => EndpointQuery.Select(current, new() { Protocol = Transport.Udp, Ports = value }), "port");
        }

        foreach (string value in new[] { "192.0.2.256", "192.0.2.1/33", "2001:db8::/129",
                     "192.0.2.3/24", "2001:db8::1/64", "127.1", "010.0.0.1", "3221225985",
                     "2001:::1", "fe80::1%2", "[2001:db8::1]", "192.0.2.0/", "192.0.2.0/-1",
                     "192.0.2.0/+1", "192.0.2.0/24/1", "::ffff:192.0.2.01" })
            BadField("ips", JsonSerializer.Serialize(new[] { value }));

        foreach (string value in new[] { "", "example..test", "*example.test", "a.*.example.test",
                     "*.example.test..", "-bad.example.test", "https://a.example.test", new string('a', 64) + ".test" })
            BadField("urls", JsonSerializer.Serialize(new[] { value }));

        foreach (string value in new[] { "*", "*.*", "**.test", "*c*dn.test", "a.b*c.test", "a.*cdn.test",
                     "a.*..test", "a.*.bad_.test", "a.*.-bad", "a.*.bad-", "a.*.test..",
                     "*" + new string('a', 63) + ".test", string.Join('.', Enumerable.Repeat(new string('a', 63), 4)) })
            Reject<FormatException>(() => Program.WorldwidePattern(value), "endpoints[0]: urls[0]: Invalid DNS");

        foreach (string host in new[] { "autodiscover.*.onmicrosoft.com", "autodiscover..onmicrosoft.com",
                     ".autodiscover.t.onmicrosoft.com", "autodiscover.t-.onmicrosoft.com", "x*cdn.example.test" })
            Reject<FormatException>(() => EndpointQuery.Select(current, new() { Domain = host }), "DNS");

        Reject<FormatException>(() => EndpointQuery.Select(current, new() { AddressOrCidr = "::/129" }), "prefix");
        Reject<FormatException>(() => EndpointQuery.Select(current, new() { Domain = "*.example.test" }), "DNS");
        Reject<FormatException>(() => FieldMatch.Missing("extra/name"), "path");
        Reject<FormatException>(() => FieldMatch.Missing("/extra~2"), "escape");
        Reject<FormatException>(() => FieldMatch.Missing("/extra~"), "escape");
        Reject<ArgumentException>(() => EndpointQuery.Select(current, new() { Ports = "53" }), "protocol");
        Reject<ArgumentOutOfRangeException>(() => EndpointQuery.Select(current, new() { Category = (Category)99 }), "Category");
        Reject<ArgumentOutOfRangeException>(() => EndpointQuery.Select(current, new() { Id = 0 }), "Id");
        Reject<ArgumentException>(() => EndpointQuery.Select(current, new() { Change = ChangeKind.Added }), "previous");
        Reject<ArgumentException>(() => EndpointQuery.Select(current, new(), samples: [new(999, ResultState.Skipped)]), "id");
        Reject<ArgumentException>(() => EndpointQuery.Select(current, new(), samples: [new(101, ResultState.Unmeasured)]), "state");
    }
}
