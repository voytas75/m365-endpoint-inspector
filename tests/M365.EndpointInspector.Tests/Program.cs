using System.Text.Json;
using System.Text.Json.Nodes;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Tests;

internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "--csv-file")
            {
                var imported = WorldwideCsvImporter.Import(File.ReadAllBytes(args[1]), DateTimeOffset.UnixEpoch);
                Console.WriteLine($"CSV acceptance: count={imported.Entries.Count}; areas={string.Join(',', imported.Entries.Values.GroupBy(e => e.ServiceArea).Select(g => $"{g.Key}:{g.Count()}"))}; requiredTrue={imported.Entries.Values.Count(e => e.Required == true)}; urls={imported.Entries.Values.Sum(e => e.DnsPatterns.Count)}; ips={imported.Entries.Values.Sum(e => e.Networks.Count)}; version={imported.Version ?? "Unknown"}");
                return 0;
            }
            var previous = Load("synthetic-1.json");
            var current = Load("synthetic-2.json");
            (string Name, Action Run)[] cases =
            [
                ("simulated observed Worldwide CSV schema", CsvImportCases.Run),
                ("snapshot provenance and source", () => ImportCases(current)),
                ("known fields", () => SourceFilters(current)),
                ("nested JSON field paths", () => FieldCases(current)),
                ("TCP and UDP", () => Ports(current)),
                ("DNS labels and memberships", () => Domains(current)),
                ("Worldwide DNS wildcard labels", WorldwideDomains),
                ("IP addresses and overlapping networks", () => Networks(current)),
                ("service hits and Common dependencies", () => Services(current, previous)),
                ("snapshot changes and equality", () => Changes(current, previous)),
                ("synthetic result states", () => States(current)),
                ("malformed snapshots and queries", () => InvalidInput.Check(current))
            ];
            int failed = 0;
            foreach (var (name, run) in cases)
            {
                try
                {
                    run();
                    Console.WriteLine($"PASS {name}");
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.Error.WriteLine($"FAIL {name}: {ex}");
                }
            }
            Console.WriteLine($"{cases.Length - failed}/{cases.Length} groups passed.");
            return failed == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fixture setup failed: {ex}");
            return 1;
        }
    }

    private static EndpointSnapshot Load(string file) =>
        SyntheticImporter.Import(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", file)));

    internal static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException($"Expected [{expected}], got [{actual}].");
    }

    private static void Ids(IEnumerable<EndpointMatch> rows, params int[] expected)
    {
        Equal(string.Join(",", expected), string.Join(",", rows.Select(r => r.Entry.Id)));
    }

    private static void Hits(EndpointSnapshot current, EndpointFilter filter, params int[] expected) =>
        Ids(EndpointQuery.Select(current, filter).Hits, expected);

    private static FieldMatch At(string path, string json)
    {
        using var value = JsonDocument.Parse(json);
        return FieldMatch.Equal(path, value.RootElement);
    }

    private static void ImportCases(EndpointSnapshot current)
    {
        Equal("Synthetic", current.Provenance);
        Equal("synthetic-2", current.Version);
        Equal(new DateTimeOffset(2026, 1, 2, 12, 0, 0, TimeSpan.Zero), current.ImportedAt);
        Equal(5, current.Entries.Count);
        var entry = current.Entries[101];
        Equal("443, 5000-5004", entry.Source.GetProperty("tcpPorts").GetString());
        Equal("2001:db8:1::20/123", entry.Source.GetProperty("ips")[1].GetString());
        Equal("*.mail.example.test", entry.Source.GetProperty("urls")[0].GetString());
        Equal(2, entry.Source.GetProperty("extra").GetProperty("routing")[0].GetProperty("weight").GetInt32());
        Equal(false, current.Entries[103].Source.TryGetProperty("expressRoute", out _));
        Equal(JsonValueKind.Null, current.Entries[104].Source.GetProperty("expressRoute").ValueKind);
        Equal<bool?>(false, current.Entries[102].ExpressRoute);
        Hits(current, new(), 101, 102, 104, 103, 106);
    }

    private static void SourceFilters(EndpointSnapshot current)
    {
        Hits(current, new() { Id = 104 }, 104);
        Hits(current, new() { ServiceArea = ServiceArea.Exchange }, 101);
        Hits(current, new() { ServiceAreaDisplayName = "Synthetic Sites" }, 102);
        Hits(current, new() { Notes = "blue" }, 102);
        Hits(current, new() { Category = Category.Allow }, 102, 104);
        Hits(current, new() { Required = true }, 101, 103);
        Hits(current, new() { Required = false }, 102, 106);
        Hits(current, new() { ExpressRoute = true }, 101);
        Hits(current, new() { ExpressRoute = false }, 102);
        Hits(current, new() { Required = false, ExpressRoute = true });
        Hits(current, new() { Fields = [At("/tcpPorts", "\"443, 5000-5004\"")] }, 101);
        Hits(current, new() { Fields = [At("/udpPorts", "\"3478-3481\"")] }, 101);
        Hits(current, new() { Fields = [At("/ips/1", "\"2001:db8:1::28\"")] }, 102);
        Hits(current, new() { Fields = [At("/urls/0", "\"*.mail.example.test\"")] }, 101);
    }

    private static void FieldCases(EndpointSnapshot current)
    {
        Hits(current, new() { Fields = [At("/extra/routing/0/tag", "\"blue\"")] }, 101);
        Hits(current, new() { Fields = [At("/extra/routing/0", """{"weight":2.0,"tag":"blue"}""")] }, 101);
        Hits(current, new() { Fields = [At("/extra/a~1b/~0name", "\"scoped\"")] }, 101);
        Hits(current, new() { Fields = [At("/extra/routing/0/tag", "\"weight\"")] });
        Hits(current, new() { Fields = [At("/tag", "\"blue\"")] });
        Hits(current, new() { Fields = [At("/extra/routing/0/weight", "\"2\"")] });
        Hits(current, new() { Fields = [At("/extra/routing/00/weight", "2")] });
        Hits(current, new() { Fields = [At("/expressRoute", "null")] }, 104);
        Hits(current, new() { Fields = [FieldMatch.Missing("/expressRoute")] }, 103, 106);
        Hits(current, new() { Fields = [FieldMatch.Missing("/notes"), At("/category", "\"Allow\"")] }, 104);
    }

    private static void Ports(EndpointSnapshot current)
    {
        void Tcp(string value, params int[] ids) =>
            Hits(current, new() { Protocol = Transport.Tcp, Ports = value }, ids);
        void Udp(string value, params int[] ids) =>
            Hits(current, new() { Protocol = Transport.Udp, Ports = value }, ids);

        Tcp("4999");
        Tcp("5000", 101);
        Tcp("5002", 101);
        Tcp("5004", 101);
        Tcp("5005");
        Tcp("5004-5010", 101);
        Tcp("80, 443", 101, 102, 103);
        Tcp("3479");
        Udp("3478", 101);
        Udp("3479", 101);
        Udp("3481", 101);
        Udp("3482");
        Udp("5002", 102);
        Udp("1", 106);
        Udp("65535", 106);
        Hits(current, new() { Protocol = Transport.Tcp }, 101, 102, 103);
        Hits(current, new() { Protocol = Transport.Udp }, 101, 102, 106);
    }

    private static void Domains(EndpointSnapshot current)
    {
        void Host(string name, params int[] ids) => Hits(current, new() { Domain = name }, ids);
        Host("A.MaIl.Example.Test", 101);
        Host("a.b.mail.example.test.", 101);
        Host("mail.example.test");
        Host("notmail.example.test");
        Host("a.mail.example.test.invalid");
        Host("x.shared.example.test");
        var shared = EndpointQuery.Select(current, new() { Domain = "SHARED.example.TEST." });
        Ids(shared.Hits, 101, 102, 103);
        Equal<Category?>(Category.Optimize, shared.HighestPriorityCategory);
        Equal<Category?>(null, EndpointQuery.Select(current, new() { Domain = "absent.example.test" }).HighestPriorityCategory);
    }

    internal static EndpointSnapshot WorldwidePattern(string pattern)
    {
        var row = new { id = 1, serviceArea = "Common", category = "Allow", urls = new[] { pattern } };
        return WorldwideImporter.ImportFile(JsonSerializer.SerializeToElement(new[] { row }), DateTimeOffset.UnixEpoch);
    }

    private static void WorldwideDomains()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "simulated-worldwide-wildcard.json")));
        var reported = WorldwideImporter.ImportFile(json.RootElement, DateTimeOffset.UnixEpoch);
        Equal("autodiscover.*.onmicrosoft.com", reported.Entries[3].DnsPatterns[1]);
        foreach (string host in new[] { "autodiscover.tenant.onmicrosoft.com", "AUTODISCOVER.a.B.OnMicrosoft.Com." })
            Hits(reported, new() { Domain = host }, 3);
        foreach (string host in new[] { "autodiscover.onmicrosoft.com", "onmicrosoft.com", "x.autodiscover.t.onmicrosoft.com",
                     "notautodiscover.t.onmicrosoft.com", "autodiscover.t.notonmicrosoft.com",
                     "autodiscover.t.onmicrosoft.com.invalid", "t.autodiscover.onmicrosoft.com" })
            Hits(reported, new() { Domain = host });

        var suffix = WorldwidePattern("*CdN.Example.Test.");
        Equal("*CdN.Example.Test", suffix.Entries[1].DnsPatterns.Single());
        Hits(suffix, new() { Domain = "cdn.example.test" }, 1);
        Hits(suffix, new() { Domain = "nested.edgecdn.example.test" }, 1);
        Hits(suffix, new() { Domain = "example.test" });
        Hits(suffix, new() { Domain = "cdn.x.example.test" });

        var several = WorldwidePattern("start.*.mid.*.*");
        Hits(several, new() { Domain = "start.a.b.mid.c.d" }, 1);
        Hits(several, new() { Domain = "start.mid.a.b" });
        Hits(several, new() { Domain = "start.a.mid.b" });
        Hits(several, new() { Domain = "start.a.notmid.b.c" });
        Hits(several, new() { Domain = "a.start.b.mid.c.d" });

        var fixedLabels = WorldwidePattern("*cdn.*.mid.end");
        Hits(fixedLabels, new() { Domain = "edgecdn.a.mid.b.mid.end" }, 1);
        Hits(fixedLabels, new() { Domain = "cdn.mid.end" });
        Hits(fixedLabels, new() { Domain = "cdn.a.mid.b.end" });
    }

    private static void Networks(EndpointSnapshot current)
    {
        void Ip(string value, params int[] ids) => Hits(current, new() { AddressOrCidr = value }, ids);
        Ip("192.0.2.31");
        Ip("192.0.2.32", 101);
        Ip("192.0.2.40", 101, 102);
        Ip("192.0.2.63", 101);
        Ip("192.0.2.64");
        Ip("192.0.2.48/28", 101);
        Ip("192.0.2.0/24", 101, 102);
        Ip("192.0.2.64/27");
        Ip("0.0.0.0/0", 101, 102, 104);
        Ip("::ffff:192.0.2.40");
        Ip("2001:db8:1::1f");
        Ip("2001:db8:1::20", 101);
        Ip("2001:DB8:1::28/128", 101, 102);
        Ip("2001:db8:1::3f", 101);
        Ip("2001:db8:1::40");
        Ip("2001:db8:1::30/124", 101);
        Ip("2001:db8:1::/64", 101, 102);
        Ip("2001:db8:1::40/123");
        Ip("::/0", 101, 102, 104);
        Ip("198.51.100.9/32", 104);
    }

    private static void Services(EndpointSnapshot current, EndpointSnapshot previous)
    {
        var service = EndpointQuery.Select(current, new()
        {
            ServiceArea = ServiceArea.Exchange, Protocol = Transport.Tcp, Ports = "443",
            Required = true, Category = Category.Optimize, Change = ChangeKind.Modified, Result = ResultState.Success
        }, previous, [new(101, ResultState.Success)]);
        Ids(service.Hits, 101);
        Ids(service.CommonDependencies, 104, 103);

        var empty = EndpointQuery.Select(current, new() { ServiceArea = ServiceArea.Exchange, Id = 102 });
        Ids(empty.Hits);
        Ids(empty.CommonDependencies, 104, 103);
        Equal<Category?>(null, empty.HighestPriorityCategory);
        var common = EndpointQuery.Select(current, new() { ServiceArea = ServiceArea.Common, Required = true });
        Ids(common.Hits, 103);
        Ids(common.CommonDependencies);
        Ids(EndpointQuery.Select(current, new()).CommonDependencies);
    }

    private static void Changes(EndpointSnapshot current, EndpointSnapshot previous)
    {
        void Change(ChangeKind kind, params int[] ids) =>
            Ids(EndpointQuery.Select(current, new() { Change = kind }, previous).Hits, ids);
        Change(ChangeKind.Added, 106);
        Change(ChangeKind.Removed, 105);
        Change(ChangeKind.Modified, 101, 104);
        Change(ChangeKind.Unchanged, 102, 103);
        Ids(EndpointQuery.Select(current, new()
        {
            Change = ChangeKind.Removed, Domain = "retired.example.test", Protocol = Transport.Tcp, Ports = "8443"
        }, previous).Hits, 105);
        Ids(EndpointQuery.Select(previous, new() { Change = ChangeKind.Added }, current).Hits, 105);
        Ids(EndpointQuery.Select(current, new() { Change = ChangeKind.Unchanged }, current).Hits, 101, 102, 104, 103, 106);
        Equal<ChangeKind?>(null, EndpointQuery.Select(current, new() { Id = 101 }).Hits.Single().Change);

        // Change only a scalar, the order of an array, and membership in Common.
        var edited = JsonNode.Parse(current.Source.GetRawText())!;
        var entries = edited["endpoints"]!.AsArray();
        var sites = entries.Single(row => row!["id"]!.GetValue<int>() == 102)!;
        sites["required"] = true;
        var common = entries.Single(row => row!["id"]!.GetValue<int>() == 103)!;
        common["urls"] = new JsonArray("common.example.test", "shared.example.test");
        entries.Remove(entries.Single(row => row!["id"]!.GetValue<int>() == 104));
        var updated = SyntheticImporter.Import(edited.ToJsonString());
        Ids(EndpointQuery.Select(updated, new() { Change = ChangeKind.Modified }, current).Hits, 102, 103);
        var service = EndpointQuery.Select(updated, new() { ServiceArea = ServiceArea.Exchange }, current);
        Ids(service.CommonDependencies, 103);
        Ids(EndpointQuery.Select(updated, new() { ServiceArea = ServiceArea.Common, Change = ChangeKind.Removed }, current).Hits, 104);
    }

    private static void States(EndpointSnapshot current)
    {
        SyntheticResult[] samples =
        [
            new(101, ResultState.Success), new(101, ResultState.Failure), new(101, ResultState.Success),
            new(102, ResultState.Inconclusive), new(103, ResultState.NotTestable),
            new(104, ResultState.Skipped), new(104, ResultState.Cancelled)
        ];
        void State(ResultState state, params int[] ids) =>
            Ids(EndpointQuery.Select(current, new() { Result = state }, samples: samples).Hits, ids);
        State(ResultState.Success, 101);
        State(ResultState.Failure, 101);
        State(ResultState.Inconclusive, 102);
        State(ResultState.NotTestable, 103);
        State(ResultState.Skipped, 104);
        State(ResultState.Cancelled, 104);
        State(ResultState.Unmeasured, 106);
        var row = EndpointQuery.Select(current, new() { Id = 101 }, samples: samples).Hits.Single();
        Equal("Success,Failure", string.Join(",", row.Results));
    }
}
