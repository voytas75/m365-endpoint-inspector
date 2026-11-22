using System.Text;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Tests;

internal static class CsvImportCases
{
    // Synthetic data simulating the observed schema, never a production response.
    internal const string Sample = "1,Common,Symulacja,\" *.example.test, test.example.test\",\"192.0.2.0/24, 2001:db8::/32\",\"443, 5000-5004\",,False,Allow,True,\"Zażółć, \"\"tekst\"\"\nciąg dalszy\"";
    internal static void Run()
    {
        EndpointSnapshot Read(string s) => WorldwideCsvImporter.Import(Encoding.UTF8.GetBytes(s), DateTimeOffset.UnixEpoch);
        string csv = WorldwideCsvImporter.Header + "\r\n" + Sample;
        var snapshot = Read("\ufeff" + csv);
        Program.Equal(1, snapshot.Entries.Count);
        Program.Equal(2, snapshot.Entries[1].DnsPatterns.Count);
        Program.Equal(2, snapshot.Entries[1].Networks.Count);
        Program.Equal<bool?>(true, snapshot.Entries[1].Required);
        Program.Equal(false, snapshot.Entries[1].Source.TryGetProperty("udpPorts", out _));
        Program.Equal("Zażółć, \"tekst\"\nciąg dalszy", snapshot.Entries[1].Notes);
        Program.Equal("Zażółć, \"tekst\"\r\nciąg dalszy", Read(csv.Replace("\nciąg", "\r\nciąg")).Entries[1].Notes);
        Read(csv + "\r\n");
        var extra = Read(WorldwideCsvImporter.Header + ",future\n" + Sample + ",opaque");
        Program.Equal("opaque", extra.Entries[1].Source.GetProperty("future").GetString());
        // Flexible header order.
        Read("future," + WorldwideCsvImporter.Header + "\nopaque," + Sample);
        void Reject(string s)
        {
            try { Read(s); throw new InvalidOperationException("Accepted invalid CSV"); }
            catch (FormatException ex) { if (!ex.Message.Contains("row") || !ex.Message.Contains("column")) throw; }
        }
        Reject(csv + ",extra");
        Reject(csv.Replace("Common", "UnknownArea"));
        Reject(csv.Replace("Allow", "UnknownCategory"));
        Reject(csv.Replace("False", "yes"));
        Reject(csv.Replace("True", "1")); Reject(csv.Replace("True", ""));
        Reject(csv.Replace("5000-5004", "70000")); Reject(csv.Replace("*.example.test", "bad*pattern.example.test"));
        Reject(csv.Replace("192.0.2.0/24", "192.0.2.999")); Reject(csv.Replace("1,Common", "0,Common"));
        Reject(csv + "\n" + Sample); Reject(csv.Replace("notes", "id")); Reject(csv.Replace(",notes", ""));
        Reject(csv + "\"oops"); Reject(csv.Replace("ciąg dalszy\"", "ciąg dalszy"));
        Reject(WorldwideCsvImporter.Header + "\n" + Sample.Replace("Symulacja", new string('a', WorldwideCsvImporter.CellLimit + 1)));
        Reject(new string('a', WorldwideCsvImporter.ByteLimit + 1));
        Reject(string.Join(',', Enumerable.Range(0, 65).Select(i => "x" + i)) + "\n1");
        Reject(WorldwideCsvImporter.Header + "\n" + string.Join('\n', Enumerable.Repeat(Sample, 10001)));
        Reject("Section,ID,RowOrigin\nHits,1,Current");
        try { WorldwideCsvImporter.Import([0xff], DateTimeOffset.UnixEpoch); throw new Exception("UTF8 accepted"); }
        catch (FormatException) { }
    }
}
