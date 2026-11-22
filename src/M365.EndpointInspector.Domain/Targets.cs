using System.Globalization;
using System.Net;
using System.Net.Sockets;

namespace M365.EndpointInspector.Domain;

public sealed class PortRange
{
    public int First { get; }
    public int Last { get; }

    private PortRange(int first, int last) => (First, Last) = (first, last);

    public bool Overlaps(PortRange other) => First <= other.Last && other.First <= Last;

    public static IReadOnlyList<PortRange> ParseList(string text)
    {
        var ranges = new List<PortRange>();
        foreach (string part in text.Split(','))
        {
            string[] bounds = part.Trim(' ').Split('-');
            if (bounds.Length > 2)
                throw new FormatException($"Invalid port range '{part}'.");

            int first = DecimalField.Parse(bounds[0], 1, 65535, "port");
            int last = bounds.Length == 1 ? first : DecimalField.Parse(bounds[1], 1, 65535, "port");
            if (last < first)
                throw new FormatException($"Reversed port range '{part}'.");
            ranges.Add(new PortRange(first, last));
        }
        return ranges.AsReadOnly();
    }
}

public sealed class IpPrefix
{
    private readonly IPNetwork network;

    private IpPrefix(IPNetwork network) => this.network = network;

    // An individual address is a one-address prefix. Keep the two families distinct,
    // including IPv4-mapped IPv6 addresses.
    public bool Overlaps(IpPrefix other) =>
        network.BaseAddress.AddressFamily == other.network.BaseAddress.AddressFamily &&
        (network.Contains(other.network.BaseAddress) || other.network.Contains(network.BaseAddress));

    public static IpPrefix Parse(string text)
    {
        string[] parts = text.Split('/');
        string literal = parts[0];
        if (parts.Length > 2 || literal.Length == 0 ||
            literal.Any(c => !(char.IsAsciiHexDigit(c) || c is '.' or ':')))
            throw new FormatException($"Invalid IP address or CIDR '{text}'.");

        // IPAddress.TryParse also accepts legacy abbreviated, integer and octal IPv4 syntax.
        if (literal.Contains('.') || !literal.Contains(':'))
        {
            string dotted = literal[(literal.LastIndexOf(':') + 1)..];
            string[] octets = dotted.Split('.');
            if (octets.Length != 4)
                throw new FormatException($"Invalid dotted IPv4 address '{dotted}'.");
            foreach (string octet in octets)
                DecimalField.Parse(octet, 0, 255, "IPv4 octet");
        }

        if (!IPAddress.TryParse(literal, out var address))
            throw new FormatException($"Invalid IP address '{literal}'.");

        int width = address.AddressFamily == AddressFamily.InterNetwork ? 32 : 128;
        int length = parts.Length == 1 ? width : DecimalField.Parse(parts[1], 0, width, "prefix length");
        // A CIDR must name its network base. Reject host bits before constructing IPNetwork.
        byte[] bytes = address.GetAddressBytes();
        for (int bit = length; bit < width; bit++)
        {
            if ((bytes[bit / 8] & (0x80 >> (bit % 8))) != 0)
                throw new FormatException($"Nonzero host bits in '{text}'.");
        }
        return new IpPrefix(new IPNetwork(address, length));
    }
}

internal static class DecimalField
{
    public static int Parse(string text, int min, int max, string field)
    {
        if (text.Length == 0 || (text.Length > 1 && text[0] == '0') ||
            !text.All(char.IsAsciiDigit) ||
            !int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) ||
            value < min || value > max)
            throw new FormatException($"Invalid {field} '{text}' (expected {min}..{max}).");
        return value;
    }
}

internal static class DnsPattern
{
    public static string Normalize(string text, bool allowWildcard, bool worldwide = false)
    {
        string name = text.EndsWith('.') ? text[..^1] : text;
        if (name.Length > 253 || !ValidLabels(name, allowWildcard, worldwide))
            throw new FormatException($"Invalid DNS name or pattern '{text}'.");
        return name;
    }

    private static bool ValidLabels(string name, bool allowWildcard, bool worldwide)
    {
        string[] labels = name.Split('.');
        bool hasLiteral = false;
        for (int i = 0; i < labels.Length; i++)
        {
            string label = labels[i];
            if (allowWildcard && label == "*" && (worldwide || i == 0))
                continue;
            if (label.Length is < 1 or > 63 || label[0] == '-' || label[^1] == '-')
                return false;

            int start = worldwide && allowWildcard && i == 0 && label[0] == '*' ? 1 : 0;
            foreach (char c in label.AsSpan(start))
                if (!char.IsAsciiLetterOrDigit(c) && c != '-')
                    return false;
            hasLiteral = true;
        }
        return hasLiteral;
    }

    public static bool Matches(string pattern, string host)
    {
        if (!pattern.Contains('*'))
            return host.Equals(pattern, StringComparison.OrdinalIgnoreCase);

        string[] hosts = host.Split('.');
        // Each position counts the host labels consumed by the pattern so far.
        var previous = new bool[hosts.Length + 1];
        var next = new bool[previous.Length];
        previous[0] = true;
        foreach (string label in pattern.Split('.'))
        {
            next[0] = false;
            for (int count = 1; count < next.Length; count++)
            {
                next[count] = label switch
                {
                    // Consume the first label for this star, or extend it by one label.
                    "*" => previous[count - 1] || next[count - 1],
                    // Only the initial label can have this form. Preserve the leading suffix match.
                    ['*', ..] => hosts[count - 1].AsSpan().EndsWith(label.AsSpan(1), StringComparison.OrdinalIgnoreCase),
                    _ => previous[count - 1] && hosts[count - 1].Equals(label, StringComparison.OrdinalIgnoreCase)
                };
            }
            (previous, next) = (next, previous);
        }
        return previous[^1];
    }
}
