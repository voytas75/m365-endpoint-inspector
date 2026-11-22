using System.Text.Json;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Presentation;

public sealed record BrowserFilter
{
    public ServiceArea? Area { get; init; }
    public Category? Category { get; init; }
    public bool? Required { get; init; }
    public Transport? Protocol { get; init; }
    public string Ports { get; init; } = "";
    public string Domain { get; init; } = "";
    public string Address { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Notes { get; init; } = "";
    public string Pointer { get; init; } = "";
    public string JsonValue { get; init; } = "";
    public ChangeKind? Change { get; init; }

    public IReadOnlyList<FilterFieldError> Validate(BrowserDocument document, BrowserDocument? reference = null)
    {
        var errors = new List<FilterFieldError>();
        void Check(FilterField field, string code, string message, Action validate)
        {
            try { validate(); }
            catch (Exception ex) when (ex is FormatException or ArgumentException or JsonException)
            { errors.Add(new(field, code, message)); }
        }
        if (Ports.Trim().Length > 0)
        {
            if (Protocol == null)
                errors.Add(new(FilterField.Protocol, "required", "Select TCP or UDP to apply a port filter."));
            Check(FilterField.Ports, "syntax", $"Correct the port or range “{Ports}”: numbers 1–65535, e.g. 443, 5000-5004; start must not exceed end.",
                () => PortRange.ParseList(Ports.Trim()));
        }
        if (Address.Trim().Length > 0)
            Check(FilterField.Address, "syntax", "Enter a complete IPv4 or IPv6 address, or CIDR with a network address (no host bits), e.g. 192.0.2.32/27; IPv4 prefix: 0–32, IPv6: 0–128.",
                () => IpPrefix.Parse(Address.Trim()));
        if (Domain.Trim().Length > 0 && document.Snapshot is { } snapshot)
        {
            if (!UsesPublishedPattern)
                // EndpointQuery exposes the canonical Domain DNS validator.
                Check(FilterField.Domain, "syntax", "Enter a DNS name without spaces or a URL, e.g. mail.example.test, or a complete pattern from the open data.",
                    () => EndpointQuery.Select(snapshot, new EndpointFilter { Domain = Domain.Trim() }));
            else
            {
                string pattern = Domain.Trim();
                if (pattern.EndsWith('.')) pattern = pattern[..^1];
                bool HasPattern(Endpoint entry) => entry.DnsPatterns.Contains(pattern, StringComparer.OrdinalIgnoreCase);
                if (!snapshot.Entries.Values.Any(HasPattern) && reference?.Snapshot?.Entries.Values.Any(HasPattern) != true)
                    errors.Add(new(FilterField.Domain, "published-pattern", "Enter the complete pattern exactly as in the open data, or a concrete DNS name, e.g. mail.example.test."));
            }
        }
        if (Pointer.Length > 0)
            Check(FilterField.Pointer, "syntax", "Start the path with /, e.g. /extra/routing/0/tag. Encode ~ as ~0 and / within a name as ~1.",
                () => FieldMatch.Missing(Pointer));
        else if (!string.IsNullOrWhiteSpace(JsonValue))
            errors.Add(new(FilterField.Pointer, "required", "Enter a JSON Pointer path for this value, e.g. /extra/tag."));
        if (!string.IsNullOrWhiteSpace(JsonValue))
            Check(FilterField.JsonValue, "syntax", "Enter a valid JSON value: quoted text, e.g. \"blue\", a number, true or null. An empty field means the property is missing.",
                () => { using var value = JsonDocument.Parse(JsonValue); });
        if (Change != null && reference == null)
            errors.Add(new(FilterField.Change, "reference-required", "Open a Synthetic reference snapshot or select All changes."));
        return errors.AsReadOnly();
    }

    internal bool UsesPublishedPattern => Domain.Contains('*');

    internal EndpointFilter ToQuery()
    {
        if (Ports.Trim().Length > 0 && Protocol is null)
            throw new FormatException("Select TCP or UDP for a port query.");
        return new EndpointFilter
        {
            ServiceArea = Area, Category = Category, Required = Required, Protocol = Protocol,
            Ports = Optional(Ports.Trim()), Domain = UsesPublishedPattern ? null : Optional(Domain.Trim()),
            AddressOrCidr = Optional(Address.Trim()),
            ServiceAreaDisplayName = Optional(DisplayName), Notes = Optional(Notes),
            Fields = FieldCriteria(), Change = Change
        };
    }

    private IReadOnlyList<FieldMatch> FieldCriteria()
    {
        if (Pointer.Length == 0)
        {
            if (!string.IsNullOrWhiteSpace(JsonValue))
                throw new FormatException("Specify a JSON Pointer path for this value.");
            return [];
        }
        if (string.IsNullOrWhiteSpace(JsonValue)) return [FieldMatch.Missing(Pointer)];
        using var doc = JsonDocument.Parse(JsonValue);
        return [FieldMatch.Equal(Pointer, doc.RootElement)];
    }

    private static string? Optional(string text) => text.Length == 0 ? null : text;
}

public enum FilterField { Protocol, Ports, Domain, Address, Pointer, JsonValue, Change }

public sealed record FilterFieldError(FilterField Field, string Code, string Message);
