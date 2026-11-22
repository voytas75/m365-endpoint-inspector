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
                errors.Add(new(FilterField.Protocol, "required", "Wybierz TCP lub UDP, aby zastosować filtr portów."));
            Check(FilterField.Ports, "syntax", $"Popraw port lub zakres „{Ports}”: liczby 1–65535, np. 443, 5000-5004; początek nie większy niż koniec.",
                () => PortRange.ParseList(Ports.Trim()));
        }
        if (Address.Trim().Length > 0)
            Check(FilterField.Address, "syntax", "Podaj pełny IPv4 lub IPv6 albo CIDR z adresem sieci (bez bitów hosta), np. 192.0.2.32/27; prefiks IPv4: 0–32, IPv6: 0–128.",
                () => IpPrefix.Parse(Address.Trim()));
        if (Domain.Trim().Length > 0 && document.Snapshot is { } snapshot)
        {
            if (!UsesPublishedPattern)
                // EndpointQuery exposes the canonical Domain DNS validator.
                Check(FilterField.Domain, "syntax", "Podaj nazwę DNS bez spacji i adresu URL, np. mail.example.test, albo pełny wzorzec z otwartych danych.",
                    () => EndpointQuery.Select(snapshot, new EndpointFilter { Domain = Domain.Trim() }));
            else
            {
                string pattern = Domain.Trim();
                if (pattern.EndsWith('.')) pattern = pattern[..^1];
                bool HasPattern(Endpoint entry) => entry.DnsPatterns.Contains(pattern, StringComparer.OrdinalIgnoreCase);
                if (!snapshot.Entries.Values.Any(HasPattern) && reference?.Snapshot?.Entries.Values.Any(HasPattern) != true)
                    errors.Add(new(FilterField.Domain, "published-pattern", "Wpisz pełny wzorzec dokładnie z otwartych danych albo konkretną nazwę DNS, np. mail.example.test."));
            }
        }
        if (Pointer.Length > 0)
            Check(FilterField.Pointer, "syntax", "Zacznij ścieżkę od /, np. /extra/routing/0/tag. Znak ~ zapisz jako ~0, a / wewnątrz nazwy jako ~1.",
                () => FieldMatch.Missing(Pointer));
        else if (!string.IsNullOrWhiteSpace(JsonValue))
            errors.Add(new(FilterField.Pointer, "required", "Podaj ścieżkę JSON Pointer dla tej wartości, np. /extra/tag."));
        if (!string.IsNullOrWhiteSpace(JsonValue))
            Check(FilterField.JsonValue, "syntax", "Podaj poprawną wartość JSON: tekst w cudzysłowie, np. \"blue\", liczbę, true lub null. Puste pole oznacza brak właściwości.",
                () => { using var value = JsonDocument.Parse(JsonValue); });
        if (Change != null && reference == null)
            errors.Add(new(FilterField.Change, "reference-required", "Otwórz migawkę odniesienia Synthetic albo wybierz Wszystkie zmiany."));
        return errors.AsReadOnly();
    }

    internal bool UsesPublishedPattern => Domain.Contains('*');

    internal EndpointFilter ToQuery()
    {
        if (Ports.Trim().Length > 0 && Protocol is null)
            throw new FormatException("Wybierz TCP lub UDP dla zapytania o porty.");
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
                throw new FormatException("Wskaż ścieżkę JSON Pointer dla tej wartości.");
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
