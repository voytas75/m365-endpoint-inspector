using System.Text.Json;
using System.Text.Json.Serialization;

namespace M365.EndpointInspector.Presentation;

public sealed record GridColumn(string Id, int Order, double Width);
public sealed record GridSort(string Id, bool Descending);

/// <summary>Only bounded UI metadata. Unknown/malformed metadata never invalidates source data.</summary>
public sealed record GridLayout(GridColumn[] Columns, GridSort[] Sorts)
{
    public static IReadOnlyList<string> Ids { get; } = Array.AsReadOnly(new[]
        { "id", "area", "category", "required", "domains", "addresses", "tcp", "udp", "name", "expressRoute", "change" });
    public static IReadOnlyList<double> Widths { get; } = Array.AsReadOnly(new double[]
        { 54, 90, 83, 86, 180, 165, 118, 100, 160, 100, 92 });
    public static GridLayout Default => new(Ids.Select((id, i) => new GridColumn(id, i, Widths[i])).ToArray(), []);
    private static readonly JsonSerializerOptions Json = new()
    {
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true, MaxDepth = 8
    };

    public bool IsValid() => Columns != null && Sorts != null && Columns.Length == Ids.Count &&
        Columns.All(c => c != null && Ids.Contains(c.Id) && c.Order >= 0 && c.Order < Ids.Count &&
            double.IsFinite(c.Width) && c.Width >= 20 && c.Width <= 2000) &&
        Columns.Select(c => c.Id).Distinct().Count() == Ids.Count &&
        Columns.Select(c => c.Order).Distinct().Count() == Ids.Count &&
        Sorts.Length <= Ids.Count && Sorts.All(s => s != null && Ids.Contains(s.Id)) &&
        Sorts.Select(s => s.Id).Distinct().Count() == Sorts.Length;

    public static GridLayout Read(JsonElement? metadata)
    {
        if (metadata is not { ValueKind: JsonValueKind.Object } value || value.GetRawText().Length > 8192)
            return Default;
        try
        {
            var layout = value.Deserialize<GridLayout>(Json);
            return layout?.IsValid() == true ? layout : Default;
        }
        catch (JsonException) { return Default; }
    }

    internal JsonElement Pack() => JsonSerializer.SerializeToElement(IsValid() ? this : Default, Json);
}
