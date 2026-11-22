namespace M365.EndpointInspector.Presentation;

public sealed record WindowPlacement(double Left, double Top, double Width, double Height, bool Maximized)
{
    internal void Validate()
    {
        if (!double.IsFinite(Left) || !double.IsFinite(Top) ||
            !double.IsFinite(Width) || !double.IsFinite(Height) || Width <= 0 || Height <= 0)
            throw new InvalidDataException("Invalid window dimensions.");
    }

    public WindowPlacement Fit(double left, double top, double width, double height, double minWidth, double minHeight)
    {
        double w = Math.Clamp(Width, Math.Min(width, minWidth), width);
        double h = Math.Clamp(Height, Math.Min(height, minHeight), height);
        return new(Math.Clamp(Left, left, left + width - w), Math.Clamp(Top, top, top + height - h), w, h, Maximized);
    }
}

public sealed record BrowserWorkspace(BrowserDocument Document, BrowserFilter Applied, BrowserFilter Draft,
    BrowserDocument? Reference = null, int? SelectedId = null, WindowPlacement? Window = null,
    GridLayout? HitsLayout = null, GridLayout? CommonLayout = null,
    bool SourceFieldsExpanded = false, bool ComparisonExpanded = false)
{
    public static BrowserWorkspace Empty { get; } = new(BrowserDocument.Empty, new(), new());

    public BrowserWorkspace Open(BrowserDocument source)
    {
        if (source.Snapshot == null)
            throw new InvalidDataException($"{source.SourceLabel}. {source.Notice}", BrowserDocument.Diagnostic(source.Details));
        return new(source, new(), new(), Window: Window, HitsLayout: HitsLayout, CommonLayout: CommonLayout,
            SourceFieldsExpanded: SourceFieldsExpanded, ComparisonExpanded: ComparisonExpanded);
    }

    public WorkspaceView Present(bool applyDraft = false, bool keepEmptySelection = false)
    {
        var results = BrowserResults.Select(Document, applyDraft ? Draft : Applied, Reference);
        string? error = results.FilterError;
        var fieldErrors = results.FieldErrors;
        var next = this;
        if (applyDraft && error == null) next = this with { Applied = Draft };
        if (applyDraft && error != null) results = BrowserResults.Select(Document, Applied, Reference);
        if (!applyDraft && Draft != Applied)
        {
            var draftResults = BrowserResults.Select(Document, Draft, Reference);
            error = draftResults.FilterError;
            fieldErrors = draftResults.FieldErrors;
        }
        var rows = results.Hits.Concat(results.Common);
        int? id = SelectedId;
        if ((id.HasValue || !keepEmptySelection) && !rows.Any(row => row.Entry.Id == id)) id = rows.FirstOrDefault()?.Entry.Id;
        return new(next with { SelectedId = id }, results, error) { FieldErrors = fieldErrors };
    }
}

public sealed record WorkspaceView(BrowserWorkspace Workspace, BrowserResults Results, string? DraftError)
{
    public IReadOnlyList<FilterFieldError> FieldErrors { get; init; } = [];
}
