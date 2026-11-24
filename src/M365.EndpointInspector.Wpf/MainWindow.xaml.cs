using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Wpf;

public partial class MainWindow : Window
{
    private sealed record Choice<T>(T? Value, string Label) where T : struct;

#if DEBUG
    private const bool CompiledForDebug = true;
#else
    private const bool CompiledForDebug = false;
#endif
    private static readonly BuildInfo Build = new(typeof(MainWindow).Assembly, CompiledForDebug);
    private static readonly TimeProvider Clock = TimeProvider.System;
    private readonly WorkspaceStore store = new(AppFolder(Environment.SpecialFolder.LocalApplicationData, "Workspace"), Clock);
    private readonly WorkspaceWriter writer;
    private BrowserWorkspace view = BrowserWorkspace.Empty;
    private BrowserWorkspace? lastQueued;
    private CancellationTokenSource? operation;
    private TaskCompletionSource? operationFinished;
    private bool canSave, rendering, closing, closeReady, wasMaximized;

    private static string AppFolder(Environment.SpecialFolder folder, string name) =>
        Path.Combine(Environment.GetFolderPath(folder), "M365 Endpoint Inspector", name);

    public MainWindow()
    {
        writer = new(store, TimeSpan.FromMilliseconds(500));
        InitializeComponent();
        DataContext = Build;
        HitsGrid.LayoutChanged += (_, _) => CaptureChanges();
        CommonGrid.LayoutChanged += (_, _) => CaptureChanges();
        SourceFields.Expanded += (_, _) => CaptureChanges();
        SourceFields.Collapsed += (_, _) => CaptureChanges();
        Comparison.Expanded += (_, _) => CaptureChanges();
        Comparison.Collapsed += (_, _) => CaptureChanges();
        var area = SystemParameters.WorkArea;
        MinWidth = Math.Min(MinWidth, area.Width);
        MinHeight = Math.Min(MinHeight, area.Height);
        Width = Math.Min(Width, area.Width);
        Height = Math.Min(Height, area.Height);
        SetChoices(Area, Enum.GetValues<ServiceArea>(), "All", areaChoice => areaChoice.ToString());
        SetChoices(CategoryBox, Enum.GetValues<Category>(), "All", category => category.ToString());
        SetChoices(RequiredBox, new[] { true, false }, "Any", value => value ? "Yes" : "No");
        SetChoices(Protocol, Enum.GetValues<Transport>(), "Any", value => value.ToString().ToUpperInvariant());
        SetChoices(Change, Enum.GetValues<ChangeKind>(), "All", value => BrowserRow.ChangeLabel(value));
        ShowView(view.Present());
        IsEnabled = false;
        Loaded += Restore;
        FilterPanel.AddHandler(TextBox.TextChangedEvent, new TextChangedEventHandler((_, _) => CaptureChanges()));
        FilterPanel.AddHandler(ComboBox.SelectionChangedEvent, new SelectionChangedEventHandler((_, _) => CaptureChanges()));
        LocationChanged += (_, _) => CaptureChanges();
        SizeChanged += (_, _) => CaptureChanges();
        StateChanged += (_, _) =>
        {
            if (WindowState != WindowState.Minimized) wasMaximized = WindowState == WindowState.Maximized;
            CaptureChanges();
        };
    }

    private async void Restore(object sender, RoutedEventArgs e)
    {
        Loaded -= Restore;
        var (loaded, display) = await Task.Run(() =>
        {
            var state = store.Load();
            return (state, state.Workspace.Present(keepEmptySelection: true));
        });
        if (closing) return;
        if (loaded.Workspace.Window is { } window)
        {
            // WorkArea uses WPF device-independent units for the primary monitor.
            var area = SystemParameters.WorkArea;
            var fit = window.Fit(area.Left, area.Top, area.Width, area.Height, MinWidth, MinHeight);
            Left = fit.Left;
            Top = fit.Top;
            Width = fit.Width;
            Height = fit.Height;
            if (fit.Maximized) WindowState = WindowState.Maximized;
        }
        ShowView(display);
        SaveNotice.Text = loaded.Warning ?? "";
        if (view.Document.Snapshot != null)
        {
            Activity.Text = "Workspace restored from local storage.";
            CancelAction.Visibility = Visibility.Collapsed;
            ActivityPanel.Visibility = Visibility.Visible;
        }
        IsEnabled = true;
        view = view with { Window = Placement() };
        lastQueued = view;
        canSave = loaded.CanSave;
    }

    private WindowPlacement? Placement()
    {
        var bounds = WindowState == WindowState.Normal ? new Rect(Left, Top, ActualWidth, ActualHeight) : RestoreBounds;
        if (bounds.IsEmpty || !double.IsFinite(bounds.Left) || !double.IsFinite(bounds.Top) ||
            bounds.Width <= 0 || bounds.Height <= 0) return view.Window;
        return new(bounds.Left, bounds.Top, bounds.Width, bounds.Height, wasMaximized);
    }

    private async void CaptureChanges()
    {
        if (!canSave || rendering || closing) return;
        view = view with { Draft = ReadFilter(), Window = Placement(),
            HitsLayout = HitsGrid.CaptureLayout(), CommonLayout = CommonGrid.CaptureLayout(),
            SourceFieldsExpanded = SourceFields.IsExpanded, ComparisonExpanded = Comparison.IsExpanded };
        if (view == lastQueued) return;
        var captured = view;
        lastQueued = captured;
        try
        {
            if (await writer.ScheduleAsync(captured) && lastQueued == captured)
                SaveNotice.Text = "";
        }
        catch (Exception ex) when (WorkspaceStore.IsDataError(ex))
        {
            SaveNotice.Text = $"Could not save the workspace. {ex.Message}";
        }
    }

    private static void SetChoices<T>(ComboBox box, IEnumerable<T> values, string all, Func<T, string> label)
        where T : struct
    {
        box.DisplayMemberPath = nameof(Choice<T>.Label);
        box.ItemsSource = values.Select(value => new Choice<T>(value, label(value))).Prepend(new(null, all)).ToArray();
    }

    private static T? Selected<T>(ComboBox box) where T : struct => (box.SelectedItem as Choice<T>)?.Value;

    private static void Select<T>(ComboBox box, T? value) where T : struct =>
        box.SelectedItem = box.Items.Cast<Choice<T>>().First(choice => Nullable.Equals(choice.Value, value));

    private BrowserFilter ReadFilter() => new()
    {
        Area = Selected<ServiceArea>(Area), Category = Selected<Category>(CategoryBox),
        Required = Selected<bool>(RequiredBox), Protocol = Selected<Transport>(Protocol),
        Ports = Ports.Text, Domain = DomainBox.Text, Address = Address.Text,
        DisplayName = DisplayNameBox.Text, Notes = Notes.Text, Pointer = Pointer.Text,
        JsonValue = JsonValue.Text, Change = Selected<ChangeKind>(Change)
    };

    private void ShowView(WorkspaceView display)
    {
        rendering = true;
        try
        {
            view = display.Workspace;
            var filter = view.Draft;
            Select(Area, filter.Area);
            Select(CategoryBox, filter.Category);
            Select(RequiredBox, filter.Required);
            Select(Protocol, filter.Protocol);
            Select(Change, filter.Change);
            Ports.Text = filter.Ports;
            DomainBox.Text = filter.Domain;
            Address.Text = filter.Address;
            DisplayNameBox.Text = filter.DisplayName;
            Notes.Text = filter.Notes;
            Pointer.Text = filter.Pointer;
            JsonValue.Text = filter.JsonValue;

            SourcePanel.DataContext = view.Document;
            Comparison.IsEnabled = view.Document.IsSynthetic;
            var baseline = view.Reference?.Snapshot;
            Change.IsEnabled = DetachReference.IsEnabled = baseline != null;
            ReferenceInfo.Text = baseline == null ? "No reference" : $"Synthetic · {baseline.Version}";
            ComparisonCaption.Text = baseline == null ? "" :
                $"Synthetic reference: {baseline.Version}   ·   Removed entries come from the reference.";
            HitsGrid.ChangeColumn.Visibility = CommonGrid.ChangeColumn.Visibility =
                baseline == null ? Visibility.Collapsed : Visibility.Visible;
            ResultsPanel.DataContext = display.Results;
            FilterProblem.Text = display.DraftError == null ? "" :
                "⚠ Filter not applied. The table retains the last valid result. Correct the indicated fields and select Apply.";
            SourceFields.IsExpanded = view.SourceFieldsExpanded;
            Comparison.IsExpanded = view.ComparisonExpanded;
            ShowFilterErrors(display.FieldErrors);
            // Validation may reveal a collapsed section; keep that actual state for later renders/saves.
            view = view with { SourceFieldsExpanded = SourceFields.IsExpanded, ComparisonExpanded = Comparison.IsExpanded };
            if (display.DraftError != null && display.FieldErrors.Count == 0)
                FilterProblem.Text += $" {display.DraftError}";
            HitsGrid.RestoreLayout(view.HitsLayout);
            CommonGrid.RestoreLayout(view.CommonLayout);
            HitsGrid.SetRows(display.Results.Hits);
            CommonGrid.SetRows(display.Results.Common);
            var hit = display.Results.Hits.FirstOrDefault(row => row.Entry.Id == view.SelectedId);
            var common = display.Results.Common.FirstOrDefault(row => row.Entry.Id == view.SelectedId);
            HitsGrid.SelectedItem = hit;
            CommonGrid.SelectedItem = common;
            if (hit != null) HitsGrid.ScrollIntoView(hit);
            if (common != null) CommonGrid.ScrollIntoView(common);
            ShowDetail(hit ?? common);
            FilterPanel.IsEnabled = view.Document.Snapshot != null;
        }
        finally { rendering = false; }
    }

    private void ShowFilterErrors(IReadOnlyList<FilterFieldError> errors)
    {
        void Show(FilterField field, Control input, TextBlock message)
        {
            var text = string.Join(" ", errors.Where(error => error.Field == field).Select(error => error.Message));
            input.Tag = text.Length == 0 ? null : "Invalid";
            message.Text = text.Length == 0 ? "" : $"⚠ {text}";
            System.Windows.Automation.AutomationProperties.SetHelpText(input, text);
        }
        Show(FilterField.Protocol, Protocol, ProtocolError);
        Show(FilterField.Ports, Ports, PortsError);
        Show(FilterField.Domain, DomainBox, DomainBoxError);
        Show(FilterField.Address, Address, AddressError);
        Show(FilterField.Pointer, Pointer, PointerError);
        Show(FilterField.JsonValue, JsonValue, JsonValueError);
        Show(FilterField.Change, Change, ChangeError);
        if (errors.Any(error => error.Field is FilterField.Pointer or FilterField.JsonValue)) SourceFields.IsExpanded = true;
        if (errors.Any(error => error.Field == FilterField.Change)) Comparison.IsExpanded = true;
    }

    private void SelectRow(object sender, SelectionChangedEventArgs e)
    {
        if (rendering || e.OriginalSource != sender || sender is not EndpointGrid grid || grid.UpdatingLayout) return;
        var other = grid == HitsGrid ? CommonGrid : HitsGrid;
        var row = grid.SelectedItem as BrowserRow;
        if (row != null) other.UnselectAll();
        else if (other.SelectedItem != null) return;
        ShowDetail(row);
        view = view with { SelectedId = row?.Entry.Id };
        CaptureChanges();
    }

    private void ShowDetail(BrowserRow? row)
    {
        DetailPanel.DataContext = row;
        Targets.ItemsSource = row?.SearchableTargets.ToArray();
        Targets.SelectedIndex = Targets.Items.Count > 0 ? 0 : -1;
        TargetBar.Visibility = Targets.Items.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        Memberships.IsEnabled = Targets.SelectedItem != null;
    }

    private void SelectTarget(object sender, SelectionChangedEventArgs e) =>
        Memberships.IsEnabled = Targets.SelectedItem != null;

    private async void Memberships_Click(object sender, RoutedEventArgs e)
    {
        if (Targets.SelectedItem is TargetChoice target) await ApplyAsync(target.AllMemberships());
    }

    private void Reachability_Click(object sender, RoutedEventArgs e)
    {
        if (operation != null || closing) return;
        var targets = ReachabilityBatch.SelectedValues(HitsGrid.SelectedItems.Cast<BrowserRow>(),
            CommonGrid.SelectedItems.Cast<BrowserRow>());
        new ReachabilityWindow((Targets.SelectedItem as TargetChoice)?.Value, targets,
            view.Document.Snapshot, Selected<ServiceArea>(Area)) { Owner = this }.ShowDialog();
    }

    private string? PickFile(string title, bool csv = false)
    {
        var picker = new OpenFileDialog { Title = title, Filter = csv ? "JSON/CSV files (*.json;*.csv)|*.json;*.csv|All files (*.*)|*.*" : "JSON files (*.json)|*.json" };
        return picker.ShowDialog(this) == true ? picker.FileName : null;
    }

    private void ExportCsv_Click(object sender, RoutedEventArgs e)
    {
        if (operation != null || closing) return;
        try
        {
            // Capture actual displayed order and last-applied rows before entering the modal dialog.
            var export = BrowserCsvExport.Capture(view, HitsGrid.Items.Cast<BrowserRow>(),
                CommonGrid.Items.Cast<BrowserRow>(),
                ResultsPanel.DataContext is BrowserResults { ShowCommon: true }, ReadFilter());
            var picker = new SaveFileDialog
            {
                Title = "Export the last applied view to CSV",
                Filter = "CSV files (*.csv)|*.csv", DefaultExt = ".csv", AddExtension = true,
                FileName = "wymagania.csv", OverwritePrompt = true
            };
            if (picker.ShowDialog(this) != true) return;
            var target = CsvExportFile.Inspect(picker.FileName);
            // Bind approval to content inspected after the system dialog, before any write.
            if (target.Exists && MessageBox.Show(this,
                "Replace the existing file with the captured view? A file change after this approval will abort saving.",
                "Confirm CSV save", MessageBoxButton.YesNo, MessageBoxImage.Warning,
                MessageBoxResult.No) != MessageBoxResult.Yes) return;
            target.Save(export, overwriteApproved: target.Exists);
            MessageBox.Show(this, export.Message, "CSV export", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException or ArgumentException)
        {
            MessageBox.Show(this, $"CSV was not exported. {ex.Message}", "CSV export",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void Download_Click(object sender, RoutedEventArgs e)
    {
        var current = view;
        await UpdateAsync("Checking Worldwide…", async token =>
        {
            var feed = WorldwideFeed.CreateMicrosoft(AppFolder(Environment.SpecialFolder.CommonApplicationData, "Feed"), Clock, Build.Policy);
            return current.Open(await BrowserDocument.DownloadAsync(feed, Clock, token));
        });
    }

    private async void OpenFile_Click(object sender, RoutedEventArgs e)
    {
        if (PickFile("Open JSON/CSV file", csv: true) is { } path)
            await OpenAsync(token => BrowserDocument.OpenFile(path, Clock, token));
    }

    private async void OpenDemo_Click(object sender, RoutedEventArgs e) =>
        await OpenAsync(token => BrowserDocument.OpenSynthetic(Path.Combine(AppContext.BaseDirectory, "Demo", "synthetic-2.json"), Clock, token));

    private async void OpenCache_Click(object sender, RoutedEventArgs e)
    {
        var picker = new OpenFolderDialog { Title = "Select a local MicrosoftWorldwide or SimulatedWorldwide folder" };
        if (picker.ShowDialog(this) != true) return;
        string folder = picker.FolderName;
        await OpenAsync(token => BrowserDocument.OpenCache(folder, Clock, token));
    }

    private Task OpenAsync(Func<CancellationToken, BrowserDocument> read)
    {
        var current = view;
        return UpdateAsync("Opening data…", token => Task.FromResult(current.Open(read(token))));
    }

    private async void OpenReference_Click(object sender, RoutedEventArgs e)
    {
        if (PickFile("Open a Synthetic reference snapshot") is not { } path) return;
        var current = view with { Draft = ReadFilter() };
        await UpdateAsync("Opening Synthetic reference…", token => Task.FromResult(current with
            { Reference = BrowserDocument.OpenSynthetic(path, Clock, token) }), applyDraft: true);
    }

    private async void Detach_Click(object sender, RoutedEventArgs e)
    {
        var next = view with { Reference = null, Applied = view.Applied with { Change = null }, Draft = ReadFilter() with { Change = null } };
        await UpdateAsync("Detaching reference…", _ => Task.FromResult(next), applyDraft: true);
    }

    private async void Apply_Click(object sender, RoutedEventArgs e) => await ApplyAsync(ReadFilter());
    private async void Reset_Click(object sender, RoutedEventArgs e) => await ApplyAsync(new BrowserFilter());

    private Task ApplyAsync(BrowserFilter filter)
    {
        var next = view with { Draft = filter };
        return UpdateAsync("Filtering entries…", _ => Task.FromResult(next), applyDraft: true);
    }

    private async Task UpdateAsync(string activity, Func<CancellationToken, Task<BrowserWorkspace>> prepare, bool applyDraft = false)
    {
        if (operation != null || closing) return;
        using var pending = new CancellationTokenSource();
        var token = pending.Token;
        operation = pending;
        var finished = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        operationFinished = finished;
        SourceActions.IsEnabled = FilterPanel.IsEnabled = ResultsPanel.IsEnabled = false;
        ActivityPanel.Visibility = CancelAction.Visibility = Visibility.Visible;
        CancelAction.IsEnabled = true;
        Activity.Text = activity;
        ActivityDetails.IsExpanded = false;
        ActivityDetails.Visibility = Visibility.Collapsed;
        try
        {
            var display = await Task.Run(async () =>
            {
                var candidate = await prepare(token).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                return candidate.Present(applyDraft);
            }, token);
            token.ThrowIfCancellationRequested();
            ShowView(display);
            CaptureChanges();
            ActivityPanel.Visibility = Visibility.Collapsed;
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            Activity.Text = "Operation cancelled. Previous view preserved.";
        }
        catch (Exception ex) when (WorkspaceStore.IsDataError(ex))
        {
            Activity.Text = $"{ex.Message} Previous view preserved.";
            if (ex is InvalidDataException { InnerException: IOException detail })
            {
                ActivityDetailText.Text = detail.Message;
                ActivityDetails.Visibility = Visibility.Visible;
            }
        }
        finally
        {
            operation = null;
            finished.SetResult();
            operationFinished = null;
            SourceActions.IsEnabled = ResultsPanel.IsEnabled = true;
            FilterPanel.IsEnabled = view.Document.Snapshot != null;
            CancelAction.Visibility = Visibility.Collapsed;
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        operation?.Cancel();
        CancelAction.IsEnabled = false;
        Activity.Text = "Cancelling operation…";
    }

    protected override async void OnClosing(CancelEventArgs e)
    {
        base.OnClosing(e);
        if (closeReady || e.Cancel) return;
        e.Cancel = true;
        if (closing) return;
        CaptureChanges();
        closing = true;
        IsEnabled = false;
        var finished = operationFinished?.Task;
        operation?.Cancel();
        // Leave the original Closing event before calling Close(), even for a clean workspace.
        await System.Windows.Threading.Dispatcher.Yield();
        if (finished != null) await finished;
        try { await writer.FlushAsync(); }
        catch (Exception ex) when (WorkspaceStore.IsDataError(ex))
        {
            SaveNotice.Text = $"Latest changes were not saved. {ex.Message}";
            if (MessageBox.Show(this, $"{SaveNotice.Text}\nClose the window anyway?", "Workspace save",
                MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes)
            {
                closing = false;
                IsEnabled = true;
                return;
            }
        }
        closeReady = true;
        Close();
    }
}
