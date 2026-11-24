using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Wpf;

public partial class ReachabilityWindow : Window
{
    private readonly ObservableCollection<ProbeResult> rows = new();
    private readonly ReachabilityProbe probe = new();
    private readonly EndpointSnapshot? snapshot;
    private CancellationTokenSource? active;
    private int attempt;
    private readonly Dictionary<ReachabilityTarget, TargetProgress> progress = new();
    private readonly Dictionary<ReachabilityTarget, ProbeStage> running = new();
    private int total, completed, interrupted, waiting;
    private string phase = "Ready";

    private sealed class StageProgress(ReachabilityTarget target, ProbeStage stage) : INotifyPropertyChanged
    {
        public string Authority => target.Authority;
        public string StageLabel => ReachabilityReport.StageName(stage);
        public string State { get; private set; } = "Waiting";
        public bool IsRunning => State == "Running";
        public string Diagnostic { get; private set; } = "";
        public event PropertyChangedEventHandler? PropertyChanged;

        public void Update(string state, ProbeResult? result = null)
        {
            State = state;
            Diagnostic = result is null ? "" : $"{result.Timing} • {result.Detail}"
                + (stage == ProbeStage.HttpsSystemProxy ? $" • Route: {result.RouteLabel}" : "");
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(null));
        }
    }

    private sealed class TargetProgress(ReachabilityTarget target)
    {
        public StageProgress Dns { get; } = new(target, ProbeStage.Dns);
        public StageProgress Tcp { get; } = new(target, ProbeStage.DirectTcp);
        public StageProgress Https { get; } = new(target, ProbeStage.HttpsSystemProxy);
        public IEnumerable<StageProgress> Stages => [Dns, Tcp, Https];
        public StageProgress For(ProbeStage stage) => stage switch
        {
            ProbeStage.Dns => Dns,
            ProbeStage.DirectTcp => Tcp,
            _ => Https
        };
    }

    public ReachabilityWindow(string? host, IEnumerable<string> selected, EndpointSnapshot? snapshot, ServiceArea? area)
    {
        this.snapshot = snapshot;
        InitializeComponent();
        HostBox.Text = host ?? "";
        SelectedTargets.ItemsSource = selected.ToArray();
        SelectedTargets.SelectAll();
        SelectedTargets.SelectionChanged += (_, _) => UpdateControls();
        AreaBox.ItemsSource = Enum.GetValues<ServiceArea>();
        AreaBox.SelectedItem = area;
        AreaBox.SelectionChanged += (_, _) => UpdateControls();
        Results.ItemsSource = rows;
        CollectionViewSource.GetDefaultView(rows).SortDescriptions.Add(
            new(nameof(ProbeResult.OutcomePriority), System.ComponentModel.ListSortDirection.Ascending));
        HostBox.TextChanged += (_, _) => UpdateControls();
        UpdateControls();
        RenderProgress();
    }

    private void RenderProgress()
    {
        var counter = total == 0 ? "No checks started." : $"{phase} — {completed}/{total} targets completed"
            + (interrupted > 0 ? $" • {interrupted} interrupted" : "")
            + (active == null && waiting > 0 ? $" • {waiting} not started" : "");
        ProgressSummary.Text = counter;
        var text = string.Join("  •  ", running.Select(item =>
            $"{item.Key.Authority}: {ReachabilityReport.StageName(item.Value)}"));
        var activity = text.Length > 0 ? $"Running: {text}" : active != null ? "No stage active" : "";
        StageActivity.Text = activity;
        StageActivity.ToolTip = activity;
        ActivityPulse.IsIndeterminate = running.Count > 0;
        ActivityPulse.Visibility = ActivityPulse.IsIndeterminate ? Visibility.Visible : Visibility.Hidden;
    }

    private async void Run_Click(object sender, RoutedEventArgs e)
    {
        var batch = ReachabilityBatch.FromValues([HostBox.Text]);
        if (batch.Targets.Count == 0)
        {
            Status.Text = "Enter a concrete host, IP address or HTTPS URL without credentials, query or fragment.";
            return;
        }
        await RunAsync(batch);
    }

    private async void Selected_Click(object sender, RoutedEventArgs e) =>
        await RunAsync(ReachabilityBatch.FromValues(SelectedTargets.SelectedItems.Cast<string>()));

    private async void Area_Click(object sender, RoutedEventArgs e)
    {
        if (snapshot != null && AreaBox.SelectedItem is ServiceArea area)
            await RunAsync(ReachabilityBatch.FromArea(snapshot, area));
    }

    private async void Retry_Click(object sender, RoutedEventArgs e) =>
        await RunAsync(ReachabilityReport.Failures(rows), retry: true);

    private void Copy_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(ReachabilityReport.CopyText(rows));
            Status.Text = "Copied results with attempt, elapsed time and diagnostic codes. URL paths were omitted.";
        }
        catch (System.Runtime.InteropServices.ExternalException)
        {
            Status.Text = "Could not access the clipboard. Try again.";
        }
    }

    private void UpdateControls()
    {
        var idle = active == null;
        RunButton.IsEnabled = HostBox.IsEnabled = SelectedTargets.IsEnabled = idle;
        SelectedButton.IsEnabled = idle && SelectedTargets.SelectedItems.Count > 0;
        AreaBox.IsEnabled = idle && snapshot != null;
        AreaButton.IsEnabled = AreaBox.IsEnabled && AreaBox.SelectedItem is ServiceArea;
        CancelButton.IsEnabled = active is { IsCancellationRequested: false };
        if (idle)
        {
            var typed = ReachabilityBatch.FromValues(string.IsNullOrWhiteSpace(HostBox.Text) ? [] : [HostBox.Text]);
            var selected = ReachabilityBatch.FromValues(SelectedTargets.SelectedItems.Cast<string>());
            var area = snapshot != null && AreaBox.SelectedItem is ServiceArea selectedArea
                ? ReachabilityBatch.FromArea(snapshot, selectedArea) : ReachabilityBatch.FromValues([]);
            var retry = ReachabilityReport.Failures(rows);
            Preview.Text = $"Target: {typed.Summary} Selected: {selected.Summary} Area: {area.Summary} Retry: {retry.Targets.Count} target(s).";
            SkippedTargets.ItemsSource = new[] { ("Target", typed), ("Selected", selected), ("Area", area) }
                .SelectMany(scope => scope.Item2.SkippedTargets.Select(skip => $"{scope.Item1}: {skip.Value} — {skip.Reason}"))
                .ToArray();
            RetryButton.IsEnabled = retry.Targets.Count > 0;
        }
        else RetryButton.IsEnabled = false;
        CopyButton.IsEnabled = idle && rows.Count > 0;
        Results.CanUserSortColumns = idle;
    }

    private async Task RunAsync(ReachabilityBatch batch, bool retry = false)
    {
        if (active != null) return;
        var count = batch.Summary;
        if (batch.Targets.Count == 0)
        {
            Status.Text = $"No concrete targets. {count}";
            return;
        }
        if (!retry)
        {
            rows.Clear();
            attempt = 0;
        }
        var currentAttempt = ++attempt;
        using var cancellation = new CancellationTokenSource();
        active = cancellation;
        UpdateControls();
        progress.Clear();
        running.Clear();
        var targets = batch.Targets.Select(target => progress[target] = new TargetProgress(target)).ToArray();
        StageRows.ItemsSource = targets.SelectMany(target => target.Stages).ToArray();
        waiting = targets.Length;
        var resultView = CollectionViewSource.GetDefaultView(rows);
        resultView.SortDescriptions.Clear();
        total = batch.Targets.Count;
        completed = interrupted = 0;
        phase = "Checking";
        RenderProgress();
        Status.Text = $"Checking… {count}";
        try
        {
            await batch.RunAsync(probe, row => Dispatcher.Invoke(() =>
                {
                    rows.Add(row with { Attempt = currentAttempt });
                    progress[row.Target].For(row.Stage).Update(
                        row.Status == ProbeStatus.NotApplicable ? "Not applicable" : row.Status.ToString(), row);
                }), cancellation.Token,
                admitted: _ => Dispatcher.Invoke(() => waiting--),
                activity: (target, stage) => Dispatcher.Invoke(() =>
                {
                    if (stage is { } current)
                    {
                        running[target] = current;
                        progress[target].For(current).Update("Running");
                    }
                    else running.Remove(target);
                    RenderProgress();
                }),
                finished: (target, done) => Dispatcher.Invoke(() =>
                {
                    if (done) completed++;
                    else
                    {
                        interrupted++;
                        MarkUnstarted(progress[target]);
                    }
                    RenderProgress();
                }));
            phase = "Finished";
            Status.Text = $"Checks finished. {count}";
        }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
        {
            phase = "Cancelled";
            foreach (var target in targets) MarkUnstarted(target);
            Status.Text = $"Cancelled. Completed and interrupted stages are shown; unstarted stages are omitted. {count}";
        }
        finally
        {
            active = null;
            resultView.SortDescriptions.Add(new(nameof(ProbeResult.OutcomePriority), ListSortDirection.Ascending));
            RenderProgress();
            UpdateControls();
        }
    }

    private static void MarkUnstarted(TargetProgress target)
    {
        foreach (var stage in target.Stages)
            if (stage.State == "Waiting") stage.Update("Not started");
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        CancelButton.IsEnabled = false;
        Status.Text = "Cancelling checks…";
        phase = "Cancelling";
        RenderProgress();
        active?.Cancel();
    }

    protected override void OnClosed(EventArgs e)
    {
        active?.Cancel();
        base.OnClosed(e);
    }
}
