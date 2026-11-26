namespace M365.EndpointInspector.Presentation;

// Capture immutable data on the UI thread. Serialize disk writes, taking the newest pending
// snapshot after acquiring the gate so an older debounce cannot publish after a closing flush.
public sealed class WorkspaceWriter(WorkspaceStore store, TimeSpan debounce)
{
    private readonly SemaphoreSlim writes = new(1);
    private readonly object gate = new();
    private CancellationTokenSource? delay;
    private BrowserWorkspace? pending;
    // Only accessed while holding writes; reference identity distinguishes requests.
    private BrowserWorkspace? committed;

    // Returns the snapshot actually written, or the requested snapshot if another
    // operation already committed it. A cancelled or superseded wait returns null.
    public async Task<BrowserWorkspace?> ScheduleAsync(BrowserWorkspace workspace)
    {
        using var wait = new CancellationTokenSource();
        lock (gate)
        {
            pending = workspace;
            delay?.Cancel();
            delay = wait;
        }
        try
        {
            await Task.Delay(debounce, wait.Token).ConfigureAwait(false);
            return await WritePendingAsync(workspace).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (wait.IsCancellationRequested) { return null; }
        finally
        {
            lock (gate)
                if (ReferenceEquals(delay, wait)) delay = null;
        }
    }

    public Task FlushAsync()
    {
        lock (gate) delay?.Cancel();
        return WritePendingAsync();
    }

    private async Task<BrowserWorkspace?> WritePendingAsync(BrowserWorkspace? requested = null)
    {
        await writes.WaitAsync().ConfigureAwait(false);
        try
        {
            BrowserWorkspace? captured;
            lock (gate) captured = pending;
            if (captured == null) return ReferenceEquals(requested, committed) ? committed : null;
            await Task.Run(() => store.Save(captured)).ConfigureAwait(false);
            committed = captured;
            lock (gate)
                if (ReferenceEquals(pending, captured)) pending = null;
            return captured;
        }
        finally { writes.Release(); }
    }
}
