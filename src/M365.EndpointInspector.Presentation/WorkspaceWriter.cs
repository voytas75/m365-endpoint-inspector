namespace M365.EndpointInspector.Presentation;

// Capture immutable data on the UI thread. Serialize disk writes, taking the newest pending
// snapshot after acquiring the gate so an older debounce cannot publish after a closing flush.
public sealed class WorkspaceWriter(WorkspaceStore store, TimeSpan debounce)
{
    private readonly SemaphoreSlim writes = new(1);
    private readonly object gate = new();
    private CancellationTokenSource? delay;
    private BrowserWorkspace? pending;

    public async Task<bool> ScheduleAsync(BrowserWorkspace workspace)
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
            return await WritePendingAsync().ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (wait.IsCancellationRequested) { return false; }
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

    private async Task<bool> WritePendingAsync()
    {
        await writes.WaitAsync().ConfigureAwait(false);
        try
        {
            BrowserWorkspace? captured;
            lock (gate) captured = pending;
            if (captured == null) return false;
            await Task.Run(() => store.Save(captured)).ConfigureAwait(false);
            lock (gate)
                if (ReferenceEquals(pending, captured)) pending = null;
            return true;
        }
        finally { writes.Release(); }
    }
}
