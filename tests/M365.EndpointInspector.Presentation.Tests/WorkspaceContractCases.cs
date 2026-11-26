using M365.EndpointInspector.Presentation;
using M365.EndpointInspector.Feed.Tests;

namespace M365.EndpointInspector.Presentation.Tests;

// Opt-in discovery probes: a reproduced contract gap must return a nonzero exit code.
internal static class WorkspaceContractCases
{
    internal static async Task<int> Run()
    {
        string root = Path.Combine(Environment.GetEnvironmentVariable("M365_TEST_ARTIFACT_ROOT") ?? Path.GetTempPath(),
            "workspace-contract-" + Guid.NewGuid().ToString("N"));
        (string Name, Func<string, Task> Run)[] cases =
        [
            ("Close flush waits for A, publishes B, and cannot republish A", Order),
            ("Save result identifies the snapshot actually published", SaveResult),
            ("Failed close flush retains B for retry without another edit", Retry)
        ];
        int failed = 0;
        for (int i = 0; i < cases.Length; i++)
        {
            try
            {
                await cases[i].Run(Path.Combine(root, i.ToString()));
                Console.WriteLine($"PASS {cases[i].Name}");
            }
            catch (Exception error)
            {
                failed++;
                Console.Error.WriteLine($"FAIL {cases[i].Name}: {error}");
            }
        }
        Console.WriteLine($"Workspace contract: {cases.Length - failed} PASS, {failed} FAIL. Fixtures: {root}");
        return failed == 0 ? 0 : 1;
    }

    private static BrowserWorkspace State(string value) => BrowserWorkspace.Empty with { Draft = new() { Notes = value } };
    private static string Read(string root) => new WorkspaceStore(root, TimeProvider.System).Load().Workspace.Draft.Notes;

    private static async Task Order(string root)
    {
        var clock = new SaveBarrier();
        var store = new WorkspaceStore(root, clock);
        store.Load();
        var writer = new WorkspaceWriter(store, Timeout.InfiniteTimeSpan);
        var a = writer.ScheduleAsync(State("A"));
        var first = writer.FlushAsync();
        await clock.Entered.Task;
        var b = writer.ScheduleAsync(State("B"));
        Task closing = writer.FlushAsync();
        Task repeated = writer.FlushAsync();
        try
        {
            Check.True(!closing.IsCompleted && !repeated.IsCompleted);
        }
        finally { clock.Release.TrySetResult(); }
        await first;
        Check.True(await a is null && await b is null);
        await Task.WhenAll(closing, repeated);
        Check.Equal("B", Read(root));
        string backup = Path.Combine(root, "workspace.previous.json");
        using var json = System.Text.Json.JsonDocument.Parse(File.ReadAllBytes(backup));
        Check.Equal("A", json.RootElement.GetProperty("Workspace").GetProperty("Draft").GetProperty("Notes").GetString()!);
        await writer.FlushAsync();
        Check.Equal("B", Read(root));
    }

    private static async Task SaveResult(string root)
    {
        var clock = new SaveBarrier(pauseSecond: true);
        var store = new WorkspaceStore(root, clock);
        store.Load();
        var writer = new WorkspaceWriter(store, TimeSpan.Zero);
        var stateA = State("A");
        var stateB = State("B");
        var stateC = State("C");
        var a = writer.ScheduleAsync(stateA);
        await clock.Entered.Task;
        // A holds the writer semaphore. Both requests reach its wait queue synchronously.
        var b = writer.ScheduleAsync(stateB);
        var c = writer.ScheduleAsync(stateC);
        clock.Release.TrySetResult();
        await clock.SecondEntered.Task;
        try
        {
            var firstResult = await a;
            Check.True(ReferenceEquals(stateA, firstResult));
            Check.True(!ReferenceEquals(stateC, firstResult)); // Same identity check as CaptureChanges.
            Check.True(!b.IsCompleted && !c.IsCompleted);
            Check.Equal("A", Read(root));
        }
        finally { clock.SecondRelease.TrySetResult(); }
        var saved = await Task.WhenAll(a, b, c);
        Check.Equal("C", Read(root));
        Check.True(ReferenceEquals(stateA, saved[0]));
        Check.True(!ReferenceEquals(stateC, saved[0])); // A cannot acknowledge pending C.
        Check.True(!ReferenceEquals(stateB, saved[1]));
        Check.True(ReferenceEquals(stateC, saved[1])); // B's worker actually wrote C.
        Check.True(ReferenceEquals(stateC, saved[2])); // C observes its own commit.
    }

    private static async Task Retry(string root)
    {
        var store = new WorkspaceStore(root, TimeProvider.System);
        store.Load();
        store.Save(State("A"));
        var writer = new WorkspaceWriter(store, Timeout.InfiniteTimeSpan);
        var b = writer.ScheduleAsync(State("B"));
        string active = Path.Combine(root, "workspace.json");
        byte[] committed = File.ReadAllBytes(active);
        using (var lease = new FileStream(Path.Combine(root, "workspace.lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            await Check.Throws<IOException>(async () => { await writer.FlushAsync(); });
            Check.True(committed.SequenceEqual(File.ReadAllBytes(active)));
        }
        Check.True(await b is null);
        await writer.FlushAsync(); // no second ScheduleAsync, no edit
        Check.Equal("B", Read(root));
        await writer.FlushAsync();
    }

    // Decode consults the clock after the candidate is flushed but before it is published.
    // First save has no previous generation, so this hook marks that phase, not recovery.
    private sealed class SaveBarrier(bool pauseSecond = false) : TimeProvider
    {
        private int entered;
        internal TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource SecondEntered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource SecondRelease { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public override DateTimeOffset GetUtcNow()
        {
            int phase = Interlocked.Increment(ref entered);
            if (phase == 1)
            {
                Entered.SetResult();
                Release.Task.GetAwaiter().GetResult();
            }
            else if (phase == 2 && pauseSecond)
            {
                SecondEntered.SetResult();
                SecondRelease.Task.GetAwaiter().GetResult();
            }
            return DateTimeOffset.UnixEpoch;
        }
    }
}
