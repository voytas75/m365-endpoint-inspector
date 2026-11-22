using M365.EndpointInspector.Storage;
using System.Text;
using static M365.EndpointInspector.Feed.Tests.SimulatedTransport;

namespace M365.EndpointInspector.Feed.Tests;

internal static class StorageCases
{
    internal static async Task Generations()
    {
        using var f = new Fixture();
        await f.Seed();
        f.Hour();
        await f.Seed(Second);
        string active = Path.Combine(f.DirectoryPath, "active.json");
        // A syntactically and structurally valid edit must still fail the persisted digest.
        File.WriteAllText(active, File.ReadAllText(active).Replace(Second, Third));
        Check.Equal(CacheGeneration.Previous, f.Feed.ReadCache().Generation);
        File.WriteAllText(active, "{");
        var fallback = f.Feed.ReadCache();
        Check.Equal(CacheGeneration.Previous, fallback.Generation);
        Check.Equal(First, fallback.Snapshot!.Version);
        Check.Equal(TimeSpan.FromHours(1), fallback.Age);
        Check.True(fallback.Error!.Contains("active"));
        f.Hour();
        await f.Seed(Third);
        Check.Equal(First, f.Feed.ReadCache(true).Snapshot!.Version);

        // A wrong origin cannot masquerade as an official download, even by copying the whole file.
        string real = Path.Combine(f.Root, "MicrosoftWorldwide");
        Directory.CreateDirectory(real);
        File.Copy(active, Path.Combine(real, "active.json"));
        Check.Equal(CacheGeneration.Unavailable, WorldwideFeed.CreateMicrosoft(f.Root).ReadCache().Generation);

        File.WriteAllText(active, "[]");
        File.WriteAllText(Path.Combine(f.DirectoryPath, "previous.json"), "null");
        var unavailable = f.Feed.ReadCache();
        Check.Equal(CacheGeneration.Unavailable, unavailable.Generation);
        Check.True(unavailable.Snapshot == null && unavailable.Age == null && unavailable.Error != null);
    }

    internal static async Task Atomicity()
    {
        using var f = new Fixture();
        await f.Seed();
        f.Hour();
        await f.Seed(Second);
        string active = Path.Combine(f.DirectoryPath, "active.json");
        string previous = Path.Combine(f.DirectoryPath, "previous.json");
        byte[] initial = File.ReadAllBytes(active);
        byte[] backup = File.ReadAllBytes(previous);
        string unrelated = Path.Combine(f.DirectoryPath, "unrelated.tmp");
        File.WriteAllText(unrelated, "preserve");
        foreach (string fault in new[] { "partial-write", "validation", "cancel-at-commit", "create-conflict" })
        {
            using var cancelled = new CancellationTokenSource();
            void Write()
            {
                AtomicFile.Write(active, stream =>
                {
                    stream.Write(Encoding.UTF8.GetBytes("{"));
                    if (fault == "partial-write") throw new IOException("simulated partial write");
                }, candidate =>
                {
                    Check.Equal("{", File.ReadAllText(candidate));
                    if (fault == "validation") throw new FormatException("simulated invalid candidate");
                    if (fault == "cancel-at-commit") cancelled.Cancel();
                }, replace: fault != "create-conflict", previous, cancelled.Token);
            }
            if (fault == "cancel-at-commit")
                await Check.Throws<OperationCanceledException>(() => Task.Run(Write));
            else if (fault == "validation")
                await Check.Throws<FormatException>(() => Task.Run(Write));
            else
                await Check.Throws<IOException>(() => Task.Run(Write));
            Check.True(initial.SequenceEqual(File.ReadAllBytes(active)));
            Check.True(backup.SequenceEqual(File.ReadAllBytes(previous)));
            Check.Equal("preserve", File.ReadAllText(unrelated));
            Check.Equal(Second, f.Feed.ReadCache().Snapshot!.Version);
        }

        f.Hour();
        f.Transport.Reply("version", Version(Third));
        f.Transport.Queue("endpoints", _ =>
        {
            // Simulate an external writer installing a different, valid generation mid-refresh.
            File.Copy(previous, active, overwrite: true);
            return Task.FromResult(Json(Rows));
        });
        var conflict = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Failed, conflict.Status);
        Check.Equal(First, conflict.Cache.Snapshot!.Version);
        Check.Equal(First, f.Feed.ReadCache(true).Snapshot!.Version);
    }

    internal static async Task PathsAndIdentity()
    {
        using var f = new Fixture();
        await Check.Throws<ArgumentException>(() => Task.FromResult(new WorldwideFeed("../relative", f.Client, f.Clock)));
        await Check.Throws<ArgumentException>(() =>
            Task.FromResult(new WorldwideFeed(Path.Combine(f.Root, "..", "escape"), f.Client, f.Clock)));
        await f.Seed();
        string state = Path.Combine(f.DirectoryPath, "machine-state.json");
        string original = File.ReadAllText(state);
        File.WriteAllText(state, "{}");
        Check.Equal(RefreshStatus.Failed, (await f.Feed.RefreshAsync()).Status);
        File.Delete(state);
        Check.Equal(RefreshStatus.Failed, (await f.Feed.RefreshAsync()).Status);
        Check.Equal(2, f.Transport.Ids.Count);
        File.WriteAllText(state, original);
        f.Hour();

        string preserved = Path.Combine(f.Root, "preserved.json");
        File.WriteAllText(preserved, "keep");
        string backup = Path.Combine(f.DirectoryPath, "previous.json");
        File.CreateSymbolicLink(backup, preserved);
        f.Transport.Pair(Second);
        var rejected = await f.Feed.RefreshAsync();
        Check.Equal(RefreshStatus.Failed, rejected.Status);
        Check.Equal(First, rejected.Cache.Snapshot!.Version);
        Check.Equal("keep", File.ReadAllText(preserved));

        string link = Path.Combine(f.Root, "linked-root");
        Directory.CreateSymbolicLink(link, f.Root);
        using var linked = new Fixture(link);
        Check.Equal(RefreshStatus.Failed, (await linked.Feed.RefreshAsync()).Status);
        Check.Equal(0, linked.Transport.Ids.Count);
    }
}
