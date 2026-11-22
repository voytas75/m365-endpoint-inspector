namespace M365.EndpointInspector.Storage;

internal static class AtomicFile
{
    // Callers hold the writer lease. Both the candidate and the backup are on the same volume.
    internal static void Write(string path, Action<Stream> write, Action<string> validate,
        bool replace, string? backup, CancellationToken token)
    {
        string temporary = Path.Combine(Path.GetDirectoryName(path)!, Path.GetRandomFileName());
        var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        try
        {
            using (file)
            {
                write(file);
                file.Flush(flushToDisk: true);
            }
            validate(temporary);
            token.ThrowIfCancellationRequested();
            if (replace)
                File.Replace(temporary, path, backup);
            else
                File.Move(temporary, path);
        }
        catch
        {
            File.Delete(temporary);
            throw;
        }
    }
}
