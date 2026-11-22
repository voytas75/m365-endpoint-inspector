using System.Security.Cryptography;

namespace M365.EndpointInspector.Presentation;

// An explicit approval binds existing content. Publication never overwrites a newly appeared file.
public sealed class CsvExportFile
{
    private readonly string path;
    private readonly byte[]? expected;
    public bool Exists => expected != null;
    private CsvExportFile(string path, byte[]? expected) => (this.path, this.expected) = (path, expected);

    public static CsvExportFile Inspect(string path)
    {
        path = Path.GetFullPath(path);
        RejectLink(path);
        return new(path, File.Exists(path) ? Digest(path) : null);
    }

    public void Save(BrowserCsvExport export, bool overwriteApproved, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        if (Exists && !overwriteApproved) throw new InvalidOperationException("Nadpisanie nie zostało zatwierdzone.");
        string temporary = Path.Combine(Path.GetDirectoryName(path)!, Path.GetRandomFileName());
        string backup = temporary + ".original";
        bool displaced = false;
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(export.CopyBytes());
                stream.Flush(true);
            }
            token.ThrowIfCancellationRequested();
            RejectLink(path);
            if (expected != null)
            {
                // Move first, then validate the actual displaced object, closing the check/overwrite race.
                File.Move(path, backup, overwrite: false);
                displaced = true;
                RejectLink(backup);
                if (!CryptographicOperations.FixedTimeEquals(expected, Digest(backup)))
                    throw new IOException("Plik zmienił się po zatwierdzeniu. Nie zapisano eksportu.");
            }
            // Same-directory atomic publication. If another writer creates the target, refuse.
            File.Move(temporary, path, overwrite: false);
        }
        catch
        {
            if (displaced)
            {
                try { File.Move(backup, path, overwrite: false); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    throw new IOException($"Nie zapisano eksportu. Inny zapis zajął cel; zachowano oryginał w {backup}.");
                }
            }
            throw;
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
        // Keep the previous file as recovery evidence if cleanup is denied; publication already succeeded.
        if (displaced)
        {
            try { File.Delete(backup); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static byte[] Digest(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return SHA256.HashData(stream);
    }
    private static void RejectLink(string path)
    {
        if (File.Exists(path) && (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("Eksport nie zastępuje dowiązań plików.");
    }
}
