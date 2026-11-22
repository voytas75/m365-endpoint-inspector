namespace M365.EndpointInspector.Storage;

internal static class LocalFile
{
    // Check each component, including dangling links. The root is operator-owned.
    internal static string CheckedPath(string path)
    {
        if (!Path.IsPathFullyQualified(path) ||
            path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Contains(".."))
            throw new ArgumentException("Wymagana jest pełna ścieżka bez segmentów nadrzędnych.");
        var file = new FileInfo(path);
        RejectLink(file);
        for (var parent = file.Directory; parent != null; parent = parent.Parent)
            RejectLink(parent);
        return file.FullName;
    }

    private static void RejectLink(FileSystemInfo entry)
    {
        if (entry.LinkTarget != null || entry.Exists && entry.Attributes.HasFlag(FileAttributes.ReparsePoint))
            throw new IOException($"Odmowa dostępu przez dowiązanie: {entry.Name}");
    }

    internal static byte[] ReadBytes(string path, int limit)
    {
        using var stream = new FileStream(CheckedPath(path), FileMode.Open, FileAccess.Read,
            FileShare.Read | FileShare.Delete);
        long length = stream.Length;
        if (length > limit) throw new InvalidDataException($"Plik przekracza limit {limit} bajtów.");
        var bytes = new byte[(int)length];
        stream.ReadExactly(bytes);
        if (stream.ReadByte() != -1) throw new IOException("Plik zmienił się podczas odczytu.");
        return bytes;
    }
}
