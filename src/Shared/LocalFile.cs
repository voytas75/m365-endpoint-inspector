namespace M365.EndpointInspector.Storage;

internal static class LocalFile
{
    // Check each component, including dangling links. The root is operator-owned.
    internal static string CheckedPath(string path)
    {
        if (!Path.IsPathFullyQualified(path) ||
            path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Contains(".."))
            throw new ArgumentException("A full path without parent segments is required.");
        var file = new FileInfo(path);
        RejectLink(file);
        for (var parent = file.Directory; parent != null; parent = parent.Parent)
            RejectLink(parent);
        return file.FullName;
    }

    private static void RejectLink(FileSystemInfo entry)
    {
        if (entry.LinkTarget != null || entry.Exists && entry.Attributes.HasFlag(FileAttributes.ReparsePoint))
            throw new IOException($"Access through a symbolic link denied: {entry.Name}");
    }

    internal static byte[] ReadBytes(string path, int limit)
    {
        using var stream = new FileStream(CheckedPath(path), FileMode.Open, FileAccess.Read,
            FileShare.Read | FileShare.Delete);
        long length = stream.Length;
        if (length > limit) throw new InvalidDataException($"File exceeds the {limit}-byte limit.");
        var bytes = new byte[(int)length];
        stream.ReadExactly(bytes);
        if (stream.ReadByte() != -1) throw new IOException("File changed while being read.");
        return bytes;
    }
}
