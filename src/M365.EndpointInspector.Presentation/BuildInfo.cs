using System.Reflection;
using M365.EndpointInspector.Feed;

namespace M365.EndpointInspector.Presentation;

public sealed class BuildInfo
{
    public RefreshPolicy Policy { get; }
    public string Label { get; }
    public string WindowTitle => $"M365 Endpoint Inspector · {Label}";

    public BuildInfo(Assembly application, bool compiledForDebug = false)
    {
        string? info = application.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        var parts = info?.Split('+');
        if (parts is not [var version, var digest] || version.Length == 0 ||
            digest.Length != 64 || !digest.All(char.IsAsciiHexDigit))
            throw new InvalidDataException("Brak metadanych wersji lub sumy źródeł aplikacji.");

        var configuration = application.GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration;
        Policy = compiledForDebug && configuration == "Debug" ? RefreshPolicy.ManualDevelopment : RefreshPolicy.Release;
        string mode = Policy == RefreshPolicy.ManualDevelopment ? "DEV" : "RELEASE";
        Label = $"v{version} · {mode} · build {digest[..12]}";
    }
}
