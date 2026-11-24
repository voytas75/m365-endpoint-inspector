using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using M365.EndpointInspector.Domain;

namespace M365.EndpointInspector.Presentation;

public sealed record ReachabilityTarget
{
    public string Host { get; }
    public bool IsAddress { get; }

    public int Port { get; }
    public Uri HttpsUri { get; }
    public string Authority => $"{(Host.Contains(':') ? $"[{Host}]" : Host)}:{Port}";

    private ReachabilityTarget(string host, bool isAddress, int port = 443, string path = "/")
    {
        (Host, IsAddress, Port) = (host, isAddress, port);
        HttpsUri = new UriBuilder("https", host, port, path).Uri;
    }

    public static ReachabilityTarget? Parse(string input)
    {
        var value = input.Trim();
        if (!value.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return ParseHost(value);
        if (value.IndexOfAny(['?', '#', '\\']) >= 0 || value.Any(c => char.IsControl(c) || char.IsWhiteSpace(c))) return null;
        var authority = value[8..].Split('/', 2)[0];
        if (authority.Contains('@') || !Uri.TryCreate(value, UriKind.Absolute, out var uri) || uri.Port < 1) return null;
        // Validate the literal authority before Uri can normalize shorthand IPv4 or an empty port.
        var end = authority.StartsWith('[') ? authority.IndexOf(']') + 1 : authority.Split(':')[0].Length;
        if (end == 0) return null;
        var port = authority[end..];
        if (port.Length != 0 && (port[0] != ':' || port.Length == 1 || port[1..].Any(c => !char.IsAsciiDigit(c)))) return null;
        var literal = authority[..end];
        var target = ParseHost(literal.StartsWith('[') ? literal[1..^1] : literal);
        if (target is null || (literal.StartsWith('[') && !target.Host.Contains(':'))) return null;
        return new(target.Host, target.IsAddress, uri.Port, uri.AbsolutePath);
    }

    private static ReachabilityTarget? ParseHost(string input)
    {
        var host = input;
        if (host.Length == 0 || host.IndexOfAny(['*', '/', '%']) >= 0) return null;
        if (IPAddress.TryParse(host, out var ip))
        {
            try { IpPrefix.Parse(host); }
            catch (FormatException) { return null; }
            return new(ip.ToString(), true);
        }
        if (host.EndsWith('.')) host = host[..^1];
        if (host.Length > 253 || !char.IsAsciiLetterOrDigit(host.FirstOrDefault()) ||
            host.Split('.').Any(label => label.Length is < 1 or > 63 || label[0] == '-' || label[^1] == '-' ||
                label.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-')) ||
            Uri.CheckHostName(host) != UriHostNameType.Dns ||
            Uri.CheckHostName(input.Trim()) != UriHostNameType.Dns) return null;
        return new(host.ToLowerInvariant(), false);
    }
}

public enum ProbeStage { Dns, DirectTcp, HttpsSystemProxy }
public enum ProbeStatus { Passed, Restricted, Responded, NotApplicable, Failed, TimedOut, Cancelled }
public enum ProbeRoute { Unknown, Direct, Proxy }
public enum ProbeCode { None, IpLiteral, Resolved, Connected, HttpResponse, Deadline, UserCancelled, SocketError, HttpError, IoError, ProxyPolicyInitializationError, ProxyPolicyBypassError, ProxyPolicyResolutionError, ProxyPolicyInvalidUri }

public sealed record ProbeObservation(ProbeStatus Status, ProbeCode Code, ProbeRoute Route = ProbeRoute.Unknown,
    int? HttpStatus = null, SocketError? SocketError = null, HttpRequestError? HttpError = null);

public sealed record ProbeResult(ReachabilityTarget Target, ProbeStage Stage, ProbeObservation Observation, TimeSpan Elapsed)
{
    public int Attempt { get; init; } = 1;
    public int OutcomePriority => Status switch
    {
        ProbeStatus.Failed or ProbeStatus.TimedOut => 0,
        ProbeStatus.Cancelled or ProbeStatus.Restricted or ProbeStatus.Responded => 1,
        _ => 2
    };
    public ProbeStatus Status => Observation.Status;
    public string StageLabel => Stage == ProbeStage.DirectTcp ? $"Direct TCP {Target.Port}" : ReachabilityReport.StageName(Stage);
    public string RouteLabel => Stage == ProbeStage.HttpsSystemProxy ? Observation.Route.ToString() : "—";
    public string Detail => $"{Observation.Code}{(Observation.HttpStatus is int code ? $" HTTP {code}" : "")}"
        + $"{(Observation.SocketError is { } socket ? $" / {socket}" : "")}"
        + $"{(Observation.HttpError is { } http ? $" / {http}" : "")}";
    public string Timing => $"{Elapsed.TotalMilliseconds:F0} ms";
}

public sealed class ReachabilityProbe
{
    private readonly Func<ReachabilityTarget, ProbeStage, CancellationToken, Task<ProbeObservation>> execute;
    private readonly TimeSpan timeout;

    public ReachabilityProbe() : this(() => HttpClient.DefaultProxy, null) { }

    public ReachabilityProbe(IWebProxy proxyPolicy) : this(proxyPolicy, null) { }

    internal ReachabilityProbe(IWebProxy proxyPolicy, System.Security.Cryptography.X509Certificates.X509ChainPolicy? chainPolicy)
        : this(() => proxyPolicy, chainPolicy) { }

    internal ReachabilityProbe(Func<IWebProxy> getProxyPolicy, System.Security.Cryptography.X509Certificates.X509ChainPolicy? chainPolicy)
        : this((target, stage, token) => ExecuteAsync(target, stage, getProxyPolicy, chainPolicy, token), TimeSpan.FromSeconds(8)) { }

    public ReachabilityProbe(Func<ReachabilityTarget, ProbeStage, CancellationToken, Task<ProbeObservation>> execute, TimeSpan timeout)
        => (this.execute, this.timeout) = (execute, timeout);

    public async Task RunAsync(ReachabilityTarget target, Action<ProbeResult> publish, CancellationToken cancellation,
        Action<ProbeStage?>? activity = null)
    {
        foreach (var stage in Enum.GetValues<ProbeStage>())
        {
            if (cancellation.IsCancellationRequested) break;
            if (stage == ProbeStage.Dns && target.IsAddress)
            {
                publish(new(target, stage, new(ProbeStatus.NotApplicable, ProbeCode.IpLiteral), TimeSpan.Zero));
                continue;
            }
            activity?.Invoke(stage);
            using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            deadline.CancelAfter(timeout);
            var start = Stopwatch.GetTimestamp();
            ProbeObservation observation;
            try
            {
                observation = await execute(target, stage, deadline.Token).ConfigureAwait(false);
                deadline.Token.ThrowIfCancellationRequested();
            }
            catch (OperationCanceledException) when (deadline.IsCancellationRequested)
            {
                observation = cancellation.IsCancellationRequested
                    ? new(ProbeStatus.Cancelled, ProbeCode.UserCancelled)
                    : new(ProbeStatus.TimedOut, ProbeCode.Deadline);
            }
            catch (Exception ex) when (ex is SocketException or HttpRequestException or IOException)
            {
                observation = Failure(ex);
            }
            finally
            {
                activity?.Invoke(null);
            }
            publish(new(target, stage, observation, Stopwatch.GetElapsedTime(start)));
        }
    }

    private static ProbeObservation Failure(Exception ex, ProbeRoute route = ProbeRoute.Unknown) => ex switch
    {
        SocketException socket => new(ProbeStatus.Failed, ProbeCode.SocketError, route, SocketError: socket.SocketErrorCode),
        HttpRequestException http => new(ProbeStatus.Failed, ProbeCode.HttpError, route,
            http.StatusCode is { } code ? (int)code : null, (http.InnerException as SocketException)?.SocketErrorCode,
            http.HttpRequestError),
        _ => new(ProbeStatus.Failed, ProbeCode.IoError, route)
    };

    private static async Task<ProbeObservation> ExecuteAsync(ReachabilityTarget target, ProbeStage stage,
        Func<IWebProxy> getProxyPolicy, System.Security.Cryptography.X509Certificates.X509ChainPolicy? chainPolicy, CancellationToken token)
    {
        if (stage == ProbeStage.Dns)
        {
            var addresses = await Dns.GetHostAddressesAsync(target.Host, token).ConfigureAwait(false);
            if (addresses.Length == 0) throw new SocketException((int)SocketError.NoData);
            return new(ProbeStatus.Passed, ProbeCode.Resolved);
        }
        if (stage == ProbeStage.DirectTcp)
        {
            using var socket = new TcpClient();
            await socket.ConnectAsync(target.Host, target.Port, token).ConfigureAwait(false);
            return new(ProbeStatus.Passed, ProbeCode.Connected);
        }

        // Snapshot the policy decision for this request. Do not infer HTTPS routing from the TCP stage.
        var destination = target.HttpsUri;
        var policyError = ProbeCode.ProxyPolicyInitializationError;
        Uri? proxy;
        try
        {
            var proxyPolicy = getProxyPolicy();
            policyError = ProbeCode.ProxyPolicyBypassError;
            var bypassed = proxyPolicy.IsBypassed(destination);
            policyError = ProbeCode.ProxyPolicyResolutionError;
            proxy = bypassed ? null : proxyPolicy.GetProxy(destination);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException &&
            (ex is not OperationCanceledException || !token.IsCancellationRequested))
        {
            return new(ProbeStatus.Failed, policyError);
        }
        token.ThrowIfCancellationRequested();

        // SocketsHttpHandler treats a null proxy result as direct, including Windows automatic policy.
        if (proxy is not null && proxy != destination && (!proxy.IsAbsoluteUri || proxy.UserInfo.Length != 0 ||
            proxy.Scheme is not ("http" or "https" or "socks4" or "socks4a" or "socks5")))
            return new(ProbeStatus.Failed, ProbeCode.ProxyPolicyInvalidUri);

        var route = proxy is null || proxy == destination ? ProbeRoute.Direct : ProbeRoute.Proxy;
        using var http = new HttpClient(new SocketsHttpHandler
        {
            UseProxy = route == ProbeRoute.Proxy, Proxy = route == ProbeRoute.Proxy ? new WebProxy(proxy) : null,
            SslOptions = new() { CertificateChainPolicy = chainPolicy },
            AllowAutoRedirect = false, UseCookies = false, Credentials = null, PreAuthenticate = false
        }) { Timeout = Timeout.InfiniteTimeSpan };
        using var request = new HttpRequestMessage(HttpMethod.Head, destination);
        try
        {
            using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);
            return ClassifyResponse((int)response.StatusCode, route);
        }
        catch (HttpRequestException ex)
        {
            // A failed connection does not prove a route. A CONNECT rejection does prove proxy contact.
            return Failure(ex, ex.HttpRequestError == HttpRequestError.ProxyTunnelError && route == ProbeRoute.Proxy
                ? ProbeRoute.Proxy : ProbeRoute.Unknown);
        }
    }
    public static ProbeObservation ClassifyResponse(int status, ProbeRoute route) => new(status switch
    {
        401 or 403 => ProbeStatus.Restricted,
        >= 200 and < 300 => ProbeStatus.Passed,
        _ => ProbeStatus.Responded
    }, ProbeCode.HttpResponse, route, status);
}
