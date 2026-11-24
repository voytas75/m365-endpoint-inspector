using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using M365.EndpointInspector.Presentation;

internal static class NativeCases
{
    public static async Task RunAsync()
    {
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        var token = deadline.Token;
        using var key = RSA.Create(2048);
        var certificateRequest = new CertificateRequest("CN=127.0.0.1", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using var selfSigned = certificateRequest.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
        using var certificate = X509CertificateLoader.LoadPkcs12(selfSigned.Export(X509ContentType.Pkcs12), null);
        using var origin = Listen();
        int port = Port(origin);
        var directTarget = ReachabilityTarget.Parse($"https://127.0.0.1:{port}/probe")!;
        var tls = RejectCertificateAsync(origin, certificate, token);
        var directRows = await ProbeAsync(directTarget, new WebProxy(), token);
        await tls;
        Require(directRows[1] is { Status: ProbeStatus.Passed, Stage: ProbeStage.DirectTcp } && directTarget.Port == port, "direct TCP explicit port");
        Require(directRows[2].Observation is { Status: ProbeStatus.Failed, HttpError: HttpRequestError.SecureConnectionError, Route: ProbeRoute.Unknown }, "TLS rejection");
        Console.WriteLine("PASS native loopback direct TCP and untrusted TLS rejection");

        using var proxy = Listen();
        var proxyPolicy = new WebProxy($"http://127.0.0.1:{Port(proxy)}");
        var refusal = RejectTunnelAsync(proxy, token);
        var proxied = await ProbeAsync(directTarget, proxyPolicy, token);
        Require(await refusal == $"CONNECT 127.0.0.1:{port} HTTP/1.1", "CONNECT authority");
        Require(proxied[2].Observation is { Status: ProbeStatus.Failed, HttpError: HttpRequestError.ProxyTunnelError, Route: ProbeRoute.Proxy }, "proxy tunnel refusal");
        proxy.Stop();
        var refused = await ProbeAsync(directTarget, proxyPolicy, token);
        Require(refused[2].Observation is { Status: ProbeStatus.Failed, HttpError: HttpRequestError.ConnectionError, Route: ProbeRoute.Unknown }, "proxy socket refusal");
        Console.WriteLine("PASS native loopback proxy CONNECT rejection and proxy connection refusal");

        using var rootKey = RSA.Create(2048);
        var rootRequest = new CertificateRequest("CN=Loopback test root", rootKey, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        rootRequest.CertificateExtensions.Add(new X509BasicConstraintsExtension(true, false, 0, true));
        rootRequest.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.KeyCertSign, true));
        using var root = rootRequest.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-2), DateTimeOffset.UtcNow.AddDays(2));
        var san = new SubjectAlternativeNameBuilder();
        san.AddIpAddress(IPAddress.Loopback);
        certificateRequest.CertificateExtensions.Add(san.Build());
        using var leaf = certificateRequest.Create(root, DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1), RandomNumberGenerator.GetBytes(16));
        using var keyedLeaf = leaf.CopyWithPrivateKey(key);
        // Schannel needs a temporary key container; no certificate is installed in a store.
        using var server = X509CertificateLoader.LoadPkcs12(keyedLeaf.Export(X509ContentType.Pkcs12), null);
        var policy = new X509ChainPolicy { TrustMode = X509ChainTrustMode.CustomRootTrust, DisableCertificateDownloads = true, RevocationMode = X509RevocationMode.NoCheck };
        policy.CustomTrustStore.Add(root);
        foreach (var (code, outcome) in new[] { (204, ProbeStatus.Passed), (401, ProbeStatus.Restricted), (403, ProbeStatus.Restricted), (302, ProbeStatus.Responded), (500, ProbeStatus.Responded) })
        {
            using var trustedOrigin = Listen();
            var trustedTarget = ReachabilityTarget.Parse($"https://127.0.0.1:{Port(trustedOrigin)}/probe")!;
            var served = ServeTrustedAsync(trustedOrigin, server, code, token);
            var observed = await ProbeAsync(trustedTarget, code == 204 ? new FixturePolicy(_ => null) : new WebProxy(), token, policy);
            Require(observed[2].Observation.Code == ProbeCode.HttpResponse, "direct policy decision reaches HTTPS");
            Require(await served == "HEAD /probe HTTP/1.1", "TLS request preserves method and path");
            Require(observed[2].Observation == new ProbeObservation(outcome, ProbeCode.HttpResponse, ProbeRoute.Direct, code), $"trusted TLS HTTP {code}");
        }
        Console.WriteLine("PASS native isolated trusted TLS HEAD: HTTP 204/401/403/302/500 (including null policy result)");

        origin.Stop();
        const string secret = "fixture-secret";
        foreach (var (getPolicy, expected) in new (Func<IWebProxy>, ProbeCode)[]
        {
            (() => throw new InvalidOperationException(secret), ProbeCode.ProxyPolicyInitializationError),
            (() => new FixturePolicy(_ => null, _ => throw new InvalidOperationException(secret)), ProbeCode.ProxyPolicyBypassError),
            (() => new FixturePolicy(_ => throw new InvalidOperationException(secret)), ProbeCode.ProxyPolicyResolutionError),
            (() => new FixturePolicy(_ => throw new OperationCanceledException(secret)), ProbeCode.ProxyPolicyResolutionError),
            (() => new FixturePolicy(_ => new Uri("relative", UriKind.Relative)), ProbeCode.ProxyPolicyInvalidUri),
            (() => new FixturePolicy(_ => new Uri("ftp://127.0.0.1")), ProbeCode.ProxyPolicyInvalidUri),
            (() => new FixturePolicy(_ => new Uri($"http://user:{secret}@127.0.0.2")), ProbeCode.ProxyPolicyInvalidUri),
        })
        {
            var rows = new List<ProbeResult>();
            await new ReachabilityProbe(getPolicy, null).RunAsync(directTarget, rows.Add, token);
            Require(rows[2].Observation == new ProbeObservation(ProbeStatus.Failed, expected), $"policy diagnostic {expected}");
            var exported = ReachabilityReport.CopyText(rows);
            Require(!exported.Contains(secret) && !exported.Contains("127.0.0.2"), "policy diagnostics omit proxy address and credentials");
        }
        Console.WriteLine("PASS offline policy initialization, bypass, resolution and invalid URI diagnostics");

        foreach (var (code, outcome) in new[] { (204, ProbeStatus.Passed), (401, ProbeStatus.Restricted), (403, ProbeStatus.Restricted), (302, ProbeStatus.Responded), (500, ProbeStatus.Responded) })
        {
            var observation = ReachabilityProbe.ClassifyResponse(code, ProbeRoute.Direct);
            Require(observation == new ProbeObservation(outcome, ProbeCode.HttpResponse, ProbeRoute.Direct, code), $"HTTP {code} classification");
        }
        Require(directRows.All(row => row.Elapsed >= TimeSpan.Zero) && directRows[2].Elapsed > TimeSpan.Zero, "elapsed timing");
        Require(!ReachabilityReport.CopyText(proxied).Contains("/probe"), "clipboard excludes path");
        Require(ReachabilityReport.Failures(proxied).Targets.Single() == directTarget, "retry preserves port and path");
        Console.WriteLine("PASS structured HTTP classification, timing, clipboard privacy and retry identity");
    }

    private static async Task<List<ProbeResult>> ProbeAsync(ReachabilityTarget target, IWebProxy proxy, CancellationToken token, X509ChainPolicy? policy = null)
    {
        var rows = new List<ProbeResult>();
        await new ReachabilityProbe(proxy, policy).RunAsync(target, rows.Add, token);
        Require(rows.Count == 3, "published all stages");
        return rows;
    }

    private sealed class FixturePolicy(Func<Uri, Uri?> resolve, Func<Uri, bool>? bypass = null) : IWebProxy
    {
        public ICredentials? Credentials { get; set; }
        public Uri? GetProxy(Uri destination) => resolve(destination);
        public bool IsBypassed(Uri destination) => bypass?.Invoke(destination) ?? false;
    }

    private static TcpListener Listen()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return listener;
    }

    private static int Port(TcpListener listener) => ((IPEndPoint)listener.LocalEndpoint).Port;

    private static async Task RejectCertificateAsync(TcpListener listener, X509Certificate2 certificate, CancellationToken token)
    {
        using (await listener.AcceptTcpClientAsync(token)) { } // Standalone TCP probe.
        using var client = await listener.AcceptTcpClientAsync(token);
        using var tls = new SslStream(client.GetStream());
        try { await tls.AuthenticateAsServerAsync(new() { ServerCertificate = certificate }, token); }
        catch (Exception ex) when (ex is IOException or System.Security.Authentication.AuthenticationException) { }
    }

    private static async Task<string?> ServeTrustedAsync(TcpListener listener, X509Certificate2 cert, int code, CancellationToken token)
    {
        using (await listener.AcceptTcpClientAsync(token)) { }
        using var client = await listener.AcceptTcpClientAsync(token);
        using var tls = new SslStream(client.GetStream());
        await tls.AuthenticateAsServerAsync(new() { ServerCertificate = cert }, token);
        using var reader = new StreamReader(tls, Encoding.ASCII);
        var request = await reader.ReadLineAsync(token);
        while (await reader.ReadLineAsync(token) is { Length: > 0 }) { }
        await tls.WriteAsync(Encoding.ASCII.GetBytes($"HTTP/1.1 {code} Fixture\r\nLocation: /must-not-follow\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"), token);
        return request;
    }

    private static async Task<string?> RejectTunnelAsync(TcpListener listener, CancellationToken token)
    {
        using var client = await listener.AcceptTcpClientAsync(token);
        using var stream = client.GetStream();
        using var reader = new StreamReader(stream, Encoding.ASCII);
        var request = await reader.ReadLineAsync(token);
        while (await reader.ReadLineAsync(token) is { Length: > 0 }) { }
        await stream.WriteAsync("HTTP/1.1 407 Proxy Authentication Required\r\nProxy-Authenticate: Basic realm=\"fixture\"\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"u8.ToArray(), token);
        return request;
    }

    private static void Require(bool condition, string check)
    {
        if (!condition) throw new InvalidOperationException($"Reachability fixture failed: {check}");
    }
}
