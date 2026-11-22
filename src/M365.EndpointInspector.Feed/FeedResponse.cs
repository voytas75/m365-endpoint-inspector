using System.Net;
using System.Text.Json;

namespace M365.EndpointInspector.Feed;

internal static class FeedResponse
{
    internal static async Task<JsonDocument> ReadAsync(HttpResponseMessage response, int limit, CancellationToken token)
    {
        if (response.StatusCode != HttpStatusCode.OK)
            throw new HttpRequestException("Expected a complete HTTP 200 response.", null, response.StatusCode);
        if (response.Headers.CacheControl?.NoStore == true)
            throw new InvalidDataException("The service forbids storage of this response.");
        if (response.Content.Headers.ContentRange != null || response.Content.Headers.ContentLength > limit)
            throw new InvalidDataException("Partial or oversized feed response.");
        using var input = await response.Content.ReadAsStreamAsync(token).ConfigureAwait(false);
        using var body = new MemoryStream();
        byte[] buffer = new byte[16384];
        int read;
        while ((read = await input.ReadAsync(buffer, token).ConfigureAwait(false)) != 0)
        {
            if (body.Length + read > limit)
                throw new InvalidDataException("Feed response exceeds its size limit.");
            body.Write(buffer, 0, read);
        }
        if (response.Content.Headers.ContentLength is long length && length != body.Length)
            throw new InvalidDataException("Incomplete feed response.");
        var bytes = body.GetBuffer().AsMemory(0, (int)body.Length);
        int offset = bytes.Span.StartsWith(System.Text.Encoding.UTF8.Preamble) ? 3 : 0;
        return JsonDocument.Parse(bytes[offset..]);
    }

    internal static DateTimeOffset? RetryAfter(HttpResponseMessage response, DateTimeOffset now)
    {
        var retry = response.Headers.RetryAfter;
        return retry?.Delta is { } delta ? now + delta : retry?.Date;
    }

    internal static DateTimeOffset? ServerDeadline(HttpResponseMessage response, DateTimeOffset now, DateTimeOffset? retryAfter)
    {
        var freshUntil = response.Content.Headers.Expires;
        if (response.Headers.CacheControl?.MaxAge is { } maxAge)
        {
            var remaining = maxAge - (response.Headers.Age ?? TimeSpan.Zero);
            freshUntil = remaining > TimeSpan.Zero ? now + remaining : null;
        }
        var next = retryAfter;
        if (freshUntil is { } fresh && (next == null || fresh > next)) next = fresh;
        return next > now ? next : null;
    }
}
