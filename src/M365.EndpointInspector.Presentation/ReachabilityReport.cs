namespace M365.EndpointInspector.Presentation;

public static class ReachabilityReport
{
    public static ReachabilityBatch Failures(IEnumerable<ProbeResult> rows) => ReachabilityBatch.FromTargets(rows
        .Where(row => row.Status != ProbeStatus.Cancelled)
        .GroupBy(row => (row.Target, row.Stage))
        .Select(group => group.MaxBy(row => row.Attempt)!)
        .Where(row => row.Status is ProbeStatus.Failed or ProbeStatus.TimedOut)
        .Select(row => row.Target).Distinct());

    public static string StageName(ProbeStage stage) => stage switch
    {
        ProbeStage.Dns => "DNS",
        ProbeStage.DirectTcp => "Direct TCP",
        ProbeStage.HttpsSystemProxy => "HTTPS (system proxy policy)",
        _ => "Unknown"
    };

    // Export validated authorities and structured observations, never paths or exception prose.
    public static string CopyText(IEnumerable<ProbeResult> rows) => string.Join(Environment.NewLine,
        rows.OrderBy(row => row.OutcomePriority)
            .Select(row => $"{row.Target.Authority}\t{row.StageLabel}\t{row.Status}\t{row.Attempt}\t{row.RouteLabel}\t{row.Timing}\t{row.Detail}")
            .Prepend("Host and port\tStage\tOutcome\tAttempt\tHTTPS route\tElapsed\tDiagnostic codes"));
}
