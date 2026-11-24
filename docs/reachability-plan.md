# Reachability diagnostics plan

Baseline: `b4429e9`, clean. Native Windows .NET SDK 10.0.404 confirmed. The unchanged Presentation harness failed at `AssertReadablePreview` during CSV import.

Scope: user-initiated diagnostics in a separate WPF dialog; no all-catalog option, external verification, authentication, certificate overrides, configuration changes or publication.

1. Implemented: typed or selected concrete host; DNS, direct TCP 443 and HTTPS using the system proxy policy. Focused harness first failed compilation for missing target type, then passed target validation and independent failure/deadline/HTTP outcome cases. Native WPF Debug build: 0 warnings, 0 errors. `git diff --check` passed.
2. Implemented: selected targets or one whole loaded service area. Fixture passed scope, wildcard/CIDR skips, normalized deduplication, live reports and cancellation with only four workers admitted. Native WPF Debug build: 0 warnings, 0 errors. UNC executable overwrite failed on repeat harness build; DLL execution with `-p:UseAppHost=false` passed. `git diff --check` passed.
3. Implemented: retry hosts with a failed or timed-out stage, preserve other hosts, and copy an allowlisted host/stage/outcome table. Focused harness first failed compilation for the missing report type, then passed all five reported groups including the native localhost DNS probe. No external network checks were run.

## Final verification

- Windows SDK 10.0.404: `dotnet build M365.EndpointInspector.slnx -p:UseAppHost=false` and `dotnet build src/M365.EndpointInspector.Wpf -c Debug --no-restore` passed with zero warnings/errors.
- `dotnet tests/M365.EndpointInspector.Reachability.Tests/bin/Debug/net10.0/M365.EndpointInspector.Reachability.Tests.dll`: all five reported groups passed. TCP/HTTPS outcomes in the harness use controlled transports; the native DNS test stops before either connection.
- Domain harness: 12/12 passed. Presentation repeated the clean-baseline CSV preview failure. Feed returned 0/17 with simulated retrieval, storage and timeout failures in this environment; a clean-baseline Feed run was not made.
- Pending: manual Windows dialog/clipboard interaction and production TCP/HTTPS system proxy behavior. No live endpoint validation is in scope.

## Independent review — changes required before acceptance

- Rebuilt solution and WPF on Windows SDK 10.0.404, both without warnings/errors. Reachability passed five groups; Domain passed 12/12. Independently reproduced Presentation `AssertReadablePreview` and Feed 0/17 on both the working tree and a fresh archive of HEAD under Windows/WSL UNC. These observed failures are not new in this change; the underlying cause remains unverified.
- Corrected and verified locally: explicit HTTPS URLs accept a path and optional port, but reject credentials, queries and fragments. Host/IP targets use `/` and port 443; a URL’s effective port applies to both direct TCP and HTTPS. These checks do not infer catalog port requirements or test UDP. Results and clipboard show host and port, not the URL path. Retries preserve the original path and port.
- HTTPS reports `Passed` for 2xx, `Restricted` for 401/403, and `Responded` for other HTTP statuses; none establishes service health. The route is separate from proxy policy: `Direct` or `Proxy` is reported on a response, `Proxy` on a CONNECT rejection, and `Unknown` on failures without route evidence. The dialog shows safe codes and elapsed stage time.
- Open: show per-target skip reasons and scope counts before starting, prioritize failed results, include timing and safe diagnostic codes in clipboard output, and preserve per-attempt identity when retries are interrupted. Verify selection behavior across both grids and manual clipboard interaction on Windows.
- Correction verification: Windows SDK 10.0.404 built `M365.EndpointInspector.slnx` and `src/M365.EndpointInspector.Wpf -c Debug` with `--no-restore -p:UseAppHost=false` (zero warnings/errors). The Reachability DLL runner passed eight groups and Domain passed 12/12. Real loopback sockets exercised direct TCP on an ephemeral port, untrusted TLS rejection under normal validation, an HTTP 407 CONNECT rejection, and refusal to connect to a stopped proxy. HTTP 204/401/403/302/500 were exercised at the production classification seam, not over successful TLS; no trust-store changes, certificate overrides, system proxy resolver acceptance, or production endpoint tests were performed. The presentation/feed baseline failures above were not rerun in this slice.
- Next: validate the batch and operator workflow, including manual Windows dialog and clipboard interaction.

## Local Windows trusted TLS checkpoint

- Windows SDK 10.0.404: portable solution build (`--no-restore -p:UseAppHost=false`) passed with 0 warnings/errors; Reachability DLL passed nine groups. Fresh loopback listeners served real HTTPS HEAD responses 204/401/403/302/500 using a fixture-only custom root and IP SAN. Schannel required a temporary private-key container; no certificate was installed, and no validation callback or OS proxy changes were used.
- Existing explicit proxy fixtures cover rejection and refusal, not successful proxy tunneling or real Windows system-proxy discovery. Manual UI and clipboard interaction remain unverified.

## Batch and reporting checkpoint

- Windows SDK 10.0.404: portable solution and WPF Debug builds (`--no-restore -p:UseAppHost=false`) passed with 0 warnings/errors. Reachability DLL passed all ten groups, Domain DLL passed 12/12, and `git diff --check` passed. Presentation and Feed baseline failures were not rerun. Local changes remain uncommitted.

- Preview shows executable, skipped and duplicate counts before starting; skipped values are listed with reasons by scope. The only scopes are typed target, selected rows, a single loaded service area and retry.
- Results and clipboard put failures/timeouts first by default. The report retains attempt numbers and interrupted stage rows without overwriting earlier evidence. Retry eligibility uses the latest non-cancelled outcome for each target/stage. Clipboard output includes elapsed time, route and structured diagnostic codes, not paths or raw errors.
- Automated fixtures cover scope counts, skip reasons, projection from both grid selection inputs, dialog-list subset planning, cancellation and retry/export behavior. The main window clears the other grid when an operator selects a row in one grid; the projection accepts both inputs, but that is not evidence of simultaneous interactive selection. Ctrl/Shift interaction, visual sorting and clipboard access still need manual Windows acceptance.

## Manual Windows acceptance — first run

- Operator reported a copied `localhost:443` table: DNS `Passed / Resolved` (3 ms), direct TCP `Failed / ConnectionRefused` (4072 ms), HTTPS `Failed / Unknown / ProxyPolicy` (0 ms). The dialog, single-target execution and clipboard output with the three stages were exercised. Batch interaction is still pending.
- Native Windows SDK 10.0.404 reproduced the same three outcomes in the production probe. For `localhost:443`, `HttpClient.DefaultProxy` returned `IsBypassed=false` and `GetProxy=null` without throwing. The old probe rejected `null` before opening HTTPS; the .NET HTTP handler interprets it as direct.
- Fixed: a successful null result, explicit bypass or destination return selects direct. Initialization, bypass and resolution exceptions and invalid proxy URIs produce distinct safe `ProxyPolicy*` codes and do not fall back to direct. No proxy addresses, credentials or exception text appear in diagnostics. A transport failure still reports route `Unknown` unless there is route evidence.
- Verified: the new trusted loopback null-policy fixture failed before the fix and returned `204 / Direct` afterward. The native production probe then reported DNS `Passed / Resolved`, TCP `ConnectionRefused`, and HTTPS `Failed / Unknown / HttpError / ConnectionRefused / ConnectionError` on `localhost:443`. Windows portable solution and WPF builds passed with zero warnings/errors; Reachability passed all 11 groups and Domain passed 12/12. No external targets or system proxy/security changes; Presentation and Feed baseline failures were not rerun.
- Remaining manual gate: reopen the rebuilt WPF dialog and rerun/copy `localhost:443` to confirm the new HTTPS diagnostic in the operator's session; then complete batch selection, cancellation/retry and clipboard acceptance. Successful system-proxy tunneling was not verified.

## Operator usage

Open **Reachability…**. Type a host or IP, or use the prefilled detail target. For multiple targets, select rows within either browser grid using Ctrl/Shift, then refine their targets in the dialog. Check the preview and skipped-target list before starting. **Run entire area** reads one selected service area from the loaded snapshot, irrespective of display filters; it never checks the entire catalog. Wildcards and CIDR ranges are skipped, not expanded. Checks use at most four concurrent hosts and an eight-second deadline per stage. Cancelling leaves available results visible. **Retry failures** rechecks all stages for failed/timed-out hosts only; **Copy results** includes hostnames or IPs, attempts, timing, route and diagnostic codes but omits URL paths and raw exception details, so review before sharing. An HTTPS response is transport evidence, not service health or authorization.


## Stacked layout and progress previews

- The operator chose **A — Stacked**. Split and Steps, their layout selector, and their navigation/filter wiring were removed. Scope previews and skips remain above the report; retry, cancellation, and copy keep their existing behavior.
- **Progress view** is switchable during the same run. **Compact bar** shows completed/total targets and cancellation state. **Live activity** shows only target stages explicitly entered by the probe; IP-literal DNS skips are not active stages. **Queue status** shows exact totals and a bounded sample: all active workers, up to four waiting targets and four recent finished targets, with overflow counts. Cancelled targets are distinguished as interrupted or not started and do not inflate the completed count. Progress covers the current attempt; report rows retain attempt history on retry.
- Operator follow-up on rebuilt `localhost:443`: copy confirmed; DNS resolved, direct TCP was refused, and HTTPS reported `ConnectionRefused` with route `Unknown`, not the previous `ProxyPolicy` outcome. This slice did not change route semantics.
- Windows SDK 10.0.404: solution and isolated WPF Debug builds passed with zero warnings/errors. The normal WPF output was locked by a running .NET process and was not overwritten. Reachability passed all 12 groups and Domain passed 12/12. Presentation and Feed baseline failures were not rerun. No external probes, commits or pushes.
- Native STA smoke exercised all three views during one 20-target run with a controlled transport: one completed, four active, 15 waiting, then four interrupted and 15 not started after cancellation. Results remained available. Evidence: `~/.hermes/artifacts/m365-endpoint-inspector/progress-preview/native-smoke.log`. The smoke did not show the window or write the clipboard.
- The previews above were superseded by the three selectable designs below.

## Selectable progress designs — historical comparison

- A — Stacked remains the only layout. **Progress view** switches between **1 — Stage rows** (fixed DNS/TCP/HTTPS rows from the start, active highlight and animated activity strip), **2 — Target matrix** (one stable row per target, three status columns and selectable diagnostics), and **3 — Progress bar** (prominent completed-target bar and running stage text next to Cancel, with the ordinary attempt results below).
- The views share the same run and actual stage callbacks. An interrupted stage says `Cancelled`; stages never entered say `Not started`. Literal-IP DNS skips are results, not active stages. Progress rows represent the current attempt; retry and clipboard output still use the accumulated attempt results. The ordinary grid appends during the run and restores failure-first sorting afterward.
- Windows .NET SDK 10.0.404 built the isolated app with zero warnings/errors. The retained production DLL opened its native main window. A separate visible Windows STA smoke switched all three views during one 20-target fixture batch, checked stable row identities and matrix selection, cancelled with one completed, four interrupted and 15 not started, then retried the failed host and verified both attempts in the clipboard projection. All 12 Reachability groups passed. No live external probes, config changes, commits or pushes.
- Retained build, smoke source and evidence: `~/.hermes/artifacts/m365-endpoint-inspector/progress-designs/`. From WSL: `"/mnt/c/Program Files/dotnet/dotnet.exe" "$(wslpath -w "$HOME/.hermes/artifacts/m365-endpoint-inspector/progress-designs/app/M365.EndpointInspector.Wpf.dll")"`. Open **Reachability…** and use **Progress view** to compare.
- Remaining manual gate: choose a design after checking visual fit and keyboard/mouse interaction during an approved loopback or fixture run. The smoke verified the clipboard projection, not a write to the Windows clipboard.


## Final progress design — Stage rows

- The operator chose **Stage rows** after visual comparison. **A — Stacked** remains the only layout. Removed the **Progress view** switcher, Target matrix, bar-only view, and their dedicated wiring. The active-stage highlight, animated activity strip, completed-target counter, attempt results, cancellation, retry, and copy remain.
- Windows .NET SDK 10.0.404 built the isolated WPF Debug app and Reachability harness offline with zero warnings/errors. All 12 Reachability groups passed. A visible native STA fixture against the retained build confirmed Stage rows are the only progress view, row identities stay stable during the run, the activity pulse and counter update, cancellation keeps partial results and distinguishes interrupted from not-started work, and retry retains both attempts in the clipboard projection.
- Retained app and evidence: `~/.hermes/artifacts/m365-endpoint-inspector/stage-rows-final/`. Launch from WSL: `"/mnt/c/Program Files/dotnet/dotnet.exe" "$(wslpath -w "$HOME/.hermes/artifacts/m365-endpoint-inspector/stage-rows-final/app/M365.EndpointInspector.Wpf.dll")"`.
- Operator supplied a multi-target clipboard report from manual acceptance: DNS, TCP and HTTPS results were present, with direct HTTPS 200 responses, other HTTP responses, TLS failures and deadline outcomes. This confirms a batch run and clipboard output. The supplied rows do not demonstrate cancellation, retry or a successful Windows proxy tunnel.
- Manual cancellation confirmed in the copied report: four active HTTPS stages ended as `Cancelled / UserCancelled`, and completed stage results were retained. `Not started` is a UI-only state, not a copied probe result; its visual confirmation and manual retry remain pending. The submitted attempt has no `Failed` or `TimedOut` rows to retry.
- Manual retry confirmed on `localhost:443`: attempts 1 and 2 remained in the copied report, each with DNS success and TCP/HTTPS connection refusal. This confirms rerunning a failed target and retaining history, not multi-target retry selection.
- Operator visually confirmed `Not started` after the cancelled batch. Together with the copied `Cancelled` outcomes and retained results, this closes the manual cancellation check.
- Deferred by the operator: successful HTTPS through the Windows system proxy is unverified because the current station has no proxy. Resume on a Windows station with an existing proxy configuration and an approved HTTPS target; acceptance requires an HTTP response and route `Proxy`. No configuration changes, commits or pushes were made by the implementation.
