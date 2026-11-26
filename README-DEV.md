# Developer guide

## Build and test

Use the exact .NET SDK specified in [global.json](global.json). `M365.EndpointInspector.slnx` builds the portable libraries and console test harnesses; it **does not include the WPF project**. Build the GUI explicitly on Windows:

```powershell
dotnet build src/M365.EndpointInspector.Wpf -c Debug
```

To prepare a portable release on Windows, publish to a new directory with `dotnet publish src/M365.EndpointInspector.Wpf -c Release -r win-x64 --self-contained true -o <publish-directory>`. Include `LICENSE`, the user instructions and the logo referenced by the README. ZIP the directory contents, then record the source revision and ZIP SHA-256. Verify the extracted executable on Windows x64 before distribution.

Run the offline console harnesses from the repository root (Windows or Linux/WSL):

```text
dotnet run --project tests/M365.EndpointInspector.Tests
dotnet run --project tests/M365.EndpointInspector.Feed.Tests
dotnet run --project tests/M365.EndpointInspector.Presentation.Tests
dotnet run --project tests/M365.EndpointInspector.Reachability.Tests
```

Use `--no-restore` if assets have already been restored. The Feed harness includes a filesystem test that requires permission to create symbolic links; a permission failure does not prove a cache defect. The portable harnesses do not replace testing the GUI on Windows.

The reachability harness uses synthetic data and controlled transports; see [the reachability plan](docs/reachability-plan.md) for Windows acceptance status.

## Regression Gate (GitHub Actions)

[Regression Gate](.github/workflows/regression.yml) runs on pushes and pull requests targeting `main`. It uses the SDK in `global.json` and reports the tested `GITHUB_SHA` (the merge-result SHA for pull requests).

- Ubuntu builds the portable solution in Release and runs Domain, Feed, Presentation and Reachability as separate console harnesses. They use synthetic inputs, simulated Worldwide transport and local loopback, not production Microsoft 365 endpoints or a user workspace. SDK setup and restore are separate from application traffic.
- Feed and Presentation require file and directory symlinks. If the runner cannot create them, those suites and link coverage are `BLOCKED`; the gate fails. A link capability probe alone is insufficient: both harnesses must report their link groups as passed. Checks that did not start are `NOT RUN`, never `PASS`.
- Windows restores and builds the WPF project separately in Release with warnings treated as errors. CI does not accept the GUI, qualify the final ZIP, publish assets, deploy, or change branch protection.

## Released baseline and regression notes

- The stable Windows x64 ZIP is published in [v1.0.0](https://github.com/voytas75/m365-endpoint-inspector/releases/tag/v1.0.0). Its source is tag `v1.0.0` (`c23bf50`), not the `main` tip used before release (`2147812`). Start release reproduction from the tag, not from a working tree with pending changes. The published ZIP SHA-256 is `5c8c487a9e01e3ffc92f1043545689a02ec12bcb986c46277f91796f3432f959`. This identifies the accepted archive; it does not imply byte-for-byte reproducible builds. Record Windows acceptance against the final ZIP hash; a rebuilt archive does not inherit candidate acceptance.
- Run Windows acceptance from a local Windows volume, as a standard user. A symlink fixture failure before the assertions does not establish a product defect or a pass. Qualify that fixture separately in an environment where file and directory links can be created; do not elevate the application. For JSON preview layout assertions, normalize CRLF/LF in the assertion, not in the application or the exported data.
- The workspace is per user (`%LOCALAPPDATA%`); the Worldwide cache is machine-scoped (`%ProgramData%`). Changing users alone does not provide a cold-cache test. The startup order is restored workspace, usable Worldwide cache, then one initial refresh attempt if there is no restored snapshot and saving is allowed. Unreadable workspace state must not be treated as a clean start. Format 1 accepts the legacy `Window.Monitor` field without using or rewriting it; other unknown fields are still rejected.
- Reachability is user initiated. Keep normal TLS validation and distinguish direct TCP from HTTPS using system proxy policy. Loopback and controlled transport tests do not qualify a successful Windows system proxy tunnel. The official Microsoft connectivity test is a browser link, not report integration.

## Project structure

- `Domain`: snapshot models, imports, and filters.
- `Feed`: Worldwide retrieval and cache.
- `Presentation`: workspace persistence, browser state, and CSV export.
- `Wpf`: Windows UI and interaction wiring.

## Compatibility and test data

- Keep maintained UI text, diagnostics, comments, and documentation in English. Do not translate persisted identifiers or CSV contract values as part of UI changes.
- Workspace uses Format 1; test compatibility with earlier files when adding persisted settings. Windows builds launched from different folders under the same user share the workspace. Back it up before GUI tests that replace the working session.
- The view export has a fixed 23-column schema, including legacy Polish headers and values. Grid layout changes do not reorder CSV columns. The view export and the Microsoft Worldwide CSV input are different formats.
- Use synthetic fixtures for tests. Do not commit real workspaces, caches, credentials, exports, or local build artifacts.
