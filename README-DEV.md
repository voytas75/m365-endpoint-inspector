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
