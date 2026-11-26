<p align="center">
  <img src="assets/readme-logo.svg" alt="M365 Endpoint Inspector logo" width="128" />
</p>

# M365 Endpoint Inspector

A local browser of Microsoft 365 network requirements for Windows, written in C# using .NET 10 and WPF. The project is in development.

- Downloads the Worldwide list on request or opens a local JSON or CSV file.
- Filters and organizes requirements, shows entry details, and exports the view to CSV.
- Locally preserves the source, filters, selection, table layout, and expanded filter sections. Two sample synthetic snapshots let you test entry comparison without downloading data.

## Look up a domain or IP address

1. Open a local JSON or CSV file, or download the Microsoft 365 Worldwide list.
2. Enter a domain, IP address, or CIDR range and select **Apply**.
3. Inspect the matching entries for their service area, category, and TCP/UDP ports.
4. Export the filtered view to CSV to share the result.

You can review a saved list without downloading a new one. A match tells you what the source lists, not whether the endpoint is reachable from your network.

## Building and running

To try the Windows x64 release candidate, extract the entire ZIP to a writable folder and run `M365.EndpointInspector.Wpf.exe`. The package includes the .NET runtime; no SDK or separate .NET installation is needed. Workspace data is stored in the Windows user profile, not beside the executable.

**Reachability** runs DNS, direct TCP and HTTPS checks on request. An HTTP response does not prove service health. Successful HTTPS tunneling through a Windows system proxy has not been verified. **Open Microsoft 365 connectivity test** opens the official site in the default browser without passing targets or results.

On Windows, with the .NET SDK specified in `global.json`, run these commands from the repository directory:

```powershell
dotnet run --project src/M365.EndpointInspector.Wpf
```

Microsoft is the source of the Worldwide requirements. The service documentation describes data access and versioning, while the address list explains ranges, categories, and service areas. [Microsoft 365 IP Address and URL web service](https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide) · [Microsoft 365 URLs and IP address ranges](https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide)

For project structure, test commands, and compatibility rules, see the [developer guide](README-DEV.md).

## AI-assisted development

Parts of the application code and tests were developed with an AI assistant running in [OpenClaw](https://github.com/openclaw/openclaw), the agent harness. The maintainer directed the work and tested selected builds on Windows.

The model used to prepare this disclosure was GPT-6.2. The available development records do not reliably identify the model for every earlier coding session, so that identifier should not be read as attribution for the entire codebase. The AI tooling was used during development; it is not required to run the application.

The project is available under the [MIT license](LICENSE).
