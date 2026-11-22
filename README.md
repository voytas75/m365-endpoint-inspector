# M365 Endpoint Inspector

A local browser of Microsoft 365 network requirements for Windows, written in C# using .NET 10 and WPF. Development version **v0.3.4-dev**.

- Downloads the Worldwide list on request or opens a local JSON or CSV file.
- Filters and organizes requirements, shows entry details, and exports the view to CSV.
- Locally preserves the source, filters, selection, table layout, and expanded filter sections. Two sample synthetic snapshots let you test entry comparison without downloading data.

The application displays Microsoft requirements; it does not yet measure endpoint availability.

## Building and running

On Windows, with the .NET SDK specified in `global.json`, run these commands from the repository directory:

```powershell
dotnet run --project src/M365.EndpointInspector.Wpf
```

Microsoft is the source of the Worldwide requirements. The service documentation describes data access and versioning, while the address list explains ranges, categories, and service areas. [Microsoft 365 IP Address and URL web service](https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide) · [Microsoft 365 URLs and IP address ranges](https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide)

For project structure, test commands, and compatibility rules, see the [developer guide](README-DEV.md).

The project is available under the [MIT license](LICENSE).
