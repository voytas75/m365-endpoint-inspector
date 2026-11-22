# M365 Endpoint Inspector

Lokalna przeglądarka wymagań sieciowych Microsoft 365 dla Windows, napisana w C# na .NET 10 i WPF. Wersja rozwojowa **v0.3.4-dev**.

- Pobiera wykaz Worldwide na żądanie albo otwiera lokalny plik JSON lub CSV.
- Filtruje i porządkuje wymagania, pokazuje szczegóły wpisów i eksportuje widok do CSV.
- Zachowuje lokalnie źródło, filtry, zaznaczenie i układ tabel. Dwie przykładowe migawki syntetyczne pozwalają sprawdzić porównywanie wpisów bez pobierania danych.

Aplikacja pokazuje wymagania Microsoft; nie mierzy jeszcze dostępności endpointów.

## Budowanie i uruchamianie

Na Windows, z .NET SDK wskazanym w `global.json`, wykonaj w katalogu repozytorium:

```powershell
dotnet build M365.EndpointInspector.slnx
dotnet run --project src/M365.EndpointInspector.Wpf
```

Źródłem wymagań dla planowanej edycji Worldwide jest Microsoft. Dokumentacja usługi opisuje dostęp do danych i wersjonowanie, a wykaz adresów objaśnia zakresy, kategorie i obszary usług. [Usługa internetowa Microsoft 365 IP Address and URL](https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide) · [Adresy URL i zakresy adresów IP Microsoft 365](https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide)

Projekt jest udostępniany na [licencji MIT](LICENSE).
