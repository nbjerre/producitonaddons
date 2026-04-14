# WebAPI

ASP.NET Web API til SAP Business One integration for ProductionAddon og worksheet-funktionalitet.

## Formål

API'et samler forretningskald til blandt andet:

- produktionsordrer fra salgsordrer
- planlægning via PlanUnlimited
- stykliste/BOM-funktioner
- brugere, medarbejdere, konti og salgsmedarbejdere
- worksheets og worksheet-linjer

## Kør lokalt

1. Tilpas `appsettings.json` eller `appsettings.Development.json`.
2. Start projektet med Visual Studio eller `dotnet run`.
3. Brug /swagger til at teste apikald

## Base URL

Lokal udvikling bruger typisk en adresse som:

`http://localhost:5230`

Se den aktuelle port i launch settings eller konsollen ved opstart.

## OpenAPI

OpenAPI er aktiveret i development-miljø via [Program.cs](Program.cs).

- OpenAPI JSON: `/openapi/v1.json`
- Swagger UI: `/swagger`

Swagger UI er nu den primære endpoint-dokumentation under lokal udvikling. XML summaries og response metadata ligger direkte i controllers og modeller, så endpoint-beskrivelserne holdes tæt på koden.

## Generelt svarformat

De fleste endpoints returnerer et JSON-objekt med et eller flere af disse felter:

- `success`: om operationen lykkedes
- `data`: læste data fra SAP eller interne mappings
- `message`: bruger- eller driftsrelevant statusbesked
- `error`: fejlbesked ved validerings- eller integrationsfejl

Nogle endpoints returnerer ekstra statusfelter som `created`, `cancelled`, `released`, `finished`, `documentId`, `openUrl` eller detaljer fra eksterne processer.

## Test Matrix (Eksamen)

Denne aflevering fokuserer integrationstest på tre centrale flows:

| Flow | Endpoint | Testfil | Hvad testes |
|---|---|---|---|
| Create-all | `POST /api/sales-production/create-all` | `Tests/WebAPI.Tests/Integration/CreateAllIntegrationTests.cs` | Gyldig salgsordrelinje giver oprettet produktionsordre (`created=true`, `createdCount=1`). |
| Print | `POST /api/sales-production/print-production-orders` | `Tests/WebAPI.Tests/Integration/CrystalPrintIntegrationTests.cs` | 1) Ingen printable linjer giver `printableCount=0`. 2) Manglende Crystal-konfiguration returnerer tydelig Crystal-fejl. |
| Plan-all | `POST /api/plan/order/btnpu` | `Tests/WebAPI.Tests/Integration/PlanUnlimitedIntegrationTests.cs` | Korrekt mapping til PlanUnlimited-knappen `BtnPU`. |

### Kør tests

```powershell
dotnet test WebAPI/Tests/WebAPI.Tests/WebAPI.Tests.csproj -v minimal
```

Bemærk: testprojektet targeter `net10.0` og kræver en .NET 10 SDK.