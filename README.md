# Utviklingsradar kliniske og pasientrettede systemer

En enkel React/Vite-app som leser Excel-eksporten for "Nye tjenester" og viser faseboard per `Assignment Group`.

## Bruk

1. Legg nyeste `.xlsx` i `input/`.
2. Kjør `npm install` første gang.
3. Kjør `npm run dev` for lokal visning.
4. Åpne adressen Vite viser i terminalen.

Hvis `input/` er tom, brukes nyeste `.xlsx` i prosjektroten som fallback.

## Dataimport

Kjør import manuelt med:

```bash
npm run import:data
```

Importen genererer `src/data/requests.json` fra Excel-filen.

## Eksport

I appen kan du eksportere:

- valgt board som PNG
- alle boards som PNG, én fil per `Assignment Group`
- PowerPoint med én slide per `Assignment Group`

Noen nettlesere ber om godkjenning ved nedlasting av mange PNG-filer samtidig.

## Build

```bash
npm run build
```

## Vercel

Prosjektet er klart for statisk deploy til Vercel.

- Build command: `npm run build`
- Output directory: `dist`
- Install command: `npm install`

Excel-data importeres under build fra nyeste `.xlsx` i `input/`. Appen er ikke autentisert, så en deployet URL må regnes som delbar for alle som får lenken.
