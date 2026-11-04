# Audience Coverage - trial build

A browser-only prototype of the Audience Coverage dashboard. It shows an electorate office which groups it has contacted enough, which it hasn't, and where to go next.

## How it handles data

- A loaded file is read inside the browser tab and is never uploaded. There is no server.
- The import gate removes names, dates of birth (turned into age brackets first), addresses and phone/email values (turned into yes/no flags), and reduces interaction notes to a date and a channel code before anything is counted.
- Groups under 10 electors are never shown.
- The only thing saved in the browser is the office profile: rules about what tag values mean. It holds no personal information.
- A Content Security Policy stops the page loading or sending anything outside its own address. Fonts and the spreadsheet reader are hosted here, not loaded from third parties.

## Files

- `index.html`, `app.js` - the app.
- `sample/demo-electorate-export.csv` - 18,000 fictional electors. Every name, suburb, postcode and contact is invented.
- `vendor/xlsx.full.min.js` - SheetJS 0.18.5, used only to read .xlsx files (licence in `vendor/`).
- `fonts/` - Public Sans and JetBrains Mono (SIL Open Font Licence).

Real office data must not be loaded into this public trial build until the phase 4 security work and legal advice are complete.
