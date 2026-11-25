# vibecodingForDummies

En enkel lånelogg for IKT-materiell.

## Kjøre appen

Åpne `index.html` i en nettleser. Data lagres lokalt i nettleseren.

## Excel-dashbord

`dashboard.html` er et dashbord som bruker en Excel-fil som datakilde. Åpne siden og velg en fil, eller dra den inn på siden (.xlsx, .xls, .ods eller .csv). Filen leses bare i nettleseren og lastes ikke opp noe sted.

- Kolonnene gjenkjennes automatisk som tall, dato eller tekst. Norske tall- og datoformater støttes.
- Hvert ark i arbeidsboken blir en egen fane.
- Filtre, nøkkeltall, to grafer og en sorterbar tabell bygges ut fra innholdet.
- «Last inn på nytt» leser filen på nytt etter at du har lagret endringer i Excel.
- «Last ned eksempelfil» gir en Excel-fil som viser formatet.

Siden henter bibliotekene SheetJS og Chart.js fra cdnjs, så den trenger internettilgang.
