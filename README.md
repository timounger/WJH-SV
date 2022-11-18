\mainpage WJH-SV

\tableofcontents

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/timounger/WJH-SV)](https://github.com/timounger/WJH-SV/releases/latest)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://github.com/timounger/WJH-SV/blob/master/LICENSE.md)

# Wirtschaftliche Jugendhilfe -  Sozialversicherung

## Über ℹ️

Das Tool erzeugt die Dokumente für die Rückerstattung der Sozialversicherungsbeiträge durch die Wirtschaftliche Jugendhilfe.

## Ausführen 🚀

Es handelt sich sich hierbei um ein Kommandozeilentool. Die Parameter sind beim Tool Aufruf beschrieben.

Für das Ausführen der Executable kann einfachheithalber eine Batch Datei (z.B. `run_WJH-SV.bat`) auf gleicher Ordnerebene erstellt werden.

Inhalt der Batch Datei z.B.

``` bat
WJH-SV.exe --year 2022 --file Tabelle_2022.xlsx --sheet Tabelle
pause
```

## Eigangsdaten ⬇️

Die Berechnung erfolgt aufgrund der Eingangsdaten in Form einer Excel Tabelle (`.xlsx`). Der Datei- sowie der Seitenname müssen in den Aufrufparametern angegeben werden.

Folgende Eigenschaften müssen die Inhalte der Datei aufweisen:

* Die Tabelleneinträge beginnen ab Zeile 3 bzw. die Spaltenbeschriftung bereits in Zeile 2.
* Alle folgenden Tabelleneinträge werden gewertet.
* Die Tabelle darf nicht durch eine leere Zeile unterbrochen sein, da die folgenden EInträge nicht gewertet werden.
* Spalte A: Nachname des Betreuers
* Spalte B: Vorname des Betreuers
* Spalte G: Nachname des Kindes
* Spalte H: Vorname des Kindes
* Spalte I: Wohnort des Kindes
* Spalte M: Buchungsdatum
* Spalte N: `Bezeichnung`  - Ausgezahlter Betrag oder Abzüge an den Betreuer
* Spalte Q: Vermerk für spezielle Betreuung z.B. Vertretung

## Ausgangsdaten ➡️

Die Ausgabe der Excel Dokumente (`.xlsx`) erfolgt in dem Ordner `Output`. In diesem Ordner wird bei jedem erneuten Ausführen ein Ordner mit dem aktuellen Datum erstellt z.B `SV_Berechnung_YYYY-MM-DD_HHhMMmSSs`.

## Verarbeitungslogik 🔃

Ab folgendem Tag des vorherigen Monats erfolgt die Buchung für den nächsten Monat:

* Januar: ab 22.12
* Juli: ab 23.06
* Sonstige: ab 25.xx

Enthält die Spalte `Bezeichnung` eines der folgenden Wörter, erfolgt die Buchung in die untere Tabelle für erhöhtes Pflegegeld:

* `außergewöhnlich`
* `Vertretung`
* `erhöhter Förderbedarf`
