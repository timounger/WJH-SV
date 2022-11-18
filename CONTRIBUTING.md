\page CONTRIBUTING Contributing

[TOC]

## Architektur 🏭️

Das WJH-SV Tool besteht aus einer einzigen Klasse welche alle Funktionalitäten erfüllt.

@startuml cdBonPrinterSWComponents
skinparam titleFontSize 30
skinparam titleFontStyle bold
skinparam packageBorderColor black
skinparam packageFontSize 18
skinparam groupFontStyle bold
skinparam componentBorderColor black
skinparam interfaceBorderColor black
skinparam CollectionsBorderColor black

title SW components of WJH-SV Tool

node "wjh_sv.png" as icon #yellow

package "WJH-SV" as WJH-SV #ededed {
  component "**wjh_sv**" as wjh #63d8ff
  note left of wjh
    Application entry point
  end note
}

node "WJH-SV.exe" as wjh_exe #plum

package "Python Libraries (mostly third party)" #lightgreen {
 [PyInstaller]
 [openpyxl]
}

[wjh] -[#black]r-> [icon]
[wjh] -[#black]-> [openpyxl]
[PyInstaller] -[#black]d-> wjh_exe

@enduml

---

## Versionsverwaltung

Die Version ist in `Source/wjh_sv.py` hinterlegt und wird manuell inkrementiert. Im Ordner `Executable` liegt ein Hilfsskript (`generate_version_file.py`) welches das aktuelle Versionsinfo-File generiert, das zur Exe-Generierung benötigt wird.

---

## Exe-Generierung 🔧

Die Generierung der EXE wird mithilfe des `pyinstaller` gemacht. In der Datei `Executable/generate_executable.bat` sind die dafür notwendigen Parameter spezifiziert. Durch Ausführen des Batch-Skriptes wird im Ordner `Executable/bin` die EXE erzeugt.

---

## GitHub Release Schritte

### Vorbereitung

* [ ] Versionierung hochzählen
* [ ] `B_DEBUG` im Code auf `False` stellen
* [ ] Versionen der Drittanbieterpakete (packages.txt) (und optional Python Version) auf den neuesten Stand aktualisieren
* [ ] Liste der erlaubten, in EXE inkludierten Paketen aktualisieren (`l_allowed_third_party_packages` in `Executable_check_include_packages.py`
* [ ] Nicht benötigte Pakete ggf. explizit exkludieren (--exclude-modules` in `Executable/generate_executable.bat`)

### Tests

* [ ] EXE generieren, alle Funktionalitäten stichprobenartig testen

### Freigabe

* [ ] Nach Merge in master: Commit taggen (z.B `WJH-SV - 1.0.0.0` & Release erstellen mit Links auf Executable in Package Registry
