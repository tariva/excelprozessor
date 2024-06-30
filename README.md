# excelprozessor

Excel-Datenprozessor
Diese Anwendung ermöglicht die Verarbeitung und den Export von Excel-Daten basierend auf einer konfigurierten Spaltenzuordnung.

Übersicht
Die Anwendung liest Daten aus einer Quelldatei (Excel), filtert und transformiert sie gemäß einer Konfigurationsdatei und exportiert die verarbeiteten Daten in eine Zieldatei (Excel).

Konfigurationsdatei
Die Konfigurationsdatei (config.json) steuert das Verhalten der Anwendung. Sie enthält folgende Parameter:

mapping: Ein Objekt, das die Zuordnung von Spaltenüberschriften zwischen der Quelldatei und der Zieldatei definiert.
sourceFile: Der Name der Quelldatei.
destFile: Der Name der Zieldatei.
adjustmentkeys: (Optional) Schlüssel für das alte System.
keyColumn: Der Schlüsselwert, der in der Quelldatei mit dem Wert in der Zieldatei verglichen wird.
worksheetName: Der Name des Arbeitsblatts im alten System.
sourceMappingRow: Die Zeile in der Quelldatei, in der die Spaltenüberschriften stehen.
sourceStartDataRow: Die Zeile in der Quelldatei, ab der die Daten gelesen werden.
sourceWorksheetName: Der Tabellenblattname in der Quelldatei.
destMappingRow: Die Zeile in der Zieldatei, in der die Spaltenüberschriften stehen.
destStartDataRow: Die Zeile in der Zieldatei, ab der die Daten eingefügt werden.
destWorksheetName: Der Tabellenblattname in der Zieldatei.
Beispiel einer Konfigurationsdatei (config.json)

```json


{
  "mapping": {
    "M11": "M11",
    "M9": "M9",
    "S21": "S21",
    "S22": "S22",
    "STR": "STR",
    "M10": "M10"
  },
  "sourceFile": "ATMOSA-RoLI-20240603-5.xlsx",
  "destFile": "VORLAGE-ATMOSA-IMPORT-EXPORT.xlsx",
  "adjustmentkeys": [],
  "keyColumn": "STR",
  "worksheetName": "Rohrschellen",
  "sourceMappingRow": 1,
  "sourceStartDataRow": 3,
  "sourceWorksheetName": "RoLI",
  "destMappingRow": 2,
  "destStartDataRow": 5,
  "destWorksheetName": "Rohr_RH"
}
```

Ablauf der Anwendung
Vorbereitung der Umgebung:

Ein temporäres Verzeichnis (tmp) wird erstellt, um Dateien zwischenzuspeichern.
Auswahl der Dateien:

Der Benutzer wählt die Quelldatei und die Zieldatei aus dem angegebenen Verzeichnis aus.
Verarbeitung der Quelldatei:

Die Daten aus der Quelldatei werden gelesen und basierend auf der Konfiguration gefiltert und transformiert.
Es wird sichergestellt, dass die Schlüsselspalte (keyColumn) in beiden Dateien vorhanden ist.
Erstellung eines Mappings:

Ein Mapping der Spalten zwischen der Quelldatei und der Zieldatei wird erstellt.
Die Daten aus der Quelldatei werden anhand der Schlüsselspalte in die entsprechenden Zeilen der Zieldatei kopiert.
Export der verarbeiteten Daten:

Die transformierten Daten werden in eine neue Excel-Datei exportiert.
Die Datei wird unter dem im Konfigurationsparameter outputPath angegebenen Pfad gespeichert.
Fehlerbehandlung
Die Anwendung umfasst eine Fehlerbehandlung für häufige Dateizugriffsfehler wie:

EBUSY: Wenn die Ressource beschäftigt oder gesperrt ist.
ENOENT: Wenn die Datei nicht gefunden wird.
