# Newmaterialmasterdatavalidation

**Autor:** Max Keller (DC/BDC)
**Datum:** 29.07.2024
**Version:** 1.0

## 1. Übersicht

Dieses Python-Skript dient zur Validierung von Materialkurzbezeichnungen (`Typkurzbezeichnung`) in einer Excel-Datei. Es führt eine mehrstufige Analyse durch, um festzustellen, ob eine Bezeichnung potenziell unzulässige oder unerwünschte Wörter enthält. Das Skript aktualisiert die Eingabedatei nicht direkt, sondern erstellt eine neue Excel-Datei mit den Validierungsergebnissen in zwei zusätzlichen Spalten.

Der Validierungsprozess wird nur für Zeilen durchgeführt, bei denen die Materialart (`MArt`) einem vordefinierten Wert entspricht (z.B. 'BREX' oder 'KAUF').

Die Validierungslogik umfasst:
- **Stufe 1:** Prüfung auf ungültige Sonderzeichen.
- **Stufe 2:** Prüfung auf Basis von Wortlisten:
    - Deutsche, französische und englische Wörterbücher.
    - Eine Liste deutscher "Stoppwörter" (Füllwörter wie "und", "für" etc.).
    - Eine Positivliste mit erlaubten Wörtern.
    - Eine Prüfung auf deutsche und französische Teilwörter in Komposita (zusammengesetzte Wörter).

## 2. Voraussetzungen

Stellen Sie sicher, dass Python 3.x auf Ihrem System installiert ist. Die folgenden Python-Bibliotheken werden benötigt:

- **pandas:** Zum Einlesen und Verarbeiten der Excel-Daten.
- **openpyxl:** Wird von pandas intern zum Lesen und Schreiben von `.xlsx`-Dateien verwendet.

Sie können die benötigten Bibliotheken einfach über `pip` installieren:
```bash
pip install pandas openpyxl
```

## 3. Ordnerstruktur
Damit das Skript die Eingabedateien und Wörterbücher korrekt findet, muss die folgende Ordnerstruktur eingehalten werden:

.
├── main.py                 # Das Hauptskript zum Ausführen
├── utils/
│   ├── __init__.py
│   ├── io_utils.py         # Hilfsfunktionen für Datei-I/O
│   └── validator.py        # Die Kernlogik der Textvalidierung
└── data/
 ├── input_data/
 │   └── TEST_Data_for_REGEX_V2.xlsx  # Ihre Eingabedatei
 ├── output_data/
 │   └── (hier wird output.xlsx erstellt)
 └── dictionaries/
     ├── german_stopwords.txt
     ├── german_words.txt
     ├── french_words.txt
     ├── english_words.txt
     └── positive_words.txt


## 4. Konfiguration
Alle wichtigen Einstellungen können direkt am Anfang der main.py-Datei angepasst werden.

--- CONFIGURATION ---
Dateinamen und Spalten
SOURCE_FILENAME = 'data\input_data\TEST_Data_for_REGEX_V2.xlsx'
OUTPUT_FILENAME = 'data\output_data\output.xlsx'
SOURCE_COLUMN_NAME = 'Typkurzbezeichnung' # Spalte mit dem zu prüfenden Text
CONDITION_COLUMN = 'MArt'                 # Spalte für die Bedingung
TARGET_FLAG_COLUMN = 'B'                  # Ergebnisspalte für das Flag (1/"")
TARGET_REASON_COLUMN = 'lsg'              # Ergebnisspalte für den Grund

Bedingungswerte
CONDITION_VALUES = ['BREX', 'KAUF']       # Validierung nur für diese Werte in 'MArt'


SOURCE_FILENAME: Pfad zu Ihrer Excel-Eingabedatei.
OUTPUT_FILENAME: Pfad und Name für die zu erstellende Excel-Ausgabedatei.
SOURCE_COLUMN_NAME: Name der Spalte, die die zu validierenden Materialbezeichnungen enthält.
CONDITION_COLUMN / CONDITION_VALUES: Steuern, welche Zeilen basierend auf dem Wert in dieser Spalte validiert werden.
TARGET_FLAG_COLUMN / TARGET_REASON_COLUMN: Namen der Spalten in der Zieldatei, in die die Ergebnisse geschrieben werden.
## 5. Ausführung
Um das Skript auszuführen, öffnen Sie eine Kommandozeile (Terminal, PowerShell etc.), navigieren Sie in das Hauptverzeichnis des Projekts (dorthin, wo main.py liegt) und führen Sie den folgenden Befehl aus:

python main.py


Das Skript gibt während der Ausführung Statusmeldungen auf der Konsole aus:

--- Starting: Advanced Material Master Data Validation ---
Loading data from 'TEST_Data_for_REGEX_V2.xlsx'...
Initializing validator and loading resources...
Running conditional validation pipeline...
Duplicating template from '...' to '...'...
Opening the new workbook to write results...
Writing data to column 'B'...
Writing data to column 'lsg'...
Saving final workbook to 'data\output_data\output.xlsx'...
Successfully saved the file!


**SUCCESS: Workflow completed without errors.** 



Nach erfolgreicher Ausführung finden Sie die Ergebnisdatei output.xlsx im Ordner data/output_data/.

## 6. Fehlerbehandlung
Das Skript verfügt über eine grundlegende Fehlerbehandlung für häufige Probleme:

FileNotFoundError: Wird ausgelöst, wenn die Eingabedatei oder ein Wörterbuch nicht gefunden wird. Überprüfen Sie die Pfade in der Konfiguration und die Ordnerstruktur.
ValueError: Wird ausgelöst, wenn eine der benötigten Spalten (Typkurzbezeichnung, MArt) in der Excel-Datei fehlt.
PermissionError: Tritt typischerweise auf, wenn die Ausgabe-Excel-Datei bereits existiert und in einem anderen Programm (z.B. Microsoft Excel) geöffnet ist. Schließen Sie die Datei und führen Sie das Skript erneut aus.
