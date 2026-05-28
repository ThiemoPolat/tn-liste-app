# TN-Liste Web-App

Diese kleine Web-App befüllt eine vorhandene Word-Vorlage für mündliche Prüfungen.

## Funktion

1. Word-Vorlage `.docx` hochladen
2. TN-Liste einfügen
3. Auf **Vorlage befüllen** klicken
4. Fertige Word-Datei herunterladen

Die App verändert nicht den Kopfbereich der Vorlage. Sie nutzt die vorhandene Tabelle und füllt nur die mittlere Spalte mit den Teilnehmerpaaren.

## Installation

Python installieren, dann im Ordner der App:

```bash
pip install -r requirements.txt
```

## Starten

```bash
streamlit run app.py
```

Danach öffnet sich die App im Browser.

## TN-Listen-Format

Die App erwartet pro Teilnehmer ungefähr dieses Muster:

```text
GK
Ginette Kouowou
D384291
Initial
PM
Pamela Marthe Huguette MBACK
D384295
Initial
```

Die Teilnehmer werden automatisch paarweise eingefügt:

```text
Ginette Kouowou (1)
Pamela Marthe Huguette MBACK (2)
```