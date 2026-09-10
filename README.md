# MaLo-ID & Stromzähler-Zuordnung

Streamlit-App zur Verwaltung der MaLo-ID- und Stromzähler-Zuordnung für Mehrfamilienhäuser
(bis zu 15 Parteien pro Haus), inklusive Excel-Import und Excel-/PDF-Export.

## Lokale Installation (macOS)

1. Terminal öffnen und in den Projektordner wechseln:
   ```
   cd Pfad/zum/malo_zuordnung
   ```
2. Virtuelle Umgebung anlegen und aktivieren (einmalig):
   ```
   python3 -m venv venv
   source venv/bin/activate
   ```
3. Abhängigkeiten installieren:
   ```
   pip install -r requirements.txt
   ```
4. App starten:
   ```
   streamlit run app.py
   ```
   Der Browser öffnet sich automatisch unter http://localhost:8501.

Beim erneuten Start reicht ab Schritt 2 nur noch `source venv/bin/activate` und
`streamlit run app.py` – die Daten bleiben in einer lokalen SQLite-Datenbank unter
`~/.malo_zuordnung/malo_zuordnung.db` erhalten.

## Was wurde überarbeitet

- Datenbankverbindungen werden jetzt zuverlässig geschlossen (kein Verbindungsleck mehr
  bei häufigen Interaktionen).
- Häuser werden in der Objektliste nach Hausnummer numerisch statt alphabetisch sortiert
  (z. B. "2" vor "10").
- Excel-Import verträgt jetzt doppelt zugeordnete Spalten (z. B. "Name" und "Mieter"
  gleichzeitig in der Quelldatei), ohne abzubrechen.
- Neu: Bestehende Zuordnungen können direkt in der Tabelle "Zuordnungsliste" bearbeitet
  und über den Button "Änderungen speichern" persistiert werden – vorher war nur Anlegen
  und Löschen möglich.
