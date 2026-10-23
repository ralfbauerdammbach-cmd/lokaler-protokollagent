# Lokaler Protokollagent

Kostenloser Windows-Agent zur Transkription und Word-Protokollerstellung mit Faster Whisper und Ollama.

## Start
1. Repository auf Windows klonen.
2. Ollama starten und ggf. `ollama pull qwen2.5:7b` ausführen.
3. `STARTEN.bat` doppelklicken.
4. Dateien auswählen und Ergebnisprotokoll erzeugen.

Benötigt Python 3.10+; fehlende Python-Pakete werden beim Start installiert. Erste Modell-Downloads benötigen Internet. Keine kostenpflichtigen APIs.

## Grenzen
Noch keine automatische Sprecherdiarisierung, keine Entdoppelung überlappender Teile. Zeitstempel beginnen pro Datei neu. Ergebnisse vor Weitergabe anhand des Transkripts prüfen.

## Datenschutz
Keine Audioaufnahmen, Transkripte oder Protokolle in GitHub hochladen. `.gitignore` beachten. Einwilligungen zur Aufzeichnung sicherstellen.

## Status
Windows-End-to-End-Test mit echter Aufnahme steht aus.
