# [Projektname] - PPTX to LaTeX Beamer Converter

> **Verwandelt PowerPoint-Präsentationen in LaTeX Beamer-Folien.**
> *Powered by LLMs (Ollama) & Docling.*

##  Features
* **Präzise Layout-Analyse:** Nutzt `python-pptx` und `Docling` zur Extraktion von Geometrie und Inhalt.
* **LLM-Powered:** Ein KI-Agent wandelt JSON-Daten in LaTeX-Code um.
* **Multimedia-Support:** Bettet Videos (`.mp4`) und Bilder automatisch ein.
* **Absolute Positionierung:** Pixelgenaue Nachbildung des PPTX-Layouts mittels `textpos`.

## Architektur & Pipeline
Das Tool arbeitet in einer 5-stufigen Pipeline:
1.  **Media Extraction:** Extrahiert Videos und Bilder.
2.  **Layout Analysis:** Erfasst Koordinaten und Inhaltstypen (Text, Code, Tabellen, Media).
3.  **Data Optimization:** Merging von Docling-JSON und extrahierten Mediendaten.
4.  **LaTeX Generation:** Ein LLM-Agent generiert den Code pro Folie.
5.  **Compilation:** Automatische Kompilierung mittels `pdflatex` zu PDF.

## Voraussetzungen
* **Python 3.10+**
* **LaTeX Distribution:** TeX Live.
* **LLM Server:** Lokales Ollama.
* **Adobe Acrobat Reader:** Erforderlich für die Videowiedergabe im PDF.

## Anwendung
```bash
1. Stelle sicher, dass Ollama läuft und das Modell verfügbar ist:   "qwen3:8b"
2. Passe ggf. die Einstellungen in main.py (class Config).
3. Pipeline ausführen Starte das Skript:
python main.py