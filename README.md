# PPTX to LaTeX Beamer Converter

Konvertiert PowerPoint-Präsentationen in LaTeX Beamer-Folien inkl. Bilder und Videos.

## Voraussetzungen
* **Python 3.10+**
* **LaTeX:** TeX Live
* **Ollama:** Muss lokal laufen
* **Videowiedergabe:** Adobe Acrobat Reader 

## Installation

1. **Abhängigkeiten installieren:**
   ```bash
   pip install -r requirements.txt
   ```
2. **Modell laden:**
   ```bash
    ollama pull qwen2.5-coder:7b
    ```
3. **Starten:**
   ```bash
    python main.py
    ```