# [Projektname] - PPTX to LaTeX Beamer Converter

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![LaTeX](https://img.shields.io/badge/latex-beamer-green)
![Status](https://img.shields.io/badge/status-active-success)

> **Verwandelt PowerPoint-Präsentationen in LaTeX Beamer-Folien.**
> *Powered by LLMs (Ollama/OpenAI) & Docling.*

---

## Table of Contents

- [Über das Projekt](#-über-das-projekt)
- [Features](#-features)
- [Architektur & Pipeline](#-architektur--pipeline)
- [Voraussetzungen](#-voraussetzungen)
- [Installation](#-installation)
- [Verwendung](#-verwendung)
- [Multimedia & Video-Support](#-multimedia--video-support)
- [Konfiguration](#-konfiguration)
- [Troubleshooting](#-troubleshooting)
- [Lizenz](#-lizenz)

---

## 📖 Über das Projekt
[...]

##  Features
* **Präzise Layout-Analyse:** Nutzt `python-pptx` und `Docling` zur Extraktion von Geometrie und Inhalt.
* **LLM-Powered:** Ein KI-Agent wandelt JSON-Daten in LaTeX-Code um.
* **Multimedia-Support:** Bettet Videos (`.mp4`) und Bilder automatisch ein.
* **Absolute Positionierung:** Pixelgenaue Nachbildung des PPTX-Layouts mittels `textpos`.

## 🏗 Architektur & Pipeline
Das Tool arbeitet in einer 5-stufigen Pipeline:
1.  **Media Extraction:** Extrahiert Videos, Poster-Bilder und Grafiken (Deep XML Scan).
2.  **Layout Analysis:** Erfasst Koordinaten und Inhaltstypen (Text, Code, Tabellen).
3.  **Data Optimization:** Merging von Docling-JSON und extrahierten Mediendaten.
4.  **LaTeX Generation:** Ein LLM-Agent generiert den Code pro Folie.
5.  **Compilation:** Automatische Kompilierung mittels `pdflatex` zu PDF.

## 🛠 Voraussetzungen
* **Python 3.10+**
* **LaTeX Distribution:** TeX Live.
* **Perl:** (Wird oft von LaTeX-Skripten benötigt).
* **LLM Server:** Lokales Ollama.

## 📦 Installation
```bash
git clone [https://github.com/dein-user/projektname.git](https://github.com/dein-user/projektname.git)
cd projektname
pip install -r requirements.txt