# 🇮🇳 India State & District Explorer

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![HTML5 / CSS3](https://img.shields.io/badge/Frontend-HTML5%20%2F%20Glassmorphism-purple.svg)](#)
[![JSON Database](https://img.shields.io/badge/Data-GeoJSON%20%2F%20JSON-orange.svg)](#)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

An interactive, responsive web application and data tool that provides comprehensive geographical and administrative information on all States and Union Territories of India, along with their respective districts.

---

## ✨ Features

- 🗺️ **Interactive Geographic Explorer**: Select any Indian State or Union Territory to view all associated administrative districts.
- 🎨 **Modern Glassmorphic UI**: Beautiful dark aesthetic with radial gradients and responsive cards.
- 📊 **Automated Presentation Generator**: Includes Python scripts (`create_ppt.py`, `create_simple_ppt.py`) to generate presentation slide decks from geographical data.
- ⚡ **Lightweight Server**: Instant lightweight HTTP server script (`server.py`) with zero external complex dependencies.

---

## 📁 Repository Structure

```text
india-state-district-explorer/
├── index.html                           # Main explorer web interface
├── states-and-districts.json            # Master administrative dataset of India
├── server.py                            # Lightweight Python HTTP server
├── data.py                              # Data processing utility
├── create_ppt.py                        # Automated PowerPoint presentation builder
├── create_simple_ppt.py                 # Simplified slide generator
├── State_District_Explorer_Presentation.pptx # Pre-generated presentation
└── README.md
```

---

## 🚀 Quick Start

### 1. Clone & Navigate
```bash
git clone https://github.com/harshitmeena200515-cyber/india-state-district-explorer.git
cd india-state-district-explorer
```

### 2. Run the Local Server
```bash
python server.py
```
Open `http://localhost:8000` in your web browser.

---

## 📜 License

MIT License. Free for educational and open-source usage.
