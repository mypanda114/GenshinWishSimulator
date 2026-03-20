# Genshin Wish Simulator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/pandas-3.0+-green.svg)](https://pandas.pydata.org/)
[![Plotly](https://img.shields.io/badge/plotly-6.6+-orange.svg)](https://plotly.com/)

[English](README.md) | [中文](README-zh-CN.md)

---

## 📖 Introduction
A personal study-oriented wish simulator based on official rules and community consensus, focusing on **Version 6.4 Phase 2** limited character banners (Skirk, Escoffier) and weapon banner (Azurelight, Symphonist of Scents). It supports target setting, weapon banner strategic decision-making, duplicate character conversion statistics, and generates detailed Excel, chart, and Word analysis reports.

---

## ✨ Core Features
- **Accurate Wish Simulation**: Fully reproduces official probabilities, with soft pity using community-recognized models (character banner: +6% per pull after 73; weapon banner: segmented increments).
- **Target Setting & Prediction**: Set your desired goals at startup (e.g., "7 Skirk", "1 Azurelight"), and the system provides estimated pulls needed based on community expectation values.
- **Weapon Banner Strategic Pauses**: Pauses after each 5-star pull, dynamically offering strategy options based on your goals (including the "cancel Epitomized Path and reset" optimization strategy for dual limited weapons).
- **Automatic Stop on Goal Achievement**: Pauses and asks whether to continue when a preset goal is reached, avoiding resource waste.
- **Duplicate Character Conversion Rules**: Strictly follows official descriptions, distinguishing between the 2nd-7th and 8th+ copies of characters for Starglitter/Stella Fortuna conversion, with terminal and report hints.
- **Real Pulls for Limited Items**: Excel includes a "Limited Item Details" sheet, and the Word report features a new section displaying the actual pulls (1–180) for each acquisition of limited characters/weapons, making it easy to verify whether you lost the 50/50.
- **Optimized Weapon Banner Interaction**: When no Epitomized Path is specified, the previously set path is automatically used; confirmation is only required when there is no existing path and none is specified.
- **Adaptive Charts**: The Top 10 5-star items chart automatically adjusts its size based on the number of items, avoiding overly wide bars or excessive white space.
- **Word Report Font Compatibility**: Prioritizes Microsoft YaHei; falls back to SimSun if unavailable, ensuring consistent bold formatting.
- **Fixed 4-star Starglitter Display**: Uses item names to determine character/weapon type, ensuring all 4-star weapons correctly award Starglitter and display it.
- **In-depth Data Analysis**: Generates Excel records, four independent PNG charts, and a Word report containing basic statistics, key metrics, goal achievement analysis, strategy execution logs, luck ratings, and fun Easter eggs.
- **Speed Mode Selection**: Supports "Realistic Pace" (1-second pause every 10 pulls) and "Fast Mode" (no pauses) for different scenarios.
- **Colorful Terminal Output**: 5-star items in bold gold, 4-star in bold purple, with special marks for overflows.
- **Multi-language Support**: All prompts are managed via an i18n module, supporting switching between Chinese and English.

---

## 🎴 Banner Info (Version 6.4 Phase 2)
| Banner Code | 5-star UP Character | 4-star UP Characters | Epitomized Path Code | 5-star UP Weapons | 4-star UP Weapons |
|-------------|---------------------|----------------------|----------------------|--------------------|--------------------|
| **C1** | Skirk | Dahlia, Candace, Charlotte | **W1** | Azurelight | The Flute, The Bell, Favonius Lance, Favonius Codex, Rust |
| **C2** | Escoffier | Dahlia, Candace, Charlotte | **W2** | Symphonist of Scents | The Flute, The Bell, Favonius Lance, Favonius Codex, Rust |

> **Note**: C1 and C2 share pity counters and guarantee status.

---

## 🚀 Quick Start
1. **Install Python 3.9+** and ensure it's added to PATH.
2. **Clone or download this project**:
   ```bash
   git clone https://github.com/yourname/GenshinWishSimulator.git
   cd GenshinWishSimulator
   ```
3. **Create and activate a virtual environment** (recommended):
   ```bash
   python -m venv venv
   .\venv\Scripts\activate   # Windows
   ```
4. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
5. **Run the simulator**:
   ```bash
   python src/genshin_wish_simulator.py
   ```

---

## 🎮 Commands
| Command | Example | Description |
|---------|---------|-------------|
| `C1 50` | C1 50 | 50 pulls on Skirk's banner |
| `C2 1` | C2 1 | 1 pull on Escoffier's banner |
| `W 20 W1` | W 20 W1 | 20 pulls on weapon banner, Epitomized Path set to Azurelight |
| `W 3 W2` | W 3 W2 | 3 pulls on weapon banner, Epitomized Path set to Symphonist of Scents |
| `Y` | Y | Repeat last pull command (count, banner, and Epitomized Path) |
| `S` | S | End pulls and export data |

> **Optimization**: If no Epitomized Path is specified for the weapon banner, the previously set path is automatically used. Confirmation is only required when there is no existing path and none is specified.

---

## 📤 Export Options
After finishing pulls, choose:
1. **Save Excel records only** (character records, weapon records, limited item details, summary sheet, strategy log)
2. **Generate Plotly charts** (four independent PNG images: Character 5-star Pie Chart, Weapon 5-star Pie Chart, 4-star TOP10, 5-star TOP10)
3. **Generate Word report** (includes all KPIs, goal analysis, strategy path, luck rating, limited item details)
4. **Export all**

All files are automatically saved in `output/GenshinWishSim_YYYYMMDD_HHMMSS/`.

---

## 📁 Project Structure
```
GenshinWishSimulator/
├── docs/               # Project documentation
├── src/                # Source code
│   ├── genshin_wish_simulator.py   # Main script
│   ├── i18n.py                     # Translation module
│   └── locales/                    # Language files
│       ├── zh-CN.json              # Chinese translations
│       └── en.json                  # English translations
├── output/             # Wish data output (timestamped subdirectories)
├── requirements.txt    # Dependency list
├── LICENSE             # License file
├── README.md           # This file (English)
├── README-zh-CN.md     # Chinese version
└── .gitignore          # Git ignore file
```

---

## 📦 Dependencies
```txt
pandas==3.0.1
numpy==2.4.3
openpyxl==3.1.5
plotly==6.6.0
python-docx==1.2.0
kaleido==1.2.0
```

Install with: `pip install -r requirements.txt`

---

## 🌍 Multi-language Support
The simulator supports switching between Chinese and English. All prompts are managed via `i18n.py`. Default language is Chinese. To switch to English, modify the initialization line in `src/genshin_wish_simulator.py`:
```python
init_i18n("en")   # Change from "zh-CN" to "en"
```

---

## 🤝 Contributing
Welcome to submit bug reports or feature suggestions via [Issues](https://github.com/yourname/GenshinWishSimulator/issues). If you'd like to contribute code, please fork this repository, create your feature branch, and submit a Pull Request. Ensure your code style is consistent with the existing codebase and passes basic tests.

---

## 📌 Notes
- This tool is **recommended for personal study only**. Simulation results do not represent actual in-game probabilities.
- Character and weapon names are used solely for describing game content. All rights belong to miHoYo. No official images are used.
- This project is licensed under the **MIT License**.

---

[⬆ Back to Top](#genshin-wish-simulator)
```