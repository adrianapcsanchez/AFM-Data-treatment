# 🧪 AFM Data Treatment Automation – Bruker JPK Focused

Welcome to the **AFM Data Automation Toolkit**, a collection of scripts and macros designed to streamline **Atomic Force Microscopy (AFM)** data processing workflows, especially for data generated with **Bruker JPK (jpkspm)** software.

## 📌 What is this repository?

This repository helps you:
- Automate repetitive AFM data processing tasks.
- Extract and organize key measurements such as **Adhesion Force** and **Young’s Modulus**.
- Generate quick summary outputs in Excel.
- Reduce manual errors and save analysis time.

## 🧰 Current tools

### 1) VBA macro: `ProcessTSVsinOne`
- Batch reads `.tsv` files.
- Detects and extracts **Adhesion [N]** values.
- Converts values to **pN**.
- Computes **average, median, and stdev.p**.
- Writes a central `Resume` sheet.

### 2) VBA macro: `TSVsYoungModulus`
- Batch reads `.tsv` files.
- Detects and extracts **Young's Modulus [Pa]**.
- Converts values to **kPa**.
- Computes **average, median, and stdev.p**.
- Writes a central `Resume` sheet.

### 3) Python script: `YMProfiles.py`
- Recursively finds `.txt/.TXT` force-curve files.
- Cleans and parses force-curve points.
- Uses interpolation + PCA + KMeans to select **4 representative curves**.
- Exports selected curves to `YMProfiles.xlsx`.
- Automatically adds a scatter chart with offset curves.

## ⚠️ Important JPK export instruction (required before using the Python script)

For this code to work with your Young Modulus curves, use **Batch Processing** in JPK and:
1. Select **Save Force Curves**.
2. Check the option **Save in TXT**.
3. Start the YM curve analysis.

The export folder will be saved in the **same folder as your original data**. Then, for this script, select the **`curves` folder inside the new folder created by JPK**.

## 🚀 How to use

### VBA macros
1. Open Excel with macros enabled.
2. Import/run `ProcessTSVsinOne` or `TSVsYoungModulus`.
3. Select your input folder and save the output `.xlsx` file.

### Python representative YM curves script
1. Install dependencies:
   ```bash
   pip install numpy pandas scipy scikit-learn openpyxl
   ```
2. Run:
   ```bash
   python YMRepresentativeCurves.py
   ```
3. Select:
   - Input folder (the exported JPK `curves` folder).
   - Output folder for the Excel file.
4. The script creates `YMProfiles.xlsx` with the selected curves and chart.

## 🔧 Requirements
- Microsoft Excel (for VBA macros).
- Python 3.9+ (for `YMRepresentativeCurves.py`).
- Bruker JPK exports (`.tsv` for macros, `.txt` force curves for Python script).

## 👤 Author
Developed and maintained by Adriana.
Contributions and suggestions are welcome.
