# 🧪 AFM Data treatment Automation – Bruker JPK software Focused

Welcome to the **AFM Data Automation Toolkit**, a collection of scripts, macros, and ideas designed to streamline the data processing workflow of **Atomic Force Microscopy (AFM)** measurements, specifically those generated using **Bruker JPK software - jpkspm software**.

## 📌 What is This Repository?

This repository is intended as a practical and evolving toolkit to:
- Automate repetitive data treatment tasks from AFM force measurements.
- Extract and organize relevant measurement data, such as **adhesion forces** and **Young Modulus**
- Provide ready-to-use Excel VBA macros
- Help researchers and lab technicians save time and reduce human error in data processing.

## 🧰 Current Features

- ✅ **Excel VBA Macros** for batch processing `.tsv` files exported from Bruker JPK:
  - Automatically reads adhesion values labeled as `"Adhesion [N]"`.
  - Converts values to picoNewtons (pN).
  - Computes summary statistics (average, median).
  - Outputs a clean, formatted "Resume" sheet.
  - 📂 **Folder-level automation** to process multiple JPK export files at once.
  - 🧼 Sheet name cleaning, dynamic column detection, and safety checks.

- ✅ **Excel VBA Macros** for batch processing `.tsv` files exported from Bruker JPK:
  -Automatically reads adhesion values labeled as `"Young Modulus [Pa]"`.
  -📂 Batch processes .tsv files from a selected folder
-  📉 Extracts Young’s Modulus [Pa] data and converts it to KPa
 - Converts values to kiloPascal (kPa).
-  📊 Automatically calculates:
      -Average
      -Median
      -Population standard deviation
-📈 Compiles all results into a centralized Resume sheet
-💾 Prompts user to save the final compiled .xlsx file

### 3) Python script: `YMcells.py`
From all graphs, the code will select 4 representative curves, from 4 clusters 
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


- ✅ **New adds coming soon**

## 🧪 About AFM and Bruker JPK

**Atomic Force Microscopy (AFM)** is widely used in material science and cell biology for measuring surface forces at the nanoscale. 
Bruker JPK software exports raw data in `.tsv` format, which can be large and inconsistent in structure. 
This toolkit helps bridge that gap by automating the conversion, cleaning, and summarization of key metrics such as **adhesion forces**.

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
- Microsoft Excel (with macro support enabled)
- Basic familiarity with VBA if customizing
- Bruker JPK `.tsv` export files as input


## 📦 Coming Soon
Automatic graphs and  all columns in one

## 👤 Author
Developed and maintained by [Adriana].  
Contributions and suggestions are welcome — feel free to open issues or pull requests.

---

📁 This repository is meant to grow with ideas and tools that make AFM data analysis more efficient and reproducible.
