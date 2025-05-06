# 🧪 AFM Data Automation Toolkit – Bruker JPK software ~Focus~

Welcome to the **AFM Data Automation Toolkit**, a collection of scripts, macros, and ideas designed to streamline the data processing workflow of **Atomic Force Microscopy (AFM)** measurements, specifically those generated using **Bruker JPK software**.

## 📌 What is This Repository?

This repository is intended as a practical and evolving toolkit to:
- Automate repetitive data treatment tasks from AFM force measurements.
- Extract and organize relevant measurement data, such as **adhesion forces**.
- Provide ready-to-use Excel VBA macros and Python scripts.
- Help researchers and lab technicians save time and reduce human error in data processing.

## 🧰 Current Features

- ✅ **Excel VBA Macros** for batch processing `.tsv` files exported from Bruker JPK:
  - Automatically reads adhesion values labeled as `"Adhesion [N]"`.
  - Converts values to picoNewtons (pN).
  - Computes summary statistics (average, median).
  - Outputs a clean, formatted "Resume" sheet.
  - 📂 **Folder-level automation** to process multiple JPK export files at once.
  - 🧼 Sheet name cleaning, dynamic column detection, and safety checks.

- ✅ **New adds coming soon**

## 🧪 About AFM and Bruker JPK

**Atomic Force Microscopy (AFM)** is widely used in material science and cell biology for measuring surface forces at the nanoscale. 
Bruker JPK software exports raw data in `.tsv` format, which can be large and inconsistent in structure. 
This toolkit helps bridge that gap by automating the conversion, cleaning, and summarization of key metrics such as **adhesion forces**.

## 🚀 How to Use
1. Download or clone this repository.
2. Open the Excel file containing the macros (or add them to your own).
3. Run the macro `ProcessTSVsInBatches` and select your `.tsv` data folder.
4. Let the macro process your data and generate a summarized workbook.

## 🔧 Requirements
- Microsoft Excel (with macro support enabled)
- Basic familiarity with VBA if customizing
- Bruker JPK `.tsv` export files as input

## 📦 Coming Soon
Atuomatic graphs and  all columns in one

## 👤 Author
Developed and maintained by [Adriana].  
Contributions and suggestions are welcome — feel free to open issues or pull requests.

---

📁 This repository is meant to grow with ideas and tools that make AFM data analysis more efficient and reproducible.
