# 📂 Data Combining Tool

A Python-based solution for **combining a large number of Excel and CSV datasets** into one consistent dataset.

I built this because **Excel Power Query** often struggles or crashes when dealing with **1,000+ files** or datasets containing **millions of rows**.  
This tool handles that scale with ease — while keeping track of where every record came from.

---

## 🚀 What This Tool Does

- ✅ Combine **thousands of Excel and CSV files** efficiently  
- ✅ Automatically add:
  - `source_directory` → folder path of each file  
  - `source_filename` → file name of each record  
- ✅ Handle **datasets larger than 10 million rows** by splitting into multiple Excel files (each up to 500,000 rows)  
- ✅ Detect and log:
  - ⚠️ **Mismatched files** (wrong or missing columns)  
  - ❌ **Failed files** (unreadable or corrupted files)

---

## 🔧 Key Features

### 🧩 Large-Scale Combining
- Works with both **.xlsx** and **.csv** files  
- Handles **millions of rows** using automatic file splitting  
- Processes files in **sorted order** for consistent and repeatable results  

### 📑 Column Normalization
- Standardizes column headers to match a **predefined schema**  
- If headers differ but column count matches, columns are **renamed by position**  
- Files with missing or extra columns are **flagged as mismatched**

### 🗂️ Source Tracking
- Adds two columns to every record:
  - `source_directory`
  - `source_filename`  
- Makes it easy to trace back where each row originated  

### ⚠️ Error Handling & Reporting
- Generates clear reports:
  - `mismatched_files_report.xlsx` → files with missing or extra columns  
  - `failed_files.xlsx` → files that couldn’t be read (with error reason)  
- Skips empty rows automatically  

---

## 📊 Workflow

1. Collect all `.xlsx` and `.csv` files under the input folder  
2. Normalize column headers to match the standard schema  
3. Combine valid files into one dataset  
4. Split into multiple Excel files if total rows exceed 500,000  
5. Generate reports:
   - `mismatched_files_report.xlsx`
   - `failed_files.xlsx`

---

## 📁 Output Files

| File | Description |
|------|--------------|
| `all_combined_data.xlsx` | Output file (if ≤ 500,000 rows) |
| `all_combined_data_1.xlsx`, `all_combined_data_2.xlsx`, ... | Split output files (if > 500,000 rows) |
| `mismatched_files_report.xlsx` | Lists files with missing or extra columns |
| `failed_files.xlsx` | Lists unreadable or broken files |

---

## ⚙️ Configuration

Edit the paths at the top of the script:

```python
# Input folder containing all datasets
folder_path = r"D:\data\cleaningData\ရန်ကုန်"

# Output folder for combined data and reports
output_dir = r"D:\data\sampleCombinedData\ygn_combined_data"
