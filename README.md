# 📊 RPA Automation Suite – Data Integration & File Standardization

## 📌 Overview
This repository contains two Robotic Process Automation (RPA) scripts designed to automate data ingestion and file standardization processes in a manufacturing analytics environment.

The automations focus on:
- Automated extraction of data from Excel Query (.iqy) -- sharepoint default file -- and transform to excel (xlsX)
- Schema-aligned data integration into a historical database (columns to columns)
- Standardization of file naming conventions, and save on your repository

These processes reduce manual workload, improve data consistency, and ensure reliable datasets for analytics and reporting.

---

## 🤖 RPA-001: Automated data ingestion & Schema Alignment by Sharepoint default file

### 🧠 Technical Classification
- ETL Automation (Extract, Transform, Load)
- Schema Alignment Process
- Automated Data Integration Pipeline
- Excel Query Table Extraction (IQY → XLSX)


### 🔎 Description
This automation extracts data from `.iqy` files (Excel Web Queries), converts them into structured Excel format, and integrates the data into a historical database while enforcing schema consistency.

---

### ⚙️ Process Flow

#### 1. File Detection
- Scans the Downloads folder for `.iqy` files
- Selects the most recent file

#### 2. Data Extraction
- Opens the file using Excel COM automation (`win32com`)
- Executes:
  - `RefreshAll()` to query the source system
  - Waits for asynchronous query completion

#### 3. File Conversion
- Converts `.iqy` file into `.xlsx`
- Overwrites existing files silently

#### 4. Data Transformation
- Loads:
  - New dataset
  - Historical database (`Historic.xlsx`)
- Cleans column names (removes spaces)
- Aligns schema:
  - Keeps only matching columns
  - Adds missing columns with `NULL`

#### 5. Data Load
- Reorders columns to match database structure
- Appends new records
- Overwrites the existing database


### 🎯 Business Value
- Eliminates manual data extraction
- Ensures schema consistency
- Prevents integration errors
- Supports reliable historical tracking

---

## 🤖 RPA-002: Regex-Based File Naming Standardization

### 🧠 Technical Classification
- File System Automation
- Pattern Recognition (Regex)
- Data Label Normalization


### 🔎 Description
This automation standardizes file names based on a predefined pattern, ensuring consistency and improving file organization for downstream processes.

---

### ⚙️ Process Flow

#### 1. File Detection
- Searches for files matching pattern:

#### 2. Pattern Extraction
- Uses **Regular Expressions (regex)** to extract day (`XX`)

#### 3. File Renaming
- Converts filenames into standard date format:
