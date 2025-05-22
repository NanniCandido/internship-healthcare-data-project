# 🏥 PCC Preprocessing Data – Internship Project

**Institution**: Nova Scotia Health Authority  
**Duration**: April – May 2025  
**Role**: Data Analyst Intern  
**Note**: Due to confidentiality, code and data files cannot be shared.

---

## 📌 Project Overview

During my internship with the **Primary Health Care** team at Nova Scotia Health, I developed a set of automated Python scripts to **standardize and consolidate appointment and triage data** collected from clinics across multiple zones in Nova Scotia. Each clinic uses a different **EMR (Electronic Medical Record)** system, resulting in highly inconsistent data formats.

The processed data feeds **Tableau dashboards**, supporting health leaders in making data-informed decisions and improving **access to primary care**.

---

## ⚙️ Tools & Technologies

- **Python**
  - `pandas` – data manipulation
  - `argparse` – command-line input validation
  - `win32com.client` – Excel automation via Win32 API
- **JSON** – configuration files for column mapping
- **Excel / CSV** – input and output formats
- **Tableau** – interactive dashboards (consuming processed data)

---

## 🧠 Key Functionalities

- Accepts **raw CSV or Excel files** (inconsistent structure)
- Removes:
  - Extra/empty rows
  - Duplicate headers
  - Invalid columns
- **Dynamically identifies the correct header row** using known values
- Uses a `.json` config file to **map and rename columns**
- Fills missing values and adds **Zone, PCC, and type of care (PCC or VCNS)**
- Appends to a **master Excel file** using Win32 API (preserving formatting and preventing duplication)

---

## ✅ Results & Impact

- ✅ **Reduced manual work** previously required to clean and merge data
- ✅ **Improved consistency and quality** of data used in official Tableau dashboards
- ✅ **Scalable solution** — can be reused by multiple zones with minimal adjustments
- ✅ **Contributed directly** to better visibility and decision-making in **primary care delivery**

---

## 📚 Skills Demonstrated

- Python scripting for real-world data workflows
- Data cleaning and preprocessing in `pandas`
- Excel automation using Win32 API
- Configuration-driven development with JSON
- Command-line tool design with `argparse`
- Collaboration with healthcare professionals and analysts

---

## 🔒 Disclaimer

This repository is **documentation-only** due to the **sensitive and proprietary nature** of the data and scripts developed for Nova Scotia Health.

---

## 📬 Contact

Elaine Candido da Silva – [LinkedIn](https://www.linkedin.com/in/elaine-da-silva-candido/) | nannicandido@gmail.com
