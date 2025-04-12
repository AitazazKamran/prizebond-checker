# 🏆 Prize Bond Checker using Python & Selenium

This project automates checking **Pakistani prize bonds** from the [HamariWeb Prize Bond Portal](https://hamariweb.com/finance/prizebonds/) using **Python, Selenium**, and **Excel files**.

---

## 📌 Features

📅 Automatically selects bond denominations (Rs. 100, 200, 750, etc.)  
🔢 Submits and checks bond numbers from an Excel file  
📈 Saves winning results into a new Excel file (`results.xlsx`)  
🚫 Handles both win and "No Win" responses  
🚀 Works headlessly with browser automation  

---

## 📂 Input Format (`Testing_data.xlsx`)

The input Excel file must have **bond amounts as column headers** and bond numbers listed under them. For example:

| 100.00 | 200.00 | 750.00 |
|--------|--------|--------|
| 000009 | 000090 | 009999 |
| 000010 | 000070 | 000890 |
|        |        | 009090 |

Make sure:
- All bond numbers are numeric
- The format is consistent with Excel columns as bond types

---

## 🚀 How to Run

### 1. Install Dependencies

```bash
pip install selenium pandas openpyxl
```

### 2. Download ChromeDriver

- Download it from: [https://sites.google.com/chromium.org/driver/](https://sites.google.com/chromium.org/driver/)
- Make sure the **ChromeDriver version matches** your installed Chrome browser.
- Add ChromeDriver to your system’s **PATH**.

### 3. Place Files

- Put `prizebond_checker.py` and `Testing_data.xlsx` in the same folder.

### 4. Run the Script

```bash
python prizebond_checker.py
```

### 5. Check Results

- Output is saved in `results.xlsx`, with separate sheets for each bond amount.

---

## 🧰 Tech Stack

- **Python 3**
- **Selenium WebDriver**
- **Pandas**
- **OpenPyXL**
- **HamariWeb Prize Bond Checker**

---



## 👨‍💼 Aitazaz Kamran


