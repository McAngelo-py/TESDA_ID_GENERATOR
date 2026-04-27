# 🎓 TESDA ID Generator

A powerful, local-only Python desktop application designed for batch generating TESDA IDs by replacing placeholders in `.docx` templates. It features a modern, user-friendly interface with support for dark and light modes.

---

## ✨ Features

- **Template Based**: Use any `.docx` file as a template.
- **Auto-Detection**: Scans your template for specific placeholders automatically.
- **Bulk Processing**: Generate dozens of IDs in seconds.
- **CSV Support**: Import student/person data directly from Excel or CSV files.
- **Smart ID Generation**: Automatically increments ID numbers (e.g., `2026-000`, `2026-001`, etc.).
- **Privacy First**: All processing happens locally on your computer. No data is sent to the internet.
- **Theming**: Toggle between a clean Light Mode and a sleek Dark Mode.

---

## 🛠️ Getting Started

### 1. Installation
Ensure you have Python installed, then install the required libraries:
```bash
pip install -r requirements.txt
```

### 2. Run the App
```bash
python src/main.py
```

---

## 📝 How to Use

### Step 1: Prepare your Template
Your Word document (`.docx`) must contain specific placeholder text that the app will look for. The app is case-insensitive but the text must match these exactly:

| Data Field | Placeholder in Word |
| :--- | :--- |
| **Name** | `NAME HERE` (or custom via UI) |
| **ID Number** | `2026-000` (auto-increments) |
| **Course** | `COURSE HERE` (selected via dropdown) |
| **Student Code** | `CODE HERE` |
| **Address** | `HOME ADDRESS HERE` |
| **Blood Type** | `BLOOD TYPE HERE` |
| **Sex** | `SEX HERE` |
| **Emergency Name** | `EMERGENCY NAME HERE` |
| **Emergency Contact** | `EMERGENCY NUMBER HERE` |
| **Emergency Address** | `EMERGENCY ADDRESS HERE` |

> **Note:** If your template has two IDs per page (Front & Back), the app detects this and repeats the same person's data twice automatically.

### Step 2: Load Template & Auto-Detect
1. Click **Upload .docx** and select your template.
2. Click **Auto-detect Placeholders**. The app will show you exactly how many placeholders it found.

### Step 3: Input Data
You can either type the data manually into the text boxes (one person per line) or use a **CSV file**.

#### **CSV Format Options:**
Prepare a CSV without headers in one of these column orders:
- **9 Columns**: `Name, Code, Address, Blood Type, Sex, Gender, Emergency Name, Emergency Number, Emergency Address`
- **8 Columns**: `Name, Address, Blood Type, Sex, Gender, Emergency Name, Emergency Number, Emergency Address`
- **4 Columns**: `Name, Address, Blood Type, Sex`
- **1 Column**: `Name` (only)

### Step 4: Generate
Click **Generate Updated ID Documents**. A new Word file with a timestamp will be created in the same folder as your template.

---

## 🎨 Themes
Click the **🌙 Dark Mode** or **☀ Light Mode** button in the top right corner to switch the interface style to your preference.

---

## 📦 Requirements
- Python 3.x
- `python-docx`
- `Pillow` (for the UI icons)
