# Office2PDF v5.x (Office Ultimate PDF Tool)

> Designed for office workers who need reliable, offline batch PDF processing.


![1](https://github.com/user-attachments/assets/4e234712-dffc-4182-9ad2-eb7483f36f3d)
![2](https://github.com/user-attachments/assets/fc89aae9-fba6-4a31-94aa-e61b9c8eb659)



**A practical, office-focused batch PDF conversion & processing tool.**  
Convert Office files and images to PDF in bulk, then apply watermarks, page numbers, splitting/merging, and flexible naming rules with just a few clicks.

> **Offline conversion:** This app uses Microsoft Office via COM automation.  
> Microsoft Office must be installed, and conversions are performed locally (offline).

---

## ⚠️ Maintenance & Support Policy

**Important: The author is not a professional software developer.**  
This tool was created to improve personal efficiency, and it may be difficult for the author to provide complex fixes or advanced support.

- **For engineers:** Pull requests for improvements, bug fixes, and new features are welcome.
- **For users:** This software is provided **“as-is”**, without warranty.

---

## 🌟 Key Features (v5.x)

### Batch Conversion
- Convert the following file types to PDF:
  - **Word**: `.doc`, `.docx`
  - **Excel**: `.xls`, `.xlsx`, `.xlsm`
  - **PowerPoint**: `.ppt`, `.pptx`
  - **Images**: `.jpg`, `.jpeg`, `.png`
  - **PDF**: re-save/normalize and process

### Flexible Split / Merge
- **Merge all** inputs into a single PDF
- **Split by page** (Word / PowerPoint / PDF / Excel)
- **Split by sheet** (Excel)

### Range Selection
- Specify target pages (Word / PowerPoint / PDF) by double-clicking an item:
  - Examples: `1-3,5,8-` / `2` / `-3` / `1-`
- For Excel: select one or more **sheet names**

### Watermarks & Page Numbers
- Up to **2 watermarks** with position, opacity, font, size, and color
- **Page numbers** with format text (e.g. `- {n} / {total} -`)
  - *(Current build)* page numbers are rendered in **black, fixed 10.5pt**
- **Preview**: generate a preview PDF using the **first page** before running the full batch

### Naming Rules & Presets
- Flexible output naming templates using tags (see below)
- Save/load presets (JSON config)

### Security & Optimization
- Set a PDF password in batch
- Remove metadata
- Compress PDF (when supported by your pypdf version)

### Usability
- Drag & drop file registration
- Reorder items, remove, clear list
- Output options:
  - Same folder as source, or custom folder
- Post-process actions:
  - Open PDF / open output folder / clear list

---

## 🏷️ Available Tags

Tags can be used in **Naming Rules** and **Watermark Text**.  
They are automatically replaced during processing.

- `{name}`: original filename (without extension)
- `{sheet}`: Excel sheet name (empty for non-Excel)
- `{parent}`: parent folder name
- `{seq}`: global sequence number
- `{fseq}`: sequence number per input file
- `{pseq}`: page sequence number (when splitting/pages are counted)
- `{total}`: total number of input files
- `{ptotal}`: total pages in the current file/unit
- `{username}`: Windows login username
- `{rand}`: random 4-digit number (e.g. `0427`)

### Date/Time Tag
- `{date:yyyy-mm-dd}` → `2026-02-06`
- `{date:yyyy/mm/dd HH:MM}` → `2026/02/06 16:54`

Supported placeholders:
- `yyyy` year / `mm` month / `dd` day
- `HH` hour / `MM` minute / `SS` second

---

## 🚀 How to Use

### Requirements
- Windows 10 / 11 (64-bit)
- **Microsoft Office installed** (Excel / Word / PowerPoint)

### Basic Steps
1. Launch the app.
2. Drag & drop files into the list.
3. (Optional) Double-click an item to set page range / sheet selection.
4. Configure:
   - Watermark(s) / page numbers
   - Split/merge options
   - Naming rules and output folder
5. Click **Start Conversion**.

---

## 📦 Install

1. Download the latest release from **Releases**.
2. Extract `Office2PDF.zip`.
3. Run `Office2PDF.exe` (no installer required).

> If you do not have a release page yet, create one on GitHub:  
> `Releases` → `Draft a new release`

---

## 🛠️ Run from Source / Build Executable

### 1) Environment
- Python **3.9+** (Windows)
- Microsoft Office installed

### 2) Install dependencies
```bash
pip install pywin32 pypdf Pillow reportlab tkinterdnd2
````

### 3) Run

```bash
python Office2PDF.py
```

### 4) Build `.exe` (PyInstaller)

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --name Office2PDF Office2PDF.py
```

After build, the executable will be available in:

*   `dist/Office2PDF.exe`

> Tip: Some environments require keeping PowerPoint visible during conversion.  
> If PowerPoint conversion fails, avoid forcing `Visible=False`.

***

## ☕ Support the Project

If you find this software useful, support would be greatly appreciated.

*   Amazon eGift Card (Email delivery):  
    <https://www.amazon.co.jp/Amazon-eGift-Card-Flower-Bouquet-Animated/dp/B004N3APDM/>
*   Recipient email: `the.desert.of.the.real@gmail.com`

***

## ⚖️ License

MIT License — see ./LICENSE
