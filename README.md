# Attendance Sheet Maker 📋

> **Built for FACE Club — REVA University, Bengaluru**

A streamlined web tool designed for **FACE Club** (Forum of Aspiring Computer Engineers) at **REVA University** to automatically generate standardized, multi-group attendance workbooks (`.xlsx`) for badging events, hackathons, club meetings, and workshops.

🌐 **Live Application**: [https://zynssam.github.io/Attendance-Sheet-Maker/](https://zynssam.github.io/Attendance-Sheet-Maker/)

---

## 🎯 Purpose

During university club activities and badging sessions, attendance sheets need to be submitted across multiple groups and semesters with strict institutional formatting:
- Separate sheets for **3rd Sem** and **5th Sem**.
- Proper student details: **Sl. No**, **Class Section (A–F)**, **Student Name**, and **USN / SRN**.
- Dedicated faculty coordinator blocks for **Group 1**, **Group 2**, and **Group 3**.
- Preserved signature footer spacing for the **Extracurricular Vertical Head** and **School Director / HoDs**.

This tool eliminates manual data entry and formatting errors by taking raw USN lists and instantly producing official, ready-to-sign Excel files.

---

## ✨ Features

- **⚡ Instant 100% Client-Side Processing**: No servers or database dependencies at runtime. Everything runs right in your browser with zero latency.
- **🔍 Smart USN Matching**: Paste raw student USNs (space, comma, or newline-separated). The app automatically resolves student names, branch, section, and semester from the FACE Club registry.
- **📑 Grouped Excel Templates**: Inlines and preserves authentic REVA University workbook styles, column widths, merged headers, Link Sheets, and formulas.
- **✍️ Signature Clearance**: Automatically positions the signature footer row with a clean 4-row spacing after the last student entry.
- **🌗 Responsive UI**: Clean, minimalistic interface with automatic light/dark mode support.
- **🚀 Automated CI/CD**: Automatically rebuilds and deploys to GitHub Pages upon git push.

---

## 👥 Coordinators & Groups

| Group | Faculty Coordinator | Applicable Semesters |
| :--- | :--- | :--- |
| **Group 1** | Prof. Mala B M | 3rd Sem, 5th Sem, Link Sheet |
| **Group 2** | Prof. Ravi B K | 3rd Sem, 5th Sem, Link Sheet |
| **Group 3** | Prof. Manasa K N | 3rd Sem, 5th Sem, Link Sheet |

---

## 🛠️ Project Structure

```text
├── data/                                # Master records & Excel templates
│   ├── Facemembers_2026-27_Grouped.xlsx # Registered FACE Club member list
│   ├── Attendance_21-9_Group1_Badging.xlsx
│   ├── Attendance_21-9_Group2_Badging.xlsx
│   └── Attendance_21-9_Group3_Badging.xlsx
├── .github/workflows/deploy.yml         # GitHub Actions automated deployment
├── build.py                             # Inlines templates/members into static HTML
├── template.html                        # App UI & client-side XML workbook generator
├── index.html                           # Compiled static SPA entrypoint
├── attendance_app.html                  # Standalone offline single-file app
├── requirements.txt                     # Build dependencies
├── test_build.py                        # Automated build & data integrity tests
└── test_browser.py                      # Headless workbook & XML structure tests
```

---

## 💻 Local Development

### 1. Clone & Install Dependencies
```bash
git clone https://github.com/zynssam/Attendance-Sheet-Maker.git
cd Attendance-Sheet-Maker
pip install -r requirements.txt
```

### 2. Updating Data or UI
When updating the member registry in `data/Facemembers_2026-27_Grouped.xlsx` or tweaking styles in `template.html`, rebuild the static distribution:
```bash
python3 build.py
```
This generates updated `index.html` and `attendance_app.html` files.

### 3. Run Tests
```bash
python3 test_build.py
python3 test_browser.py
```

---

## 🚀 Deployment

The project is hosted on **GitHub Pages**. Pushing any change to the `main` branch automatically triggers the `.github/workflows/deploy.yml` pipeline, which:
1. Installs Python dependencies.
2. Executes `build.py` to compile the app.
3. Deploys the static site to [https://zynssam.github.io/Attendance-Sheet-Maker/](https://zynssam.github.io/Attendance-Sheet-Maker/).

---

## 🏫 Credits

Developed for **FACE Club (Forum of Aspiring Computer Engineers)**  
**School of Computer Science & Engineering**, REVA University, Bengaluru.
