# High School Graduation Exam Scheduler

This software is designed to generate an optimal daily schedule for a graduation oral exam committee and the students. The program minimizes the idle times of the exam committee and ensures a fair and minimal waiting time for students in the room.

The project offers both a **Python CLI (Command Line Interface)** version and an easy-to-use, interactive **browser-based web interface**.

---

## Key Features

- **Intelligent Optimization**: By interleaving language subjects (0 minutes preparation time) as "slot fillers", the program keeps the committee active while other students are preparing for non-language subjects.
- **Dual Interface**:
  - Command Line Python version ([vizsga_beosztas.py](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/vizsga_beosztas.py)) with Excel export capability.
  - Web application ([index.html](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/index.html)) featuring manual data entry, Excel import, and interactive visualizations.
- **Automatic Multi-Day Splitting**: Based on the specified daily time limit, the software automatically splits the examinees into separate days.
- **Entry Time Alignment**: Students do not wait in the room from the start of the day; instead, they enter exactly at the beginning of their preparation preceding their first exam.
- **Excel Compatibility**: Reads examinees based on a template and exports detailed daily schedules (with separate sheets for the committee and the students).

---

## Examination Rules and Constraints

- **Exam Duration**: 15 minutes per student.
- **Break**: 2 minutes of mandatory break for the committee between exams.
- **Preparation Time**: 
  - Non-language subjects: **minimum 20 minutes**.
  - Language subjects: **0 minutes** (the exam can start immediately).
- **Classroom Capacity**: A maximum of **5 students** can stay in the exam room simultaneously.
- **Presence**: A student must remain in the room from the start of preparation for their first exam until the end of their last exam.

---

## User Guide

### 1. Browser Version (`index.html`)
No special installation or server required.
1. Open the [index.html](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/index.html) file in any modern web browser.
2. Enter the parameters (start/end times, preparation time, etc.).
3. Load test data, import an Excel file, or manually add rows.
4. Click the **"Beosztás készítése"** (Generate Schedule) button.
5. View the results on the *Statisztika* (Statistics), *Bizottság beosztása* (Committee Schedule), and *Diákok beosztása* (Students Schedule) tabs, and download the generated Excel sheet.

*(Note: The web interface automatically validates parameters; if the specified timeframe is too small to conduct even a single exam, the calculation will not start, avoiding browser freezing.)*

### 2. Python Version (`vizsga_beosztas.py`)
Requires Python 3 and the `openpyxl` library (used for Excel file management).

#### Installation:
```bash
pip install openpyxl
```

#### Running the script:
```bash
python vizsga_beosztas.py
```
During execution, the program prompts for:
- The start and end of exams (by hour, e.g., 8 and 16).
- The name of the input Excel file (default: `vizsga_bemenet_sablon.xlsx`).

After completion, the detailed schedule is displayed in the console, and a `vizsga_beosztas.xlsx` file is generated.

---

## Input Excel Format

The Excel file used for importing must follow this structure:
- **Column A** contains the name of the examinee.
- **Starting from Column B** (number of columns is arbitrary), the examinee's subjects are listed.
- **Marking Language Subjects**: The name of language subjects must be prefixed with an **asterisk (`*`)** character (e.g., `*Angol nyelv`, `*Német nyelv`).

The sample template can be found in the project directory: [vizsga_bemenet_sablon.xlsx](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/vizsga_bemenet_sablon.xlsx).

---

## Additional Documentation

For detailed information about the internal workings, simulation model, and optimization algorithms of the software, please refer to the [ARCHITECTURE.md](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/ARCHITECTURE.md) document.
A magyar nyelvű leírásért lásd a [README.md](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/README.md) fájlt.
