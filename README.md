# Texas Power Guru

A daily-updated electricity plan comparison tool for Texas.

## How it works

Your existing Python script updates the Excel file daily.
Add one call to `build_from_excel.py` at the end of your script and
it rebuilds `index.html` automatically with the fresh data.

## File structure

```
texas_power_guru/
├── build_from_excel.py   ← add a call to this in your daily script
├── template.html         ← app shell (do not edit)
├── index.html            ← rebuilt each day (upload this to GitHub)
├── requirements.txt
└── README.md
```

## Setup

1. Install dependencies (one time):
       pip install pandas openpyxl

2. Open build_from_excel.py and set EXCEL_PATH to wherever your
   script saves the Excel file.

3. Add this to the END of your existing daily Python script:

       import subprocess
       subprocess.run(["python", r"C:\path\to\build_from_excel.py"])

   Or import and call directly:

       from build_from_excel import build_html
       build_html(excel_path=r"C:\path\to\your_file.xlsx")

4. Each day after your script runs, upload the new index.html to
   your GitHub repository. GitHub Pages will serve the updated app.

## Running manually in Spyder

Open build_from_excel.py and press F5, or run from the console:

    runfile('build_from_excel.py', args='C:/path/to/your_file.xlsx')
