from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from openpyxl import load_workbook
import re

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# === Your processing logic (UNMODIFIED) ===
def process_excel(input_path, output_path):
    wb = load_workbook(input_path)
    ws = wb["STC Summary"]

    headers = [
        "Source Row",
        "1 – Active Coverage (In-network)", "1 – Active Coverage (Out-of-network)", "1 – Active Coverage (Network Not Applicable)", "1 – Active Coverage (Unknown)",
        "A – Coinsurance (In-network)", "A – Coinsurance (Out-of-network)", "A – Coinsurance (Network Not Applicable)", "A – Coinsurance (Unknown)",
        "B – Copayment (In-network)", "B – Copayment (Out-of-network)", "B – Copayment (Network Not Applicable)", "B – Copayment (Unknown)",
        "C – Deductible (In-network)", "C – Deductible (Out-of-network)", "C – Deductible (Network Not Applicable)", "C – Deductible (Unknown)",
        "F – Limitation (In-network)", "F – Limitation (Out-of-network)", "F – Limitation (Network Not Applicable)", "F – Limitation (Unknown)",
        "G – Out-of-Pocket (In-network)", "G – Out-of-Pocket (Out-of-network)", "G – Out-of-Pocket (Network Not Applicable)", "G – Out-of-Pocket (Unknown)",
        "I – Non-covered (In-network)", "I – Non-covered (Out-of-network)", "I – Non-covered (Network Not Applicable)", "I – Non-covered (Unknown)"
    ]

    for col_idx, header in enumerate(headers, start=2):
        ws.cell(row=1, column=col_idx, value=header)

    def sort_stcs(stcs):
        nums, alphas = [], []
        for s in stcs:
            s = s.strip()
            if s.isdigit():
                nums.append(int(s))
            elif s:
                alphas.append(s.upper())
        nums_sorted = [str(n) for n in sorted(nums)]
        alphas_sorted = sorted(set(alphas))
        return nums_sorted + alphas_sorted

    def parse_271(text):
        if not text or not str(text).strip():
            return None
        segments = [seg.strip() for seg in str(text).split("~") if seg.strip()]
        benefit_types = ['1', 'A', 'B', 'C', 'F', 'G', 'I']
        categories = ['Y', 'N', 'W', 'UNKNOWN']
        result = {bt: {cat: [] for cat in categories} for bt in benefit_types}

        for seg in segments:
            if not seg.startswith("EB*"):
                continue
            parts = seg.split("*")
            while len(parts) < 13:
                parts.append("")
            eb01 = parts[1].strip()
            eb03 = parts[3].strip()
            eb12 = parts[12].strip().upper()

            if eb12 in ['Y', 'N', 'W']:
                category = eb12
            elif eb12 in ['', 'U']:
                category = 'UNKNOWN'
            else:
                category = 'UNKNOWN'

            if eb01 in result:
                stcs = eb03.split("^") if "^" in eb03 else ([eb03] if eb03 else [])
                result[eb01][category].extend(stcs)

        for bt in result:
            for cat in result[bt]:
                unique_stcs = []
                seen = set()
                for s in result[bt][cat]:
                    if s and s.upper() not in seen:
                        seen.add(s.upper())
                        unique_stcs.append(s)
                sorted_stcs = sort_stcs(unique_stcs)
                result[bt][cat] = ",".join(sorted_stcs) if sorted_stcs else "-"

        return result

    max_row = ws.max_row
    for row in range(2, max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        if not cell_value or not str(cell_value).strip():
            continue
        parsed = parse_271(cell_value)
        if not parsed:
            continue

        ws.cell(row=row, column=2, value=f"A{row}")

        col_offset = 3
        for bt in ['1', 'A', 'B', 'C', 'F', 'G', 'I']:
            for cat in ['Y', 'N', 'W', 'UNKNOWN']:
                ws.cell(row=row, column=col_offset, value=parsed[bt][cat])
                col_offset += 1

    for col in range(2, len(headers) + 2):
        # Added a safety check for empty columns to prevent max() error
        vals = [len(str(ws.cell(row=r, column=col).value)) for r in range(1, max_row + 1)]
        max_length = max(vals) if vals else 10
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = min(max_length + 2, 60)

    wb.save(output_path)

# === Routes ===

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return redirect(url_for('index'))
    file = request.files["file"]
    if file.filename == '':
        return redirect(url_for('index'))

    if file.filename.endswith(".xlsx"):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        return redirect(url_for("process_file", filename=file.filename))
    return "Invalid file type. Please upload an Excel file."

@app.route("/process/<filename>")
def process_file(filename):
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_filename = f"processed_{filename}"
    output_path = os.path.join(PROCESSED_FOLDER, output_filename)

    # Execute your existing logic
    process_excel(input_path, output_path)

    # Redirect to the new designed template
    return render_template("download.html", filename=output_filename)

@app.route("/download/<filename>")
def download_file(filename):
    filepath = os.path.join(PROCESSED_FOLDER, filename)
    return send_file(filepath, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
