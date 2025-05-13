import openpyxl
from openpyxl.styles import PatternFill
import os
import glob

# Step 1: Clean Gusto
def normalize_and_clean_gusto(wb):
    if "Gusto Report" not in wb.sheetnames:
        print("Sheet 'Gusto Report' not found.")
        return

    ws = wb["Gusto Report"]

    if "Cleaned Gusto" in wb.sheetnames:
        del wb["Cleaned Gusto"]
    ws_out = wb.create_sheet("Cleaned Gusto")

    headers = ["Clinician Name", "Date", "Job title", "Job hours"]
    ws_out.append(headers)

    row = 1
    max_row = ws.max_row
    out_row = 2

    while row <= max_row:
        cell_value = str(ws.cell(row=row, column=1).value)
        if cell_value.startswith("Hours for "):
            clinician_name = cell_value.replace("Hours for ", "").strip()
            if "," in clinician_name:
                last, first = map(str.strip, clinician_name.split(","))
                clinician_name = f"{first} {last}"

            header_row = [str(ws.cell(row=row + 1, column=col).value).strip() for col in range(1, ws.max_column + 1)]
            j = row + 2
            while j <= max_row:
                entry = str(ws.cell(j, 1).value)
                if entry.startswith("Hours for ") or entry == "":
                    break

                data = {header_row[col - 1]: ws.cell(j, col).value for col in range(1, len(header_row) + 1)}
                base_date = data.get("Date")

                for num in range(1, 5):
                    title = data.get(f"Job {num} title")
                    hours = data.get(f"Job {num} hours")
                    if title or hours:
                        ws_out.cell(row=out_row, column=1).value = clinician_name
                        ws_out.cell(row=out_row, column=2).value = base_date
                        ws_out.cell(row=out_row, column=2).number_format = 'mm/dd/yyyy'
                        ws_out.cell(row=out_row, column=3).value = title
                        ws_out.cell(row=out_row, column=4).value = hours
                        out_row += 1
                j += 1
            row = j - 1
        row += 1

# Step 2: Compare to CR Report
def get_mapped_service_codes(job_title):
    mappings = {
        "ASSISTANT BEHAVIOR CONSULTANT (ABC -ABA)": ["97153"],
        "BEHAVIOR CONSULTANT": ["H0032"],
        "BEHAVIORAL HEALTH TECHNICIAN": ["H2021", "SUB BHT"],
        "BEHAVIORAL HEALTH TECHNICIAN (BHT-ABA)": ["97153", "SUB BHT"],
        "BHT-ABA CENTER": ["97154", "Center Support"],
        "BHT-ABA CENTER GROUP": ["97154", "Center Support"],
        "MOBILE THERAPY": ["H2019"],
        "TRAINING AND SUPERVISION": ["BHT Supervision", "Group Supervision", "Individual Supervision"],
    }
    return mappings.get(job_title.strip().upper(), [""])

def compare_cr_to_gusto(wb):
    if "Cleaned Gusto" not in wb.sheetnames or "CR Report" not in wb.sheetnames:
        print("Required sheets not found.")
        return

    ws_gusto = wb["Cleaned Gusto"]
    ws_cr = wb["CR Report"]

    # Create output sheet
    if "Comparison Report" in wb.sheetnames:
        del wb["Comparison Report"]
    ws_out = wb.create_sheet("Comparison Report")
    headers = [
        "Name", "Date", "Job Title (Gusto)", "Service Code (CR)",
        "Gusto Hours", "CR Hours", "Match", "Note"
    ]
    ws_out.append(headers)

    gusto_data = [
        {
            "name": row[0],
            "date": row[1].strftime('%Y-%m-%d') if row[1] else None,
            "job_title": row[2],
            "hours": row[3]
        }
        for row in ws_gusto.iter_rows(min_row=2, values_only=True)
    ]
    cr_data = [
        {
            "name": row[0],
            "service_code": str(row[1]).strip() if row[1] is not None else "",
            "date": row[2].strftime('%Y-%m-%d') if row[2] else None,
            "hours": row[3]
        }
        for row in ws_cr.iter_rows(min_row=2, values_only=True)
    ]

    TOLERANCE = 0.25
    used_cr = [False] * len(cr_data)
    output_row = 2
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for g in gusto_data:
        mapped_codes = get_mapped_service_codes(g["job_title"] or "")
        found_match = False
        for i, c in enumerate(cr_data):
            if used_cr[i]:
                continue
            if (g["name"] == c["name"] and g["date"] == c["date"] and c["service_code"] in mapped_codes):
                ws_out.append([
                    g["name"], g["date"], g["job_title"], c["service_code"],
                    g["hours"], c["hours"],
                    "Yes" if abs((g["hours"] or 0) - (c["hours"] or 0)) <= TOLERANCE else "No",
                    "Matched" if abs((g["hours"] or 0) - (c["hours"] or 0)) <= TOLERANCE else "Hour mismatch"
                ])
                if abs((g["hours"] or 0) - (c["hours"] or 0)) > TOLERANCE:
                    for col in range(1, 9):
                        ws_out.cell(row=output_row, column=col).fill = red_fill
                output_row += 1
                used_cr[i] = True
                found_match = True
                break
        if not found_match:
            ws_out.append([
                g["name"], g["date"], g["job_title"], "", g["hours"], "", "No", "No matching CR entry"
            ])
            for col in range(1, 9):
                ws_out.cell(row=output_row, column=col).fill = red_fill
            output_row += 1

    for i, c in enumerate(cr_data):
        if not used_cr[i]:
            ws_out.append([
                c["name"], c["date"], "", c["service_code"], "", c["hours"], "No", "No matching Gusto entry"
            ])
            for col in range(1, 9):
                ws_out.cell(row=output_row, column=col).fill = red_fill
            output_row += 1

    print("✅ Comparison complete.")

# MAIN
if __name__ == "__main__":
    files = glob.glob("*_Payroll.xlsx")
    if not files:
        print("No payroll files found.")
    else:
        latest_file = max(files, key=os.path.getctime)
        print(f"📂 Processing latest payroll file: {latest_file}")
        wb = openpyxl.load_workbook(latest_file)
        normalize_and_clean_gusto(wb)
        compare_cr_to_gusto(wb)
        wb.save(latest_file)
        print(f"✅ All results saved to '{latest_file}'.")
