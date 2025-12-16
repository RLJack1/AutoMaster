import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import random
import math
from openpyxl import load_workbook

COUNTRY_LOCATION_MAP = {
    "Belgium": [
         "PH1_BE_ING_Brussels", "PH2_BE_Brussels"
    ],
    "Philippines": [
        "PH1_PH_World_Plaza", "PH2_PH_One_Ayala_Tower", "PH2_PH_Manila"
    ],
    "Romania": [
         "PH1_RO_Bucharest", "PH1_RO_Vladimir"
    ],
    "Great Britain": [
         "PH3_GB_London"
    ],
    "Slovakia": [
        "PH3_SK_Bratislava"
    ],
    "United States": [
        "PH1_US_Washington"
    ],
    "Germany": [
        "PH3_DE_Ulm", "PH2_DE_Stuttgart"
    ],
    "China": [
        "PH2_CN_Shanghai"
    ],
    "Ukraine": [
        "PH2_UA_Kyiv"
    ],
    "Poland": [
        "PH1_PL_Katowice"
    ],
    "Czechia": [
        "PH2_CZ_Skaltiz"
    ],
    "Hungary": [
        "PH2_HU_Budapest"
    ]
}

# ---------------- HELPER FUNCTIONS ---------------- 

def normalize_raw_name(raw_name):
    
    # Converts 'First Last - ID' -> 'Last, First'
    
    if pd.isna(raw_name):
        return None

    raw_name = str(raw_name)

    if "-" not in raw_name:
        return raw_name.strip()

    try:
        name_part = raw_name.split("-")[0].strip()
        first, last = name_part.split(" ", 1)
        return f"{last}, {first}"
    except:
        return raw_name.strip()
    
def format_assignment_group(group):
    if pd.isna(group):
        return ""
    
    # Switch-case style normalization for Assignment Group
    
    group = str(group).lower()

    if "compensation" in group or "benefits" in group:
        return "Performance & Rewards"
    elif "expenses" in group:
        return "Expense"
    elif "hcm" in group:
        return "Human Capital Management"
    elif "organizational management" in group:
        return "Org Management"
    elif "international mobility" in group:
        return "International Mobility"
    elif "learning coordination" in group:
        return "Learning"
    elif "people services" in group:
        return "Contact Center"
    elif "recruitment admin" in group:
        return "Recruitment Admin"
    elif "reporting" in group:
        return "Reporting"
    elif "travels" in group:
        return "Travel"
    else:
        return group.title()


def calculate_sample_percentage(tenure):
    if tenure < 6:
        return 0.15
    elif tenure < 12:
        return 0.10
    else:
        return 0.05
    
def format_location(raw_location):
    
    # Maps raw location values to standardized country names
    # based on COUNTRY_LOCATION_MAP
    
    if pd.isna(raw_location):
        return ""

    raw_location = str(raw_location).upper()

    for country, identifiers in COUNTRY_LOCATION_MAP.items():
        for identifier in identifiers:
            if identifier.upper() in raw_location:
                return country

    return raw_location.title()

def find_column(df, possible_names):
    
    # Finds a column in df whose header matches one of the possible_names.
    # Matching is case-insensitive.
    
    for col in df.columns:
        col_name = str(col).strip().lower()
        for name in possible_names:
            if col_name == name.strip().lower():
                return col
    raise ValueError(f"Missing required column: {possible_names}")

def find_qcl_start_row(ws):
    
    # Finds the first empty row below 'Control Check No.' in column B
    
    for row in range(1, ws.max_row + 1):
        cell_value = ws[f"B{row}"].value
        if cell_value and str(cell_value).strip().lower() == "control check no.":
            # Start checking below header
            r = row + 1
            while ws[f"B{r}"].value:
                r += 1
            return r
    raise ValueError("Could not find 'Control Check No.' in QCL template.")

def detect_header_row(file_path, sheet_name=0, required_columns=None):
    
    # Detects the row number containing required column headers.
    # Returns row index to be used as header=.
    
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    required_columns = [col.lower() for col in required_columns]

    for idx, row in preview.iterrows():
        row_values = [str(cell).strip().lower() for cell in row.values if pd.notna(cell)]
        if all(req in row_values for req in required_columns):
            return idx

    raise ValueError(f"Could not find header row containing: {required_columns}")



# ---------------- MAIN PROCESS ---------------- #

def process_files():
    raw_file = raw_entry.get()
    member_file = member_entry.get()
    qcl_template = qcl_entry.get()

    if not raw_file or not member_file or not qcl_template:
        messagebox.showerror("Error", "Please select all three Excel files.")
        return

    try:
        raw_df = pd.read_excel(raw_file)

        location_col = find_column(raw_df, ["Location"])
        assigned_col = find_column(raw_df, ["Assigned to", "Assigned To"])
        case_id_col = find_column(raw_df, ["Number"])
        service_col = find_column(raw_df, ["HR Service"])
        assignment_group_col = find_column(raw_df, ["Assignment group"])

        member_header_row = detect_header_row(
            member_file,
            required_columns=["Team", "Tenure"]
        )

        members_df = pd.read_excel(
            member_file,
            header=member_header_row
        )
        
        member_name_col = find_column(members_df, ["Team"])
        tenure_col = find_column(members_df, ["Tenure"])

        employee_cases = {}

        for _, row in raw_df.iterrows():
            assigned_to = str(row[assigned_col]).strip()
            normalized = normalize_raw_name(assigned_to)
            employee_cases.setdefault(normalized, []).append(row)

        wb = load_workbook(qcl_template)
        ws = wb["QA Checks"]

        current_row = find_qcl_start_row(ws)
        control_no = 1

        output_text.delete("1.0", tk.END)

        for _, member in members_df.iterrows():
            member_name = str(member[member_name_col]).strip()
            tenure = member[tenure_col]

            if member_name not in employee_cases or pd.isna(tenure):
                continue

            cases = employee_cases[member_name]
            sample_size = max(1, math.ceil(len(cases) * calculate_sample_percentage(tenure)))
            sampled_cases = random.sample(cases, min(sample_size, len(cases)))

            output_text.insert(
                tk.END,
                f"{member_name} | Sampled {len(sampled_cases)} cases\n"
            )

            for case in sampled_cases:
                ws[f"B{current_row}"] = control_no
                ws[f"C{current_row}"] = format_assignment_group(case[assignment_group_col])
                ws[f"D{current_row}"] = format_location(case[location_col])
                ws[f"E{current_row}"] = normalize_raw_name(case[assigned_col])
                ws[f"F{current_row}"] = case[case_id_col]
                ws[f"G{current_row}"] = case[service_col]

                control_no += 1
                current_row += 1

        output_path = qcl_template.replace(".xlsx", "_POPULATED.xlsx")
        wb.save(output_path)

        messagebox.showinfo("Success", f"QCL generated:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# ---------------- GUI SETUP ---------------- #

root = tk.Tk()
root.title("HR QA Case Sampling & QCL Generator")
root.geometry("820x700")

# Circle Dropdown (No logic yet)
tk.Label(root, text="Circle Selection:").pack(pady=5)
circle_var = tk.StringVar()
circle_dropdown = tk.OptionMenu(
    root,
    circle_var,
    "Circle 1 (HCM, OM)",
    "Circle 2 (Expense, Reporting, Travel)",
    "Circle 3 (Learning)",
    "Circle 4 (International Mobility)",
    "Circle 5 (Performance & Rewards)",
    "Circle 6 (Contact Center)",
    "Circle 7 (Recruitment Admin)"
)
circle_dropdown.pack()
circle_var.set("Circle 1 (HCM, OM)")

def file_input(label):
    tk.Label(root, text=label).pack(pady=4)
    frame = tk.Frame(root)
    frame.pack()
    entry = tk.Entry(frame, width=65)
    entry.pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Browse", command=lambda: browse(entry)).pack(side=tk.LEFT)
    return entry

def browse(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

raw_entry = file_input("Raw HR Case File:")
member_entry = file_input("Member List File:")
qcl_entry = file_input("QCL Template File:")

tk.Button(
    root,
    text="Process Files & Generate QCL",
    command=process_files,
    width=35
).pack(pady=10)

output_text = scrolledtext.ScrolledText(root, wrap=tk.WORD)
output_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

root.mainloop()