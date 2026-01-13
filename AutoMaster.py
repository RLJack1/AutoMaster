import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import random
import math
from openpyxl import load_workbook
from PIL import Image, ImageTk
import os
import sys
import warnings
import shutil

# Suppress openpyxl warnings for .xlsm files
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Dictionary containing the country codes for all available countries serviced by CPS
COUNTRY_LOCATION_MAP = {
    "Australia": {
        "formatted_prefixes": [],
        "arbitrary_codes": [
            "SYD", "TUG", "WYN", "AU", "OVR", 
        ]
    },
    "Belgium": {
        "formatted_prefixes": ["PH3_BE_ING_BANK", "PH3_BE"],
        "arbitrary_codes": [
            "BE", "M1 03 14"
        ]
    },
    "China": {
        "formatted_prefixes": ["PH2_CN"],
        "arbitrary_codes": [
            
        ]
    },
    "Czech Republic": {
        "formatted_prefixes": ["PH2_CZ"],
        "arbitrary_codes": [
            
        ]
    },
    "France": {
        "formatted_prefixes": [],
        "arbitrary_codes": [
            "France", "RIC"
        ]
    },
    "Germany": {
        "formatted_prefixes": ["PH2_DE"],
        "arbitrary_codes": [
            "Germany"
        ]
    },
    "Hong Kong": {
        "formatted_prefixes": ["PH2_HK"],
        "arbitrary_codes": [
            "Hong Kong"
        ]
    },
    "Hungary": {
        "formatted_prefixes": ["PH2_HU"],
        "arbitrary_codes": [
            "HU"
        ]
    },
    "Ireland": {
        "formatted_prefixes": ["PH2_IE"],
        "arbitrary_codes": [
            
        ]
    },
    "Italy": {
        "formatted_prefixes": ["PH2_ITA"],
        "arbitrary_codes": [
            "Italy"
        ]
    },
    "Japan": {
        "formatted_prefixes": ["PH2_JP"],
        "arbitrary_codes": [
        ]
    },
    "Luxembourg": {
        "formatted_prefixes": ["PH2_LU"],
        "arbitrary_codes": [
            "LU"
        ]
    },
    "Netherlands": {
        "formatted_prefixes": ["Guest NL"],
        "arbitrary_codes": [
            "ACT", "ALP", "AME", "BMG", 
            "CDR", "DP", "HBP", "KBK",
            "KCC", "KCM", "KFL", "KFR", 
            "KGD", "KMH", "KNF", "KQB",
            "KQB", "KSY", "KXS", "KZK", 
            "KZR", "LZ", "RBG", "WBM",
            "WP", "NL", "AT", "IE", 
            "RO", "KNZ", "KBE", "Guest NL - Guest.NL",
            "KLQ", "KQT", "DEA"
        ]
    },
    "Philippines": {
        "formatted_prefixes": ["PH1_PH", "Guest PH", "PH2_PH"],
        "arbitrary_codes": [
            "PH", "Guest PH - Guest.PH"
        ]
    },
    "Poland": {
        "formatted_prefixes": ["PH1_PL"],
        "arbitrary_codes": [
            "PL", "PH2_SK_Brati", "P-PULAWSKA", "P-CHORZOWSKA", "P-CHORZ.50", "DR01R1402", "DR09R0102"
        ]
    },
    "Romania": {
        "formatted_prefixes": ["PH1_RO"],
        "arbitrary_codes": [
            "PH", "Guest PH - Guest.PH"
        ]
    },
    "Russian Federation": {
        "formatted_prefixes": [],
        "arbitrary_codes": [
            "RUSMCW001"
        ]
    },
    "Singapore": {
        "formatted_prefixes": ["PH2_SG"],
        "arbitrary_codes": [
            "SG", "Guest Asia - Guest.Asia"
        ]
    },
    "Slovakia": {
        "formatted_prefixes": ["PH1_SK","PH2_SK"],
        "arbitrary_codes": [
            "SK"
        ]
    },
    "South Korea": {
        "formatted_prefixes": ["PH2_KR"],
        "arbitrary_codes": [
        ]
    },
    "Spain": {
        "formatted_prefixes": [],
        "arbitrary_codes": [
            "Spain", "Madrid_Pobla", "Madrid_Poblados"
        ]
    },
    "Sri Lanka": {
        "formatted_prefixes": ["PH2_LK"],
        "arbitrary_codes": [
        ]
    },
    "Switzerland": {
        "formatted_prefixes": ["PH2_CH"],
        "arbitrary_codes": [
        ]
    },
    "Taiwan": {
        "formatted_prefixes": ["PH2_TW"],
        "arbitrary_codes": [
            "TW"
        ]
    },
    "Turkey": {
        "formatted_prefixes": ["PH2_TR"],
        "arbitrary_codes": [
            "TR", "4127"
        ]
    },
    "Ukraine": {
        "formatted_prefixes": ["PH2_UA"],
        "arbitrary_codes": [
        ]
    },
    "United Kingdom": {
        "formatted_prefixes": ["PH2_GB"],
        "arbitrary_codes": [
        ]
    },
    "United States": {
        "formatted_prefixes": ["PH2_US"],
        "arbitrary_codes": [
        ]
    },
}

# ---------------- HELPER FUNCTIONS ---------------- 

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_image(image_path, width=None, height=None):
    """Load and resize image for tkinter"""
    try:
        img = Image.open(get_resource_path(image_path))
        if width and height:
            img = img.resize((width, height), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception as e:
        print(f"Error loading image {image_path}: {e}")
        return None

# ---------------- HELPER FUNCTIONS ---------------- 

def normalize_raw_name(raw_name):
    """
    Converts various name formats to 'Last, First'
    Handles: 'First Last - ID', 'First Last', 'Last, First'
    """
    if pd.isna(raw_name):
        return None
    
    raw_name = str(raw_name).strip()
    
    # Handle empty strings
    if not raw_name:
        return None
    
    # Remove ID if present (e.g., "First Last - ID123")
    if "-" in raw_name:
        raw_name = raw_name.split("-")[0].strip()
    
    # If already in "Last, First" format, return as-is
    if "," in raw_name:
        return raw_name.strip()
    
    # Convert "First Last" to "Last, First"
    parts = raw_name.split()
    if len(parts) >= 2:
        first = parts[0]
        last = " ".join(parts[1:])  # Handle multi-part last names
        return f"{last}, {first}"
    
    # If single name, return as-is
    return raw_name

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
    
def extract_formatted_prefix(location_code):
    """
    Extracts Department-Country prefix from formatted codes.
    Example: PH1_BE_Brussels -> PH1_BE
    """
    parts = location_code.split("_")
    if len(parts) >= 2:
        return f"{parts[0]}_{parts[1]}"
    return location_code


def format_location(raw_location):
    if pd.isna(raw_location):
        return ""

    raw_location = str(raw_location).upper().strip()

    # Extract formatted prefix if applicable
    prefix = extract_formatted_prefix(raw_location)

    for country, rules in COUNTRY_LOCATION_MAP.items():
        # Check formatted prefixes
        if prefix in rules.get("formatted_prefixes", []):
            return country

        # Check arbitrary/manual codes
        for code in rules.get("arbitrary_codes", []):
            if code.upper() in raw_location:
                return country

    # Fallback if no match
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
    raw_file = qa_checks_entry.get()
    qa_review_file = qa_review_entry.get()
    qcl_template = qcl_entry.get()
    member_file = member_entry.get()

    if not raw_file or not member_file or not qcl_template:
        messagebox.showerror("Error", "Please select all required files.")
        return

    try:
        # Read raw data file
        raw_df = pd.read_excel(raw_file)
        location_col = find_column(raw_df, ["Location"])
        assigned_col = find_column(raw_df, ["Assigned to", "Assigned To"])
        case_id_col = find_column(raw_df, ["Number"])
        service_col = find_column(raw_df, ["HR Service"])
        assignment_group_col = find_column(raw_df, ["Assignment group"])

        # Read member list file
        member_header_row = detect_header_row(
            member_file,
            required_columns=["Team", "Tenure"]
        )
        members_df = pd.read_excel(member_file, header=member_header_row)
        member_name_col = find_column(members_df, ["Team"])
        tenure_col = find_column(members_df, ["Tenure"])

        # Build dictionary of cases per employee
        # KEY FIX: Normalize names from raw data
        employee_cases = {}
        for _, row in raw_df.iterrows():
            assigned_to = str(row[assigned_col]).strip()
            normalized = normalize_raw_name(assigned_to)
            if normalized:  # Only add if normalization succeeded
                employee_cases.setdefault(normalized, []).append(row)

        # DEBUGGING: Print matching statistics
        print(f"\n=== MATCHING STATISTICS ===")
        print(f"Total cases in raw file: {len(raw_df)}")
        print(f"Unique agents with cases: {len(employee_cases)}")
        print(f"Total agents in member list: {len(members_df)}")
        
        # Copy template and preserve all features
        file_ext = os.path.splitext(qcl_template)[1]
        output_path = qcl_template.replace(file_ext, f"_POPULATED{file_ext}")
        shutil.copy2(qcl_template, output_path)
        
        wb = load_workbook(
            output_path, 
            keep_vba=True,
            data_only=False,
            keep_links=True
        )
        
        ws = wb["QA Checks"]
        current_row = find_qcl_start_row(ws)
        control_no = 1
        
        # Track processing statistics
        matched_agents = 0
        total_cases_written = 0
        unmatched_agents = []

        # Process each member
        for _, member in members_df.iterrows():
            member_name_raw = str(member[member_name_col]).strip()
            # KEY FIX: Normalize member names too
            member_name = normalize_raw_name(member_name_raw)
            tenure = member[tenure_col]
            
            # Skip if no valid name or tenure
            if not member_name or pd.isna(tenure):
                continue
            
            # Check if this member has cases
            if member_name not in employee_cases:
                unmatched_agents.append(member_name)
                continue
            
            matched_agents += 1
            cases = employee_cases[member_name]
            
            # Calculate sample size based on tenure
            sample_size = max(1, math.ceil(len(cases) * calculate_sample_percentage(tenure)))
            sampled_cases = random.sample(cases, min(sample_size, len(cases)))
            
            print(f"Agent: {member_name} | Tenure: {tenure} | Total Cases: {len(cases)} | Sample: {len(sampled_cases)}")

            # Write sampled cases to QCL
            for case in sampled_cases:
                ws[f"B{current_row}"].value = control_no
                ws[f"C{current_row}"].value = format_assignment_group(case[assignment_group_col])
                ws[f"D{current_row}"].value = format_location(case[location_col])
                ws[f"E{current_row}"].value = normalize_raw_name(case[assigned_col])
                ws[f"F{current_row}"].value = case[case_id_col]
                ws[f"G{current_row}"].value = case[service_col]
                control_no += 1
                current_row += 1
                total_cases_written += 1

        # Save workbook
        wb.save(output_path)
        wb.close()
        
        # Print final statistics
        print(f"\n=== FINAL RESULTS ===")
        print(f"Matched agents: {matched_agents}")
        print(f"Total cases written: {total_cases_written}")
        print(f"Unmatched agents: {len(unmatched_agents)}")
        if unmatched_agents[:5]:  # Show first 5 unmatched
            print(f"Sample unmatched: {unmatched_agents[:5]}")
        
        messagebox.showinfo(
            "Success", 
            f"QCL generated successfully!\n\n"
            f"Output: {output_path}\n\n"
            f"Matched agents: {matched_agents}\n"
            f"Cases written: {total_cases_written}\n"
            f"Unmatched agents: {len(unmatched_agents)}"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))
        import traceback
        traceback.print_exc()

# ---------------- GUI SETUP ---------------- #

root = tk.Tk()
root.title("AutoMaster")
root.geometry("600x700")
root.minsize(500, 600)  # Set minimum size to prevent elements from overlapping
root.configure(bg="#BFBFBF")

# Set window icon
try:
    icon_image = Image.open(get_resource_path("icon.png"))
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, icon_photo)
except Exception as e:
    print(f"Could not load icon: {e}")

# Create main container with scrollbar for small screens
main_canvas = tk.Canvas(root, bg="#BFBFBF", highlightthickness=0)
scrollbar = tk.Scrollbar(root, orient="vertical", command=main_canvas.yview)
scrollable_frame = tk.Frame(main_canvas, bg="#BFBFBF")

scrollable_frame.bind(
    "<Configure>",
    lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
)

main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
main_canvas.configure(yscrollcommand=scrollbar.set)

# Pack canvas and scrollbar
main_canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Enable mouse wheel scrolling
def on_mousewheel(event):
    main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

main_canvas.bind_all("<MouseWheel>", on_mousewheel)

# Logo frame
logo_frame = tk.Frame(scrollable_frame, bg="white", height=120)
logo_frame.pack(fill=tk.X, padx=20, pady=(20, 15))
logo_frame.pack_propagate(False)

# Try to load logo image, fallback to text if not found
logo_img = load_image("logo.png", width=200, height=80)
if logo_img:
    logo_label = tk.Label(logo_frame, image=logo_img, bg="white")
    logo_label.image = logo_img  # Keep a reference
    logo_label.pack(expand=True)
else:
    # Fallback to text logo
    logo_label = tk.Label(logo_frame, text="A\nM", font=("Arial", 36, "bold"), 
                          bg="white", fg="#FF6B35")
    logo_label.pack(expand=True)

# Circle Selection
circle_label = tk.Label(scrollable_frame, text="1.", font=("Arial", 14, "bold"), 
                        bg="#BFBFBF", fg="black")
circle_label.pack(anchor=tk.W, padx=30, pady=(10, 5))

circle_frame = tk.Frame(scrollable_frame, bg="white", bd=2, relief=tk.SOLID, highlightbackground="#FF6B35", 
                        highlightthickness=2)
circle_frame.pack(fill=tk.X, padx=30, pady=5)

circle_var = tk.StringVar(value="Circle 1 (HCM, OM)")
circle_dropdown = tk.OptionMenu(
    circle_frame,
    circle_var,
    "Circle 1 (HCM, OM)",
    "Circle 2 (Expense, Reporting, Travel)",
    "Circle 3 (Learning)",
    "Circle 4 (International Mobility)",
    "Circle 5 (Performance & Rewards)",
    "Circle 6 (Contact Center)",
    "Circle 7 (Recruitment Admin)"
)
circle_dropdown.config(font=("Arial", 11), bg="white", fg="gray", 
                       relief=tk.FLAT, highlightthickness=0, width=35)
circle_dropdown.pack(fill=tk.X, padx=5, pady=5)

# File upload section
file_section_label = tk.Label(scrollable_frame, text="2.", font=("Arial", 14, "bold"), 
                              bg="#BFBFBF", fg="black")
file_section_label.pack(anchor=tk.W, padx=30, pady=(15, 5))

file_container = tk.Frame(scrollable_frame, bg="white", bd=2, relief=tk.SOLID)
file_container.pack(fill=tk.X, padx=30, pady=5)

def create_file_row(parent, icon_text, button_text, status_text):
    row_frame = tk.Frame(parent, bg="white")
    row_frame.pack(fill=tk.X, padx=10, pady=8)
    
    icon_frame = tk.Frame(row_frame, bg="#00C853", width=40, height=40)
    icon_frame.pack(side=tk.LEFT, padx=(0, 10))
    icon_frame.pack_propagate(False)
    icon_label = tk.Label(icon_frame, text=icon_text, font=("Arial", 16, "bold"), 
                         bg="#00C853", fg="white")
    icon_label.pack(expand=True)
    
    button = tk.Button(row_frame, text=button_text, font=("Arial", 10, "bold"),
                      bg="#FF6B35", fg="white", width=18, relief=tk.FLAT, cursor="hand2")
    button.pack(side=tk.LEFT)
    
    status = tk.Label(row_frame, text=status_text, font=("Arial", 9),
                     bg="gray", fg="white", anchor=tk.W, padx=10)
    status.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
    
    entry = tk.Entry(row_frame)
    
    def browse_file():
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            filename = path.split("/")[-1]
            status.config(text=filename, bg="green")
    
    button.config(command=browse_file)
    return entry

qa_checks_entry = create_file_row(file_container, "📊", "QA CHECKS RAW", "No raw data selected")
qa_review_entry = create_file_row(file_container, "📊", "QA REVIEW RAW", "No raw data selected")

# QCL Template
qcl_label = tk.Label(scrollable_frame, text="3.", font=("Arial", 14, "bold"), 
                     bg="#BFBFBF", fg="black")
qcl_label.pack(anchor=tk.W, padx=30, pady=(15, 5))

qcl_frame = tk.Frame(scrollable_frame, bg="white", bd=2, relief=tk.SOLID)
qcl_frame.pack(fill=tk.X, padx=30, pady=5)
qcl_entry = create_file_row(qcl_frame, "📄", "QCL TEMPLATE", "No template selected")

# Member List
member_label = tk.Label(scrollable_frame, text="4.", font=("Arial", 14, "bold"), 
                        bg="#BFBFBF", fg="black")
member_label.pack(anchor=tk.W, padx=30, pady=(15, 5))

member_frame = tk.Frame(scrollable_frame, bg="white", bd=2, relief=tk.SOLID)
member_frame.pack(fill=tk.X, padx=30, pady=5)
member_entry = create_file_row(member_frame, "👤", "MEMBER LIST", "No member list selected")

# AutoMate Button
automate_btn = tk.Button(scrollable_frame, text="AutoMate", font=("Arial", 14, "bold"),
                        bg="#FF6B35", fg="white", width=20, height=2,
                        relief=tk.FLAT, cursor="hand2", command=process_files)
automate_btn.pack(pady=30)

root.mainloop()