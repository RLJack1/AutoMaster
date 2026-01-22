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
            "SYD", "TUG", "WYN", "AU", "OVR", "Guest AU - Guest.AU"
        ]
    },
    "Austria": {
        "formatted_prefixes": ["PH2_AT"],
        "arbitrary_codes": [
            "SYD", "TUG", "WYN", "AU", "OVR", "Guest AU - Guest.AU"
        ]
    },
    "Belgium": {
        "formatted_prefixes": ["PH3_BE_ING_BANK", "PH3_BE"],
        "arbitrary_codes": [
            "BE", "M1 03 14", "Guest BE - Guest.BE"
        ]
    },
    "China": {
        "formatted_prefixes": ["PH2_CN"],
        "arbitrary_codes": [
            "CN", "China"
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
            "Germany", "Guest DE - Guest.DE"
        ]
    },
    "Hong Kong": {
        "formatted_prefixes": ["PH2_HK"],
        "arbitrary_codes": [
            "Hong Kong", "HK", "Hk"
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
            "Japan"
        ]
    },
    "Luxembourg": {
        "formatted_prefixes": ["PH2_LU"],
        "arbitrary_codes": [
            "LU"
        ]
    },
    "Mexico": {
        "formatted_prefixes": [],
        "arbitrary_codes": [
            "MX"
        ]
    },
    "Netherlands": {
        "formatted_prefixes": ["Guest NL"],
        "arbitrary_codes": [
            "ACT", "ALP", "AME", "AT",
            "BMG", 
            "CDR", 
            "DEA", "DP", 
            "HBP",
            "IE",
            "KBK","KCC", "KCM", "KFL", "KFR", "KGD", "KMH", "KNF", "KQB", "KQB", "KSY", "KXS", "KZK", "KZR", "KNZ", "KBE", "KLQ", "KQT",
            "LZ",
            "NL", "RO",
            "RBG", 
            "WP","WBM",
            "Guest NL - Guest.NL",
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
            "PL", "PL Bank", "PH2_SK_Brati", "P-PULAWSKA", "P-CHORZOWSKA", "P-CHORZ.50", "DR01R1402", "DR09R0102"
        ]
    },
    "Portugal": {
        "formatted_prefixes": [],
        "arbitrary_codes": [
            "PT",
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
            "RUSMCW001", "RU"
        ]
    },
    "Singapore": {
        "formatted_prefixes": ["PH2_SG"],
        "arbitrary_codes": [
            "SG", "Guest Asia - Guest.Asia", "Guest ASIA - Guest.ASIA"
        ]
    },
    "Slovakia": {
        "formatted_prefixes": ["PH1_SK","PH2_SK"],
        "arbitrary_codes": [
            "SK",
        ]
    },
    "South Korea": {
        "formatted_prefixes": ["PH2_KR"],
        "arbitrary_codes": [
            "Kr", "KR"
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
            "CH"
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
            "Guest UK - Guest.UK", "United Kingdom"
        ]
    },
    "United States": {
        "formatted_prefixes": ["PH2_US"],
        "arbitrary_codes": [
        ]
    },
    "Vietnam": {
        "formatted_prefixes": [""],
        "arbitrary_codes": [
            "Vn"
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
        return 0.15 / 2
    elif tenure < 12:
        return 0.10 / 2
    else:
        return 0.05 / 2
    
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

def infer_location_from_opened_for(opened_for_value):
    """
    Infers location from 'Opened for' field by checking for guest account patterns
    like "Guest XX - Guest.XX" where XX is a country code
    Returns the country name if found, empty string otherwise
    """
    if pd.isna(opened_for_value):
        return ""
    
    opened_for_str = str(opened_for_value).strip()
    
    # Check against all arbitrary codes in the country location map
    for country, rules in COUNTRY_LOCATION_MAP.items():
        for code in rules.get("arbitrary_codes", []):
            # Check if the code appears in the opened_for field (case-insensitive)
            if code.upper() in opened_for_str.upper():
                return country
    
    return ""

def get_case_location(location_value, opened_for_value):
    """
    Gets the location for a case, trying Location first, then falling back to Opened for
    """
    # Try to get location from Location column first
    location = format_location(location_value)
    
    # If location is missing or couldn't be determined, try to infer from Opened for
    if not location or location == str(location_value).title():
        inferred = infer_location_from_opened_for(opened_for_value)
        if inferred:
            return inferred
    
    return location

def find_column(df, possible_names):
    # Finds a column in df whose header matches one of the possible_names.
    # Matching is case-insensitive.
    
    for col in df.columns:
        col_name = str(col).strip().lower()
        for name in possible_names:
            if col_name == name.strip().lower():
                return col
    raise ValueError(f"Missing required column: {possible_names}")

def find_qcl_start_row(ws, search_column="B", search_text="control check no."):

    # Finds the first empty row below a specified search text in a specified column

    for row in range(1, ws.max_row + 1):
        cell_value = ws[f"{search_column}{row}"].value
        if cell_value and str(cell_value).strip().lower() == search_text.lower():
            r = row + 1
            while ws[f"{search_column}{r}"].value:
                r += 1
            return r
    raise ValueError(f"Could not find '{search_text}' in column {search_column} of QCL template.")

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

def extract_corpkey_from_name(formatted_name):

    # Extracts corp key from "FirstName LastName - CorpKey" format
    # Returns the corp key or None if not found

    if pd.isna(formatted_name):
        return None

    formatted_name = str(formatted_name).strip()

    # Check if it contains " - " separator
    if " - " in formatted_name:
        parts = formatted_name.split(" - ")
        if len(parts) >= 2:
            return parts[-1].strip()  # Return the last part (corp key)

    return None

# ---------------- MAIN PROCESS ---------------- #

def process_qa_reviews_reopened_cases(qa_review_file, member_file, wb, control_no_start):

    # Process QA Reviews - Reopened Cases sheet
    # Returns: (next_control_no, total_cases_written, matched_agents, unmatched_agents)

    try:
        # Read the "Reopened Cases" sheet from QA Review file
        reopened_df = pd.read_excel(qa_review_file, sheet_name="Reopened Cases")

        # Find required columns in Reopened Cases sheet
        # All columns follow format: "F09 - SNOW HR Cases[Column Name]"
        # We'll search for columns that contain the target name within brackets
        def find_column_containing(df, search_terms):
            """Find column that contains any of the search terms (case-insensitive)"""
            for col in df.columns:
                col_lower = str(col).strip().lower()
                for term in search_terms:
                    # Check if the term is in the column name (handles both bracket and non-bracket formats)
                    if f"[{term.strip().lower()}]" in col_lower or term.strip().lower() in col_lower:
                        return col
            raise ValueError(f"Missing required column containing: {search_terms}")

        assignment_group_col = find_column_containing(reopened_df, ["assignment group"])  # Column D -> QCL Column C
        country_code_col = find_column_containing(reopened_df, ["country code"])  # Column G -> QCL Column D
        user_id_col = find_column_containing(reopened_df, ["user id"])  # Column L (corp key only) -> QCL Column E
        number_col = find_column_containing(reopened_df, ["number"])  # Column E -> QCL Column F
        hr_service_col = find_column_containing(reopened_df, ["hr service"])  # Column C -> QCL Column G
        reopened_reason_col = find_column_containing(reopened_df, ["re-opened reason", "reopened reason"])  # Column AA -> QCL Column I

        # Read member list file
        member_header_row = detect_header_row(
            member_file,
            required_columns=["Team", "Tenure"]
        )
        members_df = pd.read_excel(member_file, header=member_header_row)

        # Find formatted name column (with corp key) and Team column (LastName, FirstName format)
        formatted_name_col = None
        for col in members_df.columns:
            sample_values = members_df[col].dropna().astype(str).head(5)
            if any(" - " in val for val in sample_values):
                formatted_name_col = col
                break

        member_name_col = formatted_name_col if formatted_name_col else find_column(members_df, ["Team"])
        team_col = find_column(members_df, ["Team"])  # Column A with "LastName, FirstName" format
        tenure_col = find_column(members_df, ["Tenure"])

        # Build corpkey to member mapping
        corpkey_to_member = {}
        for _, member in members_df.iterrows():
            member_name_raw = str(member[member_name_col]).strip()
            team_name = str(member[team_col]).strip()  # Get the "LastName, FirstName" format
            tenure = member[tenure_col]

            if not member_name_raw or member_name_raw.lower() == 'nan' or pd.isna(tenure):
                continue

            corpkey = extract_corpkey_from_name(member_name_raw)
            if corpkey:
                corpkey_to_member[corpkey] = {
                    'name': member_name_raw,
                    'team_name': team_name,  # Store the Team column name for display
                    'tenure': tenure,
                    'cases': []
                }

        # Build dictionary of cases per corpkey
        for _, row in reopened_df.iterrows():
            user_id = str(row[user_id_col]).strip()
            if user_id and user_id.lower() != 'nan':
                if user_id in corpkey_to_member:
                    corpkey_to_member[user_id]['cases'].append(row)

        # Access the Reopened Cases sheet
        if "Reopened Cases" not in wb.sheetnames:
            raise ValueError("QCL template does not contain 'Reopened Cases' sheet")

        ws = wb["Reopened Cases"]
        current_row = find_qcl_start_row(ws, search_column="B", search_text="control check no.")
        control_no = control_no_start

        # Track statistics
        matched_agents = 0
        total_cases_written = 0
        unmatched_agents = []

        # Process each member with cases
        for corpkey, member_data in corpkey_to_member.items():
            cases = member_data['cases']

            if len(cases) == 0:
                continue

            matched_agents += 1
            tenure = member_data['tenure']
            member_name = member_data['name']

            print(f"Agent: {member_name} | Corp Key: {corpkey} | Tenure: {tenure} | Total Cases: {len(cases)}")

            # Write ALL cases to QCL - Reopened Cases sheet (no sampling)
            for case in cases:
                ws[f"B{current_row}"].value = control_no
                ws[f"C{current_row}"].value = format_assignment_group(case[assignment_group_col])
                ws[f"D{current_row}"].value = format_location(case[country_code_col])
                ws[f"E{current_row}"].value = member_data['team_name']  # Use Team column (LastName, FirstName format)
                ws[f"F{current_row}"].value = case[number_col]
                ws[f"G{current_row}"].value = case[hr_service_col]
                ws[f"I{current_row}"].value = case[reopened_reason_col]  # Column I for Reopened Reason

                control_no += 1
                current_row += 1
                total_cases_written += 1

        # Find unmatched agents (members with no cases)
        for corpkey, member_data in corpkey_to_member.items():
            if len(member_data['cases']) == 0:
                unmatched_agents.append(member_data['name'])

        return control_no, total_cases_written, matched_agents, unmatched_agents

    except Exception as e:
        raise Exception(f"Error processing QA Reviews - Reopened Cases: {str(e)}")

def process_qa_reviews_breached_cases(qa_review_file, member_file, wb, control_no_start):

    # Process QA Reviews - Breached Cases sheet
    # Returns: (next_control_no, total_cases_written, matched_agents, unmatched_agents)

    try:
        # Read the "Breached Cases" sheet from QA Review file
        breached_df = pd.read_excel(qa_review_file, sheet_name="Breached Cases")

        # Find required columns in Breached Cases sheet
        def find_column_containing(df, search_terms):
            """Find column that contains any of the search terms (case-insensitive)"""
            for col in df.columns:
                col_lower = str(col).strip().lower()
                for term in search_terms:
                    # Check if the term is in the column name (handles both bracket and non-bracket formats)
                    if f"[{term.strip().lower()}]" in col_lower or term.strip().lower() in col_lower:
                        return col
            raise ValueError(f"Missing required column containing: {search_terms}")

        # Column mapping for Breached Cases:
        assignment_group_col = find_column_containing(breached_df, ["assignment group"])  # Column N -> QCL Column B
        country_code_col = find_column_containing(breached_df, ["country code"])  # Column O -> QCL Column C
        user_id_col = find_column_containing(breached_df, ["user id"])  # Column M (corp key only) -> QCL Column D
        number_col = find_column_containing(breached_df, ["number"])  # Column C -> QCL Column E
        hr_service_col = find_column_containing(breached_df, ["hr service"])  # Column D -> QCL Column F
        resolution_type_col = find_column_containing(breached_df, ["resolution type"])  # Column E -> QCL Column G
        breach_reason_col = find_column_containing(breached_df, ["sla breach reason"])  # Column AA -> QCL Column I

        # Read member list file
        member_header_row = detect_header_row(
            member_file,
            required_columns=["Team", "Tenure"]
        )
        members_df = pd.read_excel(member_file, header=member_header_row)

        # Find formatted name column (with corp key) and Team column (LastName, FirstName format)
        formatted_name_col = None
        for col in members_df.columns:
            sample_values = members_df[col].dropna().astype(str).head(5)
            if any(" - " in val for val in sample_values):
                formatted_name_col = col
                break

        member_name_col = formatted_name_col if formatted_name_col else find_column(members_df, ["Team"])
        team_col = find_column(members_df, ["Team"])  # Column A with "LastName, FirstName" format
        tenure_col = find_column(members_df, ["Tenure"])

        # Build corpkey to member mapping
        corpkey_to_member = {}
        for _, member in members_df.iterrows():
            member_name_raw = str(member[member_name_col]).strip()
            team_name = str(member[team_col]).strip()  # Get the "LastName, FirstName" format
            tenure = member[tenure_col]

            if not member_name_raw or member_name_raw.lower() == 'nan' or pd.isna(tenure):
                continue

            corpkey = extract_corpkey_from_name(member_name_raw)
            if corpkey:
                corpkey_to_member[corpkey] = {
                    'name': member_name_raw,
                    'team_name': team_name,  # Store the Team column name for display
                    'tenure': tenure,
                    'cases': []
                }

        # Build dictionary of cases per corpkey
        for _, row in breached_df.iterrows():
            user_id = str(row[user_id_col]).strip()
            if user_id and user_id.lower() != 'nan':
                if user_id in corpkey_to_member:
                    corpkey_to_member[user_id]['cases'].append(row)

        # Access the Breached Cases sheet
        if "Breached Cases" not in wb.sheetnames:
            raise ValueError("QCL template does not contain 'Breached Cases' sheet")

        ws = wb["Breached Cases"]
        current_row = find_qcl_start_row(ws, search_column="B", search_text="control check no.")
        control_no = control_no_start

        # Track statistics
        matched_agents = 0
        total_cases_written = 0
        unmatched_agents = []

        # Process each member with cases
        for corpkey, member_data in corpkey_to_member.items():
            cases = member_data['cases']

            if len(cases) == 0:
                continue

            matched_agents += 1
            tenure = member_data['tenure']
            member_name = member_data['name']

            print(f"Agent: {member_name} | Corp Key: {corpkey} | Tenure: {tenure} | Total Cases: {len(cases)}")

            # Write ALL cases to QCL - Breached Cases sheet (no sampling)
            for case in cases:
                ws[f"B{current_row}"].value = control_no
                ws[f"C{current_row}"].value = format_assignment_group(case[assignment_group_col])
                ws[f"D{current_row}"].value = format_location(case[country_code_col])
                ws[f"E{current_row}"].value = member_data['team_name']  # Use Team column (LastName, FirstName format)
                ws[f"F{current_row}"].value = case[number_col]
                ws[f"G{current_row}"].value = case[hr_service_col]
                ws[f"H{current_row}"].value = case[resolution_type_col]  # Column H for SLA Breach Type
                ws[f"I{current_row}"].value = case[breach_reason_col]  # Column I for Breach Reason

                control_no += 1
                current_row += 1
                total_cases_written += 1

        # Find unmatched agents (members with no cases)
        for corpkey, member_data in corpkey_to_member.items():
            if len(member_data['cases']) == 0:
                unmatched_agents.append(member_data['name'])

        return control_no, total_cases_written, matched_agents, unmatched_agents

    except Exception as e:
        raise Exception(f"Error processing QA Reviews - Breached Cases: {str(e)}")

def process_files():
    raw_file = qa_checks_entry.get()
    qa_review_file = qa_review_entry.get()
    qcl_template = qcl_entry.get()
    member_file = member_entry.get()

    if not raw_file or not member_file or not qcl_template:
        messagebox.showerror("Error", "Please select all required files.")
        return

    try:
        # Copy template and preserve all features first
        file_ext = os.path.splitext(qcl_template)[1]
        output_path = qcl_template.replace(file_ext, f"_POPULATED{file_ext}")
        shutil.copy2(qcl_template, output_path)

        wb = load_workbook(
            output_path,
            keep_vba=True,
            data_only=False,
            keep_links=True
        )

        # Initialize statistics
        qa_checks_stats = None
        qa_reviews_reopened_stats = None
        qa_reviews_breached_stats = None

        # ========== PROCESS QA CHECKS ==========
        if raw_file:
            print("\n========== PROCESSING QA CHECKS ==========")

            # Read raw data file
            raw_df = pd.read_excel(raw_file)
            location_col = find_column(raw_df, ["Location"])
            assigned_col = find_column(raw_df, ["Assigned to", "Assigned To"])
            case_id_col = find_column(raw_df, ["Number"])
            service_col = find_column(raw_df, ["HR Service"])
            assignment_group_col = find_column(raw_df, ["Assignment group"])
            opened_for_col = find_column(raw_df, ["Opened for", "Opened For"])

            # Read member list file
            member_header_row = detect_header_row(
                member_file,
                required_columns=["Team", "Tenure"]
            )
            members_df = pd.read_excel(member_file, header=member_header_row)

            # Find formatted name column (with corp key) and Team column (LastName, FirstName format)
            formatted_name_col = None
            for col in members_df.columns:
                sample_values = members_df[col].dropna().astype(str).head(5)
                if any(" - " in val for val in sample_values):
                    formatted_name_col = col
                    break

            member_name_col = formatted_name_col if formatted_name_col else find_column(members_df, ["Team"])
            team_col = find_column(members_df, ["Team"])  # Column A with "LastName, FirstName" format
            tenure_col = find_column(members_df, ["Tenure"])

            # Build corpkey to member mapping
            corpkey_to_member = {}
            for _, member in members_df.iterrows():
                member_name_raw = str(member[member_name_col]).strip()
                team_name = str(member[team_col]).strip()  # Get the "LastName, FirstName" format
                tenure = member[tenure_col]

                if not member_name_raw or member_name_raw.lower() == 'nan' or pd.isna(tenure):
                    continue

                corpkey = extract_corpkey_from_name(member_name_raw)
                if corpkey:
                    corpkey_to_member[corpkey] = {
                        'name': member_name_raw,
                        'team_name': team_name,  # Store the Team column name for display
                        'tenure': tenure,
                        'cases': []
                    }

            # Build dictionary of cases per corpkey
            for _, row in raw_df.iterrows():
                assigned_to = str(row[assigned_col]).strip()
                if assigned_to and assigned_to.lower() != 'nan':
                    # Extract corpkey from assigned_to field
                    corpkey = extract_corpkey_from_name(assigned_to)
                    if corpkey and corpkey in corpkey_to_member:
                        corpkey_to_member[corpkey]['cases'].append(row)

            # DEBUGGING: Print matching statistics
            print(f"\n=== MATCHING STATISTICS ===")
            print(f"Total cases in raw file: {len(raw_df)}")
            print(f"Total agents in member list: {len(members_df)}")
            print(f"Agents with corp keys: {len(corpkey_to_member)}")

            ws = wb["QA Checks"]
            current_row = find_qcl_start_row(ws)
            control_no = 1

            # Track processing statistics
            matched_agents = 0
            total_cases_written = 0
            unmatched_agents = []

            # Process each member with cases
            for corpkey, member_data in corpkey_to_member.items():
                cases = member_data['cases']

                if len(cases) == 0:
                    unmatched_agents.append(member_data['name'])
                    continue

                matched_agents += 1
                tenure = member_data['tenure']
                member_name = member_data['name']

                # Calculate sample size based on tenure
                sample_size = max(1, math.ceil(len(cases) * calculate_sample_percentage(tenure)))
                sampled_cases = random.sample(cases, min(sample_size, len(cases)))

                print(f"Agent: {member_name} | Corp Key: {corpkey} | Tenure: {tenure} | Total Cases: {len(cases)} | Sample: {len(sampled_cases)}")

                # Write sampled cases to QCL
                for case in sampled_cases:
                    ws[f"B{current_row}"].value = control_no
                    ws[f"C{current_row}"].value = format_assignment_group(case[assignment_group_col])
                    ws[f"D{current_row}"].value = get_case_location(case[location_col], case[opened_for_col])
                    ws[f"E{current_row}"].value = member_data['team_name']  # Use Team column (LastName, FirstName format)
                    ws[f"F{current_row}"].value = case[case_id_col]
                    ws[f"G{current_row}"].value = case[service_col]
                    control_no += 1
                    current_row += 1
                    total_cases_written += 1

            # Print final statistics
            print(f"\n=== QA CHECKS FINAL RESULTS ===")
            print(f"Matched agents: {matched_agents}")
            print(f"Total cases written: {total_cases_written}")
            print(f"Unmatched agents: {len(unmatched_agents)}")
            if unmatched_agents[:5]:  # Show first 5 unmatched
                print(f"Sample unmatched: {unmatched_agents[:5]}")

            qa_checks_stats = {
                'matched': matched_agents,
                'cases': total_cases_written,
                'unmatched': len(unmatched_agents)
            }

        # ========== PROCESS QA REVIEWS - REOPENED CASES ==========
        if qa_review_file:
            print("\n========== PROCESSING QA REVIEWS - REOPENED CASES ==========")

            try:
                control_no_start = 1  # Start from 1 for QA Reviews
                _, cases_written, matched, unmatched_list = process_qa_reviews_reopened_cases(
                    qa_review_file,
                    member_file,
                    wb,
                    control_no_start
                )

                print(f"\n=== QA REVIEWS REOPENED CASES FINAL RESULTS ===")
                print(f"Matched agents: {matched}")
                print(f"Total cases written: {cases_written}")
                print(f"Unmatched agents: {len(unmatched_list)}")
                if unmatched_list[:5]:
                    print(f"Sample unmatched: {unmatched_list[:5]}")

                qa_reviews_reopened_stats = {
                    'matched': matched,
                    'cases': cases_written,
                    'unmatched': len(unmatched_list)
                }
            except Exception as e:
                print(f"Warning: Could not process QA Reviews - Reopened Cases - {str(e)}")
                qa_reviews_reopened_stats = None

        # ========== PROCESS QA REVIEWS - BREACHED CASES ==========
        if qa_review_file:
            print("\n========== PROCESSING QA REVIEWS - BREACHED CASES ==========")

            try:
                control_no_start = 1  # Start from 1 for Breached Cases
                _, cases_written, matched, unmatched_list = process_qa_reviews_breached_cases(
                    qa_review_file,
                    member_file,
                    wb,
                    control_no_start
                )

                print(f"\n=== QA REVIEWS BREACHED CASES FINAL RESULTS ===")
                print(f"Matched agents: {matched}")
                print(f"Total cases written: {cases_written}")
                print(f"Unmatched agents: {len(unmatched_list)}")
                if unmatched_list[:5]:
                    print(f"Sample unmatched: {unmatched_list[:5]}")

                qa_reviews_breached_stats = {
                    'matched': matched,
                    'cases': cases_written,
                    'unmatched': len(unmatched_list)
                }
            except Exception as e:
                print(f"Warning: Could not process QA Reviews - Breached Cases - {str(e)}")
                qa_reviews_breached_stats = None

        # Save workbook
        wb.save(output_path)
        wb.close()

        # Build success message
        success_msg = "QCL generated successfully!\n\n"
        success_msg += f"Output: {output_path}\n\n"

        if qa_checks_stats:
            success_msg += "=== QA CHECKS ===\n"
            success_msg += f"Matched agents: {qa_checks_stats['matched']}\n"
            success_msg += f"Cases written: {qa_checks_stats['cases']}\n"
            success_msg += f"Unmatched agents: {qa_checks_stats['unmatched']}\n\n"

        if qa_reviews_reopened_stats:
            success_msg += "=== QA REVIEWS (Reopened Cases) ===\n"
            success_msg += f"Matched agents: {qa_reviews_reopened_stats['matched']}\n"
            success_msg += f"Cases written: {qa_reviews_reopened_stats['cases']}\n"
            success_msg += f"Unmatched agents: {qa_reviews_reopened_stats['unmatched']}\n\n"

        if qa_reviews_breached_stats:
            success_msg += "=== QA REVIEWS (Breached Cases) ===\n"
            success_msg += f"Matched agents: {qa_reviews_breached_stats['matched']}\n"
            success_msg += f"Cases written: {qa_reviews_breached_stats['cases']}\n"
            success_msg += f"Unmatched agents: {qa_reviews_breached_stats['unmatched']}"

        messagebox.showinfo("Success", success_msg)

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
    "Circle 6 (Recruitment Admin)"
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