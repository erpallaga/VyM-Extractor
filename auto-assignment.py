import pandas as pd
from datetime import datetime

###############################################################################
#                          HELPER FUNCTIONS
###############################################################################

def parse_date_str(date_val):
    """
    Converts a cell value to a datetime.date if possible, trying common formats.
    """
    if isinstance(date_val, datetime):
        return date_val.date()
    if isinstance(date_val, str):
        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]:
            try:
                return datetime.strptime(date_val, fmt).date()
            except ValueError:
                continue
    return None

def weeks_since_assignment(last_date, current_date):
    """Return how many (float) weeks have passed from last_date to current_date."""
    return (current_date - last_date).days / 7.0

def load_people_data(filename="people_data.xlsx"):
    """
    Loads two sheets from 'people_data.xlsx':
      - 'people' (with roles/columns = YES/NO + optional Mod columns + 'Estudiante')
      - 'AssignmentHistory' (Name, Part, AssignmentDate).
    """
    xls = pd.ExcelFile(filename)
    df_people = pd.read_excel(xls, sheet_name="people")
    if "AssignmentHistory" in xls.sheet_names:
        df_history = pd.read_excel(xls, sheet_name="AssignmentHistory")
    else:
        df_history = pd.DataFrame(columns=["Name", "Part", "AssignmentDate"])
    return df_people, df_history

def save_people_data(df_people, df_history, filename="people_data.xlsx"):
    """
    Overwrite the existing 'people_data.xlsx' with updated 'people' and 
    'AssignmentHistory' sheets.
    """
    with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
        df_people.to_excel(writer, sheet_name="people", index=False)
        df_history.to_excel(writer, sheet_name="AssignmentHistory", index=False)

def load_weekly_programs(filename="weekly_programs.xlsx"):
    """Reads the entire weekly_programs file into a DataFrame."""
    return pd.read_excel(filename)

def get_date_columns(df):
    """
    From df.columns, pick out which are valid date strings (or datetimes).
    Return them in a list (the original strings).
    """
    date_cols = []
    for col in df.columns:
        dt = parse_date_str(col)
        if dt:
            date_cols.append(col)
    return date_cols

def get_last_assignment_date(df_history, person_name, part):
    """
    Finds the most recent AssignmentDate in df_history for (person_name, part).
    Returns a datetime.date or 1900-01-01 if none found.
    """
    relevant = df_history[
        (df_history["Name"] == person_name) &
        (df_history["Part"] == part)
    ]
    if relevant.empty:
        return datetime(1900,1,1).date()
    # Sort descending
    sorted_dates = relevant["AssignmentDate"].sort_values(ascending=False)
    most_recent = sorted_dates.iloc[0]
    parsed = parse_date_str(most_recent)
    return parsed if parsed else datetime(1900,1,1).date()

def compute_score_and_lastdate(df_people, df_history, idx, part_key, meeting_date):
    """
    Returns (score, last_date) for a given person & part.
     - score = (weeks since last assignment) * <part_key> Mod
     - last_date = their most recent assignment date of that part
    """
    row = df_people.loc[idx]
    name = row["Hermano"]

    last_date = get_last_assignment_date(df_history, name, part_key)
    wks = weeks_since_assignment(last_date, meeting_date)

    mod_col = part_key + " Mod"  # e.g. "Tesoros Mod" or "Estudiante Mod"
    mod_val = row.get(mod_col, 1.0)
    score = wks * float(mod_val)
    return score, last_date

def get_top_candidates(
    df_people, 
    df_history, 
    part_key, 
    meeting_date, 
    assigned_so_far, 
    top_n=3, 
    required_gender=None, 
    must_be_estudiante=False
):
    """
    Return a list of up to 'top_n': (idx, score, last_date)
     - part_key: which part name we store in df_history (e.g. "Tesoros", "EBC", or "Estudiante")
     - assigned_so_far: set of names already assigned this date
     - required_gender: "V" for men, "M" for women, or None for no filter
     - must_be_estudiante: if True, only pick rows with Estudiante=YES
    Sort by 'score' descending => highest first.
    """
    candidates_idx = []
    for idx, person in df_people.iterrows():
        # Must be active
        if str(person.get("Activo?", "NO")).upper() != "YES":
            continue

        name = person["Hermano"]

        # If we require "Estudiante=YES"
        if must_be_estudiante:
            if str(person.get("Estudiante", "NO")).upper() != "YES":
                continue

        # If we require a certain gender
        if required_gender is not None:
            gen = str(person.get("Género", "")).upper()
            if gen != required_gender.upper():
                continue

        # If name is already assigned that day, skip
        if name in assigned_so_far:
            continue

        # If not "Estudiante", we require that part_key=YES in df_people
        if part_key not in ["Estudiante"]:
            if str(person.get(part_key, "NO")).upper() != "YES":
                continue
        
        candidates_idx.append(idx)
    
    if not candidates_idx:
        return []

    # Compute (score, last_date)
    scored = []
    for cidx in candidates_idx:
        sc, ldate = compute_score_and_lastdate(df_people, df_history, cidx, part_key, meeting_date)
        scored.append((cidx, sc, ldate))
    
    # Sort descending
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:top_n]

def pick_candidate_interactively(top_candidates, df_people, part_label, date_str, assignment_text="", top_n=3):
    """
    Show up to 'top_n' in console, with score & last assignment date.
    Let user pick 1..top_n or 'skip'. We'll repeatedly ask if input invalid.
    Return the chosen 'Hermano' name or None if 'skip'.
    """
    if not top_candidates:
        print(f"\nNo eligible candidates for {part_label} on {date_str}. Skipping.")
        return None

    # Show the assignment text if provided
    if assignment_text:
        print(f"\n--- {part_label} on {date_str} ---")
        print(f"Assignment: {assignment_text}")
    else:
        print(f"\n--- {part_label} on {date_str} ---")

    for i, (idx_cand, sc, ldate) in enumerate(top_candidates):
        if i >= top_n:
            break
        hermano = df_people.at[idx_cand, "Hermano"]
        # Format last date, or "None"
        if ldate > datetime(1900,1,1).date():
            last_str = ldate.strftime("%d/%m/%Y")
        else:
            last_str = "None"
        print(f"{i+1}) {hermano} (score={sc:.2f}, last={last_str})")

    # Repeatedly prompt until we get a valid entry
    while True:
        choice = input(f"Choose 1..{top_n} or 'skip': ").strip().lower()
        if choice == "skip":
            return None
        try:
            cnum = int(choice)
            if 1 <= cnum <= len(top_candidates):
                chosen_idx = top_candidates[cnum-1][0]
                return df_people.at[chosen_idx, "Hermano"]
        except:
            pass

        print(f"Invalid input! Type a number from 1..{len(top_candidates)} or 'skip'.")

def add_history(df_history, hermano_name, part_key, mtg_date):
    """
    Append a new row to df_history using dd/mm/yyyy format for AssignmentDate.
    """
    date_str = mtg_date.strftime("%d/%m/%Y")
    new_row = pd.DataFrame([{
        "Name": hermano_name,
        "Part": part_key,
        "AssignmentDate": date_str
    }])
    df_history = pd.concat([df_history, new_row], ignore_index=True)
    return df_history

###############################################################################
#             PRESIDENCIA, TESOROS, PERLAS, LECTURA, SMM, NVC, EBC
###############################################################################

def identify_tesoros_perlas(weekly_df, row_idx, col):
    """
    For Tesoros (1.) or Perlas (2.) in certain rows.
    Return (part_key, assignment_text).
    """
    if row_idx not in weekly_df.index:
        return None, None

    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip()

    low = txt.lower()
    if low.startswith("1."):
        return "Tesoros", txt
    elif low.startswith("2."):
        return "Perlas", txt
    else:
        return None, None

def identify_nvc_type(weekly_df, row_idx, col):
    """
    For rows [13..17], read the cell:
      - if includes "estudio bíblico de la congregación" => ("EBC", text)
      - if includes "necesidades de la congregación" => ("Necesidades", text)
      - else => ("NVC", text) if not blank
      - (None, None) if blank
    """
    if row_idx not in weekly_df.index:
        return None, None

    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip()
    low = txt.lower()

    if len(low) == 0:
        return None, None
    elif "estudio bíblico de la congregación" in low:
        return "EBC", txt
    elif "necesidades de la congregación" in low:
        return "Necesidades", txt
    else:
        return "NVC", txt

def identify_lectura(weekly_df, row_idx, col):
    """
    If row_idx is 6, check if it starts with '3. Lectura de la Biblia'.
    Return (True, assignment_text) if yes, else (False, None).
    """
    if row_idx not in weekly_df.index:
        return False, None
    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip()
    if txt.lower().startswith("3. lectura de la biblia"):
        return True, txt
    return False, None

def identify_smm(weekly_df, row_idx, col):
    """
    If row_idx in [8..11], check if there's any text. 
    If there's text, we treat it as an SMM assignment. Return that text.
    If blank => (False, None).
    """
    if row_idx not in weekly_df.index:
        return False, None
    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip()
    if len(txt) == 0:
        return False, None
    return True, txt  # there's text => SMM assignment

###############################################################################
#                          MAIN ASSIGNMENT LOGIC
###############################################################################

def main_assignment():
    df_people, df_history = load_people_data("people_data.xlsx")
    weekly_df = load_weekly_programs("weekly_programs.xlsx")
    date_cols = get_date_columns(weekly_df)

    final_rows = [
        "PRESIDENCIA",
        "TESOROS",
        "PERLAS",
        "LECTURA",  
        "SMM1",
        "SMM2",
        "SMM3",
        "SMM4",
        "NVC1",
        "NVC2",
        "EBC"
    ]
    df_final = pd.DataFrame(index=final_rows, columns=date_cols)

    part_mapping = {
        "PRESIDENCIA": "Presidencia",
        "TESOROS": "Tesoros",
        "PERLAS": "Perlas",
        "NVC1": "NVC",        
        "NVC2": "NVC",
        "EBC": "EBC",
        "LECTURA": "Estudiante",  # men only
        "SMMx": "Estudiante",     # men/women
    }

    for col in date_cols:
        mtg_date = parse_date_str(col)
        if not mtg_date:
            continue

        assigned_today = set()

        # 1) PRESIDENCIA
        top3 = get_top_candidates(
            df_people, df_history,
            part_mapping["PRESIDENCIA"],
            mtg_date, assigned_today,
            top_n=3
        )
        chosen = pick_candidate_interactively(top3, df_people, "PRESIDENCIA", col)
        if chosen:
            df_final.at["PRESIDENCIA", col] = chosen
            assigned_today.add(chosen)
            df_history = add_history(df_history, chosen, part_mapping["PRESIDENCIA"], mtg_date)

        # 2) Tesoros & Perlas (rows 3..4 in your final code)
        tesoros_found = False
        perlas_found = False
        for r in range(3, 5):
            part_key, assignment_text = identify_tesoros_perlas(weekly_df, r, col)
            if part_key == "Tesoros" and not tesoros_found:
                top_candidates = get_top_candidates(
                    df_people, df_history,
                    "Tesoros",
                    mtg_date,
                    assigned_today,
                    top_n=3
                )
                chosen_t = pick_candidate_interactively(top_candidates, df_people, "TESOROS", col, assignment_text)
                if chosen_t:
                    df_final.at["TESOROS", col] = chosen_t
                    assigned_today.add(chosen_t)
                    df_history = add_history(df_history, chosen_t, "Tesoros", mtg_date)
                tesoros_found = True
            elif part_key == "Perlas" and not perlas_found:
                top_candidates = get_top_candidates(
                    df_people, df_history,
                    "Perlas",
                    mtg_date,
                    assigned_today,
                    top_n=3
                )
                chosen_p = pick_candidate_interactively(top_candidates, df_people, "PERLAS", col, assignment_text)
                if chosen_p:
                    df_final.at["PERLAS", col] = chosen_p
                    assigned_today.add(chosen_p)
                    df_history = add_history(df_history, chosen_p, "Perlas", mtg_date)
                perlas_found = True

        # 3) LECTURA (row 5 in your final code)
        is_lectura, lectura_text = identify_lectura(weekly_df, 5, col)
        if is_lectura:
            top7 = get_top_candidates(
                df_people, df_history,
                part_key="Estudiante",
                meeting_date=mtg_date,
                assigned_so_far=assigned_today,
                top_n=7,
                required_gender="V",       # men
                must_be_estudiante=True
            )
            chosen_lec = pick_candidate_interactively(top7, df_people, "LECTURA BIBLIA", col, lectura_text, top_n=7)
            if chosen_lec:
                df_final.at["LECTURA", col] = chosen_lec
                assigned_today.add(chosen_lec)
                df_history = add_history(df_history, chosen_lec, "Lectura", mtg_date)

        # 4) SMM (rows 7..9 in your final code)
        smm_index = 1
        for r in range(7, 10):
            if smm_index > 4:
                break
            has_smm, smm_text = identify_smm(weekly_df, r, col)
            if has_smm:
                print(f"\n=== SMM Part row={r} on {col} ===")
                print(f"Title: {smm_text}")
                if "Discurso" in smm_text.lower():
                    # always men
                    choice = "v"
                else:
                    # loop until v/m/skip
                    while True:
                        raw = input("Assign to VARÓN (V), MUJER (M), or skip? [V/M/skip]: ").strip().lower()
                        if raw == "skip" or raw in ["v","m",""]:
                            choice = raw
                            break
                        print("Invalid input. Type 'V', 'M', or 'skip'. (Empty => skip)")

                if choice == "" or choice == "skip":
                    # skip
                    pass
                elif choice in ["v","m"]:
                    required_gender = "V" if choice=="v" else "M"
                    top7 = get_top_candidates(
                        df_people,
                        df_history,
                        part_key="Estudiante",
                        meeting_date=mtg_date,
                        assigned_so_far=assigned_today,
                        top_n=7,
                        required_gender=required_gender,
                        must_be_estudiante=True
                    )
                    smm_label = f"SMM{smm_index}"
                    chosen_smm = pick_candidate_interactively(top7, df_people, smm_label, col, smm_text, top_n=7)
                    if chosen_smm:
                        df_final.at[smm_label, col] = chosen_smm
                        assigned_today.add(chosen_smm)
                        df_history = add_history(df_history, chosen_smm, "SMM", mtg_date)
                smm_index += 1

        # 5) NVC / Necesidades / EBC (rows 13..14 in your final code)
        nvc_parts_info = []
        for r in range(13, 15):
            cat, txt = identify_nvc_type(weekly_df, r, col)
            if cat:
                nvc_parts_info.append((cat, txt))

        ebc_entries = [(c, t) for (c, t) in nvc_parts_info if c == "EBC"]
        other_nvc_list = [(c, t) for (c, t) in nvc_parts_info if c != "EBC"]

        # NVC1
        if len(other_nvc_list) > 0:
            cat1, text1 = other_nvc_list[0]
            part_col_1 = "NVC" if cat1 == "NVC" else "Necesidades"
            top3 = get_top_candidates(
                df_people, df_history,
                part_key=part_col_1,
                meeting_date=mtg_date,
                assigned_so_far=assigned_today,
                top_n=3
            )
            chosen_1 = pick_candidate_interactively(top3, df_people, cat1.upper(), col, text1)
            if chosen_1:
                df_final.at["NVC1", col] = chosen_1
                assigned_today.add(chosen_1)
                df_history = add_history(df_history, chosen_1, part_col_1, mtg_date)

        # NVC2
        if len(other_nvc_list) > 1:
            cat2, text2 = other_nvc_list[1]
            part_col_2 = "NVC" if cat2 == "NVC" else "Necesidades"
            top3 = get_top_candidates(
                df_people, df_history,
                part_key=part_col_2,
                meeting_date=mtg_date,
                assigned_so_far=assigned_today,
                top_n=3
            )
            chosen_2 = pick_candidate_interactively(top3, df_people, cat2.upper(), col, text2)
            if chosen_2:
                df_final.at["NVC2", col] = chosen_2
                assigned_today.add(chosen_2)
                df_history = add_history(df_history, chosen_2, part_col_2, mtg_date)

        # EBC
        if len(ebc_entries) > 0:
            ebc_cat, ebc_text = ebc_entries[0]
            top3 = get_top_candidates(
                df_people, df_history,
                part_key="EBC",
                meeting_date=mtg_date,
                assigned_so_far=assigned_today,
                top_n=3
            )
            chosen_ebc = pick_candidate_interactively(top3, df_people, "EBC", col, ebc_text)
            if chosen_ebc:
                df_final.at["EBC", col] = chosen_ebc
                assigned_today.add(chosen_ebc)
                df_history = add_history(df_history, chosen_ebc, "EBC", mtg_date)

        # SAVE after finishing this date
        save_people_data(df_people, df_history, "people_data.xlsx")
        print(f"\nSaved partial progress after finishing assignments for {col}.")
        
    # End of columns
    df_final.to_excel("final_assignments.xlsx")
    print("\nFinished all assignments!")
    print("Check 'final_assignments.xlsx' and 'people_data.xlsx' for results.")

###############################################################################
#                              ENTRY POINT
###############################################################################

if __name__ == "__main__":
    main_assignment()
