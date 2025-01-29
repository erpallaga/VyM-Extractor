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
    Load two sheets from 'people_data.xlsx':
      - 'people' (with roles/columns = YES/NO + optional Mod columns)
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

def compute_score(df_people, df_history, idx, part_key, meeting_date):
    """
    Score = (weeks since last assignment) * <part_key> Mod.
    The bigger => assigned sooner.
    """
    row = df_people.loc[idx]
    name = row["Hermano"]
    last_date = get_last_assignment_date(df_history, name, part_key)
    wks = weeks_since_assignment(last_date, meeting_date)

    mod_col = part_key + " Mod"  # e.g. "Tesoros Mod"
    mod_val = row.get(mod_col, 1.0)
    return wks * float(mod_val)

def get_top_candidates(df_people, df_history, part_key, meeting_date, assigned_so_far, top_n=3):
    """
    Return top N candidates for a given part_key and date, 
    excluding anyone in 'assigned_so_far' (they already have a part that day).
    """
    candidates_idx = []
    for idx, person in df_people.iterrows():
        if str(person.get("Activo?", "NO")).upper() != "YES":
            continue
        # Must have 'YES' in the <part_key> column
        if str(person.get(part_key, "NO")).upper() != "YES":
            continue
        # Exclude if they're already assigned that day
        hermano_name = person["Hermano"]
        if hermano_name in assigned_so_far:
            continue
        candidates_idx.append(idx)
    
    if not candidates_idx:
        return []

    # Compute scores
    scored = []
    for cidx in candidates_idx:
        sc = compute_score(df_people, df_history, cidx, part_key, meeting_date)
        scored.append((cidx, sc))
    
    # Sort descending by score
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:top_n]

def pick_candidate_interactively(top_candidates, df_people, part_label, date_str):
    """
    Show top 3 in console, let user pick 1,2,3 or skip.
    Return the chosen 'Hermano' name or None if skipped.
    """
    if not top_candidates:
        print(f"\nNo eligible candidates for {part_label} on {date_str}. Skipping.")
        return None

    print(f"\n=== {part_label} on {date_str} ===")
    for i, (idx_cand, sc) in enumerate(top_candidates):
        hermano = df_people.at[idx_cand, "Hermano"]
        print(f"{i+1}) {hermano} (score={sc:.2f})")

    choice = input("Choose 1, 2, 3 or 'skip': ").strip().lower()
    if choice == "skip":
        return None

    try:
        cnum = int(choice)
        if 1 <= cnum <= len(top_candidates):
            chosen_idx = top_candidates[cnum-1][0]
        else:
            chosen_idx = top_candidates[0][0]
    except:
        chosen_idx = top_candidates[0][0]

    return df_people.at[chosen_idx, "Hermano"]

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
#                     LOGIC TO IDENTIFY TESOROS & PERLAS
###############################################################################

def identify_tesoros_perlas(weekly_df, row_idx, col):
    """
    Suppose row_idx is in [4..7]. We read the cell:
      - If it starts with "1." => "Tesoros"
      - If it starts with "2." => "Perlas"
      - Otherwise => None
    """
    if row_idx not in weekly_df.index:
        return None

    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip().lower()

    if txt.startswith("1."):
        return "Tesoros"
    elif txt.startswith("2."):
        return "Perlas"
    else:
        return None  # e.g. "3. Lectura..." or blank => skip

###############################################################################
#         LOGIC TO IDENTIFY NVC, NECESIDADES, EBC from row content
###############################################################################

def identify_nvc_type(weekly_df, row_idx, col):
    """
    For rows [13..17], we read the cell:
      - if includes "estudio bíblico de la congregación" => "EBC"
      - if includes "necesidades de la congregación" => "Necesidades"
      - else => "NVC" (if not blank)
    """
    if row_idx not in weekly_df.index:
        return None

    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip().lower()

    if len(txt) == 0:
        return None
    elif "estudio bíblico de la congregación" in txt:
        return "EBC"
    elif "necesidades de la congregación" in txt:
        return "Necesidades"
    else:
        return "NVC"

###############################################################################
#                          MAIN ASSIGNMENT LOGIC
###############################################################################

def main_assignment():
    # 1) Load data
    df_people, df_history = load_people_data("people_data.xlsx")
    weekly_df = load_weekly_programs("weekly_programs.xlsx")
    date_cols = get_date_columns(weekly_df)

    # We'll output to a DataFrame with 6 final rows
    final_rows = ["PRESIDENCIA", "TESOROS", "PERLAS", "NVC1", "NVC2", "EBC"]
    df_final = pd.DataFrame(index=final_rows, columns=date_cols)

    # Map these 6 row labels to columns in df_people
    part_mapping = {
        "PRESIDENCIA": "Presidencias",
        "TESOROS": "Tesoros",
        "PERLAS": "Perlas",
        "NVC1": "NVC",        # or "Necesidades" if identified
        "NVC2": "NVC",        # same logic
        "EBC": "EBC",
    }

    for col in date_cols:
        mtg_date = parse_date_str(col)
        if not mtg_date:
            continue

        # We'll track who is assigned on this date
        assigned_today = set()

        #######################################################################
        # 1) PRESIDENCIA (hard-coded every date)
        #######################################################################
        top3 = get_top_candidates(df_people, df_history, part_mapping["PRESIDENCIA"], mtg_date, assigned_today)
        chosen = pick_candidate_interactively(top3, df_people, "PRESIDENCIA", col)
        if chosen:
            df_final.at["PRESIDENCIA", col] = chosen
            assigned_today.add(chosen)
            df_history = add_history(df_history, chosen, part_mapping["PRESIDENCIA"], mtg_date)

        #######################################################################
        # 2) TESOROS (rows 4..7 look for "1.")
        #######################################################################
        tesoros_found = False
        for r in range(3, 7):
            part_key = identify_tesoros_perlas(weekly_df, r, col)
            if part_key == "Tesoros":
                # Assign
                top3 = get_top_candidates(df_people, df_history, part_key, mtg_date, assigned_today)
                chosen_t = pick_candidate_interactively(top3, df_people, "TESOROS", col)
                if chosen_t:
                    df_final.at["TESOROS", col] = chosen_t
                    assigned_today.add(chosen_t)
                    df_history = add_history(df_history, chosen_t, part_key, mtg_date)
                tesoros_found = True
                break
        # if not found, leave TESOROS blank

        #######################################################################
        # 3) PERLAS (rows 4..7 look for "2.")
        #######################################################################
        perlas_found = False
        for r in range(3, 7):
            part_key = identify_tesoros_perlas(weekly_df, r, col)
            if part_key == "Perlas":
                top3 = get_top_candidates(df_people, df_history, part_key, mtg_date, assigned_today)
                chosen_p = pick_candidate_interactively(top3, df_people, "PERLAS", col)
                if chosen_p:
                    df_final.at["PERLAS", col] = chosen_p
                    assigned_today.add(chosen_p)
                    df_history = add_history(df_history, chosen_p, part_key, mtg_date)
                perlas_found = True
                break
        # if not found, leave PERLAS blank

        #######################################################################
        # 4) NVC / Necesidades / EBC (rows 13..17)
        #######################################################################
        nvc_parts = []  # each item is "EBC" or "Necesidades" or "NVC"
        for r in range(13, 18):
            cat = identify_nvc_type(weekly_df, r, col)
            if cat:
                nvc_parts.append(cat)  # e.g. "EBC", "Necesidades", "NVC"

        # Separate out EBC vs. the rest
        ebc_count = sum(1 for x in nvc_parts if x == "EBC")
        needs_nvc_list = [x for x in nvc_parts if x != "EBC"]  # "NVC" or "Necesidades"

        # NVC1
        if len(needs_nvc_list) > 0:
            cat1 = needs_nvc_list[0]  # "NVC" or "Necesidades"
            part_col_1 = "NVC" if cat1 == "NVC" else "Necesidades"
            top3 = get_top_candidates(df_people, df_history, part_col_1, mtg_date, assigned_today)
            chosen_1 = pick_candidate_interactively(top3, df_people, cat1.upper(), col)
            if chosen_1:
                df_final.at["NVC1", col] = chosen_1
                assigned_today.add(chosen_1)
                df_history = add_history(df_history, chosen_1, part_col_1, mtg_date)

        # NVC2
        if len(needs_nvc_list) > 1:
            cat2 = needs_nvc_list[1]  # "NVC" or "Necesidades"
            part_col_2 = "NVC" if cat2 == "NVC" else "Necesidades"
            top3 = get_top_candidates(df_people, df_history, part_col_2, mtg_date, assigned_today)
            chosen_2 = pick_candidate_interactively(top3, df_people, cat2.upper(), col)
            if chosen_2:
                df_final.at["NVC2", col] = chosen_2
                assigned_today.add(chosen_2)
                df_history = add_history(df_history, chosen_2, part_col_2, mtg_date)

        # EBC
        if ebc_count > 0:
            top3 = get_top_candidates(df_people, df_history, "EBC", mtg_date, assigned_today)
            chosen_ebc = pick_candidate_interactively(top3, df_people, "EBC", col)
            if chosen_ebc:
                df_final.at["EBC", col] = chosen_ebc
                assigned_today.add(chosen_ebc)
                df_history = add_history(df_history, chosen_ebc, "EBC", mtg_date)

    # End of all columns
    df_final.to_excel("final_assignments.xlsx")

    save_people_data(df_people, df_history, "people_data.xlsx")

    print("\nFinished all assignments!")
    print("Check 'final_assignments.xlsx' and 'people_data.xlsx' for results.")

###############################################################################
#                              ENTRY POINT
###############################################################################

if __name__ == "__main__":
    main_assignment()
