import pandas as pd
from datetime import datetime

###############################################################################
#                          HELPER FUNCTIONS
###############################################################################

def parse_date_str(date_val):
    """Convert a cell value to a datetime.date if possible."""
    if isinstance(date_val, datetime):
        return date_val.date()
    if isinstance(date_val, str):
        # Attempt various common formats
        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]:
            try:
                return datetime.strptime(date_val, fmt).date()
            except ValueError:
                continue
    return None

def weeks_since_assignment(last_date, current_date):
    """Return how many weeks (float) between two datetime.date objects."""
    return (current_date - last_date).days / 7.0

def load_people_data(filename="people_data.xlsx"):
    """
    Load two sheets from people_data.xlsx:
      - 'people' (columns for each part=YES/NO, plus a 'Mod' column for weighting)
      - 'AssignmentHistory' (list of past assignments).
    """
    xls = pd.ExcelFile(filename)
    
    # Read 'people'
    df_people = pd.read_excel(xls, sheet_name="people")

    # Read or create 'AssignmentHistory'
    if "AssignmentHistory" in xls.sheet_names:
        df_history = pd.read_excel(xls, sheet_name="AssignmentHistory")
    else:
        df_history = pd.DataFrame(columns=["Name", "Part", "AssignmentDate"])
    
    return df_people, df_history

def save_people_data(df_people, df_history, filename="people_data.xlsx"):
    """
    Overwrite the people_data.xlsx with updated 'people' and 'AssignmentHistory'.
    (We typically won't change 'people' but we do update 'AssignmentHistory'.)
    """
    with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
        df_people.to_excel(writer, sheet_name="people", index=False)
        df_history.to_excel(writer, sheet_name="AssignmentHistory", index=False)

def load_weekly_programs(filename="weekly_programs.xlsx"):
    """
    We only care about columns that represent meeting dates.
    We won't use the content—just the column headers.
    """
    return pd.read_excel(filename)

def get_all_date_columns(df):
    """
    From df.columns, pick out which ones can be parsed as dates.
    Return them in a list (as the original column strings).
    """
    date_cols = []
    for col in df.columns:
        d = parse_date_str(col)
        if d:
            date_cols.append(col)
    return date_cols

def get_last_assignment_date(df_history, person_name, part):
    """
    Look up the most recent assignment date for (person_name, part)
    in df_history. Returns a datetime.date or 1900-01-01 if none found.
    """
    relevant = df_history[
        (df_history["Name"] == person_name) &
        (df_history["Part"] == part)
    ]
    if relevant.empty:
        return datetime(1900,1,1).date()
    # Sort descending by date
    sorted_dates = relevant["AssignmentDate"].sort_values(ascending=False)
    most_recent = sorted_dates.iloc[0]
    # Because we store it as string "dd/mm/yyyy", parse it back
    parsed = parse_date_str(most_recent)
    return parsed if parsed else datetime(1900,1,1).date()

def compute_score(df_people, df_history, idx, part_key, meeting_date):
    """
    Score = (weeks since last assignment) * (modifier).
    The bigger => the earlier we suggest them.
    
    part_key is "Presidencias" or "Tesoros" or "Perlas", etc.
    If we have <part_key> Mod column, we use that weighting factor.
    """
    row = df_people.loc[idx]
    name = row["Hermano"]

    # 1) how long since last assignment
    last_date = get_last_assignment_date(df_history, name, part_key)
    wks = weeks_since_assignment(last_date, meeting_date)

    # 2) find the modifier
    mod_col = part_key + " Mod"
    mod_val = row.get(mod_col, 1.0)  # default 1.0 if no column

    return wks * float(mod_val)

def get_top_candidates(df_people, df_history, part_key, meeting_date, top_n=3):
    """
    Filter df_people rows to Active=YES, <part_key>=YES,
    then rank them by compute_score descending. Return top N (index, score).
    """
    # 1) gather eligible
    cand_indices = []
    for idx, person in df_people.iterrows():
        if str(person.get("Activo?", "NO")).upper() != "YES":
            continue
        yesno = str(person.get(part_key, "NO")).upper()
        if yesno == "YES":
            cand_indices.append(idx)
    
    if not cand_indices:
        return []

    # 2) compute scores
    scored = []
    for cidx in cand_indices:
        sc = compute_score(df_people, df_history, cidx, part_key, meeting_date)
        scored.append((cidx, sc))
    
    # 3) sort descending
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:top_n]


###############################################################################
#                           MAIN ASSIGNMENT LOGIC
###############################################################################

def main_assignment():
    # 1) Load people + history
    df_people, df_history = load_people_data("people_data.xlsx")

    # 2) Load weekly programs (just to get date columns)
    df_programs = load_weekly_programs("weekly_programs.xlsx")
    date_cols = get_all_date_columns(df_programs)
    # e.g. date_cols might be ["04/03/2025","18/03/2025","25/03/2025",...]

    # 3) Build an empty final_assignments DataFrame with:
    #    - Index = [ "PRESIDENCIA", "TESOROS", "PERLAS" ]
    #    - Columns = date_cols
    parts_list = ["PRESIDENCIA", "TESOROS", "PERLAS"]
    df_final = pd.DataFrame(index=parts_list, columns=date_cols)

    # We need to map these row labels to the columns in df_people:
    # "PRESIDENCIA" => "Presidencias", "TESOROS" => "Tesoros", "PERLAS" => "Perlas"
    part_mapping = {
        "PRESIDENCIA": "Presidencia",
        "TESOROS": "Tesoros",
        "PERLAS": "Perlas"
    }

    # 4) For each date, for each part, propose top 3, user picks or skip
    for date_col in date_cols:
        # parse the date in python
        mtg_date = parse_date_str(date_col)
        if not mtg_date:
            continue  # skip if not valid

        for part_name in parts_list:
            part_key = part_mapping[part_name]  # e.g. "Tesoros" or "Presidencias"
            
            # Find top 3
            top_cands = get_top_candidates(df_people, df_history, part_key, mtg_date, top_n=3)
            if not top_cands:
                print(f"\nNo eligible candidates for {part_name} on {date_col}. Skipping.")
                df_final.at[part_name, date_col] = ""
                continue

            print(f"\n=== {part_name.upper()} on {date_col} ===")
            for i, (idx_cand, sc) in enumerate(top_cands):
                person_name = df_people.at[idx_cand, "Hermano"]
                print(f"{i+1}) {person_name} (score={sc:.2f})")

            choice = input("Choose 1, 2, 3 or 'skip': ").strip().lower()
            if choice == "skip":
                # leave blank
                chosen_idx = None
            else:
                try:
                    choice_num = int(choice)
                    if 1 <= choice_num <= len(top_cands):
                        chosen_idx = top_cands[choice_num-1][0]
                    else:
                        chosen_idx = top_cands[0][0]  # fallback
                except:
                    chosen_idx = top_cands[0][0]  # fallback

            if chosen_idx is None:
                df_final.at[part_name, date_col] = ""
            else:
                chosen_name = df_people.at[chosen_idx, "Hermano"]
                df_final.at[part_name, date_col] = chosen_name

                # Save assignment to history
                # Format the meeting date as dd/mm/yyyy
                date_str = mtg_date.strftime("%d/%m/%Y")
                new_assignment = pd.DataFrame([{
                    "Name": chosen_name,
                    "Part": part_key,
                    "AssignmentDate": date_str
                }])
                df_history = pd.concat([df_history, new_assignment], ignore_index=True)

    # 5) Save final_assignments.xlsx
    df_final.to_excel("final_assignments.xlsx")

    # 6) Save updated people_data with new assignment history
    save_people_data(df_people, df_history, "people_data.xlsx")

    print("\nAll done! Check 'final_assignments.xlsx' for the results.")

###############################################################################
#                              SCRIPT ENTRY POINT
###############################################################################

if __name__ == "__main__":
    main_assignment()
