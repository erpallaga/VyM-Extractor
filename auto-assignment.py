import pandas as pd
from datetime import datetime

###############################################################################
# CONFIG: SMM SUBPARTS => unify last-date among them
###############################################################################
SMM_SUBPARTS = {
    "Discurso",
    "Haga Revisitas",
    "Empiece conversaciones",
    "Haga discípulos",
    "Explique sus creencias"
}

# We might also consider "Lectura" in SMM_SUBPARTS if you want to unify it,
# but for now we keep it separate.

# Set of "specialized" parts that map to 'Estudiante' in df_people
SPECIALIZED_PARTS = {
    "Lectura",
    *SMM_SUBPARTS
}

###############################################################################
# HELPER FUNCTIONS
###############################################################################

def parse_date_str(date_val):
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
    return (current_date - last_date).days / 7.0

def load_people_data(filename="people_data.xlsx"):
    xls = pd.ExcelFile(filename)
    df_people = pd.read_excel(xls, sheet_name="people")
    if "AssignmentHistory" in xls.sheet_names:
        df_history = pd.read_excel(xls, sheet_name="AssignmentHistory")
    else:
        df_history = pd.DataFrame(columns=["Name","Part","AssignmentDate"])
    return df_people, df_history

def save_people_data(df_people, df_history, filename="people_data.xlsx"):
    with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
        df_people.to_excel(writer, sheet_name="people", index=False)
        df_history.to_excel(writer, sheet_name="AssignmentHistory", index=False)

def load_weekly_programs(filename="weekly_programs.xlsx"):
    return pd.read_excel(filename)

def get_date_columns(df):
    date_cols = []
    for col in df.columns:
        dt = parse_date_str(col)
        if dt:
            date_cols.append(col)
    return date_cols

###############################################################################
# SMM UNIFICATION
###############################################################################

def is_smm_subpart(part_key):
    return part_key in SMM_SUBPARTS

def get_unified_smm_last_date(df_history, person_name):
    relevant = df_history[
        (df_history["Name"] == person_name) &
        (df_history["Part"].isin(SMM_SUBPARTS))
    ]
    if relevant.empty:
        return datetime(1900,1,1).date()
    sorted_dates = relevant["AssignmentDate"].sort_values(ascending=False)
    most_recent = sorted_dates.iloc[0]
    parsed = parse_date_str(most_recent)
    return parsed if parsed else datetime(1900,1,1).date()

def get_last_assignment_date(df_history, person_name, part_key):
    """
    If part_key in SMM_SUBPARTS => unify among them,
    else do the usual.
    """
    if is_smm_subpart(part_key):
        return get_unified_smm_last_date(df_history, person_name)
    else:
        # normal logic
        relevant = df_history[
            (df_history["Name"] == person_name) &
            (df_history["Part"] == part_key)
        ]
        if relevant.empty:
            return datetime(1900,1,1).date()
        sorted_dates = relevant["AssignmentDate"].sort_values(ascending=False)
        most_recent = sorted_dates.iloc[0]
        parsed = parse_date_str(most_recent)
        return parsed if parsed else datetime(1900,1,1).date()

def get_recent_smm_assignments(df_history, person_name, how_many=3):
    """
    Return up to 'how_many' of the person's most recent SMM_SUBPARTS,
    each as (date_str, subpart). If none, empty list.
    """
    relevant = df_history[
        (df_history["Name"] == person_name) &
        (df_history["Part"].isin(SMM_SUBPARTS))
    ]
    if relevant.empty:
        return []
    relevant = relevant.sort_values("AssignmentDate", ascending=False)
    results=[]
    for _, row in relevant.head(how_many).iterrows():
        dt_obj = parse_date_str(row["AssignmentDate"])
        dt_str = dt_obj.strftime("%d/%m/%Y") if dt_obj else "None"
        part_name = row["Part"]
        results.append((dt_str, part_name))
    return results

###############################################################################
# MAPPING specialized => "Estudiante" in df_people
###############################################################################

def get_people_column_for_part(part_key):
    if part_key in SPECIALIZED_PARTS:
        return "Estudiante"
    return part_key

###############################################################################
# compute_score_and_lastdate
###############################################################################

def compute_score_and_lastdate(df_people, df_history, idx, part_key, meeting_date):
    name = df_people.at[idx,"Hermano"]
    last_date = get_last_assignment_date(df_history, name, part_key)
    wks = weeks_since_assignment(last_date, meeting_date)

    people_col = get_people_column_for_part(part_key)
    mod_col = people_col + " Mod"
    mod_val = df_people.at[idx, mod_col] if mod_col in df_people.columns else 1.0
    score = wks * float(mod_val)
    return score, last_date

###############################################################################
# get_top_candidates
###############################################################################

def get_top_candidates(df_people, df_history, part_key, meeting_date,
                       assigned_so_far, top_n=3, required_gender=None):
    """
    Filter df_people => top N by (weeks since last assignment) * mod.
    If part_key in SPECIALIZED_PARTS => check 'Estudiante=YES' column, etc.
    If part_key in SMM_SUBPARTS => unify last-date among them.
    """
    people_col = get_people_column_for_part(part_key)
    valid_idx=[]
    for idx, row in df_people.iterrows():
        if str(row.get("Activo?", "NO")).upper()!="YES":
            continue
        name = row["Hermano"]
        if name in assigned_so_far:
            continue
        if required_gender:
            gen = str(row.get("Género","")).upper()
            if gen!=required_gender.upper():
                continue

        if str(row.get(people_col,"NO")).upper()!="YES":
            continue

        valid_idx.append(idx)

    if not valid_idx:
        return []

    scored=[]
    for cidx in valid_idx:
        sc, ldate = compute_score_and_lastdate(df_people, df_history, cidx, part_key, meeting_date)
        scored.append((cidx, sc, ldate))

    scored.sort(key=lambda x:x[1], reverse=True)
    return scored[:top_n]

###############################################################################
# pick_candidate_interactively
###############################################################################

def pick_candidate_interactively(top_candidates, df_people, df_history,
                                 real_part_key, part_label, date_str,
                                 assignment_text="", top_n=3):
    """
    Show top_n with extended info if real_part_key in SMM_SUBPARTS => show last 3 SMM.
    """
    if not top_candidates:
        print(f"\nNo eligible candidates for {part_label} on {date_str}. Skipping.")
        return None

    if assignment_text:
        print(f"\n--- {part_label} on {date_str} ---\nAssignment: {assignment_text}")
    else:
        print(f"\n--- {part_label} on {date_str} ---")

    max_count = min(top_n, len(top_candidates))
    for i in range(max_count):
        cidx, sc, ldate = top_candidates[i]
        hermano = df_people.at[cidx,"Hermano"]
        if ldate.year>1900:
            last_str = ldate.strftime("%d/%m/%Y")
        else:
            last_str = "None"

        info_line = f"{i+1}) {hermano} (score={sc:.2f}, last={last_str})"

        # if part_key is in SMM_SUBPARTS => show last 3 SMM
        if is_smm_subpart(real_part_key):
            rec_smm = get_recent_smm_assignments(df_history, hermano, how_many=3)
            for idx_r, (dt_val, subp) in enumerate(rec_smm, start=1):
                info_line += f"\n   Last SMM(n-{idx_r}): {dt_val} ({subp})"

        print(info_line)

    while True:
        choice=input(f"Choose 1..{max_count} or 'skip': ").strip().lower()
        if choice=="skip":
            return None
        try:
            cnum=int(choice)
            if 1<=cnum<=max_count:
                chosen_idx = top_candidates[cnum-1][0]
                return df_people.at[chosen_idx,"Hermano"]
        except:
            pass

        print(f"Invalid input. Type a number from 1..{max_count} or 'skip'.")

###############################################################################
# add_history
###############################################################################

def add_history(df_history, hermano_name, part_key, mtg_date):
    date_str = mtg_date.strftime("%d/%m/%Y")
    new_row = pd.DataFrame([{
        "Name": hermano_name,
        "Part": part_key,
        "AssignmentDate": date_str
    }])
    df_history = pd.concat([df_history, new_row], ignore_index=True)
    return df_history

###############################################################################
# ask_gender_or_skip
###############################################################################

def ask_gender_or_skip():
    while True:
        ans=input("Assign to Varón (V), Mujer (M), or skip? [V/M/skip]: ").strip().lower()
        if ans in ["v","m"]:
            return ans.upper()
        if ans in ["skip",""]:
            return "skip"
        print("Invalid input. Type 'V','M','skip' or press Enter for skip.")

###############################################################################
# identify_ Tesoros/Perlas, NVC, Lectura, SMM
###############################################################################

def identify_tesoros_perlas(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return None,None
    val = weekly_df.at[row_idx, col]
    if not isinstance(val, str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip().lower()
    if txt.startswith("1."):
        return "Tesoros", val
    elif txt.startswith("2."):
        return "Perlas", val
    return None,None

def identify_nvc_type(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return None,None
    val = weekly_df.at[row_idx, col]
    if not isinstance(val,str):
        val = str(val) if pd.notna(val) else ""
    txt = val.strip().lower()
    if len(txt)==0:
        return None,None
    elif "estudio bíblico de la congregación" in txt:
        return "EBC",val
    elif "necesidades de la congregación" in txt:
        return "Necesidades", val
    return "NVC", val

def identify_lectura(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return False,None
    val = weekly_df.at[row_idx,col]
    if not isinstance(val,str):
        val=str(val) if pd.notna(val) else ""
    txt=val.strip().lower()
    if txt.startswith("3. lectura de la biblia"):
        return True,val
    return False,None

def identify_smm(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return False,None
    val = weekly_df.at[row_idx,col]
    if not isinstance(val,str):
        val=str(val) if pd.notna(val) else ""
    txt=val.strip()
    if len(txt)==0:
        return False,None
    return True, txt

###############################################################################
# MAIN ASSIGNMENT
###############################################################################

def main_assignment():
    df_people, df_history = load_people_data("people_data.xlsx")
    weekly_df = load_weekly_programs("weekly_programs.xlsx")
    date_cols = get_date_columns(weekly_df)

    # final rows now including LECTOR EBC + ORACION at the end
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
        "EBC",
        "LECTOR EBC",
        "ORACION"
    ]
    df_final = pd.DataFrame(index=final_rows, columns=date_cols)

    for col in date_cols:
        mtg_date = parse_date_str(col)
        if not mtg_date:
            continue
        assigned_today = set()

        # 1) PRESIDENCIA
        cand_pres = get_top_candidates(df_people, df_history, "Presidencia",
                                       mtg_date, assigned_today, top_n=3)
        chosen_pres = pick_candidate_interactively(cand_pres, df_people, df_history,
                                                   "Presidencia","PRESIDENCIA",col)
        if chosen_pres:
            df_final.at["PRESIDENCIA", col] = chosen_pres
            assigned_today.add(chosen_pres)
            df_history = add_history(df_history, chosen_pres, "Presidencia", mtg_date)

        # 2) TESOROS & PERLAS
        tesoros_found=False
        perlas_found=False
        for r in range(3,5):
            part_key, text_val = identify_tesoros_perlas(weekly_df, r, col)
            if part_key=="Tesoros" and not tesoros_found:
                cand_t = get_top_candidates(df_people, df_history,
                                            "Tesoros", mtg_date,
                                            assigned_today, top_n=3)
                chosen_t = pick_candidate_interactively(cand_t, df_people, df_history,
                                                        "Tesoros","TESOROS",
                                                        col, assignment_text=text_val,
                                                        top_n=3)
                if chosen_t:
                    df_final.at["TESOROS", col] = chosen_t
                    assigned_today.add(chosen_t)
                    df_history = add_history(df_history, chosen_t, "Tesoros", mtg_date)
                tesoros_found=True
            elif part_key=="Perlas" and not perlas_found:
                cand_p = get_top_candidates(df_people, df_history,
                                            "Perlas", mtg_date,
                                            assigned_today, top_n=3)
                chosen_p = pick_candidate_interactively(cand_p, df_people, df_history,
                                                        "Perlas","PERLAS",
                                                        col, assignment_text=text_val,
                                                        top_n=3)
                if chosen_p:
                    df_final.at["PERLAS", col] = chosen_p
                    assigned_today.add(chosen_p)
                    df_history = add_history(df_history, chosen_p, "Perlas", mtg_date)
                perlas_found=True

        # 3) LECTURA row=5 => men only
        is_lec, lect_text = identify_lectura(weekly_df, 5, col)
        if is_lec:
            cand_lec = get_top_candidates(df_people, df_history,
                                          "Lectura", mtg_date,
                                          assigned_today, top_n=7,
                                          required_gender="V")
            chosen_lec = pick_candidate_interactively(cand_lec, df_people, df_history,
                                                      "Lectura","LECTURA BIBLIA",
                                                      col, assignment_text=lect_text,
                                                      top_n=7)
            if chosen_lec:
                df_final.at["LECTURA", col] = chosen_lec
                assigned_today.add(chosen_lec)
                df_history = add_history(df_history, chosen_lec, "Lectura", mtg_date)

        # 4) SMM (rows=7..9)
        smm_index=1
        for r_smm in range(7,10):
            if smm_index>4:
                break
            has_smm, smm_text = identify_smm(weekly_df, r_smm, col)
            if has_smm:
                print(f"\n=== SMM Part row={r_smm} on {col} ===")
                print(f"Title: {smm_text}")

                low_smm = smm_text.lower()
                subpart=None
                required_g=None
                if "discurso" in low_smm:
                    subpart="Discurso"
                    required_g="V"  # men only
                elif "revisitas" in low_smm:
                    subpart="Haga Revisitas"
                    ans = ask_gender_or_skip()
                    if ans=="skip":
                        smm_index+=1
                        continue
                    if ans in ["V","M"]:
                        required_g=ans
                elif "empiece conversaciones" in low_smm:
                    subpart="Empiece conversaciones"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        smm_index+=1
                        continue
                    if ans in ["V","M"]:
                        required_g=ans
                elif "haga discípulos" in low_smm:
                    subpart="Haga discípulos"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        smm_index+=1
                        continue
                    if ans in ["V","M"]:
                        required_g=ans
                elif "explique sus creencias" in low_smm:
                    subpart="Explique sus creencias"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        smm_index+=1
                        continue
                    if ans in ["V","M"]:
                        required_g=ans
                else:
                    # fallback => "Estudiante"
                    subpart="Estudiante"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        smm_index+=1
                        continue
                    if ans in ["v","m"]:
                        required_g=ans.upper()

                cands_smm = get_top_candidates(df_people, df_history,
                                               subpart, mtg_date,
                                               assigned_today, top_n=7,
                                               required_gender=required_g)
                label_smm = f"SMM{smm_index}"
                chosen_smm = pick_candidate_interactively(cands_smm, df_people, df_history,
                                                          subpart, label_smm, col,
                                                          assignment_text=smm_text,
                                                          top_n=7)
                if chosen_smm:
                    df_final.at[label_smm, col] = chosen_smm
                    assigned_today.add(chosen_smm)
                    df_history = add_history(df_history, chosen_smm, subpart, mtg_date)
                smm_index+=1

        # 5) NVC / Necesidades / EBC => rows=13..15
        nvc_info=[]
        for r_nvc in range(13,16):
            cat, txt_val = identify_nvc_type(weekly_df, r_nvc, col)
            if cat:
                nvc_info.append((cat, txt_val))

        ebc_ent = [(c,t) for (c,t) in nvc_info if c=="EBC"]
        other_nvc= [(c,t) for (c,t) in nvc_info if c!="EBC"]

        # NVC1
        if len(other_nvc)>0:
            cat1, txt1= other_nvc[0]
            part1 = "NVC" if cat1=="NVC" else "Necesidades"
            cands_n1 = get_top_candidates(df_people, df_history, part1,
                                          mtg_date, assigned_today, top_n=3)
            chosen_1 = pick_candidate_interactively(cands_n1, df_people, df_history,
                                                    part1, cat1.upper(), col,
                                                    assignment_text=txt1, top_n=3)
            if chosen_1:
                df_final.at["NVC1", col] = chosen_1
                assigned_today.add(chosen_1)
                df_history = add_history(df_history, chosen_1, part1, mtg_date)

        # NVC2
        if len(other_nvc)>1:
            cat2, txt2 = other_nvc[1]
            part2 = "NVC" if cat2=="NVC" else "Necesidades"
            cands_n2 = get_top_candidates(df_people, df_history, part2,
                                          mtg_date, assigned_today, top_n=3)
            chosen_2 = pick_candidate_interactively(cands_n2, df_people, df_history,
                                                    part2, cat2.upper(), col,
                                                    assignment_text=txt2, top_n=3)
            if chosen_2:
                df_final.at["NVC2", col] = chosen_2
                assigned_today.add(chosen_2)
                df_history = add_history(df_history, chosen_2, part2, mtg_date)

        # EBC
        if len(ebc_ent)>0:
            ebc_cat, ebc_txt = ebc_ent[0]
            cands_ebc = get_top_candidates(df_people, df_history, "EBC",
                                           mtg_date, assigned_today, top_n=3)
            chosen_ebc = pick_candidate_interactively(cands_ebc, df_people, df_history,
                                                      "EBC", "EBC", col,
                                                      assignment_text=ebc_txt, top_n=3)
            if chosen_ebc:
                df_final.at["EBC", col] = chosen_ebc
                assigned_today.add(chosen_ebc)
                df_history = add_history(df_history, chosen_ebc, "EBC", mtg_date)

        #######################################################################
        # 6) LECTOR EBC => "Lector EBC" in df_people
        #######################################################################
        cand_le = get_top_candidates(df_people, df_history, "Lector EBC",
                                     mtg_date, assigned_today, top_n=3)
        chosen_le = pick_candidate_interactively(cand_le, df_people, df_history,
                                                 "Lector EBC", "LECTOR EBC", col,
                                                 top_n=3)
        if chosen_le:
            df_final.at["LECTOR EBC", col] = chosen_le
            assigned_today.add(chosen_le)
            df_history = add_history(df_history, chosen_le, "Lector EBC", mtg_date)

        #######################################################################
        # 7) ORACION => "Oraciones" in df_people
        #######################################################################
        cand_or = get_top_candidates(df_people, df_history, "Oraciones",
                                     mtg_date, assigned_today, top_n=3)
        chosen_or = pick_candidate_interactively(cand_or, df_people, df_history,
                                                 "Oraciones", "ORACION", col,
                                                 top_n=3)
        if chosen_or:
            df_final.at["ORACION", col] = chosen_or
            assigned_today.add(chosen_or)
            df_history = add_history(df_history, chosen_or, "Oraciones", mtg_date)

        # Save partial progress
        save_people_data(df_people, df_history, "people_data.xlsx")
        print(f"\nSaved partial progress after finishing assignments for {col}.")

    # end of columns
    df_final.to_excel("final_assignments.xlsx")
    print("\nDone! Check 'final_assignments.xlsx' and 'people_data.xlsx' for results.")


###############################################################################
# Script entry
###############################################################################

if __name__=="__main__":
    main_assignment()
