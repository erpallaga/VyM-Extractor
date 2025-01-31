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

SPECIALIZED_PARTS = {
    "Lectura",  # men only
    *SMM_SUBPARTS
}

###############################################################################
# STRIP " Sala B" SUFFIX
###############################################################################

def strip_sala_b_suffix(part_key: str) -> str:
    if part_key.endswith(" Sala B"):
        return part_key[:-7].strip()
    return part_key

def is_smm_subpart(part_key: str) -> bool:
    base = strip_sala_b_suffix(part_key)
    return base in SMM_SUBPARTS

###############################################################################
# LOAD/SAVE
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
    date_cols=[]
    for col in df.columns:
        dt = parse_date_str(col)
        if dt:
            date_cols.append(col)
    return date_cols

###############################################################################
# UNIFIED SMM LAST DATE
###############################################################################

def get_unified_smm_last_date(df_history, person_name):
    all_rows = df_history.loc[df_history["Name"]==person_name].copy()
    if all_rows.empty:
        return datetime(1900,1,1).date()

    def check_smm(p):
        return strip_sala_b_suffix(str(p)) in SMM_SUBPARTS

    all_rows["isSMM"] = all_rows["Part"].apply(check_smm).copy()
    sub = all_rows[all_rows["isSMM"]==True]
    if sub.empty:
        return datetime(1900,1,1).date()
    sub = sub.sort_values("AssignmentDate", ascending=False)
    parsed = parse_date_str(sub["AssignmentDate"].iloc[0])
    return parsed if parsed else datetime(1900,1,1).date()

def weeks_since_assignment(last_date, current_date):
    return (current_date - last_date).days / 7.0

def get_last_assignment_date(df_history, person_name, part_key):
    base = strip_sala_b_suffix(part_key)
    if base in SMM_SUBPARTS:
        return get_unified_smm_last_date(df_history, person_name)
    if base=="Lectura":
        # unify "Lectura" + "Lectura Sala B"
        sub = df_history[df_history["Name"]==person_name].copy()
        if sub.empty:
            return datetime(1900,1,1).date()
        def isLect(p):
            return strip_sala_b_suffix(str(p))=="Lectura"
        sub["isLect"] = sub["Part"].apply(isLect)
        sub = sub[sub["isLect"]==True]
        if sub.empty:
            return datetime(1900,1,1).date()
        sub = sub.sort_values("AssignmentDate", ascending=False)
        parsed = parse_date_str(sub["AssignmentDate"].iloc[0])
        return parsed if parsed else datetime(1900,1,1).date()
    # else direct match
    sub = df_history[
        (df_history["Name"]==person_name) &
        (df_history["Part"]==part_key)
    ]
    if sub.empty:
        return datetime(1900,1,1).date()
    sub = sub.sort_values("AssignmentDate", ascending=False)
    parsed = parse_date_str(sub["AssignmentDate"].iloc[0])
    return parsed if parsed else datetime(1900,1,1).date()

###############################################################################
# MAPPING to 'Estudiante'
###############################################################################

def get_people_column_for_part(part_key):
    base = strip_sala_b_suffix(part_key)
    if base in SPECIALIZED_PARTS:
        return "Estudiante"
    return base

###############################################################################
# compute_score_and_lastdate
###############################################################################

def compute_score_and_lastdate(df_people, df_history, idx, part_key, meeting_date):
    name = df_people.at[idx,"Hermano"]
    last_date = get_last_assignment_date(df_history, name, part_key)
    wks = weeks_since_assignment(last_date, meeting_date)
    col_name = get_people_column_for_part(part_key)
    mod_col = col_name + " Mod"
    mod_val = df_people.at[idx, mod_col] if mod_col in df_people.columns else 1.0
    score = wks * float(mod_val)
    return score, last_date

###############################################################################
# get_top_candidates
###############################################################################

def get_top_candidates(df_people, df_history, part_key, meeting_date,
                       assigned_so_far, top_n=3, required_gender=None):
    col_name = get_people_column_for_part(part_key)
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
        if str(row.get(col_name,"NO")).upper()!="YES":
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

def get_recent_smm_assignments(df_history, person_name, how_many=3):
    allr = df_history.loc[df_history["Name"]==person_name].copy()
    if allr.empty:
        return []
    def is_smm(x):
        return strip_sala_b_suffix(str(x)) in SMM_SUBPARTS
    allr["isSMM"] = allr["Part"].apply(is_smm)
    sub = allr[allr["isSMM"]==True]
    if sub.empty:
        return []
    sub = sub.sort_values("AssignmentDate", ascending=False)
    out=[]
    for _, row in sub.head(how_many).iterrows():
        dt_str = "None"
        dt_ = parse_date_str(row["AssignmentDate"])
        if dt_:
            dt_str = dt_.strftime("%d/%m/%Y")
        out.append((dt_str, row["Part"]))
    return out

def pick_candidate_interactively(top_candidates, df_people, df_history,
                                 part_key, part_label, date_str,
                                 assignment_text="", top_n=3):
    if not top_candidates:
        print(f"\nNo eligible candidates for {part_label} on {date_str}. Skipping.")
        return None

    base = strip_sala_b_suffix(part_key)
    if assignment_text:
        print(f"\n--- {part_label} on {date_str} ---\nAssignment: {assignment_text}")
    else:
        print(f"\n--- {part_label} on {date_str} ---")

    max_count = min(top_n, len(top_candidates))
    for i in range(max_count):
        cidx, sc, ldate = top_candidates[i]
        hermano = df_people.at[cidx,"Hermano"]
        last_str = "None" if ldate.year<1901 else ldate.strftime("%d/%m/%Y")
        info_line = f"{i+1}) {hermano} (score={sc:.2f}, last={last_str})"

        # if base in SMM_SUBPARTS => show 3 last SMM
        if base in SMM_SUBPARTS:
            rec_smm = get_recent_smm_assignments(df_history, hermano, how_many=3)
            for idxr, (dts, subp) in enumerate(rec_smm, start=1):
                info_line += f"\n   Last SMM(n-{idxr}): {dts} ({subp})"

        print(info_line)

    while True:
        choice = input(f"Choose 1..{max_count} or 'skip': ").strip().lower()
        if choice=="skip":
            return None
        try:
            cnum=int(choice)
            if 1<=cnum<=max_count:
                chosen_idx= top_candidates[cnum-1][0]
                return df_people.at[chosen_idx,"Hermano"]
        except:
            pass
        print(f"Invalid input! Type a number from 1..{max_count} or 'skip'.")

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
        ans = input("Assign to Varón (V), Mujer (M), or skip? [V/M/skip]: ").strip().lower()
        if ans in ["v","m"]:
            return ans.upper()
        if ans in ["skip",""]:
            return "skip"
        print("Invalid input. Type 'V','M','skip' or Enter for skip.")

###############################################################################
# Identify Tesoros/Perlas, NVC, Lectura, SMM
###############################################################################

def identify_tesoros_perlas(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return None,None
    val= weekly_df.at[row_idx,col]
    if not isinstance(val,str):
        val= str(val) if pd.notna(val) else ""
    txt= val.strip().lower()
    if txt.startswith("1."):
        return "Tesoros", val
    elif txt.startswith("2."):
        return "Perlas", val
    return None,None

def identify_nvc_type(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return None,None
    val= weekly_df.at[row_idx,col]
    if not isinstance(val,str):
        val= str(val) if pd.notna(val) else ""
    txt= val.strip().lower()
    if len(txt)==0:
        return None,None
    elif "estudio bíblico de la congregación" in txt:
        return "EBC", val
    elif "necesidades de la congregación" in txt:
        return "Necesidades", val
    else:
        return "NVC", val

def identify_lectura(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return False,None
    val= weekly_df.at[row_idx,col]
    if not isinstance(val,str):
        val= str(val) if pd.notna(val) else ""
    txt= val.strip().lower()
    if txt.startswith("3. lectura de la biblia"):
        return True,val
    return False,None

def identify_smm(weekly_df, row_idx, col):
    if row_idx not in weekly_df.index:
        return False,None
    val= weekly_df.at[row_idx,col]
    if not isinstance(val,str):
        val=str(val) if pd.notna(val) else ""
    txt= val.strip()
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
        "ORACION",
        "LECTURA Sala B",
        "SMM1 Sala B",
        "SMM2 Sala B",
        "SMM3 Sala B",
        "SMM4 Sala B"
    ]
    df_final = pd.DataFrame(index=final_rows, columns=date_cols)

    for col in date_cols:
        mtg_date = parse_date_str(col)
        if not mtg_date:
            continue

        assigned_today = set()

        #--- 1) PRESIDENCIA
        cand_pres = get_top_candidates(df_people, df_history, "Presidencia",
                                       mtg_date, assigned_today, top_n=3)
        chosen_pres = pick_candidate_interactively(cand_pres, df_people, df_history,
                                                   "Presidencia","PRESIDENCIA",
                                                   col)
        if chosen_pres:
            df_final.at["PRESIDENCIA", col] = chosen_pres
            assigned_today.add(chosen_pres)
            df_history = add_history(df_history, chosen_pres, "Presidencia", mtg_date)

        #--- 2) TESOROS & PERLAS (rows 3..4)
        tesoros_found=False
        perlas_found=False
        for r in range(3,5):
            part_key, text_val = identify_tesoros_perlas(weekly_df, r, col)
            if part_key=="Tesoros" and not tesoros_found:
                cand_t = get_top_candidates(df_people, df_history,
                                            "Tesoros", mtg_date, assigned_today, top_n=3)
                chosen_t = pick_candidate_interactively(cand_t, df_people, df_history,
                                                        "Tesoros","TESOROS", col,
                                                        assignment_text=text_val)
                if chosen_t:
                    df_final.at["TESOROS", col] = chosen_t
                    assigned_today.add(chosen_t)
                    df_history = add_history(df_history, chosen_t, "Tesoros", mtg_date)
                tesoros_found=True
            elif part_key=="Perlas" and not perlas_found:
                cand_p = get_top_candidates(df_people, df_history,
                                            "Perlas", mtg_date, assigned_today, top_n=3)
                chosen_p = pick_candidate_interactively(cand_p, df_people, df_history,
                                                        "Perlas","PERLAS", col,
                                                        assignment_text=text_val)
                if chosen_p:
                    df_final.at["PERLAS", col] = chosen_p
                    assigned_today.add(chosen_p)
                    df_history = add_history(df_history, chosen_p, "Perlas", mtg_date)
                perlas_found=True

        #--- 3) LECTURA => men only
        is_lec, lect_text = identify_lectura(weekly_df, 5, col)
        if is_lec:
            cand_lec = get_top_candidates(df_people, df_history,
                                          "Lectura", mtg_date,
                                          assigned_today, top_n=3,
                                          required_gender="V")
            chosen_lec = pick_candidate_interactively(cand_lec, df_people, df_history,
                                                      "Lectura","LECTURA", col,
                                                      assignment_text=lect_text,
                                                      top_n=3)
            if chosen_lec:
                df_final.at["LECTURA", col] = chosen_lec
                assigned_today.add(chosen_lec)
                df_history = add_history(df_history, chosen_lec, "Lectura", mtg_date)

        #--- 4) SMM => rows 7..9
        smm_index=1
        smm_texts_main = [None,None,None,None]  # store each row's text for replication
        for r_smm in range(7,10):
            if smm_index>4:
                break
            has_smm, smm_text = identify_smm(weekly_df, r_smm, col)
            smm_texts_main[smm_index-1] = smm_text if has_smm else None

            if has_smm:
                print(f"\n=== SMM Part row={r_smm} on {col} ===")
                print(f"Title: {smm_text}")

                subpart = None
                required_g=None
                low_smm = smm_text.lower()
                if "discurso" in low_smm:
                    subpart="Discurso"
                    required_g="V"
                elif "revisitas" in low_smm:
                    subpart="Haga Revisitas"
                    ans=ask_gender_or_skip()
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
                    subpart="Estudiante"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        smm_index+=1
                        continue
                    if ans in ["v","m"]:
                        required_g=ans.upper()

                cand_smm = get_top_candidates(df_people, df_history, subpart,
                                              mtg_date, assigned_today,
                                              top_n=5, required_gender=required_g)
                label_smm = f"SMM{smm_index}"
                chosen_smm = pick_candidate_interactively(cand_smm, df_people, df_history,
                                                          subpart, label_smm,
                                                          col, assignment_text=smm_text,
                                                          top_n=5)
                if chosen_smm:
                    df_final.at[label_smm, col] = chosen_smm
                    assigned_today.add(chosen_smm)
                    df_history = add_history(df_history, chosen_smm, subpart, mtg_date)

            smm_index+=1

        #--- 5) NVC / Necesidades / EBC => rows=13..15
        nvc_parts=[]
        for rr in range(13,16):
            cat, txt = identify_nvc_type(weekly_df, rr, col)
            if cat:
                nvc_parts.append((cat, txt))

        ebc_ent = [(c,t) for (c,t) in nvc_parts if c=="EBC"]
        other_nvc= [(c,t) for (c,t) in nvc_parts if c!="EBC"]

        # NVC1
        if len(other_nvc)>0:
            cat1, txt1= other_nvc[0]
            part1 = "NVC" if cat1=="NVC" else "Necesidades"
            cand_n1 = get_top_candidates(df_people, df_history, part1,
                                         mtg_date, assigned_today, top_n=3)
            chosen_1 = pick_candidate_interactively(cand_n1, df_people, df_history,
                                                    part1,"NVC1", col,
                                                    assignment_text=txt1,
                                                    top_n=3)
            if chosen_1:
                df_final.at["NVC1", col] = chosen_1
                assigned_today.add(chosen_1)
                df_history = add_history(df_history, chosen_1, part1, mtg_date)

        # NVC2
        if len(other_nvc)>1:
            cat2, txt2 = other_nvc[1]
            part2= "NVC" if cat2=="NVC" else "Necesidades"
            cand_n2 = get_top_candidates(df_people, df_history, part2,
                                         mtg_date, assigned_today, top_n=3)
            chosen_2 = pick_candidate_interactively(cand_n2, df_people, df_history,
                                                    part2,"NVC2", col,
                                                    assignment_text=txt2,
                                                    top_n=3)
            if chosen_2:
                df_final.at["NVC2", col] = chosen_2
                assigned_today.add(chosen_2)
                df_history = add_history(df_history, chosen_2, part2, mtg_date)

        # EBC
        if len(ebc_ent)>0:
            ebc_cat, ebc_txt = ebc_ent[0]
            cand_ebc = get_top_candidates(df_people, df_history, "EBC",
                                          mtg_date, assigned_today, top_n=3)
            chosen_ebc = pick_candidate_interactively(cand_ebc, df_people, df_history,
                                                      "EBC","EBC", col,
                                                      assignment_text=ebc_txt,
                                                      top_n=3)
            if chosen_ebc:
                df_final.at["EBC", col] = chosen_ebc
                assigned_today.add(chosen_ebc)
                df_history = add_history(df_history, chosen_ebc, "EBC", mtg_date)

        #--- 6) LECTOR EBC
        cand_le = get_top_candidates(df_people, df_history, "Lector EBC",
                                     mtg_date, assigned_today, top_n=3)
        chosen_le = pick_candidate_interactively(cand_le, df_people, df_history,
                                                 "Lector EBC","LECTOR EBC", col,
                                                 top_n=3)
        if chosen_le:
            df_final.at["LECTOR EBC", col] = chosen_le
            assigned_today.add(chosen_le)
            df_history = add_history(df_history, chosen_le, "Lector EBC", mtg_date)

        #--- 7) ORACION
        cand_or = get_top_candidates(df_people, df_history, "Oraciones",
                                     mtg_date, assigned_today, top_n=3)
        chosen_or = pick_candidate_interactively(cand_or, df_people, df_history,
                                                 "Oraciones","ORACION", col,
                                                 top_n=3)
        if chosen_or:
            df_final.at["ORACION", col] = chosen_or
            assigned_today.add(chosen_or)
            df_history = add_history(df_history, chosen_or, "Oraciones", mtg_date)

        #--- 8) SALA B prompt
        sala_b_ans = input("\nIs there SALA B for this week? [Y/N]: ").strip().lower()
        if sala_b_ans=="y":
            # replicate LECTURA Sala B if LECTURA in main was assigned
            if is_lec:
                # same text
                # men only
                cand_le_sb = get_top_candidates(df_people, df_history,
                                                "Lectura Sala B",
                                                mtg_date, assigned_today,
                                                top_n=3, required_gender="V")
                chosen_le_sb = pick_candidate_interactively(cand_le_sb, df_people, df_history,
                                                            "Lectura Sala B","LECTURA Sala B",
                                                            col, assignment_text=lect_text,
                                                            top_n=3)
                if chosen_le_sb:
                    df_final.at["LECTURA Sala B", col] = chosen_le_sb
                    assigned_today.add(chosen_le_sb)
                    df_history = add_history(df_history, chosen_le_sb, "Lectura Sala B", mtg_date)

            # replicate SMM1..SMM4 from main
            # smm_texts_main array => replicate in SALA B
            for i_sb in range(1,5):
                label_main = f"SMM{i_sb}"
                text_main = smm_texts_main[i_sb-1]  # might be None
                if not text_main:
                    # no assignment in main => skip
                    continue
                # parse it again for sub-part
                print(f"\n=== {label_main} Sala B on {col} ===")
                print(f"Title: {text_main}")
                low_sb = text_main.lower()
                subpart_b=None
                required_b=None

                if "discurso" in low_sb:
                    subpart_b="Discurso Sala B"
                    required_b="V"
                elif "revisitas" in low_sb:
                    subpart_b="Haga Revisitas Sala B"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        continue
                    if ans in ["V","M"]:
                        required_b=ans
                elif "empiece conversaciones" in low_sb:
                    subpart_b="Empiece conversaciones Sala B"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        continue
                    if ans in ["V","M"]:
                        required_b=ans
                elif "haga discípulos" in low_sb:
                    subpart_b="Haga discípulos Sala B"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        continue
                    if ans in ["V","M"]:
                        required_b=ans
                elif "explique sus creencias" in low_sb:
                    subpart_b="Explique sus creencias Sala B"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        continue
                    if ans in ["V","M"]:
                        required_b=ans
                else:
                    subpart_b="Estudiante Sala B"
                    ans=ask_gender_or_skip()
                    if ans=="skip":
                        continue
                    if ans in ["v","m"]:
                        required_b=ans.upper()

                # now get top candidates
                label_smm_b = f"SMM{i_sb} Sala B"
                cand_smm_b = get_top_candidates(df_people, df_history, subpart_b,
                                                mtg_date, assigned_today, top_n=5,
                                                required_gender=required_b)
                chosen_smm_b = pick_candidate_interactively(cand_smm_b, df_people, df_history,
                                                            subpart_b, label_smm_b, col,
                                                            assignment_text=text_main,
                                                            top_n=5)
                if chosen_smm_b:
                    df_final.at[label_smm_b, col] = chosen_smm_b
                    assigned_today.add(chosen_smm_b)
                    df_history = add_history(df_history, chosen_smm_b, subpart_b, mtg_date)

        # save partial progress for this date
        save_people_data(df_people, df_history, "people_data.xlsx")
        print(f"\nSaved partial progress after finishing assignments for {col}.")

    # end of columns
    df_final.to_excel("final_assignments.xlsx")
    print("\nAll done! Check 'final_assignments.xlsx' & 'people_data.xlsx' for final results.")


###############################################################################
# ENTRY POINT
###############################################################################

if __name__=="__main__":
    main_assignment()
