#Link to Dataset: https://data.sba.gov/en/dataset/7-a-504-foia/resource/6898b986-a895-47b4-bb7e-c6b286b23a7b
#Note on filters for CPR calculation:
#1) Exclude loans that are CANCLD (never became a real loan), COMMIT (not yet disbursed), and CLSIN (closed inactive) - filter in column AJ
#2) Exclude revolvers since CPR is only for loans with a defined payoff path - filter in column AN

#Imports libraries and loads the data
import pandas as pd
import numpy as np
file_path = r"foia-7a-fy2020-present-asof-260331.csv"
df = pd.read_csv(file_path, low_memory=False)

#Cleans data
df["loanstatus"] = (df["loanstatus"].astype(str).str.strip().str.upper().str.replace(" ", ""))
df["naicsdescription"] = df["naicsdescription"].astype(str).str.strip()

#Filter 1 - removes revolvers
if "revolverstatus" in df.columns:
    df = df[df["revolverstatus"] == 0]

#Filter 2 - only keeps valid loans
valid_statuses = ["PIF", "CURR", "DELINQ", "PSTDUE", "DEFERD", "CHGOFF", "LIQUID", "PURCH(NOTC/O)"]
df = df[df["loanstatus"].isin(valid_statuses)]

#CPR calculation
df["event"] = np.where(df["loanstatus"] == "PIF", 1, 0)
cpr_table = df.groupby("naicsdescription").agg(total_loans=("event", "count"), prepays=("event", "sum")).reset_index()

#Removes small sample industries (<500 count in dataset)
cpr_table = cpr_table[cpr_table["total_loans"] > 500]

#Finds CPR
cpr_table["CPR (%)"] = (cpr_table["prepays"] / cpr_table["total_loans"] * 100).round(2)
high_cpr = cpr_table.sort_values("CPR (%)", ascending=False)
low_cpr = cpr_table.sort_values("CPR (%)", ascending=True)

#Outputs results in table
print("\nHigh CPR (bad for IO)\n")
print(high_cpr[["naicsdescription", "CPR (%)"]].head(10))
print("\nLow CPR (good for IO)\n")
print(low_cpr[["naicsdescription", "CPR (%)"]].head(10))