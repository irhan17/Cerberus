#Link to Dataset: https://data.sba.gov/en/dataset/7-a-504-foia/resource/6898b986-a895-47b4-bb7e-c6b286b23a7b
#Findings quantitatively validate the knowledge that CPR is inversely correlated to interest rates

#Imports libraries and loads data
import pandas as pd
import numpy as np
file_path = r"foia-7a-fy2020-present-asof-260331.csv"
df = pd.read_csv(file_path, low_memory=False)
df["loanstatus"] = (df["loanstatus"].astype(str).str.strip().str.upper().str.replace(" ", ""))
df["initialinterestrate"] = pd.to_numeric(df["initialinterestrate"], errors="coerce")

#Filter 1 - removes revolvers
df = df[df["revolverstatus"] == 0]

#Filter 2 - only takes valid loans
valid_statuses = ["PIF", "CURR", "DELINQ", "PSTDUE", "DEFERD", "CHGOFF", "LIQUID", "PURCH(NOTC/O)"]
df = df[df["loanstatus"].isin(valid_statuses)]
df["event"] = (df["loanstatus"] == "PIF").astype(int)

#Drops missing interest rate data
df = df.dropna(subset=["initialinterestrate"])

#Builds rate buckets (quartiles)
df["rate_bucket"] = pd.qcut(df["initialinterestrate"], q=4)

#Calculates CPR
rate_cpr = df.groupby("rate_bucket").agg(total_loans=("event", "count"), prepays=("event", "sum")).reset_index()
rate_cpr["CPR (%)"] = (rate_cpr["prepays"] / rate_cpr["total_loans"] * 100).round(2)

#Gets quartile cuts and renames buckets with ranges
quantiles = df["initialinterestrate"].quantile([0, 0.25, 0.5, 0.75, 1])
q = quantiles.values
labels = [f"{q[0]:.2f}% - {q[1]:.2f}%",
          f"{q[1]:.2f}% - {q[2]:.2f}%",
          f"{q[2]:.2f}% - {q[3]:.2f}%",
          f"{q[3]:.2f}% - {q[4]:.2f}%"]
rate_cpr["rate_bucket"] = labels

#Displays output
print("\nCPR vs Interest Rate\n")
print(rate_cpr)