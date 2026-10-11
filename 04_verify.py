"""Cross-check Excel formula results against independent pandas/Python logic."""
import sys, pandas as pd
sys.path.insert(0, "src")
from dq_rules import *
X = "excel/MDM_Vendor_DQ_Governance.xlsx"
ok = True
def cmp(name, a, b):
    global ok
    flag = "OK " if a == b else "FAIL"; ok &= (a == b); print(f"{flag} {name}: excel={a} python={b}")

raw = pd.read_csv("data/processed/vendor_master_raw_with_group.csv", dtype=str, keep_default_na=False)
gold = pd.read_csv("data/processed/vendor_master_golden.csv", dtype=str, keep_default_na=False)
for nm, df in (("Raw", raw), ("Golden", gold)):
    xl = pd.read_excel(X, f"Vendor_Master_{nm}", dtype=str, keep_default_na=False)
    py = pd.DataFrame([run_checks(r) for _, r in df.iterrows()])
    for c in CHECKS:
        bad = (xl[c].values != py[c].values).sum()
        cmp(f"{nm} {c} row mismatches", int(bad), 0)
sc = pd.read_excel(X, "DQ_Scorecard", header=3).iloc[:10]
print(sc.iloc[:, [0, 4, 7, 8, 9]].to_string())
dbk = pd.read_excel(X, "Dashboard", header=None)
print(dbk.iloc[2:21, :3].to_string())

# independent invoice logic
ap = pd.read_csv("data/raw/ap_invoices_raw.csv"); m = pd.read_csv("data/processed/vendor_id_map.csv")
g = pd.read_csv("data/processed/vendor_master_golden.csv", dtype=str)
ap["gid"] = ap.Vendor_ID.map(dict(zip(m.Raw_Vendor_ID, m.Golden_Vendor_ID)))
ap["raw_seq"] = ap.groupby(["Vendor_ID", "Invoice_No", "Invoice_Amount_INR"]).cumcount() + 1
ap["g_seq"] = ap.groupby(["gid", "Invoice_No", "Invoice_Amount_INR"], dropna=False).cumcount() + 1
dup = ap.g_seq > 1
inact = ap.gid.map(dict(zip(g.Vendor_ID, g.Status))).eq("Inactive")
dbv = dbk.set_index(0)[[1, 2]]
cmp("Orphans", int(dbv.loc["Orphan vendor ID (not in master)", 1]), int(ap.gid.isna().sum()))
cmp("Orphan amt", int(dbv.loc["Orphan vendor ID (not in master)", 2]), int(ap[ap.gid.isna()].Invoice_Amount_INR.sum()))
cmp("Inactive", int(dbv.loc["Invoices on inactive vendors", 1]), int(inact.sum()))
cmp("Dup count", int(dbv.loc["Duplicate invoices (after MDM)", 1]), int(dup.sum()))
cmp("Dup amt", int(dbv.loc["Duplicate invoices (after MDM)", 2]), int(ap[dup].Invoice_Amount_INR.sum()))
only = dup & (ap.raw_seq == 1)
cmp("Only-after-MDM count", int(dbv.loc["Duplicates found ONLY after vendor merge", 1]), int(only.sum()))
cmp("Only-after-MDM amt", int(dbv.loc["Duplicates found ONLY after vendor merge", 2]), int(ap[only].Invoice_Amount_INR.sum()))
print("ALL OK" if ok else "SOME FAILED")
