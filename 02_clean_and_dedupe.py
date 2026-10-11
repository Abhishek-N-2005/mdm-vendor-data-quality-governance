"""Clean, de-duplicate and build golden vendor records. Standardises formats; never invents missing values."""
from difflib import SequenceMatcher
import pandas as pd
from dq_rules import *

raw = pd.read_csv("data/raw/vendor_master_raw.csv", dtype=str, keep_default_na=False)

# 1) standardise every raw record
cl = raw.copy()
cl["Vendor_Name"] = cl["Vendor_Name"].map(clean_name)
cl["PAN"] = cl["PAN"].str.strip().str.upper()
cl["GSTIN"] = cl["GSTIN"].str.strip().str.upper()
cl["Email"] = cl["Email"].str.strip().str.lower()
cl["Phone"] = cl["Phone"].map(clean_phone)
cl["Bank_Account"] = cl["Bank_Account"].str.strip()
cl["IFSC"] = cl["IFSC"].str.strip().str.upper()
cl["Country"] = cl["Country"].map(clean_country)
cl["Payment_Terms"] = cl["Payment_Terms"].map(clean_terms)
cl["_key"] = cl["Vendor_Name"].map(norm_name)
for c, v in zip(CHECKS, zip(*[run_checks(r).values() for _, r in cl.iterrows()])):
    cl[c] = v
cl["_score"] = (cl[CHECKS] == "Valid").sum(axis=1)

# 2) duplicate groups: same normalised name OR same valid PAN (union-find)
parent = list(range(len(cl)))
def find(a):
    while parent[a] != a:
        parent[a] = parent[parent[a]]; a = parent[a]
    return a
def union(a, b): parent[find(a)] = find(b)
for col, ok in (("_key", lambda s: s != ""), ("PAN", lambda s: pan_check(s) == "Valid")):
    first = {}
    for i, v in enumerate(cl[col]):
        if ok(v):
            union(i, first.setdefault(v, i))
cl["_grp"] = [find(i) for i in range(len(cl))]

# 3) survivorship -> golden record
golden, idmap = [], []
for g, d in cl.groupby("_grp"):
    d = d.sort_values(["_score", "Created_Date", "Vendor_ID"], ascending=[False, True, True])
    s = d.iloc[0].copy()
    def best(col, chk):
        for _, r in d.iterrows():
            if chk(r) == "Valid": return r[col]
        return s[col]
    pan = best("PAN", lambda r: r["PAN_Check"])
    gst = next((r["GSTIN"] for _, r in d.iterrows() if gstin_check(r["GSTIN"], pan) == "Valid"), s["GSTIN"])
    pair = next(((r["Bank_Account"], r["IFSC"]) for _, r in d.iterrows()
                 if r["Bank_Check"] == "Valid" and r["IFSC_Check"] == "Valid"), None)
    bank, ifsc = pair if pair else (best("Bank_Account", lambda r: r["Bank_Check"]), best("IFSC", lambda r: r["IFSC_Check"]))
    gid = s["Vendor_ID"]
    golden.append(dict(
        Vendor_ID=gid, Vendor_Name=s["Vendor_Name"], PAN=pan, GSTIN=gst, Email=best("Email", lambda r: r["Email_Check"]),
        Phone=best("Phone", lambda r: r["Phone_Check"]), Bank_Account=bank, IFSC=ifsc, Country="India",
        Payment_Terms=best("Payment_Terms", lambda r: r["Terms_Check"]),
        Status="Active" if (d["Status"] == "Active").any() else "Inactive", Created_Date=d["Created_Date"].min(),
        Source_Records=len(d)))
    for _, r in d.iterrows():
        rule = "Survivor" if r["Vendor_ID"] == gid else (
            "Normalised name match" if r["_key"] == s["_key"] else
            "PAN match" if r["PAN"] == s["PAN"] and r["PAN"] else "Transitive match")
        idmap.append(dict(Raw_Vendor_ID=r["Vendor_ID"], Golden_Vendor_ID=gid, Match_Rule=rule,
                          Raw_Name=raw.loc[raw.Vendor_ID == r["Vendor_ID"], "Vendor_Name"].iloc[0].strip(), Golden_Name=s["Vendor_Name"]))
gold = pd.DataFrame(golden).sort_values("Vendor_ID").reset_index(drop=True)
imap = pd.DataFrame(idmap).sort_values(["Golden_Vendor_ID", "Match_Rule", "Raw_Vendor_ID"]).reset_index(drop=True)
gchk = pd.DataFrame([run_checks(r) for _, r in gold.iterrows()])
gold = pd.concat([gold, gchk], axis=1)

# 4) fuzzy candidates (NOT auto-merged) + steward queue
fz = []
for i in range(len(gold)):
    for j in range(i + 1, len(gold)):
        a, b = norm_name(gold.Vendor_Name[i]), norm_name(gold.Vendor_Name[j])
        r = SequenceMatcher(None, a, b).ratio()
        if r >= 0.90:
            fz.append(dict(Vendor_A=gold.Vendor_ID[i], Name_A=gold.Vendor_Name[i], PAN_A=gold.PAN[i],
                           Vendor_B=gold.Vendor_ID[j], Name_B=gold.Vendor_Name[j], PAN_B=gold.PAN[j],
                           Similarity=round(r, 3),
                           Recommended_Action=("Likely DIFFERENT vendors - PANs differ, keep separate" if gold.PAN[i] and gold.PAN[j] and gold.PAN[i] != gold.PAN[j]
                                               else "Probable duplicate - steward to confirm (PAN missing on one record)")))
fuzzy = pd.DataFrame(fz, columns=["Vendor_A","Name_A","PAN_A","Vendor_B","Name_B","PAN_B","Similarity","Recommended_Action"])
CRIT = {"PAN_Check", "GSTIN_Check", "Bank_Check", "IFSC_Check"}
q = []
for _, r in gold.iterrows():
    bad = [(c, r[c]) for c in CHECKS if r[c] != "Valid"]
    if bad:
        crit = any(c in CRIT for c, _ in bad)
        q.append(dict(Vendor_ID=r.Vendor_ID, Vendor_Name=r.Vendor_Name, Status=r.Status,
                      Open_Issues="; ".join(f"{c.replace('_Check','')}: {v}" for c, v in bad), Issue_Count=len(bad),
                      Priority="High" if crit else "Medium",
                      Action="Request supporting document / confirm with vendor before next payment run" if crit
                      else "Update contact details in next vendor maintenance cycle"))
steward = pd.DataFrame(q).sort_values(["Priority", "Issue_Count", "Vendor_ID"], ascending=[True, False, True])

# 5) raw group id for the workbook
raw = raw.merge(imap[["Raw_Vendor_ID", "Golden_Vendor_ID"]], left_on="Vendor_ID", right_on="Raw_Vendor_ID").drop(columns="Raw_Vendor_ID")
raw.to_csv("data/processed/vendor_master_raw_with_group.csv", index=False)
gold.drop(columns=CHECKS).to_csv("data/processed/vendor_master_golden.csv", index=False)
imap.to_csv("data/processed/vendor_id_map.csv", index=False)
fuzzy.to_csv("data/processed/duplicate_review_candidates.csv", index=False)
steward.to_csv("data/processed/steward_action_queue.csv", index=False)
print("raw", len(raw), "golden", len(gold), "merged", len(raw) - len(gold), "fuzzy", len(fuzzy), "steward", len(steward))
print(imap.Match_Rule.value_counts().to_dict())
print(fuzzy[["Vendor_A","Vendor_B","Similarity"]].to_string())
