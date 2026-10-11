"""Generate a SYNTHETIC vendor master and AP invoice dataset with realistic data-quality issues.
All names, IDs and amounts are fictional. Seeded for reproducibility."""
import random, string
from datetime import date, timedelta
import pandas as pd

random.seed(42)
PREFIX = ["Apex","Bluewave","Crestline","Delta","Evergreen","Falcon","Greenfield","Horizon","Ironbridge",
          "Jupiter","Kaveri","Lotus","Metro","Nova","Orchid","Pinnacle","Quantum","Redwood","Sunrise","Titan",
          "Unity","Vertex","Western","Zenith","Aaryan","Bharat","Chola","Deccan","Eastern","Ganga"]
INDUSTRY = ["Facility Services","Electricals","Security Solutions","Housekeeping","Engineering Works",
            "Logistics","Printing","Water Solutions","Power Systems","IT Services","Maintenance",
            "Stationers","Pest Control","Landscaping"]
LEGAL = ["Private Limited","Pvt Ltd","LLP","Limited"]
STATES = ["27","29","33","07","24","36","19"]
BANKS = ["HDFC","ICIC","SBIN","UTIB","KKBK"]
L = string.ascii_uppercase

def rl(n): return "".join(random.choice(L) for _ in range(n))
def rd(n): return "".join(random.choice(string.digits) for _ in range(n))

pairs = random.sample([(p, i) for p in PREFIX for i in INDUSTRY], 200)
rows = []
for k, (p, i) in enumerate(pairs):
    pan = rl(3) + "C" + rl(1) + rd(4) + rl(1)
    gst = random.choice(STATES) + pan + "1Z" + random.choice(L + string.digits)
    slug = (p + i).lower().replace(" ", "")
    rows.append(dict(
        Vendor_ID=f"V{1001+k}", Vendor_Name=f"{p} {i} {random.choice(LEGAL)}", PAN=pan, GSTIN=gst,
        Email=f"accounts@{slug}.in", Phone=random.choice("6789") + rd(9), Bank_Account=rd(random.randint(11, 14)),
        IFSC=random.choice(BANKS) + "0" + rd(6), Country="India", Payment_Terms=random.choice(["Net 15","Net 30","Net 45","Net 60"]),
        Status="Inactive" if random.random() < 0.10 else "Active",
        Created_Date=(date(2019,1,1) + timedelta(days=random.randint(0, 2300))).isoformat(), _p=p, _i=i))

base = random.sample(range(200), 28)
def fmt_variant(r):
    n = r["Vendor_Name"].replace("Private Limited", "Pvt. Ltd.").replace("Pvt Ltd", "PVT LTD").replace("LLP", "L.L.P.")
    return n.upper() if random.random() < .5 else n.replace(" ", "  ", 1)
ABBR = {"Services":"Svcs","Engineering":"Engg","Solutions":"Soln","Maintenance":"Maint.","Electricals":"Elec.",
        "Logistics":"Logis.","Printing":"Prtg","Stationers":"Stat."}
def abbr_variant(r):
    n = r["Vendor_Name"]
    for k, v in ABBR.items(): n = n.replace(k, v)
    return n if n != r["Vendor_Name"] else r["Vendor_Name"].replace(r["_p"], r["_p"][:4] + ".")
def typo_variant(r):
    n = r["Vendor_Name"]; j = n.index(r["_i"].split()[0]) + 3
    return n[:j] + n[j+1:]
dups = []
for idx, b in enumerate(base):
    r = dict(rows[b]); r["Vendor_ID"] = f"V{1201+idx}"
    r["Created_Date"] = (date.fromisoformat(r["Created_Date"]) + timedelta(days=random.randint(60, 900))).isoformat()
    if idx < 15: r["Vendor_Name"] = fmt_variant(r)
    elif idx < 23: r["Vendor_Name"] = abbr_variant(r)
    else: r["Vendor_Name"] = typo_variant(r); r["PAN"] = ""; r["GSTIN"] = ""
    if random.random() < .3: r["Phone"] = random.choice("6789") + rd(9)
    dups.append(r)
rows += dups

for r in rows:
    x = random.random()
    if r["PAN"]:
        if x < .07: r["PAN"] = ""
        elif x < .11: r["PAN"] = r["PAN"].lower()
        elif x < .15: r["PAN"] = r["PAN"][:9]
    x = random.random()
    if r["GSTIN"]:
        if x < .08: r["GSTIN"] = ""
        elif x < .12: r["GSTIN"] = r["GSTIN"][:2] + rl(5) + r["GSTIN"][7:]
        elif x < .14: r["GSTIN"] = r["GSTIN"][:12]
    x = random.random()
    if x < .05: r["Email"] = ""
    elif x < .09: r["Email"] = r["Email"].replace("@", "") if random.random() < .5 else r["Email"].replace(".in", "")
    elif x < .25: r["Email"] = r["Email"].upper()
    ph = r["Phone"]; x = random.random()
    if x < .04: r["Phone"] = ""
    elif x < .08: r["Phone"] = ph[:9]
    elif x < .50: r["Phone"] = f"+91 {ph[:5]} {ph[5:]}"
    elif x < .70: r["Phone"] = f"{ph[:5]}-{ph[5:]}"
    if random.random() < .05: r["Bank_Account"] = ""
    x = random.random()
    if x < .05: r["IFSC"] = ""
    elif x < .09: r["IFSC"] = r["IFSC"].lower()
    elif x < .12: r["IFSC"] = r["IFSC"][:4] + "1" + r["IFSC"][5:]
    r["Country"] = random.choices(["India","IN","INDIA","india"], [70,15,10,5])[0]
    t = r["Payment_Terms"]; n = t.split()[1]
    r["Payment_Terms"] = random.choice([t, t, t.upper().replace(" ", ""), f"{n} Days", f"{n}D", t.lower()])
    if random.random() < .05: r["Vendor_Name"] = "  " + r["Vendor_Name"] + " "

twin = {}
for b, d in zip(base, dups):
    twin[f"V{1001+b}"] = d["Vendor_ID"]; twin[d["Vendor_ID"]] = f"V{1001+b}"
random.shuffle(rows)
vm = pd.DataFrame(rows)
cols = ["Vendor_ID","Vendor_Name","PAN","GSTIN","Email","Phone","Bank_Account","IFSC","Country","Payment_Terms","Status","Created_Date"]
vm[cols].to_csv("data/raw/vendor_master_raw.csv", index=False)

ids = vm["Vendor_ID"].tolist()
inv = [dict(Vendor_ID=random.choice(ids)) for _ in range(484)]
inv += [dict(Vendor_ID=f"V{9001+n%5}") for n in range(10)]
for d in inv:
    d["Invoice_Date"] = date(2025,1,1) + timedelta(days=random.randint(0, 272))
    d["Invoice_Amount_INR"] = int(round(random.lognormvariate(11.2, 0.7), -2))
    d["Payment_Status"] = random.choice(["Paid","Paid","Paid","Pending","Pending"])
random.shuffle(inv)
for k, d in enumerate(inv): d["Invoice_No"] = f"INV-2025-{k+1:04d}"
cand = [d for d in inv if d["Vendor_ID"] in twin]
plain = [d for d in inv if not d["Vendor_ID"].startswith("V9") and d["Vendor_ID"] not in twin]
copies = random.sample(cand, 10) + random.sample(plain, 6)
for k, d in enumerate(copies):
    c = dict(d)
    if k < 10: c["Vendor_ID"] = twin[d["Vendor_ID"]]
    c["Invoice_Date"] = d["Invoice_Date"] + timedelta(days=random.randint(0, 5))
    c["Payment_Status"] = random.choice(["Paid", "Pending"])
    inv.append(c)
ap = pd.DataFrame(inv).sort_values(["Invoice_Date", "Invoice_No"], kind="stable")
ap["Invoice_Date"] = ap["Invoice_Date"].astype(str)
ap[["Invoice_No","Vendor_ID","Invoice_Date","Invoice_Amount_INR","Payment_Status"]].to_csv("data/raw/ap_invoices_raw.csv", index=False)
print("vendors", len(vm), "invoices", len(ap))
