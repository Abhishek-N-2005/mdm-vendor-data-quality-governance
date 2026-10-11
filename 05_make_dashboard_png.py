"""Create the dashboard preview image from the recalculated workbook values."""
import pandas as pd, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
X = "excel/MDM_Vendor_DQ_Governance.xlsx"
sc = pd.read_excel(X, "DQ_Scorecard", header=3).iloc[:8]
d = pd.read_excel(X, "Dashboard", header=None).set_index(0)
v = lambda k, c=1: d.loc[k, c]
fig = plt.figure(figsize=(14, 8), facecolor="white")
fig.suptitle("Vendor Master & AP Data Governance - Finance Operations (synthetic data)", fontsize=16, fontweight="bold", color="#1F3864", x=0.02, ha="left")
kp = [("Raw vendor records", f"{int(v('Raw vendor records'))}"), ("Golden vendors", f"{int(v('Golden (unique) vendors'))}"),
      ("Duplicates merged", f"{int(v('Duplicate records merged'))}"),
      ("Avg DQ pass %", f"{v('Avg DQ pass % (before)')*100:.0f}% -> {v('Avg DQ pass % (after)')*100:.0f}%"),
      ("Duplicate invoices", f"{int(v('Duplicate invoices (after MDM)'))}  (Rs {v('Duplicate invoices (after MDM)',2)/1e5:.1f} L)"),
      ("Orphan-vendor invoices", f"{int(v('Orphan vendor ID (not in master)'))}  (Rs {v('Orphan vendor ID (not in master)',2)/1e5:.1f} L)")]
for i, (a, b) in enumerate(kp):
    ax = fig.add_axes([0.02 + i * 0.163, 0.80, 0.15, 0.12]); ax.axis("off")
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, color="#EAF1FB", transform=ax.transAxes))
    ax.text(.5, .62, b, ha="center", va="center", fontsize=13, fontweight="bold", color="#1F3864", transform=ax.transAxes)
    ax.text(.5, .22, a, ha="center", va="center", fontsize=9, color="#404040", transform=ax.transAxes)
a1 = fig.add_axes([0.06, 0.10, 0.50, 0.58])
x = range(len(sc)); w = .38
a1.bar([i - w/2 for i in x], sc["Pass % (Before)"] * 100, w, label="Before", color="#A6A6A6")
a1.bar([i + w/2 for i in x], sc["Pass % (After)"] * 100, w, label="After", color="#2E75B6")
a1.set_xticks(list(x)); a1.set_xticklabels(sc["Rule"], rotation=25, ha="right", fontsize=9)
a1.set_ylim(0, 105); a1.set_ylabel("Pass %"); a1.set_title("Data quality pass rate by field", loc="left", fontweight="bold")
a1.legend(frameon=False); [a1.spines[s].set_visible(False) for s in ("top", "right")]
a2 = fig.add_axes([0.66, 0.10, 0.31, 0.58])
labs = ["Orphan vendor ID", "Inactive vendors", "Duplicate invoices", "  already Paid", "  still Pending", "Dup. found only\nafter vendor merge"]
keys = ["Orphan vendor ID (not in master)", "Invoices on inactive vendors", "Duplicate invoices (after MDM)", "  of which duplicate already Paid", "  of which still Pending (stoppable)", "Duplicates found ONLY after vendor merge"]
vals = [d.loc[k, 2] / 1e5 for k in keys]
a2.barh(labs[::-1], vals[::-1], color="#C00000")
for i, t in enumerate(vals[::-1]): a2.text(t + .1, i, f"{t:.1f}", va="center", fontsize=9)
a2.set_xlabel("Amount (Rs lakh)"); a2.set_title("AP exceptions by amount", loc="left", fontweight="bold")
[a2.spines[s].set_visible(False) for s in ("top", "right")]
fig.savefig("images/dashboard_overview.png", dpi=150, bbox_inches="tight")
