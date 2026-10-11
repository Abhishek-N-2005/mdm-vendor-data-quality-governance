"""Shared data-quality rules and cleaning helpers (Vendor Master)."""
import re

LEGAL = {"pvt", "private", "ltd", "limited", "llp", "co", "company"}
VALID_TERMS = {"Net 15", "Net 30", "Net 45", "Net 60"}


def norm_name(s):
    """Normalised key used for duplicate detection."""
    s = str(s).lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return " ".join(t for t in s.split() if t not in LEGAL and t != "and")


def _upper_letters(s):
    return len(s) > 0 and all("A" <= c <= "Z" for c in s)


def pan_check(v):
    v = (v or "").strip()
    if not v:
        return "Missing"
    ok = len(v) == 10 and _upper_letters(v[:5]) and v[5:9].isdigit() and _upper_letters(v[9])
    return "Valid" if ok else "Invalid"


def gstin_check(g, pan):
    g, pan = (g or "").strip(), (pan or "").strip()
    if not g:
        return "Missing"
    if len(g) == 15 and g[13].upper() == "Z":
        return "Valid" if (pan and g[2:12].upper() == pan.upper()) else "PAN Mismatch"
    return "Invalid"


def email_check(v):
    v = (v or "").strip()
    if not v:
        return "Missing"
    at = v.find("@")
    ok = at >= 0 and " " not in v and v.find(".", at) > 0
    return "Valid" if ok else "Invalid"


def phone_digits(v):
    return (v or "").replace(" ", "").replace("-", "").replace("+", "")


def phone_check(v):
    if not (v or "").strip():
        return "Missing"
    d = phone_digits(v)
    ok = d.isdigit() and (len(d) == 10 or (len(d) == 12 and d[:2] == "91")) and int(d[-10]) >= 6
    return "Valid" if ok else "Invalid"


def bank_check(v):
    v = (v or "").strip()
    if not v:
        return "Missing"
    return "Valid" if v.isdigit() and 9 <= len(v) <= 18 else "Invalid"


def ifsc_check(v):
    v = (v or "").strip()
    if not v:
        return "Missing"
    return "Valid" if len(v) == 11 and _upper_letters(v[:4]) and v[4] == "0" else "Invalid"


def country_check(v):
    v = (v or "").strip()
    if not v:
        return "Missing"
    return "Valid" if v == "India" else "Invalid"


def terms_check(v):
    v = (v or "").strip()
    if not v:
        return "Missing"
    return "Valid" if v in VALID_TERMS else "Invalid"


CHECKS = ["PAN_Check", "GSTIN_Check", "Email_Check", "Phone_Check",
          "Bank_Check", "IFSC_Check", "Country_Check", "Terms_Check"]


def run_checks(r):
    return {
        "PAN_Check": pan_check(r["PAN"]),
        "GSTIN_Check": gstin_check(r["GSTIN"], r["PAN"]),
        "Email_Check": email_check(r["Email"]),
        "Phone_Check": phone_check(r["Phone"]),
        "Bank_Check": bank_check(r["Bank_Account"]),
        "IFSC_Check": ifsc_check(r["IFSC"]),
        "Country_Check": country_check(r["Country"]),
        "Terms_Check": terms_check(r["Payment_Terms"]),
    }


# ---------- cleaning (standardisation only - never invents data) ----------
def clean_name(s):
    s = re.sub(r"\s+", " ", (s or "").strip())
    if s.isupper():
        s = s.title()
    s = re.sub(r"\bPvt\.?\s*Ltd\.?", "Pvt Ltd", s, flags=re.I)
    return s


def clean_phone(v):
    d = phone_digits(v)
    if phone_check(v) == "Valid":
        return "+91 " + d[-10:]
    return (v or "").strip()


def clean_terms(v):
    m = re.search(r"(\d+)", v or "")
    if m and f"Net {int(m.group(1))}" in VALID_TERMS:
        return f"Net {int(m.group(1))}"
    return (v or "").strip()


def clean_country(v):
    return "India" if (v or "").strip().upper() in {"IN", "INDIA"} else (v or "").strip()
