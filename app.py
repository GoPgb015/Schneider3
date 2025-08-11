from flask import Flask, render_template, request, jsonify
from pathlib import Path
import pandas as pd

app = Flask(__name__)

# Files
EF_PATH = Path("emission_factors.xlsx")   # Scope 1 & 2
S3_PATH = Path("scope3_library.xlsx")     # Scope 3 drill-down

SCOPES = ["Scope 1", "Scope 2", "Scope 3"]

COUNTRIES = [
    "Afghanistan","Albania","Algeria","Andorra","Angola","Antigua and Barbuda","Argentina","Armenia",
    "Australia","Austria","Azerbaijan","Bahamas","Bahrain","Bangladesh","Barbados","Belarus","Belgium",
    "Belize","Benin","Bhutan","Bolivia","Bosnia and Herzegovina","Botswana","Brazil","Brunei","Bulgaria",
    "Burkina Faso","Burundi","Cabo Verde","Cambodia","Cameroon","Canada","Central African Republic",
    "Chad","Chile","China","Colombia","Comoros","Congo","Costa Rica","Croatia","Cuba","Cyprus","Czechia",
    "Democratic Republic of the Congo","Denmark","Djibouti","Dominica","Dominican Republic","Ecuador",
    "Egypt","El Salvador","Equatorial Guinea","Eritrea","Estonia","Eswatini","Ethiopia","Fiji","Finland",
    "France","Gabon","Gambia","Georgia","Germany","Ghana","Greece","Grenada","Guatemala","Guinea",
    "Guinea-Bissau","Guyana","Haiti","Honduras","Hungary","Iceland","India","Indonesia","Iran","Iraq",
    "Ireland","Israel","Italy","Jamaica","Japan","Jordan","Kazakhstan","Kenya","Kiribati","Kuwait",
    "Kyrgyzstan","Laos","Latvia","Lebanon","Lesotho","Liberia","Libya","Liechtenstein","Lithuania",
    "Luxembourg","Madagascar","Malawi","Malaysia","Maldives","Mali","Malta","Marshall Islands",
    "Mauritania","Mauritius","Mexico","Micronesia","Moldova","Monaco","Mongolia","Montenegro","Morocco",
    "Mozambique","Myanmar","Namibia","Nauru","Nepal","Netherlands","New Zealand","Nicaragua","Niger",
    "Nigeria","North Korea","North Macedonia","Norway","Oman","Pakistan","Palau","Palestine","Panama",
    "Papua New Guinea","Paraguay","Peru","Philippines","Poland","Portugal","Qatar","Romania","Russia",
    "Rwanda","Saint Kitts and Nevis","Saint Lucia","Saint Vincent and the Grenadines","Samoa",
    "San Marino","Sao Tome and Principe","Saudi Arabia","Senegal","Serbia","Seychelles","Sierra Leone",
    "Singapore","Slovakia","Slovenia","Solomon Islands","Somalia","South Africa","South Korea",
    "South Sudan","Spain","Sri Lanka","Sudan","Suriname","Sweden","Switzerland","Syria","Taiwan",
    "Tajikistan","Tanzania","Thailand","Timor-Leste","Togo","Tonga","Trinidad and Tobago","Tunisia",
    "Turkey","Turkmenistan","Tuvalu","Uganda","Ukraine","United Arab Emirates","United Kingdom",
    "United States","Uruguay","Uzbekistan","Vanuatu","Vatican City","Venezuela","Vietnam","Yemen",
    "Zambia","Zimbabwe"
]

# -------- Loaders --------
def load_emission_factors():
    """Generic factors for Scope 1 & 2: columns Title, Unit, Factor."""
    if not EF_PATH.exists():
        return []
    df = pd.read_excel(EF_PATH)
    for c in ("Title", "Unit", "Factor"):
        if c not in df.columns:
            raise ValueError("emission_factors.xlsx must have Title, Unit, Factor")
    df["Title"] = df["Title"].astype(str).str.strip()
    df["Unit"] = df["Unit"].astype(str).str.strip()
    df["Factor"] = pd.to_numeric(df["Factor"], errors="coerce").fillna(0)
    return df.to_dict("records")

def load_scope3_library():
    """Normalize Scope-3 workbook to: Category, Type, Group, Subtype, Unit, Factor."""
    if not S3_PATH.exists():
        return []
    xls = pd.ExcelFile(S3_PATH)
    recs = []

    def fnum(x):
        try: return float(x)
        except Exception: return 0.0

    for sheet in xls.sheet_names:
        df = pd.read_excel(S3_PATH, sheet_name=sheet)
        df.columns = [c.strip() for c in df.columns]

        if set(["Category","Type","ActivityGroup","Subtype","Unit","Factor"]).issubset(df.columns):
            d = df.copy()
            d["Sheet"] = sheet
            d.rename(columns={"ActivityGroup":"Group"}, inplace=True)
            d["Factor"] = d["Factor"].map(fnum)
            recs += d[["Sheet","Category","Type","Group","Subtype","Unit","Factor"]].to_dict("records")
            continue

        if set(["Category","Type","ActivityGroup","Haul","Class","Unit","Factor"]).issubset(df.columns):
            d = df.copy()
            d["Sheet"] = sheet
            d["Group"] = d["ActivityGroup"].astype(str).str.strip()
            d["Subtype"] = d["Haul"].astype(str).str.strip() + " · " + d["Class"].astype(str).str.strip()
            d["Factor"] = d["Factor"].map(fnum)
            recs += d[["Sheet","Category","Type","Group","Subtype","Unit","Factor"]].to_dict("records")
            continue

        if set(["Category","Type","ActivityGroup","Fuel","Unit","Factor"]).issubset(df.columns):
            d = df.copy()
            d["Sheet"] = sheet
            d["Group"] = d["ActivityGroup"].astype(str).str.strip()
            d["Subtype"] = d["Fuel"].astype(str).str.strip()
            d["Factor"] = d["Factor"].map(fnum)
            recs += d[["Sheet","Category","Type","Group","Subtype","Unit","Factor"]].to_dict("records")
            continue

        if set(["Activity","Type","Unit","Year","Factor"]).issubset(df.columns):
            d = df.copy()
            d["Sheet"] = sheet
            d["Category"] = "Fuel & Energy (WTT)"
            d["Group"] = d["Activity"].astype(str).str.strip()
            d["Subtype"] = d["Type"].astype(str).str.strip() + " · " + d["Year"].astype(str)
            d["Type"] = "Upstream"
            d["Factor"] = d["Factor"].map(fnum)
            recs += d[["Sheet","Category","Type","Group","Subtype","Unit","Factor"]].to_dict("records")
            continue

        if set(["Activity","Type","Unit","Factor"]).issubset(df.columns):
            d = df.copy()
            d["Sheet"] = sheet
            d["Category"] = "Water"
            d["Group"] = d["Activity"].astype(str).str.strip()
            d["Subtype"] = d["Type"].astype(str).str.strip()
            d["Type"] = "Upstream"
            d["Factor"] = d["Factor"].map(fnum)
            recs += d[["Sheet","Category","Type","Group","Subtype","Unit","Factor"]].to_dict("records")
            continue

        if "Activity" in df.columns and "Material" in df.columns and "Unit" in df.columns:
            wide = df.copy()
            idv = ["Activity","Material","Unit"]
            long = wide.melt(id_vars=idv, var_name="Subcol", value_name="Factor")
            long = long.dropna(subset=["Factor"])
            long["Sheet"] = sheet
            long["Category"] = "Construction"
            long["Type"] = "Upstream"
            long["Group"] = long["Activity"].astype(str).str.strip()
            long["Subtype"] = long["Material"].astype(str).str.strip() + " · " + long["Subcol"].astype(str).str.strip()
            long["Factor"] = long["Factor"].map(fnum)
            recs += long[["Sheet","Category","Type","Group","Subtype","Unit","Factor"]].to_dict("records")
            continue
    return recs

# -------- Routes --------
@app.route("/")
def index():
    f12 = load_emission_factors()
    s3 = load_scope3_library()
    activities = sorted({r["Title"] for r in f12})
    return render_template(
        "index.html",
        scopes=SCOPES,
        activities=activities,
        factors12=f12,
        scope3=s3,
        countries=COUNTRIES
    )

@app.route("/calculate", methods=["POST"])
def calculate():
    f12 = load_emission_factors()
    s3 = load_scope3_library()
    rows = request.json.get("rows", [])

    map12 = {(r["Title"], r["Unit"]): float(r["Factor"]) for r in f12}
    map3  = {(r["Category"], r["Type"], r["Group"], r["Subtype"], r["Unit"]): float(r["Factor"]) for r in s3}

    totals = {s: 0.0 for s in SCOPES}
    grand = 0.0

    for r in rows:
        scope = r.get("scope")
        qty = float(r.get("quantity") or 0)
        if scope == "Scope 3":
            key = (
                (r.get("category") or "").strip(),
                (r.get("flowType") or "").strip(),       # NOTE: this is the 'Type' in the workbook
                (r.get("group") or "").strip(),
                (r.get("subtype") or "").strip(),
                (r.get("unit") or "").strip()
            )
            factor = map3.get(key, 0.0)
        else:
            key = ((r.get("activity") or "").strip(), (r.get("unit") or "").strip())
            factor = map12.get(key, 0.0)

        emission = qty * factor
        if scope in totals:
            totals[scope] += emission
        grand += emission

    return jsonify({
        "total": round(grand, 4),
        "Scope 1": round(totals["Scope 1"], 4),
        "Scope 2": round(totals["Scope 2"], 4),
        "Scope 3": round(totals["Scope 3"], 4),
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
