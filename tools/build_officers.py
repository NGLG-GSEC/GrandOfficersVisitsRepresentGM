"""Δημιουργεί το officers.js από το Excel «Τάξις και Προβάδισμα».

Χρήση:  python3 tools/build_officers.py <αρχείο.xlsx>

Προβάδισμα: η σειρά των αξιωμάτων ακολουθεί το Excel (Σύνταγμα).
Μέσα στο ίδιο αξίωμα προηγείται ο αρχαιότερος (επετηρίδα: έτος ανώτατου
βαθμού, μικρότερο πρώτα). Όταν ο τίτλος (στ. B) διαφέρει από τον ανώτατο
βαθμό (στ. E), δηλαδή για εν ενεργεία αξιωματικούς, το έτος λαμβάνεται από το
διάταγμα της στ. J (π.χ. «502/2026»). Χωρίς έτος -> στο τέλος του αξιώματος.
Όσοι έχουν «Μετέστη» στη στήλη J παραλείπονται.
"""
import json
import re
import sys

import openpyxl

SAME_TIER = [
    # Επαρχιακός / Περιφερειακός Μ.Δ. έχουν την ίδια θέση στην ιεραρχία
    (re.compile(r"^(Πρώην )?(Επαρχιακός ή Περιφερειακός|Επαρχιακός|Περιφερειακός) Μέγας Διδάσκαλος$"),
     lambda m: (m.group(1) or "") + "Επαρχιακός ή Περιφερειακός Μέγας Διδάσκαλος"),
]


def tier_of(title):
    for rx, fn in SAME_TIER:
        m = rx.match(title)
        if m:
            return fn(m)
    return title


def main(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        aa, title, _, name, highest, _, _, _, year, current = r[:10]
        if not name or not title:
            continue
        if current and "Μετέστη" in str(current):
            continue
        if highest and str(highest).strip() != str(title).strip():
            m = re.search(r"/\s*((?:19|20)\d{2})", str(current or ""))
            year = int(m.group(1)) if m else year
        rows.append({
            "aa": int(aa), "title": str(title).strip(), "name": str(name).strip(),
            "year": int(year) if year else None,
        })

    tier_order = {}
    for row in rows:  # σειρά πρώτης εμφάνισης = σειρά Συντάγματος
        tier_order.setdefault(tier_of(row["title"]), len(tier_order))

    rows.sort(key=lambda x: (tier_order[tier_of(x["title"])],
                             x["year"] is None, x["year"] or 0, x["aa"]))

    out = []
    for i, row in enumerate(rows, 1):
        out.append({"rank": i, "name": row["name"], "title": row["title"],
                    "year": row["year"], "group": tier_of(row["title"])})

    sys.stdout.write("// Παράγεται από tools/build_officers.py — μην το αλλάζετε με το χέρι.\n")
    sys.stdout.write("const OFFICERS = " + json.dumps(out, ensure_ascii=False, indent=0) + ";\n")


if __name__ == "__main__":
    main(sys.argv[1])
