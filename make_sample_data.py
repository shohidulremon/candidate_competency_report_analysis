"""
make_sample_data.py
Creates two FAKE Excel files that mimic an online assessment export, so the
report generator can be demonstrated without any real candidate data.

Every name, email and phone number here is invented:
  - emails use example.com, a domain reserved for documentation
  - phone numbers use the 012 prefix, which is not assigned to any
    Bangladeshi mobile operator

Run:  python make_sample_data.py
"""
import math
import random
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

SEED = 42
N_CANDIDATES = 300
OUT = Path(__file__).resolve().parent / "sample_data"

FIRST = ["Arif", "Nusrat", "Tanvir", "Farhana", "Rakib", "Sadia", "Imran", "Tasnim", "Mahin",
         "Rumana", "Shakil", "Anika", "Fahim", "Lamia", "Sabbir", "Mithila", "Zahid", "Priya",
         "Rafsan", "Tahmina", "Nayeem", "Sumaiya", "Ashik", "Jannat", "Rifat", "Sharmin"]
LAST = ["Rahman", "Hossain", "Ahmed", "Islam", "Chowdhury", "Karim", "Akter", "Sarker", "Uddin",
        "Haque", "Das", "Begum", "Kabir", "Mollah", "Siddique", "Talukder", "Roy", "Mia"]

# question -> maximum mark (56 questions, 90 marks in total)
QMAX = {q: 2 for q in range(1, 16)}            # Industry + Functional knowledge
QMAX.update({q: 1 for q in range(16, 51)})     # Non-technical aptitude
QMAX.update({q: 2 for q in range(51, 56)})     # Excel practical
QMAX[56] = 15                                  # Written communication

SECTIONS = [("Industry Knowledge (10)", range(1, 6)),
            ("Functional Knowledge (20)", range(6, 16)),
            ("Non-Technical Aptitude (35)", range(16, 51)),
            ("Excel Practical (10)", range(51, 56)),
            ("Written Communication (15)", [56])]


def main():
    rnd = random.Random(SEED)
    used, raw_rows, sum_rows = set(), [], []
    for i in range(1, N_CANDIDATES + 1):
        first, last = rnd.choice(FIRST), rnd.choice(LAST)
        email = f"{first.lower()}.{last.lower()}{i:03d}@example.com"
        phone = "012" + "".join(str(rnd.randint(0, 9)) for _ in range(8))
        ability = rnd.gauss(0, 1)

        scores = {}
        for q, mx in QMAX.items():
            if q == 56:      # essay marked out of 15 in half marks
                mark = round(max(0, min(15, 8 + 3 * ability + rnd.gauss(0, 2))) * 2) / 2
            else:
                difficulty = rnd.uniform(-1.5, 1.5)
                p = 1 / (1 + math.exp(-(1.2 * ability - difficulty)))
                mark = sum(1 for _ in range(mx) if rnd.random() < p)
            scores[f"question_{q}"] = mark
        total = sum(scores.values())

        # integrity: most candidates low risk, a few high, ~10% missing
        r = rnd.random()
        if r < 0.10:
            cheat = None
        elif r < 0.85:
            cheat = round(rnd.uniform(0, 20), 1)
        else:
            cheat = round(rnd.uniform(20, 85), 1)

        name = f"{first} {last}" if rnd.random() < 0.8 else None    # some exports lack names
        raw_phone = phone
        sum_phone = phone
        if rnd.random() < 0.03:
            sum_phone = "012" + "".join(str(rnd.randint(0, 9)) for _ in range(8))  # mismatch -> N/A
        if rnd.random() < 0.02:
            raw_phone = sum_phone = "Null"

        raw_rows.append({"Email": email, "Name": name, "Phone Number": raw_phone,
                         **scores, "Total Mark (90)": total})
        row = {"Email": email, "Name": name, "Phone Number": sum_phone}
        for col, qs in SECTIONS:
            row[col] = sum(scores[f"question_{q}"] for q in qs)
        row["Total Mark (90)"] = total
        row["Cheating Probability (%)"] = cheat if cheat is not None else "Null"
        sum_rows.append(row)

    OUT.mkdir(exist_ok=True)
    files = {"sample_raw_scores.xlsx": pd.DataFrame(raw_rows),
             "sample_summary.xlsx": pd.DataFrame(sum_rows)}
    for fname, df in files.items():
        path = OUT / fname
        df.to_excel(path, index=False, sheet_name="Results")
        style(path)
        print(f"Created {path.name}: {len(df)} fake candidates")


def style(path):
    """Header styling, Arial font, sensible column widths, frozen header row."""
    wb = load_workbook(path)
    ws = wb.active
    ws["A1"].parent.freeze_panes = "B2"
    head_fill = PatternFill("solid", start_color="1E3A8A")
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="Arial", size=10)
    for c in ws[1]:
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        head = str(col[0].value)
        width = 34 if head == "Email" else 20 if head in ("Name", "Phone Number") else \
            12 if head.startswith("question") else 16
        ws.column_dimensions[col[0].column_letter].width = width
    ws.row_dimensions[1].height = 32
    wb.save(path)


if __name__ == "__main__":
    main()
