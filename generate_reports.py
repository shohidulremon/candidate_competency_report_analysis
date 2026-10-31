"""
Candidate Competency Report Generator
=====================================
Turns two assessment exports (question-level scores + section summary) into a
branded 3-page PDF competency report for every candidate.

What it does
  1. Reads and cleans both Excel files (tolerant of small header changes).
  2. Cross-checks the files: section totals and phone numbers must agree.
  3. Adjusts each score for integrity:
         Weighted Score = Total Mark x (1 - Cheating Probability % / 100)
  4. Ranks candidates on the weighted score and assigns a recommendation tier.
  5. Draws one PDF per candidate plus an _index.csv summary.

Usage
  python generate_reports.py        # all candidates
  python generate_reports.py 5      # quick test: top 5 only
"""
import math
import re
import sys
from pathlib import Path
from statistics import NormalDist

import pandas as pd
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

# =============================================================================
# SETTINGS - the only block you normally need to edit
# =============================================================================
HERE = Path(__file__).resolve().parent
RAW_FILE = HERE / "sample_data" / "sample_raw_scores.xlsx"
SUMMARY_FILE = HERE / "sample_data" / "sample_summary.xlsx"
OUTPUT_DIR = HERE / "reports"
ASSESSMENT_NAME = "Management Trainee Assessment"
RECOMMENDED_TOP = 50     # ranks 1-50          -> Recommended
PROMISING_TOP = 200      # ranks 51-200        -> Promising, below -> Not Recommended

# =============================================================================
# COMPETENCY FRAMEWORK - which questions belong to which section
# =============================================================================
SECTIONS = [
    dict(key="IK", name="Industry Knowledge", max=10, qs=range(1, 6), subs=[],
         desc="Understanding of core industry practice: quality standards, client specifications, "
              "operational performance and planning fundamentals."),
    dict(key="FK", name="Functional Knowledge", max=20, qs=range(6, 16), subs=[],
         desc="Applied knowledge of the role's functional area: processes, terminology, documentation "
              "and day-to-day decision making."),
    dict(key="NTA", name="Non-Technical Aptitude", max=35, qs=range(16, 51),
         desc="Situational judgement across five behavioural competencies that predict success in "
              "early-career roles.",
         subs=[("Integrity & Ethics", range(16, 23)),
               ("Communication & Collaboration", range(23, 30)),
               ("Accountability & Ownership", range(30, 37)),
               ("Learning Agility", range(37, 44)),
               ("Analytical Reasoning", range(44, 51))]),
    dict(key="EX", name="Excel Practical", max=10, qs=range(51, 56),
         desc="Hands-on spreadsheet tasks covering formulas, sorting and data-quality checks.",
         subs=[("Functions & calculations", [51, 53]),
               ("Navigation & sorting", [52, 54]),
               ("Data quality checks", [55])]),
    dict(key="WR", name="Written Communication", max=15, qs=[56], subs=[],
         desc="A short professional message assessed on clarity of request, impact, specificity "
              "and constructive tone."),
]
QMAX = {q: 2 for q in range(1, 16)}
QMAX.update({q: 1 for q in range(16, 51)})
QMAX.update({q: 2 for q in range(51, 56)})
QMAX[56] = 15
TOTAL_MAX = sum(QMAX.values())  # 90

# The 9 bars of the competency profile chart
PROFILE = [(s["name"], list(s["qs"])) for s in SECTIONS if s["key"] in ("IK", "FK")]
PROFILE += [(n, list(q)) for n, q in SECTIONS[2]["subs"]]
PROFILE += [(s["name"], list(s["qs"])) for s in SECTIONS if s["key"] in ("EX", "WR")]

# =============================================================================
# DATA LOADING & CLEANING
# =============================================================================
def _norm(col):
    return re.sub(r"[^a-z0-9]", "", str(col).lower())


def find_col(df, *candidates, required=True):
    """Find a column even if spacing, case or punctuation differ slightly."""
    lookup = {_norm(c): c for c in df.columns}
    for cand in candidates:
        if _norm(cand) in lookup:
            return lookup[_norm(cand)]
    for cand in candidates:            # prefix match, e.g. 'Cheating Probability' vs '... (%)'
        for k, v in lookup.items():
            if k.startswith(_norm(cand)):
                return v
    if required:
        sys.exit(f"ERROR: column '{candidates[0]}' not found. Columns are: {list(df.columns)}")
    return None


def parse_pct(value):
    """'12%', '12', 12.0 -> 12.0 ; 'Null', blank -> None."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    s = str(value).strip().replace("%", "")
    if s.lower() in ("", "null", "na", "n/a", "none", "-"):
        return None
    try:
        return max(0.0, min(100.0, float(s)))
    except ValueError:
        return None


def _digits(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    d = re.sub(r"\D", "", s)
    if d.startswith("880"):
        d = d[3:]
    return d.lstrip("0")


def phone_for(a, b):
    """Show a phone only when both files agree on it."""
    da, db = _digits(a), _digits(b)
    if not da and not db:
        return "Not provided"
    if da != db:
        return "N/A"
    return f"+880 {da[:4]}-{da[4:]}" if len(da) == 10 else "Not provided"


def load_data():
    for f in (RAW_FILE, SUMMARY_FILE):
        if not f.exists():
            sys.exit(f"ERROR: file not found: {f}\nEdit RAW_FILE / SUMMARY_FILE in the SETTINGS block.")
    raw = pd.read_excel(RAW_FILE)
    summ = pd.read_excel(SUMMARY_FILE)

    r_email, s_email = find_col(raw, "Email"), find_col(summ, "Email")
    raw = raw.dropna(subset=[r_email]).copy()
    raw["_email"] = raw[r_email].astype(str).str.strip().str.lower()
    summ = summ.dropna(subset=[s_email]).copy()
    summ["_email"] = summ[s_email].astype(str).str.strip().str.lower()

    s_total = find_col(summ, "Total Mark (90)", "Total Mark")
    r_total = find_col(raw, "Total Mark (90)", "Total Mark")
    raw = raw.rename(columns={r_total: "_raw_total"})
    summ["_total"] = pd.to_numeric(summ[s_total], errors="coerce").fillna(0)
    summ["_cheat"] = summ[find_col(summ, "Cheating Probability (%)", "Cheating Probability",
                                   "Cheating Probabilty")].map(parse_pct)
    name_col = find_col(summ, "Name", required=False)
    summ["_name"] = summ[name_col] if name_col else None

    # attach question-level detail; for duplicate raw rows keep the one whose total matches
    qcols = {q: find_col(raw, f"question_{q}") for q in QMAX}
    keep = ["_email", "_raw_total"] + list(qcols.values())
    rphone = find_col(raw, "Phone Number", "Phone", required=False)
    if rphone:
        keep.append(rphone)
    df = summ.merge(raw[keep], on="_email", how="left", suffixes=("", "_raw"))
    df["_match"] = (df["_raw_total"] - df["_total"]).abs() < 1e-6
    df = df.sort_values("_match", ascending=False).drop_duplicates("_email")
    missing = df[qcols[1]].isna().sum()
    if missing:
        print(f"WARNING: {missing} candidate(s) in the summary have no question-level data; skipped.")
        df = df[df[qcols[1]].notna()]
    for q, c in qcols.items():
        df[f"q{q}"] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    sphone = find_col(summ, "Phone Number", "Phone", required=False)
    rp_col = (rphone + "_raw") if rphone and rphone + "_raw" in df.columns else rphone
    df["_phone"] = [phone_for(a, b) for a, b in
                    zip(df[sphone] if sphone else [None] * len(df),
                        df[rp_col] if rp_col else [None] * len(df))]

    # section scores from the question data, cross-checked against the summary file
    mism = 0
    for s in SECTIONS:
        df["S_" + s["key"]] = df[[f"q{q}" for q in s["qs"]]].sum(axis=1)
        scol = find_col(summ, f"{s['name']} ({s['max']})", s["name"], required=False)
        if scol:
            mism += int(((pd.to_numeric(df[scol], errors="coerce") - df["S_" + s["key"]]).abs() > 1e-6).sum())
    if mism:
        print(f"WARNING: {mism} section total(s) differ between the two files; question data was used.")
    return df.reset_index(drop=True)


def has(v):
    """True when a value is present (not blank/None/NaN)."""
    return v is not None and not (isinstance(v, float) and math.isnan(v))


def score(df):
    df["_cheat"] = [v if has(v) else None for v in df["_cheat"]]
    df["_weighted"] = [t * (1 - ch / 100) if has(ch) else None for t, ch in zip(df["_total"], df["_cheat"])]
    ranked = pd.Series([has(v) for v in df["_weighted"]], index=df.index)
    df["_rank"] = None
    df.loc[ranked, "_rank"] = df.loc[ranked, "_weighted"].rank(ascending=False, method="min").astype(int)

    def tier(rank):
        if not has(rank):
            return "Not Ranked"
        return "Recommended" if rank <= RECOMMENDED_TOP else \
            "Promising" if rank <= PROMISING_TOP else "Not Recommended"
    df["_reco"] = df["_rank"].map(tier)
    return df, int(ranked.sum())


# =============================================================================
# DRAWING
# =============================================================================
W, H = A4
M = 40
CW = W - 2 * M
NAVY = HexColor("#1E3A8A"); NAVY_SOFT = HexColor("#DBE4F7"); INK = HexColor("#1F2937")
MUTED = HexColor("#6B7280"); LINE = HexColor("#E5E7EB"); SOFT = HexColor("#F5F7FB")
GREEN = HexColor("#16A34A"); GREEN_BAR = HexColor("#86D07A"); LOST = HexColor("#F3C1BB")
AMBER = HexColor("#D97706"); RED = HexColor("#DC2626"); GREY = HexColor("#9CA3AF")
F, FB = "Helvetica", "Helvetica-Bold"
TIER_COLOR = {"Recommended": GREEN, "Promising": AMBER, "Not Recommended": RED, "Not Ranked": GREY}


def fmt(x):
    return "%g" % round(float(x), 2)


def wrap(text, font, size, width):
    lines, cur = [], ""
    for w in text.split():
        t = (cur + " " + w).strip()
        if stringWidth(t, font, size) <= width:
            cur = t
        else:
            lines.append(cur)
            cur = w
    return lines + ([cur] if cur else [])


def text(c, x, y, s, font=F, size=10, color=INK, align="left"):
    c.setFont(font, size)
    c.setFillColor(color)
    {"left": c.drawString, "right": c.drawRightString, "center": c.drawCentredString}[align](x, y, s)


def para(c, x, y, s, width, size=9, color=MUTED, leading=None, font=F):
    leading = leading or size * 1.4
    for line in wrap(s, font, size, width):
        text(c, x, y, line, font, size, color)
        y -= leading
    return y


def pill(c, x, y, label, color, align="right", size=9):
    w = stringWidth(label, FB, size) + 18
    x0 = x - w if align == "right" else x
    c.setFillColor(color)
    c.roundRect(x0, y - 5, w, 18, 9, stroke=0, fill=1)
    text(c, x0 + w / 2, y + 0.5, label, FB, size, white, "center")


def split_bar(c, x, y, w, h, frac):
    frac = max(0.0, min(1.0, frac))
    c.setFillColor(LOST)
    c.roundRect(x, y, w, h, h / 2, stroke=0, fill=1)
    if frac > 0:
        c.setFillColor(GREEN_BAR)
        c.roundRect(x, y, max(h, w * frac), h, h / 2, stroke=0, fill=1)


def heading(c, y, title):
    text(c, M, y, title, FB, 13, NAVY)
    c.setStrokeColor(LINE)
    c.setLineWidth(1)
    c.line(M, y - 7, W - M, y - 7)
    return y - 26


def header_footer(c, page, n_pages):
    text(c, M, H - 52, "Candidate Competency Report", FB, 20, NAVY)
    text(c, W - M, H - 52, ASSESSMENT_NAME, F, 9.5, MUTED, "right")
    c.setStrokeColor(NAVY)
    c.setLineWidth(2)
    c.line(M, H - 68, W - M, H - 68)
    c.setStrokeColor(LINE)
    c.setLineWidth(0.5)
    c.line(M, 40, W - M, 40)
    text(c, M, 27, "Confidential - for recruitment use only", F, 8, MUTED)
    text(c, W - M, 27, f"Page {page} of {n_pages}", F, 8, MUTED, "right")


def page_one(c, r, n_ranked, cohort):
    y = H - 100
    # candidate card
    c.setFillColor(SOFT)
    c.roundRect(M, y - 52, CW, 62, 8, stroke=0, fill=1)
    name = r["_name"] if isinstance(r["_name"], str) and r["_name"].strip() else None
    text(c, M + 16, y - 14, name or r["_email"], FB, 17, INK)
    sub = f"{r['_email']}   |   {r['_phone']}" if name else r["_phone"]
    text(c, M + 16, y - 34, sub, F, 9.5, MUTED)
    pill(c, W - M - 16, y - 22, r["_reco"], TIER_COLOR[r["_reco"]])
    y -= 82

    # KPI tiles
    ranked = has(r["_rank"])
    kpis = [("Total Score", f"{fmt(r['_total'])} / {TOTAL_MAX}"),
            ("Cheating Probability", f"{fmt(r['_cheat'])}%" if has(r["_cheat"]) else "No data"),
            ("Weighted Score", f"{r['_weighted']:.2f}" if ranked else "-"),
            ("Overall Rank", f"{int(r['_rank'])} of {n_ranked}" if ranked else "Not ranked")]
    tw, gap = (CW - 3 * 10) / 4, 10
    for i, (label, value) in enumerate(kpis):
        x = M + i * (tw + gap)
        c.setStrokeColor(LINE)
        c.setFillColor(white)
        c.roundRect(x, y - 50, tw, 56, 6, stroke=1, fill=1)
        text(c, x + 12, y - 12, label.upper(), FB, 7.5, MUTED)
        text(c, x + 12, y - 36, value, FB, 15, NAVY)
    y -= 82

    # bell curve
    y = heading(c, y, "Position among candidates")
    ch, x0 = 92, M + 10
    base = y - ch
    pdf = lambda z: math.exp(-z * z / 2)
    xs = [-3 + i * 0.05 for i in range(121)]
    px = lambda z: x0 + (z + 3) / 6 * (CW - 20)
    py = lambda z: base + pdf(z) * (ch - 14)
    if ranked:
        frac = 1 - (int(r["_rank"]) - 0.5) / n_ranked
        zc = max(-2.9, min(2.9, NormalDist().inv_cdf(frac)))
        p = c.beginPath()
        p.moveTo(px(-3), base)
        for z in xs:
            if z > zc:
                break
            p.lineTo(px(z), py(z))
        p.lineTo(px(zc), py(zc))
        p.lineTo(px(zc), base)
        p.close()
        c.setFillColor(NAVY_SOFT)
        c.drawPath(p, stroke=0, fill=1)
    c.setStrokeColor(NAVY if ranked else GREY)
    c.setLineWidth(1.6)
    p = c.beginPath()
    p.moveTo(px(-3), py(-3))
    for z in xs:
        p.lineTo(px(z), py(z))
    c.drawPath(p, stroke=1, fill=0)
    c.setStrokeColor(LINE)
    c.setLineWidth(1)
    c.line(x0, base, x0 + CW - 20, base)
    if ranked:
        c.setStrokeColor(NAVY)
        c.setDash(3, 2)
        c.line(px(zc), base, px(zc), base + ch - 4)
        c.setDash()
        label = f"This candidate  |  Rank {int(r['_rank'])} of {n_ranked}"
        lw = stringWidth(label, FB, 8.5) + 14
        lx = max(M, min(W - M - lw, px(zc) - lw / 2))
        c.setFillColor(NAVY)
        c.roundRect(lx, base + ch - 2, lw, 15, 4, stroke=0, fill=1)
        text(c, lx + lw / 2, base + ch + 2.5, label, FB, 8.5, white, "center")
    else:
        text(c, W / 2, base + ch / 2 - 30, "Not ranked - integrity data unavailable", FB, 10, GREY, "center")
    text(c, x0, base - 13, "Lower weighted score", F, 8, MUTED)
    text(c, x0 + CW - 20, base - 13, "Higher weighted score", F, 8, MUTED, "right")
    y = base - 42

    # section summary
    y = heading(c, y, "Section summary")
    for s in SECTIONS:
        v = r["S_" + s["key"]]
        text(c, M, y, s["name"], FB, 10)
        text(c, M + 196, y, f"{fmt(v)} / {s['max']}", F, 10, INK, "right")
        split_bar(c, M + 214, y - 1, CW - 290, 9, v / s["max"])
        text(c, W - M, y, f"{round(v / s['max'] * 100)}%", FB, 10, NAVY, "right")
        text(c, M + 214, y - 13, f"Cohort average {cohort['S_' + s['key']]:.1f}", F, 7.5, MUTED)
        y -= 31

    y -= 6
    y = heading(c, y, "How to read this report")
    para(c, M, y,
         "Scores are shown as marks obtained against the maximum for each section. The overall rank uses "
         "the weighted score, which reduces the total mark in proportion to the proctoring system's "
         f"cheating probability. Candidates ranked 1-{RECOMMENDED_TOP} are Recommended, "
         f"{RECOMMENDED_TOP + 1}-{PROMISING_TOP} are Promising, and the rest Not Recommended. "
         "Candidates without integrity data are not ranked.", CW, 9)


def page_two(c, r, cohort):
    y = H - 100
    y = heading(c, y, "Anti-cheating monitor")
    c.setFillColor(SOFT)
    c.roundRect(M, y - 112, CW, 122, 8, stroke=0, fill=1)
    ch = r["_cheat"]
    if not has(ch):
        level, col = "No integrity data", GREY
    elif ch < 20:
        level, col = "Low risk", GREEN
    elif ch < 50:
        level, col = "Moderate risk", AMBER
    else:
        level, col = "High risk", RED
    text(c, M + 16, y - 12, "Cheating probability", FB, 10)
    pill(c, W - M - 16, y - 14, level, col)
    bx, bw, by = M + 16, CW - 32, y - 62
    for lo, hi, cc in ((0, 20, GREEN), (20, 50, AMBER), (50, 100, RED)):
        c.setFillColor(cc)
        c.setFillAlpha(0.25)
        c.rect(bx + bw * lo / 100, by, bw * (hi - lo) / 100, 10, stroke=0, fill=1)
    c.setFillAlpha(1)
    for t in (0, 20, 50, 100):
        text(c, bx + bw * t / 100, by - 12, f"{t}%", F, 7.5, MUTED, "center")
    if has(ch):
        mx = bx + bw * ch / 100
        c.setFillColor(INK)
        p = c.beginPath()
        p.moveTo(mx, by + 11)
        p.lineTo(mx - 5, by + 19)
        p.lineTo(mx + 5, by + 19)
        p.close()
        c.drawPath(p, stroke=0, fill=1)
        text(c, mx, by + 22, f"{fmt(ch)}%", FB, 9, INK, "center")
        expl = (f"Weighted score = {fmt(r['_total'])} x (1 - {fmt(ch)}%) = {r['_weighted']:.2f}. "
                "The weighted score, not the raw total, determines rank and recommendation.")
    else:
        expl = ("No cheating probability was recorded for this candidate, so a weighted score cannot be "
                "calculated and the candidate is not ranked.")
    para(c, M + 16, by - 34, expl, CW - 32, 9, INK)
    y -= 146

    y = heading(c, y, "Competency profile")
    # legend
    c.setFillColor(NAVY)
    c.rect(M + 190, y - 2, 10, 8, stroke=0, fill=1)
    text(c, M + 204, y - 1, "This candidate", F, 8, MUTED)
    c.setStrokeColor(AMBER)
    c.setLineWidth(2.5)
    c.line(M + 290, y - 3, M + 290, y + 7)
    text(c, M + 296, y - 1, "Cohort average", F, 8, MUTED)
    y -= 24
    lx, bw = M + 190, CW - 230
    for name, qs in PROFILE:
        mx = sum(QMAX[q] for q in qs)
        v = sum(r[f"q{q}"] for q in qs)
        avg = cohort["P_" + name]
        text(c, M, y, name, F, 9.5)
        c.setFillColor(SOFT)
        c.rect(lx, y - 2, bw, 12, stroke=0, fill=1)
        c.setFillColor(NAVY)
        c.rect(lx, y - 2, bw * v / mx, 12, stroke=0, fill=1)
        c.setStrokeColor(AMBER)
        c.setLineWidth(2.5)
        c.line(lx + bw * avg / mx, y - 5, lx + bw * avg / mx, y + 13)
        text(c, W - M, y, f"{round(v / mx * 100)}%", FB, 9.5, NAVY, "right")
        y -= 32
    for g in (0, 25, 50, 75, 100):
        text(c, lx + bw * g / 100, y + 8, f"{g}%", F, 7, MUTED, "center")


def score_card(c, y, r, s, cohort):
    sub_h = 22 * len(s["subs"])
    desc = wrap(s["desc"], F, 9, CW - 32)
    h = 62 + 12 * len(desc) + sub_h
    c.setStrokeColor(LINE)
    c.setFillColor(white)
    c.roundRect(M, y - h, CW, h, 8, stroke=1, fill=1)
    c.setFillColor(NAVY)
    c.rect(M, y - h + 8, 3, h - 16, stroke=0, fill=1)
    v = r["S_" + s["key"]]
    text(c, M + 16, y - 22, s["name"], FB, 12)
    text(c, W - M - 16, y - 22, f"{fmt(v)} / {s['max']}", FB, 12, NAVY, "right")
    yy = para(c, M + 16, y - 40, s["desc"], CW - 32, 9)
    split_bar(c, M + 16, yy - 4, CW - 120, 9, v / s["max"])
    text(c, W - M - 16, yy - 4, f"Avg {cohort['S_' + s['key']]:.1f}", F, 8.5, MUTED, "right")
    yy -= 22
    for name, qs in s["subs"]:
        mx = sum(QMAX[q] for q in qs)
        sv = sum(r[f"q{q}"] for q in qs)
        text(c, M + 28, yy, name, F, 9)
        text(c, M + 230, yy, f"{fmt(sv)} / {mx}", F, 9, INK, "right")
        split_bar(c, M + 244, yy, CW - 276, 6, sv / mx)
        yy -= 22
    return y - h - 10


def build(path, r, n_ranked, cohort):
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setTitle(f"Competency Report - {r['_name'] if isinstance(r['_name'], str) else r['_email']}")
    header_footer(c, 1, 3); page_one(c, r, n_ranked, cohort); c.showPage()
    header_footer(c, 2, 3); page_two(c, r, cohort); c.showPage()
    header_footer(c, 3, 3)
    y = heading(c, H - 100, "Competency score cards")
    for s in SECTIONS:
        y = score_card(c, y, r, s, cohort)
    c.save()


def safe(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")[:60]


# =============================================================================
# RUN
# =============================================================================
if __name__ == "__main__":
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None
    df, n_ranked = score(load_data())
    cohort = {"S_" + s["key"]: df["S_" + s["key"]].mean() for s in SECTIONS}
    for name, qs in PROFILE:
        cohort["P_" + name] = df[[f"q{q}" for q in qs]].sum(axis=1).mean()

    df["_order"] = df["_rank"].fillna(10 ** 9).astype(float)
    df = df.sort_values(["_order", "_total"], ascending=[True, False])
    if limit:
        df = df.head(limit)
    OUTPUT_DIR.mkdir(exist_ok=True)
    index = []
    for i, (_, r) in enumerate(df.iterrows(), 1):
        rank = f"{int(r['_rank']):03d}" if has(r["_rank"]) else "NR"
        who = r["_name"] if isinstance(r["_name"], str) and r["_name"].strip() else r["_email"].split("@")[0]
        fname = f"{rank}_{safe(who)}.pdf"
        build(OUTPUT_DIR / fname, r, n_ranked, cohort)
        index.append({"file": fname, "email": r["_email"], "name": r["_name"], "phone": r["_phone"],
                      "total_out_of_90": r["_total"], "cheating_pct": r["_cheat"],
                      "weighted_score": round(r["_weighted"], 2) if has(r["_weighted"]) else None,
                      "rank": int(r["_rank"]) if has(r["_rank"]) else None, "recommendation": r["_reco"]})
        if i % 25 == 0 or i == len(df):
            print(f"  {i}/{len(df)} reports done")
    pd.DataFrame(index).to_csv(OUTPUT_DIR / "_index.csv", index=False)
    print(f"Done. {len(df)} PDFs + _index.csv saved to {OUTPUT_DIR}")
