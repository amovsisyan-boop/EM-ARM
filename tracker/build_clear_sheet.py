#!/usr/bin/env python3
"""Plain-language shortlist CSV from the tracker workbook (Pipeline tab). Output stays local (PII)."""
import csv, re, sys, datetime, openpyxl
SRC = sys.argv[1]; OUT = sys.argv[2]
wb = openpyxl.load_workbook(SRC)
rows = [r for r in wb["Pipeline"].iter_rows(min_row=2, values_only=True) if (r[10] or "").startswith("W")]

SENIOR = {"Sr EM": "Senior Manager level (target)", "EM": "Manager level (confirm scope)", "Lead (verify)": "Team lead (confirm they manage people)",
          "Director+": "Director or above (only if open to Sr Manager)", "IC": "Individual contributor (probably not a fit)", "Unknown": "Title not known (check LinkedIn first)"}
STAGE = {"Sourced": "Not reviewed yet", "Profile Review": "Profile reviewed only", "Outreach Sent": "Messaged only", "Replied / Interested": "Replied, screen not booked",
         "Recruiter Screen": "Recruiter screen", "HM Screen": "Hiring-manager screen", "CP1 – Coding": "Coding interview (older EM loop)",
         "System Design": "System design", "CP3 – Final Interview": "Final interview", "Offer": "Offer", "Hired": "Hired"}
RISK = {"Relocation/remote": "Relocation or remote-work issue", "Timing": "Timing / availability",
        "Stack gap": "Stack gap (not .NET/Java/C#)", "Retention risk": "Long tenure, hard to move", "Thin EM exp": "Limited EM experience",
        "Comp": "Compensation expectations", "Director-only": "Wants Director role only"}
JUNK = re.compile(r"(pending review from may 22|did we contact .*|if yes, what was the outcome\??|if not, let's reach out.*|let me know if.*|should we talk with him\??|^yes$|^no$|^maybe$|^not interested$|^no reply$|^meeting day.*|^contacted$)", re.I)
SUBS = [(r"\bno prio\b ?-? ?|\bnot (top )?prio\w* ?-? ?|\bNot Prio\b ?-? ?", "Lower priority: "), (r"\bRS\b", "recruiter screen"), (r"\bHM\b", "HM"),
        (r"\bCP1\b", "coding interview"), (r"\bCP3\b", "final interview"),
        (r"Last time was not interested, we can reach out again", "Said no last time; notes say we can try again"),
        (r"Didn't reply last time, can be reached out again", "Did not reply last time; can try again")]
def clean(n):
    parts, seen, out = re.split(r"\s*\|\s*", n or ""), set(), []
    for p in parts:
        p = re.sub(r"\s+", " ", p).strip(" -;")
        if not p or JUNK.search(p): continue
        for a, b in SUBS: p = re.sub(a, b, p, flags=re.I)
        if re.match(r"(those who|let's check interest|let's do (ta|screening)|can we consider|i will need|i know |recommended)", p, re.I): continue
        k = re.sub(r"\W+", "", p.lower())[:32]
        if k in seen or len(k) < 3: continue
        seen.add(k); out.append(p[0].upper() + p[1:])
    s = "; ".join(out)
    s = re.sub(r"(;\s*)+", "; ", s)
    return s if len(s) <= 230 else s[:s.rfind(" ", 0, 227)].rstrip(";, ") + "…"

def stand(r):
    st, stage, wave = r[8], r[7], r[10]
    if wave.startswith("W0"):
        return {"Replied / Interested": "Replied – interested", "Recruiter Screen": "In process – recruiter screen done", "HM Screen": "In process – past hiring-manager screen",
                "CP1 – Coding": "In process – reached coding interview", "System Design": "In process – reached system design",
                "CP3 – Final Interview": "In process – reached final interview"}.get(stage, "In process")
    if wave.startswith("W1"):
        return {"On Hold": "On hold – paused earlier", "Nurture – Re-engage Later": "Warm – timing wasn't right", "Not a Fit (Profile)": "Closed earlier – notes say try again"}.get(st, "Paused earlier")
    if wave.startswith("W3"): return "Messaged – no reply yet" if st != "No Response" else "Messaged – no reply after follow-ups"
    return "Not contacted yet"
def nxt(r):
    w, st = r[10], r[7]
    if w.startswith("W0"):
        n = {"Replied / Interested": "book recruiter screen", "Recruiter Screen": "book hiring-manager screen", "HM Screen": "book technical retrospective",
             "CP1 – Coding": "book system design", "System Design": "book final interviews", "CP3 – Final Interview": "check final-interview outcome"}.get(st, "confirm next step")
        return f"Confirm still interested and available, then {n}", "Mon 5 Oct"
    if w.startswith("W1"): return "Send a personal message with the Growth AI pitch; ask about interest and timing", "Wed 7 Oct"
    if w.startswith("W3"): return "Send follow-up message (use the 4-step follow-up plan)", "Thu 8 Oct"
    return ("Quick hiring-manager look at the profile, then send first message", "Tue 13 Oct") if r[9] != "A" else ("Send first message", "Tue 6 Oct")

out = []
for r in rows:
    detail = clean(r[20]); has_info = bool(r[3] or r[4] or detail)
    if r[10].startswith("W2") and not has_info: continue          # nothing known beyond a URL: stays in the full tracker
    if r[5] == "IC" and r[10].startswith("W2"): continue
    risks = "; ".join(RISK[t.strip()] for t in (r[19] or "").split(",") if t.strip() in RISK)
    na, when = nxt(r)
    role = " @ ".join(x for x in (r[3], r[4]) if x) or "(not recorded – open LinkedIn)"
    if "Status conflict" in (r[23] or ""):
        risks = "; ".join(x for x in (risks, "Records disagree between tabs – verify before contacting") if x)
    out.append([r[9] or "C", r[1], role, SENIOR[r[5]], stand(r), STAGE[r[7]], detail, risks, na, when, "", "", "", r[2]])
pr = {"A": 0, "B": 1, "C": 2}
out.sort(key=lambda x: (pr[x[0]], {"In": 0, "Re": 0}.get(x[4][:2], 1), x[4]))
hdr = ["Priority (A = first)", "Candidate", "Current role", "Seniority fit for Sr Manager role", "Where they stand today", "Furthest step reached", "What we know", "Watch-outs",
       "Next step", "Target date", "Owner", "Date contacted", "Outcome / notes", "LinkedIn"]
with open(OUT, "w", encoding="utf8", newline="") as f:
    w = csv.writer(f, lineterminator="\n"); w.writerow(hdr); w.writerows(out)
print(len(out), "rows"); import collections; print(collections.Counter(x[4] for x in out).most_common())
