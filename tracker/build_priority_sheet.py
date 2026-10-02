#!/usr/bin/env python3
"""v3 shortlist: re-ranked by the shortlisting rule (backend/distributed + hands-on AI first, strong leaders from other
stacks kept). Signals come ONLY from existing notes/titles/companies; everything else is 'check on LinkedIn'.
Usage: build_priority_sheet.py <tracker.xlsx> <out.csv>   (output contains candidate data: keep out of git)"""
import csv, re, sys, collections, openpyxl
SRC, OUT = sys.argv[1], sys.argv[2]
rows = [r for r in openpyxl.load_workbook(SRC)["Pipeline"].iter_rows(min_row=2, values_only=True) if (r[10] or "").startswith("W")]

SENIOR = {"Sr EM": "Senior Manager level (target)", "EM": "Manager level (confirm scope)", "Lead (verify)": "Team lead (confirm they manage people)",
          "Director+": "Director or above (only if open to Sr Manager)", "IC": "Individual contributor (probably not a fit)", "Unknown": "Title not known (check LinkedIn first)"}
STAGE = {"Sourced": "Not reviewed yet", "Profile Review": "Profile review", "Outreach Sent": "Messaged", "Replied / Interested": "Replied – screen not booked",
         "Recruiter Screen": "Recruiter screen done", "Screen Round": "Screen Round", "Full Loop": "Full Loop", "Post Loop": "Post Loop", "Offer": "Offer", "Hired": "Hired"}
JUNK = re.compile(r"(pending review from may 22|did we contact .*|if yes, what was the outcome\??|if not, let's reach out.*|let me know if.*|should we talk with him\??|^yes$|^no$|^maybe$|^not interested$|^no reply$|^meeting day.*|^contacted$)", re.I)
SUBS = [(r"\bno prio\b ?-? ?|\bnot (top )?prio\w* ?-? ?|\bNot Prio\b ?-? ?", "Lower priority: "), (r"\bRS\b", "recruiter screen"), (r"\bCP ?(1|I)\b", "coding interview (older loop)"),
        (r"Last time was not interested, we can reach out again", "Said no last time; notes say we can try again"),
        (r"Didn't reply last time, can be reached out again", "Did not reply last time; can try again")]
def clean(n, cap=170):
    seen, out = set(), []
    for p in re.split(r"\s*\|\s*", n or ""):
        p = re.sub(r"\s+", " ", p).strip(" -;")
        if not p or JUNK.search(p) or re.match(r"(those who|let's check interest|let's do (ta|screening)|can we consider|i will need|i know |recommended)", p, re.I): continue
        for a, b in SUBS: p = re.sub(a, b, p, flags=re.I)
        k = re.sub(r"\W+", "", p.lower())[:32]
        if k in seen or len(k) < 3: continue
        seen.add(k); out.append(p[0].upper() + p[1:])
    s = re.sub(r"(;\s*)+", "; ", "; ".join(out))
    return s if len(s) <= cap else s[:s.rfind(" ", 0, cap - 3)].rstrip(";, ") + "…"

BACK = re.compile(r"backend|back-end|distributed|microservice|high-?load|platform|\.net|c#|java\b|aws|azure|event-driven|cloud|enterprise (product|architecture)|saas|data platform|\bgo\b|\bnode", re.I)
FRONT = re.compile(r"\bfe\b|front-?end|\bqa\b|software quality|test engineering|ios|android|mobile|ruby", re.I)
OTHER = re.compile(r"c\+\+|eda\b|embedded|sip/rtp|\bpython\b|golang|js focused|js/node", re.I)
AI = re.compile(r"\bai\b|\bml\b|machine learning|llm|genai|automation|data/ai|agent", re.I)
AI_CO = re.compile(r"krisp|superannotate|podcastle|veeam", re.I)

def signals(r):
    txt = " ".join(str(x or "") for x in (r[3], r[4], r[20]))
    be = "Unlikely (from notes)" if FRONT.search(txt) else "Likely (from notes)" if BACK.search(txt) else "Different stack (from notes)" if OTHER.search(txt) else "Unknown – check LinkedIn"
    ai = "Possible (from notes/company)" if (AI.search(txt) or AI_CO.search(txt)) else "Unknown – check LinkedIn"
    return be, ai, txt

def score(r, be, ai):
    s, why = 0, []
    wave = r[10]
    if wave.startswith("W0"): s += 5; why.append("in process")
    elif wave.startswith("W1"): s += 2; why.append("warm earlier")
    elif wave.startswith("W3"): s += 1
    else: s += 1
    lv = {"Sr EM": 3, "EM": 2, "Lead (verify)": 1, "Director+": 0, "Unknown": 0, "IC": -2}[r[5]]; s += lv
    if lv >= 2: why.append("manager level")
    if be.startswith("Likely"): s += 2; why.append("backend signal")
    elif be.startswith("Unlikely"): s -= 2; why.append("FE/QA/mobile")
    elif be.startswith("Different"): s -= 1; why.append("other stack")
    if ai.startswith("Possible"): s += 3; why.append("AI signal")
    t = (r[19] or "")
    if "Retention" in t: s -= 1; why.append("hard to move")
    if "Thin EM" in t: s -= 1; why.append("thin EM exp")
    if r[8] == "Not a Fit (Profile)": s -= 1
    if re.search(r"failed|rejected", str(r[20] or ""), re.I) and "Status conflict" in (r[23] or ""): s -= 4; why.append("failed a round per one tab")
    if r[9] == "A": s += 1; why.append("TA-approved")
    return s, ", ".join(why)

def nxt(r):
    w, st = r[10], r[7]
    if w.startswith("W0"):
        n = {"Replied / Interested": "book the recruiter screen", "Recruiter Screen": "book the Screen Round (Technical Retrospective + People Mgmt / HM)",
             "Screen Round": "book the Full Loop (System Design + XFN)", "Full Loop": "book the remaining Full Loop rounds, then Post Loop",
             "Post Loop": "check the debrief / offer decision"}.get(st, "confirm next step")
        return f"Confirm still interested, then {n}", "Mon 5 Oct"
    if w.startswith("W1"): return "Personal message with the Growth AI pitch; ask about interest and timing", "Wed 7 Oct"
    if w.startswith("W3"): return "Follow-up message (4-step follow-up plan)", "Thu 8 Oct"
    return ("Send first message", "Tue 6 Oct") if r[9] == "A" else ("Quick HM look at the profile, then send first message", "Tue 13 Oct")
def stand(r):
    st, w = r[8], r[10]
    if w.startswith("W0"): return "In process"
    if w.startswith("W1"): return {"On Hold": "On hold – paused earlier", "Nurture – Re-engage Later": "Warm – timing wasn't right"}.get(st, "Closed earlier – notes say try again")
    if w.startswith("W3"): return "Messaged – no reply"
    return "Not contacted yet"

out = []
for r in rows:
    detail = clean(r[20])
    if r[10].startswith("W2") and not (r[3] or r[4] or detail): continue
    if r[5] == "IC" and r[10].startswith("W2"): continue
    be, ai, _ = signals(r); sc, why = score(r, be, ai)
    risks = r[19] or ""
    wt = "; ".join(x for x in (("Stack gap (not .NET/Java/C#)" if "Stack gap" in risks else ""), ("Hard to move (long tenure)" if "Retention" in risks else ""),
          ("Limited EM experience" if "Thin EM" in risks else ""), ("Relocation / remote issue" if "Relocation" in risks else ""), ("Timing / availability" if "Timing" in risks else ""),
          ("Comp expectations" if "Comp" in risks else ""), ("Records disagree between tabs – verify first" if "Status conflict" in (r[23] or "") else "")) if x)
    na, when = nxt(r)
    need = "Yes" if (r[5] in ("Unknown", "Lead (verify)") or be.startswith("Unknown") or ai.startswith("Unknown")) else "Optional"
    out.append([sc, r[1], " @ ".join(x for x in (r[3], r[4]) if x) or "(not recorded)", SENIOR[r[5]], stand(r), STAGE[r[7]], detail, be, ai, why, wt, na, when, need, r[2]])
out.sort(key=lambda x: (-x[0], x[1].lower()))
n = len(out)
def tier(i, sc): return "P1" if i < 20 else "P2" if i < 70 else "P3"
final = [[tier(i, x[0]), i + 1] + x[1:14] + ["", "", "", "", "", x[14]] for i, x in enumerate(out)]
hdr = ["Priority (P1 = first)", "Rank", "Candidate", "Current role", "Seniority fit for Sr Manager role", "Where they stand today", "Stage in current interview process", "What we know (from existing notes)",
       "Backend / distributed signal", "Hands-on AI signal", "Why this priority", "Watch-outs", "Next step", "Target date", "LinkedIn check needed?",
       "LinkedIn checked? (Y/N)", "LinkedIn findings (backend, AI, scope, tenure)", "Owner", "Date contacted", "Outcome / notes", "LinkedIn"]
with open(OUT, "w", encoding="utf8", newline="") as f:
    w = csv.writer(f, lineterminator="\n"); w.writerow(hdr); w.writerows(final)
print(n, "rows", collections.Counter(x[0] for x in final), collections.Counter(x[8] for x in final), collections.Counter(x[9] for x in final))
