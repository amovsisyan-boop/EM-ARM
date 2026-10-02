#!/usr/bin/env python3
"""Consolidate the legacy (Sr.) EM Armenia sheets into one clean tracker workbook.

Usage: python3 tracker/build_tracker.py --src <dir> --out output/EM_Armenia_Hiring_Tracker.xlsx

<dir> must hold exports of the four legacy Google Sheets:
  f1.csv           Armenia Engineering Candidates - Full Pipeline (first tab, CSV)
  f2_EM.xlsx       "EM"  (tabs V1, V2, 02. 2025)
  f3_Arm.xlsx      "Arm Potential Staff + (IC & Manager) Tier 1-2"
  f4_Hiring.xlsx   "[Engineering]EM and Director Hiring"
Candidate data is confidential: keep <dir> and the output out of git.
"""
import argparse, collections, csv, re, urllib.parse
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

RESTART = "2026-10-02"

# ----------------------------------------------------------------- taxonomy
LEGACY_STAGES = ["Sourced", "Profile Review", "Outreach Sent", "Replied / Interested", "Recruiter Screen", "HM Screen", "CP1 – Coding",
                 "System Design", "CP3 – Final Interview", "Offer", "Hired"]
LIDX = {s: i for i, s in enumerate(LEGACY_STAGES)}

# Current process: "Engineering Interview Process – Proposal" (EM / eDir loop)
STAGES = [
    ("Sourced", "On the list, not yet reviewed."),
    ("Profile Review", "TA / HM reviewing the LinkedIn profile. Exit: Proceed or Not a Fit."),
    ("Outreach Sent", "First message sent, waiting for reply. Follow-up cadence applies."),
    ("Replied / Interested", "Candidate answered and is open to a call. Exit: recruiter screen booked."),
    ("Recruiter Screen", "Motivation, level, scope, comp range, relocation / office, process walkthrough."),
    ("Screen Round", "ADVANCE GATE – both must pass: Technical Retrospective (45 min, no coding) + People Management / HM round (45 min)."),
    ("Full Loop", "System Design (45 min, HackerRank canvas) + Cross-Org / XFN collaboration (PM or Design director). Coding (60 min) ONLY for Frontline Managers; not required for Sr Manager and above."),
    ("Post Loop", "Behavioral with People Business Partner (PBP) + VP Bar Raiser. Every EM loop includes exactly one PBP interview."),
    ("Offer", "Debrief quorum decided; offer approved / extended / negotiating."),
    ("Hired", "Offer accepted."),
]
STAGE_NAMES = [s for s, _ in STAGES]
SIDX = {s: i for i, s in enumerate(STAGE_NAMES)}
NEWMAP = {"Sourced": "Sourced", "Profile Review": "Profile Review", "Outreach Sent": "Outreach Sent", "Replied / Interested": "Replied / Interested",
          "Recruiter Screen": "Recruiter Screen", "HM Screen": "Screen Round", "CP1 – Coding": "Full Loop", "System Design": "Full Loop",
          "CP3 – Final Interview": "Post Loop", "Offer": "Offer", "Hired": "Hired"}
NEXT_STEP = {"Sourced": "profile review", "Profile Review": "outreach", "Outreach Sent": "a reply", "Replied / Interested": "Recruiter Screen",
             "Recruiter Screen": "Screen Round (Technical Retrospective + People Mgmt / HM round)", "Screen Round": "Full Loop (System Design + XFN)",
             "Full Loop": "Post Loop (Behavioral / PBP + VP Bar Raiser)", "Post Loop": "debrief and offer decision", "Offer": "offer close", "Hired": "onboarding"}

STATUSES = [
    ("Active", "Open", "In process; a next step is owed by us or the candidate."),
    ("No Response", "Open", "Contacted, no reply after the full cadence (4 touches). Recycle after 90 days."),
    ("On Hold", "Paused", "Process paused (candidate timing, our headcount, pending decision). Needs a re-engage date."),
    ("Nurture – Re-engage Later", "Paused", "Warm but not now ('next year', 'busy 6 months', 'good profile, no role'). Re-engage on date."),
    ("Not Interested", "Closed", "Candidate declined / not interested. Recycle only if note says 'for now/later'."),
    ("Not a Fit (Profile)", "Closed", "We closed it at profile review (experience, level, stack, background)."),
    ("Rejected by ST", "Closed", "We rejected after an interview stage."),
    ("Withdrew", "Closed", "Candidate withdrew mid-process."),
    ("Hired", "Closed", "Hired – remove from outreach."),
]
STATUS_NAMES = [s for s, _, _ in STATUSES]
CLOSED = {s for s, c, _ in STATUSES if c == "Closed"}

WAVES = [
    ("W0 Validate & continue", "In-process at RS or later (or replied). Confirm still live, book the next stage."),
    ("W1 Re-engage warm", "Previously interested / on hold / nurture / 'reconsider'. Personal message, new Growth AI pitch."),
    ("W2 Fresh outreach", "Never contacted. Approved profiles first."),
    ("W3 Follow-up (no reply)", "Contacted before, no answer. Run the cadence."),
    ("Hold / closed", "Do not contact unless note changes."),
    ("Done", "Hired."),
]
WAVE_NAMES = [w for w, _ in WAVES]
WORDER = {w: i for i, w in enumerate(WAVE_NAMES)}

LEVELS = ["Sr EM", "EM", "Lead (verify)", "Director+", "IC", "Unknown"]
INTEREST = ["Open to Sr EM / EM", "Director-leaning", "IC-leaning"]
PRIORITY = ["A", "B", "C"]
OWNERS = ["Anush Movsisyan", "Ruzanna Kamalyan", "Lusine Khachatryan", "Rebeka Vahanyan", "Mary", "Unassigned"]
CADENCE = [(1, 3), (2, 4), (3, 7), (4, None)]  # touch -> days until next touch

SRC_PREC = {"Full Pipeline": 1, "FY26-27 Sr Mgr list": 2, "Hiring: NEW (Sr) EM": 3, "Hiring: NEW Director": 4,
            "Hiring: OLD Director": 5, "Hiring: OLD EM": 6, "Staff+ Priority IC": 7, "Vahe meetings": 8}

NURTURE_RE = re.compile(
    r"(next year|not interested (now|currently|at the moment|until|till|untill)|not interested for now|currently not interested|"
    r"not interested currently|for now|right now|at the moment|few months|4-6 months|several months|months|busy|"
    r"till fall|until fall|untill fall|strong nurture|taking a break|can consider|may consider|might be interested|still interested|"
    r"reconsider|probably still interested|reached out again|reach out again|check again|revisit|could be a good candidate|"
    r"can follow up|follow up in)", re.I)
NORESP_RE = re.compile(r"(no reply|ghosted|did not reply|didn't reply|stopped responding|did not join and stopped)", re.I)


# ----------------------------------------------------------------- helpers
def clean(v):
    if v is None:
        return ""
    s = str(v).replace("\\&", "&").replace("\\#", "#").replace("\\|", "|").replace("\\[", "[").replace("\\]", "]")
    return re.sub(r"[ \t]+", " ", s).strip()


def slug(u):
    if not u:
        return ""
    u = urllib.parse.unquote(str(u).strip()).lower()
    m = re.search(r"linkedin\.com/(?:in|pub)/([^/?#\s]+)", u)
    return m.group(1).strip("/") if m else ""


def li_url(u, sl):
    return f"https://www.linkedin.com/in/{sl}" if sl else clean(u)


def namekey(n):
    n = re.sub(r"\b(phd|ph\.d\.?|mba|pmp)\b", "", (n or "").lower())
    toks = sorted(re.findall(r"[a-zа-яё]+", n))
    return "".join(toks)


def name_from_slug(sl):
    s = re.sub(r"-[0-9a-f]{6,}$", "", sl)
    s = re.sub(r"\d+", "", s).replace("-", " ").strip()
    return s.title() if s else sl


def level_of(title):
    t = (title or "").lower().strip()
    if not t or t in ("—", "-"):
        return "Unknown"
    if re.search(r"head of unit|team lead|tech lead|technical lead|lead (software|backend|developer|engineer)|technical owner", t) \
            and not re.search(r"manager", t):
        return "Lead (verify)"
    if re.search(r"\b(vp|vice president|cto|cio|chief|head of|executive director|director|founder|co-?founder|deputy)\b", t) \
            and "head of unit" not in t:
        return "Director+"
    if re.search(r"(senior|sr\.?|group|principal)[^|,]*(manager|mgr|\bem\b)|senior (em|manager)|sr\.? ?(em|manager)|manager ?2|managing engineer", t):
        return "Sr EM" if not re.search(r"managing engineer|manager ?2", t) else "EM"
    if re.search(r"manager|mgr|\bem\b|engineering leader|head of unit", t):
        return "EM"
    if re.search(r"lead", t):
        return "Lead (verify)"
    if re.search(r"engineer|developer|architect|swe|programmer|analyst|professor|scientist", t):
        return "IC"
    return "Unknown"


def stage_from_text(t):
    t = (t or "").lower()
    if not t:
        return None
    if re.search(r"\boffer\b", t): return "Offer"
    if re.search(r"cp ?3|arch(itecture)? disc|final|ashot (and|&) pbp|\bnick\b|culture fit", t): return "CP3 – Final Interview"
    if re.search(r"system ?design|\bsd\b", t): return "System Design"
    if re.search(r"cp ?(1|i\b|ii\b|2)|codepair|code ?pair|coding|technical interview|failed coding", t): return "CP1 – Coding"
    if re.search(r"hm screen|hm interview|after hm|hrbp|after ashot|ashot", t) and not re.search(r"before hm", t): return "HM Screen"
    if re.search(r"\brs\b|recruiter screen|ta screen|screening|\bta interview|ta call", t): return "Recruiter Screen"
    if re.search(r"replied|follow-up received|expressed interest", t): return "Replied / Interested"
    if re.search(r"contacted|reach", t): return "Outreach Sent"
    if re.search(r"review|shortlist|pending", t): return "Profile Review"
    return None


def furthest(*stages):
    ss = [s for s in stages if s]
    return max(ss, key=lambda s: LIDX[s]) if ss else None


# ----------------------------------------------------------------- loading
def rows_of(path, tab_prefix, hyperlinks=False):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = [w for w in wb if w.title.startswith(tab_prefix)][0]
    out = []
    for r in ws.iter_rows():
        vals = []
        for c in r:
            v = clean(c.value)
            if hyperlinks and c.hyperlink and c.hyperlink.target and "linkedin" in c.hyperlink.target.lower():
                v = c.hyperlink.target
            vals.append(v)
        if any(vals):
            out.append(vals)
    return out


class Entry:
    def __init__(self, src, name="", url="", **kw):
        self.src, self.name, self.url = src, clean(name), clean(url)
        self.slug = slug(url)
        if self.name.lower().startswith("http") or not self.name:
            self.name = name_from_slug(slug(self.name) or self.slug) if (slug(self.name) or self.slug) else self.name
        self.title = kw.get("title", ""); self.company = kw.get("company", "")
        self.location = kw.get("location", ""); self.skills = kw.get("skills", "")
        self.status = kw.get("status"); self.stage = kw.get("stage")
        self.status_raw = kw.get("status_raw", ""); self.stage_raw = kw.get("stage_raw", "")
        self.note = kw.get("note", ""); self.kind = kw.get("kind", "status")   # status | pool
        self.hint = kw.get("hint", "")      # em / director / ic / mobile / vahe
        self.approved = kw.get("approved", False)
        self.assess = kw.get("assess", ""); self.owner = kw.get("owner", ""); self.level_note = kw.get("level_note", "")
        self.type = kw.get("type", "")


def derive(status_raw, stage_raw, notes, review="", mary="", final=""):
    """Map legacy free-text to (stage, status, approved)."""
    all_txt = " ".join([status_raw, stage_raw, notes, final, mary]).lower()
    s = (final or status_raw or "").lower().strip()
    stage = furthest(stage_from_text(stage_raw), stage_from_text(final), stage_from_text(notes if re.search(r"\b(rs|cp ?\d|cp i|system design|hm screen)\b", notes.lower()) else ""))
    status, approved = None, False
    rev = review.lower()
    m = mary.lower()
    approved = "yes - proceed" in rev or "shortlist" in s or "top candidate" in all_txt
    if "hired" in s: status = "Hired"
    elif "withdr" in s or "withdr" in stage_raw.lower(): status = "Withdrew"
    elif "reject" in s or s.startswith("decline after") or "declined after" in s or re.match(r"^(after (hm screen|technical interview|ashot))", s): status = "Rejected by ST"
    elif s == "decline" or "decline -" in s: status = "Not Interested"
    elif "on hold" in s or "on hold" in stage_raw.lower(): status = "On Hold"
    elif "not interested" in s or s == "not interested": status = "Not Interested"
    elif s in ("reviewed - no", "not shortlisted", "no - not a fit"): status = "Not a Fit (Profile)"
    elif s in ("rs", "interviewing", "system design", "successfully passed system design", "cp 1", "in progress", "in process"): status = "Active"
    elif "might be interested" in s: status = "Active"; stage = furthest(stage, "Replied / Interested")
    elif s.startswith("before hm screen"):
        status = "Not a Fit (Profile)" if re.search(r"(^no\b|not a fit|not much|little|limited|no em|fe |qa|failed|declined|previous)", notes.lower()) else "Active"
        if re.search(r"failed|declined|previous interview|codepair|cpi", notes.lower()): status = "Rejected by ST"; stage = furthest(stage, "CP1 – Coding")
    elif s == "not interested" or s.startswith("not interested"): status = "Not Interested"
    elif s == "pending review": status = "Active"; stage = "Profile Review"
    elif s == "reviewed": status = "Not a Fit (Profile)" if re.search(r"(^no\b|no em|not much|little|<|limited|too ic|only tech lead|skip)", notes.lower()) else "Active"
    elif s in ("contacted", "needs review", "review", "shortlist") or s.startswith("shortlist"):
        status = "Active"
        if s == "contacted": stage = furthest(stage, "Outreach Sent")
        else: stage = furthest(stage, "Profile Review")
    elif s in ("n/a", ""):
        status = None
    # secondary signals from Mary / review columns (hiring-tab layout)
    if status is None:
        if "not a fit" in m or "no - not a fit" in rev: status = "Not a Fit (Profile)"
        elif "not priority" in m: status = "Nurture – Re-engage Later"
        elif "replied" in m: status = "Active"; stage = furthest(stage, "Replied / Interested")
        elif "contacted" in m: status = "Active"; stage = furthest(stage, "Outreach Sent")
        elif rev: status = "Active"; stage = furthest(stage, "Profile Review")
    nl = notes.lower()
    if status in (None, "Active") and re.search(r"withdrew|withdrawing", nl): status = "Withdrew"; stage = furthest(stage, "HM Screen" if "hm screen" in nl else None)
    elif status in (None, "Active") and re.search(r"(declined|decline) after (hm screen|cp ?i|codepair|cp ?ii)|failed (cp3|system design|coding)|did not pass", nl):
        status = "Rejected by ST"
    if status == "Active" and stage is None and (status_raw or review): stage = "Profile Review"
    # refine
    if status in ("Not Interested", "Withdrew", "On Hold") and NURTURE_RE.search(notes + " " + final):
        status = "Nurture – Re-engage Later" if status != "On Hold" else "On Hold"
    if status in ("Active", "Not Interested") and NORESP_RE.search(notes):
        if status == "Active" and (stage in (None, "Outreach Sent", "Profile Review")):
            status = "No Response"; stage = furthest(stage, "Outreach Sent")
        elif status == "Not Interested" and not re.search(r"not interested", notes.lower().replace("no reply", "")):
            status = "No Response"; stage = furthest(stage, "Outreach Sent")
    return stage, status, approved


# ---------- build the entry list
def load_entries(src):
    E = []
    # f1 Full Pipeline (CSV, 8 cols)
    with open(f"{src}/f1.csv", encoding="utf8") as fh:
        for i, r in enumerate(csv.reader(fh)):
            if i == 0 or not any(r): continue
            r = [clean(x) for x in r] + [""] * 8
            name, typ, title, comp, url, st, sg, note = r[:8]
            stage, status, appr = derive(st, sg, note)
            E.append(Entry("Full Pipeline", name, url, title=title, company=comp, status=status, stage=stage,
                           status_raw=st, stage_raw=sg, note=note, hint="ic" if typ == "IC" else "em", approved=appr, type=typ))
    # f3 Sheet14 (FY26-27 Sr Manager list)
    for r in rows_of(f"{src}/f3_Arm.xlsx", "Sheet14", hyperlinks=True)[2:]:
        r += [""] * 9
        name, url, st, sg, comp, title, note, origin, extra = r[:9]
        if not name or name == "Name": continue
        stage, status, appr = derive(st, sg, note)
        if st.lower() == "cp 1": stage = furthest(stage, "System Design") if "system design" in sg.lower() else "CP1 – Coding"
        E.append(Entry("FY26-27 Sr Mgr list", name, url, title=title, company=comp, status=status, stage=stage,
                       status_raw=st, stage_raw=sg, note=" | ".join(x for x in (note, extra) if x), hint="em", approved=appr))
    # f4 hiring tabs
    for tab, src_label, hint, rev_i in (("NEW -(Sr)", "Hiring: NEW (Sr) EM", "em", 3), ("NEW - Director", "Hiring: NEW Director", "director", 3),
                                         ("OLD - Director", "Hiring: OLD Director", "director", 3)):
        rows = rows_of(f"{src}/f4_Hiring.xlsx", tab, hyperlinks=True)
        for r in rows[1:]:
            r += [""] * 10
            if tab.startswith("OLD - Director"):
                name, url, recn, review, hm, mary, reach, final = r[:8]; extra = ""
            else:
                name, url, recn, review, hm, mary, reach, final, extra = r[:9]
            if not name: continue
            notes = " | ".join(x for x in (recn, hm, reach, extra) if x)
            stage, status, appr = derive("", "", reach + " " + extra, review=review, mary=mary, final=final)
            if "failed cp3" in notes.lower(): stage = "CP3 – Final Interview"
            if final.lower() == "rs" and not stage: stage = "Recruiter Screen"
            if final.lower().startswith("meeting with ashot"): stage = "CP3 – Final Interview"; status = "Active"
            if re.search(r"(for|consider(ed)? for|fit for) (a )?(senior )?em\b|em position|senior em|sr\.? ?em|reporting team em", notes.lower()) and hint == "director": hint = "em"
            E.append(Entry(src_label, name, url, status=status, stage=stage, status_raw=" / ".join(x for x in (review, mary, final) if x),
                           stage_raw=reach, note=notes, hint=hint, approved=appr))
    # f4 OLD - EM (URL + reasons only)
    rows = rows_of(f"{src}/f4_Hiring.xlsx", "OLD - EM")
    for r in rows[1:]:
        r += [""] * 5
        url, why, decl, d, e = r[:5]
        if not slug(url):
            continue   # junk rows are reported in the cleanup log
        notes = " | ".join(x for x in (why, d, e) if x)
        stage, status, appr = derive(decl, "", notes)
        dl = decl.lower()
        if re.search(r"after hm|after ashot|hrbp|declined after hm|decline after hm", dl): stage = furthest(stage, "HM Screen")
        if "failed system design" in notes.lower(): stage = "System Design"
        if not decl and re.fullmatch(r".{0,40}", notes) and re.search(r"c\+\+|fe|python|golang|java|c#|ruby", notes.lower()):
            status, stage = "Active", "Profile Review"
        E.append(Entry("Hiring: OLD EM", name_from_slug(slug(url)), url, status=status, stage=stage, status_raw=decl, note=notes, hint="em", approved=False))
    # f3 Priority IC (status-bearing, Staff+ req)
    for r in rows_of(f"{src}/f3_Arm.xlsx", "Priority", hyperlinks=True)[1:]:
        r += [""] * 14
        name, url, title, comp, lvl, fit, sg, st, owner, det, nxt, n1, n2, n3 = r[:14]
        if not name: continue
        stage, status, appr = derive(st, sg, " ".join((det, nxt)))
        E.append(Entry("Staff+ Priority IC", name, url, title=title, company=comp, status=status, stage=stage, status_raw=st, stage_raw=sg,
                       note=" | ".join(x for x in (det, nxt, n1, n2, n3) if x), hint="ic", owner=owner, level_note=lvl))
    for r in rows_of(f"{src}/f3_Arm.xlsx", "Potential to meet")[1:]:
        r += [""] * 8
        name, url, notes, st, brief, slot, office, who = r[:8]
        note = " | ".join(x for x in (notes, st, f"Vahe meeting day {brief} {slot}".strip() if brief else "") if x)
        offer = bool(re.search(r"offer out|offer negotiat", note.lower()))
        E.append(Entry("Vahe meetings", name, url, status="Active" if offer else None, stage="Offer" if offer else None, status_raw=st, note=note, hint="vahe", owner=who))
    # pool sources
    for tab, lab in (("V1", "EM sourcing V1"), ("V2", "EM sourcing V2")):
        for r in rows_of(f"{src}/f2_EM.xlsx", tab)[1:]:
            r += [""] * 11
            fn, ln, head, loc, title, comp, em, ph, url, n1, n2 = r[:11]
            E.append(Entry(lab, f"{fn} {ln}".strip(), url, title=title, company=comp, location=loc, skills=head, kind="pool",
                           note=" | ".join(x for x in (n1, n2) if x)))
    for r in rows_of(f"{src}/f2_EM.xlsx", "02. 2025"):
        E.append(Entry("EM sourcing 02.2025", r[0], r[1] if len(r) > 1 else "", kind="pool"))
    for r in rows_of(f"{src}/f3_Arm.xlsx", "Engineering Leadership")[1:]:
        r += [""] * 7
        E.append(Entry("Staff+ Leadership list", r[0], r[1], company=r[2], title=r[3], skills=r[4], assess=r[5], kind="pool"))
    for r in rows_of(f"{src}/f3_Arm.xlsx", "Mobile")[1:]:
        r += [""] * 7
        E.append(Entry("Staff+ Mobile list", r[0], r[1], company=r[2], title=r[3], skills=r[4], assess=r[6], kind="pool", hint="mobile-mgr" if r[5] == "Manager" else "mobile-ic"))
    for r in rows_of(f"{src}/f3_Arm.xlsx", "IC  ")[1:]:
        r += [""] * 7
        E.append(Entry("Staff+ IC list", r[0], r[1], company=r[2], title=r[3], skills=r[4], assess=r[5], note=r[6], kind="pool", hint="ic"))
    full = rows_of(f"{src}/f3_Arm.xlsx", "Full list ")
    company_assess = {}
    for r in full[1:]:
        r += [""] * 11
        nm = f"{r[0]} {r[1]}".strip()
        if r[10] and r[4]: company_assess.setdefault(r[4], r[10])
        E.append(Entry("Staff+ Full list", nm, r[3], title=r[2], company=r[4], skills=r[5], assess=r[9], note=r[6], kind="pool"))
    companies = [r[0] for r in rows_of(f"{src}/f4_Hiring.xlsx", "Sourcing List")][1:]
    return E, companies, company_assess


# ----------------------------------------------------------------- grouping & merge
def group(entries):
    parent = list(range(len(entries)))
    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]; i = parent[i]
        return i
    by_slug = {}
    for i, e in enumerate(entries):
        if e.slug:
            if e.slug in by_slug: parent[find(i)] = find(by_slug[e.slug])
            else: by_slug[e.slug] = i
    # entries with no slug: attach by unique name key
    by_name = collections.defaultdict(set)
    for i, e in enumerate(entries):
        if e.slug and len(namekey(e.name)) >= 8: by_name[namekey(e.name)].add(find(i))
    for i, e in enumerate(entries):
        if not e.slug:
            nk = namekey(e.name)
            if len(nk) >= 8 and len(by_name.get(nk, ())) == 1:
                parent[find(i)] = find(next(iter(by_name[nk])))
    def ctoken(e):
        m = re.findall(r"[a-z]{3,}", (e.company or "").lower())
        return m[0] if m else ""
    seen = {}
    for i, e in enumerate(entries):
        nk, ct = namekey(e.name), ctoken(e)
        if len(nk) >= 8 and ct:
            key = (nk, ct)
            if key in seen: parent[find(i)] = find(seen[key])
            else: seen[key] = i
    groups = collections.defaultdict(list)
    for i, e in enumerate(entries): groups[find(i)].append(e)
    return list(groups.values())


def sclass(s):
    if s in CLOSED: return "closed"
    if s in ("On Hold", "Nurture – Re-engage Later"): return "paused"
    return "open"


def tidy_notes(parts, cap=900):
    seen, out = set(), []
    for src, n in parts:
        for chunk in re.split(r"\s*\|\s*", n):
            c = chunk.strip()
            k = re.sub(r"\W+", "", c.lower())
            if c and k not in seen and len(k) > 1:
                seen.add(k); out.append(c)
    s = " | ".join(out)
    return s if len(s) <= cap else s[:cap - 1] + "…"


def risks(text):
    t = text.lower(); tags = []
    pats = [("Relocation/remote", r"relocat|remote|abroad|uk\b|abu dhabi|kazakhstan|us/canada|move to us"),
            ("Comp", r"compensation|comp expect|base comp"), ("Coding concern (Frontline only)", r"coding interview|codepair|passing coding|uncertain about moving forward with a coding"),
            ("Timing", r"busy|timing|time for interview|fall|few months|4-6 months|several months|maternity|vacation"),
            ("Retention risk", r"embedded|long tenure|locked in|stock|unlikely to leave|only one place"),
            ("Stack gap", r"\bfe\b|front-?end|\bqa\b|test engineering|ios|android|mobile|c\+\+|edA|python|ruby|golang|sre"),
            ("Thin EM exp", r"little|<\s?2 years|less than a year|limited|no em|not much (em|managerial)|only tech lead|1 year"),
            ("Director-only", r"only.*director|director only|interested only in d")]
    for tag, p in pats:
        if re.search(p, t): tags.append(tag)
    return ", ".join(tags)


def interest_of(text):
    t = text.lower()
    if re.search(r"interested only in ic|only in ic roles", t): return "IC-leaning"
    if re.search(r"(sr|senior) (em|manager)", t) and re.search(r"interested|open|fit", t) and not re.search(r"not interested", t): return "Open to Sr EM / EM"
    if re.search(r"director", t) and re.search(r"interested|only|prefer|\bdir\b", t): return "Director-leaning"
    if re.search(r"interested in em role", t): return "Open to Sr EM / EM"
    return ""


def merge(groups, companies):
    P, O, POOL, log = [], [], [], collections.Counter()
    junk = []
    for g in groups:
        st = [e for e in g if e.kind == "status"]
        pool = [e for e in g if e.kind == "pool"]
        st.sort(key=lambda e: SRC_PREC[e.src])
        best_url = next((e.url for e in g if e.slug), "") or next((e.url for e in g if e.url), "")
        sl = next((e.slug for e in g if e.slug), "")
        title = next((e.title for e in sorted(g, key=lambda e: (e.src != "Full Pipeline", e.src)) if e.title and e.title not in ("—", "-")), "")
        comp = next((e.company for e in sorted(g, key=lambda e: (e.src != "Full Pipeline", e.src)) if e.company and e.company not in ("—", "-")), "")
        name_src = [e for e in g if e.src != "Hiring: OLD EM" and e.name]
        name = (name_src[0].name if name_src else g[0].name)
        name_derived = not name_src
        loc = next((e.location for e in g if e.location), "")
        skills = next((e.skills for e in g if e.skills), "")
        assess = next((e.assess for e in g if e.assess), "")
        srcs = []
        for e in g:
            if e.src not in srcs: srcs.append(e.src)
        all_notes = tidy_notes([(e.src, e.note) for e in g] + [("", assess)] if assess and not st else [(e.src, e.note) for e in g])
        legacy = " ; ".join(f"{e.src}: {e.status_raw or '–'} → {e.stage_raw or '–'}" for e in st if (e.status_raw or e.stage_raw))
        # a concrete outcome beats a placeholder "Active / Profile Review" from a higher-precedence tab
        status = next((e.status for e in st if e.status and not (e.status == "Active" and e.stage in (None, "Profile Review", "Sourced"))), None) \
            or next((e.status for e in st if e.status), None)
        stage = furthest(*[e.stage for e in st]) if st else None
        approved = any(e.approved for e in st)
        classes = {sclass(e.status) for e in st if e.status}
        conflict = len(classes) > 1
        hints = {e.hint for e in g if e.hint}
        lvl = level_of(title)
        textall = " ".join(e.note for e in g) + " " + " ".join(e.status_raw + " " + e.stage_raw for e in g)
        interest = interest_of(textall)
        has_em_src = bool({e.src for e in st} & {"Hiring: NEW (Sr) EM", "FY26-27 Sr Mgr list", "Hiring: OLD EM"}) or "em" in hints and any(e.src == "Full Pipeline" and e.type == "People Manager" for e in st)
        dir_only = (lvl == "Director+" and interest == "Director-leaning" and not {e.src for e in st} & {"Hiring: NEW (Sr) EM", "FY26-27 Sr Mgr list"})
        if "vahe" in hints and not has_em_src: dest = "other"; req = "Staff+ / CTO (Vahe meetings)"
        elif {e.src for e in st} <= {"Staff+ Priority IC"} and st: dest = "other"; req = "Staff+ IC"
        elif "ic" in hints and not has_em_src and not (hints & {"director"}) and (st or any(e.hint == "ic" for e in pool)): dest = "other"; req = "Staff+ IC"
        elif has_em_src and not dir_only: dest = "pipeline"
        elif "director" in hints or lvl == "Director+" or dir_only:
            dest = "other"; req = "Director+"
            if has_em_src and not dir_only: dest = "pipeline"
        elif st and "em" in hints: dest = "pipeline"
        elif st: dest = "other"; req = "Staff+ IC"
        else: dest = "pool"
        if dest == "pool" and any(e.hint == "mobile-ic" for e in pool): dest, req = "other", "Mobile IC"
        if dest == "pool" and any(e.hint == "ic" for e in pool): dest, req = "other", "Staff+ IC"
        rec = dict(name=name, url=li_url(best_url, sl), title=title, company=comp, level=lvl, interest=interest, stage=stage, status=status,
                   approved=approved, notes=all_notes, srcs=", ".join(srcs), legacy=legacy, conflict=conflict, name_derived=name_derived,
                   no_url=not sl, loc=loc, skills=skills, assess=assess, owner=next((e.owner for e in g if e.owner), ""),
                   level_note=next((e.level_note for e in g if e.level_note), ""), ndup=len(g), textall=textall)
        if dest == "pipeline": P.append(rec)
        elif dest == "other": rec["req"] = req; O.append(rec)
        else: POOL.append(rec)
    return P, O, POOL


# ----------------------------------------------------------------- derived fields
def finish_pipeline(rec):
    status, stage = rec["status"], rec["stage"]
    if status is None:
        status = "Active"; stage = stage or ("Profile Review" if rec["approved"] else "Sourced")
    if stage is None: stage = "Outreach Sent" if status in ("No Response",) else "Profile Review"
    rec["detail_stage"] = stage                 # most specific legacy step reached
    stage = NEWMAP[stage]
    rec["status"], rec["stage"] = status, stage
    notes_l = rec["textall"].lower()
    reconsider = bool(NURTURE_RE.search(rec["notes"]))
    idx = SIDX[stage]
    if status == "Hired": wave = "Done"
    elif status in ("Rejected by ST", "Not a Fit (Profile)"):
        wave = "W1 Re-engage warm" if (reconsider and status == "Not a Fit (Profile)" and re.search(r"reach out again|reconsider|still interested|consider him|revisit", notes_l)) else "Hold / closed"
    elif status in ("Not Interested", "Withdrew"): wave = "Hold / closed"
    elif status in ("Nurture – Re-engage Later", "On Hold"): wave = "W1 Re-engage warm"
    elif status == "No Response": wave = "W3 Follow-up (no reply)"
    else:
        if idx >= SIDX["Recruiter Screen"] or stage == "Replied / Interested": wave = "W0 Validate & continue"
        elif stage == "Outreach Sent": wave = "W3 Follow-up (no reply)"
        else: wave = "W2 Fresh outreach"
    if rec["interest"] == "Director-leaning" and wave in ("W1 Re-engage warm", "W2 Fresh outreach", "W3 Follow-up (no reply)") and rec["level"] in ("Director+",):
        wave = "Hold / closed"
    rec["director_hold"] = (wave == "Hold / closed" and rec["interest"] == "Director-leaning" and status not in CLOSED)
    rec["wave"] = wave
    # priority (market-informed: product/SaaS scale + stack fit beats long-tenured EDA/outsourcing C++ profiles)
    closed = wave in ("Hold / closed", "Done")
    pr = ""
    if not closed:
        strong_co = re.search(r"picsart|krisp|miro|superannotate|podcastle|connectwise|disqo|adobe|synopsys|digitain|mentorcliq|servicetitan|epam|axcient|playrix", (rec["company"] or "").lower())
        if (status == "Active" and idx >= SIDX["Recruiter Screen"]) or rec["approved"] or wave == "W0 Validate & continue": pr = "A"
        elif rec["level"] in ("Sr EM", "EM") and (strong_co or wave == "W1 Re-engage warm"): pr = "A" if wave == "W1 Re-engage warm" and idx >= SIDX["Outreach Sent"] else "B"
        elif rec["level"] in ("Sr EM", "EM", "Lead (verify)", "Unknown"): pr = "B"
        else: pr = "C"
        if re.search(r"c\+\+|eda|embedded|long tenure|front-?end|\bqa\b", notes_l) and pr == "A" and wave == "W2 Fresh outreach": pr = "B"
    rec["priority"] = pr
    nxt = {"W0 Validate & continue": f"Confirm still live → schedule {NEXT_STEP[stage]}",
           "W1 Re-engage warm": "Personal re-engage msg (new Growth AI scope); confirm interest + timing",
           "W2 Fresh outreach": "HM quick-review → send outreach #1" if stage in ("Sourced", "Profile Review") and not rec["approved"] else "Send outreach #1",
           "W3 Follow-up (no reply)": "Follow-up #1 (follow cadence)"}.get(wave, "")
    date = {"W0 Validate & continue": "2026-10-05", "W1 Re-engage warm": "2026-10-07", "W3 Follow-up (no reply)": "2026-10-08"}.get(wave, "")
    if wave == "W2 Fresh outreach": date = "2026-10-06" if pr == "A" else ("2026-10-13" if pr == "B" else "")
    rec["next"], rec["next_date"] = nxt, date
    flags = []
    if rec["conflict"]: flags.append("Status conflict across tabs – verify")
    if rec["name_derived"]: flags.append("Name derived from URL – confirm")
    if rec["no_url"]: flags.append("No LinkedIn URL")
    if status == "Hired": flags.append("Marked hired in legacy – verify")
    if rec["director_hold"]: flags.append("Director-leaning – route to Director req, not (Sr) EM")
    if rec["ndup"] > 1: flags.append(f"Merged {rec['ndup']} rows")
    rec["check"] = "; ".join(flags)
    rec["risks"] = risks(rec["textall"])
    return rec


def pool_tier(rec, target_cos):
    t, loc, comp = rec["title"], rec["loc"], rec["company"]
    lvl = level_of(t)
    why = []
    nonarm = bool(loc) and not re.search(r"armenia|yerevan|^$", loc.lower())
    nonEng = re.search(r"professor|scrum|agile|project manager|product manager|it project|delivery manager|qa |test|hr |recruit", (t or "").lower())
    if nonarm: tier, why = "X", [f"Outside Armenia ({loc})"]
    elif nonEng: tier, why = "X", ["Not an engineering-management role"]
    elif lvl in ("Sr EM", "EM"): tier = "A"
    elif lvl == "Lead (verify)": tier = "B"; why = ["Lead title – check people-management scope"]
    elif lvl == "Unknown": tier = "B"; why = ["No title – review profile"]
    elif lvl == "Director+": tier = "C"; why = ["Above band – route to Director req unless open to Sr EM"]
    elif lvl == "IC": tier = "X"; why = ["IC / architect title – Staff+ req"]
    else: tier = "B"
    if re.search(r"embedded|long tenure|front-?end|not a backend|ios|android|php", (rec["assess"] + " " + rec["skills"]).lower()) and tier == "A":
        tier = "B"; why.append("Stack / retention concern in prior assessment")
    if re.search(r"\.net|c#|java|aws|microservice", (rec["skills"] + rec["assess"]).lower()) and tier in ("A", "B"): why.append("Stack signal: .NET/Java")
    tc = any(c.lower() in (comp or "").lower() for c in target_cos if len(c) > 3)
    return tier, "; ".join(why), "Y" if tc else ""


# ----------------------------------------------------------------- company market view
# Recruiter judgment on the Armenian market (validate with TA before relying on it).
MARKET = {
    "Picsart": ("Tier 1 – product at scale", "Biggest local product company; most EM/Sr EM depth, consumer-scale distributed systems. Competitive comp; strong retention."),
    "Krisp": ("Tier 1 – AI product", "AI-native; strongest Growth-AI narrative match. Hard to move, scarce EM bench."),
    "SuperAnnotate": ("Tier 1 – AI/data platform", "AI/ML platform. Smaller scale; check EM years/people scope."),
    "Miro": ("Tier 1 – product at scale", "Large Yerevan eng hub, AWS/microservices/DDD; strong EM bench."),
    "Adobe": ("Tier 1 – enterprise SaaS", "Deep Java/C#/C++ benches; confirm titles (Manager vs Sr Mgr). Stable, hard to move."),
    "Synopsys": ("Tier 2 – EDA (C++)", "Large mgmt layer but EDA/C++ stack gap; ok for Sr EM only if hands-on in services/backends."),
    "Siemens EDA": ("Tier 2 – EDA (C++)", "Same EDA/C++ profile; evaluate stack transfer."),
    "AMD": ("Tier 2 – semiconductor (C++)", "Large mgr pool; C++/systems stack; lower product-SaaS fit."),
    "VMware": ("Tier 1 – enterprise (post-Broadcom)", "Post-acquisition mobility: strongest source of experienced managers; retention is stock-driven."),
    "Broadcom": ("Tier 1 – enterprise (post-Broadcom)", "Same as VMware; long tenure and unvested stock slow decisions."),
    "ConnectWise": ("Tier 1 – .NET B2B SaaS", "Best stack match (.NET/B2B SaaS); several leaders already approached – cross-check history."),
    "DISQO": ("Tier 1 – product (Java)", "Data-driven product; Java; good EM bench."),
    "EPAM": ("Tier 2 – outsourcing", "Large bench of 'EM/Resource Manager/Head of Unit' – verify product ownership and hands-on depth."),
    "Digitain": ("Tier 2 – gaming/sports (.NET)", "On-stack (.NET) but very long-tenured; Deputy-CTO-level people unlikely to move."),
    "BetConstruct": ("Tier 3 – gaming (PHP/JS)", "Long tenure, PHP-heavy, high retention; low yield."),
    "Playrix": ("Tier 2 – gaming", "Product/game studio; different engineering profile."),
    "Podcastle": ("Tier 1 – AI product", "AI/creator product; small EM bench."),
    "Synergy": ("Tier 2 – enterprise Java", "Java/Kotlin microservices, multi-tenant; good Sr EM/Staff pipeline."),
    "Questrade": ("Tier 2 – remote fintech", "Remote-from-Yerevan engineers; check office-attendance expectation."),
    "Align": ("Tier 2 – MedTech product", "Large Armenia R&D org; Sr Mgr titles."),
    "Axcient": ("Tier 2 – SaaS (backup)", "Product SaaS; check scale."),
    "Webb Fontaine": ("Tier 2 – trade-tech", "Seasoned leaders; sensitive to Director-level scope."),
    "Instigate": ("Tier 3 – services", "Services/R&D; low-medium yield."),
    "Plata": ("Tier 2 – fintech", "Fast-growing fintech; strong engineers, compensation-competitive."),
}


def market_for(comp):
    for k, v in MARKET.items():
        if k.lower() in comp.lower(): return v
    return ("", "")


# ----------------------------------------------------------------- workbook
NAVY, TEAL, GREY, LIGHT = "1F2A44", "2A7F8E", "6B7280", "EEF2F7"
HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BFONT = Font(name="Arial", size=10)
HFILL = PatternFill("solid", fgColor=NAVY)
thin = Side(style="thin", color="D5DAE1")
BORDER = Border(bottom=thin)


def sheet(wb, title, headers, widths, rows, freeze="C2", tab=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for r in rows: ws.append(r)
    for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w
    for c in ws[1]:
        c.font, c.fill = HFONT, HFILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = BFONT; c.alignment = Alignment(vertical="center"); c.border = BORDER
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{L(len(headers))}{max(ws.max_row, 2)}"
    if tab: ws.sheet_properties.tabColor = tab
    return ws


def dv_list(ws, col, formula, n, warn=True):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True)
    dv.errorStyle = "warning" if warn else "stop"; dv.errorTitle = "Not in list"; dv.error = "Pick a value from the list (see Lists tab)."
    ws.add_data_validation(dv); dv.add(f"{col}2:{col}{n}")


def build(args):
    entries, companies, company_assess = load_entries(args.src)
    groups = group(entries)
    P, O, POOL = merge(groups, companies)
    target_cos = [c for c in companies if c.lower() != "armenia"]
    P = [finish_pipeline(r) for r in P]
    nkc = collections.Counter(namekey(r["name"]) for r in P + O + POOL if len(namekey(r["name"])) >= 8)
    for r in P:
        if nkc[namekey(r["name"])] > 1:
            r["check"] = "; ".join(x for x in (r["check"], "Same name as another row (different LinkedIn) – confirm not a duplicate") if x)
    for r in O:
        st, sg, _ = None, None, None
        r["status"] = r["status"] or ""; r["stage"] = r["stage"] or ""
    P.sort(key=lambda r: (WORDER[r["wave"]], r["priority"] or "Z", -SIDX[r["stage"]], r["name"].lower()))
    for i, r in enumerate(P, 1): r["id"] = f"EM-{i:03d}"
    O.sort(key=lambda r: (r["req"], r["status"] == "", r["name"].lower()))
    for r in POOL: r["tier"], r["why"], r["tc"] = pool_tier(r, target_cos)
    POOL.sort(key=lambda r: ("ABCX".index(r["tier"]), r["tc"] != "Y", r["name"].lower()))

    wb = openpyxl.Workbook(); wb.remove(wb.active)
    nP = len(P) + 1

    # ---------------- Lists
    lists = wb.create_sheet("Lists"); lists.sheet_properties.tabColor = GREY
    cols = {"A": ("Stage", STAGE_NAMES), "B": ("Status", STATUS_NAMES), "C": ("Status class", [c for _, c, _ in STATUSES]),
            "D": ("Restart wave", WAVE_NAMES), "E": ("Level fit", LEVELS), "F": ("Interest in role", INTEREST), "G": ("Priority", PRIORITY),
            "H": ("Owner", OWNERS), "I": ("Touch #", [c for c, _ in CADENCE]), "J": ("Days to next touch", [d if d is not None else "stop" for _, d in CADENCE]),
            "K": ("Pool decision", ["Promote to Pipeline", "Review with HM", "Skip"]), "L": ("Company sourcing status", ["To source", "Sourcing", "Sourced", "Skip"]),
            "M": ("Pool tier", ["A", "B", "C", "X"])}
    for col, (h, vals) in cols.items():
        lists[f"{col}1"] = h; lists[f"{col}1"].font = HFONT; lists[f"{col}1"].fill = HFILL
        for i, v in enumerate(vals, 2): lists[f"{col}{i}"] = v; lists[f"{col}{i}"].font = BFONT
        lists.column_dimensions[col].width = 26
    # ---------------- Pipeline
    H = ["ID", "Candidate", "LinkedIn", "Current Title", "Current Company", "Level Fit", "Interest in Role", "Stage", "Status", "Priority",
         "Restart Wave", "Owner", "Touches (restart)", "Last Touch", "Auto Follow-up", "Next Action", "Next Action Date", "Due Date", "Due Flag",
         "Risks / Flags", "Latest Notes (consolidated)", "Source Tabs", "Legacy Status → Stage", "Data Check", "Dup Check", "Furthest step (old-loop detail)"]
    W = [8, 24, 36, 30, 24, 13, 18, 22, 22, 8, 24, 16, 10, 12, 13, 44, 14, 12, 12, 26, 90, 28, 50, 34, 18, 26]
    rows = []
    for i, r in enumerate(P, 2):
        rows.append([r["id"], r["name"], r["url"], r["title"], r["company"], r["level"], r["interest"], r["stage"], r["status"], r["priority"],
                     r["wave"], "", "", "", f'=IF(OR(M{i}="",M{i}=0,N{i}=""),"",IF(M{i}>=4,"",N{i}+INDEX(Lists!$J$2:$J$4,M{i})))',
                     r["next"], r["next_date"] and __import__("datetime").date.fromisoformat(r["next_date"]),
                     f'=IF(OR(I{i}="Not Interested",I{i}="Not a Fit (Profile)",I{i}="Rejected by ST",I{i}="Withdrew",I{i}="Hired"),"",IF(AND(O{i}="",Q{i}=""),"",MIN(IF(O{i}="",99999,O{i}),IF(Q{i}="",99999,Q{i}))))',
                     f'=IF(R{i}="","",IF(R{i}<TODAY(),"OVERDUE",IF(R{i}<=TODAY()+7,"This week","")))',
                     r["risks"], r["notes"], r["srcs"], r["legacy"], r["check"],
                     f'=IF(C{i}="","",IF(COUNTIF($C$2:$C${nP},C{i})>1,"DUPLICATE",IF(COUNTIF(\'Other Roles\'!$C:$C,C{i})>0,"ALSO IN OTHER ROLES","")))', r["detail_stage"]])
    ws = sheet(wb, "Pipeline", H, W, rows, freeze="C2", tab=TEAL)
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.column_letter in "NOQR": c.number_format = "dd-mmm-yy"
        r[0].font = Font(name="Arial", size=10, color=GREY); r[1].font = Font(name="Arial", size=10, bold=True)
    dv_list(ws, "F", "=Lists!$E$2:$E$7", nP); dv_list(ws, "G", "=Lists!$F$2:$F$4", nP); dv_list(ws, "H", "=Lists!$A$2:$A$11", nP, False)
    dv_list(ws, "I", "=Lists!$B$2:$B$10", nP, False); dv_list(ws, "J", "=Lists!$G$2:$G$4", nP); dv_list(ws, "K", "=Lists!$D$2:$D$7", nP)
    dv_list(ws, "L", "=Lists!$H$2:$H$7", nP); dv_list(ws, "M", "=Lists!$I$2:$I$5", nP)
    rng = f"A2:{L(len(H))}{nP}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['OR($I2="Not Interested",$I2="Not a Fit (Profile)",$I2="Rejected by ST",$I2="Withdrew",$I2="Hired")'], font=Font(color="9AA3AF"), fill=PatternFill("solid", bgColor="F3F4F6")))
    for txt, col in (("Active", "D1FAE5"), ("No Response", "FEF3C7"), ("On Hold", "FDE68A"), ("Nurture – Re-engage Later", "DBEAFE")):
        ws.conditional_formatting.add(f"I2:I{nP}", FormulaRule(formula=[f'$I2="{txt}"'], fill=PatternFill("solid", bgColor=col)))
    ws.conditional_formatting.add(f"S2:S{nP}", FormulaRule(formula=['$S2="OVERDUE"'], font=Font(bold=True, color="B91C1C"), fill=PatternFill("solid", bgColor="FEE2E2")))
    ws.conditional_formatting.add(f"S2:S{nP}", FormulaRule(formula=['$S2="This week"'], font=Font(bold=True, color="92400E"), fill=PatternFill("solid", bgColor="FEF3C7")))
    ws.conditional_formatting.add(f"Y2:Y{nP}", FormulaRule(formula=['LEN($Y2)>0'], font=Font(bold=True, color="B91C1C")))
    ws.conditional_formatting.add(f"J2:J{nP}", FormulaRule(formula=['$J2="A"'], font=Font(bold=True, color="065F46")))
    ws.conditional_formatting.add(f"X2:X{nP}", FormulaRule(formula=['LEN($X2)>0'], font=Font(color="B45309")))

    # ---------------- Sourcing pool
    nPool = len(POOL) + 1
    rows = [[r["name"], r["url"], r["title"], r["company"], r["loc"], r["tier"], r["why"], (r["skills"] or "")[:200], r["srcs"], r["tc"], r["assess"][:300],
             "", "", ""] for r in POOL]
    rows = [r[:10] + [f'=IF(B{i}="","",IF(COUNTIF(Pipeline!$C:$C,B{i})>0,"In Pipeline",IF(COUNTIF(\'Other Roles\'!$C:$C,B{i})>0,"In Other Roles","")))'] + r[10:] for i, r in enumerate(rows, 2)]
    # columns: A name B url C title D company E loc F tier G why H skills I source J target K tracked L prior assessment M decision N notes O reviewer
    rows = [r[:11] + [r[11], r[12], r[13]] for r in rows]
    hp = ["Candidate", "LinkedIn", "Title", "Company", "Location", "Tier", "Why / flags", "Headline / skills", "Source tab", "Target co?", "Already tracked?",
          "Prior assessment", "Decision", "Notes"]
    rows = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13]] for r in rows]
    wp = sheet(wb, "Sourcing Pool", hp, [24, 36, 32, 26, 12, 6, 40, 40, 26, 9, 16, 50, 20, 30], rows, freeze="C2", tab="D97706")
    dv_list(wp, "M", "=Lists!$K$2:$K$4", nPool); dv_list(wp, "F", "=Lists!$M$2:$M$5", nPool)
    wp.conditional_formatting.add(f"A2:N{nPool}", FormulaRule(formula=['$F2="X"'], font=Font(color="9AA3AF")))
    wp.conditional_formatting.add(f"K2:K{nPool}", FormulaRule(formula=['LEN($K2)>0'], font=Font(bold=True, color="B91C1C")))

    # ---------------- Other roles
    nO = len(O) + 1
    rows = [[r["name"], r["req"], r["url"], r["title"], r["company"], r["status"], r["stage"], r["level_note"], r["owner"], r["notes"], r["srcs"], r["legacy"]] for r in O]
    wo = sheet(wb, "Other Roles", ["Candidate", "Belongs to req", "LinkedIn", "Title", "Company", "Status (legacy-normalised)", "Stage", "Level assessment",
                                   "Owner", "Notes", "Source tabs", "Legacy status → stage"], [24, 22, 36, 30, 24, 22, 20, 24, 16, 80, 28, 44], rows, freeze="C2", tab=GREY)

    # ---------------- Target companies
    seen = collections.OrderedDict()
    for c in target_cos: seen.setdefault(c.strip(), "f4 sourcing list")
    for r in P + POOL + O:
        c = (r["company"] or "").strip()
        if c and c not in ("—", "-"): seen.setdefault(c, "seen in candidates")
    crow = []
    for i, (c, how) in enumerate(seen.items(), 2):
        tier, view = market_for(c)
        assess = next((v for k, v in company_assess.items() if k.lower() == c.lower()), "")
        crow.append([c, tier, view or assess[:300], how,
                     f'=COUNTIF(Pipeline!$E:$E,"*"&A{i}&"*")', f'=COUNTIF(\'Sourcing Pool\'!$D:$D,"*"&A{i}&"*")', "Sourced" if c in {r["company"] for r in P + POOL} else "To source", ""])
    crow.sort(key=lambda r: (r[1] == "", r[1], r[0].lower()))
    for i, r in enumerate(crow, 2):
        r[4] = f'=COUNTIF(Pipeline!$E:$E,"*"&A{i}&"*")'; r[5] = f'=COUNTIF(\'Sourcing Pool\'!$D:$D,"*"&A{i}&"*")'
    wc = sheet(wb, "Target Companies", ["Company", "Market tier (recruiter view)", "Market view / assessment", "Origin", "In Pipeline", "In Pool", "Sourcing status", "Notes"],
               [28, 30, 80, 18, 11, 9, 16, 30], crow, freeze="B2", tab="7C3AED")
    dv_list(wc, "G", "=Lists!$L$2:$L$5", len(crow) + 1)

    # ---------------- Dashboard
    d = wb.create_sheet("Dashboard", 0); d.sheet_properties.tabColor = NAVY
    d["A1"] = "(Sr.) Engineering Manager – Armenia · Growth AI"; d["A1"].font = Font(name="Arial", size=16, bold=True, color=NAVY)
    d["A2"] = f"Restart tracker · restart date {RESTART} · all numbers are live formulas on the Pipeline tab"; d["A2"].font = Font(name="Arial", size=10, color=GREY)
    d["A4"] = "Funnel by stage"; d["A4"].font = Font(name="Arial", size=12, bold=True)
    hdr = ["Stage", "Active", "Paused (hold / nurture)", "Closed", "Total ever", "Reached this stage (cumulative)"]
    for j, h in enumerate(hdr, 1):
        c = d.cell(5, j, h); c.font, c.fill = HFONT, HFILL; c.alignment = Alignment(wrap_text=True, vertical="center")
    for k, s in enumerate(STAGE_NAMES):
        r = 6 + k
        d.cell(r, 1, s)
        d.cell(r, 2, f'=COUNTIFS(Pipeline!$H$2:$H${nP},$A{r},Pipeline!$I$2:$I${nP},"Active")+COUNTIFS(Pipeline!$H$2:$H${nP},$A{r},Pipeline!$I$2:$I${nP},"No Response")')
        d.cell(r, 3, f'=COUNTIFS(Pipeline!$H$2:$H${nP},$A{r},Pipeline!$I$2:$I${nP},"On Hold")+COUNTIFS(Pipeline!$H$2:$H${nP},$A{r},Pipeline!$I$2:$I${nP},"Nurture*")')
        d.cell(r, 4, f'=E{r}-B{r}-C{r}')
        d.cell(r, 5, f'=COUNTIF(Pipeline!$H$2:$H${nP},$A{r})')
        d.cell(r, 6, f'=SUM(E{r}:E${5 + len(STAGE_NAMES)})')
    last = 5 + len(STAGE_NAMES)
    d.cell(last + 1, 1, "Total").font = Font(bold=True)
    for j in range(2, 6): d.cell(last + 1, j, f"=SUM({L(j)}6:{L(j)}{last})").font = Font(bold=True)
    d["H4"] = "Needs attention"; d["H4"].font = Font(name="Arial", size=12, bold=True)
    kp = [("OVERDUE actions", f'=COUNTIF(Pipeline!$S$2:$S${nP},"OVERDUE")'), ("Due this week", f'=COUNTIF(Pipeline!$S$2:$S${nP},"This week")'),
          ("Duplicate LinkedIn rows", f'=COUNTIF(Pipeline!$Y$2:$Y${nP},"DUPLICATE")'), ("Rows with data check flags", f'=COUNTIF(Pipeline!$X$2:$X${nP},"?*")'),
          ("Priority A, open", f'=COUNTIFS(Pipeline!$J$2:$J${nP},"A")'), ("Pool: Tier A not yet promoted", f"=COUNTIFS('Sourcing Pool'!$F$2:$F${nPool},\"A\",'Sourcing Pool'!$M$2:$M${nPool},\"\")")]
    for k, (a, b) in enumerate(kp):
        d.cell(5 + k, 8, a); d.cell(5 + k, 9, b).font = Font(bold=True)
    d["A20"] = "By restart wave"; d["A20"].font = Font(name="Arial", size=12, bold=True)
    for j, h in enumerate(["Wave", "Candidates", "What to do"], 1):
        c = d.cell(21, j, h); c.font, c.fill = HFONT, HFILL
    for k, (w, desc) in enumerate(WAVES):
        d.cell(22 + k, 1, w); d.cell(22 + k, 2, f'=COUNTIF(Pipeline!$K$2:$K${nP},A{22 + k})'); d.cell(22 + k, 3, desc)
    d["A30"] = "By status"; d["A30"].font = Font(name="Arial", size=12, bold=True)
    for j, h in enumerate(["Status", "Candidates", "Class"], 1):
        c = d.cell(31, j, h); c.font, c.fill = HFONT, HFILL
    for k, (s, cl, _) in enumerate(STATUSES):
        d.cell(32 + k, 1, s); d.cell(32 + k, 2, f'=COUNTIF(Pipeline!$I$2:$I${nP},A{32 + k})'); d.cell(32 + k, 3, cl)
    d["H13"] = "By level fit (open pipeline)"; d["H13"].font = Font(name="Arial", size=12, bold=True)
    for k, lv in enumerate(LEVELS):
        d.cell(14 + k, 8, lv); d.cell(14 + k, 9, f'=COUNTIFS(Pipeline!$F$2:$F${nP},H{14 + k},Pipeline!$K$2:$K${nP},"W*")')
    for col, w in zip("ABCDEFGHI", [26, 12, 20, 10, 11, 20, 4, 34, 12]): d.column_dimensions[col].width = w
    for row in d.iter_rows(min_row=5, max_row=40):
        for c in row:
            if c.font == Font(): c.font = BFONT
    ch = BarChart(); ch.type = "bar"; ch.style = 10; ch.title = "Open candidates by stage"
    ch.add_data(Reference(d, min_col=2, min_row=5, max_row=5 + len(STAGE_NAMES)), titles_from_data=True); ch.set_categories(Reference(d, min_col=1, min_row=6, max_row=5 + len(STAGE_NAMES)))
    ch.height, ch.width = 9, 16; ch.y_axis.delete = False; ch.x_axis.delete = False
    d.add_chart(ch, "K4")

    # ---------------- Playbook (README)
    pb = wb.create_sheet("Playbook", 1); pb.sheet_properties.tabColor = NAVY
    pb.column_dimensions["A"].width = 30; pb.column_dimensions["B"].width = 120
    lines = [("HOW TO USE THIS TRACKER", None),
             ("One row per person", "LinkedIn URL is the unique key. Pipeline = people considered for (Sr.) EM / EM. 'Other Roles' = Director+/Staff+/Mobile people kept for their own reqs (nothing deleted). 'Sourcing Pool' = leads never worked – promote a lead by copying its row into Pipeline and setting Decision."),
             ("Daily (10 min)", "Filter Due Flag = OVERDUE / This week → do the Next Action → set Touches (+1), Last Touch (today), and move Stage/Status. Auto Follow-up fills itself from the cadence."),
             ("Weekly (Mon, 30 min)", "Dashboard review: funnel counts, overdue, Priority A. Promote 10–15 Tier-A leads from Sourcing Pool. Close out anyone at 4 touches with no reply → Status 'No Response' (recycle in 90 days)."),
             ("Rules", "Never type free text in Stage/Status/Wave – use the dropdowns. Add context to 'Latest Notes' with date + initials. Never delete a row: close it with a Status so we never re-approach the same person blindly."),
             ("", None),
             ("RESTART WAVES", None)] + [(w, d_) for w, d_ in WAVES] + [("", None), ("OUTREACH CADENCE (repeating follow-ups)", None),
             ("Touch 1", "Personal LinkedIn message (role, Growth AI scope, Yerevan office). Auto follow-up in 3 days."),
             ("Touch 2 (+3d)", "Short follow-up with one concrete hook (AI initiative / team / scope). Next in 4 days."),
             ("Touch 3 (+7d total)", "Alternate channel (email / referral warm intro). Next in 7 days."),
             ("Touch 4 (+14d total)", "Break-up note leaving the door open. Then Status = No Response; recycle after 90 days."),
             ("", None), ("STAGES", None)] + [(f"{i}. {s}", d_) for i, (s, d_) in enumerate(STAGES)] + [("", None), ("STATUSES", None)] + [(s, f"[{c}] {d_}") for s, c, d_ in STATUSES] + [
             ("", None), ("ARMENIA MARKET NOTES (recruiter judgment – validate)", None),
             ("Where the EM/Sr EM bench is", "Picsart, Krisp, Miro, SuperAnnotate, Podcastle (product/AI), Adobe/VMware/Broadcom/Synopsys/AMD/Siemens EDA (enterprise, deep manager layers), ConnectWise/DISQO/Synergy/Digitain (.NET/Java SaaS), EPAM/Playrix (large outsourcing/gaming benches)."),
             ("Stack fit", "Role stack is C#/.NET/Java-centric with a coding screen. EDA/semiconductor C++ managers (Synopsys, Siemens EDA, AMD, Mentor) have the title but a stack-transfer gap; test hands-on depth early at RS. Outsourcing 'EM' titles often mean resource management – verify people + product ownership."),
             ("Mobility signals", "VMware/Broadcom talent is the best source of seasoned managers right now but is stock-locked; long-tenured BetConstruct/Digitain leaders rarely move. Strong signals: recently promoted, new CEO/org change, team reorganisations, open-to-work, laid-off colleagues."),
             ("Level reality check", "Many 'Senior Manager' / 'Director' titles in Armenia map to Sr EM scope here. Directors repeatedly say they are 'Director-only' – route them to the Director req rather than burning cycles. Check Level Fit + Interest in Role before outreach."),
             ("Candidate friction to pre-empt in outreach", "Coding interview for managers (top drop-off reason: 'decided to drop after learning he has to pass coding'), in-office/no-remote policy, relocation-only candidates abroad, compensation expectations vs band. Say all three in Touch 1."),
             ("Referral channel", "Many leads are known by team members ('Lusine knows her', 'Narek knows him'). Run a weekly 15-min referral check with the Armenia leadership team before cold outreach."),
             ("", None), ("ASSUMPTIONS", None),
             ("Legacy statuses", "Statuses were carried over as of each legacy tab's last edit (dates unknown). All 'Active' rows are therefore in Wave W0 'Validate & continue' – confirm they are still live before acting."),
             ("Precedence", "When the same person appears with different statuses, the most recent consolidated source wins (Full Pipeline > FY26-27 list > Hiring NEW (Sr) EM > Director tabs > OLD EM); furthest stage reached is kept. Conflicts are flagged in Data Check."),
             ("Proposed dates", f"Next Action Dates are proposed starting {RESTART} (W0 Mon 5-Oct, W2-A 6-Oct, W1 7-Oct, W3 8-Oct). Edit freely."),
             ("Role brief", "See 'Role Brief' tab and the Google Doc 'Growth AI – (Sr.) EM Armenia: Role Brief, Pitch & Recruiter Screening Kit'. Interview stages follow the 'Engineering Interview Process – Proposal' (EM / eDir loop).")]
    for i, (a, b) in enumerate(lines, 1):
        pb.cell(i, 1, a);
        if b is None and a: pb.cell(i, 1).font = Font(name="Arial", size=12, bold=True, color="FFFFFF"); pb.cell(i, 1).fill = HFILL; pb.cell(i, 2).fill = HFILL
        else:
            pb.cell(i, 1).font = Font(name="Arial", size=10, bold=True); pb.cell(i, 2, b); pb.cell(i, 2).font = BFONT
            pb.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top"); pb.cell(i, 1).alignment = Alignment(vertical="top", wrap_text=True)

    # ---------------- Role brief
    rb = wb.create_sheet("Role Brief", 2); rb.sheet_properties.tabColor = NAVY
    rb.column_dimensions["A"].width = 28; rb.column_dimensions["B"].width = 120
    brief = [("Role", "Senior Manager, Software Engineering – Growth AI / DemandGen, Yerevan (JR112035). Reports to the Armenia Engineering Director. Grade 34. (Workday title reads 'Director, Software Engineering' – confirm level.)"),
             ("Domain", "AI-first initiative: autonomous lead generation, AI agents (Voice, SMS, Email), self-optimizing campaigns. One of six pillars of the Agentic OS. Core stack .NET / C# / ASP.NET Core / Azure / SQL Server."),
             ("Interview loop (EM / eDir)", "Screen (gate, both must pass): Technical Retrospective 45m + People Management / HM round 45m. Full Loop: System Design 45m + Cross-Org / XFN 45m. Coding 60m ONLY for Frontline Managers (not required for Sr Manager and above; unguarded AI default, HM may choose guarded). Post Loop: Behavioral with PBP + VP Bar Raiser. Binary Yes/No scoring, feedback in 24h, no single round overrides another."),
             ("Leveling signals for managers", "People Management, XFN and Behavioral are the primary leveling signals. Down-leveling to the highest level cleared is a normal outcome."),
             ("Status of the process", "'Engineering Interview Process – Proposal' is being piloted; the EM loop is still being built out. Confirm the current version with TA before quoting details to candidates."),
             ("Profile signals", "Hands-on technical credibility (retrospective on a real project), people leadership (coaching, performance, hiring), AI/ML product curiosity, ambiguity tolerance, cross-functional influence, stack adaptability."),
             ("Disqualifiers seen repeatedly", "Under 2–3 years of real EM experience; QA/test or pure delivery management; resource-manager titles without ownership; frontend/mobile-only with no backend appetite; consulting-only; Director-only ambition; relocation- or remote-only."),
             ("Comp / level band", "(fill in)"), ("Hiring manager / panel", "(fill in)"), ("Headcount / target start", "(fill in)")]
    for i, (a, b) in enumerate(brief, 1):
        rb.cell(i, 1, a).font = Font(name="Arial", size=10, bold=True); rb.cell(i, 2, b).font = BFONT
        rb.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top"); rb.cell(i, 1).alignment = Alignment(vertical="top")

    # ---------------- Cleanup log
    cl = wb.create_sheet("Cleanup Log"); cl.sheet_properties.tabColor = GREY
    raw = len(entries)
    junk = []
    for r in rows_of(f"{args.src}/f4_Hiring.xlsx", "OLD - EM")[1:]:
        if not slug(r[0]): junk.append(r[0][:60])
    st_cnt = collections.Counter(r["status"] for r in P)
    log = [("Legacy rows read (all tabs, incl. duplicates)", raw), ("Unique people after de-duplication", len(groups)),
           ("→ Pipeline ((Sr) EM / EM track)", len(P)), ("→ Other Roles (Director+, Staff+ IC, Mobile IC)", len(O)), ("→ Sourcing Pool (never worked)", len(POOL)),
           ("People merged from 2+ rows", sum(1 for r in P + O + POOL if r["ndup"] > 1)), ("Pipeline rows with status conflicts across tabs", sum(1 for r in P if r["conflict"])),
           ("Pipeline names derived from LinkedIn URL (OLD EM tab had no names)", sum(1 for r in P if r["name_derived"])),
           ("Pipeline rows without LinkedIn URL", sum(1 for r in P if r["no_url"])),
           ("Pool rows excluded as Tier X (out-of-country / non-EM / IC)", sum(1 for r in POOL if r["tier"] == "X"))]
    cl.append(["Metric", "Value"]); [cl.append(list(x)) for x in log]
    cl.append([]); cl.append(["Junk / malformed legacy rows skipped (OLD - EM tab, no LinkedIn URL)", ""]); [cl.append([j, ""]) for j in junk]
    cl.append([]); cl.append(["Other cleanups applied", ""])
    for t in ["LinkedIn URLs normalised to https://www.linkedin.com/in/<slug> (tracking params, trailing slashes, locale suffixes stripped)",
              "Google-Sheets markdown escapes (\\&, \\#, \\|) removed from text",
              "Free-text statuses/stages (30+ variants, e.g. 'RS', 'Recruiter Screen', 'system design', 'successfully passed system design') mapped to 11 stages × 9 statuses",
              "Person-name collisions NOT merged when LinkedIn slugs differ (e.g. two different 'Davit Petrosyan' profiles)",
              "'Hired' marker on the OLD EM tab carried over but flagged for verification"]: cl.append([t, ""])
    cl.column_dimensions["A"].width = 110; cl.column_dimensions["B"].width = 12
    for c in cl[1]: c.font, c.fill = HFONT, HFILL

    order = ["Dashboard", "Playbook", "Role Brief", "Pipeline", "Sourcing Pool", "Target Companies", "Other Roles", "Cleanup Log", "Lists"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.save(args.out)
    print("saved", args.out)
    print("raw", raw, "unique", len(groups), "P", len(P), "O", len(O), "POOL", len(POOL))
    print("pipeline status", st_cnt.most_common()); print("wave", collections.Counter(r["wave"] for r in P).most_common())
    print("stage", collections.Counter(r["stage"] for r in P).most_common()); print("prio", collections.Counter(r["priority"] for r in P).most_common())
    print("pool tier", collections.Counter(r["tier"] for r in POOL).most_common(), "other req", collections.Counter(r["req"] for r in O).most_common())
    return P, O, POOL


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--src", required=True); ap.add_argument("--out", default="output/EM_Armenia_Hiring_Tracker.xlsx")
    build(ap.parse_args())
