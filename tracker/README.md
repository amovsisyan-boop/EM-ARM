# (Sr.) EM Armenia – hiring tracker builder

Builds one clean tracker workbook from four legacy Google Sheets (full pipeline, EM sourcing lists,
Staff+/leadership lists, EM & Director hiring tabs).

**Candidate data is confidential and is deliberately not stored in this repository.**
`output/`, `src/`, `*.xlsx` and `*.csv` are git-ignored. Only the build script and this process doc are tracked.

## Rebuild

```bash
pip install openpyxl beautifulsoup4 lxml
# put exports of the four legacy sheets in <src>: f1.csv, f2_EM.xlsx, f3_Arm.xlsx, f4_Hiring.xlsx
python3 tracker/build_tracker.py --src <src> --out output/EM_Armenia_Hiring_Tracker.xlsx
```

Import into Google Sheets: **File → Import → Upload → Replace spreadsheet** (dropdowns, formulas and
conditional formatting carry over). Then share the sheet with the hiring team.

## Workbook tabs

| Tab | Purpose |
|---|---|
| Dashboard | Live funnel by stage/status/wave, overdue actions, level mix |
| Playbook | How to run the tracker, cadence, stage/status definitions, Armenia market notes, assumptions |
| Role Brief | Role summary + disqualifiers seen in legacy notes (paste the JD here) |
| Pipeline | One row per (Sr) EM / EM candidate – the working tab |
| Sourcing Pool | Leads never worked, tiered A/B/C/X, with "already tracked?" check |
| Target Companies | Company list with market tier and live candidate counts |
| Other Roles | Director+, Staff+ IC, Mobile and CTO-level people kept for their own reqs |
| Cleanup Log | What was merged/normalised and what was skipped |
| Lists | Dropdown sources (stages, statuses, waves, cadence) |

## Stages (ordered)

Sourced → Profile Review → Outreach Sent → Replied / Interested → Recruiter Screen → HM Screen →
CP1 – Coding → System Design → CP3 – Final Interview → Offer → Hired

## Statuses

Active · No Response · On Hold · Nurture – Re-engage Later · Not Interested · Not a Fit (Profile) ·
Rejected by ST · Withdrew · Hired

## Restart waves

W0 Validate & continue · W1 Re-engage warm · W2 Fresh outreach · W3 Follow-up (no reply) · Hold / closed · Done

## Follow-up cadence (repeating)

Touch 1 → +3 days → Touch 2 → +4 days → Touch 3 → +7 days → Touch 4 (break-up) → Status `No Response`, recycle after 90 days.
`Auto Follow-up`, `Due Date` and `Due Flag` are formulas driven by `Touches` and `Last Touch`.

## Merge rules

- LinkedIn slug is the identity key. Rows without a URL attach to a person only on an unambiguous name match.
- Same name + same employer with different slugs are merged; same name with a different employer is flagged, not merged.
- Status precedence: Full Pipeline > FY26-27 list > Hiring NEW (Sr) EM > Director tabs > OLD EM, but a concrete
  outcome (rejected, withdrew, not interested, nurture) beats a placeholder "Active / Profile Review".
- Furthest stage reached across all tabs is kept. Cross-tab status conflicts are flagged in `Data Check`.
