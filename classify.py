#!/usr/bin/env python3
"""
Stage 2: turn judges_rows.csv into the per-district numbers for columns C and D.

Scope (the "posted" count) counts posts held by a named
officer in that judicial district. It drops:
  - High Court judges the district pages display at the top (Chief Justice,
    Administrative/Portfolio judge)
  - officers shown as on deputation to tribunals outside the district
  - vacant posts (already dropped in stage 1)
  - "FAC"/in-charge placeholder rows, where the name cell holds the covering
    officer's substantive post rather than a person

Bench (column D) counts those officers who preside over proceedings. Family
courts, MACT, labour courts, commercial courts, POCSO and other special courts
all count. Only roles that do not sit -- DLSA secretaries, registry officers,
Lok Adalat chairpersons, mediation-centre judges, OSDs and court managers --
are counted as non-bench.

    python classify.py --rows judges_rows.csv --out judges_summary.csv

Writes judges_summary.csv (one row per district) and judges_unknown.csv (every
designation the rules could not place, so the patterns can be extended).
"""

import argparse
import csv
import re
from collections import OrderedDict, Counter, defaultdict

# ------------------------------------------------------------------ scope
HIGH_COURT = re.compile(
    r"high\s+court|hon.?ble\s+(mr|mrs|ms|smt|sri|shri)?\.?\s*justice|"
    r"chief\s+justice|administrative\s+judge|portfolio\s+judge|"
    r"inspecting\s+judge", re.I)
DEPUTATION = re.compile(
    r"\bon\s+deputation\b|officers?\s+on\s+deputation|"
    r"\brpsc\b|revenue\s+board|board\s+of\s+revenue|public\s+service\s+commission",
    re.I)

# The name cell holds a post, not a person: "( FAC ) III Additional Civil
# Judge, Kakinada" means the post is vacant and covered in full additional
# charge by someone who is already listed against their own court.
NOT_A_PERSON = re.compile(
    r"^\(?\s*(fac|f\.a\.c|i/c|in\s*-?\s*charge)\s*\)?", re.I)
PERSONISH = re.compile(
    r"\b(sri|shri|smt|sh|ms|mr|mrs|dr|kum|kumari|thiru|md|mohd)\b\.?", re.I)
DESIGNATIONISH = re.compile(
    r"\b(judge|magistrate|munsiff|munsif|court|tribunal|registrar|secretary|"
    r"presiding\s+officer|nyayadhikari)\b", re.I)

# ------------------------------------------------------------------ role
# Roles that do not preside over proceedings.
NON_BENCH = re.compile(
    r"legal\s+services?\s+authorit|\bdlsa\b|\btlsa\b|\bslsa\b|"
    r"member\s+secretary|"
    r"\bregistrar\b|registry\s+officer|"
    r"officer\s+on\s+special\s+duty|\bosd\b|"
    r"permanent\s+lok\s+adalat|chairperson,?\s*(permanent\s+)?lok\s+adalat|"
    r"\bplapus\b|chairman\s*-\s*plapus|"
    r"mediation\s+cent(re|er)|"
    r"court\s+manager|system\s+officer|protocol\s+officer|"
    r"central\s+project\s+coordinator|\bcpc\b\s*$|"
    r"director,?\s+judicial\s+academ|judicial\s+academ", re.I)
# A DLSA secretary is written many ways; "secretary" alone is too broad
# (e.g. "Secretary, Legal Services") so it is paired with the authority.
SECRETARY_DLSA = re.compile(
    r"secretary[^|]{0,40}(dlsa|legal\s+services?|d\.l\.s\.a)|"
    r"(dlsa|legal\s+services?)[^|]{0,40}secretary", re.I)

# Roles that do preside over proceedings.
BENCH = re.compile(
    r"district\s*(&|and)\s*(addl\.?|additional\s+)?sessions?|"
    r"sessions?\s+judge|addl?\.?\s*sessions?|"
    r"district\s+judge|addl?\.?\s*district\s+judge|additional\s+district|"
    r"principal\s+(district|judge|civil|sub|munsiff|senior)|"
    r"civil\s+judge|sub\s*-?\s*judge|\bsub\s+court\b|"
    r"munsiff|munsif|"
    # "magist" rather than "magistrate": the sites also write Magistate,
    # Magistrete and Magisrate
    r"magist\w*|\bcjm\b|\bacjm\b|\bjmfc\b|\bjmic\b|\bjfcm\b|\bjmi{1,3}c\b|"
    r"excise\s+court|exclusive\s+special|special\s+\w+\s+court|"
    r"family\s+court|family\s+judge|"
    r"\bmact\b|motor\s+accident|"
    r"labour\s+court|industrial\s+tribunal|"
    r"commercial\s+court|"
    r"special\s+(court|judge)|spl\.?\s*(court|judge|addl)|"
    r"fast\s+track|\bpocso\b|\bndps\b|\bcbi\b|sc/st|scheduled\s+castes|"
    r"atrocit|vigilance|enquiry\s+commissioner|"
    r"gram\s+nyayalaya|nyayadhikari|"
    r"small\s+causes|rent\s+control|land\s+acquisition|"
    r"wakf\s+tribunal|revenue\s+(court|tribunal)|"
    r"consumer|"
    r"juvenile\s+justice|children.?s\s+court|"
    r"presiding\s+officer|\bpo\b\s*,?\s*mact|"
    r"chairman\s*\(district\s+judge\)|appellate\s+tribunal|"
    r"metropolitan\s+magistrate|\bmm\b|\bcmm\b|"
    r"judicial\s+(first|1st|i|second|2nd|ii|iii)\s*class|"
    r"court\s+of\s+(civil|small|session)", re.I)


# Rajasthan, Odisha and Bihar list most posts as initialisms only ("ADJ",
# "CJ & JM", "A.C.J.M. No. 2"). Matched after stripping periods so the many
# spellings collapse to one. Chief Justice is caught by scope_of before this
# runs, so a bare "CJ" here is safely a Civil Judge.
ABBREV = re.compile(
    r"\b(adj|asj|dasj|acjm|cjm|sdjm|sdjmc|jmfc|jmfco|jmic|jfcm|jmsc|jm|acj|cjjd|cjsd|"
    r"cj|jjb|po|mm|cmm|acmm|mact|dj|addl\s*dj|spl\s*jm|jmsecondclass)\b",
    re.I)


# A court someone actually presides over. Checked FIRST and regardless of
# where it appears, because an officer who sits in a courtroom counts as bench
# however many administrative duties are bolted on after it:
#   "Registrar-cum-ASJ, Civil Courts, Cuttack"      -> bench (sits as ASJ)
#   "Secretary DLSA -cum- Chief Judicial Magistrate" -> bench (sits as CJM)
# Deliberately excludes the bare cadre ranks -- Civil Judge, Sub Judge, Senior
# Civil Judge. Those name an officer's grade, not a courtroom, which is why
# "Secretary(Sr.Civil Judge) DLSA" and "DLSA SECRETARY (SUB JUDGE)" are not
# bench: nothing there is a court. A cadre rank standing on its own still is
# bench, and gets picked up by BENCH at the tier below.
PRESIDES = re.compile(
    r"magist\w*|\bcjm\b|\bacjm\b|\bjmfc\w*|\bjmic\b|\bjfcm\b|\bsdjm\b|\bjmsc\b|"
    r"\bmm\b|\bcmm\b|\bacmm\b|\bpmjjb\b|juvenile\s+justice|children.?s\s+court|"
    r"sessions?|\basj\b|\bdasj\b|\bsdj\b|"
    r"district\s+judge|\badj\b|additional\s+district|addl\.?\s*district|"
    r"family\s+court|family\s+judge|"
    r"\bmact\b|motor\s+accident|labour\s+court|industrial\s+tribunal|"
    r"appellate\s+tribunal|wakf\s+tribunal|waqf\s+tribunal|revenue\s+court|"
    r"rent\s+(control|appellate)|land\s+acquisition|"
    r"commercial\s+court|consumer|"
    r"special\s+(court|judge)|spl\.?\s*(court|judge|addl)|special\s+\w+\s+court|"
    r"track\s+court|fast\s+track|\bpocso\b|\bndps\b|\bcbi\b|\bpmla\b|sc/st|"
    r"atrocit|vigilance|excise\s+court|exclusive\s+special|\bni\s*act\b|"
    r"designated\s+court|enquiry\s+commissioner|"
    r"gram\s*a?\s*nyayalaya|nyayadhikar\w*|nyayalaya|gramanyalya|"
    r"presiding\s+officer|chairman\s*\(district\s+judge\)|"
    r"judge,?\s*(court\s+of\s+)?small\s+causes|small\s+causes\s+(judge|court\s+judge)|"
    r"court\s+of\s+(civil|session)|"
    r"judicial\s+(first|1st|i|second|2nd|ii|iii)\s*class",
    re.I)

LOOSE_BENCH = re.compile(
    r"\bjudge\b|\bjudges\b|\bcourt\b|tribunal|nyayalaya|nyayalya|nyayadhikar|"
    r"gramanyalya|grama?\s*nyaya|\bni\s*act\b|\bacd\b|\bdjs\b|"
    r"जज|न्याय|अदालत",   # जज / न्याय / अदालत
    re.I)


def strip_dots(s):
    """"A.C.J.M. No. 2" -> "ACJM No 2" (drop the dots, keep the spaces)."""
    return re.sub(r"\s+", " ", s.replace(".", "").replace(" ", " ")).strip()


def scope_of(block, name, desig):
    hay = f"{block} | {desig}"
    if DEPUTATION.search(block) or DEPUTATION.search(desig):
        return "deputation"
    if HIGH_COURT.search(hay) or HIGH_COURT.search(name):
        return "high_court"
    if NOT_A_PERSON.match(name):
        return "placeholder"
    # name cell that is plainly a designation and carries no honorific
    if DESIGNATIONISH.search(name) and not PERSONISH.search(name):
        return "placeholder"
    return "district"


def _first(s, *pats):
    """position of the earliest match of any pattern, or None."""
    hits = [m.start() for m in (p.search(s) for p in pats) if m]
    return min(hits) if hits else None


def role_of(block, desig):
    """
    bench     = presides over a courtroom, whatever else they also do
    non_bench = presides over no courtroom at all

    Decided on the designation, in three tiers:
      1. names a court they preside over -> bench, wherever it appears
      2. otherwise names an administrative-only post -> non_bench
      3. otherwise a bare cadre rank ("Civil Judge (Jr. Divn.)") -> bench
    """
    d = strip_dots(desig)
    if PRESIDES.search(d):
        return "bench"
    if SECRETARY_DLSA.search(d) or NON_BENCH.search(d):
        return "non_bench"
    if BENCH.search(d) or ABBREV.search(d):
        return "bench"

    # nothing usable in the designation: fall back to the block heading
    hay = strip_dots(f"{block} | {desig}")
    if PRESIDES.search(hay):
        return "bench"
    if SECRETARY_DLSA.search(hay) or NON_BENCH.search(hay):
        return "non_bench"
    if BENCH.search(hay) or ABBREV.search(hay):
        return "bench"
    # Last resort, and only because every non-bench role has already been
    # ruled out above: a designation that still names a judge, a court or a
    # tribunal is someone who sits in one. Catches local spellings
    # ("Itirnery Court", "Gramanyalya"), Hindi-script entries and one-off
    # special courts (NI Act, Waqf, ACD).
    if LOOSE_BENCH.search(hay):
        return "bench"
    return "unknown"


def norm_person(name):
    n = re.sub(r"\(.*?\)", " ", name.lower())
    n = re.sub(r"\b(sri|shri|smt|sh|ms|mr|mrs|dr|kum|kumari|thiru|hon.?ble|"
               r"justice|md|mohd)\b\.?", " ", n)
    n = re.sub(r"[^a-z ]", " ", n)
    return re.sub(r"\s+", " ", n).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", default="judges_rows.csv")
    ap.add_argument("--sites", default="judges_sites.csv")
    ap.add_argument("--out", default="judges_summary.csv")
    ap.add_argument("--unknown-out", default="judges_unknown.csv")
    ap.add_argument("--classified-out", default="judges_rows_classified.csv")
    args = ap.parse_args()

    rows = list(csv.DictReader(open(args.rows, encoding="utf-8-sig")))
    sites = {(s["state"], s["district"]): s
             for s in csv.DictReader(open(args.sites, encoding="utf-8-sig"))}

    order, agg = OrderedDict(), defaultdict(lambda: {
        "posts": 0, "bench": 0, "non_bench": 0, "unknown": 0,
        "high_court": 0, "deputation": 0, "placeholder": 0, "names": set()})
    unknown_rows, classified = [], []

    for r in rows:
        key = (r["state"], r["district"])
        order.setdefault(key, len(order))
        a = agg[key]
        sc = scope_of(r["block"], r["name"], r["designation"])
        role = ""
        if sc != "district":
            a[sc] += 1
        else:
            role = role_of(r["block"], r["designation"])
            a["posts"] += 1
            a[role] += 1
            a["names"].add(norm_person(r["name"]))
            if role == "unknown":
                unknown_rows.append(r)
        classified.append({**r, "scope": sc, "role": role})

    for key in sites:
        order.setdefault(key, len(order))

    out = []
    for key in sorted(order, key=lambda k: order[k]):
        state, district = key
        a = agg.get(key)
        site = sites.get(key, {})
        if a is None:
            out.append({"state": state, "district": district,
                        "officers_posted": "", "bench": "", "non_bench": "",
                        "unclassified": "", "distinct_persons": "",
                        "excluded_high_court": "", "excluded_deputation": "",
                        "excluded_placeholder": "", "url": site.get("url", ""),
                        "status": site.get("status", "no_data"),
                        "note": site.get("note", "")})
            continue
        out.append({
            "state": state, "district": district,
            "officers_posted": a["posts"], "bench": a["bench"],
            "non_bench": a["non_bench"], "unclassified": a["unknown"],
            "distinct_persons": len(a["names"]),
            "excluded_high_court": a["high_court"],
            "excluded_deputation": a["deputation"],
            "excluded_placeholder": a["placeholder"],
            "url": site.get("url", ""), "status": site.get("status", ""),
            "note": site.get("note", "")})

    cols = ["state", "district", "officers_posted", "bench", "non_bench",
            "unclassified", "distinct_persons", "excluded_high_court",
            "excluded_deputation", "excluded_placeholder", "url", "status",
            "note"]
    with open(args.out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(out)

    with open(args.classified_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()) + ["scope", "role"])
        w.writeheader()
        w.writerows(classified)

    with open(args.unknown_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["count", "designation", "example_state", "example_district"])
        c = Counter(r["designation"] for r in unknown_rows)
        first = {}
        for r in unknown_rows:
            first.setdefault(r["designation"], (r["state"], r["district"]))
        for d, n in c.most_common():
            w.writerow([n, d, first[d][0], first[d][1]])

    tot = sum(r["officers_posted"] for r in out if r["officers_posted"] != "")
    bench = sum(r["bench"] for r in out if r["bench"] != "")
    nb = sum(r["non_bench"] for r in out if r["non_bench"] != "")
    unk = sum(r["unclassified"] for r in out if r["unclassified"] != "")
    print(f"districts with data : {sum(1 for r in out if r['officers_posted'] != '')}")
    print(f"officers posted     : {tot}")
    print(f"  bench             : {bench}")
    print(f"  non-bench         : {nb}")
    print(f"  unclassified      : {unk}  ({unk / max(tot, 1):.1%})")
    print(f"excluded high court : {sum(r['excluded_high_court'] for r in out if r['excluded_high_court'] != '')}")
    print(f"excluded deputation : {sum(r['excluded_deputation'] for r in out if r['excluded_deputation'] != '')}")
    print(f"excluded placeholder: {sum(r['excluded_placeholder'] for r in out if r['excluded_placeholder'] != '')}")
    print(f"wrote {args.out}, {args.classified_out}, {args.unknown_out}")


if __name__ == "__main__":
    main()
