# dcourts-judge-scraper

Counts the judicial officers actually posted to each district court in India by
reading the courts' own websites (`*.dcourts.gov.in`), and splits them into
those who preside over a courtroom and those who do not.

No such per-district figure is published anywhere. Official judicial strength
is reported at state level and does not distinguish officers who preside over a
courtroom from those doing administrative work, so the only source of truth is
each court's own list of judges — hundreds of separate pages of names and
designations with nothing that adds them up. A run over 250 district courts
across 12 states collected ~9,300 officer records.

## Why this is harder than it looks

The obvious approach — fetch `/list-of-judges/`, find the biggest table, count
the rows — is wrong, quietly, in five different ways. Each of these was found
by checking a scraped number against a known one and asking why they differed.

**One table per court complex, not one per page.** The pages render a
`<div class="whoswho">` block for every court complex in the district, each
with its own table. Taking the largest returns a single complex and silently
drops the rest. Ernakulam reads 21 officers that way; it has 78.

**Two different pagination schemes.** Blocks are cut off at 20 rows. Blocks
that belong to a taxonomy continue through the theme's own AJAX endpoint
(`POST /wp-admin/admin-ajax.php`, `action=judge_pages`, with a nonce lifted
from the page). Uncategorised blocks instead continue at ordinary
`/list-of-judges/page/N/` URLs. A scraper that handles only one of the two
under-counts every district that uses the other — Muzaffarpur reads 20 instead
of 54.

**Vacant posts are listed as if they were people.** `Vacant`,
`Post Kept Vacant ( Incharge ) <designation>`, and `( FAC ) <designation>`
placeholders all appear as named rows. The last two name the officer's
substantive post, so counting them also double-counts the officer covering the
vacancy.

**Subdomains collide across states.** `aurangabad.dcourts.gov.in` is
*Maharashtra's* Aurangabad. Bihar's is `aurangabad.bih.dcourts.gov.in`. The
guard here fetches every site's home-page `<title>` and checks it against the
state the input says it should be — across 250 districts that caught exactly
one wrong site, but it was wrong by 47 officers.

**Officer names are not always in the column you expect.** Some states put the
designation first; some add Room Number and Location columns. Rows are read by
their header labels rather than by position.

## Pipeline

Four stages, deliberately separate, so that re-deciding what counts as "bench"
does not mean re-scraping 250 websites.

| | | |
|---|---|---|
| 1 | `scrape_rows.py` | one CSV row per officer: state, district, url, block, name, designation |
| 2 | `classify.py` | assigns scope and role to each row, aggregates per district |
| 3 | `write_workbook.py` | writes the counts into a copy of a target workbook |
| 4 | `restructure.py` | collapses rows that are not judicial districts, rebuilds totals |

Stage 1 caches every fetch to disk, so stages 2–4 re-run in about a second.

## Counting rules

Stage 2 is where the judgement lives. Two decisions drive every number:

**What counts as posted.** Filled posts, not people — an officer holding
additional charge of a second court is counted once per court, because the
question being asked is how many courtrooms are running. Excludes vacant posts,
the High Court judges the district pages display at the top, officers on
deputation to tribunals elsewhere, and FAC placeholder rows.

**What counts as bench.** Anyone who presides over a courtroom, whatever else
they also do. Decided in three tiers on the designation text:

1. names a court they preside over → bench, wherever in the string it appears
2. otherwise names an administrative-only post → non-bench
3. otherwise a bare cadre rank on its own → bench

So `Chief Judicial Magistrate-cum-Secretary, DLSA` is bench — that officer sits
— while `Secretary(Sr.Civil Judge) DLSA` is not, because there the bracketed
rank is the officer's cadre grade and no court is named. Family courts, MACT,
labour courts, commercial courts and special courts all count as bench; DLSA
secretaries, registry officers, Lok Adalat chairpersons, mediation-centre
judges and court managers do not.

On the 12-state run this left 0 rows unclassified out of ~9,300.

## Validating the output

A scraper that returns a number always returns a number, so the output was
checked rather than taken on trust:

- against published state working-strength figures, for the states where every
  district was covered: within ~2% for Karnataka, exact for Chandigarh
- every count is reproducible from the page it came from, and outliers were
  re-counted by hand against the live site before being accepted

The site is not automatically right either — these lists lag reality, and at
least one under-reports (Gangtok's cause-list dropdown shows a Family Court
that never appears in its judge list). The output is best read as *what the
courts publish*, with `judges_sites.csv` recording exactly which URL each
number came from.

## Usage

```bash
pip install requests beautifulsoup4 lxml openpyxl
```

The input workbook needs a sheet with state in column A and district in
column B.

```bash
python scrape_rows.py --xlsx districts.xlsx --workers 6 --delay 0.4
python classify.py
python write_workbook.py --xlsx districts.xlsx --out districts_filled.xlsx
```

Useful flags on stage 1: `--states "Kerala,Odisha"`, `--districts
"Kerala:Ernakulam"`, `--limit 15` for a spot check, `--refresh` to bypass the
cache, `--extra-sheet` to pick up districts listed only on a second sheet.

Districts whose subdomain cannot be guessed go in the `OVERRIDES` table at the
top of `scrape_rows.py`; a first run reports them, and DNS-probing a few
candidate spellings usually settles it. About two dozen are non-obvious —
Dakshina Kannada is `dk`, Mahendragarh is `narnaul` (the court sits at the
headquarters town), Sri Muktsar Sahib is `muktsarsahib`.

`restructure.py` is workbook-specific: its `MERGE` and `DROP` tables name the
rows to collapse. Edit them before running it on anything else.

## Notes

Court websites are slow and occasionally time out; the scraper retries, backs
off, and treats a DNS failure as fatal for that host rather than trying five
more paths against it. Default settings are 6 workers with a 0.4s delay, which
is gentle enough to be polite and finishes 250 districts in a few minutes.

Only officer names and designations are collected — information those courts
publish themselves.
