# Architecture

This document explains how the Sparta Global Data Concepts site is put together: why it's built this way, how the
pieces fit, and how to make changes safely. It's written for a developer picking the project up for the first time.
For setup and editing content, start with the [README](README.md).

---

## 1. The big picture

The site is a **single static HTML file** with an **Excel formula engine inside it**, plus two download files.

```
 Source (src/, data/)  ──►  npm run build  ──►  dist/index.html (one file, ~1.4 MB)  ──►  GitHub Pages
      you edit this       checks, then inlines     + dist/downloads/*.xlsx, *.zip
                          every file into one
```

- **Everything runs in the browser.** `src/engine.js` parses and evaluates Excel formulas against the Sparta Leisure data,
  which is embedded in the page as JSON.
- **No server, no framework, no dependencies.** The UI is plain JavaScript that builds DOM elements with a small `h()`
  helper, the same pattern as the SQL Practice Lab.
- **Learner progress lives in the browser** (`localStorage`, key `dataConcepts.v1`).

### Why these choices

| Decision | Why |
|---|---|
| **Our own formula engine** | Learners need to type real Excel formulas and be marked instantly, on any device, without installing anything. Spreadsheet libraries are large, and the good ones are GPL or paid. The engine covers the ~110 functions the course uses, in about 600 lines, and it was checked cell for cell against LibreOffice. |
| **Practise here *and* In Excel** | The browser grid makes practice frictionless. Excel itself is what learners will use at work, and the only place to practise menus such as sorting, filtering, pivot tables and charts. Both modes mark against the same expected answer. |
| **Marking by value, not formula text** | There are many correct formulas for most tasks (VLOOKUP, XLOOKUP, INDEX/MATCH…). Comparing what a formula *returns*, on every row of a filled column, accepts all of them. |
| **Answers come from model formulas** | Content stores `sol:'=COUNTIF(J:J,"Active")'`, never the number. The expected value is worked out at run time, so it can't drift from the data. |
| **Generated, made-up data** | No licensing or privacy questions, the exact problems the lessons need (skew, outliers, a real correlation, planted quality issues), and a fixed seed so it can be rebuilt identically. |
| **One self-contained file** | Works from GitHub Pages, a double-click, SharePoint or an LMS, with no paths to configure (except the two downloads). |
| **Hash routing** (`#/lesson/lookups`) | Works in a single file on any host, and the back button works. |

---

## 2. Pages and routing

`route()` in `src/app.js` reads the address after `#`, shows one `<section class="view">` from `src/index.html`, and
calls its draw function. After every route, `renderChrome()` draws the breadcrumb, the previous/next bar and the lesson
footer (completion state, **Mark lesson as complete**, next lesson).

| Address | View | Drawn by |
|---|---|---|
| `#/path` | `v-path` | `renderPath()`: hero, continue card, units and lessons |
| `#/lesson/intro` | `v-intro` | `drawIntro()`: `SITE_INTRO` |
| `#/lesson/data` | `v-data` | `drawData()`: `DATA_INTRO`, the diagram (`DATA_ERD`) and a read-only sheet viewer |
| `#/lesson/capstone` | `v-capstone` | `drawCapstone()`: `CAPSTONE_HTML` and the `CAPSTONE_TASKS` checklist |
| `#/lesson/<id>` | `v-lesson` | `openLesson()` → `renderLessonTabs()`, `renderLessonBody()` |
| `#/practice` | `v-practice` | `drawPractice()`: a free grid |
| `#/quiz`, `#/quiz/<id>` | `v-quizzes`, `v-quiz` | `renderQuizList()`, `openQuiz()` (in `src/quiz.js`) |

Any element with `data-go="lessonId"` (or `path`, `practice`, `quiz/<id>`) navigates when clicked.

---

## 3. Content

Content is plain JavaScript in `src/content/`, loaded in this order: `helpers.js`, `lessons.js`, `intro.js`,
`capstone.js`, `quizzes.js`. The README explains every field; in brief:

```js
const UNITS   = [ { title, blurb, lessons:['types','formats'] }, … ];
const LESSONS = { lookups:{ title, summary, sheet, learn:`…html…`, ex:[ { p, sheet, sol, hint, fill?, col?, askRow?, vtype?, excel? } ] },
                  types:{ title, summary, learn, check:[ { q, a:[…], c, why } ] }, … };
const PATH    = UNITS.flatMap(…);   // every lesson in order
```

A lesson has exercises (`ex`), questions (`check`), a checklist (`tasks`), or a mix (Formula errors has both
exercises and questions). Lessons with `ex` show the practice grid beside the text, unless they're `noGrid` (the
Power Query unit, which is done entirely in Excel, so every exercise is a typed answer with `steps`, `ask` and `how`).
`kind: 'intro' | 'data' | 'capstone'` gives a lesson its own page.

**Optional units.** `optional: true` on a unit adds an Optional badge. `REQUIRED` in `app.js` (every lesson not in an
optional unit) drives "Continue", the progress trail and the lesson and exercise counts on the home page.

**Helpers in lesson text** (`src/content/helpers.js`): `F(formula, sheet, {fill, col, note})` adds a runnable formula
example, `TIP()` a tip, `NOTE()` a callout, `TABS([{label, html}])` tabs, `STEPS([…])` a numbered list, and
`W('name')` an interactive piece from `WIDGETS` in `app.js`. `FIG_NORMAL`, `FIG_SKEW` and `FIG_ETL` are static SVG
figures.

**Option order.** Right answers are written in any position, and a fixed rotation (end of `lessons.js` for Check
yourself questions, the `Q()` helper in `quizzes.js`) spreads them across a, b, c and d. `npm run check` prints the
spread for each quiz.

---

## 4. The formula engine (`src/engine.js`)

No DOM code, so `scripts/check.mjs` runs exactly the same engine in Node.

### The workbook

`makeWorkbook(WORKBOOK)` wraps the embedded data. A sheet is `{ name, cols:[{name, type}], rows:[[…]] }`. As in
Excel, **row 1 is the header** and data starts on row 2. `cellValue(sheet, row, col)` uses 1-based Excel coordinates.
Values are numbers, strings, booleans, `null` (blank) or `XErr`. **Dates are Excel serial numbers** (days since
30/12/1899, with times as fractions), so date arithmetic and functions behave exactly as in Excel. Each column's
`type` (`date`, `datetime`, `money`, `bool`…) only affects how it's displayed.

### Evaluating a formula

```
 "=VLOOKUP(H2,Plans!$A$2:$C$6,3,FALSE)"
   │ tokenize()   strings, numbers, references (Sheet!A1, $A$1, A:A, 'Sheet name'!B2:C9), functions, operators
   │ parse()      recursive descent with Excel's precedence: comparison < & < + − < * / < ^ < % < unary −
   ▼ evaluate()   walks the tree with a context { sheet, row, col, dr, dc }
```

- **Filling down.** `runFormula(wb, formula, target, {dr})` shifts every *relative* reference by `dr` rows, which is
  exactly what happens when Excel fills a formula down. `$` references don't move. This is how a forgotten `$` on a
  lookup table produces #N/A lower down, as it would in Excel.
- **Ranges** are lazy `Range` objects. Whole columns (`E:E`) stop at the last row of data, so they're fast.
  Operations on ranges (`(H2:H100="PREM")*1`) produce `Arr` values, which `SUMPRODUCT` and friends consume.
  A formula whose final result is several values returns a `#SPILL` explanation.
- **Lazy functions.** `IF`, `IFS`, `IFERROR`, `IFNA` and `SWITCH` only evaluate the branch they need, as Excel does.
- **Criteria** (`COUNTIFS`, `SUMIFS`…) are parsed by `makeCriterion()`: operators, wildcards, dates and blanks. A
  range as the criteria returns one result per cell, as in Excel 365.
- **Errors carry a reason.** `XErr` has an Excel code (`#N/A`) and a plain-English `why`, which the grid shows.
  Unknown functions get a spelling hint, and functions that only make sense in Excel (`FILTER`, `UNIQUE`,
  `GETPIVOTDATA`) say so (`NOT_HERE`).
- **UK conventions.** `DATEVALUE` and typed answers read `dd/mm/yyyy`, ISO dates, and `12 Jun 1990`.

To add a function, add a `def('NAME', minArgs, maxArgs, (args, ctx) => …)` line in the function library section. Use
`scalar()` for single-value arguments, `nums()` to collect numbers the way SUM does, and return an `ERR(code, why)`
for errors.

### Marking (`sameValue`, `readTyped`)

- **Practise here** (`markFormula()` in `app.js`): the learner's formula runs in the answer cell (row 2 of the first
  empty column). For `fill` exercises it runs on rows 2 to 301, and every row must match the model formula's result.
  The message names the first row that differs, and explains `$` signs when row 2 is right but a later row isn't.
- **In Excel** (`markTyped()`): `readTyped()` reads `£1,234.50`, `35%`, `0.35`, `31/12/2025`, `TRUE` or text.
- **Numbers** match within 0.005, or, for typed answers, when the learner's value equals the expected one rounded to the
  same number of decimal places (at least 1 for a non-whole answer). **Text** ignores capitals and surrounding spaces.

---

## 5. The practice grid (`makeSheetView` in `app.js`)

One component drives the lesson grid, Excel practice, the data pop-up in quizzes, and the read-only viewers.

| Part | Behaviour |
|---|---|
| Formula bar | Enter runs it. **F4** cycles `A1 → $A$1 → A$1 → $A1` on the reference at the cursor |
| Clicking while typing | Adds the cell (or column, from its letter) to the formula, with the sheet name if it's another sheet. Clicking again replaces it, and Shift-click makes a range, as in Excel |
| Answer column | Pink, to the right of the data. Single answers appear in row 2. **Fill down** shows the result in every visible row |
| Highlights | References in the formula are tinted in the grid, a different colour for each reference |
| Sheet tabs | Switch the view. The pink dot marks the sheet the formula belongs to |
| Rows | The first 200 are drawn, with **Show more rows** for the rest. Formulas always use every row |

`set({sheet, col, fill, vtype, formula, run})` points the grid at a task. `onRun` lets the lesson mark the formula.

---

## 6. Interactive pieces and diagrams

`WIDGETS` in `app.js` maps a name to a function that draws into a `<div data-widget>` (from `W()` in content). They
compute from the embedded data, so they always match it: `histogram`, `scatter`, `anscombe`, `truncate`,
`colourblind` (SVG `feColorMatrix` filters simulate protanopia, deuteranopia and greyscale), `formats`, `sheet`,
`schema`, `typerecord` and `typecards` (from `DATA_TYPES`), and `erdexplain` (from `ERD_RELS`).

The Data visualisation unit's `gallery`, `makeovers`, `heatmap` and `bignumbers` share a small chart toolkit
(`hbars`, `lines`, `pie`) and one set of totals, `AGG()`, worked out once from the workbook: visits by site and
weekday, monthly visits, active members by site and plan, bookings by class, and so on. Charts are hand-drawn SVG with
no chart library, use the Okabe-Ito colour-blind-safe palette where categories need colours, and label values
directly instead of relying on a legend.

`renderERD()` draws entity relationship diagrams with crow's-foot ends, in the same style as the SQL Practice Lab.
Models are plain data: `DATA_ERD`, `SCHEMA_STAR` and `SCHEMA_SNOW` in `src/content/intro.js`, each a list of tables
with positions and `[column, 'PK' | 'FK', referencedTable]` rows.

---

## 7. Progress and completion

Everything is in one object, `P`, saved to `localStorage['dataConcepts.v1']`:

| Field | Holds |
|---|---|
| `exDone` | Solved exercises, keyed `<lessonId>.<index>` |
| `drafts` | The formula in the grid for each lesson tab, typed Excel answers (`…:x`) and Excel practice |
| `checks`, `checkDone` | Check yourself answers per lesson, and whether all are right |
| `mode`, `modeSet` | Practise here or In Excel (the learner's choice; Excel-only exercises remember it per exercise) |
| `quizzes` | Per quiz: `{ attempts, best, last }` |
| `introDone`, `dataSheets`, `capstone`, `markedDone` | The Foundations pages, the capstone checklist, and lessons ticked by hand |
| `tasks` | Ticked checklist items, keyed `<lessonId>.<index>` |

`lessonProg(id)` counts exercises solved plus one for an all-correct check. `FOUND` holds the rules for the three
special pages. If the shape of `P` ever changes incompatibly, change the key to `dataConcepts.v2`.

---

## 8. The data (`scripts/generate_data.py`)

A seeded generator (numpy, `default_rng(2026)`) builds the whole company, so the patterns the lessons rely on are real:

- **Members** join with seasonality (a January spike), and the 2025 joiners depend on each site's **marketing spend**,
  which gives the correlation in the statistics unit (r ≈ 0.81).
- **Visit lengths** are log-normal, so positively skewed. Nine forgotten check-outs (7 to 12 hours) are the outliers.
- **Off-peak plans** only check in before 16:00 on weekdays, and aqua classes only run at sites with a pool.
- **MembersRaw** has planted problems: case, spaces, three date formats, a 1900 date of birth, a fee in pence, plan and
  site variants, blanks, and repeated rows.
- **The capstone export** (2026) is the same company nine months later, with different problems: a double export,
  check-out times before check-in, forgotten check-outs, orphaned member IDs, phone numbers without their leading zero,
  a site exporting UK-style text dates, Booleans written four ways, and JSON and XML files. The generator counts what
  it planted into `data/capstone_facts.json`, which the capstone page quotes. It also includes `sites.csv` and
  `bookings.csv`, so the capstone can be done without Power Query.
- **The Power Query practice data** is the 2025 lesson data in raw form: `members.json` with messy name case, spaces and
  lower-case plan codes, one visits CSV per site (Walsall's with UK-style text dates, Wolverhampton's with 150 rows
  exported twice), `plans.csv` and `sites.xml`. Cleaned correctly, every answer equals a formula on the lesson
  workbook, which is how the tasks are checked.

The generated files are committed, so the build needs only Node.

---

## 9. Build and checks

`npm run build` runs `scripts/check.mjs`, then `scripts/build.mjs`.

**`check.mjs`** loads the engine and content into a Node `vm` context (`scripts/content.mjs`), exactly as the page does,
and fails the build if:

- a unit lists a lesson that doesn't exist, or a lesson twice
- a lesson has no title, summary or learn text, or no exercises or questions
- every column of the main sheets has exactly one type in `DATA_TYPES` (for the Data types lesson)
- an exercise has no prompt or hint, uses an unknown sheet, or its model formula gives an error (on any filled row it
  checks), or its `askRow` is invalid
- a formula example (`F()`) in a lesson gives an error
- a Check yourself question has no right answer, or two identical options
- a quiz has a duplicate id, no questions, a bad time limit, or a multiple-choice question without four different options

It prints how the right answers are spread across a, b, c and d for each quiz.

**`build.mjs`** inlines the CSS, the data, the engine, the content and the app into `dist/index.html`, inlines the
logos as data URIs, and copies `public/` (the downloads) into `dist/`. It refuses to build if a script would contain
`</script`.

**`.github/workflows/deploy.yml`** runs the build on every push to `main` and publishes `dist/` to GitHub Pages.

---

## 10. Design decisions and trade-offs

| Decision | Benefit | Trade-off |
|---|---|---|
| Own formula engine | Instant, offline, any device, no licence questions | Not every Excel function. Dynamic-array functions (FILTER, UNIQUE, SORT) explain that they're for Excel |
| Value-based marking | Any correct formula is accepted | A hard-coded number would also pass in the grid. It's practice, not assessment |
| Answers in the page | Works offline, no server | A determined learner could find them. Fine for practice, not for assessment |
| Generated data | Exactly the teaching patterns needed, no privacy issues | It's synthetic. The README says so, and so does the footer |
| One file | Hosts anywhere | Needs a build step. Content edits go through GitHub, where the checks run automatically |
| localStorage | Private, zero set-up | Progress is per browser and device. Trainers can't see learners' scores |
