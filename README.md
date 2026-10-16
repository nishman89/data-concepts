# Data Concepts | Sparta Global

A learning path for the Data Concepts module. It covers data types, Excel, basic statistics, data quality, critical
thinking, visualisation and data storage, with a capstone project at the end and an optional Power Query unit after it. The Excel exercises can be done in the
browser, where formulas run on the real data and are marked straight away, or in Excel with a downloadable workbook.

Everything uses **Sparta Leisure**, a made-up chain of six leisure centres with nine linked sheets: sites, plans,
price history, members, classes, bookings, visits, monthly site figures, and a deliberately messy extract for the cleaning
lessons.

It's a sister site to the SQL Practice Lab and the ISTQB study site, and uses the same look.

## What's in it

| | |
|---|---|
| **Learning path** | 30 lessons in 10 units, built from the course slides and Excel notes, plus an optional Power Query unit (4 lessons) |
| **Excel exercises** | 76, each marked against the real data. Learners choose **Practise here** (a formula grid in the browser) or **In Excel** (download the workbook and type the answer) |
| **Check yourself** | Multiple-choice questions on the concept lessons, with an explanation for every answer |
| **Interactive pieces** | A histogram with a bin slider, a correlation scatter plot, Anscombe's Quartet, a truncated-axis demo, a colour-blindness simulator, the file formats side by side, and star and snowflake schemas |
| **Quizzes** | An end-of-unit quiz for units 2 to 9 and a final quiz: 70 questions, 65% pass mark, optional timer |
| **Capstone project** | A raw, messy 2026 export (CSV, JSON and XML) to clean, analyse and present to the board, in Excel with Power BI optional |
| **Power Query (optional)** | Install Excel, then import JSON, CSV and XML, clean with recorded steps, combine a folder of files and merge tables, using a raw version of the 2025 data. Answers are checked on the site. Optional lessons don't count towards completing the path |
| **Excel practice** | A free formula grid on all nine sheets |

Progress is saved in each learner's browser (`localStorage`). Nothing is sent anywhere.

## Running it locally

You need [Node.js](https://nodejs.org) 20 or later. There are no other dependencies.

```bash
npm run build      # checks every lesson and question, then builds dist/
npm run check      # just the checks
npm run dev        # rebuilds whenever a file in src/ or data/ changes
```

Then open `dist/index.html` in a browser. The download buttons need the files in `dist/downloads/`, so keep the
`dist/` folder together.

## Publishing

Every push to `main` runs `.github/workflows/deploy.yml`, which runs the checks, builds the site and publishes `dist/`
to **GitHub Pages**. If a check fails, nothing is published, and the failed run in the **Actions** tab says what's
wrong. To set it up the first time, go to **Settings > Pages** and set **Source** to **GitHub Actions**.

The site is plain static files, so `dist/` can also be hosted anywhere else, or zipped and shared.

---

## Updating the content

All the words learners see are in **`src/content/`**, as plain text inside JavaScript files. You don't need to
understand the app code to change them.

| To change… | Edit |
|---|---|
| A lesson's text, exercises or Check yourself questions | `src/content/lessons.js` |
| The order of lessons and units | `UNITS` at the top of `src/content/lessons.js` |
| A quiz | `src/content/quizzes.js` |
| The capstone brief or its checklist | `src/content/capstone.js` |
| How this site works, and Meet Sparta Leisure | `src/content/intro.js` |

### The easy way: edit on GitHub

You don't need to install anything to make a change.

1. Open the file on GitHub, for example `src/content/lessons.js`, and select the **pencil icon** (Edit this file).
2. Make your change. Use **Ctrl+F** to find the lesson you want.
3. Select **Commit changes**. Commit straight to `main` for a small fix, or create a branch and pull request if you'd
   like someone to review it.
4. Open the **Actions** tab. After a minute or two the run goes green ✓ and the live site is updated.
   If it goes red ✗, open the run and expand **npm run build**. The message names the lesson and the exercise and says
   what's wrong. Fix it the same way, and the next commit tries again. **The live site isn't touched until the checks pass.**

### Editing safely

The content files are JavaScript, so a few characters matter:

- Lesson text sits between backticks: `` learn:`…` ``. Inside it, write ordinary HTML: `<p>`, `<strong>`,
  `<code>`, `<table class="mini">` and so on. If you need a backtick in the text, write `` \` ``.
- Prompts, hints and questions sit between single quotes: `'…'`. An apostrophe inside them needs a backslash:
  `'Excel\'s menus'`.
- Every item in a list ends with a comma.

If you break one of these rules, the check stops with a "SyntaxError" and a line number.

### Adding an Excel exercise

Exercises live in a lesson's `ex` list. Each one is an object like this:

```js
{ p:'How many members are **Active**?', sheet:'Members', sol:'=COUNTIF(J:J,"Active")', hint:'COUNTIF(the Status column, "Active").' },
```

| Field | Meaning |
|---|---|
| `p` | The task. `**bold**` and `` `code` `` work here |
| `sheet` | The sheet the formula goes on. The answer cell is row 2 of the first empty column (the pink column) |
| `sol` | A formula that gives the right answer. **You never type the answer itself**: the site works it out from this formula, so it can't be wrong |
| `hint` | Shown when the learner asks for a hint |
| `fill: true` | The formula fills down a new column, and every row is checked (the first 300). Write `sol` for row 2 |
| `col` | The heading of that new column, such as `'Age'` |
| `askRow` | For `fill` exercises in **In Excel** mode: which row's result the learner types. The default is row 2 |
| `vtype: 'date'` | The answer is a date, so it's shown and typed as a date |
| `excel: true` | The task is done with Excel's menus (sorting, filtering, pivot tables), so it opens in **In Excel** mode |
| `steps` | A list of numbered instructions shown above the answer box (used by the Power Query tasks) |
| `ask` | The question above the answer box, instead of the default "Type the answer you get in Excel" |
| `how` | Shown by **Show the answer** instead of the model formula. Use it when the task isn't done with a formula |

Learners can use any formula that gives the same values, not just yours. Numbers count as the same within 0.005, and
text ignores capitals.

The check runs your `sol` against the data and stops the build if it gives an error. Run `npm run check` (or watch the
Actions tab) after adding exercises.

### Adding a Check yourself question

In a lesson's `check` list:

```js
{ q:'Which column is **continuous**?', a:['<code>PlanCode</code>','<code>AvgTempC</code>','<code>Capacity</code>'], c:1, why:'Temperature is measured on a scale…' },
```

`a` is the list of options, `c` is the position of the right one, **counting from 0**, and `why` is shown when it's
answered correctly. Write the right answer anywhere: the site rotates the options so right answers are spread across
a, b, c and d.

### Adding a quiz question

In `src/content/quizzes.js`, use `Q(…)` for multiple choice (always four options) and `V(…)` for a value from the data:

```js
Q(4, 'Question text', ['option 1','option 2','option 3','option 4'], 1, 'Why option 2 is right.'),
V(4, 'How many visits were made at <strong>S06</strong>?', 'Visits', '=COUNTIF(C:C,"S06")', 'COUNTIF of SiteID.'),
```

The first number is the unit the question belongs to, which is used for the score-by-unit table.

### Adding a lesson

1. Add an entry to `LESSONS` in `src/content/lessons.js`, with an id, `title`, `summary`, `learn`, and at least one of
   `ex` (with `sheet` set to the sheet the grid should open on), `check` or `tasks`.
2. Add its id to a unit's `lessons` list in `UNITS`, in the order it should appear.

Other lesson options:

| Field | Meaning |
|---|---|
| `tasks` | A checklist shown at the end of the Learn tab, such as the Power Query set-up steps. The lesson is complete when everything is ticked |
| `noGrid: true` | For lessons done entirely in Excel (the Power Query unit): no practice grid, and every exercise is answered by typing a value. They're called "tasks" on the page |

A unit with `optional: true` is shown with an **Optional** badge. Its lessons don't count towards finishing the path,
and "Continue" skips them.

**Power Query answers** use the same trick as every other exercise: the practice files are a raw version of the 2025
workbook, so once a learner cleans them correctly, the answer matches a formula on the workbook. For example, the
"remove duplicates" task's `sol` is `=COUNTA(Visits!A:A)-1`.

In the `learn` text, `F('=AVERAGE(E:E)', 'Visits')` adds a formula example with a **Try this** button, and
`TIP('…')` adds a tip box. `W('histogram')` and the other names in `WIDGETS` (in `src/app.js`) add an interactive chart.

### Changing the data

The data is generated by `scripts/generate_data.py` (Python 3 with `numpy` and `openpyxl`). It writes
`data/workbook.json` (used by the site), `public/downloads/sparta-leisure.xlsx`,
`public/downloads/sparta-leisure-capstone.zip`, `public/downloads/sparta-leisure-powerquery.zip` and `data/capstone_facts.json` (the "check your numbers" figures on the
capstone page). The random seed is fixed, so running it again gives exactly the same files.

```bash
pip install numpy openpyxl
python3 scripts/generate_data.py
npm run build
```

Exercises don't store answers, so they stay correct when the data changes. Any lesson text that quotes a specific
number would need checking by hand.

---

## Project layout

```
src/
  index.html          page shell: header, views, footer
  styles.css          the shared Sparta look, plus the grid and chart styles
  engine.js           the Excel formula engine (no page code, so the checks use it too)
  app.js              routing, the learning path, lessons, the practice grid, charts and diagrams
  quiz.js             the quiz engine
  content/            everything learners read (see above)
  assets/             logos and favicon
data/                 the generated workbook data and capstone facts
public/downloads/     the Excel workbook, the capstone ZIP and the Power Query practice ZIP
scripts/              generate_data.py, check.mjs, build.mjs
```

For how it all fits together, see [ARCHITECTURE.md](ARCHITECTURE.md).
