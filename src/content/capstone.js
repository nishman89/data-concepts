/* ---------- Capstone project ----------
   Learners clean and analyse the raw 2026 export (public/downloads/sparta-leisure-capstone.zip) in Excel, with Power BI
   optional. The "check your numbers" figures come from data/capstone_facts.json, written by scripts/generate_data.py. */
const CAP_Q = (n, task, check, revise) => `<li class="cap-q" data-testid="capstone-q-${n.replace('.','-')}"><div class="cap-n">${n}</div><div><p>${task}</p>${check?`<div class="cap-check"><b>Check your numbers</b>${check}</div>`:''}${revise?`<p class="cap-revise">Revise: ${revise}</p>`:''}</div></li>`;
const REV = (id) => `<a href="#/lesson/${id}">${LESSONS[id].title}</a>`;
const N = n => Number(n).toLocaleString('en-GB');
const CF = CAP_FACTS;

const CAPSTONE_HTML = `
<p class="lede">Sparta Leisure's board meets next month. The council has offered to part-fund an investment, and the board wants answers backed by data. You've been given a raw export from the company's systems for January to September 2026. Clean it, analyse it, and present your recommendations.</p>

<section class="cap-before" data-testid="capstone-brief">
  <h2>The board's three questions</h2>
  <ol class="cap-brief">
    <li><strong>Where should we invest?</strong> The money could pay for a new studio at one site, or a refurbished pool. Which sites are busiest for their size, and when do they get crowded?</li>
    <li><strong>Should we keep the Student plan?</strong> It's the cheapest plan. Do students stay, and do they use the gym enough to be worth the discount?</li>
    <li><strong>Did the New Year campaign work?</strong> The marketing manager says: <em>"January sign-ups were more than four times December's. Our New Year campaign worked, so let's double the budget next January."</em></li>
  </ol>
</section>

<h2 class="dot">Before you start</h2>
<p><a class="btn primary" href="downloads/sparta-leisure-capstone.zip" download data-testid="capstone-download">Download the capstone data (ZIP)</a></p>
<p>Unzip it. Inside are:</p>
<table class="mini"><tr><th>File</th><th>Holds</th></tr>
<tr><td><code>members.csv</code></td><td>every member who joined up to 30 September 2026</td></tr>
<tr><td><code>visits.csv</code></td><td>every gym check-in and check-out in 2026</td></tr>
<tr><td><code>bookings.json</code></td><td>every class booking, from the booking system. <code>bookings.csv</code> is the same data, if your Excel can't import JSON.</td></tr>
<tr><td><code>sites.xml</code></td><td>the six leisure centres</td></tr>
<tr><td><code>plans.csv</code>, <code>classes.csv</code></td><td>the 2026 plans and the class menu</td></tr>
<tr><td><code>marketing.csv</code></td><td>marketing spend by site and month, January 2025 to September 2026</td></tr></table>
<p>The export has <strong>not</strong> been cleaned. Expect the kinds of problems from the <a href="#/lesson/quality">Data quality</a> unit, and some you haven't seen. The 2025 workbook from the lessons is useful too: it has a full year of data to compare against.</p>
${TIP('Use Excel on a Windows or Mac computer: <strong>Data &gt; Get Data</strong> (Power Query) imports CSV, JSON and XML, and records your cleaning steps so you can repeat them. Excel for the web can open the CSV files but can\'t import JSON or XML. Power BI Desktop is optional (see the end of this brief).')}

<h2 class="dot">What you'll hand in</h2>
<table class="mini"><tr><th>File</th><th>Contains</th></tr>
<tr><td><code>sparta-leisure-analysis.xlsx</code></td><td>Your cleaned data, one sheet per table, with a <strong>Cleaning log</strong> sheet, your analysis, and a <strong>Dashboard</strong> sheet</td></tr>
<tr><td><code>star-schema.png</code> (or PDF)</td><td>A diagram of your data model</td></tr>
<tr><td><code>recommendations.pptx</code> (or a 2-page report)</td><td>Your answers to the board's three questions: no more than 8 slides, for a 10-minute presentation</td></tr>
<tr><td><code>sparta-leisure.pbix</code> (optional)</td><td>A Power BI report on the same data</td></tr></table>

<h2 class="dot">Stage 1: Import and profile</h2>
<p class="cap-when">Uses <strong>Unit 2</strong> and <strong>Unit 3</strong></p>
<ol class="cap-list">
${CAP_Q('1.1', 'Import each file into its own sheet. Use Power Query (<strong>Data &gt; Get Data</strong>) for the JSON and XML files. Name the sheets after the files.', `members.csv has <strong>${N(CF.membersRowsRaw)}</strong> rows of data, visits.csv has <strong>${N(CF.visitsRowsRaw)}</strong>, and bookings.json has <strong>${N(CF.bookingsRows)}</strong> bookings.`, REV('formats'))}
${CAP_Q('1.2', 'For each table, write a short <strong>data dictionary</strong>: each column\'s name, its data type (numeric, categorical, Boolean, datetime…), and what it means. Note any column whose type Excel guessed wrongly.', null, REV('types'))}
${CAP_Q('1.3', 'Profile the data before changing anything. Count the blanks in each column, list the distinct values of the categorical columns (a pivot table makes this quick), and find each numeric column\'s minimum and maximum.', `<code>Phone</code> has <strong>${N(CF.membersPhoneMissingZero)}</strong> numbers that have lost their leading zero, and <code>Gender</code> is blank in <strong>${N(CF.membersBlankGender)}</strong> rows.`, REV('functions'))}
</ol>

<h2 class="dot">Stage 2: Clean, and log what you did</h2>
<p class="cap-when">Uses <strong>Unit 6</strong></p>
<p>Keep a <strong>Cleaning log</strong> sheet with these columns: <em>Table</em>, <em>Problem</em>, <em>Quality dimension</em>, <em>Rows affected</em>, <em>What you did</em>, and <em>Why</em>. Anyone reading it should be able to repeat your cleaning and understand every decision.</p>
<ol class="cap-list">
${CAP_Q('2.1', 'Remove <strong>exact duplicate</strong> rows from members and visits. One of the exports was run twice.', `members.csv has <strong>${CF.membersDuplicates}</strong> duplicate rows, leaving <strong>${N(CF.membersUnique)}</strong> members. visits.csv has <strong>${CF.visitsDuplicates}</strong>, leaving <strong>${N(CF.visitsRowsRaw - CF.visitsDuplicates)}</strong> visits.`, REV('cleaning'))}
${CAP_Q('2.2', '<strong>Standardise</strong> the categories: genders, plan codes and site IDs each appear in several forms. Map them all to the codes used in plans.csv and sites.xml.', `<strong>${CF.membersSiteAsTown}</strong> members have a town name where their home site ID should be.`, REV('cleaning'))}
${CAP_Q('2.3', 'Fix the <strong>types</strong>: dates of birth are written three different ways, Walsall\'s visits use a different date format from the other sites, and the booking system records <code>attended</code> in four different ways. Turn them into real dates and real TRUE/FALSE values.', `<code>attended</code> is empty for <strong>${CF.bookingsAttendedBlank}</strong> bookings. Decide what to do with them, and log it.`, REV('types'))}
${CAP_Q('2.4', 'Work out each visit\'s <strong>length in minutes</strong> from CheckIn and CheckOut. Then decide what to do with impossible and suspicious visits, and log your reasons.', `You should find <strong>${CF.visitsNegative}</strong> visits that end before they start (a clock fault), and <strong>${CF.visitsForgotCheckout}</strong> that the system closed at 23:59 because the member forgot to check out.`, REV('distributions'))}
${CAP_Q('2.5', 'Check <strong>validity</strong> and <strong>referential integrity</strong>. Find impossible dates of birth, and visits by member IDs that don\'t exist in members.csv. Pad phone numbers back to 11 digits.', `There are <strong>${CF.membersImpossibleDob}</strong> impossible dates of birth and <strong>${CF.visitsOrphan}</strong> visits by unknown members.`, REV('missing'))}
</ol>

<h2 class="dot">Stage 3: Link and model</h2>
<p class="cap-when">Uses <strong>Unit 4</strong> and <strong>Unit 9</strong></p>
<ol class="cap-list">
${CAP_Q('3.1', 'Use lookups to bring the details you need into your visits and bookings sheets: each member\'s plan and age band, and each site\'s name and capacity.', null, REV('lookups'))}
${CAP_Q('3.2', 'Draw a <strong>star schema</strong> for Sparta Leisure, with visits as the fact table. Show the dimension tables, their keys, and the grain of the fact table. Add a sentence on how a snowflake schema would differ.', null, REV('warehouses'))}
</ol>

<h2 class="dot">Stage 4: Analyse</h2>
<p class="cap-when">Uses <strong>Unit 4</strong> and <strong>Unit 5</strong></p>
<ol class="cap-list">
${CAP_Q('4.1', '<strong>Describe the visits.</strong> Find the mean, median, mode and standard deviation of visit length, before and after your outlier decisions, and draw a histogram. Is the distribution skewed? Which average would you quote to the board, and why?', null, REV('distributions'))}
${CAP_Q('4.2', '<strong>Where to invest.</strong> Compare the sites: visits per month, visits per place of capacity, and how busy each one is by hour and day of the week. Which sites are crowded at peak time (Monday to Friday, 16:00 to 20:59)? Which of them has a pool?', null, REV('conditional'))}
${CAP_Q('4.3', '<strong>The Student plan.</strong> For each plan, work out the number of members who joined, the share who have cancelled, the average visits per active member per month, and the monthly income. How do students compare?', null, REV('pivots'))}
${CAP_Q('4.4', '<strong>The campaign.</strong> Count sign-ups by month for 2025 and 2026 (use JoinDate). Compare January 2026 with December 2025, then compare January 2026 with January 2025. Bring in the marketing spend, and work out the correlation between spend and sign-ups.', `January 2026 had <strong>${CF.joinsJan2026}</strong> sign-ups against <strong>${CF.joinsDec2025}</strong> in December 2025. A year earlier, January 2025 had <strong>${CF.joinsJan2025}</strong> against December 2024's <strong>${CF.joinsDec2024}</strong>.`, REV('correlation'))}
</ol>

<h2 class="dot">Stage 5: Think critically</h2>
<p class="cap-when">Uses <strong>Unit 7</strong></p>
<ol class="cap-list">
${CAP_Q('5.1', 'Run the marketing manager\'s claim through the <strong>SIMPLE</strong> checklist. What\'s the obvious alternative explanation? What does comparing January with January show? What data is missing that would settle the question? Write one paragraph you could say to the marketing manager.', null, REV('simple'))}
${CAP_Q('5.2', 'For each of your own findings in Stage 4, note one alternative explanation and how confident you are. Separate what the data shows (facts) from what you think it means (suppositions).', null, REV('alternatives'))}
</ol>

<h2 class="dot">Stage 6: Visualise and present</h2>
<p class="cap-when">Uses <strong>Unit 4</strong> and <strong>Unit 8</strong></p>
<ol class="cap-list">
${CAP_Q('6.1', 'Build a one-screen <strong>Dashboard</strong> sheet from pivot charts, with slicers for site and plan connected to every chart. Choose each chart for the relationship it shows, and follow the design advice: no 3D, no truncated bar axes, colour used consistently and accessibly.', null, REV('dashboards'))}
${CAP_Q('6.2', 'Write your <strong>recommendations</strong>: one slide (or section) per board question, each with a clear answer, the evidence for it, and how confident you are. Finish with what data Sparta Leisure should start collecting.', null, REV('good-viz'))}
</ol>

<div class="cap-optional" data-testid="capstone-power-bi"><b>Optional: Power BI.</b> Load your cleaned tables into Power BI Desktop, create the relationships to match your star schema in the <strong>Model</strong> view, and build a report page that answers one of the board's questions. Power BI makes the star schema real: the visits table sits in the middle, and filters flow out from the dimensions.</div>

<h2 class="dot">What good looks like</h2>
<table class="mini"><tr><th>Area</th><th>A strong project…</th></tr>
<tr><td>Cleaning</td><td>finds every planted problem, fixes each one in a way that can be repeated, and logs it with a reason</td></tr>
<tr><td>Analysis</td><td>uses the right measures (median where data is skewed, rates rather than totals when sites differ in size), with formulas or pivot tables rather than typed numbers</td></tr>
<tr><td>Critical thinking</td><td>tests the campaign claim fairly, considers alternatives for its own findings, and is honest about uncertainty</td></tr>
<tr><td>Visualisation</td><td>has charts that are trustworthy, accessible and elegant, with each one chosen for its relationship</td></tr>
<tr><td>Communication</td><td>answers the board's questions directly, in plain language, with clear recommendations</td></tr></table>`;

/* [id, label, optional] */
const CAPSTONE_TASKS = [
  ['1', 'Stage 1: import every file, write the data dictionary and profile the data'],
  ['2', 'Stage 2: clean the data and complete the cleaning log'],
  ['3', 'Stage 3: link the tables and draw the star schema'],
  ['4', 'Stage 4: analyse the visits, the sites, the plans and the campaign'],
  ['5', 'Stage 5: test the campaign claim with SIMPLE'],
  ['6', 'Stage 6: build the dashboard and write the recommendations'],
  ['pbi', 'Build a Power BI report', true],
];
