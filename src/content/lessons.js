/* ---------- Learning path content ----------
   UNITS sets the order of the path. Each lesson in LESSONS has:
     title, summary          shown on the path and at the top of the lesson
     learn                   HTML for the Learn tab (F() makes a runnable formula, W() an interactive piece)
     sheet                   the sheet the practice grid opens on (Excel lessons)
     ex                      Excel exercises: { p, sheet, sol, hint, fill?, col?, askRow?, vtype?, excel? }
                               sol is a model formula, written for the answer cell: row 2 of the first empty column.
                               fill: the formula is filled down a new column called col, and every row is checked.
                               excel: the task is done with Excel's menus (sort, filter, pivot tables), so the lesson
                               opens in "In Excel" mode. The model formula still works out the expected answer.
     check                   multiple-choice questions instead of exercises: { q, a:[options], c, why }
     kind                    'intro' | 'data' | 'capstone' for lessons with their own page */

const UNITS = [
  { title:'Foundations', blurb:'How this site works, and the leisure centre data you\'ll use throughout.', lessons:['intro','data'] },
  { title:'Data types and file formats', blurb:'What kinds of data there are, and the files they travel in.', lessons:['types','formats'] },
  { title:'Excel essentials', blurb:'Formulas, references and the everyday functions, on real data.', lessons:['formulas','functions','text-dates','logic','errors'] },
  { title:'Excel for analysis', blurb:'Join sheets together with lookups, summarise with conditions and pivot tables, and build a dashboard.', lessons:['lookups','conditional','sort-filter','pivots','dashboards'] },
  { title:'Basic statistics', blurb:'Averages, spread, distributions, outliers and correlation, worked out in Excel.', lessons:['averages','spread','distributions','correlation'] },
  { title:'Data quality and cleaning', blurb:'Spot what\'s wrong with data, fix it, and decide what to do about gaps.', lessons:['quality','cleaning','missing'] },
  { title:'Critical thinking for data', blurb:'Question the data, and the conclusions people draw from it.', lessons:['simple','alternatives'] },
  { title:'Data visualisation', blurb:'Choose the right chart, design it well, and use colour with care.', lessons:['good-viz','charts','colour'] },
  { title:'Data storage', blurb:'Pipelines, databases, warehouses, lakes and lakehouses.', lessons:['pipelines','warehouses','lakes'] },
  { title:'Capstone project', blurb:'Analyse a messy real-world export and present your recommendations.', lessons:['capstone'] },
  { title:'Power Query', optional:true, blurb:'An optional extra: import, clean and combine data with steps Excel records and repeats for you. You\'ll need Excel installed on your computer.', lessons:['pq-setup','pq-import','pq-clean','pq-combine'] },
];

const LESSONS = {
/* ===================== Unit 1 ===================== */
intro:{ kind:'intro', title:'How this site works', summary:'The learning path, practising Excel here or in Excel itself, and saving your progress.' },
data:{ kind:'data', title:'Meet Sparta Leisure', summary:'Six leisure centres, nine sheets of data, and how they link together.' },

/* ===================== Unit 2 ===================== */
types:{ title:'Data types', summary:'Numeric and non-numeric, continuous, discrete and categorical, Boolean, dates and times, and key-value pairs.',
learn:`
<p>Before you analyse anything, work out what <em>kind</em> of data each column holds. The type decides which calculations make sense, which charts work, and how the data should be stored.</p>
<h2 class="dot">One record, many types</h2>
<p>Here's a real row from Sparta Leisure's data, with every value labelled. Switch sheets, or show another row, and notice how one record mixes several types. Each type is explained below.</p>
${W('typerecord')}
<h2 class="dot">Numeric and non-numeric</h2>
<p><strong>Numeric</strong> data is quantitative: it's only numbers, so you can do arithmetic with it. Whole numbers are <strong>integers</strong> (<code>Capacity</code> is 220) and numbers with a decimal point are <strong>decimals</strong>, or floats (<code>MonthlyFee</code> is 32.99).</p>
<p><strong>Non-numeric</strong> data is qualitative: text, made of any characters. In programming it's called a <strong>string</strong>, such as <code>SiteName</code> = Telford Central.</p>
${TIP('A number isn\'t always numeric data. A phone number or a member ID is made of digits, but adding two phone numbers together means nothing. Store those as text.')}
<h2 class="dot">Continuous, discrete and categorical</h2>
<table class="mini"><tr><th>Type</th><th>What it means</th><th>At Sparta Leisure</th></tr>
<tr><td><strong>Continuous</strong></td><td>Numeric data measured on a scale. It can be split into finer and finer amounts, as precise as your measuring allows.</td><td><code>AvgTempC</code>, a visit's exact length, a member's height</td></tr>
<tr><td><strong>Discrete</strong></td><td>Counted in whole units that can't be split further and still mean something.</td><td><code>Capacity</code>, <code>MaxPlaces</code>, the number of visits</td></tr>
<tr><td><strong>Categorical</strong></td><td>Grouped into categories. Usually text, sometimes coded as numbers.</td><td><code>PlanCode</code>, <code>Gender</code>, <code>Category</code>, <code>Intensity</code></td></tr></table>
<p>Some categories have a natural order (<code>Intensity</code>: Low, Medium, High), which is called <strong>ordinal</strong> data. Others don't (<code>Town</code>), which is called <strong>nominal</strong>.</p>
<h2 class="dot">Boolean</h2>
<p>A Boolean has exactly two possible values: <strong>TRUE</strong> (1) or <strong>FALSE</strong> (0). <code>HasPool</code> and <code>Attended</code> are Booleans. Excel and most programming languages combine them with logical operators:</p>
<table class="mini"><tr><th>Operation</th><th>Excel</th><th>Other common forms</th></tr>
<tr><td>Conjunction (both true)</td><td><code>AND(…)</code></td><td><code>&amp;</code> <code>&amp;&amp;</code> <code>*</code></td></tr>
<tr><td>Disjunction (either true)</td><td><code>OR(…)</code></td><td><code>|</code> <code>||</code> <code>+</code></td></tr>
<tr><td>Equivalence</td><td><code>=</code></td><td><code>==</code> <code>EQV</code></td></tr>
<tr><td>Exclusive or (one or the other)</td><td><code>XOR(…)</code> or <code>&lt;&gt;</code></td><td><code>^</code> <code>!=</code></td></tr>
<tr><td>Negation</td><td><code>NOT(…)</code></td><td><code>!</code> <code>~</code></td></tr></table>
<h2 class="dot">Dates and times</h2>
<p>Date and time data comes in three main types: a <strong>date</strong> (<code>JoinDate</code>), a <strong>time</strong> (<code>StartTime</code>) and a <strong>datetime</strong> (<code>CheckIn</code>). They work differently in every system, so check the documentation. The same day can be written many ways: 03/04/2006, 2006-04-03, 3 Apr 2006, April 3rd 2006. In the UK 03/04/2006 is the 3rd of April, but in the US it's the 4th of March.</p>
<p>Excel stores a date as a <strong>serial number</strong>: the number of days since 31 December 1899. 1 January 2025 is 45658, and a time is a fraction of a day, so 18:00 is 0.75. This is why you can subtract one date from another to get the number of days between them.</p>
${F('=DATE(2025,1,1)','Members',{note:'Try it: this gives a serial number. Excel shows it as a date only because of the cell\'s format.'})}
${TIP('The safest format to share dates in is <strong>ISO 8601</strong>: <code>2025-01-31</code>. It sorts correctly as text, and no one can confuse the day and the month.')}
<h2 class="dot">Every type in Sparta Leisure's data</h2>
<p>Each card shows the columns of one type, with real values from the workbook. Select a column to see more of its values.</p>
${W('typecards')}
${TIP('Some columns could go either way. <code>DurationMins</code> is time, which is continuous, but it\'s recorded in whole minutes, so it looks discrete. <code>MonthlyFee</code> is money, which only goes down to a penny. Treat a measured amount as continuous when its unit is small compared with its values: it will still make sense to average it and draw it as a histogram.')}
<h2 class="dot">Key-value pairs</h2>
<p>A key-value pair (KVP) links two pieces of data. The <strong>key</strong> is a unique label, and the <strong>value</strong> is the data it identifies. Some keys have one value, and some can have several:</p>
<pre class="code">site_id: S01
name: Telford Central
has_pool: true
classes: [Spin, Yoga, Aqua Fit]</pre>
<p>Key-value pairs are behind dictionaries and hash tables in programming, JSON and YAML files, and NoSQL databases.</p>`,
check:[
  {q:'Which Sparta Leisure column holds <strong>continuous</strong> data?', a:['<code>PlanCode</code>','<code>AvgTempC</code>','<code>Capacity</code>','<code>HasPool</code>'], c:1, why:'Temperature is measured on a scale and can be as precise as the thermometer allows. Capacity is a count (discrete), PlanCode is a category and HasPool is a Boolean.'},
  {q:'<code>Intensity</code> holds Low, Medium or High. What kind of data is it?', a:['Continuous','Discrete numeric','Categorical, with an order (ordinal)','Boolean'], c:2, why:'It\'s categorical because each value is a group, and ordinal because the groups have a natural order.'},
  {q:'A member\'s phone number, 07700900123, should be stored as…', a:['an integer, because it\'s all digits','a decimal','text (a string)','a Boolean'], c:2, why:'You never do arithmetic on a phone number, and storing it as a number loses the leading zero. Store it as text.'},
  {q:'In Excel, what is the date 2 January 2025 if 1 January 2025 is 45658?', a:['45659','45658.5','2025-01-02 as text','It can\'t be worked out'], c:0, why:'Each day adds 1 to the serial number.'},
  {q:'<code>CheckIn</code> on the Visits sheet holds values like <code>01/01/2025 06:01</code>. What type is it?', a:['Date','Time','Datetime','Text'], c:2, why:'It holds a date and a time together, so it\'s a datetime.'},
  {q:'Which column is <strong>ordinal</strong> data?', a:['<code>Category</code> (Cardio, Strength…)','<code>Intensity</code> (Low, Medium, High)','<code>Town</code>','<code>PlanCode</code>'], c:1, why:'Low, Medium and High have a natural order. The others are groups with no order, which is nominal data.'},
  {q:'Which is a key-value pair?', a:['<code>Telford Central</code>','<code>town: Telford</code>','<code>TRUE</code>','<code>45658</code>'], c:1, why:'<code>town</code> is the key and <code>Telford</code> is its value.'},
]},
formats:{ title:'File formats', summary:'TXT, CSV, XLSX, JSON, YAML, XML and DAT: what each one is and when you\'ll meet it.',
learn:`
<p>Data moves between systems in files, and each format has its strengths. Below are the same three Sparta Leisure members in five formats. Compare how each one marks where a value starts and ends.</p>
${W('formats')}
<table class="mini"><tr><th>Format</th><th>What it is</th><th>You'll meet it when…</th></tr>
<tr><td><code>.txt</code></td><td>Plain, unformatted text. Any program can open it.</td><td>reading logs, notes and simple exports</td></tr>
<tr><td><code>.csv</code></td><td>Comma-separated values: plain text, one record per line, a comma between fields.</td><td>moving data between databases, spreadsheets and contact managers. It's the most common exchange format.</td></tr>
<tr><td><code>.xlsx</code></td><td>An Excel workbook: one or more worksheets, plus formulas, charts, styles and formatting.</td><td>sharing analysis with people who work in Excel</td></tr>
<tr><td><code>.json</code></td><td>JavaScript Object Notation: key-value pairs and ordered lists, as lightweight, human-readable text.</td><td>web apps talk to servers, and when you download data from an API</td></tr>
<tr><td><code>.yml</code> / <code>.yaml</code></td><td>"YAML Ain't Markup Language": key-value pairs written with colons and indentation.</td><td>configuring software and data pipelines</td></tr>
<tr><td><code>.xml</code></td><td>Extensible Markup Language: custom tags describe the structure, readable by people and machines.</td><td>older systems, documents, and data feeds between organisations</td></tr>
<tr><td><code>.dat</code></td><td>A generic data file, in binary or text, written by one program for its own use.</td><td>an application's own files. Usually only that application can read it.</td></tr></table>
${TIP('A CSV has no types: everything in it is text until a program decides otherwise. When Excel opens a CSV it guesses, and it sometimes guesses wrong. It can drop the leading zero from a phone number, or turn a product code like <code>MAR1</code> into a date. Use <strong>Data &gt; From Text/CSV</strong> to choose each column\'s type yourself.')}
<h2 class="dot">Try it</h2>
<p>Open a plain-text editor, such as Notepad, Notepad++ or VS Code. Then type up a small set of data (three of your favourite films, say) as <code>.txt</code>, <code>.csv</code>, <code>.json</code>, <code>.yml</code> and <code>.xml</code>. Open the CSV in Excel to see how it's read.</p>`,
check:[
  {q:'Which format stores <strong>formulas, charts and several sheets</strong>?', a:['.csv','.xlsx','.txt','.json'], c:1, why:'A CSV holds one table of plain values. An Excel workbook (.xlsx) holds worksheets, formulas, charts and formatting.'},
  {q:'A web app asks a server for a member\'s classes. Which format is the reply most likely to be in?', a:['.dat','.xml','.json','.xlsx'], c:2, why:'JSON is the usual format for data between web apps and servers. XML does the same job in many older systems.'},
  {q:'Which of these describes YAML?', a:['Tags in angle brackets, like &lt;name&gt;','Key-value pairs written with colons and indentation','Values separated by commas','A binary file for one program'], c:1, why:'YAML uses <code>key: value</code> and indentation. XML uses tags, and CSV uses commas.'},
  {q:'You open a CSV of members in Excel and the phone numbers have lost their leading 0. Why?', a:['CSV files can\'t store zeros','Excel guessed the column was a number','The file is corrupt','JSON would have the same problem'], c:1, why:'The CSV holds text, but Excel decided the column was a number, and numbers don\'t have leading zeros. Import the column as text instead.'},
]},

/* ===================== Unit 3: Excel essentials ===================== */
formulas:{ title:'Formulas and cell references', summary:'Operators, the order Excel works things out, and relative and absolute references.', sheet:'Plans',
learn:`
<p>Every formula starts with an equals sign. After it you can use numbers, cell references, operators and functions. Press <kbd>Enter</kbd> and the cell shows the result. The formula itself stays in the formula bar.</p>
<table class="mini"><tr><th>Operator</th><th>Does</th><th>Example</th></tr>
<tr><td><code>+</code> <code>-</code></td><td>add, subtract</td><td><code>=C2+D2</code></td></tr>
<tr><td><code>*</code> <code>/</code></td><td>multiply, divide</td><td><code>=C2*12</code></td></tr>
<tr><td><code>^</code></td><td>raise to a power</td><td><code>=2^3</code> is 8</td></tr>
<tr><td><code>&amp;</code></td><td>join text</td><td><code>=A2&amp;" plan"</code></td></tr>
<tr><td><code>=</code> <code>&lt;&gt;</code> <code>&lt;</code> <code>&gt;</code> <code>&lt;=</code> <code>&gt;=</code></td><td>compare, giving TRUE or FALSE</td><td><code>=C2&gt;30</code></td></tr></table>
<p>Try one on the <strong>Plans</strong> sheet beside this lesson: the cost of a year on the first plan.</p>
${F('=C2*12','Plans',{fill:true,col:'AnnualFee',note:'Each row uses its own fee, because C2 becomes C3, C4… as the formula fills down.'})}
<h2 class="dot">The order Excel works things out</h2>
<p>Excel follows the same rules as maths (often remembered as BODMAS or BIDMAS): brackets first, then powers, then multiply and divide, then add and subtract. So <code>=C2*12+D2</code> is the annual fee plus the joining fee, but <code>=C2*(12+D2)</code> is something quite different.</p>
<h2 class="dot">Relative and absolute references</h2>
<p>When you copy or fill a formula down, Excel moves its references with it. <code>=C2*12</code> in row 2 becomes <code>=C3*12</code> in row 3. That's a <strong>relative</strong> reference, and it's usually what you want.</p>
<p>Sometimes one reference must stay put, such as a total or a rate in a single cell. Add <code>$</code> signs to make it <strong>absolute</strong>: <code>$E$2</code> never moves. Press <kbd>F4</kbd> while the cursor is in a reference to cycle through <code>E2</code>, <code>$E$2</code>, <code>E$2</code> and <code>$E2</code>. The F4 key works in the grid here too.</p>
${F('=E2/SUM($E$2:$E$7)','Sites',{fill:true,col:'Share',note:'Each site\'s share of the total capacity. Remove the $ signs and fill it down again to see what goes wrong.'})}
${TIP('Put numbers you might change, such as a VAT rate or a target, in a cell of their own and refer to it with an absolute reference. Writing <code>=C2*1.2</code> hides the 1.2 inside the formula, where no one will think to update it.')}
<h2 class="dot">Using the practice grid</h2>
<ul class="plain"><li>Type a formula in the bar above the grid and press <kbd>Enter</kbd>. It goes in the pink answer column.</li>
<li>While you type, select a cell or a column letter in the grid to add it to your formula. Switch sheets with the tabs at the bottom to refer to another sheet.</li>
<li><strong>Fill down</strong> copies your formula to every row, the way double-clicking the fill handle does in Excel.</li></ul>`,
ex:[
  {p:'On the **Plans** sheet, work out the **annual fee** of each plan: 12 times the monthly fee.', sheet:'Plans', fill:true, col:'AnnualFee', sol:'=C2*12', hint:'The monthly fee is in column C. Multiply it by 12, then fill down.'},
  {p:'Work out the **cost of the first year** for each plan: 12 months of fees plus the joining fee.', sheet:'Plans', fill:true, col:'FirstYear', sol:'=C2*12+D2', hint:'Multiplying happens before adding, so you don\'t need brackets: monthly fee × 12, then + the joining fee in column D.'},
  {p:'If a member uses all of their classes, how much does each class cost them? Divide the monthly fee by `ClassesPerMonth`.', sheet:'Plans', fill:true, col:'CostPerClass', sol:'=C2/F2', hint:'ClassesPerMonth is column F.'},
  {p:'On the **Sites** sheet, work out each site\'s **share of the total capacity** (as a decimal). The total must stay the same as you fill down.', sheet:'Sites', fill:true, col:'Share', sol:'=E2/SUM($E$2:$E$7)', hint:'Divide E2 by the SUM of E2:E7. Make the SUM\'s range absolute with $ signs (press F4) so it doesn\'t slide down.', askRow:4},
  {p:'On the **PriceHistory** sheet, work out how much each plan\'s fee **rose from 2024 to 2025**, as a decimal fraction of the 2024 fee.', sheet:'PriceHistory', fill:true, col:'Rise', sol:'=(E2-D2)/D2', hint:'(new − old) ÷ old. The brackets matter: without them Excel divides first.'},
]},
functions:{ title:'Everyday functions', summary:'SUM, AVERAGE, MIN, MAX, COUNT, COUNTA and ROUND.', sheet:'Visits',
learn:`
<p>A function is a ready-made calculation. You give it <strong>arguments</strong> in brackets, separated by commas, and it returns a result. Excel has over 400 of them, listed under <strong>Formulas &gt; Function Library</strong>.</p>
<p><code>=SUM(E2:E50)</code> adds up E2 to E50. <code>=SUM(E:E)</code> adds up the whole of column E, however many rows it has, and ignores the text heading in E1. Whole columns save you from updating ranges when data is added.</p>
${F('=SUM(E:E)','Visits',{note:'Every minute spent in a Sparta Leisure gym in 2025.'})}
<table class="mini"><tr><th>Function</th><th>Returns</th></tr>
<tr><td><code>SUM</code></td><td>the total</td></tr>
<tr><td><code>AVERAGE</code></td><td>the mean: the total divided by how many numbers there are</td></tr>
<tr><td><code>MIN</code>, <code>MAX</code></td><td>the smallest and largest number</td></tr>
<tr><td><code>COUNT</code></td><td>how many cells hold <strong>numbers</strong>. Dates count, because dates are numbers.</td></tr>
<tr><td><code>COUNTA</code></td><td>how many cells are <strong>not empty</strong>, whatever they hold</td></tr>
<tr><td><code>COUNTBLANK</code></td><td>how many cells are empty</td></tr>
<tr><td><code>ROUND(number, places)</code></td><td>a number rounded to that many decimal places</td></tr></table>
${F('=ROUND(AVERAGE(E:E),1)','Visits',{note:'One function inside another: the average visit length, rounded to 1 decimal place.'})}
<p>Shortcut: select the cell below a column of numbers and press ${KEYS('Alt','=')} to insert a <code>SUM</code>. <strong>Formulas &gt; AutoSum</strong> has the other common functions.</p>
${TIP('<code>COUNTA(A:A)</code> counts the heading in A1 as well. To count records, take 1 away: <code>=COUNTA(A:A)-1</code>.')}`,
ex:[
  {p:'How many **minutes** did members spend in the gym altogether? (Use the **Visits** sheet.)', sheet:'Visits', sol:'=SUM(E:E)', hint:'SUM the DurationMins column, E.'},
  {p:'What was the **average visit length**, rounded to 1 decimal place?', sheet:'Visits', sol:'=ROUND(AVERAGE(E:E),1)', hint:'Put AVERAGE(E:E) inside ROUND( … , 1).'},
  {p:'How long was the **longest** visit, in minutes?', sheet:'Visits', sol:'=MAX(E:E)', hint:'MAX finds the largest number.'},
  {p:'How many **members** are on the Members sheet? Don\'t count the heading.', sheet:'Members', sol:'=COUNTA(A:A)-1', hint:'COUNTA counts cells that aren\'t empty, including the heading in A1.'},
  {p:'How many members have a **cancel date**? Use a function that counts numbers, since dates are numbers.', sheet:'Members', sol:'=COUNT(K:K)', hint:'COUNT only counts numbers, so it skips the heading and the empty cells. CancelDate is column K.'},
  {p:'What is the **cheapest** monthly fee on the Plans sheet?', sheet:'Plans', sol:'=MIN(C:C)', hint:'MIN finds the smallest number.'},
]},
'text-dates':{ title:'Text and dates', summary:'Join and split text, change its case, and work with dates and times.', sheet:'Members',
learn:`
<h2 class="dot">Joining text</h2>
<p>Use <code>&amp;</code> to join values together, with any extra text in double quotes: <code>=B2&amp;" "&amp;C2</code> gives a full name. <code>CONCAT</code> and <code>TEXTJOIN</code> do the same job for longer lists.</p>
${F('=B2&" "&C2','Members',{fill:true,col:'FullName'})}
<table class="mini"><tr><th>Function</th><th>Returns</th><th>Example</th></tr>
<tr><td><code>LEFT(text, n)</code> / <code>RIGHT(text, n)</code></td><td>the first or last n characters</td><td><code>=LEFT("TF3",2)</code> is TF</td></tr>
<tr><td><code>MID(text, start, n)</code></td><td>n characters from position start</td><td><code>=MID("M0042",2,4)</code> is 0042</td></tr>
<tr><td><code>LEN(text)</code></td><td>the number of characters</td><td><code>=LEN("Yoga")</code> is 4</td></tr>
<tr><td><code>UPPER</code>, <code>LOWER</code>, <code>PROPER</code></td><td>CAPITALS, lower case, Title Case</td><td><code>=PROPER("aLICE")</code> is Alice</td></tr>
<tr><td><code>TRIM(text)</code></td><td>the text without extra spaces</td><td><code>=TRIM("  Ann ")</code> is Ann</td></tr>
<tr><td><code>TEXT(value, format)</code></td><td>a number or date as text, in a format you choose</td><td><code>=TEXT(D2,"dddd")</code> is the day's name</td></tr></table>
<h2 class="dot">Dates and times</h2>
<p>Because dates are numbers, you can do arithmetic with them. <code>=I2+30</code> is 30 days after the join date, and <code>=K2-I2</code> is how many days someone was a member.</p>
<table class="mini"><tr><th>Function</th><th>Returns</th></tr>
<tr><td><code>DATE(y, m, d)</code></td><td>a date from a year, month and day</td></tr>
<tr><td><code>YEAR</code>, <code>MONTH</code>, <code>DAY</code></td><td>one part of a date</td></tr>
<tr><td><code>HOUR</code>, <code>MINUTE</code></td><td>one part of a time</td></tr>
<tr><td><code>WEEKDAY(date, 2)</code></td><td>the day of the week as a number, Monday = 1 to Sunday = 7</td></tr>
<tr><td><code>DATEDIF(start, end, "Y")</code></td><td>whole years between two dates (use "M" for months and "D" for days)</td></tr>
<tr><td><code>TODAY()</code></td><td>today's date, recalculated every time the workbook opens</td></tr></table>
${F('=DATEDIF(E2,DATE(2025,12,31),"Y")','Members',{fill:true,col:'Age',note:'Each member\'s age on 31 December 2025.'})}
${TIP('Prefer a fixed date such as <code>DATE(2025,12,31)</code> to <code>TODAY()</code> when you report on data for a set period. With <code>TODAY()</code>, the answer changes every day you open the file.')}
<p>Two less obvious points: <code>DATEDIF</code> isn't listed in Excel's function menu, but it works. And to see a date as a date, the cell needs a date format: <strong>Home &gt; Number Format &gt; Short Date</strong>.</p>`,
ex:[
  {p:'Make a **full name** column: first name, a space, then last name.', sheet:'Members', fill:true, col:'FullName', sol:'=B2&" "&C2', hint:'B2&" "&C2. The space goes inside double quotes.'},
  {p:'Make a column of each member\'s **postcode area**: the first two characters of the postcode.', sheet:'Members', fill:true, col:'Area', sol:'=LEFT(F2,2)', hint:'LEFT(text, number of characters). The postcode is in column F.'},
  {p:'Make a column with the **year each member joined**.', sheet:'Members', fill:true, col:'JoinYear', sol:'=YEAR(I2)', hint:'JoinDate is column I. YEAR pulls out the year.'},
  {p:'Work out each member\'s **age on 31 December 2025**, in whole years.', sheet:'Members', fill:true, col:'Age', sol:'=DATEDIF(E2,DATE(2025,12,31),"Y")', hint:'DATEDIF(date of birth, end date, "Y"). Build the end date with DATE(2025,12,31).', askRow:5},
  {p:'On the **Visits** sheet, show the **name of the day** each visit happened, such as Monday.', sheet:'Visits', fill:true, col:'Day', sol:'=TEXT(D2,"dddd")', hint:'TEXT(date, "dddd") gives the full day name. "ddd" gives Mon, Tue…'},
  {p:'On the **Visits** sheet, show the **hour** each member checked in (0 to 23).', sheet:'Visits', fill:true, col:'Hour', sol:'=HOUR(D2)', hint:'CheckIn (column D) holds a date and a time. HOUR pulls out the hour.'},
  {p:'Make an **email username** for each member: their member ID in lower case, followed by `@spartaleisure.example`.', sheet:'Members', fill:true, col:'Email', sol:'=LOWER(A2)&"@spartaleisure.example"', hint:'LOWER(A2), then & and the rest in double quotes.'},
]},
logic:{ title:'Logical functions', summary:'IF, AND, OR, NOT and nested IFs: formulas that make decisions.', sheet:'Visits',
learn:`
<p><code>IF</code> checks a condition and returns one value if it's TRUE and another if it's FALSE:</p>
<pre class="code">=IF(condition, value if TRUE, value if FALSE)</pre>
${F('=IF(E2>90,"Long","Normal")','Visits',{fill:true,col:'VisitType'})}
<h2 class="dot">AND and OR</h2>
<p><code>AND</code> is TRUE only when <em>every</em> condition is TRUE. <code>OR</code> is TRUE when <em>at least one</em> is. On their own they return TRUE or FALSE, and inside an <code>IF</code> they let one decision depend on several things.</p>
${F('=AND(H2="PREM",J2="Active")','Members',{fill:true,col:'ActivePremium'})}
${F('=IF(OR(H2="STU",H2="SEN"),"Concession","Full price")','Members',{fill:true,col:'PriceBand'})}
<p><code>NOT</code> turns TRUE into FALSE and FALSE into TRUE. <code>XOR</code> is TRUE when an odd number of its conditions are TRUE, so with two conditions, one or the other but not both.</p>
<h2 class="dot">Nested IFs</h2>
<p>To choose between more than two results, put another <code>IF</code> in the FALSE part. Excel checks the conditions in order and stops at the first TRUE one.</p>
${F('=IF(HOUR(D2)<12,"Morning",IF(HOUR(D2)<17,"Afternoon","Evening"))','Visits',{fill:true,col:'Session'})}
<p>In Excel 365 and 2019, <code>IFS</code> is easier to read: <code>=IFS(HOUR(D2)&lt;12,"Morning",HOUR(D2)&lt;17,"Afternoon",TRUE,"Evening")</code>. The final <code>TRUE</code> catches everything else.</p>
${TIP('Text in a condition needs double quotes, and it isn\'t case-sensitive: <code>H2="prem"</code> matches PREM. Numbers don\'t need quotes: <code>E2&gt;90</code>.')}`,
ex:[
  {p:'On the **Members** sheet, label each member **Current** if their status is Active, and **Former** if not.', sheet:'Members', fill:true, col:'Group', sol:'=IF(J2="Active","Current","Former")', hint:'Status is column J. Text values need double quotes.'},
  {p:'On the **Visits** sheet, label visits over 90 minutes **Long**, and the rest **Normal**.', sheet:'Visits', fill:true, col:'VisitType', sol:'=IF(E2>90,"Long","Normal")', hint:'IF(E2>90, "Long", "Normal")', askRow:3},
  {p:'On the **Members** sheet, show TRUE for members who are on the **Premium** plan (PREM) **and** still Active, and FALSE for everyone else.', sheet:'Members', fill:true, col:'ActivePremium', sol:'=AND(H2="PREM",J2="Active")', hint:'AND on its own returns TRUE or FALSE. PlanCode is H, Status is J.'},
  {p:'Label members on the Student (STU) **or** Senior (SEN) plan **Concession**, and everyone else **Full price**.', sheet:'Members', fill:true, col:'PriceBand', sol:'=IF(OR(H2="STU",H2="SEN"),"Concession","Full price")', hint:'Put OR(…) as the condition of an IF.'},
  {p:'On the **Visits** sheet, label each visit **Morning** (before 12:00), **Afternoon** (before 17:00) or **Evening**.', sheet:'Visits', fill:true, col:'Session', sol:'=IF(HOUR(D2)<12,"Morning",IF(HOUR(D2)<17,"Afternoon","Evening"))', hint:'Use HOUR(D2). Put a second IF where the first one\'s FALSE value would go.'},
  {p:'Peak time is **Monday to Friday, from 16:00 until 20:59**. Label each visit **Peak** or **Off-peak**.', sheet:'Visits', fill:true, col:'Peak', sol:'=IF(AND(WEEKDAY(D2,2)<=5,HOUR(D2)>=16,HOUR(D2)<21),"Peak","Off-peak")', hint:'Three conditions inside one AND: WEEKDAY(D2,2)<=5, HOUR(D2)>=16 and HOUR(D2)<21.'},
]},
errors:{ title:'Formula errors', summary:'What each error means, how to fix it, and when IFERROR is the right answer.', sheet:'MembersRaw',
learn:`
<p>When Excel can't work out a formula, it shows an error instead of a result. Each one points to a different problem.</p>
<table class="mini"><tr><th>Error</th><th>Means</th><th>Usual fix</th></tr>
<tr><td><code>#DIV/0!</code></td><td>dividing by zero or by an empty cell</td><td>check the divisor, or catch it with IFERROR</td></tr>
<tr><td><code>#N/A</code></td><td>a lookup couldn't find what it was looking for</td><td>check for typos, extra spaces and mismatched codes</td></tr>
<tr><td><code>#NAME?</code></td><td>Excel doesn't recognise a name: a misspelt function, or text without quotes</td><td>check the spelling, and put text in double quotes</td></tr>
<tr><td><code>#VALUE!</code></td><td>the wrong type of value, such as text where a number should be</td><td>check what's in the cells the formula uses</td></tr>
<tr><td><code>#REF!</code></td><td>a reference to a cell that no longer exists, often after deleting rows or columns</td><td>rebuild the reference</td></tr>
<tr><td><code>#NUM!</code></td><td>an impossible calculation, such as the square root of a negative number</td><td>check the inputs</td></tr>
<tr><td><code>#SPILL!</code></td><td>a formula that returns several values has no room to spill into</td><td>clear the cells below or beside it</td></tr>
<tr><td><code>####</code></td><td>not an error: the column is too narrow to show the value</td><td>widen the column</td></tr></table>
<p>Try these on the <strong>MembersRaw</strong> sheet, a messy extract from an old system:</p>
${F('=VLOOKUP(I2,Plans!A:C,3,FALSE)','MembersRaw',{fill:true,col:'PlanFee',note:'Fill it down: some plans are written in a way the Plans sheet doesn\'t recognise, so they show #N/A.'})}
<h2 class="dot">IFERROR</h2>
<p><code>=IFERROR(formula, value if error)</code> shows your own value instead of an error. It's useful when an error is expected and harmless. A lookup that finds nothing can show "Unknown", for example.</p>
${F('=IFERROR(VLOOKUP(I2,Plans!A:C,3,FALSE),"Unknown plan")','MembersRaw',{fill:true,col:'PlanFee'})}
${TIP('Don\'t use IFERROR to hide errors you don\'t understand. An #N/A from a lookup is often telling you the data needs cleaning. Hiding it means mistakes flow silently into your totals.')}
<p>To find what's wrong, select the cell and use <strong>Formulas &gt; Evaluate Formula</strong> to step through the calculation, or <strong>Trace Precedents</strong> to see which cells it uses.</p>`,
ex:[
  {p:'On **MembersRaw**, look up each row\'s monthly fee from the **Plans** sheet using the `Plan` column, and show **Unknown plan** wherever the lookup fails.', sheet:'MembersRaw', fill:true, col:'PlanFee', sol:'=IFERROR(VLOOKUP(I2,Plans!A:C,3,FALSE),"Unknown plan")', hint:'Wrap VLOOKUP(I2, Plans!A:C, 3, FALSE) in IFERROR( … , "Unknown plan"). Lookups get a full lesson later on.', askRow:14},
  {p:'On **SiteMonths**, work out **cancellations per new member** for each row, showing 0 if there were no new members that month.', sheet:'SiteMonths', fill:true, col:'CancelRatio', sol:'=IFERROR(D2/C2,0)', hint:'D2/C2 gives #DIV/0! when C2 is 0. Wrap it in IFERROR( … , 0).'},
],
check:[
  {q:'A cell shows <code>#NAME?</code> after you type <code>=IF(J2=Active,"Current","Former")</code>. Why?', a:['Active should be in double quotes','IF needs four arguments','J2 is empty','The column is too narrow'], c:0, why:'Without quotes, Excel thinks Active is the name of something, such as a function, and doesn\'t recognise it.'},
  {q:'Which error means a lookup didn\'t find its value?', a:['#REF!','#N/A','#VALUE!','#DIV/0!'], c:1, why:'#N/A means "not available": the value you looked up isn\'t in the table.'},
]},

/* ===================== Unit 4: Excel for analysis ===================== */
lookups:{ title:'Lookups: joining sheets', summary:'VLOOKUP, HLOOKUP, XLOOKUP and INDEX with MATCH: bring data from one sheet into another.', sheet:'Members',
learn:`
<p>Sparta Leisure's data is split across sheets that link through shared codes, just like tables in a database. The Members sheet has each member's <code>PlanCode</code>, but the fee is on the Plans sheet. A <strong>lookup</strong> uses the code to fetch the fee.</p>
<h2 class="dot">VLOOKUP</h2>
<pre class="code">=VLOOKUP(what to find, table, column number, FALSE)</pre>
<p>VLOOKUP looks <strong>down the first column</strong> of the table for what you're finding, then returns the value from the column number you give, counting the first column as 1. Always finish with <code>FALSE</code> for an exact match.</p>
${F('=VLOOKUP(H2,Plans!$A$2:$C$6,3,FALSE)','Members',{fill:true,col:'MonthlyFee'})}
${TIP('Make the table absolute (<code>$A$2:$C$6</code>), or use whole columns (<code>Plans!A:C</code>). Otherwise the table slides down as you fill the formula, and rows near the bottom start to show #N/A.')}
<p>VLOOKUP's limits: the value you look for must be in the table's <em>first</em> column, it can only return values to the right of it, and inserting a column breaks the column number.</p>
<h2 class="dot">HLOOKUP</h2>
<p>HLOOKUP is VLOOKUP turned on its side. It searches across the <strong>first row</strong> and returns from the row number you give. Use it when your headings run across the top, as the years do on the <strong>PriceHistory</strong> sheet.</p>
${F('=HLOOKUP(2023,PriceHistory!A1:E6,3,FALSE)','PriceHistory',{note:'The 2023 fee from row 3: the Standard plan.'})}
<h2 class="dot">XLOOKUP</h2>
<pre class="code">=XLOOKUP(what to find, where to look, what to return, [if not found])</pre>
<p>XLOOKUP, in Excel 365 and 2021, fixes VLOOKUP's limits. You give it the column to search and the column to return separately, so the return column can be anywhere. It matches exactly by default.</p>
${F('=XLOOKUP(G2,Sites!A:A,Sites!B:B)','Members',{fill:true,col:'SiteName'})}
<h2 class="dot">INDEX and MATCH</h2>
<p>Before XLOOKUP, analysts combined two functions. <code>MATCH</code> finds the position of a value in a list, and <code>INDEX</code> returns the value at a position. Together they can look in any direction, and they work in every version of Excel.</p>
${F('=INDEX(Classes!$B$2:$B$13,MATCH(C2,Classes!$A$2:$A$13,0))','Bookings',{fill:true,col:'ClassName'})}
${TIP('Lookups match on the exact value, so " STD" with a leading space won\'t match "STD". When a lookup shows #N/A, suspect the data before the formula.')}`,
ex:[
  {p:'On the **Members** sheet, bring in each member\'s **monthly fee** from the Plans sheet, using `VLOOKUP`.', sheet:'Members', fill:true, col:'MonthlyFee', sol:'=VLOOKUP(H2,Plans!$A$2:$C$6,3,FALSE)', hint:'Find H2 (PlanCode) in Plans!$A$2:$C$6, return column 3, with FALSE for an exact match.', askRow:6},
  {p:'Bring in the **name of each member\'s home site** from the Sites sheet.', sheet:'Members', fill:true, col:'SiteName', sol:'=XLOOKUP(G2,Sites!A:A,Sites!B:B)', hint:'XLOOKUP(G2, Sites!A:A, Sites!B:B). Or VLOOKUP(G2, Sites!A:B, 2, FALSE).'},
  {p:'On the **Visits** sheet, bring in the **plan code** of the member who made each visit.', sheet:'Visits', fill:true, col:'Plan', sol:'=XLOOKUP(B2,Members!A:A,Members!H:H)', hint:'Look for the MemberID (B2) in Members!A:A and return Members!H:H.'},
  {p:'What was the **Student** plan\'s monthly fee in **2023**? Use `HLOOKUP` on the PriceHistory sheet. (STU is in row 5.)', sheet:'PriceHistory', sol:'=HLOOKUP(2023,A1:E6,5,FALSE)', hint:'HLOOKUP(2023, A1:E6, 5, FALSE): find 2023 across row 1, return row 5.'},
  {p:'On the **Bookings** sheet, bring in each **class name** using `INDEX` and `MATCH`.', sheet:'Bookings', fill:true, col:'ClassName', sol:'=INDEX(Classes!$B$2:$B$13,MATCH(C2,Classes!$A$2:$A$13,0))', hint:'MATCH(C2, Classes!$A$2:$A$13, 0) finds the class\'s row. INDEX(Classes!$B$2:$B$13, that row) returns its name.'},
]},
conditional:{ title:'Counting and adding with conditions', summary:'COUNTIF, COUNTIFS, SUMIFS and AVERAGEIFS: summaries for just the rows you want.', sheet:'Members',
learn:`
<p>These functions answer "how many…?", "how much…?" and "what's the average…?" for only the rows that meet your conditions. They're the formula version of filtering, then reading the total.</p>
<table class="mini"><tr><th>Function</th><th>Shape</th></tr>
<tr><td><code>COUNTIF</code></td><td><code>=COUNTIF(range, criteria)</code></td></tr>
<tr><td><code>COUNTIFS</code></td><td><code>=COUNTIFS(range1, criteria1, range2, criteria2, …)</code></td></tr>
<tr><td><code>SUMIFS</code></td><td><code>=SUMIFS(range to add, range1, criteria1, …)</code></td></tr>
<tr><td><code>AVERAGEIFS</code></td><td><code>=AVERAGEIFS(range to average, range1, criteria1, …)</code></td></tr>
<tr><td><code>MAXIFS</code>, <code>MINIFS</code></td><td>the same shape as SUMIFS</td></tr></table>
${F('=COUNTIF(J:J,"Active")','Members')}
${F('=COUNTIFS(H:H,"PREM",J:J,"Active")','Members',{note:'With COUNTIFS, a row must meet every condition.'})}
${F('=AVERAGEIFS(E:E,C:C,"S02")','Visits')}
<h2 class="dot">Writing criteria</h2>
<table class="mini"><tr><th>Criteria</th><th>Matches</th></tr>
<tr><td><code>"Active"</code></td><td>exactly Active (not case-sensitive)</td></tr>
<tr><td><code>"&gt;90"</code>, <code>"&lt;=30"</code>, <code>"&lt;&gt;S01"</code></td><td>more than 90, 30 or less, anything except S01</td></tr>
<tr><td><code>"A*"</code>, <code>"DY1?"</code></td><td>wildcards: <code>*</code> is any characters, <code>?</code> is exactly one</td></tr>
<tr><td><code>"&gt;="&amp;DATE(2025,1,1)</code></td><td>on or after 1 January 2025: join the operator to a date or a cell with &amp;</td></tr>
<tr><td><code>A2</code></td><td>whatever is in cell A2. Filled down, it makes a summary table.</td></tr>
<tr><td><code>""</code></td><td>empty cells</td></tr></table>
${F('=COUNTIF(Members!$G:$G,A2)','Sites',{fill:true,col:'Members',note:'A small summary table: the number of members at each site.'})}
${TIP('In COUNTIFS and SUMIFS, every range must be the same size. Whole columns, such as <code>C:C</code>, take care of that for you.')}`,
ex:[
  {p:'How many members are **Active**?', sheet:'Members', sol:'=COUNTIF(J:J,"Active")', hint:'COUNTIF(the Status column, "Active").'},
  {p:'How many **Active** members are on the **Premium** plan (PREM)?', sheet:'Members', sol:'=COUNTIFS(H:H,"PREM",J:J,"Active")', hint:'COUNTIFS with two pairs: H:H with "PREM", and J:J with "Active".'},
  {p:'How many members **joined during 2025**?', sheet:'Members', sol:'=COUNTIFS(I:I,">="&DATE(2025,1,1),I:I,"<="&DATE(2025,12,31))', hint:'Use JoinDate (I:I) twice: ">="&DATE(2025,1,1) and "<="&DATE(2025,12,31).'},
  {p:'How many **minutes** did members spend at **Wolverhampton City** (S03)?', sheet:'Visits', sol:'=SUMIFS(E:E,C:C,"S03")', hint:'SUMIFS(the minutes, the SiteID column, "S03").'},
  {p:'What was the **average visit length at Shrewsbury Riverside** (S02)? Round it to 1 decimal place.', sheet:'Visits', sol:'=ROUND(AVERAGEIFS(E:E,C:C,"S02"),1)', hint:'AVERAGEIFS(E:E, C:C, "S02"), inside ROUND( … , 1).'},
  {p:'What fraction of class bookings were **attended**? Give it as a decimal.', sheet:'Bookings', sol:'=COUNTIF(G:G,TRUE)/(COUNTA(G:G)-1)', hint:'COUNTIF(G:G, TRUE) counts the attended ones. Divide by the number of bookings: COUNTA(G:G)-1.'},
  {p:'How much did Sparta Leisure spend on **marketing in January 2025**, across all sites?', sheet:'SiteMonths', sol:'=SUMIFS(F:F,B:B,DATE(2025,1,1))', hint:'Each Month cell holds the 1st of the month. SUMIFS(F:F, B:B, DATE(2025,1,1)).'},
  {p:'On the **Sites** sheet, count **how many members** each site has. Fill it down to build a summary table.', sheet:'Sites', fill:true, col:'Members', sol:'=COUNTIF(Members!$G:$G,A2)', hint:'COUNTIF(Members!$G:$G, A2). The criteria is the SiteID in the same row.'},
]},
'sort-filter':{ title:'Sorting and filtering', summary:'Put data in order, and show only the rows you need.', sheet:'Visits', excel:true,
learn:`
<p>Sorting and filtering are done with Excel's menus rather than formulas, so do this lesson in Excel. Download the workbook, then type what you find into the box on each exercise.</p>
<p>${DL_WORKBOOK}</p>
<h2 class="dot">Sorting</h2>
${STEPS(['Select any cell in the data (not a whole column).','Go to <strong>Data &gt; Sort</strong>.','Choose the column to sort by, and the order: smallest to largest, A to Z, or the reverse.','To break ties, select <strong>Add Level</strong> and choose a second column, such as SiteID, then LastName.'])}
${TIP('Never sort just one column on its own. The other columns stay where they are, and every row gets scrambled. Selecting a single cell first lets Excel find the whole table.')}
<h2 class="dot">Filtering</h2>
${STEPS(['Select a cell in the data and go to <strong>Data &gt; Filter</strong> (or press '+KEYS('Ctrl','Shift','L')+'). Arrows appear in the headings.','Select a heading\'s arrow and tick only the values you want, or use <strong>Text Filters</strong>, <strong>Number Filters</strong> or <strong>Date Filters</strong> for conditions.','Filter more than one column to combine conditions. A row must pass all of them.','Read the count in the status bar at the bottom: "125 of 2620 records found".','Select <strong>Data &gt; Clear</strong> to show every row again.'])}
<p>The sheets in the workbook already have filter arrows switched on.</p>
${TIP('Filtering only hides rows. <code>SUM</code> still adds up the hidden ones. To total only what you can see, use <code>=SUBTOTAL(9, E:E)</code>. Excel\'s AutoSum does this for you on a filtered list.')}`,
ex:[
  {p:'Sort the **Visits** sheet by `DurationMins`, largest to smallest. What is the **VisitID** of the longest visit?', sheet:'Visits', sol:'=INDEX(A:A,MATCH(MAX(E:E),E:E,0))', hint:'Data > Sort, by DurationMins, Largest to Smallest. The top row is the longest visit.', excel:true},
  {p:'Filter the **Members** sheet to `PlanCode` STU and `Status` Cancelled. How many records are found?', sheet:'Members', sol:'=COUNTIFS(H:H,"STU",J:J,"Cancelled")', hint:'Filter both columns, then read the count in the status bar.', excel:true},
  {p:'Filter **Members** to `HomeSiteID` S01, and use a **date filter** on `JoinDate` to show only 2025. How many members joined Telford Central in 2025?', sheet:'Members', sol:'=COUNTIFS(G:G,"S01",I:I,">="&DATE(2025,1,1),I:I,"<="&DATE(2025,12,31))', hint:'Date Filters > Between, or tick only 2025 in the date list.', excel:true},
  {p:'Sort the **Sites** sheet by `Capacity`, largest first. Which site is **third**? Type its SiteName.', sheet:'Sites', sol:'=INDEX(B:B,MATCH(LARGE(E:E,3),E:E,0))', hint:'Sort by Capacity, Largest to Smallest, and read the third row of data.', excel:true},
  {p:'Filter **Visits** to `DurationMins` greater than 300 (a Number Filter). How many visits lasted over 5 hours?', sheet:'Visits', sol:'=COUNTIF(E:E,">300")', hint:'Number Filters > Greater Than… 300.', excel:true},
]},
pivots:{ title:'Pivot tables and slicers', summary:'Summarise thousands of rows in seconds, then slice them by site, plan or date.', sheet:'Visits', excel:true,
learn:`
<p>A pivot table groups and summarises data without any formulas. Drag fields into place and Excel counts, adds up or averages them for every group. Do this lesson in Excel.</p>
<p>${DL_WORKBOOK}</p>
<h2 class="dot">Creating a pivot table</h2>
${STEPS(['Select any cell in the data, for example on the <strong>Visits</strong> sheet.','Go to <strong>Insert &gt; PivotTable</strong> and choose <strong>New Worksheet</strong>.','In the <strong>PivotTable Fields</strong> pane, drag fields into the four areas: <strong>Rows</strong> (the groups down the side), <strong>Columns</strong> (groups across the top), <strong>Values</strong> (what to calculate) and <strong>Filters</strong>.','Drag <code>SiteID</code> to Rows and <code>VisitID</code> to Values. Text fields are counted, so you get the number of visits at each site.','To change the calculation, select the field in Values, then <strong>Value Field Settings</strong>: Sum, Count, Average, Max and so on.'])}
${TIP('Dates in Rows are grouped by years, quarters and months automatically. Right-click a date and choose <strong>Group</strong> to change this.')}
<h2 class="dot">Slicers</h2>
<p>A slicer is a set of buttons that filters a pivot table. Slicers make the filter visible, so anyone reading the report can see what's selected.</p>
${STEPS(['Select a cell in the pivot table.','Go to <strong>PivotTable Analyze &gt; Insert Slicer</strong> and tick a field, such as <code>SiteID</code>.','Select a button to filter. Hold '+KEYS('Ctrl')+' to choose several, and select the clear-filter icon to show everything.','To connect one slicer to several pivot tables, right-click it and choose <strong>Report Connections</strong>.'])}
${TIP('A pivot table doesn\'t update by itself when its data changes. Select <strong>PivotTable Analyze &gt; Refresh</strong> (or '+KEYS('Alt','F5')+'). Formatting the data as a table first (<strong>Insert &gt; Table</strong>) means new rows are included when you refresh.')}`,
ex:[
  {p:'Make a pivot table from **Visits** with `SiteID` in Rows and the count of `VisitID` in Values. How many visits did **Dudley Castle Gate** (S05) have?', sheet:'Visits', sol:'=COUNTIF(C:C,"S05")', hint:'Find the S05 row in the pivot table.', excel:true},
  {p:'From **Members**, put `PlanCode` in Rows, `Status` in Columns and the count of `MemberID` in Values. How many **Off-peak** (OFF) members have **Cancelled**?', sheet:'Members', sol:'=COUNTIFS(H:H,"OFF",J:J,"Cancelled")', hint:'Read across the OFF row to the Cancelled column.', excel:true},
  {p:'From **Visits**, show the **average** of `DurationMins` for each `SiteID`. What is it for **Walsall Arboretum** (S04), to 1 decimal place?', sheet:'Visits', sol:'=ROUND(AVERAGEIFS(E:E,C:C,"S04"),1)', hint:'Value Field Settings > Average. Number Format can set 1 decimal place.', excel:true},
  {p:'From **SiteMonths**, put `Month` in Rows and the sum of `NewMembers` in Values. How many new members joined in **September 2025**, across all sites?', sheet:'SiteMonths', sol:'=SUMIFS(C:C,B:B,DATE(2025,9,1))', hint:'If Excel groups the months, expand the groups to find September.', excel:true},
  {p:'From **Bookings**, put `ClassID` in Rows and the count of `BookingID` in Values, then add a **slicer** on `SiteID` and select **S02**. How many **Aqua Fit** (C09) bookings were at Shrewsbury?', sheet:'Bookings', sol:'=COUNTIFS(C:C,"C09",D:D,"S02")', hint:'Insert Slicer > SiteID, select S02, then read the C09 row.', excel:true},
]},
dashboards:{ title:'Charts and dashboards in Excel', summary:'Turn pivot tables into charts and put them together on one page that answers a question.',
learn:`
<p>A dashboard puts the few numbers and charts that matter on one screen, so someone can see how things stand at a glance. In Excel, dashboards are usually built from pivot tables, pivot charts and slicers.</p>
<p>${DL_WORKBOOK}</p>
<h2 class="dot">Build a small dashboard</h2>
${STEPS([
 'Decide the question first, for example "How busy is each site, and when?". Every chart on the dashboard should help answer it.',
 'Make a pivot table for each chart, each on its own sheet: visits by <code>SiteID</code>, visits by month, and the average <code>DurationMins</code> by site.',
 'Select a cell in each pivot table and choose <strong>PivotTable Analyze &gt; PivotChart</strong>. Pick a chart that suits the data (the Data visualisation unit covers this): a bar chart for comparing sites, a line chart for months.',
 'Add a new sheet called <strong>Dashboard</strong>. Cut each chart (not the pivot table) and paste it there. Line the charts up in a grid.',
 'Insert a slicer for <code>SiteID</code>, then right-click it, choose <strong>Report Connections</strong> and tick every pivot table. Now one selection filters every chart.',
 'Add a title that says what the page shows, and two or three headline numbers in large text (total visits, active members) linked to cells with <code>=</code>.',
 'Tidy up: remove chart junk such as 3D effects, unneeded gridlines and legends that repeat the title. Hide the field buttons on the charts (<strong>PivotChart Analyze &gt; Field Buttons</strong>).'])}
<p>Microsoft has a full walkthrough: <a href="https://support.microsoft.com/en-gb/office/create-and-share-a-dashboard-with-excel-and-microsoft-groups-ad92a34d-38d0-4fdd-b8b1-58379aae746e" target="_blank" rel="noopener">Create and share a dashboard with Excel</a>.</p>
${TIP('Keep a dashboard to one screen. If someone has to scroll, it\'s a report. Lead with the most important number at the top left, where people look first.')}
<h2 class="dot">Practise more</h2>
<ul class="plain"><li><a href="https://support.microsoft.com/en-gb/office/excel-video-training-9bc05390-e94c-46af-a5b3-d7c22f6990bb" target="_blank" rel="noopener">Excel video training</a> from Microsoft Support</li>
<li><a href="https://excelexercises.com/lessons.html" target="_blank" rel="noopener">Excel Exercises</a>: practice problems with instant feedback</li></ul>`,
check:[
  {q:'You want one slicer to filter three pivot charts. What do you use?', a:['Data &gt; Filter on each chart','Report Connections on the slicer','A separate slicer for each chart','Conditional formatting'], c:1, why:'Report Connections links a slicer to several pivot tables, and the charts built on them follow.'},
  {q:'The data behind your pivot table has changed. What do you do so the dashboard shows the new figures?', a:['Nothing: it updates itself','Refresh the pivot tables','Rebuild every chart','Save the file as CSV'], c:1, why:'Pivot tables keep a copy of the data. Refresh (or Refresh All) to update them.'},
  {q:'What should you decide before building any charts?', a:['The colour scheme','The question the dashboard answers','How many slicers to add','The font'], c:1, why:'Start from the question. It tells you which numbers and charts belong on the page and which don\'t.'},
]},

/* ===================== Unit 5: Statistics ===================== */
averages:{ title:'Averages', summary:'Mean, median and mode: what each tells you, and when to use which.', sheet:'Visits',
learn:`
<p>An average sums up a set of numbers with one typical value. There are several kinds, and they can give very different answers.</p>
<h2 class="dot">Mean</h2>
<p>The <strong>mean</strong> is the total divided by the count. It's the most commonly used average, and it's easy to work out. It's the value you could replace every number with and still get the same total. But it's easily pulled about by <strong>outliers</strong>: a few very large or small values.</p>
<p>Try it: the mean of 5.00, 7.25, 6.12, 9.24, 8.00 and 9.21 is 7.47. Now mistype the last value as 921, and the mean jumps to 159.44, larger than five of the six values.</p>
<h2 class="dot">Median</h2>
<p>The <strong>median</strong> is the middle value once the numbers are in order. With an odd count, it's the middle number. With an even count, it's the mean of the two middle numbers. It splits the data into two equal halves, and it <strong>handles outliers well</strong>: for both sets above, the median is 7.63.</p>
<h2 class="dot">Mode</h2>
<p>The <strong>mode</strong> is the value that appears most often, a bit like a vote where the winner takes all. It's always a value that's really in the data, and it's the only average that works for categories ("the most popular plan"). There can be more than one mode, or none at all.</p>
${F('=AVERAGE(E:E)','Visits')}${F('=MEDIAN(E:E)','Visits')}${F('=MODE.SNGL(E:E)','Visits')}
<p>Sparta Leisure's mean visit is longer than its median. A handful of very long visits pull the mean up, while the median stays put. You'll see why in the Distributions lesson.</p>
<h2 class="dot">Other averages</h2>
<table class="mini"><tr><th>Average</th><th>Used for</th><th>Excel</th></tr>
<tr><td>Geometric mean</td><td>growth rates and percentages, such as a fee rising by different amounts each year</td><td><code>GEOMEAN</code></td></tr>
<tr><td>Harmonic mean</td><td>rates, something "per" something else, such as an average speed</td><td><code>HARMEAN</code></td></tr>
<tr><td>Weighted average</td><td>when some values count for more than others, such as course grades</td><td><code>=SUMPRODUCT(values, weights)/SUM(weights)</code></td></tr></table>`,
ex:[
  {p:'What is the **mean** visit length, in minutes?', sheet:'Visits', sol:'=AVERAGE(E:E)', hint:'AVERAGE of DurationMins.'},
  {p:'What is the **median** visit length?', sheet:'Visits', sol:'=MEDIAN(E:E)', hint:'MEDIAN of DurationMins.'},
  {p:'What is the **modal** (most common) visit length?', sheet:'Visits', sol:'=MODE.SNGL(E:E)', hint:'MODE.SNGL, or MODE in older versions of Excel.'},
  {p:'On **SiteMonths**, what is the **median** number of new members a site gets in a month?', sheet:'SiteMonths', sol:'=MEDIAN(C:C)', hint:'MEDIAN of NewMembers, column C.'},
  {p:'On **SiteMonths**, work out the **average marketing spend per new member** across the year, weighting each month by its new members: the total spend divided by the total new members. Round it to 2 decimal places.', sheet:'SiteMonths', sol:'=ROUND(SUM(F:F)/SUM(C:C),2)', hint:'Total spend (SUM of F) ÷ total new members (SUM of C). The mean of each month\'s cost per member would treat quiet months the same as busy ones.'},
]},
spread:{ title:'Spread and standard deviation', summary:'Range, quartiles and standard deviation: how far the values sit from the average.', sheet:'Visits',
learn:`
<p>Two sets of data can share an average and still be completely different. Spread describes how far values sit from the middle.</p>
<h2 class="dot">Range and quartiles</h2>
<p>The <strong>range</strong> is the largest value minus the smallest. It's simple, but one outlier can stretch it. <strong>Quartiles</strong> split ordered data into four equal groups. The lower quartile (Q1) has a quarter of the values below it, and the upper quartile (Q3) has a quarter above it. The <strong>interquartile range</strong> (Q3 − Q1) is the spread of the middle half, and outliers don't affect it.</p>
${F('=QUARTILE.INC(E:E,3)-QUARTILE.INC(E:E,1)','Visits',{note:'The interquartile range of visit lengths.'})}
<h2 class="dot">Standard deviation</h2>
<p>The <strong>standard deviation</strong> (SD) measures how far values typically sit from the mean. A larger SD means more spread. To work it out by hand:</p>
${STEPS(['Subtract the mean from each value.','Square each of those differences.','Add up the squares.','Divide by the count (this gives the <strong>variance</strong>).','Take the square root.'])}
${F('=STDEV.P(E:E)','Visits')}
<p>Excel has two versions. <code>STDEV.P</code> is for a whole <strong>population</strong>: every visit in 2025. <code>STDEV.S</code> is for a <strong>sample</strong> drawn from a bigger population, and divides by the count minus 1 to make up for the missing data. With thousands of values the two are almost the same.</p>
<p>For data with a <strong>normal distribution</strong>, the SD tells you a lot:</p>
${FIG_NORMAL}
${TIP('Like the mean, the SD is easily skewed by outliers. A few 10-hour "visits" from people who forgot to check out make Sparta Leisure\'s SD much bigger than it would otherwise be.')}`,
ex:[
  {p:'What is the **standard deviation** of visit lengths? Treat the visits as a whole population.', sheet:'Visits', sol:'=STDEV.P(E:E)', hint:'STDEV.P(E:E). STDEV.S gives almost the same answer here, and is also accepted.'},
  {p:'What is the **range** of visit lengths?', sheet:'Visits', sol:'=MAX(E:E)-MIN(E:E)', hint:'MAX minus MIN.'},
  {p:'What is the **upper quartile** (Q3) of visit lengths?', sheet:'Visits', sol:'=QUARTILE.INC(E:E,3)', hint:'QUARTILE.INC(range, 3).'},
  {p:'What **fraction** of visits are within 1 standard deviation of the mean? Is it close to the 68% a normal distribution would give?', sheet:'Visits', sol:'=COUNTIFS(E:E,">="&AVERAGE(E:E)-STDEV.P(E:E),E:E,"<="&AVERAGE(E:E)+STDEV.P(E:E))/COUNT(E:E)', hint:'COUNTIFS(E:E, ">="&(mean − SD), E:E, "<="&(mean + SD)), divided by COUNT(E:E).'},
]},
distributions:{ title:'Distributions and outliers', summary:'Histograms, the normal curve, skew, and what to do with extreme values.', sheet:'Visits',
learn:`
<p>How data is <strong>distributed</strong>, meaning where the values bunch up and how they tail off, decides which averages and charts you can trust.</p>
<h2 class="dot">Histograms</h2>
<p>A histogram puts values into <strong>bins</strong> (ranges of equal width) and draws a bar for how many fall in each. Too many bins make the chart spiky and hard to read. Too few hide the shape. Try it:</p>
${W('histogram')}
<h2 class="dot">The normal distribution</h2>
<p>Also called the <strong>bell curve</strong>. The data is symmetrical: any value is as likely to be above the mean as below it, and the mean and median sit together in the middle. Heights, and errors in measurements, are often roughly normal.</p>
<h2 class="dot">Skew</h2>
${FIG_SKEW}
<p><strong>Positive skew</strong> has its tail towards the larger values. Usually the mode is less than the median, which is less than the mean (but not always). It can point to a <strong>floor effect</strong>: a minimum value, such as zero, but no limit on the maximum. Household incomes and visit lengths are positively skewed.</p>
<p><strong>Negative skew</strong> has its tail towards the smaller values, and can point to a <strong>ceiling effect</strong>: a maximum possible value, such as 100% in an easy exam.</p>
<p>Excel's <code>SKEW</code> function gives a number: positive for a tail to the right, negative for a tail to the left, and near 0 for symmetrical data.</p>
<h2 class="dot">Outliers</h2>
<p>A few very extreme values can have a big effect on the mean and the standard deviation. They might be meaningful (a marathon training session) or errors in data entry or transfer (a member who forgot to check out). A height of 2.72 m is extreme but has really happened. A height of 17.53 m is an error.</p>
<p>A common rule of thumb flags a value as an outlier if it's more than <strong>1.5 × the interquartile range</strong> above Q3 or below Q1.</p>
<p>There are no hard and fast rules about what to do with them:</p>
<ul class="plain"><li>Remove obvious errors and impossible values.</li><li>Keep meaningful values that you need to account for and that might happen again.</li><li>Be open about it: say what you removed and why. It's a very subjective area, so take great care.</li></ul>`,
ex:[
  {p:'What is the **skew** of visit lengths? (Positive means a tail to the right.)', sheet:'Visits', sol:'=SKEW(E:E)', hint:'SKEW(E:E).'},
  {p:'How many visits lasted **more than 5 hours** (300 minutes)? These are almost certainly members who forgot to check out.', sheet:'Visits', sol:'=COUNTIF(E:E,">300")', hint:'COUNTIF(E:E, ">300").'},
  {p:'What is the **mean** visit length if you leave out visits over 300 minutes?', sheet:'Visits', sol:'=AVERAGEIF(E:E,"<=300")', hint:'AVERAGEIF(E:E, "<=300"). Compare it with the mean of every visit.'},
  {p:'Using the **1.5 × IQR rule**, how many visits are high outliers? (Count values above Q3 + 1.5 × (Q3 − Q1).)', sheet:'Visits', sol:'=COUNTIF(E:E,">"&(QUARTILE.INC(E:E,3)+1.5*(QUARTILE.INC(E:E,3)-QUARTILE.INC(E:E,1))))', hint:'Build the limit with QUARTILE.INC(E:E,3) and QUARTILE.INC(E:E,1), then COUNTIF(E:E, ">"&limit).'},
]},
correlation:{ title:'Correlation', summary:'How two measures move together, r and R², and why correlation isn\'t causation.', sheet:'SiteMonths',
learn:`
<p>A <strong>correlation</strong> describes the straight-line relationship between two continuous variables: as one changes, how does the other change? A <strong>scatter plot</strong> shows it, with one dot per row.</p>
${W('scatter')}
<h2 class="dot">The correlation coefficient, r</h2>
<p><strong>r</strong> is always between −1 and 1:</p>
<ul class="plain"><li><strong>r = 1</strong>: perfect positive correlation. As x increases, y increases.</li>
<li><strong>r = 0</strong>: no correlation. As x changes, there's no consistent change in y.</li>
<li><strong>r = −1</strong>: perfect negative correlation. As x increases, y decreases.</li></ul>
${F('=CORREL(F:F,C:C)','SiteMonths',{note:'Marketing spend against new members.'})}
<h2 class="dot">R²</h2>
<p><strong>R²</strong> (r squared) is often used instead, as a percentage between 0 and 100%. It's worked out from how far each point is from the line of best fit, and it's the share of the variation in one variable that the other can explain. R² = 75% means 75% of the changes in y can be accounted for by changes in x, and vice versa: R² can't tell you which one causes which.</p>
${F('=RSQ(C:C,F:F)','SiteMonths')}
<h2 class="dot">Correlation does not imply causation</h2>
<p>Two things moving together doesn't mean one causes the other. Both might be driven by a third thing, or it might be chance. At Sparta Leisure, colder months have more new members. The cold doesn't make people join: January is cold, and January is when people make New Year's resolutions. Marketing spend is also highest in January. So how much of that strong correlation is marketing, and how much is January?</p>
<p>Tyler Vigen's <a href="http://tylervigen.com/spurious-correlations" target="_blank" rel="noopener">spurious correlations</a> collects real data that correlates by pure coincidence.</p>`,
ex:[
  {p:'What is the correlation (r) between **marketing spend** and **new members**?', sheet:'SiteMonths', sol:'=CORREL(F:F,C:C)', hint:'CORREL(MarketingSpend, NewMembers): columns F and C.'},
  {p:'What is **R²** for marketing spend and new members?', sheet:'SiteMonths', sol:'=RSQ(C:C,F:F)', hint:'RSQ(known y values, known x values), or CORREL(…)^2.'},
  {p:'What is the correlation between the **average temperature** and **new members**?', sheet:'SiteMonths', sol:'=CORREL(G:G,C:C)', hint:'AvgTempC is column G.'},
  {p:'What is the correlation between the **average temperature** and the number of **visits**?', sheet:'SiteMonths', sol:'=CORREL(G:G,E:E)', hint:'Visits is column E.'},
]},

/* ===================== Unit 6: Data quality ===================== */
quality:{ title:'Data quality', summary:'Validity, accuracy, completeness, consistency, uniformity and redundancy.', sheet:'MembersRaw',
learn:`
<p><strong>Data quality</strong> is how well data represents the real situation it describes. Poor-quality data leads to wrong answers, however clever the analysis. Before cleaning anything, find out what's wrong.</p>
<h2 class="dot">Spot the problems</h2>
<p>Here's <strong>MembersRaw</strong>, an extract from Sparta Leisure's old membership system. Look through it and note every problem you can find before reading on.</p>
${W('sheet',{sheet:'MembersRaw'})}
<h2 class="dot">Six dimensions of data quality</h2>
<table class="mini"><tr><th>Dimension</th><th>The degree to which…</th><th>In MembersRaw</th></tr>
<tr><td><strong>Validity</strong></td><td>data follows its rules and constraints: data types, ranges, mandatory and unique fields, and patterns (regular expressions)</td><td>a date of birth of 01/01/1900, a fee of 3299</td></tr>
<tr><td><strong>Accuracy</strong></td><td>data reflects the true value or a standard. Absolute accuracy is very hard to achieve: data comes from samples, and has input errors.</td><td>is that 3299 really £32.99?</td></tr>
<tr><td><strong>Completeness</strong></td><td>every required data point is known: no missing values or cut-off information</td><td>blank genders, postcodes and fees</td></tr>
<tr><td><strong>Consistency</strong></td><td>values agree with each other: no contradictions, or different details for the same person</td><td>the same member listed twice</td></tr>
<tr><td><strong>Uniformity</strong></td><td>data uses the same units and formats everywhere: weights, currency, dates, addresses</td><td>dates written three ways, plans as STD, Standard and std, a site as Telford instead of S01</td></tr>
<tr><td><strong>Redundancy</strong></td><td>the same information isn't held more than once, as a complete copy or by being entered twice</td><td>repeated rows</td></tr></table>
${TIP('Phone numbers in MembersRaw have lost their leading 0, because the old system stored them as numbers. That\'s a validity problem caused by the wrong data type.')}`,
check:[
  {q:'Plans are written as <code>STD</code>, <code>Standard</code> and <code>std</code>. Which dimension does this fail?', a:['Completeness','Uniformity','Redundancy','Accuracy'], c:1, why:'The same thing is recorded in different formats, so the data isn\'t uniform.'},
  {q:'A member\'s date of birth is 01/01/1900. Which dimension is most clearly broken?', a:['Validity','Redundancy','Uniformity','Completeness'], c:0, why:'No member is 125 years old. The value breaks a sensible range rule, so it isn\'t valid. It\'s probably a placeholder someone typed when the real date wasn\'t known.'},
  {q:'The same member appears in two identical rows. What\'s the problem?', a:['Uniformity','Validity','Redundancy','Accuracy'], c:2, why:'The same information is held twice. Left in, the member is counted twice.'},
  {q:'Several rows have no postcode. Which dimension does this affect?', a:['Completeness','Consistency','Uniformity','Validity'], c:0, why:'Required data points are missing.'},
]},
cleaning:{ title:'Cleaning data', summary:'Fix case, spaces, padding, units, duplicates and types with Excel formulas.', sheet:'MembersRaw',
learn:`
<p>Data cleaning fixes the problems you found. Work on a copy, or in new columns, so you can always get back to the original.</p>
<table class="mini"><tr><th>Technique</th><th>What it means</th><th>In Excel</th></tr>
<tr><td>Irrelevant data</td><td>drop or hide columns and rows you don't need for the question. Be careful: never delete without a backup.</td><td>hide columns, or delete them in a copy</td></tr>
<tr><td>Duplicates</td><td>remove rows that are repeated in full. If only some details repeat, look closer before deciding.</td><td><strong>Data &gt; Remove Duplicates</strong>, or flag them with <code>COUNTIF</code></td></tr>
<tr><td>Type conversions</td><td>store numbers as numbers and dates as dates. A value that won't convert is probably wrong: fix it, mark it as missing, or remove it.</td><td><code>VALUE</code>, <code>DATEVALUE</code></td></tr>
<tr><td>Padding</td><td>add characters so values are a set width, such as leading zeros on IDs</td><td><code>TEXT(F2,"00000000000")</code></td></tr>
<tr><td>Typos</td><td>map misspellings to the right value, by hand or by matching patterns</td><td><code>SUBSTITUTE</code>, Find &amp; Replace</td></tr>
<tr><td>Standardisation</td><td>put every value in the same format: the same case, precision, units, currency and date format</td><td><code>PROPER</code>, <code>TRIM</code>, <code>UPPER</code>, <code>ROUND</code></td></tr>
<tr><td>Scaling and normalisation</td><td>rescale numbers to a common range so they can be compared. Normalisation squeezes values into 0 to 1: (x − min) ÷ (max − min).</td><td><code>=(E2-MIN(E:E))/(MAX(E:E)-MIN(E:E))</code></td></tr></table>
${F('=PROPER(TRIM(B2))','MembersRaw',{fill:true,col:'FirstName'})}
${F('=TEXT(F2,"00000000000")','MembersRaw',{fill:true,col:'PhoneText',note:'Pads each phone number back to 11 digits, as text.'})}
${TIP('Watch out: <strong>scaling</strong>, <strong>normalisation</strong> and <strong>standardisation</strong> are often used to mean the same thing, and different tools define them differently. Check what someone means.')}
<p>When the cleaned column is right, copy it and use <strong>Paste Special &gt; Values</strong> to replace the formulas with plain values. Then you can delete the messy original.</p>`,
ex:[
  {p:'Clean the **first names** on MembersRaw: remove extra spaces and make them Title Case.', sheet:'MembersRaw', fill:true, col:'CleanFirst', sol:'=PROPER(TRIM(B2))', hint:'TRIM inside PROPER: PROPER(TRIM(B2)).', askRow:3},
  {p:'The phone numbers lost their leading zero. Turn each one back into **11-digit text**.', sheet:'MembersRaw', fill:true, col:'PhoneText', sol:'=TEXT(F2,"00000000000")', hint:'TEXT(F2, "00000000000"), with eleven zeros.'},
  {p:'Standardise **gender** to a single capital letter: F, M or X. (Female becomes F, m becomes M, and a blank stays blank.)', sheet:'MembersRaw', fill:true, col:'GenderCode', sol:'=UPPER(LEFT(D2,1))', hint:'Take the first character with LEFT, then UPPER.'},
  {p:'Some fees were typed in **pence**. Any fee over 100 is in pence: divide it by 100. Leave the others as they are.', sheet:'MembersRaw', fill:true, col:'FeeFixed', sol:'=IF(J2>100,J2/100,J2)', hint:'IF(J2>100, J2/100, J2).', askRow:28},
  {p:'Flag **duplicate** member IDs: show TRUE where a MemberID appears more than once in column A.', sheet:'MembersRaw', fill:true, col:'IsDuplicate', sol:'=COUNTIF(A:A,A2)>1', hint:'COUNTIF(A:A, A2) counts how often this ID appears. Compare it with 1.', askRow:15},
  {p:'Convert the text dates of birth into **real dates**. (Excel reads 01/02/1990 as 1 February in the UK.)', sheet:'MembersRaw', fill:true, col:'DOBDate', sol:'=DATEVALUE(E2)', hint:'DATEVALUE turns date text into a date serial number.', vtype:'date'},
]},
missing:{ title:'Missing values', summary:'Drop, impute, hot-deck or flag: four ways to deal with gaps.', sheet:'MembersRaw',
learn:`
<p>Almost every real dataset has gaps. First count them, and ask <em>why</em> they're missing. Then choose a method.</p>
${F('=COUNTBLANK(D2:D52)','MembersRaw',{note:'Blank genders in MembersRaw.'})}
<table class="mini"><tr><th>Method</th><th>How</th><th>Watch out</th></tr>
<tr><td><strong>Drop</strong></td><td>delete rows with missing values, or drop a column that's mostly empty. You can drop based on criteria.</td><td>you lose data, and if the gaps aren't random you bias what's left</td></tr>
<tr><td><strong>Impute</strong></td><td>work out a value from other observations: the mean, the median, or a linear regression</td><td>it shrinks the spread, and makes the data look more certain than it is</td></tr>
<tr><td><strong>Hot-deck</strong></td><td>copy the value from a similar record, chosen at random or in sequence. Works for numbers and categories.</td><td>"similar" needs a sensible definition</td></tr>
<tr><td><strong>Flag</strong></td><td>record that the value is missing: fill numbers with a set value such as 0, and categories with "Missing"</td><td>a 0 must never be mistaken for a real zero</td></tr></table>
${F('=IF(G2="","Missing",G2)','MembersRaw',{fill:true,col:'PostcodeFlag'})}
${TIP('That data is missing is information in itself. If older members are less likely to give their gender, dropping those rows makes your members look younger than they are.')}`,
ex:[
  {p:'How many rows of MembersRaw (rows 2 to 52) have **no gender**?', sheet:'MembersRaw', sol:'=COUNTBLANK(D2:D52)', hint:'COUNTBLANK(D2:D52).'},
  {p:'How many rows have **no postcode**?', sheet:'MembersRaw', sol:'=COUNTBLANK(G2:G52)', hint:'Postcode is column G.'},
  {p:'**Flag** missing postcodes: show the postcode, or the word Missing if it\'s blank.', sheet:'MembersRaw', fill:true, col:'PostcodeFlag', sol:'=IF(G2="","Missing",G2)', hint:'IF(G2="", "Missing", G2).', askRow:32},
  {p:'**Impute** missing fees: where the fee is blank, look it up from the Plans sheet using the Plan code. Otherwise keep the fee as it is.', sheet:'MembersRaw', fill:true, col:'FeeFilled', sol:'=IF(J2="",VLOOKUP(I2,Plans!$A$2:$C$6,3,FALSE),J2)', hint:'IF(J2="", VLOOKUP(I2, Plans!$A$2:$C$6, 3, FALSE), J2).', askRow:24},
]},

/* ===================== Unit 7: Critical thinking ===================== */
simple:{ title:'Be sceptical: the SIMPLE checklist', summary:'Source, Irrelevant, Missing, Presentation, aLternatives and Evidence.',
learn:`
<blockquote class="quote">"Those who rule data will rule the entire world." <cite>Masayoshi Son, CEO of SoftBank</cite></blockquote>
<p>Data is only as good as the thinking around it. Two stories show why.</p>
<h2 class="dot">New Coke, 1985</h2>
<p>A new recipe was taste-tested on 200,000 people, and it almost always beat both the current version and its main competitor. Should you replace the old product? Coca-Cola did. It released New Coke to compete with Pepsi's sweeter taste, and pulled the original from the shelves. Customers were furious, and the change cost tens of millions of dollars. People buy for more than taste alone. Such a big decision needed more, and different, data.</p>
<h2 class="dot">The V-2 rocket, 1945</h2>
<p>You're at war and firing a new missile. Your agents report it's overshooting the target by 10 to 20 miles. Should you recalibrate? In the Second World War, German V-2 rockets were killing large numbers of people in London. Through double agents, the British reported that the rockets were overshooting, using places damaged by other bombs as evidence, and blamed explosions in the capital on gas leaks. The Germans recalibrated, and over half the V-2s fell short of London, in far less populated areas. The data was deliberately false, and nobody questioned where it came from.</p>
<h2 class="dot">The SIMPLE checklist</h2>
<div class="simple">
<div><b>S</b><span><strong>Source.</strong> Is it biased or unreliable? Who collected the data, and who stands to gain from it? How was it collected, in what context, and could the method have influenced the results? Does it measure what it claims to?</span></div>
<div><b>I</b><span><strong>Irrelevant.</strong> Separate the relevant from the irrelevant. With large datasets, you can't and shouldn't analyse everything. Is the data about the right topic, the right time period, and the right culture or region?</span></div>
<div><b>M</b><span><strong>Missing.</strong> Do we have all the data needed to be confident? Is someone leaving out data that's easy to get because it doesn't fit their story? What one piece of data would turn an assertion into a fact?</span></div>
<div><b>P</b><span><strong>Presentation.</strong> Is the data presented fairly: the right chart type, axes, units, scale and wording? Has anyone tried to hide the real data? Watch totals versus percentages especially.</span></div>
<div><b>L</b><span><strong>Alternatives.</strong> The most obvious explanation isn't always the right one, especially when you expected a certain answer. A different explanation could lead to a completely different conclusion.</span></div>
<div><b>E</b><span><strong>Evidence.</strong> Are the claims supported? Tell assertions ("this is caused by…") from suppositions ("this might be caused by…"), and facts ("Birmingham has a larger population than Manchester") from opinions ("Birmingham is better than Manchester").</span></div>
</div>`,
check:[
  {q:'A supplement company funds a study that finds its product boosts gym performance. Which part of SIMPLE should you check first?', a:['Presentation','Source','Irrelevant','Evidence'], c:1, why:'Who collected the data and who gains from it are Source questions. Funding doesn\'t make a study wrong, but it\'s a reason to look harder at the method.'},
  {q:'A chart of Sparta Leisure\'s new members starts its y-axis at 90 instead of 0, so a small rise looks huge. Which part of SIMPLE is this?', a:['Missing','Alternatives','Presentation','Source'], c:2, why:'The data may be right, but the axis exaggerates the change. That\'s presentation.'},
  {q:'"Our new yoga classes might be why visits rose." Is that an assertion or a supposition?', a:['An assertion','A supposition','A fact','An opinion'], c:1, why:'"Might be" makes it a supposition. An assertion would say "is why".'},
  {q:'A report on UK gym habits uses only data from 2015. Which question does that raise?', a:['Is it relevant to today?','Is the chart the right type?','Is it an opinion?','Is it uniform?'], c:0, why:'Data from the wrong time period may not be relevant to the question you\'re answering.'},
]},
alternatives:{ title:'Alternative explanations and evidence', summary:'Test a conclusion by asking what else could explain the data, and what would settle it.',
learn:`
<p>When a result matches what you hoped, it's tempting to stop looking. Instead, ask what <em>else</em> could explain it, and what data would tell the explanations apart.</p>
<h2 class="dot">Practise on three scenarios</h2>
<p>For each one, discuss the possible alternative explanations, and what further data you'd need to be more confident in the conclusion.</p>
<div class="scen"><p><b>1.</b> A numeracy intervention is tested on a class of 30 children. For six months, they spend three hours every Saturday learning a new mental maths technique. Their average maths score rises by 50%. <em>Therefore, the intervention is effective.</em></p>
<details><summary>Some alternatives</summary><p>Children improve over six months anyway, as they get older and are taught at school. Eighteen extra hours of maths on Saturdays might help whatever the technique. Children who came every week may have been the keenest, and the test may have been easier. You'd want a <strong>control group</strong> of similar children who didn't take part.</p></details></div>
<div class="scen"><p><b>2.</b> A one-year project helps farmers in Dorset apply fertiliser to their crops in winter. In the following years, these farmers have higher incomes and well-being than farmers in neighbouring regions. <em>Therefore, the project would work in the West Midlands.</em></p>
<details><summary>Some alternatives</summary><p>Dorset's farmers may have been better off to begin with, and the weather, soil, crops or prices may have differed. Did the farmers who chose to take part differ from those who didn't? And would West Midlands soils and crops respond the same way? Compare incomes <em>before and after</em>, against similar farmers who didn't take part.</p></details></div>
<div class="scen"><p><b>3.</b> A law makes it compulsory for cyclists to wear helmets. Over the next two years, the number of head injuries among cyclists halves. <em>Therefore, the law has improved safety for cyclists.</em></p>
<details><summary>Some alternatives</summary><p>Maybe fewer people cycled once helmets were compulsory, so there were fewer injuries of every kind. Roads, cycle lanes or the way injuries were recorded may have changed too. Look at injuries <em>per mile cycled</em>, and at other kinds of injury.</p></details></div>
<h2 class="dot">A Sparta Leisure example</h2>
<p>The marketing team writes: <em>"January sign-ups were more than four times December's. Our New Year campaign worked."</em> Before you agree, what would you check? Every January is busy for gyms, so compare January with January the year before, and look at what happened in years with a smaller campaign. You'll test exactly this claim in the capstone project.</p>`,
check:[
  {q:'After a summer price cut, Sparta Leisure\'s visits fell. The finance team says the price cut put people off. What\'s the strongest alternative explanation?', a:['Summer is usually quieter for gyms','The price cut was too small','Visits are always recorded wrongly','Members prefer higher prices'], c:0, why:'Gym visits dip every summer. Compare the summer with previous summers before blaming the price cut.'},
  {q:'What data would best test whether the New Year campaign raised January sign-ups?', a:['January\'s sign-ups only','January sign-ups compared with Januaries in years with smaller campaigns','The number of adverts','December\'s visits'], c:1, why:'Comparing like with like (January against January) removes the seasonal effect. Then you can see what the extra campaign added.'},
  {q:'A study compares people who chose to join a running club with people who didn\'t, and finds the members are healthier. Why can\'t we say the club made them healthier?', a:['Healthier people may be more likely to join in the first place','Running is bad for you','The sample was too large','Clubs keep poor records'], c:0, why:'People chose whether to join, so the groups were probably different before they started. This is called selection bias.'},
]},

/* ===================== Unit 8: Visualisation ===================== */
'good-viz':{ title:'What makes a good visualisation', summary:'Why we chart data at all, Anscombe\'s Quartet, and Andy Kirk\'s three principles.',
learn:`
<p>Businesses need data presented clearly because people make decisions from it, often quickly and without the raw numbers. A good chart shows patterns that a table of statistics hides.</p>
<h2 class="dot">Anscombe's Quartet</h2>
<p>Here are four datasets. Each has two variables, X and Y, with the same summary statistics: mean of X = 9, mean of Y = 7.50, variance of X = 11, variance of Y = 4.12, correlation 0.816, and line of best fit Y = 0.5X + 3. Do the statistics tell you how the data is distributed?</p>
${W('anscombe')}
<p>Only one of the four looks like what the statistics suggest. Always plot your data before you trust a summary of it.</p>
<h2 class="dot">Andy Kirk's three principles</h2>
<table class="mini"><tr><th>Principle</th><th>Means</th></tr>
<tr><td><strong>Trustworthy</strong></td><td>accurate and honest: true scales, clear sources, nothing hidden</td></tr>
<tr><td><strong>Accessible</strong></td><td>your audience can perceive it, understand it and interpret it. As Zach Gemignani puts it, data visualisation is like family photos: if you don't know the people in the picture, a beautiful composition won't keep your attention.</td></tr>
<tr><td><strong>Elegant</strong></td><td>often this means simple: remove anything that distracts from understanding. You can still add flair, if it's appropriate and fits the theme. The data artwork <em>If Vienna were an Apartment</em> is a good example.</td></tr></table>
<h2 class="dot">Tables</h2>
<p>Sometimes a table is the right choice, when people need to look up exact values. To format one well:</p>
<ul class="plain"><li>Right-align numbers (in a font where the digits line up), and left-align text.</li><li>Space the columns equally, and use colour sparingly.</li><li>Use consistent measurements with clear units, at a precision that suits your audience: exact, or to the nearest thousand.</li><li>"Remove to improve": cut every line, colour and decimal place that isn't needed.</li></ul>
<p>A <strong>highlight table</strong> (a heatmap) shades each cell by its value, so patterns jump out. In Excel, use <strong>Home &gt; Conditional Formatting &gt; Color Scales</strong>.</p>`,
check:[
  {q:'What does Anscombe\'s Quartet show?', a:['Summary statistics can hide very different data, so plot it','Correlation always means causation','Four bins is the right number for a histogram','Tables are better than charts'], c:0, why:'All four datasets share the same statistics but look completely different when plotted.'},
  {q:'Which of Kirk\'s principles does a chart break if its y-axis doesn\'t start at zero, to exaggerate a change?', a:['Elegant','Accessible','Trustworthy','None of them'], c:2, why:'A misleading scale makes the chart untrustworthy.'},
  {q:'In a table, how should you align numbers?', a:['Centred','Left','Right','It doesn\'t matter'], c:2, why:'Right-aligned numbers line up by place value, so they\'re easy to compare down a column.'},
]},
charts:{ title:'Choosing the right chart', summary:'Match the chart to the relationship in your data, and avoid the traps.',
learn:`
<p>Start from the relationship you want to show, then pick a chart that suits it.</p>
<table class="mini"><tr><th>Relationship</th><th>Shows</th><th>Good charts</th></tr>
<tr><td>Magnitude</td><td>how big things are, compared</td><td>bar chart, clustered bar</td></tr>
<tr><td>Time series</td><td>change over time</td><td>line chart, small multiples, slope chart</td></tr>
<tr><td>Ranking</td><td>order of size</td><td>sorted bar chart</td></tr>
<tr><td>Part-to-whole</td><td>how a total splits up</td><td>stacked bar, treemap, waffle (or, carefully, a doughnut)</td></tr>
<tr><td>Deviation</td><td>difference from a target or average</td><td>deviation (diverging) bar chart</td></tr>
<tr><td>Distribution</td><td>how values are spread</td><td>histogram, box plot</td></tr>
<tr><td>Correlation</td><td>how two measures move together</td><td>scatter plot</td></tr>
<tr><td>Spatial</td><td>where things are</td><td>map</td></tr></table>
<h2 class="dot">Line charts</h2>
<p>Use them for time series. Avoid lines that are too thick. Show dots for each point if the data was sampled at irregular intervals (monthly in one year, quarterly in the next). Aim for at most four coloured lines on one chart. For more, try <strong>small multiples</strong> (a grid of small charts on the same scale) or a <strong>slope chart</strong> (just the start and end points).</p>
<h2 class="dot">Bar charts</h2>
<p>Use them to compare amounts across categories or over time. Sort the bars sensibly. Aim for 4 to 8 grey gridlines at sensible intervals (10s, 1,000s). Make the gaps about half the width of a bar. Show an axis or data labels, probably not both. Avoid sideways text: switch to horizontal bars instead. And <strong>never truncate the axis</strong>. To focus on differences, use a deviation chart instead.</p>
${W('truncate')}
<h2 class="dot">Scatter plots</h2>
<p>Use them to compare two continuous measures. A trend line shows the correlation, but use it with caution: it's easy to suggest that correlation is causation. And the <strong>grain</strong> matters: is each point a visit, a member, or a site in a month?</p>
<h2 class="dot">Things to avoid</h2>
<ul class="plain"><li><strong>Never make a 3D chart.</strong> Perspective makes nearer bars and slices look bigger, so readers can't compare them.</li>
<li><strong>Be very wary of pie charts.</strong> People are poor at comparing angles: with similar slices, you can't tell whether Spain or Italy sold more. A sorted bar chart almost always works better.</li>
<li><strong>Doughnut, treemap and waffle charts</strong> show part-to-whole. A treemap copes with many categories, and a waffle (a 10 × 10 grid) makes percentages easy to read.</li>
<li><strong>Maps</strong> are for spatial questions. Shading regions by raw totals mostly shows where people live, so use rates instead.</li></ul>
<h2 class="dot">Numbers on their own</h2>
<p>Sometimes the best chart is no chart. One big number with a short explanation can say more: <em>"7,400 people receive venomous snake bites in the US every year. 5 of them die."</em> Or: <em>"90% of plastic waste has never been recycled."</em></p>`,
check:[
  {q:'You want to show each site\'s visits month by month across 2025, for all six sites. Which chart works best?', a:['One pie chart per month','Small multiples: six small line charts on the same scale','A 3D bar chart','One line chart with six bright colours'], c:1, why:'Six lines on one chart is too many. Small multiples on the same scale keep each site readable and comparable.'},
  {q:'Which chart suits the relationship between marketing spend and new members?', a:['Scatter plot','Pie chart','Treemap','Line chart'], c:0, why:'A scatter plot compares two continuous measures, one point per row.'},
  {q:'Why should you avoid pie charts for comparing similar values?', a:['They\'re hard to print','People are poor at comparing angles and areas','They can\'t show percentages','Excel doesn\'t support them'], c:1, why:'Slices of similar size look the same. Bars of similar length are much easier to compare.'},
  {q:'A bar chart\'s axis starts at 1,100 instead of 0. What happens?', a:['Differences look bigger than they are','The chart becomes more accurate','Nothing changes','The bars become 3D'], c:0, why:'A bar\'s length should be proportional to its value. Cutting off the bottom exaggerates the differences.'},
]},
colour:{ title:'Colour and accessibility', summary:'Use colour to connect and separate information, for everyone, including people with colour blindness.',
learn:`
<h2 class="dot">Using colour</h2>
<ul class="plain"><li><strong>Keep colour consistent</strong> to connect information: give each category the same colour in every chart.</li>
<li><strong>Use shade</strong> (light to dark) for ordered information, such as low to high intensity.</li>
<li><strong>Use hue</strong> (different colours) when there's no order, such as different sites.</li>
<li><strong>Think about what colours mean.</strong> Red often means bad or stop, and green good or go. A chart that shows growth in red confuses people.</li></ul>
<h2 class="dot">Colour blindness</h2>
<p>Around 5% of people are colour blind: about 8% of men and 0.5% of women. The most common types make red and green hard to tell apart:</p>
<ul class="plain"><li><strong>Protanomaly / protanopia</strong>: red weakness or blindness</li><li><strong>Deuteranomaly / deuteranopia</strong>: green weakness or blindness</li></ul>
<p>Switch between the views below to see how a red and green chart looks to someone with each type.</p>
${W('colourblind')}
<ul class="plain"><li>If in doubt, check that your chart still works in greyscale.</li><li>Shades of blue work best: blue colour blindness is extremely rare.</li><li>Don't rely on colour alone. Add labels, or use different line styles and markers.</li></ul>`,
check:[
  {q:'You\'re showing class intensity: Low, Medium and High. How should you colour it?', a:['Three unrelated bright colours','Light to dark shades of one colour','Red, green and blue','Random colours'], c:1, why:'The categories have an order, so shades of one colour (light to dark) match it.'},
  {q:'What\'s the safest way to check a chart works for colour-blind readers?', a:['Make it 3D','View it in greyscale','Use more colours','Use red and green'], c:1, why:'If the chart still makes sense in greyscale, it doesn\'t depend on telling colours apart.'},
  {q:'Which colours are hardest for the most colour-blind people to tell apart?', a:['Blue and yellow','Red and green','Black and white','Blue and orange'], c:1, why:'Red-green colour blindness (protan and deutan types) is by far the most common.'},
]},

/* ===================== Unit 9: Storage ===================== */
pipelines:{ title:'Pipelines and databases', summary:'ETL and ELT, OLTP and OLAP, and the ACID rules that keep transactions safe.',
learn:`
<h2 class="dot">ETL and ELT</h2>
<p>Data rarely stays where it's created. A <strong>pipeline</strong> moves it from source systems, such as Sparta Leisure's check-in gates, booking app and payment system, to somewhere it can be analysed.</p>
${FIG_ETL}
<p><strong>ETL</strong> (extract, transform, load) cleans and reshapes data on the way, so only tidy, structured data is stored. <strong>ELT</strong> (extract, load, transform) loads the raw data first and transforms it inside the storage system when it's needed. That has become practical as storage got cheaper and more powerful.</p>
<h2 class="dot">OLTP and OLAP</h2>
<table class="mini"><tr><th></th><th>OLTP: online transaction processing</th><th>OLAP: online analytical processing</th></tr>
<tr><td>Purpose</td><td>run the business: record every <strong>transaction</strong>, meaning any change to the database</td><td>understand the business: extract business intelligence from OLTP data</td></tr>
<tr><td>At Sparta Leisure</td><td>a member checks in, a payment is taken, a class is booked</td><td>visits by site by month, cancellation rates by plan</td></tr>
<tr><td>Design</td><td>highly <strong>normalised</strong>: each fact stored once</td><td>adds a layer of abstraction and aggregation. A semantic model describes what each element means, and data from several sources is combined across many dimensions in an <strong>OLAP cube</strong>.</td></tr>
<tr><td>Workload</td><td>heavy write, light read</td><td>heavy read, light write</td></tr></table>
<h2 class="dot">ACID</h2>
<p>Databases protect transactions with four guarantees:</p>
<table class="mini"><tr><th>Letter</th><th>Means</th></tr>
<tr><td><strong>Atomicity</strong></td><td>each transaction is a single unit. If any part fails, it all fails: a payment is never taken without the membership being created.</td></tr>
<tr><td><strong>Consistency</strong></td><td>each transaction leaves the database in a valid state, for example with every primary and foreign key relationship intact</td></tr>
<tr><td><strong>Isolation</strong></td><td>transactions running at the same time give the same result as if they'd run one after another, so two people can't book the last place in a class</td></tr>
<tr><td><strong>Durability</strong></td><td>once committed, a transaction survives a system failure, because it's written to non-volatile storage</td></tr></table>`,
check:[
  {q:'Sparta Leisure loads raw check-in logs into cheap storage and cleans them only when an analyst needs them. Is this ETL or ELT?', a:['ETL','ELT','OLTP','ACID'], c:1, why:'Load first, transform later: ELT.'},
  {q:'Which system records each check-in at the gate as it happens?', a:['OLAP','OLTP','A data mart','A slicer'], c:1, why:'Recording transactions as they happen is OLTP: heavy write, light read.'},
  {q:'A member pays for a membership, but the system crashes before the membership is created, and the payment is undone. Which ACID property is this?', a:['Atomicity','Isolation','Durability','Consistency'], c:0, why:'The payment and the membership are one transaction. If part of it fails, all of it fails.'},
]},
warehouses:{ title:'Data warehouses and dimensional modelling', summary:'Warehouses and marts, Inmon and Kimball, facts and dimensions, and star and snowflake schemas.',
learn:`
<p>A <strong>data warehouse</strong> is designed for analysis. It brings data together from many stores and supports analysis on the combined data. It only stores data that answers concrete business questions. Slices of a warehouse for one business area, such as sales or finance, are called <strong>data marts</strong>.</p>
<h2 class="dot">Two architectures</h2>
<table class="mini"><tr><th>Inmon: the enterprise data warehouse</th><th>Kimball: the dimensional data warehouse</th></tr>
<tr><td>Top-down: build one normalised warehouse first, then summarise and share it out into data marts. The data is non-volatile: never overwritten or deleted.</td><td>Bottom-up: build separate data marts first, each matching one department, then combine them. Quicker and cheaper to set up, but harder to maintain.</td></tr></table>
<h2 class="dot">Facts and dimensions</h2>
<table class="mini"><tr><th>Fact tables</th><th>Dimension tables</th></tr>
<tr><td>represent <strong>events</strong> in a business process, such as a visit or a booking. They hold <strong>measurements</strong>, often numbers that are easy to add up (DurationMins, Fee). Their <strong>grain</strong> is the level of detail: one row per visit, or one row per site per day.</td><td>represent <strong>people, things and places</strong>, such as a member, a site or a date. They hold descriptive data, often text, used for filtering and labelling (SiteName, PlanName, MonthName).</td></tr></table>
<h2 class="dot">Star and snowflake schemas</h2>
<p>In a <strong>star schema</strong>, one or more fact tables each join to several dimension tables through foreign keys. It's simple, widely used, effective for simple queries and supported by almost every BI tool, Power BI included. A <strong>snowflake schema</strong> normalises the dimensions further, for example splitting Plan out of Member. That saves disk space and reduces redundancy, but makes simple queries harder.</p>
${W('schema')}
<h2 class="dot">Activity</h2>
<p>Design a star schema, then a snowflake schema, for a video rental shop (think Blockbuster). Queries will use: the time and date, the customer's name, age and gender, the employee who served them, the product's name and type (DVD, Blu-ray, PS4…), the length of the loan, the return date and the amount paid. What is the fact table's grain?</p>
<p>Further reading: <a href="https://www.kimballgroup.com/data-warehouse-business-intelligence-resources/kimball-techniques/dimensional-modeling-techniques/" target="_blank" rel="noopener">Kimball's dimensional modelling techniques</a>, and the example data warehouses documented by <a href="https://dataedo.com" target="_blank" rel="noopener">Dataedo</a>.</p>`,
check:[
  {q:'In Sparta Leisure\'s star schema, which table is the fact table?', a:['Sites','Visits','Plans','Dates'], c:1, why:'Visits records events and holds a measurement, DurationMins. The others describe who, where and when.'},
  {q:'What makes a snowflake schema different from a star schema?', a:['It has no fact table','Its dimension tables are normalised into further tables','It only works in Excel','It stores unstructured data'], c:1, why:'A snowflake splits dimensions into more tables, such as Member → Plan, which reduces redundancy.'},
  {q:'A fact table has one row per member per day. What is this called?', a:['Its grain','Its schema','Its slicer','Its mart'], c:0, why:'The grain is the level of detail of one row.'},
  {q:'Which approach builds separate data marts first, then combines them?', a:['Inmon','Kimball','ELT','OLTP'], c:1, why:'Kimball\'s approach is bottom-up, starting with data marts.'},
]},
lakes:{ title:'Data lakes and lakehouses', summary:'Storing everything in its raw form, and the newer hybrid of lakes and warehouses.',
learn:`
<h2 class="dot">Data lakes</h2>
<p>Data lakes are newer, made possible by big data technology. They store structured and unstructured data in its <strong>original form</strong>: tables, logs, images, sensor readings. Data is processed when it's analysed, not when it's stored. This is called <strong>schema-on-read</strong>, as opposed to a warehouse's <strong>schema-on-write</strong>. Lakes are cheap compared with warehouses, and highly scalable and agile. Without care, though, they turn into "data swamps" that nobody can find anything in.</p>
<h2 class="dot">Data lakehouses</h2>
<p>A lakehouse is a very new technology that takes the best of both. It keeps a lake's flexible, cheap storage, but adds some of a warehouse's structure and control, including ACID transactions. It's flexible and scalable, and storage and processing can often be separated. Governance can be difficult.</p>
<h2 class="dot">Comparing them</h2>
<table class="mini"><tr><th></th><th>Database (OLTP)</th><th>Data warehouse</th><th>Data lake</th><th>Lakehouse</th></tr>
<tr><td>Main job</td><td>run day-to-day transactions</td><td>analysis and reporting</td><td>store everything for later</td><td>analysis and data science on one platform</td></tr>
<tr><td>Data</td><td>current, structured</td><td>cleaned, structured, historical</td><td>raw: structured and unstructured</td><td>raw and refined</td></tr>
<tr><td>Schema</td><td>on write, normalised</td><td>on write, dimensional</td><td>on read</td><td>on read, with tables on top</td></tr>
<tr><td>Users</td><td>apps and staff</td><td>analysts and managers</td><td>data scientists and engineers</td><td>all of these</td></tr>
<tr><td>Cost</td><td>moderate</td><td>higher</td><td>low</td><td>low to moderate</td></tr></table>`,
check:[
  {q:'Sparta Leisure wants to keep every raw gate log, app event and CCTV still, in case they\'re useful one day. Where should they go?', a:['An OLTP database','A data lake','A data mart','An Excel workbook'], c:1, why:'A lake stores raw, structured and unstructured data cheaply, to be processed when needed.'},
  {q:'What does schema-on-read mean?', a:['The structure is applied when the data is used','Data must be cleaned before it\'s stored','Only structured data is allowed','Reading data changes it'], c:0, why:'In a lake the data is stored as it is, and given structure when someone reads it.'},
  {q:'What does a lakehouse add to a data lake?', a:['Warehouse-style structure and ACID transactions','Paper records','Pie charts','Only unstructured data'], c:0, why:'A lakehouse adds tables, structure and ACID guarantees on top of cheap lake storage.'},
]},

capstone:{ kind:'capstone', title:'Capstone project', summary:'Clean, analyse and present the Sparta Leisure export, then make recommendations to the board.' },

/* ===================== Optional: Power Query ===================== */
'pq-setup':{ title:'Set up Excel and Power Query', summary:'Install Excel on your computer, check Power Query is there, and download the practice data.',
learn:`
<p><strong>Power Query</strong> is Excel's tool for getting data in and cleaning it. Instead of editing cells, you build a list of <strong>steps</strong> (import this file, trim these names, remove duplicates…). Excel records the steps and can repeat all of them on new data with one click. The same tool is built into Power BI.</p>
<p>This unit is optional, and it isn't needed for the capstone project. If you do it, though, you can use Power Query for the capstone's data. You'll need the <strong>desktop</strong> version of Excel, installed on your computer.</p>
<h2 class="dot">1. Install Excel</h2>
${STEPS([
 'Go to <a href="https://www.microsoft365.com" target="_blank" rel="noopener">microsoft365.com</a> and sign in with the work or college account your trainer has given you.',
 'Select <strong>Install apps</strong> (top right, or under <strong>Apps</strong>), then <strong>Microsoft 365 apps</strong>. Run the installer that downloads, and wait for it to finish. It installs Word, PowerPoint and the rest too.',
 'Open <strong>Excel</strong> from the Start menu (or Applications on a Mac), and sign in with the same account if it asks.',
 'If you can\'t install Office, or the account doesn\'t include it, ask your trainer before going any further.'])}
${TIP('<strong>Excel for the web</strong>, the version in a browser, can\'t build these queries. Neither can the Excel app on a phone or tablet. You need Excel installed on a Windows PC or a Mac.')}
<h2 class="dot">2. Check Power Query is there</h2>
${STEPS([
 'Open a blank workbook.',
 'Select the <strong>Data</strong> tab. On the left you should see a group called <strong>Get &amp; Transform Data</strong>, with a <strong>Get Data</strong> button.',
 'Select <strong>Get Data &gt; From File</strong>. You should see <strong>From Text/CSV</strong>, <strong>From JSON</strong>, <strong>From XML</strong> and <strong>From Folder</strong>.',
 'To see your version, go to <strong>File &gt; Account</strong>. It should say <strong>Microsoft 365</strong>.'])}
${TIP('These lessons are written for <strong>Excel for Windows</strong>. Excel for Mac (Microsoft 365) has Power Query too, but some options are missing or in different places. If a step isn\'t there on your Mac, work through it with someone using Windows.')}
<h2 class="dot">3. Download the practice data</h2>
<p>The practice data is Sparta Leisure's 2025 data in its raw form: a JSON file of members, one CSV of visits for each site, the plans as CSV, and the sites as XML. Some of it needs cleaning.</p>
${`<p class="task-dl"><a class="btn small" href="downloads/sparta-leisure-powerquery.zip" download data-testid="download-pq">Download the practice data (ZIP)</a></p>`}
${STEPS([
 'Save the ZIP somewhere you\'ll find it, such as your Documents folder.',
 '<strong>Unzip it</strong>: right-click the ZIP and choose <strong>Extract All</strong> (on a Mac, double-click it). Power Query can\'t read files that are still inside a ZIP.',
 'Open the <strong>sparta-leisure-powerquery</strong> folder and check it has <code>members.json</code>, <code>plans.csv</code>, <code>sites.xml</code> and a <code>visits</code> folder with six CSV files.',
 'Don\'t move or rename the files once you start. Your queries remember where each file is.'])}
<p>Open a new, blank Excel workbook and save it in the same folder, as <code>power-query-practice.xlsx</code>. You'll build every query in this unit in that one workbook.</p>`,
tasks:[
  'Excel is installed on my computer, and I\'m signed in',
  'The <strong>Data</strong> tab has <strong>Get Data</strong>, with From JSON, From XML and From Folder under <strong>From File</strong>',
  'I\'ve downloaded and unzipped the practice data, and the <code>visits</code> folder has six files',
  'I\'ve saved a blank workbook called <code>power-query-practice.xlsx</code> in the same folder',
]},
'pq-import':{ title:'Getting data in', summary:'Import JSON, CSV and XML files, set the data types, and read the Applied Steps.', noGrid:true,
learn:`
<p>Every query has the same shape: <strong>connect</strong> to a file, <strong>transform</strong> it in the Power Query Editor, then <strong>load</strong> the result into a sheet. Here's the Power Query Editor:</p>
<table class="mini"><tr><th>Part</th><th>What it's for</th></tr>
<tr><td><strong>Preview</strong> (middle)</td><td>the first rows of your data, as it looks after the steps so far</td></tr>
<tr><td><strong>Applied Steps</strong> (right)</td><td>every step, in order. Select a step to see the data at that point, and use the ✕ to delete one. Select the cog to change a step's settings.</td></tr>
<tr><td><strong>Data type icon</strong> (left of each heading)</td><td>ABC is text, 123 is a whole number, 1.2 is a decimal, a calendar is a date, and so on. Select it to change the type.</td></tr>
<tr><td><strong>Close &amp; Load</strong> (Home tab)</td><td>runs the steps and puts the result in a new sheet, as an Excel table</td></tr></table>
${TIP('To count rows, load the query, then open <strong>Data &gt; Queries &amp; Connections</strong>. The pane lists each query with "N rows loaded". The preview in the editor only shows the first rows of a big file.')}
<h2 class="dot">Data types matter</h2>
<p>Power Query guesses each column's type when you import a file (the <strong>Changed Type</strong> step). It reads dates using your computer's regional settings. Sparta Leisure's Walsall site exports dates as <code>dd/mm/yyyy</code>. On a computer set to US settings, 03/04/2025 is read as 4 March, and dates such as 25/03/2025 become errors. The fix is to set the type <strong>using locale</strong>, so you tell Power Query which country's date format the file uses.</p>
${`<p class="task-dl"><a class="btn small" href="downloads/sparta-leisure-powerquery.zip" download data-testid="download-pq">Download the practice data (ZIP)</a></p>`}
<p>Do the three tasks in order, in your <code>power-query-practice.xlsx</code> workbook.</p>`,
ex:[
  { p:'Import **members.json** and load it as a table. How many members does it load?', sheet:'Members', sol:'=COUNTA(Members!A:A)-1',
    steps:['<strong>Data &gt; Get Data &gt; From File &gt; From JSON</strong>, choose <code>members.json</code>, then <strong>Import</strong>.',
     'The editor shows a <strong>Record</strong> with two fields. Select the word <strong>List</strong> next to <code>members</code> to drill into it.',
     'Select <strong>To Table</strong> (List Tools &gt; Transform) and press <strong>OK</strong>.',
     'Select the expand icon (two arrows) in the <code>Column1</code> heading. Untick <strong>Use original column name as prefix</strong>, then <strong>OK</strong>.',
     'Set the types: select <code>dateOfBirth</code>, <code>joinDate</code> and <code>cancelDate</code> (hold Ctrl), then <strong>Transform &gt; Data Type &gt; Date</strong>.',
     'Rename the query <code>Members</code> (Query Settings, on the right), then <strong>Home &gt; Close &amp; Load</strong>.'],
    ask:'How many rows loaded? (See Data > Queries & Connections.)', hint:'JSON often arrives as a record holding a list. Drill into the list, convert it to a table, then expand the records into columns.',
    how:'Drill into the members list, To Table, expand Column1, set the date types, then Close & Load. The Queries & Connections pane shows the row count.' },
  { p:'Import the Walsall visits file, **visits/visits-S04.csv**, and make sure every check-in is read as a UK date and time. How many Walsall visits were in **March 2025**?', sheet:'Visits', sol:'=COUNTIFS(C:C,"S04",D:D,">="&DATE(2025,3,1),D:D,"<"&DATE(2025,4,1))',
    steps:['<strong>Data &gt; Get Data &gt; From File &gt; From Text/CSV</strong>, choose <code>visits-S04.csv</code>, then <strong>Transform Data</strong> (not Load).',
     'Look at the <strong>Changed Type</strong> step under Applied Steps, and at <code>CheckIn</code>. If its icon isn\'t a calendar with a clock, or some values show <em>Error</em>, the dates were misread.',
     'Right-click the <code>CheckIn</code> heading and choose <strong>Change Type &gt; Using Locale…</strong>. Set <strong>Data Type</strong> to <strong>Date/Time</strong> and <strong>Locale</strong> to <strong>English (United Kingdom)</strong>, then <strong>OK</strong>. If asked, choose <strong>Replace current</strong>.',
     'Select the arrow in the <code>CheckIn</code> heading and choose <strong>Date/Time Filters &gt; Between…</strong>. Keep rows that are after or equal to 01/03/2025 00:00 <strong>and</strong> before 01/04/2025 00:00.',
     'Rename the query <code>Walsall March</code>, then <strong>Close &amp; Load</strong>.'],
    ask:'How many rows loaded?', hint:'The type has to be set using locale English (United Kingdom), so 03/04/2025 means 3 April.',
    how:'Change Type > Using Locale (Date/Time, English (United Kingdom)), then filter CheckIn to on or after 1 March and before 1 April 2025.' },
  { p:'Import **sites.xml**, make `pool` a true/false column, and keep only the sites with a pool. What is their **total capacity**?', sheet:'Sites', sol:'=SUMIFS(E:E,D:D,TRUE)',
    steps:['<strong>Data &gt; Get Data &gt; From File &gt; From XML</strong>, choose <code>sites.xml</code>.',
     'In the Navigator, select <strong>site</strong>, then <strong>Transform Data</strong>. Each site\'s <code>id</code> attribute appears as a column called <code>Attribute:id</code>.',
     'Set <code>pool</code> to <strong>True/False</strong> and <code>capacity</code> to <strong>Whole Number</strong> (the type icon in each heading).',
     'Filter <code>pool</code> to <strong>TRUE</strong> only.',
     'Select the <code>capacity</code> column, then <strong>Transform &gt; Statistics &gt; Sum</strong>. The query now returns a single number.',
     'Rename the query <code>Pool capacity</code>, then <strong>Close &amp; Load</strong>. (Don\'t worry about losing the site list: the next lessons import it again.)'],
    ask:'What is the total capacity of the sites with a pool?', hint:'Change pool to True/False, filter it to TRUE, then use Transform > Statistics > Sum on capacity.',
    how:'From XML, choose site, set pool to True/False, filter to TRUE, then Statistics > Sum on capacity.' },
]},
'pq-clean':{ title:'Cleaning with steps', summary:'Trim, change case, replace values and filter, then refresh to repeat every step.', noGrid:true,
learn:`
<p>Power Query has a button for most of the cleaning techniques from <a href="#/lesson/cleaning">Cleaning data</a>. Each one adds a step instead of changing the data in place, so your original file is never touched.</p>
<table class="mini"><tr><th>Cleaning job</th><th>In Power Query</th><th>The formula you learnt</th></tr>
<tr><td>Remove extra spaces</td><td><strong>Transform &gt; Format &gt; Trim</strong></td><td><code>TRIM</code></td></tr>
<tr><td>Fix capital letters</td><td><strong>Format &gt; Capitalize Each Word</strong>, <strong>UPPERCASE</strong>, <strong>lowercase</strong></td><td><code>PROPER</code>, <code>UPPER</code>, <code>LOWER</code></td></tr>
<tr><td>Fix a typo or a code</td><td><strong>Transform &gt; Replace Values</strong></td><td><code>SUBSTITUTE</code></td></tr>
<tr><td>Split one column into two</td><td><strong>Transform &gt; Split Column</strong> (by delimiter or by number of characters)</td><td><code>LEFT</code>, <code>MID</code></td></tr>
<tr><td>Fill gaps from the row above</td><td><strong>Transform &gt; Fill &gt; Down</strong></td><td>copy and paste</td></tr>
<tr><td>Remove repeated rows</td><td><strong>Home &gt; Remove Rows &gt; Remove Duplicates</strong></td><td>Data &gt; Remove Duplicates</td></tr>
<tr><td>Deal with errors</td><td><strong>Home &gt; Remove Rows &gt; Remove Errors</strong>, or <strong>Transform &gt; Replace Errors</strong></td><td><code>IFERROR</code></td></tr>
<tr><td>Summarise</td><td><strong>Transform &gt; Group By</strong></td><td><code>COUNTIFS</code>, a pivot table</td></tr></table>
${TIP('Power Query is <strong>case-sensitive</strong>, unlike Excel formulas. To Power Query, <code>stu</code> and <code>STU</code> are different values, and <code>OLIVIA</code> isn\'t <code>Olivia</code>. So fix the case <em>before</em> you filter, group or merge.')}
<h2 class="dot">Refreshing</h2>
<p>This is the point of Power Query. When the data changes (next month's export, say), save the new file in the same place, then select <strong>Data &gt; Refresh All</strong>. Every step runs again on the new data. Cleaning that took an hour by hand now takes seconds, and it's done the same way every time.</p>
<p>All three tasks use your <strong>Members</strong> query from the last lesson. To change it, open <strong>Data &gt; Queries &amp; Connections</strong> and double-click <strong>Members</strong>. Do each task on a copy: right-click Members in the Queries pane on the left of the editor, choose <strong>Duplicate</strong>, and give the copy a name.</p>`,
ex:[
  { p:'Clean the names: trim spaces from `firstName` and `lastName`, and make them Capitalised Each Word. Then **how many members are called Olivia**?', sheet:'Members', sol:'=COUNTIF(Members!B:B,"Olivia")',
    steps:['Duplicate the <strong>Members</strong> query and name the copy <code>Olivias</code>.',
     'Select <code>firstName</code> and <code>lastName</code> (hold Ctrl), then <strong>Transform &gt; Format &gt; Trim</strong>.',
     'With both still selected, choose <strong>Format &gt; Capitalize Each Word</strong>.',
     'Filter <code>firstName</code> to <code>Olivia</code> only.',
     '<strong>Close &amp; Load</strong>, and read the row count in Queries &amp; Connections.'],
    ask:'How many members are called Olivia?', hint:'Some names are in CAPITALS, lower case, or have spaces round them. Without cleaning, the filter misses them.',
    how:'Trim, then Capitalize Each Word on both name columns, then filter firstName to Olivia.' },
  { p:'Some plan codes are in lower case. Make every `planCode` UPPERCASE, then use **Group By** to count the members on each plan. How many are on the **Student** plan (STU)?', sheet:'Members', sol:'=COUNTIF(Members!H:H,"STU")',
    steps:['Duplicate the <strong>Members</strong> query and name the copy <code>Members by plan</code>.',
     'Select <code>planCode</code>, then <strong>Transform &gt; Format &gt; UPPERCASE</strong>.',
     'Select <strong>Transform &gt; Group By</strong>. Group by <code>planCode</code>, with a new column called <code>Members</code> using the operation <strong>Count Rows</strong>. Then <strong>OK</strong>.',
     'You should see five rows, one per plan. If you see <code>stu</code> as well as <code>STU</code>, the UPPERCASE step is missing or in the wrong place.',
     '<strong>Close &amp; Load</strong>.'],
    ask:'How many members are on the STU plan?', hint:'Group By counts each distinct value. Fix the case first, or stu and STU are counted separately.',
    how:'UPPERCASE planCode, then Group By planCode with Count Rows, and read the STU row.' },
  { p:'Find the members who **cancelled during 2025**. How many are there?', sheet:'Members', sol:'=COUNTIFS(Members!K:K,">="&DATE(2025,1,1))',
    steps:['Duplicate the <strong>Members</strong> query and name the copy <code>Cancelled 2025</code>.',
     'Check <code>cancelDate</code> has the date type (a calendar icon). Active members have <code>null</code> there, which means empty.',
     'Select the arrow in the <code>cancelDate</code> heading, then <strong>Date Filters &gt; After or Equal To…</strong>, and enter 01/01/2025. Empty values are left out automatically.',
     '<strong>Close &amp; Load</strong>.'],
    ask:'How many members cancelled in 2025?', hint:'Use a date filter on cancelDate: on or after 1 January 2025. The data ends at 31 December 2025.',
    how:'Filter cancelDate to on or after 01/01/2025.' },
]},
'pq-combine':{ title:'Combining tables', summary:'Combine a folder of files, remove duplicates, and merge tables like a lookup.', noGrid:true,
learn:`
<p>Data often arrives in pieces: a file per site, or the plan prices in a different system from the members. Power Query puts the pieces together in two ways.</p>
<table class="mini"><tr><th></th><th>Append (and From Folder)</th><th>Merge</th></tr>
<tr><td>Does</td><td>stacks tables <strong>on top of each other</strong>: more rows</td><td>joins tables <strong>side by side</strong> on a matching column: more columns</td></tr>
<tr><td>Use it when</td><td>several files have the same columns, such as one visits file per site</td><td>you need a detail from another table, such as each member's plan fee</td></tr>
<tr><td>Like</td><td>copying one list under another</td><td><code>XLOOKUP</code>, for a whole table at once</td></tr></table>
<p><strong>From Folder</strong> appends every file in a folder in one go. Drop next month's file into the folder, select Refresh All, and it's included.</p>
<h2 class="dot">Join kinds</h2>
<p>When you merge, Power Query asks for a <strong>join kind</strong>. <strong>Left Outer</strong> keeps every row of the first table and brings in matches from the second, so it's the one that works like a lookup. <strong>Inner</strong> keeps only rows that match in both. Rows with no match are where data-quality problems hide, such as visits by member IDs that don't exist.</p>
${TIP('After a merge, the new column holds a little table in each row. Select the expand icon in its heading, and tick only the columns you need.')}`,
ex:[
  { p:'Combine all six files in the **visits** folder into one query, fix the dates, and **remove duplicate rows**. Part of Wolverhampton\'s file was exported twice. How many visits are left?', sheet:'Visits', sol:'=COUNTA(A:A)-1',
    steps:['<strong>Data &gt; Get Data &gt; From File &gt; From Folder</strong>, choose the <code>visits</code> folder, then <strong>Combine &gt; Combine &amp; Transform Data</strong>. Press <strong>OK</strong> on the Combine Files window.',
     'Power Query adds a <code>Source.Name</code> column saying which file each row came from. Right-click its heading and choose <strong>Remove</strong>.',
     'Walsall\'s dates are UK-style text. Right-click <code>CheckIn</code> and choose <strong>Change Type &gt; Using Locale…</strong>, with Date/Time and English (United Kingdom). The other sites\' dates (2025-03-01 13:05) are read correctly too.',
     'Select every column (click in the table, then '+KEYS('Ctrl','A')+'), then <strong>Home &gt; Remove Rows &gt; Remove Duplicates</strong>.',
     'Rename the query <code>Visits</code>, then <strong>Close &amp; Load</strong>.'],
    ask:'How many visits are left after removing duplicates?', hint:'Remove Duplicates only removes rows where every selected column matches. Select all the columns first.',
    how:'From Folder > Combine & Transform, remove Source.Name, set CheckIn with locale English (UK), select all columns, then Remove Duplicates.' },
  { p:'Merge your **Visits** with the sites from **sites.xml** to find each visit\'s `pool` value. How many **minutes** did members spend at the sites **with a pool**?', sheet:'Visits', sol:'=SUM(SUMIFS(E:E,C:C,{"S01","S02","S05"}))',
    steps:['Import <code>sites.xml</code> again (From XML, choose <strong>site</strong>, Transform Data). Set <code>pool</code> to True/False, and rename the query <code>Sites</code>. Use <strong>Home &gt; Close &amp; Load To… &gt; Only Create Connection</strong>, so it doesn\'t add a sheet.',
     'Open your <strong>Visits</strong> query. Select <strong>Home &gt; Merge Queries &gt; Merge Queries as New</strong>.',
     'Top table: <strong>Visits</strong>, select the <code>SiteID</code> column. Bottom table: <strong>Sites</strong>, select the <code>Attribute:id</code> column. Join Kind: <strong>Left Outer</strong>. Then <strong>OK</strong>.',
     'Expand the new <code>Sites</code> column and tick only <code>pool</code>.',
     'Filter <code>pool</code> to TRUE. Select <code>DurationMins</code>, then <strong>Transform &gt; Statistics &gt; Sum</strong>.',
     '<strong>Close &amp; Load</strong>.'],
    ask:'How many minutes were spent at the sites with a pool?', hint:'Merge on SiteID and Attribute:id with Left Outer, expand pool, filter to TRUE, then sum DurationMins.',
    how:'Merge Visits with Sites (SiteID = Attribute:id, Left Outer), expand pool, filter TRUE, then Statistics > Sum on DurationMins.' },
  { p:'Merge **Members** with **plans.csv** to get each member\'s monthly fee. What is the **total monthly fee of all Active members**, in pounds?', sheet:'Members', sol:'=SUMPRODUCT((Members!J2:J2621="Active")*SUMIF(Plans!A:A,Members!H2:H2621,Plans!C:C))',
    steps:['Import <code>plans.csv</code> (From Text/CSV, Transform Data). Check <code>MonthlyFee</code> is a decimal number, then <strong>Close &amp; Load To… &gt; Only Create Connection</strong>.',
     'Duplicate your <strong>Members</strong> query and name it <code>Income</code>. Make <code>planCode</code> UPPERCASE, as in the last lesson.',
     '<strong>Home &gt; Merge Queries</strong>: Income\'s <code>planCode</code> with plans\' <code>PlanCode</code>, Left Outer.',
     'Expand the new column and tick only <code>MonthlyFee</code>.',
     'Filter <code>status</code> to <code>Active</code>, select <code>MonthlyFee</code>, then <strong>Transform &gt; Statistics &gt; Sum</strong>.',
     '<strong>Close &amp; Load</strong>.'],
    ask:'What is the total monthly fee of all Active members? Type it to 2 decimal places.', hint:'Merging is case-sensitive: lower-case plan codes won\'t find a fee unless you make them UPPERCASE first.',
    how:'UPPERCASE planCode, merge with plans on PlanCode (Left Outer), expand MonthlyFee, filter status to Active, then Statistics > Sum.' },
]},
};

/* Spread the right answers of Check yourself questions across the options (a fixed rotation, so it's stable) */
{ let n=0; for (const L of Object.values(LESSONS)) for (const q of (L.check||[])){ const len=q.a.length, k=(n++*2+1)%len; q.a=q.a.map((_,i)=>q.a[(i-k+len)%len]); q.c=(q.c+k)%len; } }
const PATH = UNITS.flatMap(u=>u.lessons.map(id=>({id, unit:u})));
