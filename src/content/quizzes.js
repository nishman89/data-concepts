/* ---------- Quizzes ----------
   Two kinds of question:
     multiple choice  { unit, q, options:[4], answer:0-3, why }
     value            { unit, q, sheet, sol, why }  the learner types a formula (run on the sheet) or the value they
                                                    got in Excel. It's marked on the value the model formula gives. */
const QUIZ_SETTINGS = { passMark: 65 };
/* Options are written with the right answer anywhere; each question is rotated by a set amount so the right
   letters are spread across a, b, c and d. The rotation is fixed, so the order is the same every time. */
let Q_N = 0;
const Q = (unit, q, options, answer, why) => { const k=(Q_N++*3+1)%4; const o=options.map((_,i)=>options[(i-k+4)%4]); return { unit, q, options:o, answer:(answer+k)%4, why }; };
const V = (unit, q, sheet, sol, why) => ({ unit, q, sheet, sol, why });

const QUIZZES = [
{ id:'unit-2', unitQuiz:2, title:'Unit 2: Data types and file formats', blurb:'Types of data, dates, key-value pairs and file formats.', timeLimits:[10,15,0], defaultLimit:10, questions:[
  Q(2,'Which column is <strong>discrete</strong> data?',['<code>AvgTempC</code>','<code>MaxPlaces</code>','<code>Town</code>','<code>HasPool</code>'],1,'MaxPlaces is a count of whole places: discrete.'),
  Q(2,'<code>Attended</code> is TRUE or FALSE. What type is it?',['Categorical text','Continuous','Boolean','Datetime'],2,'Exactly two possible values: a Boolean.'),
  Q(2,'Why is <code>2025-03-04</code> safer than <code>03/04/2025</code> when you share data internationally?',['It\'s shorter','The day and the month can\'t be confused','Excel can\'t read the other format','It includes the time'],1,'In the US 03/04 is 4 March, and in the UK it\'s 3 April. ISO 8601 (year-month-day) is unambiguous.'),
  Q(2,'Which format uses tags such as <code>&lt;site id="S01"&gt;</code>?',['YAML','CSV','JSON','XML'],3,'XML describes structure with custom tags.'),
  Q(2,'A CSV file is…',['a binary file only Excel can open','plain text, with commas between values','a set of key-value pairs with indentation','a workbook with several sheets'],1,'Comma-separated values: plain text, one record per line.'),
  Q(2,'Which of these is NOT numeric data, even though it\'s written in digits?',['A member\'s phone number','A visit\'s length in minutes','A monthly fee','A site\'s capacity'],0,'You never do arithmetic on a phone number, so store it as text.'),
]},
{ id:'unit-3', unitQuiz:3, title:'Unit 3: Excel essentials', blurb:'Formulas, references, everyday functions, text, dates and logic.', timeLimits:[15,20,0], defaultLimit:15, questions:[
  Q(3,'You fill <code>=C2*$B$1</code> down to row 3. What does row 3 contain?',['<code>=C3*$B$2</code>','<code>=C3*$B$1</code>','<code>=C2*$B$1</code>','<code>=D2*$B$1</code>'],1,'C2 is relative, so it moves to C3. $B$1 is absolute, so it stays put.'),
  Q(3,'What does <code>=2+3*4</code> give?',['20','14','24','9'],1,'Multiplying comes first: 3*4 = 12, then + 2 = 14.'),
  Q(3,'Which function counts cells that are <strong>not empty</strong>, whatever they hold?',['COUNT','COUNTA','COUNTBLANK','SUM'],1,'COUNT only counts numbers. COUNTA counts anything that isn\'t empty.'),
  V(3,'How many members are on the <strong>Off-peak</strong> plan (OFF)?','Members','=COUNTIF(H:H,"OFF")','COUNTIF(H:H, "OFF"). You could also filter PlanCode in Excel and read the count.'),
  V(3,'What is the <strong>total capacity</strong> of all six sites?','Sites','=SUM(E:E)','SUM of the Capacity column, E.'),
  Q(3,'Which formula returns <strong>Long</strong> for visits over 90 minutes, and <strong>Normal</strong> otherwise?',['<code>=IF(E2>90,Long,Normal)</code>','<code>=IF(E2>90,"Long","Normal")</code>','<code>=IF("E2>90","Long","Normal")</code>','<code>=IF(E2>90;"Normal";"Long")</code>'],1,'Text results need double quotes, and the condition doesn\'t.'),
  Q(3,'<code>=DATEDIF(E2,DATE(2025,12,31),"Y")</code> gives…',['the number of days between the dates','the whole number of years between the dates','the year of E2','an error'],1,'With "Y", DATEDIF counts whole years.'),
  Q(3,'A cell shows <code>#DIV/0!</code>. What happened?',['A lookup failed','The formula divides by zero or an empty cell','A function name is misspelt','The column is too narrow'],1,'#DIV/0! means dividing by zero.'),
]},
{ id:'unit-4', unitQuiz:4, title:'Unit 4: Excel for analysis', blurb:'Lookups, conditional summaries, sorting, filtering and pivot tables.', timeLimits:[15,20,0], defaultLimit:15, questions:[
  Q(4,'In <code>=VLOOKUP(H2,Plans!$A$2:$C$6,3,FALSE)</code>, what does the 3 mean?',['Return the third row','Return the third column of the table','Look in column C of Members','Allow 3 errors'],1,'VLOOKUP returns from the column number you give, counting the table\'s first column as 1.'),
  Q(4,'Why does a VLOOKUP that works in row 2 show #N/A further down after being filled?',['VLOOKUP can\'t be filled','The table range wasn\'t absolute, so it slid down','FALSE should be TRUE','The data is sorted'],1,'Without $ signs the table moves as the formula is filled, and eventually misses the rows it needs.'),
  Q(4,'What can XLOOKUP do that VLOOKUP can\'t?',['Return a value from a column to the left of the one it searches','Match text','Find exact matches','Work on another sheet'],0,'XLOOKUP takes the search and return columns separately, so the return column can be anywhere.'),
  V(4,'How many visits were made at <strong>Kidderminster Wyre</strong> (S06)?','Visits','=COUNTIF(C:C,"S06")','COUNTIF(C:C, "S06"), or a pivot table with SiteID in Rows.'),
  V(4,'What was Ironbridge\'s total <strong>marketing spend for 2025</strong>, across all sites?','SiteMonths','=SUM(F:F)','SUM of MarketingSpend, column F.'),
  V(4,'How many <strong>Premium</strong> (PREM) members are at <strong>Telford Central</strong> (S01), whatever their status?','Members','=COUNTIFS(H:H,"PREM",G:G,"S01")','COUNTIFS with two pairs: PlanCode and HomeSiteID.'),
  Q(4,'You filter a list, then <code>=SUM(E:E)</code> still gives the total of every row. What adds up only the visible rows?',['=SUMIF(E:E)','=SUBTOTAL(9,E:E)','=VISIBLE(E:E)','=SUM(FILTER)'],1,'SUBTOTAL ignores rows hidden by a filter. Its 9 means sum.'),
  Q(4,'In a pivot table, you want visits counted for each site. Where does <code>SiteID</code> go?',['Values','Filters','Rows','Slicer only'],2,'The groups go in Rows (or Columns), and what you calculate goes in Values.'),
]},
{ id:'unit-5', unitQuiz:5, title:'Unit 5: Basic statistics', blurb:'Averages, spread, distributions, outliers and correlation.', timeLimits:[15,20,0], defaultLimit:15, questions:[
  Q(5,'Which average is least affected by outliers?',['Mean','Median','Range','Standard deviation'],1,'The median is the middle value, so a few extreme values barely move it.'),
  V(5,'What is the <strong>median</strong> class duration on the Classes sheet, in minutes?','Classes','=MEDIAN(D:D)','MEDIAN of DurationMins (column D).'),
  Q(5,'In a normal distribution, roughly what share of values lie within 2 standard deviations of the mean?',['50%','68%','95%','99.7%'],2,'68% within 1 SD, 95% within 2 and 99.7% within 3.'),
  Q(5,'A histogram has a long tail to the right. The distribution is…',['negatively skewed','positively skewed','normal','uniform'],1,'The tail points the way of the skew: right means positive.'),
  Q(5,'r = −0.9 means…',['no relationship','a strong negative relationship: as x rises, y falls','a weak positive relationship','x causes y'],1,'Close to −1 is a strong negative correlation. It says nothing about cause.'),
  V(5,'What is the <strong>standard deviation</strong> of NewMembers on SiteMonths? (Treat the year as a population.)','SiteMonths','=STDEV.P(C:C)','STDEV.P of column C. STDEV.S is also accepted if it\'s within rounding.'),
  Q(5,'R² = 0.65 for marketing spend and new members. What does it mean?',['Marketing causes 65% of sign-ups','65% of the variation in one can be explained by the other','65% of months had marketing','The correlation is 0.65'],1,'R² is the share of variation explained. It doesn\'t say which causes which, and r would be about 0.81.'),
]},
{ id:'unit-6', unitQuiz:6, title:'Unit 6: Data quality and cleaning', blurb:'Quality dimensions, cleaning techniques and missing values.', timeLimits:[10,15,0], defaultLimit:10, questions:[
  Q(6,'Weights recorded in a mix of kilograms and pounds fail which dimension?',['Uniformity','Redundancy','Completeness','Validity'],0,'Different units for the same measure: uniformity.'),
  Q(6,'Which formula turns <code>"  aLICE "</code> into <code>Alice</code>?',['=UPPER(TRIM(B2))','=PROPER(TRIM(B2))','=TRIM(B2)','=LEFT(B2,5)'],1,'TRIM removes the spaces, and PROPER makes it Title Case.'),
  V(6,'How many rows of <strong>MembersRaw</strong> share their MemberID with another row? (Count every row that has a repeat, including the first.)','MembersRaw','=SUMPRODUCT(--(COUNTIF(A2:A52,A2:A52)>1))','Flag duplicates with COUNTIF(A:A,A2)>1 and count the TRUEs, or use conditional formatting > Duplicate Values.'),
  Q(6,'Filling missing fees with the average fee is called…',['flagging','hot-deck','imputing','dropping'],2,'Working out a value from other observations is imputation.'),
  Q(6,'Why might dropping every row with a missing value be a bad idea?',['Excel can\'t delete rows','The remaining data can be biased if the gaps aren\'t random','It always makes the mean bigger','Missing values are never informative'],1,'If some kinds of people are more likely to have gaps, dropping them skews what\'s left.'),
  Q(6,'Normalisation, as defined in the course, rescales values to…',['0 to 100','a mean of 0','0 to 1','whole numbers'],2,'(x − min) ÷ (max − min) gives values between 0 and 1.'),
]},
{ id:'unit-7', unitQuiz:7, title:'Unit 7: Critical thinking', blurb:'The SIMPLE checklist, alternative explanations and evidence.', timeLimits:[10,15,0], defaultLimit:10, questions:[
  Q(7,'What does the L in SIMPLE stand for?',['Logic','Alternatives','Length','Labels'],1,'Identify possible aLternative explanations.'),
  Q(7,'What went wrong with New Coke?',['The taste tests were faked','People buy for more than taste, and the decision needed more and different data','Pepsi stopped selling','The sample was too small'],1,'200,000 taste tests measured only taste. Loyalty and habit mattered too.'),
  Q(7,'"Birmingham is better than Manchester" is…',['a fact','an opinion','an assertion backed by evidence','a statistic'],1,'"Better" is a judgement, so it\'s an opinion.'),
  Q(7,'Head injuries among cyclists halved after a helmet law. Which extra data would best test whether the law made cycling safer?',['Helmet sales','Injuries per mile cycled, before and after','The number of cycle shops','Weather reports'],1,'If fewer people cycled, injuries fall anyway. Rates per mile cycled control for that.'),
  Q(7,'The British fed false V-2 reports through double agents. Which question would have protected the Germans?',['Presentation','Source: is this data reliable?','Irrelevant','Alternatives'],1,'They trusted data without questioning its source.'),
]},
{ id:'unit-8', unitQuiz:8, title:'Unit 8: Data visualisation', blurb:'Good charts, the right chart for the relationship, and colour.', timeLimits:[10,15,0], defaultLimit:10, questions:[
  Q(8,'Which chart suits a <strong>time series</strong>?',['Pie chart','Line chart','Treemap','Map'],1,'Line charts show change over time.'),
  Q(8,'Which is a rule for bar charts?',['Always start the axis at the smallest value','Never truncate the axis','Use 3D for depth','Leave bars unsorted'],1,'A bar\'s length must be proportional to its value, so the axis starts at zero.'),
  Q(8,'What does a <strong>slope chart</strong> show?',['A distribution','Just the start and end values of several series','A correlation','Part-to-whole'],1,'A slope chart compares two points in time for each series.'),
  Q(8,'Which colours are safest for colour-blind readers?',['Red and green','Shades of blue','Rainbow colours','Pink and purple'],1,'Blue colour blindness is extremely rare.'),
  Q(8,'Andy Kirk\'s three principles are…',['fast, cheap, good','trustworthy, accessible, elegant','colourful, detailed, 3D','simple, small, sorted'],1,'Trustworthy, accessible and elegant.'),
]},
{ id:'unit-9', unitQuiz:9, title:'Unit 9: Data storage', blurb:'Pipelines, OLTP and OLAP, ACID, warehouses, lakes and lakehouses.', timeLimits:[10,15,0], defaultLimit:10, questions:[
  Q(9,'Which is heavy write, light read?',['OLAP','OLTP','A data warehouse','A data mart'],1,'OLTP records transactions as they happen.'),
  Q(9,'In ACID, Isolation means…',['each transaction is all or nothing','transactions running at the same time give the same result as running one after another','committed data survives failures','keys stay valid'],1,'That\'s isolation. All or nothing is atomicity.'),
  Q(9,'Dimension tables mostly hold…',['measurements to add up','descriptive data used for filtering and labelling','raw logs','backups'],1,'Facts hold measurements, and dimensions describe who, what, where and when.'),
  Q(9,'Which architecture builds the normalised warehouse first, then data marts?',['Kimball','Inmon','Snowflake','ELT'],1,'Inmon\'s approach is top-down.'),
  Q(9,'A data lake is usually…',['schema-on-write and expensive','schema-on-read and low cost','only for structured data','the same as OLTP'],1,'Lakes store raw data cheaply and apply structure when it\'s read.'),
]},
{ id:'final', title:'Final quiz', blurb:'Mixed questions across the whole path, including Excel questions you can answer here or in Excel.', timeLimits:[30,45,0], defaultLimit:30, questions:[
  Q(2,'<code>Category</code> on the Classes sheet (Cardio, Strength…) is…',['continuous','discrete','categorical','Boolean'],2,'Each value is a group with no order: categorical.'),
  Q(2,'Which file format is lightweight text built from key-value pairs and lists, and used between web apps and servers?',['DAT','JSON','XLSX','TXT'],1,'JSON.'),
  V(3,'What is the <strong>average monthly fee</strong> across the five plans, rounded to 2 decimal places?','Plans','=ROUND(AVERAGE(C:C),2)','ROUND(AVERAGE(C:C), 2).'),
  Q(3,'Which reference stays fixed when filled down, but moves when filled across?',['A1','$A$1','A$1','$A1'],2,'A$1 fixes the row only.'),
  V(3,'How many members were born <strong>before 1 January 1960</strong>?','Members','=COUNTIF(E:E,"<"&DATE(1960,1,1))','COUNTIF(DateOfBirth, "<"&DATE(1960,1,1)).'),
  V(4,'How many <strong>bookings</strong> were made for <strong>Spin</strong> (C01)?','Bookings','=COUNTIF(C:C,"C01")','COUNTIF of ClassID.'),
  V(4,'What is the <strong>average visit length</strong> at Wolverhampton City (S03), to 1 decimal place?','Visits','=ROUND(AVERAGEIFS(E:E,C:C,"S03"),1)','AVERAGEIFS(E:E, C:C, "S03"), rounded.'),
  Q(4,'Which formula finds the name of site S04?',['=VLOOKUP("S04",Sites!A:B,2,FALSE)','=VLOOKUP("S04",Sites!B:A,1,FALSE)','=HLOOKUP("S04",Sites!A:B,2,FALSE)','=MATCH("S04",Sites!B:B,0)'],0,'Search Sites column A and return column 2, with an exact match.'),
  Q(4,'A slicer…',['deletes rows','filters pivot tables with buttons that show what\'s selected','sorts a column','creates a chart'],1,'Slicers are visual filters for pivot tables and charts.'),
  V(5,'What is the <strong>mean</strong> visit length after removing visits over 300 minutes, to 1 decimal place?','Visits','=ROUND(AVERAGEIF(E:E,"<=300"),1)','AVERAGEIF(E:E, "<=300").'),
  Q(5,'Mean > median > mode usually suggests…',['negative skew','positive skew','a normal distribution','no outliers'],1,'A long right tail pulls the mean above the median.'),
  Q(5,'Which is the safest summary for a strongly skewed variable?',['Mean and SD','Median and interquartile range','Mode only','Maximum'],1,'The median and IQR aren\'t pulled about by the tail.'),
  Q(6,'Two rows hold the same member with different dates of birth. Which dimension fails?',['Consistency','Uniformity','Redundancy only','Completeness'],0,'Contradictory information about one individual is a consistency problem.'),
  Q(6,'Copying a missing value from a similar record is called…',['imputing the mean','hot-deck','flagging','scaling'],1,'Hot-deck takes the value from a similar record.'),
  Q(7,'"January sign-ups were four times December\'s, so the campaign worked." The strongest challenge is…',['the chart colours','January is always busier, so compare January with January','the source is the marketing team, so it must be false','December is an outlier that should be deleted'],1,'Seasonality is the obvious alternative explanation. Compare like with like.'),
  Q(8,'You need to show how Ironbridge\'s members split across five plans. Which is the better chart?',['3D pie chart','Sorted bar chart','Line chart','Scatter plot'],1,'A sorted bar chart compares parts easily. Avoid 3D, and be wary of pies.'),
  Q(8,'Around how many men are colour blind?',['1 in 200','1 in 12','1 in 3','1 in 1,000'],1,'About 8% of men: roughly 1 in 12.'),
  Q(9,'Which schema normalises its dimension tables?',['Star','Snowflake','Flat','Lake'],1,'A snowflake schema splits dimensions into further tables.'),
  Q(9,'Loading raw data first and transforming it later is…',['ETL','ELT','OLAP','ACID'],1,'Extract, load, then transform.'),
  Q(9,'A lakehouse adds which warehouse feature to a lake?',['ACID transactions and structure','Paper archives','Pie charts','OLTP gates'],0,'Lakehouses bring tables and ACID guarantees to lake storage.'),
]},
];
