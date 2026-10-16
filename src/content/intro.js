/* ---------- How this site works, and Meet Sparta Leisure ---------- */
const SITE_INTRO = `
<p class="lede">This site teaches the core ideas of working with data, and gives you somewhere to practise them on a realistic dataset.</p>
<h2 class="dot">The learning path</h2>
<p>The path on the home page runs from data types to data storage, with a capstone project at the end. Work through it in order: each unit builds on the ones before. A lesson is complete when you've solved its exercises or answered its questions. If you already know a topic, you can also select <strong>Mark lesson as complete</strong>.</p>
<table class="mini"><tr><th>Lessons with…</th><th>Work like this</th></tr>
<tr><td>Exercises</td><td>Read the <strong>Learn</strong> tab, then do each <strong>Exercise</strong>. They're checked automatically.</td></tr>
<tr><td>Check yourself questions</td><td>Read the lesson, then answer the questions at the end. Explore the interactive charts as you go.</td></tr></table>
<h2 class="dot">Two ways to practise Excel</h2>
<p>The Excel exercises can be done in two ways. Switch between them on any exercise, and use whichever suits you.</p>
<table class="mini"><tr><th>Practise here</th><th>In Excel</th></tr>
<tr><td>Type a formula in the grid beside the lesson. It runs in your browser on the same data, and is checked straight away. Nothing to install, and it works on a phone.</td><td>Download the Sparta Leisure workbook, do the task in Excel, and type the answer you got into the box. This is how you'll work for real, and it's the only way to practise sorting, filtering, pivot tables and charts.</td></tr></table>
${TIP('The practice grid understands the Excel functions in this course, including XLOOKUP, COUNTIFS and DATEDIF. A few things only exist in Excel itself: pivot tables, charts, and functions that "spill" several results, such as FILTER and UNIQUE.')}
<h2 class="dot">Quizzes and the capstone</h2>
<p>Each unit has a short <strong>end-of-unit quiz</strong>, and there's a <strong>final quiz</strong> across the whole path. The pass mark is 65%. The <strong>capstone project</strong> gives you a raw, messy export to clean and analyse in Excel (Power BI is optional), and asks you to make recommendations to the board.</p>
<h2 class="dot">Your progress</h2>
<p>Progress is saved in this browser only, on this device. Nothing is sent anywhere. Clearing your browser data, or switching to another browser, starts you afresh.</p>`;

const DATA_INTRO = `
<p class="lede">Sparta Leisure is a made-up chain of six leisure centres in Shropshire, the Black Country and Worcestershire. Its data is used in every Excel exercise, the statistics lessons and the capstone project.</p>
<p>Members pick a plan, belong to a home site, check in at the gym and book classes. Three of the six sites have a pool. The data covers everyone who joined up to the end of 2025, and every visit and class booking during 2025.</p>
<table class="mini"><tr><th>Sheet</th><th>One row per…</th><th>Links to</th></tr>
<tr><td><strong>Sites</strong></td><td>leisure centre</td><td></td></tr>
<tr><td><strong>Plans</strong></td><td>membership plan, with its fees and what it includes</td><td></td></tr>
<tr><td><strong>PriceHistory</strong></td><td>plan, with its monthly fee for each year across the top</td><td>Plans</td></tr>
<tr><td><strong>Members</strong></td><td>member</td><td>Sites (HomeSiteID), Plans (PlanCode)</td></tr>
<tr><td><strong>Classes</strong></td><td>type of class</td><td></td></tr>
<tr><td><strong>Bookings</strong></td><td>class booking in 2025</td><td>Members, Classes, Sites</td></tr>
<tr><td><strong>Visits</strong></td><td>gym check-in in 2025</td><td>Members, Sites</td></tr>
<tr><td><strong>SiteMonths</strong></td><td>site per month of 2025: new members, cancellations, visits, marketing spend and temperature</td><td>Sites</td></tr>
<tr><td><strong>MembersRaw</strong></td><td>row of a messy extract from an old system, for the data cleaning lessons</td><td></td></tr></table>
<p>The sheets link through shared codes, like tables in a database: <code>Visits.MemberID</code> matches <code>Members.MemberID</code>, which tells you who made each visit. In Excel, you follow these links with lookups such as <code>XLOOKUP</code>.</p>
<p>${DL_WORKBOOK}</p>`;

/* Positions and columns for the Sparta Leisure diagram (PK = primary key, FK = foreign key → table) */
const DATA_ERD = {
  tables: [
    { name:'Plans', x:20, y:20, cols:[['PlanCode','PK'],['PlanName'],['MonthlyFee'],['JoiningFee'],['PeakAccess'],['ClassesPerMonth'],['PoolAccess'],['MinAge']] },
    { name:'PriceHistory', x:20, y:280, cols:[['PlanCode','FK','Plans'],['2022'],['2023'],['2024'],['2025']] },
    { name:'Members', x:250, y:20, cols:[['MemberID','PK'],['FirstName'],['LastName'],['Gender'],['DateOfBirth'],['Postcode'],['HomeSiteID','FK','Sites'],['PlanCode','FK','Plans'],['JoinDate'],['Status'],['CancelDate']] },
    { name:'Visits', x:490, y:20, cols:[['VisitID','PK'],['MemberID','FK','Members'],['SiteID','FK','Sites'],['CheckIn'],['DurationMins']] },
    { name:'Bookings', x:490, y:190, cols:[['BookingID','PK'],['MemberID','FK','Members'],['ClassID','FK','Classes'],['SiteID','FK','Sites'],['ClassDate'],['StartTime'],['Attended']] },
    { name:'Sites', x:250, y:330, cols:[['SiteID','PK'],['SiteName'],['Town'],['HasPool'],['Capacity'],['Opened'],['Latitude'],['Longitude']] },
    { name:'Classes', x:730, y:190, cols:[['ClassID','PK'],['ClassName'],['Category'],['DurationMins'],['Intensity'],['MaxPlaces']] },
    { name:'SiteMonths', x:490, y:400, cols:[['SiteID','FK','Sites'],['Month'],['NewMembers'],['Cancellations'],['Visits'],['MarketingSpend'],['AvgTempC']] },
  ],
};
/* The star and snowflake schemas in the Data warehouses lesson */
const SCHEMA_STAR = { tables:[
  { name:'FactVisit', x:300, y:110, fact:true, cols:[['VisitKey','PK'],['DateKey','FK','DimDate'],['MemberKey','FK','DimMember'],['SiteKey','FK','DimSite'],['TimeKey','FK','DimTime'],['DurationMins'],['VisitCount']] },
  { name:'DimDate', x:30, y:20, cols:[['DateKey','PK'],['Date'],['DayName'],['MonthName'],['Quarter'],['Year'],['IsWeekend']] },
  { name:'DimMember', x:570, y:20, cols:[['MemberKey','PK'],['MemberID'],['AgeBand'],['Gender'],['PlanName'],['MonthlyFee'],['JoinYear']] },
  { name:'DimSite', x:30, y:240, cols:[['SiteKey','PK'],['SiteName'],['Town'],['HasPool'],['Capacity']] },
  { name:'DimTime', x:570, y:240, cols:[['TimeKey','PK'],['Hour'],['Session'],['IsPeak']] },
]};
const SCHEMA_SNOW = { tables:[
  { name:'FactVisit', x:300, y:110, fact:true, cols:[['VisitKey','PK'],['DateKey','FK','DimDate'],['MemberKey','FK','DimMember'],['SiteKey','FK','DimSite'],['TimeKey','FK','DimTime'],['DurationMins'],['VisitCount']] },
  { name:'DimDate', x:30, y:20, cols:[['DateKey','PK'],['Date'],['DayName'],['MonthKey','FK','DimMonth']] },
  { name:'DimMonth', x:30, y:190, cols:[['MonthKey','PK'],['MonthName'],['Quarter'],['Year']] },
  { name:'DimMember', x:570, y:20, cols:[['MemberKey','PK'],['MemberID'],['AgeBand'],['Gender'],['PlanKey','FK','DimPlan'],['JoinYear']] },
  { name:'DimPlan', x:800, y:20, cols:[['PlanKey','PK'],['PlanName'],['MonthlyFee'],['PeakAccess']] },
  { name:'DimSite', x:300, y:330, cols:[['SiteKey','PK'],['SiteName'],['TownKey','FK','DimTown'],['HasPool'],['Capacity']] },
  { name:'DimTown', x:30, y:350, cols:[['TownKey','PK'],['Town'],['County']] },
  { name:'DimTime', x:570, y:250, cols:[['TimeKey','PK'],['Hour'],['Session'],['IsPeak']] },
]};
