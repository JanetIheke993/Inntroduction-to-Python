
Excel Project

Project Title: Health Records – From Raw Data to Insights

Background

This project involved analyzing patient health records that included patient visits, diagnoses, doctors, gender, treatment, treatment costs and so on. The dataset was initially messy and required cleaning, organization, and transformation into meaningful insights. Using Microsoft Excel, I performed data cleaning, applied formulas, built pivot tables, and designed an interactive dashboard that helps visualize healthcare trends and patterns.

Objectives
Clean and organize messy healthcare data
Apply formatting and Excel functions for better readability and analysis
Use Logical, Statistical, Lookup, Math, and Date functions for calculations
Build PivotTables and PivotCharts to analyze key metrics
Present your findings clearly


Part 1: Data Cleaning
Removed duplicates:
Used Remove Duplicates from the Data tab based on PatientID to ensure unique patient records.
Converted VisitDate to proper date format:
Applied Text to Columns > Delimited > Date (DMY), then formatted as dd-mmm-yyyy.
Handled missing data:
Replaced blank Diagnosis with Unknown and blank Doctor with Unassigned using Find and Replace.
Removed unnecessary columns:
Deleted the empty Age Range column (to rebuild later with a formula).
Standardized text values and cleaned extra spaces:
Verified that Gender contained only Male, Female, Other.
Applied =TRIM() on columns Doctor, Treatment, and Diagnosis to remove extra spaces, then pasted values.


Part 2: Data Formatting
Styled headers: Applied bold text and background color to all column headers.
Formatted Cost column: Used Currency format with two decimal places.
Formatted dates: Applied custom format dd-mmm-yyyy to VisitDate.
Applied conditional formatting:
Highlighted patients older than 60 using “Greater Than” rule.
Highlighted treatments costing more than $3,000 using a custom formula =H2>3000.


Part 3: Formulas and Calculations
Created a new sheet named Formulas to perform analytical computations.

Logical Functions
Created a new column Age Group using:
Adult

Statistical Functions
Average Cost → =AVERAGE(Cleaned_Formatted!H:H)
Maximum Cost → =MAX(Cleaned_Formatted!H:H)
Minimum Cost → =MIN(Cleaned_Formatted!H:H)
Count of Diabetes patients → =COUNTIF(Cleaned_Formatted!E:E,"Diabetes")

Lookup and Reference
Created a Treatment Lookup Table for standard costs using:
I extracted the unique treatment types using the remove duplicates feature on Excel and got my four unique treatment types.
=AVERAGEIF(Cleaned_Formatted!F:F,J4,Cleaned_Formatted!H:H)
#N/A

Math and Trig
Cost with Tax: =Cost*1.1
Applied rounding functions:
=ROUND(P2,2)
=ROUNDUP(P2,2)

Date Functions
Extracted Year: =YEAR(A2)
Extracted Month: =TEXT(A2,"mmm") (to convert numeric month to text format like “Jan”)
Rebuilt Age Range using:
#N/A


Part 4: Pivot Tables & Charts
Created a new sheet titled Analysis to summarize insights.

1. Total Cost by Doctor:
Rows: Doctor
Values: Sum of Cost
Chart: Column chart titled Total Cost by Doctor

2. Gender Distribution of Patients:
Rows: Gender
Values: Count of Patients
Chart: Donut (Pie) chart titled Gender Distribution of Patients

3. Average Cost by Diagnosis:
Rows: Diagnosis
Values: Average of Cost
Chart: Horizontal Bar Chart titled Average Cost by Diagnosis

4. Total Cost by Diagnosis:
Rows: Diagnosis
Values: Sum of Cost
Chart: Column Chart titled Total Cost by Diagnosis

5. Monthly Trend of Cost:
Rows: Month
Values: Total Cost
Chart: Line Chart titled Monthly Trend of Cost

6. Age Group Distribution of Patients:
Rows: Age Group
Values: Count of Patients
Chart: 2D Pie Chart titled Age Group Distribution of Patients


Part 5: Dashboard Creation
Created a new sheet named Dashboard and designed an interactive, professional visualization.
Steps Taken:
Arranged all charts neatly in a grid layout (3 rows × 2 columns).
Added a unified dashboard title:
“Patient Data Insights Dashboard”
Applied consistent color themes, rounded chart borders, and dark background for visual appeal.
Ensured all charts were properly labeled with titles and legends.
Added Interactivity (Slicers)
Inserted slicers for key variables:
Doctor
Age Group
Gender
Month
Connected all slicers to multiple PivotTables so that they control the visuals dynamically.
Grouped slicers neatly on the right-hand side for user-friendly filtering.
Enabled users to filter the dashboard interactively e.g., view patient trends by doctor, gender, or month.

Final Touches
Ensured alignment and consistent chart sizing.
Applied rounded corners and subtle shadows to give a modern look.
Verified interactivity and responsiveness of all slicers and charts.


Key Insights from the Dashboard
The highest total cost was attributed to Dr. John’s patients.
Female patients represented a slightly higher share of total visits.
Adults made up the majority of patients, followed by Senior Adults.
Average treatment costs were highest for Cancer and Covid-19 diagnoses.
Monthly trend analysis showed spending fluctuations throughout the year.


Tools and Skills Demonstrated
Excel Tools: PivotTables, PivotCharts, Slicers, Conditional Formatting, Lookup and Date Functions
Skills: Data Cleaning, Analysis, Visualization, Dashboard Design, Data Storytelling
Outcome: A professional, interactive Excel Dashboard that transforms raw health records into actionable business insights.

