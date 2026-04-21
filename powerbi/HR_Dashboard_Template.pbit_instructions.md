# Power BI Template Setup Instructions

## File: HR_People_Analytics.pbix

Since .pbix files require the Power BI Desktop application to generate,
this file contains the complete step-by-step setup to recreate the dashboard
in Power BI Desktop in under 60 minutes.

## Data Source Connection
1. Open Power BI Desktop
2. Get Data → Excel Workbook
3. Select: HR_People_Analytics_Data.xlsx
4. Load all 5 sheets: Employee Master, Payroll Data, Attendance Data, Recruitment Data, KPI Summary

## Data Model (Star Schema)
- Fact_Employee (Employee Master) — center fact table
- Fact_Payroll (Payroll Data)
- Fact_Attendance (Attendance Data)
- Dim_Recruitment (Recruitment Data)
- Dim_Date — create using DAX (see below)

## Create Date Table (DAX)
```
Dim_Date = 
ADDCOLUMNS(
    CALENDAR(DATE(2015,1,1), DATE(2024,12,31)),
    "Year",         YEAR([Date]),
    "Month",        MONTH([Date]),
    "MonthName",    FORMAT([Date], "MMM"),
    "Quarter",      "Q" & QUARTER([Date]),
    "YearMonth",    FORMAT([Date], "YYYY-MM"),
    "WeekDay",      WEEKDAY([Date]),
    "IsWeekend",    IF(WEEKDAY([Date]) IN {1,7}, TRUE, FALSE)
)
```

## Key DAX Measures

### Attrition Rate (Rolling 12M)
```
Attrition Rate = 
VAR TotalEmp = COUNTROWS(Fact_Employee)
VAR Attrited = CALCULATE(
    COUNTROWS(Fact_Employee),
    Fact_Employee[Attrition] = "Yes"
)
RETURN DIVIDE(Attrited, TotalEmp, 0)
```

### Average Tenure (Years)
```
Avg Tenure Years = 
AVERAGEX(
    Fact_Employee,
    DATEDIFF(Fact_Employee[Hire_Date], 
             IF(ISBLANK(Fact_Employee[Exit_Date]), TODAY(), Fact_Employee[Exit_Date]),
             DAY) / 365.25
)
```

### Cost Per Hire
```
Cost Per Hire = 
DIVIDE(
    SUM(Dim_Recruitment[Cost_Per_Hire_INR]),
    SUM(Dim_Recruitment[Hired]),
    0
)
```

### Rolling 12M Attrition
```
Rolling 12M Attrition = 
CALCULATE(
    [Attrition Rate],
    DATESINPERIOD(Dim_Date[Date], LASTDATE(Dim_Date[Date]), -12, MONTH)
)
```

### YoY Headcount Change
```
YoY Headcount Change = 
VAR CurrentYear = CALCULATE(COUNTROWS(Fact_Employee), YEAR(Fact_Employee[Hire_Date]) = YEAR(TODAY()))
VAR PriorYear = CALCULATE(COUNTROWS(Fact_Employee), YEAR(Fact_Employee[Hire_Date]) = YEAR(TODAY())-1)
RETURN DIVIDE(CurrentYear - PriorYear, PriorYear, 0)
```

## Dashboard Pages (5 pages)
1. Executive Overview — KPI cards, headcount bar, attrition bar
2. Attrition Analysis — trend line, exit reasons donut, tenure heatmap
3. Performance Management — rating dist, 9-box grid, dept comparison
4. Diversity & Inclusion — gender stacked bar, age pyramid, location map
5. Payroll Analytics — salary bands, monthly trend, cost-per-hire

## Advanced Features to Enable
- Row Level Security: Manager can only see their department
- Bookmarks: Toggle between chart types on Overview page
- Drill-Through: Click any dept → goes to dept detail page
- Tooltips: Custom tooltip page showing employee details on hover
- Field Parameters: Let user switch between metrics on the fly
