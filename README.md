# Excel-Expertise
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
## Description
Front-end user interface combined with a [dynamic] backend that allows for tracking. Analytics tab built-in in order to report daily metrics.
CIO Log is sorted & filtered with 10 different filters and contains over a dozen conditional formatting rules. The data faces extraction, transformation, and loading from when the source is delivered to when the data warehouse can be interacted upon.
 - Automatic Rotation
 - Client-to-Employee Tracking
 - Daily Analytics

### Data Pipeline
- Backend: Workload Management System (Javascript)
- Staging Area: Formulas, Conditional Formatting Rules, Custom Sorting
- Frontend: Tracking interface

### CIO Log Formulas:
- =VLOOKUP(A2, 'WMS Attendance'!B:G, 2, FALSE)
- =IF(ISBLANK(VLOOKUP(A2, 'WMS Attendance'!B:G, 5, FALSE)), "", VLOOKUP(A2, 'WMS Attendance'!B:G, 5, FALSE))
- =SUMIF('Daily CIO'!D:D, "English", 'Daily CIO'!H:H)
- =SUM(SUMIF('Daily CIO'!D:D,"English/Viet", 'Daily CIO'!H:H), SUMIF('Daily CIO'!D:D, "English/Farsi",'Daily CIO'!H:H)*1)
- =SUMPRODUCT((ISNUMBER(FIND("*", 'Daily CIO'!I:M)))*1)
- =SUBTOTAL(109,['# of CIO''s])
- =COUNTIFS('pending data table'!A1:E206,">="&J14, 'pending data table'!A1:E206, "<"&K14)

There are several other basic formulas in use, as well as many common functions.

### 4 Visible tabs
- The Daily CIO tab, which is the primary interface
- The CIO Unit, which is a pilot program for the CIO Desk tbd
- WMS Attendance functions as the backend
- Analytics is a visualization tab that provides metrics for each individual day within the office
  
### 3 Hidden tabs
- A blank sheet containing VBA Macro code for automatic rotation (used previously, available for further implementation)
- A How-To Guide (since re-implemented)
- A backend of calculated data tables
