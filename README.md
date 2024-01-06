# Excel-Expertise
Automatic Rotation; Client-to-Employee Tracking; Daily Analytics

CIO Log Formulas:
=VLOOKUP(A2, 'WMS Attendance'!B:G, 2, FALSE)
=IF(ISBLANK(VLOOKUP(A2, 'WMS Attendance'!B:G, 5, FALSE)), "", VLOOKUP(A2, 'WMS Attendance'!B:G, 5, FALSE))
=SUMIF('Daily CIO'!D:D, "English", 'Daily CIO'!H:H)
=SUM(SUMIF('Daily CIO'!D:D,"English/Viet", 'Daily CIO'!H:H), SUMIF('Daily CIO'!D:D, "English/Farsi",'Daily CIO'!H:H)*1)
=SUMPRODUCT((ISNUMBER(FIND("*", 'Daily CIO'!I:M)))*1)
=SUBTOTAL(109,['# of CIO''s])
=COUNTIFS('pending data table'!A1:E206,">="&J14, 'pending data table'!A1:E206, "<"&K14)

There are several other basic formulas in use, as well as many common functions.

It has 4 visible tabs and 3 hidden ones. The Daily CIO tab, which is the primary interface. The CIO Unit, which is a pilot program for the CIO Desk tbd. WMS Attendance functions as the backend. Analytics is a visualization tab that provides metrics for each individual day within the office. The 3 hidden tabs are: a blank sheet containing VBA Macro code for automatic rotation (used previously, available for further implementation), a How-To Guide (since re-implemented), and a backend of calculated data tables.
