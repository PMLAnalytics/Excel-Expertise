# Excel-Expertise

Attendance Database

Attendance Database Formulas:
=VLOOKUP(A2,$I$2:$U$5204,6,FALSE)
=IFS(F2="IN"," ",F2="IN – Telework/Telecommuting","Telework/Telecommuting",F2="IN - Training","Training",F2="IN - SpecialProject","SpecialProject",F2="IN - Meeting","Meeting",F2="OUT - Flex","Flex",F2="OUT - Early","Early",F2="IN - Off Board","Off Board",F2="OUT - CTY Time","CTY Time",F2="OUT","No reason given",F2="OUT - Pre Approved","Pre-Approved",F2="OUT - Vacation","Vacation",F2="OUT - Late","Late",F2="OUT - LOA","LOA",F2="OUT - Outstation","Outstation")

End
