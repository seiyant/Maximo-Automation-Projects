**Seiya Nozawa-Temchenko** | @seiyant

# biweeklyHoursPlanner.py

**This script intends to gather work orders for the next 2 weeks to find the total work work hours.**

Planned work orders are on Maximo, with their respective predicted work hours.
This script lists all planned work orders to generate a figure of expected time that must be spent over the course of those 2 weeks.

# maintenanceLogReader.py

**This script intends to check if a daily maintenance log status matches the Maximo status.**

Updated work orders are logged onto a daily maintenance sheet in XLSM format by the end of each day.
This script checks whether each status logged on the daily maintenance sheet matches the status on Maximo and produces an updated summary in XLSX format.

# maintenanceLogReader_oldver.py

**This script intends to check if a daily maintenance log status matches the Maximo status.**

All work orders are logged onto a daily maintenance sheet in DOCX format by the end of each day.
This script checks whether each status logged on the daily maintenance sheet matches the status on Maximo and produces an updated summary in XLSX format.

# plannedHoursCorrection.py

**This script intends to check work order planned hours and compare them to actual hours.**

Planned work orders are on Maximo, with their respective predicted work hours.
This script looks at past work orders from a specific time frame and analyzes the mean and standard deviation of real work hours to predicted work hours.

# plannedHoursCorrectionEntry.py

**This script intends to update old work order planned hours using historic actual hours.**

Planned work orders are on Maximo, with their respective predicted work hours.
This script reads the data plannedHoursCorrection.py scraped from Maximo and uses it to create new predicted hours for work orders and updates them in Maximo.

# workOrderAutomation.py

**This script intends to shorten the paper log entry time on Maximo.**

Work orders are written on paper by contractors, and the entry process into Maximo is a repetitive process with a lot of waiting for entries to load and update. 
This script quickly transfers entered information in the process required based on the work order type.
