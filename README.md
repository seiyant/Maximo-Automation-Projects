**Seiya Nozawa-Temchenko** | @seiyant

**maintenanceLogReader.py**

This script intends to check if a daily maintenance log status matches the Maximo status.

All work orders are logged onto a daily maintenance sheet in PDF or DOCX formats by the end of each day.
This script checks whether each status logged on the daily maintenance sheet matches the status on Maximo.

**workOrderAutomation.py**

This script intends to shorten the paper log entry time on Maximo.

Work orders are written on paper by contractors, and the entry process into Maximo is a repetitive process with a lot of waiting for entries to load and update. 
This script manages to quickly transfer entered information in the process required based on the work order type.
