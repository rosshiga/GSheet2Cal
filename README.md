# GSheet2Cal
Google Apps Script and Sheet to automate calendar event entries for Gapps

Sheet template: [Google Drive](https://docs.google.com/spreadsheets/d/1j-f_kFUFQCTCEqkbFa93dTgOheXQCJPSrpDk2H-Fh-o/edit?usp=sharing)

1	Go to "File" and make a copy of this sheet for yourself, close the orginal

2	Wait up to 10 seconds for Event Adder menu to appear at top, if it does not appear refresh

3	If this is the first time, or you have added calendars, click "Event Adder" then "Update"

4	Go to Events tab and fill events as shown below, leave blanks for not needed

5	Go to "Event Adder" and "Add"

6	If event status do not up date, check Computed Event Times

7	Old events can be cleared or delete (Column A - J)


Formating: 

Time: Periods, colons, 24hr, and AM/PM supported (Cannot use 00:00)

Limitations:	

	Vaild only for dates in the future up to 9.9999 years
	
	Invaild times are interpreted midnight, so I blocked that set ttime to 00:01 for midnight events
