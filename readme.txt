<!--
Pie is the definitive general purpose ASP database tool.
(c)2000-2003
by Tim Abell
code@timwise.co.uk
Please distibute this file with the code
-->


Pie is something I've developed in my own time. If you like it and use it please send me a tenner.
If you want features & stuff maybe I'll put them in for you.

Even if you're not feeling rich, please drop me a line and let me know if you find this useful.


By the way, you will need:
IIS4 or 5 (Comes with NT4 Option pack & Win2k as standard)
And you may need to load MSDAC2.7 & Jet4SP3
(can be found at http://msdn.microsoft.com/library/default.asp?url=/downloads/list/dataaccess.asp )

Cheers.

Tim
code@timwise.co.uk
please keep this address private. - I don't like spam much.

//--------------------------------------//
Known bugs

13-08-2003
user input is not at all validated
it would be a VERY bad idea to run this code on a public server.
Status: boring

11-08-2003
dedupe.asp
when displaying results, some matches will not be picked up in the display routine due to differences that the sql statement doesn't pick up.
Status: minor inconvenience

//--------------------------------------//
Feature Requests

13-08-2003 from Me
dedupe.asp should be able to mark duplicates with the id of the record they were matched against

13-08-2003 from Me
would like to be able to produce standard queries against standard databases without the user needing to know anything

//--------------------------------------//
Change log

12-12-03
Fixed bug:
Multi recordset display was stopping if first recordset was blank.

13-08-03
modified sql server connection string to include current user details.
(shows up in enterprise manager under current info / process info)

11-08-03
dedupe.asp, cols that were highlighted with a red col heading now have a column background color.

05-08-03
Added decriptions to data types in lengths page.
removed n/a... from range output (not needed now displaying field types)

04-08-03
Already added data display pages
Added field type to lengths display.

14-04-03
Changed driver from Access to JET spec. (again?)
Added Deduplification tool to extras.

25-04-03
Modified field name interpretation in lengths.asp to cope with access keeping "" around alias field names.

30-04-03
added &debug=true stuff to lengths.asp to spit out read field type/name array & sql string.