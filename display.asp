<%
option explicit
Response.Buffer = true 'performance reasons.
dim gbDebug
gbDebug = (Request.QueryString("debug") = "true")
'Response.Redirect "debug.asp?message=" & Request.Form("sqlstr")	'debugging message
%>
<html>
<head>
<title>Get database</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
td.db    {border:#006600 solid; border-width: 0px 1px;}
th.db    {border:#006600 solid; border-width: 1px; text-align:left;}
td.b     {border:#006600 solid; border-width: 0px 1px; background-color: #CCFFFF;}
table.db {border:#006600 solid 1px;}
span.n   {color : Gray; font : xx-small;}
-->
</style>
</head>
<body bgcolor="#ffffff" text="#000000">
<%
'check db conn details still in session.
if len(session("ConnStr"))<1 then
	Response.Write "<P><B>Session timed out, db connection lost.<B><BR>Go back and refresh last page to resend conection details</B></P>"
	if Request.Form.Count>0 then _
		Response.Write "<P>SQL String:<BR>" & Request.Form("sqlstr") & "</P>"
	Response.End
end if
%>
<p><font size="+2">Results</font></p>
<%
Response.Flush
'open db
dim odbConn
set odbConn = GetDBConn
dim oRs
set oRs = server.CreateObject("ADODB.Recordset")
'check requests
if Request.Form.Count>0 then 'sql submitted
	dim sSQLStr
	sSQLStr = Request.Form("sqlstr")
	Response.Write sSQLStr
	oRs.Open sSQLStr, odbConn, 3 'open / run query
	'leave recordset open for next part
elseif Request.QueryString("tbl") <> "" then 'table selected
	Response.Write "<P>Table: <B>" & Request.QueryString("tbl") & "</B>"
	oRs.Open "SELECT * FROM " & Request.QueryString("tbl"), odbConn,3
	'leave recordset open for next part
else 'nothing selected
	Response.Write "<P>Error: no table or sql string</P>"
	set oRs = nothing
	odbConn.Close
	set odbConn = nothing
	Response.End
end if

'parsing setting
dim bParse
bParse = (session("parse") = "true")
if bParse then
	Response.Write "<BR>Parsing on, html will be interpreted. This may cause incorrect display of data.</P>"
else
	Response.Write "<BR>Parsing off, html will be displayed.</P>"
end if

'display results
WriteRSs oRS

'close objects
set oRs = nothing
odbConn.Close
set odbConn = nothing
%>
</body>
</html>

<%
sub WriteRSs(oRS)
on error resume next
	do until oRS is nothing
		Response.Write "<hr>"
		if oRs.State = 0 then 'recordset closed
			Response.Write "<P>No Recordset Returned</P>"
		else
			Response.Flush 'send cached html
			WriteRS ors, true, bParse
		end if
		set ors = oRs.NextRecordset
		if err.number <>0 then
			Response.Write "<P color=#c00>Error: " & err.Description & "</p>"
			exit sub
		end if
		if not Response.IsClientConnected then exit sub
	loop
	Response.Write "<hr>"
end sub

sub WriteRS(oTableRS, bWriteRC, bParse)
	if bWriteRC then Response.Write "<P>" & oTableRS.RecordCount & " records</P>"
	%>
	<table border=1 cellspacing="0" class=db>
	<TR>
	<%
	dim field, lev, sHtml
	'write headings
	for each field in oTableRS.Fields
		Response.Write "<TH nowrap class=db>" & field.name & "</TH>" & vbcrlf
	next
	Response.Write "</TR>" & vbcrlf
	'write data
	if oTableRS.RecordCount > 0 then oTableRS.MoveFirst
	do while not oTableRS.EOF
		lev = not lev
		Response.Write "<TR class=db>" & vbcrlf
		for each field in oTableRS.Fields
			Response.Write "<TD nowrap"
			if lev then Response.Write " class=b" else Response.Write " class=db" 
			Response.Write ">"
			if isnull(field.value) then
				'write [null] for nulls
				Response.Write "<span class=""n""> [ null ]</span>"
			elseif trim(field.value) <> "" then
				'write non blank data
				if not bParse then 
					Response.Write server.HTMLEncode(field.value)
				else
					Response.Write field.value
				end if
			else
				'write blank data
				Response.Write "&nbsp;"
			end if
			Response.Write "</TD>" & vbcrlf
		next
		Response.Write "</TR>" & vbcrlf
		oTableRS.MoveNext
		Response.Flush
		if not Response.IsClientConnected then exit sub
	loop
	%>
	</table>
	<%
end sub 'WriteRS

function GetDBConn()
	dim odbConn, sUID, sPWord, sConnStr
	sUID = session("uid")
	sPWord = session("password")
	sConnStr = session("ConnStr")
	set odbConn = server.CreateObject("ADODB.Connection")
	if gbDebug then Response.Write "<P><B>Connection:</B> " & sConnStr & "</P>"
	odbConn.Open sConnStr, sUID, sPWord
	set GetDBConn = odbConn
	set odbConn = nothing
end function 'GetDBConn
%>
