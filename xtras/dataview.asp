<%option explicit
Response.Buffer = true
'globals
dim gbDebug, godbConn, gsTable, iFirstRecords
gbDebug = (Request.QueryString("debug") = "true")
gsTable = Request.QueryString("tbl")
iFirstRecords = cint(Request.QueryString("first"))
Server.ScriptTimeout = 300
htmltop
OpenDB
main
htmlend
KillDB

sub htmltop
%>
<html>
<head>
<title>DB Xtras</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.b {  background-color: #CCFFFF;}
td,th { border:#006600 solid; border-width: 1px 1px 0px 1px}
th {text-align:left;}
table {border:#006600 solid; border-width: 1px 1px 2px 1px}
.n {
	color : Gray;
	font : xx-small;
}
.bh {  background-color: #CCFFFF; border-width: 1px 1px 0px 1px}
-->
</style>
</head>

<body bgcolor="#ffffff" text="#000000">
<%
end sub 'htmltop

sub OpenDB
	'check db conn still valid
	if len(session("ConnStr"))<1 then
		Response.Write "<P><B>Session timed out, db connection lost.<B><BR>Go back and refresh last page to resend conection details</B></P>"
		Response.End
	end if
	set godbConn = GetDBConn
end sub

sub KillDB
	godbConn.Close
	set godbConn = nothing
end sub

sub main
	dim oDataRS, sDataSQL, bAccess, aTxtFields, iFieldIndex
	set oDataRS = server.CreateObject("ADODB.Recordset")
	%><p><font size="+2">Data Cruncher - Analysis</font></p><%
	'open db
	'prepare
	Response.Write session("db")
	bAccess = (mid(session("db"),4,6) = "Access")
	Response.Write "<P>Table: " & gsTable & "</P>"
	'get record count
	sDataSQL = "SELECT Count(*) AS RC FROM [" & gsTable & "]"
	if gbDebug then Response.Write "<P>" & sDataSQL & "</P>" 'debugging
	oDataRS.Open sDataSQL, godbConn, 3
	Response.Write "<P>" & oDataRS("RC") & " Records<BR>"
	oDataRS.Close
	
	'get fields names, types, ranges, population
	aTxtFields = GetFieldList
	Response.Write "<table border=1><tr>"
	for iFieldIndex = 0 to ubound(aTxtFields)
		select case aTxtFields(iFieldIndex,1) 'data type
		case 11, 201 'skip bit, memo
			Response.Write "<td valign=top nowrap><B>" & aTxtFields(iFieldIndex,0) & "</B><br>Skipped"
			Response.Write "<BR>Type: " & aTxtFields(iFieldIndex,1)
		case else
			Response.Write "<td valign=top nowrap>"
'			Response.Write aTxtFields(iFieldIndex,1) & " - type<BR>"
			sDataSQL = "SELECT [" & aTxtFields(iFieldIndex,0) & "], Count(*) AS Qty FROM [" & gsTable & "] GROUP BY [" & aTxtFields(iFieldIndex,0) & "] ORDER BY [" & aTxtFields(iFieldIndex,0) & "]"
			if gbDebug then Response.Write sDataSQL
			oDataRS.Open sDataSQL, godbConn, 3
			WriteRS oDataRS, true, false
			oDataRS.Close
			Response.Flush
			if not Response.IsClientConnected then exit for
		end select
	next
	Response.Write "</table>"
'	sRangeSQL = BuildRangeSQL(aTxtFields, bAccess)
'	sPopSQL = BuildPopulationSQL(aTxtFields, bAccess)
	Response.Flush
	're-open rs with new sql string
'	oRangeRS.Open sRangeSQL, godbConn, 3
'	oPopRS.Open sPopSQL, godbConn, 3
'	WriteRS2 oRangeRS, oPopRS, aTxtFields, true, false, false
'	oPopRS.Close
'	oRangeRS.Close
	set oDataRS = nothing
end sub 'main

function GetFieldList()
	dim field, iFieldNo, oRS, sSQLStr, aTxtFields
	set oRs = server.CreateObject("ADODB.Recordset")
	'open recordset to get field names (1 record to limit overhead)
	sSQLStr = "SELECT TOP 1 * FROM [" & gsTable & "]"
	oRs.Open sSQLStr, godbConn, 3
	'save text field names in an array
	redim aTxtFields(oRs.Fields.Count - 1,1)
	'a(x,y)
	' x = field number
	' y = 0 : field name, y = 1 : field type
	GetTextColNames oRs, aTxtFields
	oRs.Close
	if gbDebug then
		'show array contents
		Response.Write "<P><B>Array Dump:</B> aTxtFields()</P><table><tr><th>Fieldname<th>Data Type</tr>"
		dim n
		for n = 0 to ubound(aTxtFields)
			Response.Write "<tr><td>" & aTxtFields(n,0) & " <td> " & aTxtFields(n,1)
		next
		Response.Write "</table>"
	end if
	GetFieldList = aTxtFields
end function

function BuildRangeSQL(aTxtFields, bAccess)
	'make sql string
	dim sSQLStr, iFieldNo
	sSQLStr = "SELECT"
	for iFieldNo = 0 to ubound(aTxtFields)
		'for just text fields
		select case aTxtFields(iFieldNo,1)
		case 200,202,203 'csv text, text(sql), text, memo
			sSQLStr = sSQLStr & " Max(Len([" & aTxtFields(iFieldNo,0) & "])) AS Field_"  & iFieldNo & ","
		case 2, 3, 5, 131, 7, 135 'integer, Autonumber, double, decimal, datetime, date
			'todo: need to cope with fields where result of min/max is null.
			if bAccess then
				sSQLStr = sSQLStr & " 'from: ' + CStr(Min([" & aTxtFields(iFieldNo,0) & "])) + '  to: ' + CStr(Max([" & aTxtFields(iFieldNo,0) & "])) AS Field_"  & iFieldNo & ","
			else
				sSQLStr = sSQLStr & " 'from: ' + CONVERT(varchar(12),Min([" & aTxtFields(iFieldNo,0) & "])) + ' to: ' + CONVERT(varchar(12),Max([" & aTxtFields(iFieldNo,0) & "])) AS Field_"  & iFieldNo & ","
			end if
		case 11 'bit
			if bAccess then
				sSQLStr = sSQLStr & " 'from: ' + CStr(Min([" & aTxtFields(iFieldNo,0) & "])) + '  to: ' + CStr(Max([" & aTxtFields(iFieldNo,0) & "])) + ' (boolean)' AS Field_"  & iFieldNo & ","
			else
				sSQLStr = sSQLStr & " 'n/a: type " & aTxtFields(iFieldNo,1) & "' AS Field_"  & iFieldNo & ","
'				sSQLStr = sSQLStr & " 'from: ' + CONVERT(varchar(12),Min([" & aTxtFields(iFieldNo,0) & "])) + ' to: ' + CONVERT(varchar(12),Max([" & aTxtFields(iFieldNo,0) & "])) + ' (boolean)' AS Field_"  & iFieldNo & ","
			end if
		case else
			sSQLStr = sSQLStr & " 'n/a: type " & aTxtFields(iFieldNo,1) & "' AS Field_"  & iFieldNo & ","
		end select
	next
	sSQLStr = left(sSQLStr,len(sSQLStr)-1) ' remove final comma
	sSQLStr = sSQLStr & " FROM [" & gsTable & "]"
	if gbDebug then Response.Write "<P>" & sSQLStr & "</P>" 'debugging
	BuildRangeSQL = sSQLStr
end function


sub htmlend
%>
</table>
</body>
</html>
<%
end sub

sub WriteRS(oTableRS, bWriteRC, bParse)
	if bWriteRC then Response.Write "<P>" & oTableRS.RecordCount & " records</P>"
	%>
	<table border=1 cellspacing="0" class=db>
	<TR>
	<%
	dim field, lev, sHtml, iFlush, iRecNo
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
		iFlush = iFlush + 1
		iRecNo = iRecNo + 1
		if iFlush > 10 then
			iFlush = 0
			Response.Flush
			if not Response.IsClientConnected then exit do
		end if
		if iRecNo > iFirstRecords then exit do
	loop
	%>
	</table>
	<%
end sub 'WriteRS


function GetDBConn()
	dim godbConn, sUID, sPWord, sConnStr
	sUID = session("uid")
	sPWord = session("password")
	sConnStr = session("ConnStr")
	set godbConn = server.CreateObject("ADODB.Connection")
	if gbDebug then Response.Write "<P>Connection: " & sConnStr & "</P>"
	godbConn.Open sConnStr, sUID, sPWord
	set GetDBConn = godbConn
	set godbConn = nothing
end function 'GetDBConn

sub GetTextColNames(oTableRS, aTxtFields)
'populates supplied array with field names & data types.
	dim field, iFieldNo
'	redim aTxtFields(0)
	iFieldNo = 0
	'read text field names
	for each field in oTableRS.Fields
'		Response.Write field.name & " = " & field.type & "<BR>"
'		if field.type = 200 or field.type = 202 or field.type = 203 then 'csv text or text or memo
'			redim preserve aTxtFields(iFieldNo)
			aTxtFields(iFieldNo, 0) = field.name
			aTxtFields(iFieldNo, 1) = field.type
			iFieldNo = iFieldNo + 1
'		end if
	next
'	if iFieldNo = 0 then aTxtFields(0,1) = -1
end sub 'GetColNames

function BuildPopulationSQL(aTxtFields, bAccess)
	dim sPopSQL, iFieldNo
	if gbDebug then Response.Write "<table><tr><td>SELECT " 'debugging
	sPopSQL = "SELECT "
	for iFieldNo = 0 to ubound(aTxtFields)
		select case aTxtFields(iFieldNo,1)
		case 11 'checkbox
			if bAccess then
				sPopSQL = sPopSQL & "Count( IIf([" & aTxtFields(iFieldNo,0) & "],True,Null))  AS FieldPop_"  & iFieldNo & ", "
				if gbDebug then Response.Write "<tr><td>" & "Count( IIf([" & aTxtFields(iFieldNo,0) & "],True,Null))  AS FieldPop_"  & iFieldNo & ", "'debugging
			else
			sPopSQL = sPopSQL & "Count([" & aTxtFields(iFieldNo,0) & "])  AS FieldPop_"  & iFieldNo & ", "
				if gbDebug then Response.Write "<tr><td>" &  "Count([" & aTxtFields(iFieldNo,0) & "])  AS FieldPop_"  & iFieldNo & ", "
			end if
		case 201 'ntext
				sPopSQL = sPopSQL & "'n/a' AS FieldPop_"  & iFieldNo & ", "		
		case else
			sPopSQL = sPopSQL & "Count([" & aTxtFields(iFieldNo,0) & "])  AS FieldPop_"  & iFieldNo & ", "
				if gbDebug then Response.Write "<tr><td>" & "Count([" & aTxtFields(iFieldNo,0) & "])  AS FieldPop_"  & iFieldNo & ", "
		end select
	next	
	sPopSQL = left(sPopSQL,len(sPopSQL)-2) ' remove final comma
	sPopSQL = sPopSQL & " FROM [" & gsTable & "]"
	if gbDebug then Response.Write "<tr><td> FROM [" & gsTable & "]</table>"
	if gbDebug then Response.Write "<P>" & sPopSQL & "</P>" 'debugging
	BuildPopulationSQL = sPopSQL
end function

sub WriteRS2(oTableRS, oPopRS, aTxtFields, bWriteRC, bParse, bRecNos) 'write table sideways - special case.
	if bWriteRC then Response.Write oTableRS.Fields.count & " Fields</P>"
	Response.Flush
	%>
	<table border=1 cellspacing="0">
	<TR><TH>Field Name</TH><TH>Max Chars/Range</TH><TH>Population</TH></TR>
	<%
	if bRecNos then
		dim iRecNo
		iRecNo = 1
		if oTableRS.RecordCount > 0 then oTableRS.MoveFirst
		do while not oTableRS.EOF
			Response.Write "<td>" & iRecNo
			iRecNo = iRecNo + 1
			oTableRS.MoveNext
		loop
		Response.Write "</TR>" & vbcrlf	
	end if
	dim field, lev, sHtml, iFieldIndex
	iFieldIndex = 0
	for each field in oTableRS.Fields
		'write heading
		lev = not lev
		Response.Write "<TR><TD"
		if lev then Response.Write " class=""bh"""
		Response.Write ">" & aTxtFields(iFieldIndex,0) &  "</TD>" & vbcrlf
		'write data
		if oTableRS.RecordCount > 0 then oTableRS.MoveFirst
		Response.Write "<TD"
		if lev then Response.Write " class=""bh"""
		Response.Write " nowrap align=left>"
		if isnull(field.value) then 'write [null] for nulls
			Response.Write "<span class=""n""> [ null ]</span>"
		elseif trim(field.value) <> "" then	'write non blank data
			if not bParse then 
				Response.Write "&nbsp;" & server.HTMLEncode(field.value)
			else
				Response.Write field.value
			end if
		else 'write blank data
			Response.Write "&nbsp;"
		end if
		Response.Write "</TD>" & vbcrlf
		Response.Write "<TD"
		if lev then Response.Write " class=""bh"""
		Response.Write " nowrap align=left>"
		Response.Write oPopRS.Fields(iFieldIndex)
		Response.Write "</TD></TR>" & vbcrlf
'		Response.Flush
		iFieldIndex = iFieldIndex+1
	next
	%>
	</table>
	<%
end sub 'WriteRS
%>