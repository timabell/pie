<%option explicit
Response.Buffer = true
dim gbDebug
gbDebug = (Request.QueryString("debug") = "true")
%>
<html>
<head>
<title>DB Xtras</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.b {  background-color: #CCFFFF;}
td,th { border:#006600 solid; border-width: 0px 1px}
th { border-width: 1px; text-align:left;}
table {border:#006600 solid 1px;}
.n {
	color : Gray;
	font : xx-small;
}
-->
</style>
</head>

<body bgcolor="#ffffff" text="#000000">
<%
if len(session("ConnStr"))<1 then
	Response.Write "<P><B>Session timed out, db connection lost.<B><BR>Go back and refresh last page to resend conection details</B></P>"
	if Request.Form.Count>0 then _
		Response.Write "<P>SQL String:<BR>" & Request.Form("sqlstr") & "</P>"
	Response.End
end if
%>
<p><font size="+2">Data cruncher - zero length strings</font></p>
<%
dim odbConn, oRs, sTable, sSQLStr
'open db
set odbConn = GetDBConn
set oRs = server.CreateObject("ADODB.Recordset")

sTable = Request.QueryString("tbl")
Response.Write "Table: " & sTable
sSQLStr = "SELECT TOP 1 * FROM [" & sTable & "]"
if gbDebug then Response.Write "<P>" & sSQLStr & "</P>"
Response.Flush
dim field, aTxtFields(), iFieldNo
oRs.Open sSQLStr, odbConn, 3
GetTextColNames oRs, aTxtFields
oRs.Close
if gbDebug then
	'show array contents
	Response.Write "<P><B>Array Dump:</B> aTxtFields()</P><table><tr><th>Fieldname</tr>"
	dim n
	for n = 0 to ubound(aTxtFields)
		Response.Write "<tr><td>" & aTxtFields(n)
	next
	Response.Write "</table>"
end if

if Request.QueryString("change") = "true" then
	dim iRecsAffected, iTotalRecsAffected
	iTotalRecsAffected = 0
	'make update sql strings
	for iFieldNo = 0 to ubound(aTxtFields)
		sSQLStr = "UPDATE [" & sTable & "] SET " & "[" & aTxtFields(iFieldNo) & "] = NULL "
		sSQLStr = sSQLStr & "WHERE [" & aTxtFields(iFieldNo) & "] = ''"
		if gbDebug then Response.Write "<P>" & sSQLStr
		odbConn.Execute sSQLStr, iRecsAffected
		iTotalRecsAffected = iTotalRecsAffected + iRecsAffected
		if gbDebug then Response.Write "<BR>Affected " & iRecsAffected & " rec(s)</P>"
	next
	Response.Write "<P align=center>Removed " & iTotalRecsAffected & " zero length string(s) from  data.</P>"
end if

'make sql string
sSQLStr = "SELECT * FROM [" & sTable & "] WHERE "
for iFieldNo = 0 to ubound(aTxtFields)
	sSQLStr = sSQLStr & "[" & aTxtFields(iFieldNo) & "] = '' OR "	
next
if right(sSQLStr,3)<>"OR " then 'nothing to do
	Response.Write "<P>No text fields</P>"
	set oRs = nothing
	odbConn.Close
	set odbConn = nothing
	Response.End
end if
sSQLStr = left(sSQLStr,len(sSQLStr)-4)
if gbDebug then Response.Write "<P>" & sSQLStr & "</P>"
're-open rs with new sql string
if Request.QueryString("change") = "true" then
	%>
	<P><A href="<%=Request.ServerVariables("SCRIPT_NAME")%>?tbl=<%=Request.QueryString("tbl")%>">back</A></P>
	<%
else
	%>
	<P><A href="<%=Request.ServerVariables("SCRIPT_NAME")%>?tbl=<%=Request.QueryString("tbl")%>&change=true<%if gbDebug then Response.Write "&debug=true"%>">Change zls's to null</A></P>
	<%
end if
oRs.Open sSQLStr, odbConn, 3
WriteRS ors, true, false
oRs.Close
set oRs = nothing
odbConn.Close
set odbConn = nothing

%>
</table>
</body>
</html>

<%
sub WriteRS(oTableRS, bWriteRC, bParse)
	if bWriteRC then Response.Write "<P>" & oRs.RecordCount & " records</P>"
	%>
	<table border=1 cellspacing="0">
	<TR>
	<%
	dim field, lev, sHtml
	'write headings
	for each field in oRs.Fields
		Response.Write "<TH nowrap>" & field.name & "</TH>" & vbcrlf
	next
	Response.Write "</TR>" & vbcrlf
	'write data
	if oRs.RecordCount > 0 then oRs.MoveFirst
	do while not oRs.EOF
		lev = not lev
		Response.Write "<TR>" & vbcrlf
		for each field in oRs.Fields
			Response.Write "<TD nowrap"
			if lev then Response.Write " class=""b"""
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
		oRs.MoveNext
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
	odbConn.Open sConnStr, sUID, sPWord
	set GetDBConn = odbConn
	set odbConn = nothing
end function 'GetDBConn

sub GetTextColNames(oTableRS, aTxtFields)
	dim field, iFieldNo
	redim aTxtFields(0)
	iFieldNo = 0
	'read text field names
	for each field in oTableRS.Fields
		if gbDebug then Response.Write field.name & " = " & field.type & "<BR>"
		if field.type = 202 or field.type = 203 then 'text or memo
			redim preserve aTxtFields(iFieldNo)
			aTxtFields(iFieldNo) = field.name
			iFieldNo = iFieldNo + 1
		end if
	next
end sub 'GetColNames
%>