<%option explicit%>
<%Response.Buffer = true%>
<html>
<head>
<title>Get database</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#ffffff" text="#000000">
<p><font size="+2">Open me a recordset</font></p>

<%
	'Response.Redirect "debug.asp?message=" & 	'debugging message
dim sDB, sDisplayDB, sdbConStr, sUID, sPassword, iCountRecords
sDB = Request.Form("db")
'Response.Write Request.Form("CountRecords")
iCountRecords = CBool(Request.Form("CountRecords")="true")
dim odbConn
set odbConn = server.CreateObject("ADODB.Connection")
Response.Write "<P><B>UID:</B> " & Request.Form("uid") & "<BR>"

select case Request.Form("type")
case "dsn"
	sDisplayDB = "<B>DSN Name:</B> " & sDB
	sdbConStr = sDB
	sUID = Request.Form("uid")
	sPassword = Request.Form("password")
case "sql"
	sdbConStr = "driver={SQL Server};Server=" & Request.Form("server") _
	  & ";database=" & sDB & ";APP=Pie - Online Database Tool by Tim Abell" _
	  & ";WSID=" & Request.ServerVariables("SERVER_NAME") _
	  & " (Remote: " & Request.ServerVariables("REMOTE_HOST") & " User: " & Request.ServerVariables("LOGON_USER") & ")"
	sDisplayDB = "<B>SQL Server:</B> " & Request.Form("server") & "<BR><B>Database:</B> " & sDB
	sUID = Request.Form("uid")
	sPassword = Request.Form("password")
case "mdb"
	if Request.Form("map") = "true" then _
		sDB = server.MapPath(sDB)
	sDisplayDB = "<B>Access Database:</B> " & sDB
'	sdbConStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & sDB
	sdbConStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sDB 'jet4
	sUID = Request.Form("uid")
	sPassword = Request.Form("password")
case "text"
	if Request.Form("map") = "true" then _
		sDB = server.MapPath(sDB)
	sDisplayDB = "<B>Text File Directory:</B> " & sDB
	sdbConStr = "driver={Microsoft Text Driver (*.txt; *.csv)};defaultdir=" & sDB
	sUID = ""
	sPassword = ""
case "cnstr"
	sdbConStr = sDB
	sUID = Request.Form("uid")
	sPassword = Request.Form("password")
end select
Response.Write sDisplayDB
if Request.Form("showcnstr") = "true" then Response.write "<BR><B>Connection:</B> " & sdbConStr
odbConn.Open sdbConStr, sUID, sPassword
Response.Write "</P>"
%>
<FORM id=FORM1 name=FORM1 action="display.asp" method=post>
<P><B>Use own SQL String:</B><BR>
<blockquote>
<TEXTAREA name=sqlstr rows=10 cols=70>SELECT * FROM
</TEXTAREA><BR>
<input type="submit" value="Show recordset" id=submit1 name=submit1>
</P>
</blockquote>
</FORM>
<P><B>You can use:</B></P>
<blockquote>
<p><a href="xtras">Extras</a>
</blockquote>
<P><B>Or select a table to view:</B></P>
<blockquote>
<%
session("uid") = Request.Form("uid")
session("password") = Request.Form("password")
session("ConnStr") = sdbConStr
session("db") = sDisplayDB
session("parse") = Request.Form("parse")

'write list of tables and views
'references for schema: http://msdn.microsoft.com/library/psdk/dasdk/mdae3p6g.htm, http://msdn.microsoft.com/library/psdk/his/thorref4_3ujm.htm
Response.Write "<P><B>Tables:</B><BR>"
dim oTablesRS 'create recordset to hold results
	'tables
set oTablesRS = odbConn.OpenSchema(20,array(empty,empty,empty,"TABLE"))'adSchemaTables
ListTables oTablesRS, iCountRecords 'show list
oTablesRS.Close
	'views
Response.Write "</P><P><B>Views:</B><BR>"
set oTablesRS = odbConn.OpenSchema(20,array(empty,empty,empty,"VIEW"))
ListTables oTablesRS, false
oTablesRS.Close
'release objects
set oTablesRS = nothing
odbConn.close
set odbConn = nothing

%>
</blockquote>
</body>
</html>

<%
sub ListTables(oTablesRS, RecordCounts)
	do while not oTablesRS.EOF
		Response.Write "<A href=""display.asp?tbl=" & oTablesRS("TABLE_NAME") & """>"
		Response.Write oTablesRS("TABLE_NAME")
		if recordcounts then
			Response.Write " <font color=""#666666"">(" & CountRecords(oTablesRS("TABLE_NAME")) & " records)</font>"
		end if
		Response.Write "</A>"
		Response.Write " " & oTablesRS("DESCRIPTION") & "<BR>"
		oTablesRS.MoveNext
		if not Response.IsClientConnected then Response.End
		Response.Flush
	loop
end sub

function CountRecords(TableName)
	dim oCountRS
	set oCountRS = Server.CreateObject("ADODB.Recordset")
	oCountRS.Open "SELECT COUNT(*) AS RC FROM [" & TableName & "]",odbConn,3
	CountRecords = oCountRS("RC")
	oCountRS.Close
	set oCountRS = nothing
end function
%>
