<%option explicit%>
<html>
<head>
<title>DB Xtras</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
<p><font size="+2">Data cruncher</font></p>

<P>Select table/view to work on:</P>
<blockquote>
<%
dim sUID, sPWord, sConnStr
sUID = session("uid")
sPWord = session("password")
sConnStr = session("ConnStr")

dim odbConn, oTablesRS
set odbConn = server.CreateObject("ADODB.Connection")
set oTablesRS = server.CreateObject("ADODB.Recordset")
odbConn.Open sConnStr, sUID, sPWord


Response.Write "<P><B>Tables:</B><BR>"
set oTablesRS = odbConn.OpenSchema(20,array(empty,empty,empty,"TABLE"))'adSchemaTables
	'references for above: http://msdn.microsoft.com/library/psdk/dasdk/mdae3p6g.htm, http://msdn.microsoft.com/library/psdk/his/thorref4_3ujm.htm
	ListTables oTablesRS
oTablesRS.Close
Response.Write "</P><P><B>Views:</B><BR>"
set oTablesRS = odbConn.OpenSchema(20,array(empty,empty,empty,"VIEW"))
	ListTables oTablesRS
oTablesRS.Close
set oTablesRS = nothing
odbConn.close
set odbConn = nothing
%>
</blockquote>
</body>
</html>

<%
sub ListTables(oTablesRS)
	do until oTablesRS.EOF
		Response.Write "<A href=""pickprocess.asp?tbl=" & oTablesRS("TABLE_NAME") & """>"
		Response.Write oTablesRS("TABLE_NAME") & "</A>"
		Response.Write " " & oTablesRS("DESCRIPTION") & "<BR>"
		oTablesRS.MoveNext
	loop
end sub
%>