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

<P>Select a task:</P>
<blockquote>
<P><a href="findzls.asp?tbl=<%=Request.QueryString("tbl")%>">Find Zero Length Strings</A>
<P><a href="lengths.asp?tbl=<%=Request.QueryString("tbl")%>">List data lengths</A>
<P><a href="dataview.asp?tbl=<%=Request.QueryString("tbl")%>">Data view</A> - uniques by column
 &nbsp; / &nbsp; First <a href="dataview.asp?tbl=<%=Request.QueryString("tbl")%>&first=10">10</A>,
 <a href="dataview.asp?tbl=<%=Request.QueryString("tbl")%>&first=50">50</A>,
 <a href="dataview.asp?tbl=<%=Request.QueryString("tbl")%>&first=100">100</A>,
 <a href="dataview.asp?tbl=<%=Request.QueryString("tbl")%>&first=250">250</A>,
 <a href="dataview.asp?tbl=<%=Request.QueryString("tbl")%>&first=500">500</A>
<P><a href="split.asp?tbl=<%=Request.QueryString("tbl")%>">Split delimited data</A>
<P><a href="dedupe.asp?tbl=<%=Request.QueryString("tbl")%>">Dedupe data</A>
</blockquote>
</body>
</html>