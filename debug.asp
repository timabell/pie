<html>
<head>
<title>Debug notification</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center"> 
  <table width="100%" border="0">
    <tr>
      <td width="50%"><font size="2">message from: <A HREF="<%=request.servervariables("HTTP_REFERER")%>">
      <%=request.servervariables("HTTP_REFERER")%></A>
        </font></td>
      <td width="50%"> 
        <div align="right"><font size="2"><%=now()%></font></div>
      </td>
    </tr>
  </table>
  <br>
  <b>Debug Notification<br>
    </b>
</div>
<hr align="center">
<p align="center"><br>
  <%=request.querystring("message")%></p>
<hr><P>
<%
if Request.Form.Count > 0 then
	Response.Write "<B>Submitted Form:</B></P>"
	ShowForm
else
	Response.Write "No form submitted."
end if
%>

  </body>
</html>
<%
sub ShowForm
	dim item
	for each item in Request.Form
		Response.Write "<B>" & item & "</B>, " & Request.Form(item) & "<BR>"
	next
end sub
%>