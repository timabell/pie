
<form action="encode.asp" method=post><textarea name=txtval><%=Request.Form("txtval")%></textarea>
<input type=submit>
</form>

<%
Response.Write server.HTMLEncode(server.URLEncode(Request.Form("txtval")))
%>