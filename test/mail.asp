<%
option explicit

dim oMail, sE
set oMail = server.CreateObject("CDONTS.Newmail")
sE = "tim@abell.fslife.co.uk"'"tabell@rhetorik.co.uk"
oMail.From = sE
oMail.To = sE
oMail.Subject = "arse"
oMail.Body = "tits"
oMail.Send
set oMail = nothing


%>
done