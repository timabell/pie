<%option explicit
Response.Buffer = true
server.ScriptTimeout = 900 '15 mins
'set globals
dim gsTable, gbDebug
gsTable = Request.QueryString("tbl")
gbDebug = (Request.QueryString("debug") = "true")
%>
<html>
<head>
<title>Dedupe data</title>
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

<body bgcolor="#FFFFFF" text="#000000">
<%
if len(session("ConnStr"))<1 then
	Response.Write "<P><B>Session timed out, db connection lost.<B><BR>Go back and refresh last page to resend conection details</B></P>"
	if Request.Form.Count>0 then _
		Response.Write "<P>SQL String:<BR>" & Request.Form("sqlstr") & "</P>"
	Response.End
end if

%>
<p><font size="+2">Data cruncher - Data dedupe</font></p>
<%

if Request.Form.Count > 0 then
	if gbDebug then ShowForm 'debug
	if Request.Form("deletions") = "" then
		ShowDupeList
	else
		DedupeData
	end if
else
	WriteForm
end if


%>
</body>
</html>
<%

function AddColumnSQL(sColumnName)
	dim sSQL
	sSQL = "ALTER TABLE [" & gsTable & "] ADD COLUMN [" & sColumnName & "] TEXT(255)"
	AddColumnSQL = sSQL
end function

sub ShowDupeList
	dim oDBConn, oRS, sSQL
	set oDBConn = GetDBConn
	set oRS = Server.CreateObject("ADODB.Recordset")
	sSQL = MakeDupeListSQL
	session("dedupesql") = sSQL
	if gbDebug then Response.Write sSQL
	ors.open sSQL, oDBConn, 3
	WriteDupeRS oRS, true,false
	ors.close
	set ors = nothing
	oDBConn.Close
	set oDBConn = nothing
end sub

function MakeDupeListSQL
	dim sSQL, aFields, iField
	aFields = split(Request.Form("InputField"),", ")
	sSQL = "SELECT [" & gsTable & "].* FROM [" & gsTable & "] INNER JOIN (SELECT"
	'field list
	for iField = 0 to UBound(aFields)
		sSQL = sSQL & " [" & aFields(iField) & "],"
	next
	sSQL = left(sSQL,len(sSQL)-1) 'remove last ","
	sSQL = sSQL & " FROM [" & gsTable & "] GROUP BY "
	'field list
	for iField = 0 to UBound(aFields)
		sSQL = sSQL & " [" & aFields(iField) & "],"
	next
	sSQL = left(sSQL,len(sSQL)-1) 'remove last ","
	sSQL = sSQL & " HAVING Count([" & aFields(0) & "])>1) AS DupeList ON "
	'join field list
	for iField = 0 to UBound(aFields)
		sSQL = sSQL & "([" & gsTable & "].[" & aFields(iField) & "] = DupeList.[" & aFields(iField) & "]) AND "
	next
	sSQL = left(sSQL,len(sSQL)-5) 'remove last "AND"
	sSQL = sSQL & " ORDER BY "
	'sort field list
	for iField = 0 to UBound(aFields)
		sSQL = sSQL & "DupeList.[" & aFields(iField) & "], "
	next
	if Request.Form("DeletionPriority") = "" then
		sSQL = left(sSQL,len(sSQL)-2) 'remove last ","
	else
		sSQL = sSQL & "[" & gsTable & "].[" & Request.Form("DeletionPriority") & "]"
		if Request.Form("desc") = "on" then sSQL = sSQL & " DESC"
	end if	
	MakeDupeListSQL = sSQL
end function

sub WriteForm
	dim odbConn, oRs, sSQLStr
	'open db
	set odbConn = GetDBConn
	set oRs = server.CreateObject("ADODB.Recordset")
	'open recordset to get field names
	Response.Write session("db")
	Response.Write "<P>Table: " & gsTable & "</P>"
	sSQLStr = "SELECT TOP 1 * FROM [" & gsTable & "]"
	oRs.Open sSQLStr, odbConn, 3
	'save text field names in an array
	dim field, aTxtFields(), iFieldNo
	redim aTxtFields(oRs.Fields.Count - 1,1)
	'a(x,y)
	' x = field number
	' y = 0 : field name, y = 1 : field type
	GetTextColNames oRs, aTxtFields
	oRs.Close
	set oRs = nothing
	odbConn.Close
	set odbConn = nothing
	%>
	<blockquote>
	<form name="form1" method="post" action="">
	<table border=1>
	<TR>
	<TH valign=top><P>Select column to use <br>as unique index:<BR>
	<TH valign=top><P>Select column(s) to <br>make unique:<BR>
	<TH valign=top><P>Select column to use<br>for deletion priority:<BR>
	<TR>
	<TD valign=top>
		<select name="IndexField">
		<%WriteColumnOptions aTxtFields%>
		</select>
	<TD valign=top>
		<select multiple size=<%=ubound(aTxtFields)+1%> name="InputField">
		  <%WriteColumnOptions aTxtFields%>
		</select>
	<TD valign=top nowrap>
		<select name="DeletionPriority">
		<option value="">--None--</option>
		<%WriteColumnOptions aTxtFields%>
		</select>
		<BR><BR><input type=checkbox name=desc> Sort descending
		<font size=-1>
		<BR><BR>records will be sorted by
		<BR> this field and higher records
		<BR> deleted in preference</font>
	<TR><TD colspan=2 align=center>
	    <input type="submit" name="Submit" value="Go">
	<TD>&nbsp;
	</table>
	</form>
	</blockquote>
	<p>&nbsp;</p>
	<%
end sub

sub DedupeData
	dim oDBConn, oRS, sSQL, iCount, sIndexField, bDelete
	sIndexField = Request.Form("IndexField")
	sSQL = session("dedupesql")
	if ssql = "" then
		Response.Write "Error, Lost SQL String"
		exit sub
	end if
	
	set oDBConn = GetDBConn
	bDelete = Request.Form("delete") <> ""
	if bDelete then
		sSQL = "DELETE FROM [" & gsTable & "]"
	else
		sSQL = "UPDATE [" & gsTable & "] SET [" & Request.Form("MarkField") & "] = True"
	end if
	sSQL = sSQL & " WHERE [" & sIndexField & "] In(" & Request.Form("deletions") & ")"
	if gbDebug then Response.Write "<P>" & sSQL & "</P>"
	oDBConn.Execute sSQL,iCount
	Response.Write "<P>" & iCount & " Records " & IIf(bDelete,"Deleted","Updated") & "</P><HR>"
end sub

sub GetTextColNames(oTableRS, aTxtFields)
	dim field, iFieldNo
	iFieldNo = 0
	'read text field names
	for each field in oTableRS.Fields
'		Response.Write field.name & " = " & field.type & "<BR>"
			aTxtFields(iFieldNo, 0) = field.name
			aTxtFields(iFieldNo, 1) = field.type
			iFieldNo = iFieldNo + 1
	next
end sub 'GetColNames

sub WriteColumnOptions(aTxtFields)
	dim iItem
	for iItem = 0 to UBound(aTxtFields)
		Response.Write "<option>" & aTxtFields(iItem,0) & "</option>"
	next
end sub

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

sub WriteDupeRS(oTableRS, bWriteRC, bParse)
	dim field, lev, sHtml, sFields, aFields, aPrev, bChanged, sIndexField
	if bWriteRC then Response.Write "<P>" & oTableRS.RecordCount & " duplicates </P> "
	%>
	<form name="form1" method="post" action="">
	<input type=hidden name=InputField value="<%=Request.Form("InputField")%>">
	<input type=hidden name=IndexField value="<%=Request.Form("IndexField")%>">
	<P>
	<table border=1 cellpadding=2><tr><td>
	<input type=submit value="Delete Selected" name=delete>
	<th>or:<td> <input type=submit value="Mark Selected" name=mark> in column 
		<select name="MarkField">
		<%
		for each field in oTableRS.Fields
			Response.Write "<option>" & field.name & "</option>"
		next
		%>
		</select>
		</table>
	</P>
	<table border=1 cellspacing="0" class=db>
	<%
	sFields = Request.Form("InputField")
	sIndexField = Request.Form("IndexField")
	aFields = split(sFields,", ")
	aPrev = aFields
	'write column groups
	Response.Write "<COLGROUP><COL>"
	for each field in oTableRS.Fields
		if InArray(aFields,field.name) then
			Response.Write "<COL style=""background:#e6e6fa"">" & vbcrlf
		else
			Response.Write "<COL>" & vbcrlf
		end if
	next
	Response.Write "</COLGROUP>"
	'write headings
	Response.Write "<TR><TH nowrap class=db>delete"
	for each field in oTableRS.Fields
		if InArray(aFields,field.name) then
			Response.Write "<TH nowrap class=db><font color=#FF0000>" & field.name & "</font></TH>" & vbcrlf
		else
			Response.Write "<TH nowrap class=db>" & field.name & "</TH>" & vbcrlf
		end if
	next
	Response.Write "</TR>" & vbcrlf
	'write data
	if oTableRS.RecordCount > 0 then oTableRS.MoveFirst
	do while not oTableRS.EOF
		bChanged =  not MatchesLast(oTableRS, aPrev, aFields)
		if bChanged then lev = not lev
		Response.Write "<TR class=db>" & vbcrlf
		Response.Write "<TD nowrap"
		if lev then Response.Write " class=b" else Response.Write " class=db" 
		Response.Write "><input type=checkbox name=deletions"
		if not bChanged then  Response.Write " checked"
		Response.Write " value=" & oTableRS(sIndexField) & ">"
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
	</form>
	<%
end sub 'WriteRS

function InArray(aArray, Value)
	dim iIndex
	InArray = false
	for iIndex = 0 to UBound(aArray)
		if aArray(iIndex) = value then
			InArray = true
			exit function
		end if
	next
end function

function MatchesLast(oRS, aLast, aFields)
	dim iIndex
	MatchesLast = true
	for iIndex = 0 to UBound(aFields)
'		if not(isnull(oRS(aFields(iIndex))) or isnull(aLast(iIndex))) then
		
			if lcase(oRS(aFields(iIndex))) <> lcase(aLast(iIndex)) then
				MatchesLast = false
				aLast(iIndex) = oRS(aFields(iIndex))
			end if
'			if oRS(aFields(iIndex)) <> aLast(iIndex) then
'				MatchesLast = false
'				aLast(iIndex) = oRS(aFields(iIndex))
'			end if
'		end if
	next
end function

sub ShowForm
	dim item
	for each item in Request.Form
		Response.Write "<B>" & item & ":</B> " & Request.Form(item) & "<BR>"
	next
end sub


function IIf(bTest,vTrue,vFalse)
	if bTest then
		iif = vTrue
	else
		iif = vFalse
	end if
end function
%>

