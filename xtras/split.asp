<%option explicit
Response.Buffer = true
%>
<html>
<head>
<title>Split data</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
if len(session("ConnStr"))<1 then
	Response.Write "<P><B>Session timed out, db connection lost.<B><BR>Go back and refresh last page to resend conection details</B></P>"
	if Request.Form.Count>0 then _
		Response.Write "<P>SQL String:<BR>" & Request.Form("sqlstr") & "</P>"
	Response.End
end if
'set globals
dim sTable
sTable = Request.QueryString("tbl")

%>
<p><font size="+2">Data cruncher - data splitting</font></p>
<%

if Request.Form.Count > 0 then
'ShowForm
	SplitData
else
	WriteForm
end if


%>
</body>
</html>
<%
sub SplitData
	dim sSQL1, sSQL2, sSQL3, sSQL4
	'add any new columns
	if Request.Form("outputtype1") = "new" then
		sSQL1 = AddColumnSQL(Request.Form("NewField1"))
	end if
	if Request.Form("outputtype2") = "new" then
		sSQL2 = AddColumnSQL(Request.Form("NewField2"))
	end if
	'remove chr from end of strings
	sSQL3 = TrimDataSQL
	'run splitter
	sSQL4 = MakeSplitSQL
	dim odbConn, iRecsAffected
	'open db
	set odbConn = GetDBConn
	if Request.Form("outputtype1") = "new" then
		Response.Write "<P>Add Column1:<BR>" & sSQL1 & "<br>"
		odbConn.Execute sSQL1
		Response.Write " - done.</P>"
	end if
	
	if Request.Form("outputtype2") = "new" then
		Response.Write "<P>Add Column2:<BR>" & sSQL2 & "<br>"
		odbConn.Execute sSQL2
		Response.Write " - done.</P>"
	end if
	
'	Response.Write "<P>Remove trailing chrs:<br>" & sSQL3 & "<br>"
'	odbConn.Execute sSQL3', iRecsAffected
'	Response.Write " - done."', " & iRecsAffected & " records affected.</P>"
	
	Response.Write "<P>" & sSQL4 & "<br>"
	odbConn.Execute sSQL4', iRecsAffected
	Response.Write " - done."', " & iRecsAffected & " records affected.</P>"

	odbConn.Close
	set odbConn = nothing
end sub

function TrimDataSQL
	dim sField, iWidth, sSplitChr
	sField = Request.Form("InputField")

	select case Request.Form("splittype")
	case "string"
		sSplitChr = "'%" & replace(Request.Form("splitstring"),"'","''") & "'"
		iWidth = len(Request.Form("splitstring"))
	case "chrno"
		sSplitChr = "'%' & chr(" & Request.Form("splitchrno") & ")"
		iWidth = 1
	case "crlf"
		sSplitChr = "'%' & chr(13) & chr(10)"
		iWidth = 2
	case else
		Response.Write "<P>Error in splittype"
		Response.End	
	end select
	TrimDataSQL = "UPDATE [" & sTable & "] SET [" & sField & "] = Left([" & sField & "], len([" & sField & "])-" _
					 & iWidth & ") WHERE [" & sField & "] Like " & sSplitChr
end function

function MakeSplitSQL()
		'get information on task
	sInField = Request.Form("InputField")
	dim sInField, sOutField1, sOutField2, sInChr, sWhereChr, sSplitChr, iWidth
	'split chr
	select case Request.Form("splittype")
	case "string"
		sInChr = "'" & Request.Form("splitstring") & "'"
		sWhereChr = "'%" & Request.Form("splitstring") & "%'"
		iWidth = len(Request.Form("splitstring"))
	case "chrno"
		sInChr = "chr(" & Request.Form("splitchrno") & ")"
		sWhereChr = "'%' & chr(" & Request.Form("splitchrno") & ") & '%'"
		iWidth = 1
	case "crlf"
		sInChr = "chr(13) & chr(10)"
		sWhereChr = "'%' & chr(13) & chr(10) & '%'"
		iWidth = 2
	case else
		Response.Write "<P>Error in splittype"
		Response.End	
	end select
	'fields
	if Request.Form("outputtype1") = "new" then
		sOutField1 = Request.Form("NewField1")
	else
		sOutField1 = Request.Form("OutputField1")
	end if
	if Request.Form("outputtype2") = "new" then
		sOutField2 = Request.Form("NewField2")
	else
		sOutField2 = Request.Form("OutputField2")
	end if
	'calculate whole sql string
	dim sSQL, sUpdate, sCol1, sCol2, sWhere
	sUpdate = "UPDATE [" & sTable & "] SET "
	sCol1 = "[" & sOutField1 & "] = Left([" & sInField & "], InStr([" & sInField & "], " & sInChr & ")-" & iWidth & ") & ' '"
	sCol2 = "[" & sOutField2 & "] = Right([" & sInField & "], len([" & _
		sInField & "])-InStr([" & sInField & "], " & sInChr & "))"
	sWhere = " WHERE [" & sInField & "] Like " & sWhereChr
	'concantenate
	sSQL = sUpdate & sCol1 & ", " & sCol2 & sWhere
	'display
	MakeSplitSQL = sSQL
end function

function AddColumnSQL(sColumnName)
	dim sSQL
	sSQL = "ALTER TABLE [" & sTable & "] ADD COLUMN [" & sColumnName & "] TEXT(255)"
	AddColumnSQL = sSQL
end function

sub WriteForm
	dim odbConn, oRs, sSQLStr
	'open db
	set odbConn = GetDBConn
	set oRs = server.CreateObject("ADODB.Recordset")
	'open recordset to get field names
	Response.Write session("db")
	Response.Write "<P>Table: " & sTable & "</P>"
	sSQLStr = "SELECT TOP 1 * FROM [" & sTable & "]"
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
	<form name="form1" method="post" action="">
	  <table border="1" cellspacing="0" cellpadding="3">
	    <tr> 
	      <th align="left">Split column</th>
	      <td> 
	        <select name="InputField">
	          <%WriteColumnOptions aTxtFields%>
	        </select>
	      </td>
	    </tr>
	    <tr> 
	      <th valign="top" align="left">Split on</th>
	      <td> 
	        <table border="0" cellspacing="0" cellpadding="3">
	          <tr> 
	            <td align="right">String 
	              <input type="radio" name="splittype" value="string" checked>
	            </td>
	            <td> 
	              <input type="text" name="splitstring">
	            </td>
	          </tr>
	          <tr> 
	            <td align="right">Chr Number 
	              <input type="radio" name="splittype" value="chrno">
	            </td>
	            <td>
	               <input type="text" name="splitchrno" size=10	>
	            </td>
	          </tr>
	          <tr> 
	            <td align="right">CrLf 
	              <input type="radio" name="splittype" value="crlf">
	            </td>
	            <td>&nbsp;</td>
	          </tr>
	        </table>
	      </td>
	    </tr>
	    <tr> 
	      <th align="left">Split 1st half into</th>
	      <td> 
	        <table border="0" cellspacing="0" cellpadding="3">
	          <tr> 
	            <td align="right">New column 
	              <input type="radio" name="outputtype1" value="new" checked>
	            </td>
	            <td> 
	              <input type="text" name="NewField1">
	            </td>
	          </tr>
	          <tr> 
	            <td align="right">Existing
	              <input type="radio" name="outputtype1" value="existing">
	            </td>
	            <td> 
	              <select name="OutputField1">
						<%WriteColumnOptions aTxtFields%>
	              </select>
	            </td>
	          </tr>
	        </table>
	      </td>
	    </tr>
	    <tr> 
	      <th align="left">Split 2nd half into</th>
	      <td> 
	        <table border="0" cellspacing="0" cellpadding="3">
	          <tr> 
	            <td align="right">New column 
	              <input type="radio" name="outputtype2" value="new" checked>
	            </td>
	            <td> 
	              <input type="text" name="NewField2">
	            </td>
	          </tr>
	          <tr> 
	            <td align="right">Existing
	              <input type="radio" name="outputtype2" value="existing">
	            </td>
	            <td> 
	              <select name="OutputField2">
						<%WriteColumnOptions aTxtFields%>
	              </select>
	            </td>
	          </tr>
	        </table>
	      </td>
	    </tr>	    <tr> 
	      <th align="left">&nbsp;</th>
	      <td> 
	        <input type="submit" name="Submit" value="Go">
	      </td>
	    </tr>
	  </table>
	</form>
	<p>&nbsp;</p>
	<%
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

sub ShowForm
	dim item
	for each item in Request.Form
		Response.Write "<B>" & item & "</B>, " & Request.Form(item) & "<BR>"
	next
end sub
%>

