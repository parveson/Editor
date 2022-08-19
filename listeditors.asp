<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<HTML>
<HEAD>
<TITLE>List All Editors</TITLE>
<LINK rel="stylesheet" type="text/css" href="StyleSheet1.css">
<SCRIPT LANGUAGE="JavaScript" type="text/JavaScript">
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</SCRIPT>
</HEAD>
<BODY>
<p class=small>Administration</p>
<h4 align=center>List of All Editors/Administrators</h4>
<%
dim diagnostic,email
diagnostic=false
dim prefix,fname,lname,phone,fullname,pw,entrydate,expertise,notes 
dim refdate,quot,refid,role,recordcount,partyid
role=2   '  Authors - numeric value
dim rs,I,J,Q, sDSN  ' other variables are defined in test2.asp
datapath = Server.MapPath("\database")
datapath = datapath & "\"
DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath
DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
sDSN = DSN1 & "Reviews.mdb" & DSN2
' Before storing data, verify that the party (AS AN AUTHOR) is not already stored:
Q = "SELECT * FROM [tblParty] WHERE [Role]= 1 ORDER BY [tblParty.Lname]"
' Select all records using GetRows:
' Thanks to http://www.asp101.com/samples/db_getrows.asp
dim cnnGetRows   ' ADO connection
dim strDBPath    ' Path to our Access DB (*.mdb) file
dim arrDBData    ' Array that we dump all the data into
dim iRecFirst, iRecLast
dim iFieldFirst, iFieldLast
set cnnGetRows = Server.CreateObject("ADODB.Connection")
cnnGetRows.Open sDSN
set rs=cnnGetRows.Execute(Q)
arrDBData = rs.GetRows(-1,0,Array("PartyID","Prefix","Fname","Lname","Email","Phone","Role","Expertise","Notes"))
rs.Close
cnnGetRows.Close
set cnnGetRows=nothing			
iRecFirst   = LBound(arrDBData, 2)
iRecLast    = UBound(arrDBData, 2)
iFieldFirst = LBound(arrDBData, 1)   '  = 0
iFieldLast  = UBound(arrDBData, 1)
' Display a table of the data in the array.
' Table header:
Response.Write("<table align=center bgcolor='#ffffff' border=1 cellspacing=0 width=650>")
Response.Write("<th>Name</th><th>Email</th><th>Phone</th><th>Notes</th><th>Delete?</th></tr>")
recordcount=0
For I = iRecFirst To iRecLast
	recordcount = recordcount + 1
	' A table row for each record					
	' Columns: Loop through the fields (first dimension of the array)
	For J = 0 To iFieldLast
		' Allow for special treatment of each field:
		select case J
		case 0   ' PartyID
			partyid = arrDBData(J,I)
		case 1 ' Prefix
			prefix = arrDBData(J,I)
		case 2 ' First Name
			fname=arrDBData(J,I)
		case 3  ' Last Name
			lname=arrDBData(J,I)
			fullname = prefix & " " & fname & " " & lname
			Response.Write("<td align=center>" & fullname & "</td>")
		case 4 ' Email
			email = arrDBData(J,I)
			Response.Write("<td align=center class=small><a href='mailto:" & email & "'>" & email & "</a></td>") 
		case 5 ' Phone
			Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
		case 6  ' Role
			role = arrDBData(J,I)
			if not role=1 then
				Response.Write("<H5>Error -  Role = " & role & " J=" & J & " I=" & I & " </H5>")
			end if
		case 7  ' Expertise - not used for editors
			'Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
		case 8   ' Notes
			notes = left(arrDBData(J,I),40)  '  Truncate the notes for now.
			Response.Write("<td align=center class=small>" & notes & "</td>")			
		case else
			Response.Write("Error - too many fields in array: J=" & J & " I= " & I & " data=" & arrDBData(J,I))			
		end select
	Next ' J				
	Response.Write "<td align=center><a href='confirm.asp?id=" & loginid & "&partyid=" & partyid & "'>Delete</a></td></tr>" 
Next ' I
Response.Write("</table>")
set rs = nothing
if recordcount=1 then
    Response.Write("<P>One record found.</p>")
elseif recordcount > 1 then
    Response.Write("<p>" & recordcount & " records found.</p>")
else
    Response.Write("<p>No records found.</p>")
end if
%>
<br>
<p><b><a href="addeditor.asp?id=<%=loginid%>">Add an editor</a></b></p>
<br>
<p><A href="menu.asp?id=<%=loginid%>">Return to Admin menu page</A></p>
</BODY>
</HTML>


