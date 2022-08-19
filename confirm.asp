<%@ Language=VBScript %>
<%option explicit%>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Confirm to Delete a Party</title>
<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
</head>
<body>
<p class=small>Administration</p>
<br>
<h3 align=center>Confirm/Delete a Party</h3>
<div align=center>
<%
' Called from addauthor.asp, addeditor.asp, or addreviewer.asp
dim partyid,sDSN,Q,rs,i,fullname,count,quot,email,roletext,role
partyid=CInt(Request.Querystring("partyid"))  '  Record id to be deleted 
If not isEmpty(partyid) then
	datapath = Server.MapPath("\database")
	DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath & "\"
	DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
	sDSN = DSN1 & "Reviews.mdb" & DSN2
	Q = "SELECT * FROM [tblParty] WHERE [PartyID]=" & partyid
	set rs=server.CreateObject("ADODB.recordset")
	rs.Open Q,sDSN,1,3
	' Display some fields in the record:
	if not rs.EOF then
		fullname=rs("Prefix") & " " & rs("Fname") & " " & rs("Lname")
		email=rs("Email")
		role=rs("Role")
		select case role
			case 1
				roletext="Editor"
			case 2
				roletext="Author"
			case 3
				roletext="Reviewer"
		end select
		Response.Write("<p>Name: <b>" & fullname & "</b>")
		Response.Write("<br>Email: <b>" & email & "</b>")
		Response.Write("<br>Role: <b>" & roletext & "</b></p>")
		rs.Close
		set rs=nothing
		%>
		<h4 align="center">Are you sure you want to delete this record?</h4>
		<%
		Response.Write("<p><b><a href='Javascript:history.back();'>NO - Go Back</a></b></p>")
		Response.Write("<p><b><a href='deleteparty.asp?id=" & loginid & "&partyid=" & partyid & "'>Delete Now</a></b></p>")
	else
		Response.Write("<p><strong>This record does not exist in the database.</strong></p>")
	end if
end if
%>
</div>
</body>
</html>
