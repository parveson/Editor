<%@ Language=VBScript %>
<%option explicit%>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Confirm to Delete a Record</title>
<link rel="stylesheet" type="text/css" href="../../../EntryStyle.css">
</head>
<body Class="EntryStyle">
<h3 align="center">Confirm/Delete a Record</h3>
<%
' Get the ID number of the requested record from the guestdetails.asp page:
dim refid
refid=CInt(Request.Querystring("refid"))  '  Record id of selected company
dim formvalue 
If not isEmpty(refid) then
	dim rs,i,q,fullname,fixname,count,quot
	' select the requested record:
	q = "SELECT * FROM [tblLogins] WHERE [ID] = " & refid
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open q,LoginDSN
	' Display some fields in the record:
	if not rs.EOF then
		Response.Write("<p>ID = " & rs("ID") & "</p>")
		Response.Write ("<p>Name: <strong>" & rs("Fuser") & " " & rs("Luser") & "</strong></p>")
		Response.Write("<p>Email: <strong>" & rs("Email") & "</p>")
		rs.Close
		set rs=nothing
		%>
		<h4 align="center">Are you sure you want to delete this record?</h4>
		<table align="center" border="0">
		<tr><td>
		<form name="doit" method="post" action="adelete.asp?id=<%=loginid%>">
			<input type="submit" name="doit" value="Yes, Delete">
			<input type="hidden" value="<%=refid%>" name="refid">
		</form>
		</td>
		<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td>
		<form name="back" method="post" action="memberdetails.asp?id=<%=loginid%>&amp;refid=<%=refid%>">
			<input type="submit" name="back" value="No, Cancel">
			<input type="hidden" value="<%=refid%>" name="refid">
		</form>
		</td></tr>
		</table>
		<%
	else
		Response.Write("<p><strong>This record does not exist in the database.</stron></p>")
	end if
end if
%>
</body>
</html>
