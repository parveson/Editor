<%@ Language=VBScript %>
<html>
<head>
<title>Cartoon Characters</title>
</head>
<body bgcolor="#FF5555">
<h2 align="center">Cartoon Characters:</h2>

<p>Test database using file DSN in this page, Jet OLEDB 4.0 for Access 2003</p>

<%dim rs,sDSN,Q,DSN1,DSN2 
Q = "SELECT * from tblCharacter"
'  Jet3 is for Access97 databases; Jet4 is for Access2000 databases.
	dim datapath
	datapath = Server.MapPath("\database")
	Response.Write("<p>datapath = " & datapath & "</p>")
	datapath = datapath & "\"
	DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath
	DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
	sDSN = DSN1 & "cartoon2000.mdb" & DSN2
'sDSN=Application("Jet4_ConnectionString")
Response.Write ("<p>DSN= " & sDSN & "</p>")
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open Q,sDSN,3,3

%>
<table align="center" border="2" bgcolor="white">
<tr>
	<th>Name</th>
	<th>Age</th>
	<th>Show</th>
</tr>
	<%do while not rs.EOF
		%><tr><td><%=rs.Fields("Name")%></td>
				<td><%=rs.Fields("Age")%></td>
				<td><%=rs.Fields("Show")%></td>
			</tr>
	<%rs.MoveNext
	loop
	rs.Close
	Set rs=nothing
	%>
	
</table>
<br>
<p><a href="../../bsc4_local/default.htm">Return</a></p>
</body>
</html>
