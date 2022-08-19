<%@ Language=VBScript %>
<%option explicit%>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Delete a Member Record</title>
<link rel="stylesheet" type="text/css" href="../../../EntryStyle.css">
</head>
<body Class="EntryStyle">
<h3 align="center">Delete a Member Record</h3>
<%
' Get the ID number of the requested record from the memberdetails.asp page:
dim refid
refid=CInt(Request.Form("refid"))  '  Record id of selected member
'  Open the database:
dim rs,i,j,Q
set rs=Server.CreateObject("ADODB.Recordset")
Q = "DELETE FROM [tblLogins] WHERE [ID] = " & refid
Response.Write ("<p>Query: " & Q & "</p>")
'  Set parameters to allow deletes:
rs.Open Q, LoginDSN,1,3
set rs = nothing
%>
<p><strong>Record <%=refid%> has been deleted as requested.</strong></p>
<br>
<p><a href="memberlist.asp?id=<%=loginid%>"><b>Return to Member List page</b></a></p>
<p><a href="admin.asp?id=<%=loginid%>"><b>Return to Admin page</b></a></p>

</body>
</html>
