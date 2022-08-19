<%@Language="VBScript"%>
<%option explicit
Response.expires=30
'Response.addHeader "pragma","no-cache"
'Response.addHeader "cache-control","private"
'Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
<title>Close a Review</title>
</head>
<body>
<p>Administration</p>
<br>
<%
dim rs,Q,Q2
set rs = Server.CreateObject("ADODB.Recordset")	
dim msid,assignid,reviewer,seqno,dateassigned,title,typename,versionno,author,filename
msid = Request.QueryString("msid")
if not Request.Form("submitted") then
	'  Form has not been submitted yet
	%>
	<h3 align="center">Close Review</h3>
	<p><STRONG>Are you sure you wish to close this 
review?</STRONG> <BR><BR>This will close all the MS files and review files 
from all reviewers associated with this MS (same MR#). You may view them later in the 
archive if desired.<br>
	<form method="post" action="close_review.asp?id=<%=loginid%>">
	<%
		Q = "SELECT tblFile.MsID, tblParty.Lname, tblFile.SeqNo, tblAssignment.DateAssigned, "
		Q = Q & " tblFile.Title, tblMsType.TypeName, tblFile.VersionNo, tblParty.PartyID, "
		Q = Q & " tblParty.Role, tblAssignment.AssignID, tblFile.Closed, tblFile.FileName "
		Q = Q & " FROM tblMsType INNER JOIN (tblParty INNER JOIN (tblFile INNER JOIN tblAssignment "
		Q = Q & " ON tblFile.MsID = tblAssignment.MsID) ON tblParty.PartyID = tblAssignment.PartyID) "
		Q = Q & " ON tblMsType.TypeID = tblFile.MsTypeID "
		Q = Q & " WHERE tblFile.MsID= " & msid 
		rs.Open Q,LoginDSN
		reviewer = rs("Lname")
		seqno = rs("SeqNo")
		dateassigned=rs("DateAssigned")
		title=rs("Title")
		typename=rs("TypeName")
		versionno = rs("VersionNo")
		filename = rs("FileName")
		rs.Close
		'  Query: Get author's name:
		Q2 = "SELECT tblFile.MsID, tblFile.AuID, tblParty.PartyID, tblParty.Lname "
		Q2 = Q2 & " FROM tblParty INNER JOIN tblFile ON tblParty.PartyID = tblFile.AuID "
		Q2 = Q2 & " WHERE tblFile.MsID = " & msid
		rs.Open Q2,LoginDSN
		author = rs("Lname")
		rs.Close
		Response.Write("<p>Reviewer: <b>" & reviewer & "</b><br>")
		Response.Write("MR#: <b>" & seqno & "</b><br>")
		Response.Write("Date Assigned: <b>" & dateassigned & "</b><br>")
		Response.Write("Title: <b>" & title & "&nbsp;</b><br>")
		Response.Write("Author: <b>" & author & "&nbsp;</b><br>")	
		Response.Write("Type: <b>" & typename & "&nbsp;</b><br>")
		Response.Write("Version: <b>" & versionno & "&nbsp;</b><br>")	
		Response.Write("File: <b><a href='reviews/revfiles/" & filename & "'>" & filename & "</a></b></p>")								
		Response.Write("<br><br><input type='hidden' name='seqno' value=" & seqno & ">")
		%>
		<input type="hidden" name="submitted" value="true">
		<input type="submit" name="submit" value="Close Review">
	</form>
	<%
else
	'  Form has been submitted
	seqno = Cint(Request.Form("seqno"))
	' Close all the files that have the specified seqno.  There may be several.
	Q = " UPDATE tblFile SET Closed = " & true & " WHERE [SeqNo] = " & seqno
	rs.Open Q, LoginDSN,2,3
	set rs=nothing
	'Response.Write ("<p>Query: " & Q & "</p>")
	Response.Write("<p>The review has been closed.<br>")
end if
%>
<p><A href="statusboard5.asp?id=<%=loginid%>">Return to Status Board</a></p>
</body>
</html>
