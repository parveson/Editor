<%@ Language=VBScript %>
<%option explicit%>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<link rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<title>Delete My File - for Reviewers</title>
</head>
<body>
<%
' Deletes one file that was selected by the user.
' Note: files are not deleted from the file system; this only deletes the data
' about the file in the database.  Another utility would be needed to 
' physically delete the files, or better -- archive them.  
dim diagnostic,rs,Q
dim msid,filename,title,filesize,author,uploaddate,versionno
diagnostic=false
msid = Request.QueryString("msid")
if not Request.Form("submitted") then
	Response.Write("<H3 align=center>Delete a File from the Database</H3>")
	set rs=Server.CreateObject("ADODB.Recordset")
	Q = "SELECT * FROM [tblFile] WHERE [tblFile].[MSID]=" & msid 
	rs.Open Q,LoginDSN,3,3
	filename=Trim(rs("Filename"))
	title=rs("Title")
	filesize=rs("Filesize")    ' number
	uploaddate=rs("UploadDate")
	versionno=rs("VersionNo")
	rs.Close
	Response.Write ("<p><b>File:  " & Filename & " <br><br><br><br>")
	%>
	<p>
	<table align=left border="1" cellpadding="4" cellspacing="0">
	<tr>
		<th align=center>Review File<br>Name</th>
		<th align=center>MS Title</th>
		<th align=center>Upload<br>Date</th>
		<th align=center>Version/Part</th>
	</tr>
	<%  
	Response.Write "<td><a href='../files/" & filename & "'><small>" & filename & "</small></a></td>"
	Response.Write "<td align=center><small>" & title & "</small></td>"
	Response.Write "<td align=center><small>" & uploaddate & "</small></td>"
	Response.Write "<td align=center><small>" & versionno & "</small></td></tr></table>"	
	%></p>
	<p><br><br><br>.
	<form method="post" action="deletefile.asp?id=<%=loginid%>&msid=<%=msid%>">
		<input type=hidden name=filename value="<%=filename%>">
		<input type=hidden name=submitted value="true">
		<input type="submit" value="OK - Delete">
	</form>
	</p>
	<p><a href="menu.asp?id=<%=loginid%>">NO - do not delete! </a></p>
	<%
else
	filename=Request.Form("filename")
	' Re-open database and delete file data:
	Q = "DELETE FROM [tblFile] WHERE [tblfile].[MSID] = " & msid
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q,LoginDSN,3,3
	Response.Write ("<p>The file <strong>" & filename & "</strong> has been deleted from the database.</strong></p>")
end if
set rs = nothing
%>
<p><a href="menu.asp?id=<%=loginid%>">Return to menu</a></p>
</body>
</html>
