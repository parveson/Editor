<%@ Language=VBScript %>
<%option explicit%>
<% Server.ScriptTimeout = 1200  ' 20 minutes %>
<!-- #include file="../test2.asp"  -->
<!-- Note: the following variables were previously defined in test2.asp:  -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN  -->
<html>
<head>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<title>Acknowledge Review Upload</title>
</head>
<body>
<div align="center">
<%
'  Called from addfile.asp.  
'  The loginid allows us to identify the uploader uniquely.
dim errn
dim oFileUp
dim myfile,tempname,filename,filesize,ftype,typeid,disid,uploaddate,versionno
dim title,rs,Q,versionid,seqno,editornote,authornote,filepath,partyid,msid,disposition
loginid = Request.QueryString("id")
' Form data has been submitted
' Set this to the desired directory (must have read, write & delete permissions):
filepath = Server.Mappath ("/ORGS/asa/reviews/revfiles/") 
' Instantiate FileUp 
Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
oFileUp.Path = filepath
tempname = oFileUp.UserFileName
filename = Mid(tempname, InstrRev(tempname, "\") + 1)
' Allow FileUp to append a number to the file name in case a duplicate is already stored.
oFileUp.CreateNewFile = True
ftype = oFileUp.ContentType    ' based on file extension
versionno = oFileUp.form("versionno")
seqno = oFileUp.form("seqno")
title = oFileUp.form("title")
disposition = oFileUp.form("disposition") ' =1; not used yet
typeid = Cint(oFileUp.form("typeID"))  '  Category or type of MS (standardized)
if IsNull(oFileUp.form("authornote")) then   ' optional note for author or reviewer
	authornote = " "
else
	authornote = oFileUp.form("authornote")  
end if
Response.Write("<p>Path = " & filepath & "</p>")
Response.Write("<p>Filename = " & filename & "</p>")
'  Check for errors in data entry
errn=0
if filename = "" then 
	errn=errn+1
	Response.Write("<br>File name is blank.")
end if
' If any errors, return to form
if errn>0 then
	Response.Write("</p><p><b>There are missing or invalid entries in your form.</b><br>")
	Response.Write("<p><b>Please <a href='Javascript:history.back();'>TRY AGAIN</a></b></p>")
	Response.Write("<br />")
	Response.Write("<br />")
	Response.Write("<br />")
' Else store data:
else
	filesize = oFileUp.TotalBytes
	' KEEP FILE SIZE IN BYTES NOT KB 
	'filesize = filesize/1024   ' file size in KB - NOTE -- MAX filesize = 32,768 KB
	oFileUp.Save ' Save the file in the /msfiles directory  
	' NOTE: the /msfiles directory permissions must be set to write & delete.
	' Find the real file name stored, in case it was changed by FileUp:
	dim strServerName,strFileName
	strServerName = oFileUp.Form("MyFile").ServerName
	filename = Mid(strServerName, InstrRev(strServerName, "\") + 1) 
	set oFileUp = nothing
'  Data are valid; store the metadata in the database	
	uploaddate = Date()
	disposition = 1  '  New submission code
	editornote=" "  '  For editor to use later.
	' Find PartyID of the user:
	set rs=Server.CreateObject("ADODB.Recordset")
	'  The fact that it was called from within the Reviewers directory means that
	'  the user's role is 3 (Reviewer).  
	Q = "SELECT [PartyID],[Entrytime],[Role] FROM [tblParty] WHERE [Role]= 3 AND [Entrytime] = '" & loginid & "'"
	rs.Open Q,LoginDSN
	if rs.EOF then
		rs.Close
		set rs=nothing
		'  A match was not found.  Login failed.
		Response.Redirect "default.htm?no_match3"
	else
		partyid = rs("PartyID")
		rs.Close
	end if	
	' Insert new data into table (autonumber increments MsID):	
	Q = "INSERT INTO tblFile (MsTypeID, AuID, SeqNo, UploadDate, Filename, "
	Q = Q & "VersionNo, Title, FileSize, FileType, Disposition, AuthorNote, Closed )"
	Q = Q & " VALUES (" & typeid & "," & partyid & "," & seqno & ",'" & uploaddate & "','"
	Q = Q & filename & "','" & versionno & "','" & title &"'," & filesize & ",'"
	Q = Q & ftype & "'," & disposition & ",'" & authornote & "'," & false & ")"
	Response.Write("<p>Q1 = " & Q & "</p>")
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q,LoginDSN,3,3
	' Open the File table and determine the last ID:	
	Q = "SELECT TOP 1 MsID FROM tblFile ORDER BY MsID DESC "
	rs.Open Q,LoginDSN
	msid = rs("MsID")
	rs.Close
	Response.Write("<p>New File ID = " & msid & "</p>")
	set rs = nothing
	Response.Write ("</p><p><b>Review file name = " & filename & "<br>")
	Response.Write ("has been added to the database.</b></p>")
	%>
	<br>
	<br>
	<p><a href="select_ms.asp?id=<%=loginid%>">Select another file to review</a></p>
	<br>
	<%
end if	
%>
<p><a href="menu.asp?id=<%=loginid%>">Return to menu</a></p>
</div>
</body>
</html>
