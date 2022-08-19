<%@ Language=VBScript %>
<%option explicit%>
<% Server.ScriptTimeout = 1200  ' 20 minutes %>
<!-- #include file="../test2.asp"  -->
<!-- Note: the following variables were previously defined in test2.asp:  -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN  -->
<html>
<head>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<title>Acknowledge Upload</title>
</head>
<body>
<div align="center">
<%
'  Called from addfile.asp.  
'  The loginid allows us to identify the uploader uniquely.
dim diagnostic,errn
dim oFileUp
dim myfile,tempname,filename,filesize,ftype,typeid,disid,uploaddate,versionno,authornote
dim title,rs,Q,versionid,seqno,editornote,abstract,filepath,partyid,msid,disposition
diagnostic=false
loginid = Request.QueryString("id")
' Form data has been submitted
' Set this to the desired directory (must have read, write & delete permissions):
filepath = Server.Mappath ("/ORGS/asa/authors/msfiles/") 
' Instantiate FileUp if not diagnostic
if not diagnostic then
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	oFileUp.Path = filepath
	tempname = oFileUp.UserFileName
	filename = Mid(tempname, InstrRev(tempname, "\") + 1)
	' Allow FileUp to append a number to the file name in case a duplicate is already stored.
	oFileUp.CreateNewFile = True
	ftype = oFileUp.ContentType    ' based on file extension
	title = oFileUp.form("title")
	title = Replace(Trim(Left(title,80)),"'","''")
	'author = oFileUp.form("author")
	'author = Replace(Trim(Left(author,50)),"'","''")
	' For now we are assuming the the uploader is the only author. 
	typeid = Cint(oFileUp.form("typeID"))  '  Category or type of MS (standardized)
	disid = Cint(oFileUp.form("disID"))    '  Discipline ID (from list)
	versionno = oFileUp.form("versionno")  ' Version description (standardized)
	if IsNull(oFileUp.form("authornote")) then  '  Version notes from Author, optional
		authornote=" "
	else
		authornote=oFileUp.form("authornote")
	end if
	if IsNull(oFileUp.form("abstract")) then
		abstract = " "
	else
		abstract = oFileUp.form("abstract")  '  optional field for text abstract
	end if
else
	filename=Request.Form("MyFile")
	ftype=right(filename,3)
	title=Request.Form("title")
	typeid = Cint(Request.Form("typeid"))
	disid = Cint(Request.Form("disid"))
	versionno = Request.Form("versionno")
	authornote=Replace(Request.Form("authornote"),"'","''")
	abstract=Replace(Request.Form("abstract"),"'","''")
end if   ' not diagnostic
	Response.Write("<p>Path = " & filepath & "</p>")
	Response.Write("<p>Filename = " & filename & "</p>")
'  Check for errors in data entry
errn=0
if filename = "" then 
	errn=errn+1
	Response.Write("<br>File name is blank.")
end if
if title = "" then
	errn=errn+1
	Response.Write("<br>Title field is blank.")
end if
if typeid="" then
	errn=errn+1
	Response.Write("<br>Manuscript Category was not selected.")
end if
if disid="" then
	errn=errn+1
	Response.Write("<br>A Discipline was not selected.")
end if
if versionno="" then
	errn=errn+1
	Response.Write("<br>A Version/Part description was not selected.")
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
	if not diagnostic then
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
	else
		' Dummy data for testing locally
		filename="testfile.txt"
		filesize=100
	end if
	'  Data are valid; store the metadata in the database	
	uploaddate = Date()
	seqno=0          '  Sequence no. or MR# for editor to set later - integer format.
	disposition = 1  '  New submission code
	editornote=" "  '  For editor to use later.
	' Find PartyID of the user:
	set rs=Server.CreateObject("ADODB.Recordset")
	'  The fact that it was called from within the authors directory means that
	'  the user's role is 2 (author).  
	Q = "SELECT [PartyID],[Entrytime],[Role] FROM [tblParty] WHERE [Role]= 2 AND [Entrytime] = '" & loginid & "'"
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
	' Insert new data into table (autonumber increments MsID, I hope!):	
	Q = "INSERT INTO tblFile (MsTypeID, AuID, SeqNo, UploadDate, Filename, VersionNo,"
	Q = Q & "Title, FileSize, FileType, Disposition, AuthorNote, EditorNote, Closed, Abstract )"
	Q = Q & " VALUES (" & typeid & "," & partyid & "," & seqno & ",'" & uploaddate & "','"
	Q = Q & filename & "','" & versionno & "','" & title & "'," & filesize & ",'" & ftype & "'," 
	Q = Q & disposition & ",'" & authornote & "','" & editornote & "'," & false & ",'" & abstract & "')"
	if diagnostic then 
		Response.Write("<p>Q1 = " & Q & "</p>")
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q,LoginDSN,3,3
	' Open the File table and determine the last ID:	
	Q = "SELECT TOP 1 MsID FROM tblFile ORDER BY MsID DESC "
	rs.Open Q,LoginDSN
	msid = rs("MsID")
	rs.Close
	Response.Write("<p>Last MsID = " & msid & "</p>")
	'  Insert into the table that links to the Disciplines table:
	'  (This could be expanded to insert queries for disid1, disid2, etc.)
	Q = "INSERT INTO tblFileDis ( MsID, DisID ) "
	Q = Q & " VALUES (" & msid & "," & disid & ")"
	if diagnostic then 
		Response.Write("<p>Q2 = " & Q & "</p>")
	end if
	rs.Open Q,LoginDSN,3,3
	set rs = nothing
	Response.Write ("</p><p><b>Filename = " & filename & "<br>")
	Response.Write ("has been added to the database.</b></p>")
	%>
	<br>
	<br>
	<p><a href="addfile.asp?id=<%=loginid%>">Upload another file</a></p>
	<br>
	<p><a href="menu.asp?id=<%=loginid%>">Return to menu</a></p>
	<%
end if	
%>
</div>
</body>
</html>
