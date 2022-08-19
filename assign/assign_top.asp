<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = false 
%>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<%
'  assign/assign_top.asp  - called from statusboard.asp 
'  Top part of the frameset pages - shows data on one MS.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML lang="en">
<head>
   <title>Assign Reviewers</title>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
</head>
<body><p class=small>Administration</p>
<br>
<h3>Assign Reviewers to a File</h3>
<%  
dim rs,i,j,k,Q,Q1,count,ck,refid,msid
dim diagnostic
diagnostic=false
dim formvalue,role,email,dis,done,pending,duration,note
dim filename,title,version,remail,dateassigned,reviewdays
dim lname,fname,fullname,entrydate,discipline,auid
dim cnnGetRows   ' ADO connection
dim strDBPath    ' Path to our Access DB (*.mdb) file
dim arrDBData    ' Array that we dump all the data into
dim assignid,recordcount,nextmr,versiontext
dim days,assigndate,uploaddate,seqno
dim iRecFirst, iRecLast
dim iFieldFirst, iFieldLast
msid = Request.QueryString("msid")    ' Selected file id from Status Board
Response.Write("<h5>Selected File:</h5>")
' Show the record with the selected MsID:
' Thanks to http://www.asp101.com/samples/db_getrows.asp
Q = "SELECT tblFile.MsID, tblFile.AuID, tblFile.SeqNo, tblFile.UploadDate, "
Q = Q & " tblFile.Filename, tblFile.VersionNo, tblFile.Title, tblFile.Disposition, "
Q = Q & " tblFile.AuthorNote, tblFile.Abstract, tblMsType.TypeName, tblDiscipline.DisName, "
Q = Q & " tblParty.Lname, tblParty.Role "
Q = Q & " FROM tblParty INNER JOIN (tblMsType INNER JOIN (tblFile INNER JOIN (tblDiscipline"
Q = Q & " INNER JOIN tblFileDis ON tblDiscipline.DisID = tblFileDis.DisID) ON tblFile.MsID"
Q = Q & " = tblFileDis.MsID) ON tblMsType.TypeID = tblFile.MsTypeID) ON tblParty.PartyID"
Q = Q & " = tblFile.AuID "
Q1 = Q & "WHERE tblFile.MsID = " & msid
Q1 = Q1 & " ORDER BY tblFile.UploadDate"
if not diagnostic then 
	%>
	<table align="center" border="1" cellpadding="2" cellspacing="0">
	<tr>
	<th>MR#</th>
	<th>Upload Date</th>
	<th>File</th>
	<th>Version</th>
	<th>Title</th>
	<th>Type</th>
	<th>Discipline</th>
	<th>Author</th>
	</tr>
	<% 
	set cnnGetRows = Server.CreateObject("ADODB.Connection")
	cnnGetRows.Open LoginDSN
	set rs=cnnGetRows.Execute(Q1)
	arrDBData = rs.GetRows()  ' Very fast
	rs.Close
	set rs=nothing
	cnnGetRows.Close
	set cnnGetRows=nothing
	iRecFirst   = LBound(arrDBData, 2)
	iRecLast    = UBound(arrDBData, 2)
	iFieldFirst = LBound(arrDBData, 1)   '  = 0
	iFieldLast  = UBound(arrDBData, 1)
	' Display a table of the data in the array.
	' First loop through the records (second dimension of the array)
	recordcount=0
	For I = iRecFirst To iRecLast  		' A table row for each record = I
		recordcount = recordcount + 1
		' Columns: Loop through the fields (first dimension of the array = J)
		For J = 0 To iFieldLast
			select case J
				case 0  ' MsID - display radio button
					msid=arrDBData(J,I)
				case 1  ' AuID - no display
					auid=arrDBData(J,I)
				case 2	' SeqNo = MR#
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
				case 3	' Upload date & color code
					uploaddate = arrDBData(J,I)
					Response.Write("<td align=center class=small bgcolor='" & flag(uploaddate) & "'>" & uploaddate & "</td>")
				case 4	' File name & link (HTMLEncode)
					filename = Server.HTMLEncode(arrDBData(J,I))
					Response.Write("<td align=center><a href='../authors/msfiles/" &  filename & "'>View</a></td>")
				case 5  ' VersionNo
					versiontext=stage(arrDBData(J,I))
					Response.Write("<td align=center class=small>" & versiontext & "</td>")
				case 6	' Title
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
				case 7  ' Disposition
					'
				case 8  ' Author Note
					'
				case 9  ' Abstract
					'
				case 10 ' MS Type Name
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
				case 11	' Discipline
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")				
				case 12 ' Author last name
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td></tr>")				
				case 13	' Role (should be 2)
					role=arrDBData(J,I)
					if role<>2 then
						Response.Write("ERROR - Role <> 2!")
					end if
				case else
					Response.Write("ERROR - no. of data values exceeds table size.")
			end select
		Next ' J
	Next   ' I
else	Response.Write("<p>Top frame, MSID = " & msid & "</p>")end if  ' diagnosticfunction stage(versionno)
	' Convert the numeric code into its text description
	select case versionno
		case 1
			stage="Main manuscript"
		case 2
			stage="Major revision"
		case 3
			stage="Minor revisions"
		case 4
			stage="Figure or table"
		case 5
			stage="Final"
		case 6
			stage="Other"
	end select
end function

function flag(dateassigned)
	'  Generate a color field based on date
	days = DateDiff("d",dateassigned,Date)
	flag="silver"
	if days>=0 and days <=30 then
		flag="lightgreen"
	elseif days >30 and days<=60 then
		flag="yellow"
	else
		flag="pink"
	end if
end function

%>
</body>
</html>
