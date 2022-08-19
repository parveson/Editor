<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = false 
%>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional //EN">
<html>
<head>
<title>Apply MR#</title>
<LINK rel="stylesheet" type="text/css" href="StyleSheet1.css">
<script language="Javascript" type="text/javascript">
<!-- // Begin
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.unt.edu/map/directions.htm');">directions</a>
//  End -->
</script>
</HEAD>
<BODY>
<h3>Apply MR# to a File</h3>
<% 
' Apply.asp
' Called from the Status Board, where one new file was selected.
' This page lists all existing manuscripts in one long table, 
' allowing the editor to decide which number to apply to the selected
' new file.  
dim rs,i,j,k,Q,Q1,Q2,Q3,Q4,count,ck,refid,msid
dim diagnostic
diagnostic=false
dim formvalue,occ,country,domain,role,email,dis,done,pending,duration,note
dim filename,title,email1,version,remail,dateassigned,reviewdays
dim lname,fname,fullname,entrydate,discipline,auid
dim cnnGetRows   ' ADO connection
dim strDBPath    ' Path to our Access DB (*.mdb) file
dim arrDBData    ' Array that we dump all the data into
dim assignid,recordcount,nextmr,versiontext
dim days,flag,assigndate,uploaddate,seqno
dim iRecFirst, iRecLast
dim iFieldFirst, iFieldLast
msid = Request.QueryString("msid")    ' Selected file id from Status Board
Response.Write("<h5>Selected New File:</h5>")
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
set rs=server.CreateObject("ADODB.Recordset")
rs.Open Q1,LoginDSN
if rs.EOF then
	rs.Close
	Response.Write("<p>Error - MSID not found.</p>")
else
	Response.Write("<p>MsID: <b>" & rs("msid") & "</b><br>")
	Response.Write("Upload Date: <b>" & rs("uploaddate") & "</b><br>")
	Response.Write("File Name: <b>" & rs("filename") & "</b><br>")
	Response.Write("Version No.: <b>" & rs("versionno") & "</b><br>")
	Response.Write("Title: <b>" & rs("title") & "</b><br>")
	Response.Write("Disposition: <b>" & rs("disposition") & "</b><br>")
	Response.Write("Author's Note: <b>" & rs("authornote") & "</b><br>")
	Response.Write("Abstract: <b>" & rs("abstract") & "</b><br>")
	Response.Write("MS Type: <b>" & rs("typename") & "</b><br>")
	Response.Write("Discipline: <b>" & rs("disname") & "</b><br>")
	Response.Write("Author: <b>" & rs("Lname") & "</b><br>")	
	rs.Close
end if	
if not Request.Form("submitted") then
	' Find the highest value of SeqNo in the database:
	Q2 = "SELECT TOP 1 [SeqNo] FROM [tblFile] ORDER BY 1 DESC"
	rs.Open Q2,LoginDSN
	nextmr=rs("SeqNo") + 1
	if not rs.EOF then
		rs.Close	
		%>
		<form method=post action="apply.asp?id=<%=loginid%>&msid=<%=msid%>">
		<p><b>The next available MR# is <%=nextmr%></b><br>
		<input type=hidden name=seqno value=<%=nextmr%>>
		<input type=hidden name="submitted" value="true">
		<input type=submit value="Apply this MR#">
		</form>		
		<hr>
		<h4>Files with MR#:</h4>
		<form method=post action="apply.asp?id=<%=loginid%>&msid=<%=msid%>">		
		<p><b>If files related to the above already are stored, select one of the related
		files below:</b></p>
		<table align="center" border="1" cellpadding="2" cellspacing="0">
		<tr>
		<th bgcolor="#fffacd">Select</th>
		<th>MR#</th>
		<th>Date</th>
		<th>File</th>
		<th>Version</th>
		<th>Title</th>
		<th>Type</th>
		<th>Discipline</th>
		<th>Author</th>
		</tr>
		<% 
		Q3 = Q & "WHERE tblFile.SeqNo > 0 ORDER BY tblFile.SeqNo DESC"
		set cnnGetRows = Server.CreateObject("ADODB.Connection")
		cnnGetRows.Open LoginDSN
		set rs=cnnGetRows.Execute(Q3)
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
						seqno = arrDBData(J,I)
						Response.Write("<tr><td bgcolor='#fffacd'><input type=radio name=seqno value=" & seqno & "></td>")
						Response.Write("<td class=small>" & seqno & "</td>")
					case 3	' Upload date & color code
						uploaddate = arrDBData(J,I)
						days = DateDiff("d",uploaddate,Date)
						flag="silver"
						if days>=0 and days <=30 then
							flag="lightgreen"
						elseif days >30 and days<=60 then
							flag="yellow"
						else
							flag="pink"
						end if
						Response.Write("<td align=center class=small bgcolor='" & flag & "'>" & uploaddate & "</td>")
					case 4	' File name & link (HTMLEncode)
						filename = Server.HTMLEncode(arrDBData(J,I))
						Response.Write("<td align=center><a href='authors/msfiles/" &  filename & "'>View</a></td>")
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
		Response.Write("</table><p>" & recordcount & " record(s) found.</p>")
		Response.Write("<input type=hidden name=submitted value=true>")
		Response.Write("<input type=submit value='Apply the selected MR#'>")
		Response.Write("&nbsp;&nbsp;&nbsp;<input type=reset value=Reset></form>")
	else
		Response.Write("</table><p>ERROR - MSID = " & msid & " not found.</p>")
	end if ' no records
else
	'  Form data has been submitted
	' Do an Update query to update one of the SeqNos.
	msid = Request.QueryString("msid")
	seqno = Request.Form("seqno")
	Q4 = "UPDATE tblFile SET SeqNo = " & seqno & " WHERE [MsID] = " & msid
	rs.Open Q4, LoginDSN,2,3
	set rs=nothing
	Response.Write("<p><b>The MR# has been set to <b>" & seqno & ".</b></p>")
end if  ' submitted

function stage(versionno)
	' Converts the numeric code into its text description
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

%>
<p><A href="statusboard5.asp?id=<%=loginid%>">Return to status board</A></p>

</BODY>
</HTML>
