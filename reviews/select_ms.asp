<%@ Language=VBScript %>
<%option explicit%>
<!-- #include file="../test2.asp"  -->
<html>
<head>
<title>Select a MS to Review</title>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<script language="Javascript" type="text/Javascript">
<!--
// Make a new popup window:
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=600,width=700,left=200,top=100,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
function printWindow() {
bV = parseInt(navigator.appVersion);
if (bV >= 4) window.print();
}
// -->
</script>
</head>
<body>
<h3 align="center">Select a Manuscript to Review</h3>
<P align=left>Below is a list of the files that the Editor has assigned to you.  
Click on the "View" link to select any file to download and review.  
Please note that several files may be related to a given MS -- they will all have the same MR#.  
Click on the "Select" link to submit a review of this file.</p>
<h5 align="center">Files Available for Review:</h5>
<%
dim diagnostic
diagnostic=false
dim rs,rs2,i,j,k,Q,Q2,count,ck,formvalue,closed
dim msid,assignid,versionid,refid,recordcount,versiontext,author,reviewer
dim days,assigndate,uploaddate,seqno,auid,role,partyid,reminderid
dim filename,lname,title,email,version,dateassigned,reviewdays,typename
dim cnnGetRows   ' ADO connection
dim strDBPath    ' Path to our Access DB (*.mdb) file
dim arrDBData    ' Array that we dump all the data into
dim iRecFirst, iRecLast
dim iFieldFirst, iFieldLast
'  Select a MS to review:
'  First identify the data on the reviewer who has logged in to this page
Q = "SELECT tblParty.PartyID, tblParty.Entrytime, tblParty.Email FROM tblParty "
Q = Q & " WHERE tblParty.Entrytime = '" & loginid & "'"
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open Q,LoginDSN
partyid = rs("PartyID")
rs.Close
if diagnostic then 
	Response.Write("<p>ID = " & partyid & "</p>")
end if
%>
<table border="1" cellspacing="0" cellpadding="2" ALIGN="CENTER">
<tr>
	<th>File</th>
	<th>MR#</th>
	<th>Date Assigned</th>
	<th>MS Title</th>
	<th>MS Type</th>
	<th>MS Version</th>
	<th>MS Author</th>
	<th>Review</th>
</tr>
<% 
' Now display all the files to which the editor has assigned this reviewer: 
Q = "SELECT tblFile.FileName, tblParty.Lname, tblFile.SeqNo, tblAssignment.DateAssigned, "
Q = Q & " tblFile.Title, tblMsType.TypeName, tblFile.VersionNo, tblFile.MsID, "
Q = Q & " tblParty.PartyID, tblParty.Role, tblAssignment.AssignID, tblFile.Closed  "
Q = Q & " FROM tblParty INNER JOIN (tblMsType INNER JOIN (tblFile INNER JOIN tblAssignment "
Q = Q & " ON tblFile.MsID = tblAssignment.MsID) ON tblMsType.TypeID = tblFile.MsTypeID) ON "
Q = Q & " tblParty.PartyID = tblAssignment.PartyID "
Q = Q & " WHERE (tblFile.SeqNo > 0 AND tblFile.Closed = " & false 
Q = Q & " AND tblParty.PartyID = " & partyid & ")"
Q = Q & " ORDER BY tblFile.SeqNo, tblAssignment.DateAssigned DESC"
if diagnostic then 
	Response.Write("<p>Q = " & Q & "</p>")
else
	set cnnGetRows = Server.CreateObject("ADODB.Connection")
	cnnGetRows.Open LoginDSN
	set rs=cnnGetRows.Execute(Q)
	if not rs.EOF then
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
		set rs2=Server.CreateObject("ADODB.Recordset")
		recordcount=0
		For I = iRecFirst To iRecLast  		' A table row for each record = I
			recordcount = recordcount + 1
			' Columns: Loop through the fields (first dimension of the array = J)
			For J = 0 To iFieldLast
				select case J
					case 0	' Author's file name & link (HTMLEncode)
						filename = Server.HTMLEncode(arrDBData(J,I))
						Response.Write("<tr><td align=center><a href='../authors/msfiles/" &  filename & "'>View</a></td>")
					case 1  ' AuID - no display
						auid=arrDBData(J,I)
					case 2	' SeqNo = MR#
						Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")					
					case 3	' Upload date & color code
						uploaddate = arrDBData(J,I)
						'Response.Write("<td align=center class=small>" & uploaddate & "</td>")
						Response.Write("<td align=center class=small bgcolor='" & flag(uploaddate) & "'>" & uploaddate & "</td>")
					case 4  ' VersionNo
						Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
					case 5	' Title
						Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
					case 6 ' MS Type Name
						typename = stage(arrDBData(J,I))
						Response.Write("<td align=center class=small>" & typename & "</td>")
					case 7  ' Author last name (replaces PartyID, not shown)
						msid=arrDBData(J,I)
						'  Query: FOR EACH msid, get Party Lname etc. where msid = xxx
						Q2 = "SELECT tblFile.MsID, tblFile.AuID, tblParty.PartyID, tblParty.Lname "
						Q2 = Q2 & " FROM tblParty INNER JOIN tblFile ON tblParty.PartyID = tblFile.AuID "
						Q2 = Q2 & " WHERE tblFile.MsID = " & msid
						rs2.Open Q2,LoginDSN
						author = rs2("Lname")
						rs2.Close
						Response.Write("<td align=center class=small>" & author & "</td>")
					case 8	' PartyID - not used
					case 9  '  Role
						role = arrDBData(J,I)
						if not role=3 then
							Response.Write("<td align=center class=small>ERROR: Role=" & role & "</td>")
						end if
					case 10  ' AssignID - not used
					case 11  ' Closed - not used
						Response.Write("<td align=center><a href='addfile.asp?id=" & loginid & "&msid=" &  msid & "'>Select</a></td></tr>")
					case else
						Response.Write("ERROR - no. of data values exceeds table size.")
				end select
			Next ' J
		Next   ' I
		Response.Write("</tr></table><p>" & recordcount & " record(s) found.</p>")
	else
		Response.Write("</table><p>No records found.</p>")
	end if ' no records
	set rs2=nothing
end if ' diagnostic

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

function flag(dateassigned)
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

<p><A href="menu.asp?id=<%=loginid%>">Return to menu</a></p>
<br>
<p><A href="default.htm">Return to login page</A></p>
</body>
</html>
