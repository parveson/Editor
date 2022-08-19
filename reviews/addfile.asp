<%@ Language=VBScript %>
<%option explicit
Response.Buffer = true %>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Submit a Review</title>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
</head>
<body>
<h3 align="center">Submit a Review</h3>
<p>You have selected the following file for which to submit a review:</p>
<%
' Called by select_ms.asp
dim diagnostic,msid
msid = Request.QueryString("msid") ' MsID for selected file to review (unique)
diagnostic=false
Server.ScriptTimeout=1200      ' uploads can take up to 1200 sec. or 20 minutes.
'  IF YOU CHANGE THIS PLEASE CHANGE THE HTML BELOW AND CHANGE TIMEOUT ON ackupload.asp
dim rs,rs2,i,j,k,Q,Q2,count,ck,formvalue,closed,versionno,disposition
dim assignid,versionid,refid,recordcount,versiontext,author,reviewer
dim days,assigndate,uploaddate,seqno,auid,role,partyid,reminderid
dim filename,lname,title,email,version,dateassigned,reviewdays,typename
dim cnnGetRows   ' ADO connection
dim strDBPath    ' Path to our Access DB (*.mdb) file
dim arrDBData    ' Array that we dump all the data into,
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
%>
<table border="1" cellspacing="0" cellpadding="2" ALIGN="CENTER">
<tr>
	<th>File</th>
	<th>MR#</th>
	<th>Date Assigned</th>
	<th>Title</th>
	<th>MS Type</th>
	<th>Version</th>
	<th>Author</th>
</tr>
<% 
' Display data about the file to be reviewed: 
Q = "SELECT tblFile.FileName, tblParty.Lname, tblFile.SeqNo, tblAssignment.DateAssigned, "
Q = Q & " tblFile.Title, tblFile.VersionNo, tblMsType.TypeName, tblFile.MsID, "
Q = Q & " tblParty.PartyID, tblParty.Role, tblAssignment.AssignID, tblFile.Closed  "
Q = Q & " FROM tblParty INNER JOIN (tblMsType INNER JOIN (tblFile INNER JOIN tblAssignment "
Q = Q & " ON tblFile.MsID = tblAssignment.MsID) ON tblMsType.TypeID = tblFile.MsTypeID) ON "
Q = Q & " tblParty.PartyID = tblAssignment.PartyID "
Q = Q & " WHERE tblFile.MsID = " & msid & " AND tblParty.PartyID = " & partyid
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
		'  Note: this is slight overkill, since we only get one row.
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
					case 0	' File name & link (HTMLEncode)
						filename = Server.HTMLEncode(arrDBData(J,I))
						Response.Write("<tr><td align=center><a href='../authors/msfiles/" &  filename & "'>View</a></td>")
					case 1  ' AuID - no display
						auid=arrDBData(J,I)
					case 2	' SeqNo = MR#
						seqno =  arrDBData(J,I)
						Response.Write("<td align=center class=small>" & seqno & "</td>")					
					case 3	' Upload date & color code
						uploaddate = arrDBData(J,I)
						'Response.Write("<td align=center class=small>" & uploaddate & "</td>")
						Response.Write("<td align=center class=small bgcolor='" & flag(uploaddate) & "'>" & uploaddate & "</td>")
					case 4	' Title
						title = arrDBData(J,I)
						Response.Write("<td align=center class=small>" & title & "</td>")
					case 5  ' VersionNo
						versionno = arrDBData(J,I)
						versiontext = stage(versionno)
						Response.Write("<td align=center class=small>" & versiontext & "</td>")
					case 6 ' MS Type Name
						typename = arrDBData(J,I)
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
						'Response.Write("<td align=center><a href='addfile.asp?id=" & loginid & "&msid=" &  msid & "'>Select</a></td>")
					case else
						Response.Write("ERROR - no. of data values exceeds table size.")
				end select
			Next ' J
		Next   ' I
		Response.Write("</tr></table>")
	else
		Response.Write("</tr></table><p>No records found.</p>")
	end if ' no records
	set rs2=nothing
end if ' diagnostic

function stage(versionno)
	' Converts the numeric code into its text description
	select case versionno
		case 1
			stage="Initial submission"
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
<br>
<hr>
<p><b>Instructions:</b>
<br>1. Write your review or download and markup a copy of the original file shown above.  
<br>2. Then you may submit the review by uploading the file and by filling in 
(optional) comments in the form below.  Any common file format is acceptable. 
Maximum file size is 32 MB.  Large files can take a long time to load; a maximum of 20 minutes is allowed.
If your upload lasts longer than this it will timeout. If you have any problems please 
<A href="../contact.asp">contact us.</A>&nbsp;&nbsp;
<br>
</P>
<form enctype="multipart/form-data" method="post" action="ackupload.asp?id=<%=loginid%>">
	<table WIDTH="95%" border=0>
	<tr>
	   <td align="right"><STRONG>File:</STRONG></td>
	   <td ALIGN="left"><input TYPE="file" NAME="MyFile" size="60"></td>
	</tr>
    <tr>
		<td align="right"><STRONG>Add Comments:</STRONG> <BR>
			<FONT size=1>(optional - 1500 chars. max.)</FONT>     
		</td>
		<td><TEXTAREA name="authornote" rows=8 cols=60></TEXTAREA>
		</td>
	</tr>
	<tr><td align="right"><br><br>
		<input type=hidden name="msid" value=<%=msid%>>
		<input type=hidden name="seqno" value=<%=seqno%>>
		<input type=hidden name="title" value=<%=title%>>
		<input type=hidden name="versionno" value=<%=versionno%>>
		<input type=hidden name="disposition" value=1>
		<input type=hidden name="typeid" value=9>  <!-- MS Type = Review -->
	   <input TYPE="submit" VALUE="Upload File"></td>
	   <td valign="bottom"><br>&nbsp;&nbsp;<STRONG>Click once; then wait for upload 
      to finish.</STRONG>       
	   <br>&nbsp;&nbsp;Then you will be directed to the next page. </td>
	</tr>
	</table>
	</form>
	<p><A href="menu.asp?id=<%=loginid%>">Return to menu</a></p>
</body>
</html>
