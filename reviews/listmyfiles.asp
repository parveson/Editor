<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<HTML>
<HEAD>
<TITLE>List My Files</TITLE>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<SCRIPT LANGUAGE="JavaScript" type="text/JavaScript">
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</SCRIPT>
</HEAD>
<BODY>
<h4 align=center>List My Review Files</h4>
<%
dim diagnostic,email
diagnostic=false
dim title,seqno,typename,filename,filesize,versionno,uploaddate,role,filetype
dim quot,recordcount,versionnote,versiontext,msid,fileid,days,fname,lname
role=3   '  Reviewers - numeric value
dim rs,I,J,Q  ' other variables are defined in test2.asp
Q = "SELECT * FROM tblParty WHERE Entrytime = '" & loginid & "'"
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open Q,LoginDSN
Response.Write("<p align=center>Author: <b>" & rs("Fname") & " " & rs("Lname") & "</b></p>")
rs.Close
Q = "SELECT tblFile.MsID, tblFile.UploadDate, tblFile.SeqNo, tblFile.Filename, "
Q = Q & " tblFile.FileType, tblFile.Title, tblFile.VersionNo, tblFile.MsTypeID, "
Q = Q & " tblParty.Role, tblParty.Entrytime, tblParty.Lname, tblFile.Closed "
Q = Q & " FROM tblParty INNER JOIN (tblMsType INNER JOIN tblFile ON tblMsType.TypeID = "
Q = Q & " tblFile.MsTypeID) ON tblParty.PartyID = tblFile.AuID"
Q = Q & " WHERE tblFile.Closed = " & false & " AND tblParty.Role = 3 AND tblParty.Entrytime = '" & loginid & "'"
' Select all records using GetRows:
' Thanks to http://www.asp101.com/samples/db_getrows.asp
dim cnnGetRows   ' ADO connection
dim strDBPath    ' Path to our Access DB (*.mdb) file
dim arrDBData    ' Array that we dump all the data into
dim iRecFirst, iRecLast
dim iFieldFirst, iFieldLast
set cnnGetRows = Server.CreateObject("ADODB.Connection")
cnnGetRows.Open LoginDSN
set rs=cnnGetRows.Execute(Q)
if not rs.EOF then
	' Response.Write("<p align=center><b>Reviewer: " & rs("Fname") & " " & rs("Lname") & "</b></p>")
	' Display a table of the data in the array.
	'arrDBData = rs.GetRows(-1,0,Array("SeqNo","Title","TypeName","FileName","FileType","FileSize","UploadDate","VersionNo","Role"))
	arrDBData = rs.GetRows()  ' Very fast
	rs.Close
	cnnGetRows.Close
	set cnnGetRows=nothing			
	iRecFirst   = LBound(arrDBData, 2)
	iRecLast    = UBound(arrDBData, 2)
	iFieldFirst = LBound(arrDBData, 1)   '  = 0
	iFieldLast  = UBound(arrDBData, 1)
	Response.Write("<p>Note: After a review is completed, you will be notified by the Editor, and you may ")
	Response.Write("close the file.  This removes the file from the list.  You may also close a file if ")
	Response.Write("you wish to submit a revised version.</p>")
	' Table header:
	Response.Write("<table align=center bgcolor='#ffffff' border=1 cellspacing=0 width=650>")
	Response.Write("<th>Upload Date</th><th>MR#</th><th>File</th><th>Title</th><th>Version/Part</th><th>Author</th><th>Close?</th></tr>")
	recordcount=0
	For I = iRecFirst To iRecLast
		recordcount = recordcount + 1
		' A table row for each record					
		' Columns: Loop through the fields (first dimension of the array)
		For J = 0 To iFieldLast
			' Allow for special treatment of each field:
			select case J
			case 0 ' MsID
				msid = arrDBData(J,I)
			case 1  ' Upload date
				uploaddate = arrDBData(J,I)
				Response.Write("<td align=center class=small bgcolor='" & flag(uploaddate) & "'>" & uploaddate & "</td>")
			case 2 ' SeqNo or MR#
				seqno = arrDBData(J,I)
				Response.Write("<td align=center class=small>" & seqno & "</td>")	
			case 3  ' Filename
				filename = arrDBData(J,I)
			case 4  ' File Type
				filetype = arrDBData(J,I)
				if filetype = "text/plain" or filetype = "text/html" or filetype = "text/xml" then
					 ' Place file in a new browser window if it is text or html:
					 %>
					 <td><A href="javascript:win('revfiles/<%=filename%>');">View</A> 
					 <%
				else
					' Open download dialog box if any other type of file:
					Response.Write "<td class=small><a href='revfiles/" & filename & "'>View</a></td>"
				end if
			case 5 ' Title
				title = arrDBData(J,I)
				Response.Write("<td align=center class=small>" & arrDBData(J,I) & "&nbsp;</td>")
			case 6  ' Version/Part - convert to text desc. in function stage()
				versionno = arrDBData(J,I)
				versiontext = stage(versionno)
				Response.Write("<td align=center class=small>" & versiontext & "&nbsp;</td>")
			case 7 ' MSTypeID (should be 9 - for error checking only)
				if not arrDBData(J,I)= 9 then
					Response.Write("<p>ERROR - MsType not Review")
				end if
			case 8  ' Role - should be 3, for error checking only
				role = arrDBData(J,I)
				if not role = 3 then
					Response.Write("<H5>Error -  Role = " & role & " J=" & J & " I=" & I & " </H5>")
				end if			
			case 9  ' Entrytime
			case 10  ' Lname of reviewer
				lname = arrDBData(J,I)
				Response.Write ("<td align=center class=small>" & lname & "&nbsp;</td>")
			case 11   ' Closed
			case else
				Response.Write("Error - too many fields in array: J=" & J & " I= " & I & " data=" & arrDBData(J,I))			
			end select
		Next ' J				
		Response.Write("<td align=center><a href='close_file.asp?id=" & loginid & "&msid=" & msid & "'>Close</a></td></tr>") 
	Next ' I
	Response.Write("</table>")
	set rs = nothing
	select case recordcount
		case 0
			Response.Write("<p>No files found.</p>")
		case 1
			Response.Write("<p>One file found.</p>")
		case else
			Response.Write("<p>" & recordcount & " files found.</p>")
	end select
	
else
	Response.Write("<p>No records found.</p>")
end if ' rs.EOF

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
<p><A href="menu.asp?id=<%=loginid%>">Return to menu page</A></p>
</BODY>
</HTML>


