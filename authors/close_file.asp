<%@Language="VBScript"%>
<%option explicit
Response.expires=30
'Response.addHeader "pragma","no-cache"
'Response.addHeader "cache-control","private"
'Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<link rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<title>Close a File</title>
<SCRIPT LANGUAGE="JavaScript" type="text/JavaScript">
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</SCRIPT>
</head>
<body>
<br>
<%
dim rs,Q,Q2
set rs = Server.CreateObject("ADODB.Recordset")	
dim msid,assignid,reviewer,seqno,dateassigned,title,typename,versionno
dim recordcount,author,filename,uploaddate,days,filetype,role,lname
msid = Request.QueryString("msid")
if not Request.Form("submitted") then
	'  Form has not been submitted yet
	%>
	<h3 align="center">Close a File</h3>
	<p>Note: Closing a file removes it from the list.  
	You may wish to close a file if you want to submit a revised version.</p> 
	<p><STRONG>Are you sure you want to close this file?</STRONG> 
	<form method="post" action="close_file.asp?id=<%=loginid%>">
	<%
	Q = "SELECT tblFile.MsID, tblFile.UploadDate, tblFile.SeqNo, tblFile.Filename, "
	Q = Q & " tblFile.FileType, tblFile.Title, tblFile.VersionNo, tblFile.MsTypeID, "
	Q = Q & " tblParty.Role, tblParty.Entrytime, tblParty.Lname, tblFile.Closed "
	Q = Q & " FROM tblParty INNER JOIN (tblMsType INNER JOIN tblFile ON tblMsType.TypeID = "
	Q = Q & " tblFile.MsTypeID) ON tblParty.PartyID = tblFile.AuID"
	Q = Q & " WHERE tblFile.MsID = " & msid
	' Select all records using GetRows:
	' Thanks to http://www.asp101.com/samples/db_getrows.asp
	dim cnnGetRows   ' ADO connection
	dim strDBPath    ' Path to our Access DB (*.mdb) file
	dim arrDBData    ' Array that we dump all the data into
	dim iRecFirst, iRecLast,I,J
	dim iFieldFirst, iFieldLast
	set cnnGetRows = Server.CreateObject("ADODB.Connection")
	cnnGetRows.Open LoginDSN
	set rs=cnnGetRows.Execute(Q)
	'Response.Write("<p align=center><b>Reviewer: " & rs("Fname") & " " & rs("Lname") & "</b></p>")
	' Display a table of the data in the array.
	if not rs.EOF then 
		'arrDBData = rs.GetRows(-1,0,Array("SeqNo","Title","TypeName","FileName","FileType","FileSize","UploadDate","VersionNo","Role"))
		arrDBData = rs.GetRows()  ' Very fast
		rs.Close
		cnnGetRows.Close
		set cnnGetRows=nothing			
		iRecFirst   = LBound(arrDBData, 2)
		iRecLast    = UBound(arrDBData, 2)
		iFieldFirst = LBound(arrDBData, 1)   '  = 0
		iFieldLast  = UBound(arrDBData, 1)
		' Table header:
		Response.Write("<table align=center bgcolor='#ffffff' border=1 cellspacing=0 width=650>")
		Response.Write("<th>Upload Date</th><th>MR#</th><th>File</th><th>Title</th><th>Version/Part</th><th>Author</th></tr>")
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
						 <td><A href="javascript:win('msfiles/<%=filename%>');"><small>View</A> 
						 <%
					else
						' Open download dialog box if any other type of file:
						Response.Write "<td class=small><a href='msfiles/" & filename & "'>View</a></td>"
					end if
				case 5 ' Title
					title = arrDBData(J,I)
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "&nbsp;</td>")
				case 6  ' Version/Part - convert to text desc. in function stage()
					versionno = arrDBData(J,I)
					Response.Write("<td align=center class=small>" & stage(versionno) & "&nbsp;</td>")
				case 7 ' MSTypeID
				case 8  ' Role - should be 2, for error checking only
					role = Cint(arrDBData(J,I))
					if not role = 2 then
						Response.Write("<p>Error -  Role = " & role & " J=" & J & " I=" & I & " </p>")
					end if			
				case 9  ' Entrytime
				case 10  ' Lname of author
					lname = arrDBData(J,I)
					Response.Write ("<td align=center class=small>" & lname & "&nbsp;</td>")
				case 11   ' Closed
				case else
					Response.Write("Error - too many fields in array: J=" & J & " I= " & I & " data=" & arrDBData(J,I))			
				end select
			Next ' J				
		Next ' I
		Response.Write("</table>")
		set rs = nothing
		if recordcount=0 then
			Response.Write ("<p>ERROR - Record not found.</p>")
		end if
	end if ' rs.EOF
		Response.Write("<br><br><input type='hidden' name='msid' value=" & msid & ">")
		%>
		<input type="hidden" name="submitted" value="true">
		<input type="submit" name="submit" value="Close File">
	</form>
	<%
else
	'  Form has been submitted
	msid = Cint(Request.Form("msid"))
	' Close the one file with this ID.
	Q = " UPDATE tblFile SET Closed = " & true & " WHERE [MsID] = " & msid
	rs.Open Q, LoginDSN,2,3
	set rs=nothing
	'Response.Write ("<p>Query: " & Q & "</p>")
	Response.Write("<p>The file has been closed.<br>")
end if

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
</body>
</html>
