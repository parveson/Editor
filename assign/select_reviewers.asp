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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional //EN">
<html>
<head>
<title>Select Reviewers</title>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
</HEAD>
<BODY>
<h4 align=center>List of All Reviewers</h4>
<%
'  Select reviewers for a file
dim msid,email,email2  ' File to which this selection applies
msid = Request.QueryString("msid")
email = Request.QueryString("email")
dim diagnostic
diagnostic=false
dim prefix,fname,lname,phone,fullname,pw,entrydate,expertise,notes
dim refdate,quot,refid,role,count,recordcount,partyid,notify(10)
dim rs,I,J,Q,Q1,Q2,Q3  ' other variables are defined in test2.asp
if not Request.Form("submitted") then
	Q = "SELECT tblParty.PartyID, tblParty.Prefix, tblParty.Fname, tblParty.Lname, "
	Q = Q & " tblParty.Expertise, tblParty.Notes, tblParty.Role, tblParty.Email FROM tblParty "
	Q = Q & " WHERE [tblParty.Role]= 3 ORDER BY [tblParty.Lname]"
	' Select records using GetRows:
	' Thanks to http://www.asp101.com/samples/db_getrows.asp
	dim cnnGetRows   ' ADO connection
	dim strDBPath    ' Path to our Access DB (*.mdb) file
	dim arrDBData    ' Array that we dump all the data into
	dim iRecFirst, iRecLast
	dim iFieldFirst, iFieldLast
	set cnnGetRows = Server.CreateObject("ADODB.Connection")
	cnnGetRows.Open LoginDSN
	set rs=cnnGetRows.Execute(Q)
	arrDBData = rs.GetRows()
	rs.Close
	cnnGetRows.Close
	set cnnGetRows=nothing			
	iRecFirst   = LBound(arrDBData, 2)
	iRecLast    = UBound(arrDBData, 2)
	iFieldFirst = LBound(arrDBData, 1)   '  = 0
	iFieldLast  = UBound(arrDBData, 1)
	' Display a table of the data in the array.
	set rs=Server.CreateObject("ADODB.Recordset")
	' Table header:
	Response.Write("<table align=center bgcolor='#ffffff' border=1 cellspacing=0 width=650>")
	Response.Write("<th>Name</th><th>Discipline</th><th>Expertise</th><th>Notes</th><th bgcolor='#fffacd'>Select?</th></tr>")
	Response.Write("<form method='post' action='select_reviewers.asp?id=" & loginid & "&msid=" & msid & "'>")
	recordcount=0
	For I = iRecFirst To iRecLast
		recordcount = recordcount + 1
		' A table row for each record					
		' Columns: Loop through the fields (first dimension of the array)
		For J = 0 To iFieldLast
			' Allow for special treatment of each field:
			select case J
				case 0   ' PartyID = Reviewer ID 
					partyid = arrDBData(J,I)
				case 1 ' Prefix
					prefix = arrDBData(J,I)
				case 2 ' First Name
					fname=arrDBData(J,I)
				case 3  ' Last Name
					lname=arrDBData(J,I)
					fullname = prefix & " " & fname & " " & lname
					Response.Write("<td align=center>" & fullname & "</td>")
					Response.Write("<td class=small>")
					Q = "SELECT tblPartyDis.PartyID, tblDiscipline.DisID, tblDiscipline.DisName "
					Q = Q & "FROM tblDiscipline INNER JOIN tblPartyDis ON tblDiscipline.DisID = tblPartyDis.DisID "
					Q = Q & " WHERE tblPartyDis.PartyID = " & partyid
					rs.Open Q, LoginDSN
					Do While not rs.EOF
						Response.Write (rs("DisName") & "&nbsp;")
						rs.MoveNext
					Loop
					rs.Close
					Response.Write ("</td>")
				case 4  ' Expertise 
					Response.Write("<td align=center class=small>" & arrDBData(J,I) & "</td>")
				case 5   ' Notes
					notes = left(arrDBData(J,I),20)  '  Truncate the notes for now.
					Response.Write("<td align=center class=small>" & notes & "</td>")
				case 6  ' Role
					role = arrDBData(J,I)
					if not role=3 then
						Response.Write("<H5>Error -  Role = " & role & " J=" & J & " I=" & I & " </H5>")
					end if
				case 7   ' Email of party
					email2 = arrDBData(J,I)
				case else
					Response.Write("Error - too many fields in array: J=" & J & " I= " & I & " data=" & arrDBData(J,I))			
			end select
		Next ' J	
		if not email2 = email then
			Response.Write("<td align=center bgcolor='#fffacd'><input type=checkbox name=assign" & partyid & " ></td></tr>")	
		else  '  Do not display a checkbox where the author = the reviewer
			Response.Write("<td align=center>&nbsp;</td></tr>")	
		end if
	Next ' I
	set rs = nothing
	Response.Write("</table><p>" & recordcount & " record(s) found.</p>")
	Response.Write("<input type=hidden name=msid value=" & msid & ">")
	Response.Write("<input type=hidden name=submitted value=true>")
	Response.Write("<input type=submit name=submit value='Assign Selected Reviewers'>")
	Response.Write("&nbsp;&nbsp;&nbsp;<input type=reset></form>")
else
	' Form data have been submitted
	' Assign the selected reviewers to the MS
	dim formname,item,intloop,slen,id,vid
	dim dateassigned,remindertype,reminderid
	if diagnostic then
	  Response.Write("<p>Number of form variables submitted = " & Request.Form.Count & "<br>")
		For each Item in Request.Form
			Response.Write("Element '" & item & "' Value = '" & Request.Form(Item) & "'<br>")
		Next
		Response.Write ("<p>----</p>")
		For each Item in Request.Form
			count = Request.Form(Item).Count 
			If count > 1 then
				Response.Write(Item & ":<br>")
				For intloop = 1 to count
					Response.Write ("Subkey " & intloop & " value = " & Request.Form(Item)(intloop) & "<br>")
				Next
			else
				Response.Write (Item & " = " & Request.Form(Item) & "<br>")
			end if
		Next	
		Response.Write ("<p>----</p>")
	end if
	dateassigned = Date()
	remindertype = "First"
	set rs=Server.CreateObject("ADODB.Recordset") 'Maybe could reuse the previous object here
	count = 0
	for each Item in Request.Form
		formname = Cstr(Left(Item,6))
		if formname = "assign" then  '  All the "on" checkboxes only
			count=count+1
			slen = len(Item)
			partyid = Cstr(Right(Item,slen-6))   ' ID number, string variable
			notify(count) = partyid   '  for use in next loop
			Response.Write("<p>Select reviewer: " & partyid & "</p>")
			Q = "INSERT INTO [tblReminder] (ReminderType, ReminderDate) "
			Q = Q & " VALUES ('" & remindertype & "','" & dateassigned & "')"
			rs.Open Q, LoginDSN,3,3
			Q1 = "SELECT TOP 1 ReminderID FROM tblReminder ORDER BY 1 DESC "
			rs.Open Q1, LoginDSN
			reminderid = rs("ReminderID")
			rs.Close
			Q2 = "INSERT INTO [tblAssignment] (MsID, PartyID, ReminderID, DateAssigned) "
			Q2 = Q2 & "VALUES (" & msid & "," & partyid & "," & reminderid & ",'" 
			Q2 = Q2 & dateassigned & "')"
			rs.Open Q2, LoginDSN,3,3
		end if
	next
	dim serverpath
	serverpath = Server.MapPath("select_reviewers.asp")
	for i=1 to count
		partyid = notify(i)
		'  Send an email notice to the new reviewer(s):
		'  First identify the data on the reviewer who has logged in to this page
		Q = "SELECT * FROM tblParty WHERE tblParty.PartyID = " & partyid
		rs.Open Q,LoginDSN
		fullname = rs("Prefix") & " " & rs("Fname") & " " & rs("Lname")
		email = rs("Email")
		pw = rs("Pw")
		rs.Close
		M = "Dear " & fullname & ":" & Chr(13)
		M = "I would like you to review a manuscript for possible publication in " & Chr(13)
		M = M & "Perspectives on Science and Christian Faith," & Chr(13)
		M = M & "the journal of the American Scientific Affiliation. " & Chr(13) & Chr(13)
		M = M & "You may log in to the ASA Editor's web site at " & Chr(13) 
		M = M & "www.asaeditor.org" & Chr(13)
		M = M & "using your email address " & email & " and your password: " & Chr(13)
		M = M & pw & Chr(13) & Chr(13)
		M = M & "Please review the assigned file(s) within 30 days if possible." & Chr(13) & Chr(13)
		M = M & "If this email reached you in error, or you would like to decline this assignment, " & Chr(13) 
		M = M & "please reply to this email with 'decline' in the subject line." & Chr(13) & Chr(13)  		
		M = M & "Thanks for your support of ASA!" & Chr(13) & Chr(13)
		M = M & "Sincerely," & Chr(13) & Chr(13)
		M = M & "Dr. Roman Miller, Editor " & Chr(13)
		M = M & "Perspectives in Science and Christian Faith" & Chr(13)
		M = M & "American Scientific Affiliation" & Chr(13)
		if left(serverpath,2)="c:" or diagnostic=true then
			' local web site; CDONTS not available
			Response.Write("<p>Test - an assignment message would be sent to <b>" & email & "</b>.</p>")
		else
			' Send a welcome email to the new member:
			' Use CDONTS
			dim subject,M,cc,bcc,importance,objMail
			subject="Assignment of MS for review" 
			set objMail = CreateObject("CDONTS.NewMail")
			with objMail
				.From = "paul@arveson.com"
				.To = email
			    .Cc = ""
			    .Bcc = ""
				.Subject = subject
				.Importance = 2   ' High=2; normal=1
				.BodyFormat = 1  ' HTML=0; text=1
				.Body = M
				.Send
			end with
			set objmail = nothing
		end if ' serversite
		Response.Write("<p><b>The following assignment message was sent to " & email & ":</b></p>")
		Response.Write("<p>" & M & "</p>")
	next  ' reviewer email
	set rs=nothing
end if   ' form submitted
%>
<p><a href="select_reviewers.asp?id=<%=loginid%>&msid=<%=msid%>&email=<%=email%>">Examine list again</a></p>

<p><a href="../statusboard5.asp?id=<%=loginid%>" target="_top">Return to Status Board</a></p>
</BODY>
</HTML>
