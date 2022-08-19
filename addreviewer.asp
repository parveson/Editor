<%@Language="VBScript"%>
<%option explicit
Response.expires=30
'Response.addHeader "pragma","no-cache"
'Response.addHeader "cache-control","private"
'Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Add a Reviewer</title>
<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
<script LANGUAGE="JavaScript" type="text/JavaScript">
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</script>
</head>
<body>
<p class=small>Administration</p>
<br>
<%
dim diagnostic,email
diagnostic=false
If Request.Form("submitted")<>"True" then
	' Form data have not been submitted
	%>
	<h4>Enter Reviewer Information</h4>
	<p><i><strong>All fields marked with an asterisk (*) must be entered.</strong></i></p>
	<form method="post" action="addreviewer.asp?id=<%=loginid%>">
		<table border=0>
		<tr>
		    <td align="right">Prefix:</td>
		    <td><input NAME="Prefix" size="20"></td>
		</tr>
		<tr>
		    <td align="right">First Name:</td>
		    <td><input NAME="Fname" size="20">*</td>
		</tr>
		<tr>
		    <td align="right">Last Name:</td>
		    <td><input NAME="Lname" size="30">*</td>
		</tr>
		<tr>
		    <td align="right">Phone:</td>
		    <td><input NAME="Phone" size="20">*</td>
		</tr>
		<tr>
		    <td align="right">Email:</td>
		    <td><input NAME="Email" size="50">*</td>
		</tr>
		<tr>
		    <td align="right">Expertise:</td>
		    <td><input NAME="Expertise" size="60" maxlength="60">*</td>
		</tr>
		<tr>
		    <td align="right">Notes:</td>
		    <td><textarea name="notes" rows="4" cols="60"></textarea>
		    *</td>
	  </tr>
	  </table>
	  
	  <p>Please select up to five (5) disciplines for this reviewer:</p>
	  <div align=left>
		<table border=0 cellspacing=0 cellpadding=2 class=small width=500>
		<!--<tr><th>Select</th><th>Discipline</th><th>Select</th><th>Discipline</th><th>Select</th><th>Discipline</th></tr>-->
		<% ' Show the table of disciplines:
		dim rs,I,J,Q 
		' Select records where SeqNo=0:
		' Thanks to http://www.asp101.com/samples/db_getrows.asp
		dim cnnGetRows   ' ADO connection
		dim strDBPath    ' Path to our Access DB (*.mdb) file
		dim arrDBData    ' Array that we dump all the data into
		dim iRecFirst, iRecLast,recordcount
		dim iFieldFirst, iFieldLast
		Q = "SELECT * FROM tblDiscipline"
		set cnnGetRows = Server.CreateObject("ADODB.Connection")
		cnnGetRows.Open LoginDSN
		set rs=cnnGetRows.Execute(Q)
		'arrDBData = rs.GetRows(-1,0,Array("MsID","SeqNo","UploadDate","Title","Lname","TypeName","VersionNo"))
		if not rs.EOF then
			arrDBData = rs.GetRows()
			rs.Close
			cnnGetRows.Close
			set cnnGetRows=nothing
			iRecFirst   = LBound(arrDBData, 2)
			iRecLast    = UBound(arrDBData, 2)
			iFieldFirst = LBound(arrDBData, 1)   '  = 0
			iFieldLast  = UBound(arrDBData, 1)
			' Display a table of the data in the array.
			' First loop through the records (second dimension of the array)
			recordcount=0
			For I = iRecFirst To 29 		' A table row for each record = I
				recordcount = recordcount + 1
				' Columns: Loop through the fields (first dimension of the array = J)
				For J = 0 To iFieldLast
					select case J
						case 0   ' DisID 
							Response.Write("<tr><td align=center><input type=checkbox name=check" & arrDBData(J,I) & " > </td>")	
						case 1   ' DisName
							Response.Write("<td align=left>" & arrDBData(J,I) & "</td>")
						case 2   ' DisDesc
					end select
				Next
				For J=0 to iFieldLast
					select case J
						case 0
							Response.Write("<td align=center><input type=checkbox name=check" & arrDBData(J,I+33) & "> </td>")	
						case 1
							Response.Write("<td align=left>" & arrDBData(J,I+33) & "</td>")
						case 2
					end select
				Next
				For J=0 to iFieldLast
					select case J
						case 0
							Response.Write("<td align=center><input type=checkbox name=check" & arrDBData(J,I+65) & "></td>")	
						case 1
							Response.Write("<td align=left>" & arrDBData(J,I+65) & "</td></tr>")
						case 2
					end select
				Next ' J
			Next ' I
		end if
		%>
		</table>
	</div>
	<br>
	<br>
	<input TYPE="submit" VALUE="Submit Data">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input TYPE="reset" VALUE="Reset All">
	<input type="hidden" name="submitted" value="True">
</form>
<%
else
	' Data have been submitted in the form.
	' This starts the second page, where the data will be updated.
	' Order of fields is not assumed.
	' List all the form fields for diagnostic:
	dim item
	if diagnostic then
	  Response.Write("<p>Number of form variables submitted = " & Request.Form.Count & "<br>")
		For each Item in Request.Form
			Response.Write("Element '" & item & "' Value = '" & Request.Form(Item) & "'<br>")
		Next
	end if
	' trim and replace quotes:
	dim prefix,fname,lname,phone,fullname,pw,entrydate,expertise,notes 
	prefix = left(Request.Form("Prefix"),10)
	fname = left(Replace(Request.Form("FName"),"'","''"),20)
	lname = left(Replace(Request.Form("LName"),"'","''"),30)
	fullname = prefix & " " & fname & " " & lname   
	phone = left(Replace(Request.Form("Phone"),"'","''"),20) 		
	email = Lcase(left(Replace(Request.Form("Email"),"'","''"),50)) 
	expertise = left(Replace(Request.Form("Expertise"),"'","''"),60)
	notes = left(Replace(Request.Form("Notes"),"'","''"),255)
	'  Check for errors in data entry
	dim errn
	errn=0
	if fullname = "" then 
		errn=errn+1
		Response.Write("<br>Name fields are blank.")
	end if
	if expertise ="" then
		errn=errn+1
		Response.Write("<br>Expertise field is blank.")
	end if
	if phone="" then
		errn=errn+1
		Response.Write("<br>Phone field is blank.")
	end if
	if email="" then
		errn=errn+1
		Response.Write("<br>Email field is blank.")
	end if
	if not CheckMail(email) then
		errn=errn+1
		Response.Write("<br>Email address is not valid.")
	end if
	' If any errors, return to form
	if errn>0 then
		Response.Write("<p><b>There are missing or invalid entries in your form.</b><br>")
		Response.Write("If you don't know the data, you may insert a dash (-).</p>")
		Response.Write("<p><b>Please <a href='Javascript:history.back();'>TRY AGAIN</a></b></p>")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
	' Else store data:
	else
		' Define additional fields and data not contained on form:
		dim refdate,quot,partyid,role,formname,disid,count
		entrydate= Now()
		pw = makePassword(6)
		role=3   '  Reviewers - numeric value
		' Before storing data, verify that the party (AS A REVIEWER) is not already stored:
		Q = "SELECT * FROM [tblParty] WHERE [Role]= 3 AND [Email] = '" & email & "'"
		set rs=server.CreateObject("ADODB.recordset")
		rs.Open Q,LoginDSN,1,3
		if not rs.EOF then
			Response.Write("<p><b>This Reviewer is already in the database!</b></p>")
			rs.Close
			set rs=nothing
		else
			rs.Close
			Q = "INSERT INTO [tblParty] ([Prefix],[Fname],[Lname],[Email],[Phone],[Pw],[Entrydate],[Role],[Expertise],[Notes]) "
			Q = Q & " VALUES ('" & prefix & "','" & fname & "','" & lname & "','" & email & "','" & phone & "','"
			Q = Q  & pw & "','" & entrydate & "'," & role & ",'" & expertise & "','" & notes & "');"					
			if diagnostic then 
				Response.Write ("<p>" & Q & "</p>")
			else
				' Open the database to insert the new record:
				rs.Open Q,LoginDSN,1,3
				' This stores the new record and closes the recordset
				Response.Write "<p><b>New data have been stored.</b></p>"
				Response.Write("<h5>Current data in this record:</h5>")
				' Read the data back and display:
				Q = "SELECT TOP 1 PartyID,Prefix,Fname,Lname,Email,Phone,Expertise,Notes FROM tblParty WHERE [role]= 3 ORDER BY 1 DESC "
				rs.Open Q,LoginDSN,1,3
				partyid=rs("PartyID")
				rs.Close
				Q = "SELECT * FROM tblParty WHERE PartyID = " & partyid
				rs.Open Q,LoginDSN,1,3
				Response.Write("<table bgcolor='#ffffff' border=1 width=650><tr><td>")
				Response.Write("<small>Reviewer ID=" & partyid & "</small><br>")
				Response.Write ("<p>Name: <b>" & rs("Prefix") & " " & rs("Fname") & " " & rs("Lname") & "</b></p>")
			' May have to remove extra apostrophes to make it look right. 
				Response.Write ("Email: <a href='mailto:" & rs("Email") & "'>" & rs("Email") & "</a><br>")
				Response.Write ("Phone: " & rs("Phone") & "<br>")
				Response.Write ("Expertise: " & rs("Expertise") & "<br>")
				Response.Write ("Notes: " & rs("Notes"))
				Response.Write("</td></tr></table>")
				rs.Close
				count = 0
				for each Item in Request.Form
					formname = Cstr(Left(Item,5))
					if formname = "check" then  '  All the "on" checkboxes only
						count=count+1
						if count < 6 then  ' Up to 5 disciplines allowed
							disid = Cstr(Right(Item,len(Item)-5))   ' ID number, string variable
							Response.Write("<p>Discipline " & count & ": " & disid & "</p>")
							Q = "INSERT INTO [tblPartyDis] (PartyID, DisID) "
							Q = Q & " VALUES (" & partyid & "," & disid & ")"
							rs.Open Q, LoginDSN,3,3
						end if
					end if
				next
				set rs = nothing
				Response.Write ("<p>Please review data for accuracy.</b><br>If incorrect, <a href='confirm.asp?id=" & loginid & "&partyid=" & partyid & "'>delete</a> and re-enter.</p>")
			end if   ' diagnostic
			'  Send an email notice to the new reviewer:
			M = "Dear " & fullname & ":" & Chr(13)
			M = "Your name has been included in the list of manuscript reviewers " & Chr(13)
			M = M & "for the American Scientific Affiliation. " & Chr(13)
			M = M & "If and when you are asked to review a manuscript, you may log in to the ASA web site at " & Chr(13) 
			M = M & "http://www.arveson.com/ORGS/asa/reviews/default.htm" & Chr(13)
			M = M & "using your email address " & email & " and your password: " & Chr(13)
			M = M & pw & Chr(13) & Chr(13)
			M = M & "You may log in here and change your password if desired. " & Chr(13)
			M = M & "If this email reached you in error, or you would like to unsubscribe, " & Chr(13) 
			M = M & "please reply to this email with 'unsubscribe' in the subject line." & Chr(13) & Chr(13)  
			M = M & "Thanks for your support of ASA!" & Chr(13) & Chr(13)
			M = M & "Sincerely," & Chr(13) & Chr(13)
			M = M & "Dr. Roman Miller, Editor " & Chr(13)
			M = M & "Perspectives in Science and Christian Faith" & Chr(13)
			M = M & "American Scientific Affiliation" & Chr(13)
			dim serverpath
			serverpath = Server.MapPath("addreviewer.asp")
			if left(serverpath,2)="c:" or diagnostic=true then
				' local web site; CDONTS not available
			else
				' Use CDONTS
				dim subject,M,cc,bcc,importance,objMail
				subject="Welcome to the ASA Reviewers Team!" 
				set objMail = CreateObject("CDONTS.NewMail")
				with objMail
					.From = "paul@arveson.com"
					.To = email
					.Cc = ""
					.Bcc = ""
					.Subject = subject
					.Importance = 2   ' High=2; normal=1
					.MailFormat = 1  ' 0 = MIME; 1 = plain text
					.BodyFormat = 1  ' HTML=0; text=1
					.Body = M
					.Send
				end with
				set objmail = nothing
			end if ' serversite
			Response.Write("<p><b>The following message was sent to " & email & ":</b></p>")
			Response.Write("<p>" & M & "</p>")
		end if  ' party already in database
	end if ' errors in form, errn>0
end if   '  form submitted

function makePassword(byVal maxLen)
	' Random password generator
	' by <A href="mailto:rob@aspfree.com">Robert Chartier</A>
	' Usage: newpw = makePassword(10) for a pw of length 10 chars.
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
		    whatsNext = Int((1 - 0 + 1) * Rnd + 0)
		If whatsNext = 0 Then
		'character
		   upper = 90
		   lower = 65
		Else
		   upper = 57
		   lower = 48
		End If
		strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	makePassword = strNewPass
end function

Function CheckMail(email)
	' Function to check email addresses
	' Usage: CheckMail(email,redirectpath)
	' Many thanks to http://www.aspsmith.com/re
	Dim objRegExp, blnValid
	'create a new instance of the RegExp object
	' note we do not need Server.CreateObject("")
	Set objRegExp = New RegExp
	'this is the pattern we check:
	objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
	'store the result either true or false in blnValid
	'blnValid = objRegExp.Test(email)
	CheckMail = objRegExp.Test(email)
	'If Not blnValid Then
		'do this if it is an invalid email address 
		'Response.Redirect("addreviewer.asp?id=" & loginid)
	'End If 
End Function
%>
<br>
<p><a href="menu.asp?id=<%=loginid%>">Return to Admin menu page</a></p>

</body>
</html>


