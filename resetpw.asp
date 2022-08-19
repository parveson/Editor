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
<!-- datapath,loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Reset a User's Password</title>
<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
</head>
<body>
<h3>Reset a User's Password</h3>
<% 
dim newpw,email,fuser,luser,partyid,Q1,Q,rs,diagnostic,role
diagnostic=false
If Request.Form("submitted")<>"True" then
	Response.Write("<form method='post' action='resetpw.asp?id=" & loginid & "'>")
	Response.Write("<p><b>Email address of user: </b><br><input type=text size=30 name='email'></p>")
	Response.Write("<p><b>Role: </b><br><input type=radio name='role' value=2 checked>Author<br>")
	Response.Write("<input type=radio name='role' value=3>Reviewer<br><br>")
	Response.Write("<input type=submit value='Reset Password'></p>")
	Response.Write("<input type=hidden name='submitted' value='True'>")
	Response.Write("<form>")
else
	email = Request.Form("email")
	role = Cint(Request.Form("role"))
	' Verify the user's email address and name:
	Q1 = "SELECT * FROM [tblParty] WHERE [Email] = '" & email & "' AND [Role] = " & role
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q1, LoginDSN,3,3
	if not rs.EOF then
		partyid = rs("PartyID")
		fuser=rs("Fname")
		luser=rs("Lname")
		rs.Close
		' Generate a new password:
		newpw = makePassword(6)
		' Store the new password in the database:
		Q = "UPDATE [tblParty] SET "
		Q = Q & " [Pw] = '" & newpw & "'"
		Q = Q & " WHERE (([PartyID]) = " & partyid & ")"
		if diagnostic then 
			Response.Write ("<p>" & Q & "</p>")
		else
			' This stores the new record and closes the recordset
			rs.Open Q, LoginDSN,3,3
			Response.Write "<p><b>A new password has been generated and stored.</b></p>"
			' Notify the user of his new password:
			dim serverpath
			serverpath = Server.MapPath("resetpw.asp")
			if left(serverpath,2)="c:" then
				' local web site; CDONTS not available
				Response.Write("<p>CDONTS not available; notification would be sent to <b>" & email & "</b></p>")
				Response.Write("<P>The user's new password is: <b>" & newpw & "</b></p>")
			else
				' Use CDONTS
				dim subject,recipient,M,cc,bcc,importance,objMail
				subject="ASA Administrator has reset your password"
				' Memo field special treatment:
				' Define additional fields and data:
				' Define recipients where contact form data should go:
				' Concatenate several variables into the message:
				M = "Dear " & fuser & " " & luser & ":" & Chr(13)
				M = M & "A request for a new password was received from your email address. " & Chr(13)
				M = M & "(" & email & ") on " & refdate & "." & Chr(13) & Chr(13)
				M = M & "Your password has been reset.  Your new password is: " & newpw & Chr(13) & Chr(13)
				M = M & "Please make a note of this new password." & Chr(13) & Chr(13)
				M = M & "You may log in here: " & Chr(13)
				M = M & "<a href='http://www.asaeditor.org/'>http://www.asaeditor.org/</a>" & Chr(13) & Chr(13)
				M = M & "Thank you for your interest in serving the " & Chr(13) & Chr(13)
				M = M & "American Scientific Affiliation." & Chr(13)
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
				Response.Write ("<p><b>Notification has been sent to " & email & ".</b></p>")
			end if ' serversite
			set rs=nothing
		end if
	else
		Response.Write("<p><b>This user/role is not in the database.</p>")
	end if ' rs.EOF
end if ' submitted
function makePassword(byVal maxLen)
' Random password generator
' by <A href="mailto:rob@aspfree.com">Robert Chartier</A>
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
%>

<p><a href="menu.asp?id=<%=loginid%>">Return to Admin page</a></p>

</body>
</html>
