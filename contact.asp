<%@ Language=VBScript %>
<%option explicit
' Contact.asp    
' Sends plain text email from a form.
' Using CDONTS, modeled after sendmail.pl
if Request.Form("Submit")="" then
%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
	<html>
	<head>
	<title>Contact Us</title>
	<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
	<meta NAME="DATE" CONTENT="12 Dec 2018">
	</head>
	<body>
	<div align="left">
	<h3>Contact Us</h3>
	</div>
	<p>Please select one of the following options:
	<br>
	<br>
	<strong>If you would like general information about us or our online resources, 
	please <a href="http://www.arveson.com/contact.asp">click here.</a><br><br>   
	If you have author or reviewer-related questions, please use this form:</strong>
	<form action="contact.asp" method="post">
		<p><strong>Message Category:</strong><br> 
			<input type="radio" name="subject" value="Author">I would like to submit a manuscript<br>
			<input type="radio" name="subject" value="Editor_request" checked>Message to the editor<br>
			<input type="radio" name="subject" value="Web_issue">Technical issue with this web site<br>
			<input type="radio" name="subject" value="General_question">General question
			</p>
<p><small><font color=maroon>Note: for testing purposes, all emails are currently sent to the webmaster.</font></small></p>		
		<p><strong>Your name:<br></strong>
		<input name="realname" size="30" maxlength="40">
		</p>
		<p><strong>Your email address:</strong><br>
		<input name="sender" size="30" maxlength="40">
		</p>
		<p><strong>Phone:</strong><br>
		<input name="phonex" maxlength="20">
		</p>
		<p><strong>Your message:<br>
		</strong><textarea name="messagetext" rows="10" cols="64"></textarea>
		</p>
		<p><input type="checkbox" name="optin" checked> Please send me online news announcements.
		</p>
		<p><input type="submit" name="Submit" value="Submit"></p>
	</form>         
	<%
else
	' Form has been filled in; process the data:
	dim recipient  ' email address where mail will be sent.
	' Radio buttons:
	dim subject
	subject=Cstr(trim(left(Request.Form("subject"),20)))
	select case subject
		case "Author"
			recipient = "paularveson@gmail.com"
		case "Editor_request"
			recipient = "paularveson@gmail.com"
		case "Web_issue"
			recipient = "paularveson@gmail.com"
		case "General_question"
			recipient = "paularveson@gmail.com"
	end select
	dim realname
	realname = Cstr(trim(left(Request.Form("realname"),40)))
	dim sender
	sender = Cstr(trim(left(Request.Form("sender"),40)))
	CheckMail(sender)
	dim phone
	phone = Cstr(trim(left(Request.Form("phonex"),12)))
	' Message field length limit:
	dim message
	message = Cstr(trim(left(Request.Form("messagetext"),2000)))
	'  Treat the checkboxes separately, by name:
	dim optin
	optin=Request.Form("optin")
	if optin="" then
		optin="off"
	end if
	' Define additional fields and data not contained on form:
	dim ipaddress,refdate
	refdate= Now()
	ipaddress=trim(Request.ServerVariables(33))
	dim M,cc,bcc,importance,objMail
	' Concatenate several variables into the message: 
	M = "A message to you was sent by " & Chr(13)
	M = M & " " & realname & "( " & sender & " ) phone no. " & phone & " on " & refdate & Chr(13)
	M = M & "------------------------------------------------------" & Chr(13)
	M = M & message & "   " & Chr(13)
	M = M & "Opt in: " & optin & "   " & Chr(13)
	set objMail = CreateObject("CDONTS.NewMail")
	with objMail
		.From = sender
		.To = recipient
	    .Cc = ""
		.Subject = subject
		.Importance = 2   ' High=2; normal=1
		.BodyFormat = 1   ' HTML=0; text=1
		.Body = M
		.Send
	end with	
	set objmail = nothing	
	Response.Redirect("thankyou.htm")
end if

Function CheckMail(strEmail)
'our function to check email addresses
' many thanks to http://www.aspsmith.com/re which has 
'a fantastic list of regular expressions
Dim objRegExp, blnValid
'create a new instance of the RegExp object
' note we do not need Server.CreateObject("")
Set objRegExp = New RegExp
'this is the pattern we check:
objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
'store the result either true or false in blnValid
blnValid = objRegExp.Test(strEmail)
If Not blnValid Then
	'do this if it is an invalid email address
	Response.Redirect "contact_error.htm"
End If 
End Function

%>
<p>&nbsp;</p>
<p><a href="default.htm"><strong>Return to Editor's home page</strong></a></p>
<p><a href="http://www.arveson.com/">Return to main web site</a></p>
<hr>

</body>
</html>
