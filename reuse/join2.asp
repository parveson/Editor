<%@ Language=VBScript %>
<%option explicit
' Member join form - part 2
dim diagnostic
diagnostic=false
dim email,pw,entrytime
' These variables come from join.html:
email=Left(Request.Form("Email"),80)
if email="" then
	Response.Redirect("join.asp")
	' Email cannot be blank. 
end if
CheckMail(email)
pw=left(Request.Form("pw1"),20)
if pw="" then 
	Response.Redirect("join.asp")
	'  Password cannot be blank.  Otherwise there are no additional tests. 
	'  A more stringent test could be performed here.
end if
entrytime=Left(Request.Form("timex"),20)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Add Your Data</title>
<link rel="stylesheet" type="text/css" href="../../sheet3.css">
<script LANGUAGE="JavaScript1.1" SRC="../../FormChek.js"></script>
</head>
<body>
<div align="center">
<p><strong>Please enter the following optional data to allow us 
<br>to provide you with more relevant information:</strong></p>
<form action="../thankyou.asp" method="post">
	<table bgcolor="lightgreen" border="2">
		<tr><td align="right"><strong>First name: &nbsp;</strong></td>
		<td><input name="firstname" size="30" maxlength="40"></td></tr>
		<tr><td align="right"><strong>Last name:&nbsp;</strong></td>
		<td><input name="lastname" size="30" maxlength="40"></td></tr>
		<tr><td align="right"><strong>Country:&nbsp;</strong></td>
		<td><input name="country" size="15" maxlength="15"></td></tr>
		<tr><td align="right"><strong>Your occupation:&nbsp;</strong></td>
			<td><select name="occupation" size="1">
				<option value="7" selected>select one&nbsp;
				<option value="1">student&nbsp;
				<option value="2">private sector
				<option value="3">public sector
				<option value="4">nonprofit sector&nbsp;
				<option value="5">consultant
				<option value="6">other
				</option>
			</select>
		</td></tr>
		<tr><td align="right">		
		<br>
		<br><input type="submit" value="Submit"></td></tr>
	</table>
		<input type="hidden" name="email" value="<%=email%>">
		<input type="hidden" name="pw" value="<%=pw%>">
		<input type="hidden" name="entrytime" value="<%=entrytime%>">
</form> 
</div>
       
<p><a href="../../default.html"><strong>Return to home page</strong></a></p>

<!-- #include file="../footer.htm" --> 
<%
Function CheckMail(email,redirectpath)
	' Function to check email addresses
	' Usage: CheckMail(email)
	' Many thanks to http://www.aspsmith.com/re
	Dim objRegExp, blnValid
	'create a new instance of the RegExp object
	' note we do not need Server.CreateObject("")
	Set objRegExp = New RegExp
	'this is the pattern we check:
	objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
	'store the result either true or false in blnValid
	blnValid = objRegExp.Test(email)
	If Not blnValid Then
		'do this if it is an invalid email address 
		Response.Redirect(redirectpath)
	End If 
End Function
%>
</body>
</html>
