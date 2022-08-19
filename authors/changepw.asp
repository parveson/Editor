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
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<title>Change My Password</title>
</head>
<body>
<%
dim newpw1,newpw2,email,rs,Q,pw
if not Request.Form("submitted") then
	'  Form has not been submitted yet
	%>
	<H3 align=center>Change My Password</H3>
	<p><b>Enter your new password.</b><br>
	<i><STRONG>Please make a note of your new password!</STRONG></i></p>
	<FORM  method="post" action="changepw.asp?id=<%=loginid%>">
		<p>New password:</p>
		<input type=password name=newpw1>
		<p>Enter again:</p>
			<input type=password name=newpw2>
		<br><br>
		<input type=hidden name=submitted value="true">
		<input type=submit name=submit value="Submit New Password">
		&nbsp;&nbsp;&nbsp;
		<input type=reset name=reset value="Reset">
	</FORM>
	<hr>
	<%
else
	'  Form has been submitted
	'  If passwords are equal, update the login database table
	newpw1=Request.Form("newpw1")
	newpw2=Request.Form("newpw2")
	if StrComp(newpw1,newpw2)=0 then
		pw = newpw1
		set rs=Server.CreateObject("ADODB.Recordset")
		'  Store the new password in the database, where the 
		'  unique login ID is used to identify the correct record:
		Q = " UPDATE tblParty SET Pw = '" & pw & "' WHERE [Entrytime] = '" & loginid & "'"
		rs.Open Q, LoginDSN,2,3
		set rs=nothing
		'Response.Write ("<p>Query: " & Q & "</p>")
		Response.Write("<p>Your password has been changed to: <b>" & pw & "<br>")
		Response.Write("<i>Please make a note of it now!</i></b></p>")
	 else
		  Response.Write("<p><b>Unequal password entries, <a href='changepw.asp?id=" & loginid & "'>try again</a></b></p>")
	end if
end if%>
<p><A href="menu.asp?id=<%=loginid%>" >Return to Menu</a></p>

</body>
</html>
