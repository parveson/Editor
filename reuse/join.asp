<%@ Language=VBScript %>
<%option explicit
' Member join form - part 1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Join the BSCI Network</title>
<link rel="stylesheet" type="text/css" href="../../sheet2.css">
<meta NAME="DATE" CONTENT="3 Jan 2005">
<SCRIPT language="JavaScript1.1" type="text/Javascript" src="../FormChek3.js">
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" type="text/JavaScript">
	<!--
function clicktime(form) {
// validate data:
if ((form.Email.value == "") || (form.pw1.value=="") || (form.pw2.value=="")) {
	alert("Please fill in all fields.");
	}
else {
	if (form.pw1.value==form.pw2.value) {
		objDate = new Date();
		// get ms since Jan 1, 1970:
			 var ms = objDate.getTime();
		form.timex.value = ms;
		//alert(ms);
		form.submit();
		}
	else {
		alert("Passwords do not match - please try again.");
		}
	}
}
//-->
</script>
</head>
<body>
<div align="left">
<table border="0">
	<tr><td>
	<IMG height=80 src="../../images/bsci_logo.JPG" width=131>
	<IMG height=80 src="../../images/singlepixel.GIF" width=94>
	</td>
	<td>
	<h3>Join our Global Membership Network</h3>
	</td></tr>
</table>
</div>
<p>The <strong>Balanced Scorecard Institute's Global Network</strong> 
  offers FREE white papers, announcements and other offerings that 
will help you to keep informed about the balanced scorecard, 
performance measurement and strategic management. We welcome you to join this list! 
You will normally receive less than one email alert per month. </p>

<div align="center">
<table bgcolor="lightgreen" width="500" border=1 cellspacing=0>
<tr><td>
<form action="../join2.asp" method="post">
	<table bgcolor="lightgreen" width="500" border=0>
		<tr><td align="right"><strong>Your Email address: </strong></td>
		<td>
		<input name="Email" onfocus="javascript:promptEntry(pEmail)" 
		onchange="javascript:checkEmail(this,true)" size="48" maxlength="50">&nbsp;
		</td></tr>
		<tr><td align="right"><strong>Create a password: </strong></td>
		<td><input type="password" name="pw1" size="10" maxlength="10"></td></tr>
		<tr><td align="right"><strong>Enter Password again: </strong></td>
		<td><input type="password" name="pw2" size="10" maxlength="10"></td></tr>
		<tr><td align="right">
		<input type=hidden name="timex">
		<input type="hidden" name="env_report" value="REMOTE_HOST,HTTP_USER_AGENT,HTTP_REFERER">
		<br><input type=button value="Join Now" onClick="javascript:clicktime(this.form)">
		</td></tr>
	</table>
</form>
</td></tr>
</table>
<br>
<form action="../sendpw.asp" method="post">
	<table width="550" border=0>
		<tr><td>I have already joined.  Please send my password to my &nbsp;
		<tr><td>Email address: 
		<input name="Email" onfocus="javascript:promptEntry(pEmail)" 
		onchange="javascript:checkEmail(this,true)" size="48" maxlength="50">&nbsp;
		&nbsp; 
		<input type="Submit" value="Submit">
		</td></tr>
	</table>
</form>
</div>

<p><a href="../default.asp">Members login </a></p>
	
<p><a href="../../policies/privacy_policy.html">Our privacy policy</a></p>

<p><a href="../../default.html">Return to home page</a></p>

<!-- #include file="footer.htm" -->
</body>
</html>
