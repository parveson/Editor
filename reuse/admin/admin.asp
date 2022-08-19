<%@ Language=VBScript %>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = true
' Admin.asp - Members administrators only
' This is the only entry point for admin logins, called from distr.asp
dim email
email = Request.QueryString("email")
email=Trim(Lcase(email))
%>
<!-- #include file="test2.asp"  -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Admin Page</title>
<link rel="stylesheet" type="text/css" href="../../../EntryStyle.css">
</head>
<body>
<div align="center">
<h3>Administration Page</h3>
<h5>for BSC Applications</h5>
</div>

<p><b><%=email%></b></p>

<p>You are now logged in.</p>

<p><b><a href="memberlist.asp?id=<%=loginid%>">List all members - approve for access &amp; other tasks</a></b></p>

<p><strong><a href="export.asp?id=<%=loginid%>">Generate CSV file for all unblocked members</a></strong></p>

<p><strong><a href="changepw.asp?id=<%=loginid%>">Change my (administrator's) password</strong></p>

<!--<p><b><A href="../vendors/admin/admin.asp?id=<%=loginid%>">Administer Vendors Database</A></b></p>-->

<!--<p><b><a href="../links/admin/admin.asp?id=<%=loginid%>">Administer Links Database</A></b></p>-->

<p><strong><a href="../../../adopters/admin/index.asp?id=<%=loginid%>">Administer Adopters 
Database</a>  </strong></p>
<p><strong></strong>&nbsp;</p>

<p><b><a href="../../welcome.asp?id=<%=loginid%>">Go to Members Welcome Page</a><br></b>

<br>

<hr></p>

<p><a href="../../default.asp">Return to Login page</a></p>

</body>
</html>

