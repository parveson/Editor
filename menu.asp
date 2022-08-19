<%@ Language=VBScript %>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = true
' Admin.asp - Administrators only
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
<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
</head>
<body>

<p class=small>Manuscript Review Application</p>

<h3 align=center>Administration Menu</h3>
<h5 align=center>for Editors Only</h5>

<p><A href="statusboard5.asp?id=<%=loginid%>"><STRONG>View Status 
Board</STRONG>  </A></p>

<p><A href="listauthors.asp?id=<%=loginid%>"><STRONG>List all 
Authors</STRONG>  </A>, add, delete authors</p>

<p><A href="listreviewers.asp?id=<%=loginid%>"><STRONG>List all 
Reviewers</STRONG>  </A>, add, delete reviewers</p>

<p><A href="listdisciplines.asp?id=<%=loginid%>"><STRONG>List all 
Disciplines</STRONG>  </A>, add disciplines</p>

<p><A href="listeditors.asp?id=<%=loginid%>"><STRONG>List all 
Editors</STRONG>  </A>, add, delete editors</p>

<p><A href="changepw.asp?id=<%=loginid%>"><STRONG>Change my 
password</STRONG></p>

<p><A href="resetpw.asp?id=<%=loginid%>"><STRONG>Reset a User's 
Password</STRONG>   </A></p>

<p><A href="view_archive.asp?id=<%=loginid%>"><STRONG>Transfer closed 
reviews to Archive</STRONG>    </A></p><STRONG>

<hr>
</STRONG>

<p><A href="login.htm"><STRONG>Return to Login page</STRONG>   </A></p>

<p><A href="default.htm"><STRONG>Return to home 
page</STRONG>   </A></p>

</body>
</html>

