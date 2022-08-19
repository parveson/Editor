<%@ Language=VBScript %>
<html>
<head>
<title>Map Path to a File</title>
</head>
<body bgcolor="#FFFFFF">
<h2>Physical Path to File</h2>

<%
' This script finds the complete physical path to any file 
' on the server given a relative path from the global.asa file.
dim str
Response.Write "<p>The physical path to this file is: " & Server.MapPath("mappath.asp") & "</p>"
%>

</body>
</html>
