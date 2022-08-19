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
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<head>
<title>Delete a Party</title>
<LINK rel="stylesheet" type="text/css" href="StyleSheet1.css">
</head>
<body>
<p class=small>Administration</p>
<br>
<%
' Called from addauthor.asp, addeditor.asp, or addreviewer.asp
dim partyid,Q,rs,fullname
partyid=Request.QueryString("partyid")
Q = "DELETE FROM [tblParty] WHERE [PartyID]=" & partyid
set rs=server.CreateObject("ADODB.recordset")
rs.Open Q,LoginDSN,1,3
Q = "DELETE FROM [tblPartyDis] WHERE [PartyID]=" & partyid
rs.Open Q,LoginDSN,1,3
Q = "DELETE FROM [tblAssignment] WHERE [PartyID]=" & partyid
rs.Open Q,LoginDSN,1,3
set rs=nothing
Response.Write("<p>Record ID = " & partyid & " has been deleted.</p>")
%>
<p><a href="menu.asp?id=<%=loginid%>">Return to Admin page</a></p>
</body>
</html>
