<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<HTML>
<HEAD>
<TITLE>Reviewer Notes</TITLE>
<LINK rel="stylesheet" type="text/css" href="StyleSheet1.css">
<script language="Javascript" type="text/javascript">
</script>
</HEAD>
<BODY>
<p class=small>Administration</p>
<h4 align=center>Reviewer Notes</h4>
<%
dim partyid
partyid = Request.QueryString("partyid")
dim rs,Q,cnn  ' other variables are defined in test2.asp
Q = "SELECT * FROM [tblParty] WHERE [tblParty.PartyID] = " & partyid
set cnn = Server.CreateObject("ADODB.Connection")
cnn.Open LoginDSN
set rs=cnn.Execute(Q)
dim prefix,fname,lname,fullname,notes 		
' Name
prefix = rs("Prefix")
fname=rs("Fname")
lname=rs("Lname")
fullname = prefix & " " & fname & " " & lname
Response.Write("<p><b>" & fullname & "</b></p>")
' Notes
Response.Write("<p><b>Notes: </b><br>" & rs("Notes") & "</p>")
rs.Close
set rs = nothing
%>
<p><A href="javascript:window.close()">Close window</A></p>
</BODY>
</HTML>


