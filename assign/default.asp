<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = false 
%>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<%
'  assign/default.asp  - called from statusboard.asp 
'  For use by Editor to select reviewers of a MS.
dim msid,email
msid = Request.QueryString("msid")
email = Request.QueryString("email")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN">
<HTML lang="en">
<HEAD>
	<TITLE>Assign Reviewers</TITLE>
	<meta HTTP-EQUIV="content-Type" CONTENT="text/html;charset=ISO-8859-1">
</HEAD>
<frameset border="1" frameborder="1" framespacing="0" rows="200,95%">
	<frame src="assign_top.asp?id=<%=loginid%>&msid=<%=msid%>" noresize scrolling="no">
	<frame src="select_reviewers.asp?id=<%=loginid%>&msid=<%=msid%>&email=<%=email%>" border="0" scrolling="auto">
</frameset>
</HTML>
