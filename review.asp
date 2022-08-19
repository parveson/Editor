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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional //EN">
<html>
<head>
<title>Show Review Status Data</title>
<LINK rel="stylesheet" type="text/css" href="StyleSheet1.css">
</HEAD>
<BODY>
<h3>Review Status</h3>
<%
' review.asp - called by statusboard.asp 
dim diagnostic
dim rs,I,j,k,d,Q1,Q2,Q3,Q4,sDSN,Imax,notes,status
dim msid,refid,days,flag,assigndate,reminder(10)
dim version(10),uploaddate(10),filename(10)
dim iRecFirst, iRecLast,iFieldFirst, iFieldLast
dim author1,author2,email1,email2,rphone,phone1,phone2,disciplines
dim fname,lname,title,remail,dateassigned,reviewdays
datapath = Server.MapPath("\database")
DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath & "\"
DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
sDSN = DSN1 & "Reviews.mdb" & DSN2
'sDSN=Application("cnSQL_ConnectionString")
diagnostic=false
' Select all records:
' Thanks to http://www.asp101.com/samples/db_getrows.asp
msid = Cint(Request.QueryString("msid"))
refid = Cint(Request.QueryString("r"))
'msid=1    ' TEMPORARY
'refid=1    ' TEMPORARY
' This query finds all versions of a MS to whom a specific reviewer has been assigned:
Q1 = "SELECT * FROM tblVersion INNER JOIN tblManuscript ON "
Q1 = Q1 & " tblVersion.VersionID = tblManuscript.VersionID;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open Q1, sDSN
if not rs.EOF then
	msid = rs("MsID")
	title = rs("Title")
	author1 = rs("Lname1")
	email1 = rs("Email1")
	phone1 = rs("Phone1")
	author2 = rs("Lname2")
	email2 = rs("Email2")
	phone2 = rs("Phone2")
end if
' Write Manuscript data:
'Response.Write ("<p>No. " & msid & "<br>")
Response.Write ("<p>MS Title: <b>" & title & "</b><br>")
Response.Write ("First Author: " & "<a href='mailto:" & email1 & "'>" & author1 & "</a><br>")
Response.Write ("Phone: " & phone1 & "<br>")
if not author2="" then
	Response.Write ("Second Author: " & "<a href='mailto:" & email2 & "'>" & author2 & "</a><br>")
	Response.Write ("Phone: " & phone2 & "</p>")
end if
'  Write version data:
k=0
do until rs.EOF
	k=k+1
	version(k) = rs("VersionNo")
	uploaddate(k)=rs("UploadDate")
	filename(k)=rs("FileName")
	rs.MoveNext
loop
rs.Close
Response.Write ("<p><b>Versions: </b></p>")
Response.Write ("<table border=1 cellspacing=0 cellpadding=2><tr>")
Response.Write ("<th>Version</th><th>Submit Date</th><th>Days</th><th>File Link</th></tr>")
for j=1 to k-1
	Response.Write ("<td>" & version(j) & "</td>")
	Response.Write("<td>" & uploaddate(j) & "</td>")
	days = DateDiff("d",uploaddate(j),Date)
	Response.Write ("<td>" & days & "</td>")
	Response.Write ("<td>" & filename(j) & "</td>")
next
Response.Write("</tr></table>")
Q2 = "SELECT * FROM tblParty WHERE tblParty.PartyID =  " & refid
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open Q2, sDSN
if not rs.EOF then
	' Write Reviewer data:
	fname = rs("RFname")
	lname = rs("RLname")
	disciplines=rs("Disciplines")
	remail = rs("Remail")
	rphone = rs("RPhone")
	notes = rs("Notes")
	Response.Write("<p>Reviewer name: <b>" & fname & " " & lname & "</b><br>")
	Response.Write("Disciplines: " & disciplines & "<br>")
	Response.Write("Email: " & remail & "<br>")
	Response.Write("Phone: " & rphone & "<br>")
	Response.Write("Notes: " & notes & "</p>")
end if
rs.Close
' Write assignment data:
Q3 = "SELECT tblAssignments.DateAssigned, tblReminders.ReminderType, tblReminders.ReminderDate, "
Q3 = Q3 & " tblAssignments.DateDone, tblAssignments.ReviewDays, "
Q3 = Q3 & " tblAssignments.Closed FROM tblReminders INNER JOIN (tblReviewers "
Q3 = Q3 & " INNER JOIN tblAssignments ON tblReviewers.REfID = tblAssignments.RefID) "
Q3 = Q3 & " ON tblReminders.ReminderID = tblAssignments.ReminderID "
Q3 = Q3 & " WHERE tblReviewers.RefID = " & refid 
Q3 = Q3 & " ORDER BY tblReminders.ReminderDate "
rs.Open Q3, sDSN
if not rs.EOF
	dateassigned=rs("DateAssigned")
	Response.Write("<p>Assignment Date: " & dateassigned & "<br>")
	status = rs("Closed")  ' Means review is completed 
	if status=false then
		days = DateDiff("d",dateassigned,Date)
		if days<=30 then
			flag="lightgreen"
		elseif (days>30 and days<=60) then
			flag="yellow"
		else
			flag="red"
		end if
		Response.Write("<table border=1 cellspacing=0 bgcolor='" & flag & "'><tr><td>")
		Response.Write("Days since assignment: " & days & "</td></tr></table>")
		I=0
		rs.MoveFirst
		do until rs.EOF
			reminder(I) = rs("ReminderDate")
			Response.Write(" Reminders Sent: <br>")
			Response.Write(reminder(i) & "<br>")
			I=I+1
			rs.MoveNext
		loop
		rs.Close
		Imax=I-1
		Response.Write("<table border=1 cellspacing=0>")
		for I = 1 to Imax
			Response.Write("<tr><td>" & I & "</td><td>" & reminder(I) & "</td></tr>")
		next
		Response.Write("</table>")
		' form to send a reminder goes here
		Response.Write("<form action='remind.asp?m=" & msid & "&email=" & remail & "&days=" & days & "'>")
		Response.Write("<input type=submit value='Send a reminder now'></form>")
	else  'status is true; review is closed.
		rs.Close
		' Show link to review file here
		Q4 = "SELECT tblReviews.ReviewDate, tblReviews.ReviewFile, "
		Q4 = Q4 & " tblReviews.ReviewNotes, tblReviewers.REfID, tblManuscripts.MsID, "
		Q4 = Q4 & " tblVersions.VersionNo FROM tblVersions INNER JOIN (tblReviews "
		Q4 = Q4 & " INNER JOIN (tblReviewers INNER JOIN (tblManuscripts INNER JOIN "
		Q4 = Q4 & " tblAssignments ON tblManuscripts.MsID = tblAssignments.MsID) "
		Q4 = Q4 & " ON tblReviewers.REfID = tblAssignments.RefID) ON tblReviews.ReviewID "
		Q4 = Q4 & " = tblAssignments.ReviewID) ON tblVersions.VersionID = "
		Q4 = Q4 & " tblManuscripts.VersionID "
		Q4 = Q4 & " WHERE MsID = " & msid & " AND RefID = " & refid
		rs.Open Q3, sDSN
		' There may be more than one review if there is more than one MS version.
		'  Allow for this.
		Response.Write("<p><b>This review assignment is completed.</b></p>")
		Response.Write("<p>Reviews submitted for this MS:<br>")
		I=0
		rs.MoveFirst
		do until rs.EOF
			Response.Write("<p>Review Date: " & rs("ReviewDate") & "<br>")
			Response.Write("File name: " & rs("ReviewFile") & "</p>")
			I=I+1
			rs.MoveNext
		loop
	end if ' status
else ' EOF
	Response.Write("<p>Reviewers have not yet been assigned to this MS.</p>")
end if
Response.Write("<p><b><a href='assign.asp?id=" & loginid & "&msid=" & msid & "'>" & Assign a reviewer</a></b></p>")
Response.Write("<p><a href='statusboard3.asp?id=" & loginid & "'>Return to Status Board</a></p>")
Response.Write("<p><a href='admin.asp?id=" & loginid & "'>Return to Admin menu</a></p>")
end if
%>

</BODY>
</HTML>
