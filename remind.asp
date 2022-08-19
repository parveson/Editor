<%@Language="VBScript"%>
<%option explicit
Response.expires=30
'Response.addHeader "pragma","no-cache"
'Response.addHeader "cache-control","private"
'Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Remind a Reviewer</title>
<link rel="stylesheet" type="text/css" href="StyleSheet1.css">
<script LANGUAGE="JavaScript" type="text/JavaScript">
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</script>
</head>
<body>
<p class=small>Administration</p>
<br>
<%
' Called from the Statusboard, table of files assigned to reviewers.
' Link identifies which reviewer and file.
dim diagnostic,msid,partyid,fullname,dateassigned,email
dim rs,I,j,k,d,M,Q,Q1,Q2,Q3
dim quot,role,count,serverpath,title
dim reviewdays,remindertype,reminderdate,reminderid
dim subject,recipient,cc,bcc,importance,objMail
msid = Request.QueryString("msid")
partyid = Request.QueryString("partyid")
diagnostic=false
If Request.Form("submitted")<>"True" then
	' Form data have not been submitted
	Response.Write("<h4>Reviewer's Reminders</h4>")
	' Find reminders that have already been sent to this reviewer (partyid)
	' for this file (msid)
	Q = "SELECT tblReminder.ReminderType, tblReminder.ReminderDate, tblAssignment.MsID, "
	Q = Q & " tblAssignment.PartyID, tblAssignment.ReminderID, tblAssignment.DateAssigned, "
	Q = Q & " tblAssignment.ReviewDays, tblParty.Prefix, tblParty.Fname, tblParty.Lname, "
	Q = Q & " tblParty.Email, tblParty.Role "
	Q = Q & " FROM tblParty INNER JOIN (tblReminder INNER JOIN tblAssignment ON "
	Q = Q & " tblReminder.ReminderID = tblAssignment.ReminderID) ON tblParty.PartyID = "
	Q = Q & " tblAssignment.PartyID "
	Q = Q & " WHERE tblParty.Role=3 AND tblAssignment.MsID = " & msid & " AND tblAssignment.PartyID = " & partyid
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q,LoginDSN
	' Write the reminder data
	j=0
	do while not rs.EOF
		j = j + 1 
		fullname = rs("Prefix") & " " & rs("Fname") & " " & rs("Lname")
		Response.Write("<p>" & rs("ReminderType") & " reminder sent to " & fullname & " at " & rs("Email") & " on " & rs("ReminderDate") & "</p>")
		dateassigned = rs("dateassigned")
		rs.MoveNext
	loop
	if j=0 then
		Response.Write("<p>No reminders have been sent.</p>")
	end if
	rs.Close
	set rs=nothing
	%>
	<FORM method="post" action="remind.asp?id=<%=loginid%>&msid=<%=msid%>&partyid=<%=partyid%>">
	<input TYPE="submit" VALUE="Send a Reminder Now">
	<input type="hidden" name="email" value="<%=email%>">
	<input type="hidden" name="dateassigned" value="<%=dateassigned%>">	
	<input type="hidden" name="fullname" value="<%=fullname%>">
	<input type="hidden" name="submitted" value="True">
	</FORM>
	<%
else
	' Data have been submitted in the form.
	' Note: msid and partyid have been obtained from the querystring above.
	email = Request.Form("email")
	email = "paul@arveson.com"   ' **************  DELETE WHEN READY ************
	dateassigned = Request.Form("dateassigned")
	fullname = Request.Form("fullname")
	reminderdate= Now()
	'  How many days have elapsed since assignment?
	reviewdays = DateDiff("d",dateassigned,Date)
	remindertype = "Initial"
	if reviewdays>=0 and reviewdays <=30 then
		remindertype="<30 days"
	elseif reviewdays >30 and reviewdays<=60 then
		remindertype = "<60 days"
	else
		remindertype = ">=60 days"
	end if
	Q = "INSERT INTO [tblReminder] (ReminderType, ReminderDate) "
	Q = Q & " VALUES ('" & remindertype & "','" & reminderdate & "')"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q, LoginDSN,3,3
	Q1 = "SELECT TOP 1 ReminderID FROM tblReminder ORDER BY 1 DESC "
	rs.Open Q1, LoginDSN
	reminderid = rs("ReminderID")
	rs.Close
	Q2 = "INSERT INTO [tblAssignment] (MsID, PartyID, ReminderID, DateAssigned, ReviewDays) "
	Q2 = Q2 & "VALUES (" & msid & "," & partyid & "," & reminderid & ",'" 
	Q2 = Q2 & dateassigned & "'," & reviewdays & ")"
	Q3 = "SELECT Title from tblFile WHERE tblFile.MsID = " & msid
	rs.Open Q3, LoginDSN
	title = rs("Title")
	rs.Close
	serverpath = Server.MapPath("remind.asp")
	if left(serverpath,2)="c:" or diagnostic=true then
		' local web site; CDONTS not available
		Response.Write("<p>Q2 = " & Q2 & "</p>")
		M = "Test message."
	else
		rs.Open Q2, LoginDSN,3,3
		'  Send an email notice to the reviewer:
		' Use CDONTS
		subject="A friendly reminder from your ASA Editor" 
		M = "Dear " & fullname & ":" & Chr(13)
		M = M & "Thank you for volunteering to review the manuscript:" & Chr(13)
		M = M & "'" & title & "'." & Chr(13)& Chr(13)
		M = M & "It has been " & reviewdays & " days since this review was assigned." & Chr(13)
		M = M & "Please submit your review as soon as possible.  " & Chr(13) & Chr(13)
		M = M & "Thank you, " & Chr(13) & Chr(13)
		M = M & "Roman Miller, Editor" & Chr(13)
		M = M & "American Scientific Affiliation"
		' Concatenate several variables into the message:
		set objMail = CreateObject("CDONTS.NewMail")
		with objMail
			.From = "paul@arveson.com"   '  Should be changed to MillerRJ&rica.net
			.To = email     
		    .Cc = ""
		    .Bcc = ""
			.Subject = subject
			.Importance = 2   ' High=2; normal=1
			.BodyFormat = 1  ' HTML=0; text=1
		   ' .MailFormat = 0  ' 0 = MIME; 1 = plain text
			.Body = M
			.Send
		end with
		set objMail=nothing
	end if
	Response.Write("<p><b>The following reminder message was sent to " & email & ":</b></p>")
	Response.Write("<p>" & M & "</p>")
end if   '  form submitted
set rs=nothing
%>
<br>
<p><a href="statusboard5.asp?id=<%=loginid%>">Return to Status Board</a></p>
<p><a href="menu.asp?id=<%=loginid%>">Return to Admin menu</a></p>

</body>
</html>


