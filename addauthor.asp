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
<HTML>
<HEAD>
<TITLE>Add an Author</TITLE>
<LINK rel="stylesheet" type="text/css" href="StyleSheet1.css">
<SCRIPT LANGUAGE="JavaScript" type="text/JavaScript">
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</SCRIPT>
</HEAD>
<BODY>
<p class=small>Administration</p>
<br>
<%
dim diagnostic,email
diagnostic=false
If Request.Form("submitted")<>"True" then
	' Form data have not been submitted
	%>
	<H4>Add Author Information</H4>
	<P><I><STRONG>All fields marked with an asterisk (*) must be entered.</STRONG></I></P>
	<FORM method="post" action="addauthor.asp?id=<%=loginid%>">
		<TABLE border=0>
		<TR>
		    <TD align="right">Prefix:</TD>
		    <TD><INPUT NAME="Prefix" size="20"></TD>
		</TR>
		<TR>
		    <TD align="right">First Name:</TD>
		    <TD><INPUT NAME="Fname" size="20">*</TD>
		</TR>
		<TR>
		    <TD align="right">Last Name:</TD>
		    <TD><INPUT NAME="Lname" size="30">*</TD>
		</TR>
		<TR>
		    <TD align="right">Phone:</TD>
		    <TD><INPUT NAME="Phone" size="20">*</TD>
		</TR>
		<TR>
		    <TD align="right">Email:</TD>
		    <TD><INPUT NAME="Email" size="50">*</TD>
		</TR>
		<TR>
		    <TD align="right">Expertise:</TD>
		    <TD><INPUT NAME="Expertise" size="60" maxlength="60"></TD>
		</TR>
		<TR>
		    <TD align="right">Notes:</TD>
		    <TD><textarea name="notes" rows=4 cols=60></textarea>
		    *</TD>
	  </TR>
	  <TR>
		<TD align="right"><input type=checkbox name="member"></td>
		<td>I am a member of the ASA.</td>
	 </tr>
	  <TR>
		<td><INPUT TYPE="submit" VALUE="Submit Data">
		</TD>
		<td>
			<br><br><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT TYPE = "reset" VALUE = "Reset All">
			<input type=hidden name="submitted" value="True">
		</td>
	  </TR>
	</TABLE>
	</FORM>
<%
else
	' Data have been submitted in the form.
	' This starts the second page, where the data will be updated.
	' Order of fields is not assumed.
	' List all the form fields for diagnostic:
	dim item
	if diagnostic then
	  Response.Write("<p>Number of form variables submitted = " & Request.Form.Count & "<br>")
		For each Item in Request.Form
			Response.Write("Element '" & item & "' Value = '" & Request.Form(Item) & "'<br>")
		Next
	end if
	dim prefix,fname,lname,phone,fullname,pw,entrydate,expertise,notes,partyid 
	prefix = left(Request.Form("Prefix"),10)
	fullname = prefix & " " & fname & " " & lname 
	' trim and replace quotes before storing in database: 
	fname = left(Replace(Request.Form("FName"),"'","''"),20)
	lname = left(Replace(Request.Form("LName"),"'","''"),30) 
	phone = left(Replace(Request.Form("Phone"),"'","''"),20) 		
	email = Lcase(left(Replace(Request.Form("Email"),"'","''"),50)) 
	expertise = left(Replace(Request.Form("Expertise"),"'","''"),60)
	notes = left(Replace(Request.Form("Notes"),"'","''"),255)
	'  Check for errors in data entry
	dim errn
	errn=0
	Response.Write("<b><font face='Verdana' color='red'>")
	if fullname = "" then 
		errn=errn+1
		Response.Write("<br />Name fields are blank.")
	end if
	if phone="" then
		errn=errn+1
		Response.Write("<br />Phone field is blank.")
	end if
	if email="" then
		errn=errn+1
		Response.Write("<br />Email field is blank.")
	end if
	if not CheckMail(email) then  
		errn=errn+1
		Response.Write("<br />Email address is not valid.")
	end if
	Response.Write("</font></b>")
	' If any errors, return to form
	if errn>0 then
		Response.Write("<p><b>There are missing or invalid entries in your form.</b><br>")
		Response.Write("If you don't know the data, you may insert a dash (-).</p>")
		Response.Write("<p><b>Please <a href='Javascript:history.back();'>TRY AGAIN</a></b></p>")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
	' Else store data:
	else
		' Define additional fields and data not contained on form:
		dim refdate,quot,role,editorname,header,footer
		entrydate= Now()
		pw = makePassword(6)
		role=2   '  Authors - numeric value
		dim rs,i,Q,Q1,Q2  ' other variables are defined in test2.asp
		' Get the name of the Editor who is currently logged in:
		Q1 = "SELECT [Fname],[Lname],[Entrytime] FROM tblParty "
	    Q1 = Q1 & " WHERE Entrytime = '" & loginid & "'"
	    set rs=Server.CreateObject("ADODB.Recordset")
	    rs.Open Q1,LoginDSN
	    editorname = rs("Fname") & " " & rs("Lname")
	    editorname = Replace(editorname,"''","'")
	    rs.Close
		' Before storing data, verify that the party (AS AN AUTHOR) is not already stored:
		Q2 = "SELECT * FROM [tblParty] WHERE [Role]= 2 AND [Email] = '" & email & "'"
		rs.Open Q2,LoginDSN,1,3
		if not rs.EOF then
			Response.Write("<p><b>This Author is already in the database!</b></p>")
			rs.Close
			set rs=nothing
		else
			rs.Close
			Q = "INSERT INTO [tblParty] ([Prefix],[Fname],[Lname],[Email],[Phone],[Pw],[Entrydate],[Role],[Expertise],[Notes]) "
			Q = Q & " VALUES ('" & prefix & "','" & fname & "','" & lname & "','" & email & "','" & phone & "','"
			Q = Q  & pw & "','" & entrydate & "'," & role & ",'" & expertise & "','" & notes & "');"					
			if diagnostic then 
				Response.Write ("<p>" & Q & "</p>")
			else
				' Open the database to insert the new record:
				rs.Open Q,LoginDSN,1,3
				' This stores the new record and closes the recordset
				Response.Write "<p><b>New data have been stored.</b></p>"
				Response.Write("<h5>Current data in this record:</h5>")
				'Q = "SELECT * FROM [tblParty] WHERE [role]= 2"
				Q = "SELECT TOP 1 PartyID,Prefix,Fname,Lname,Email,Phone,Expertise,Notes FROM tblParty WHERE [role]= 2 ORDER BY 1 DESC "
				rs.Open Q,LoginDSN,1,3
				'rs.MoveLast   ' Is this a valid method for finding the latest record?
				partyid=rs("PartyID")
				' Read the data back and display:
				Response.Write("<table bgcolor='#ffefd5' border=1 cellspacing=0 cellpadding=4 width=650><tr><td>")
				Response.Write("<small>ID=" & partyid & "</small><br>")
				fullname = rs("Prefix") & " " & rs("Fname") & " " & rs("Lname")
				' Have to remove extra apostrophes: 
				fullname = Replace(fullname,"''","'")
				Response.Write ("<p>Name: <b>" & fullname  & "</b></p>")
				Response.Write ("Email: <a href='mailto:" & rs("Email") & "'>" & rs("Email") & "</a><br>")
				Response.Write ("Phone: " & rs("Phone") & "<br>")
				Response.Write ("Expertise: " & rs("Expertise") & "<br>")
				Response.Write ("Notes: " & rs("Notes"))
				Response.Write("</td></tr></table>")
				set rs = nothing
				'Response.Write ("<p>Please review data for accuracy.</b><br>If incorrect, <a href='confirm.asp?id=" & loginid & "&partyid=" & partyid & "'>delete</a> and re-enter.</p>")
			end if   ' diagnostic
			' Construct an email message to send to the new party.
			fullname = Replace(fullname,"''","'")
            header = "<html><head><title>Welcome to Author</title></head><body>"
            M = "<p font family='Times New Roman,Times,Serif'>Dear "
            M = M & fullname
            M = M & ":<br /><br />Thank you for your request to submit a manuscript for "
            M = M & "<em>Perspectives on Science and Christian Faith</em>.<br /><br />"
            M = M & "To submit a manuscript or other files, you may log in to the "
            M = M & "ASA Editor Web page at<br />"
            M = M & "<a href='http://www.asaeditor.org/'>http://www.asaeditor.org/</a> .  "
            M = M & "Enter your email address and <em>this password:<br /></em><br /><b>"
            M = M & pw
            M = M & "</b><br /><br />You may log in and change your password if you wish. "
            M = M & "<br /><br />If this email reached you in error, or you would like to "
            M = M & "cancel your request, please<br />"
            M = M & "<a href='http://www.asaeditor.org/unsubscribe.asp'>click here</a> "
            M = M & "to unsubscribe.<br /><br />If you have any other questions, please write to me."
            M = M & "<br /><br />Thanks for your support of the American Scientific Affiliation!"
            M = M & "<br /><br />Sincerely,<br /><br /><em><strong>"
            M = M & editorname & " "
            M = M & "</strong></em><br />Editor,<br />"
            M = M & "<em>Perspectives on Science and Christian Faith</em><br />"
            M = M & "American Scientific Affiliation<br /></p>"
            footer = "</body></html>"
            dim serverpath
            serverpath = Server.MapPath("addauthor.asp")
            if left(serverpath,2)="c:" or diagnostic=true then
	            ' local web site; CDONTS not available
            else
	            ' Send a welcome email to the new party:
	            ' Use CDONTS
	            dim subject,M,cc,bcc,importance,objMail
	            subject=fullname & " - ASA Author rights have been granted to you" 
	            set objMail = CreateObject("CDONTS.NewMail")
	            with objMail
		            .From = "paul@arveson.com"
		            .To = email
	                .Cc = ""
	                .Bcc = ""
		            .Subject = subject
		            .Importance = 2   ' High=2; normal=1
		            .MailFormat = 0  ' 0 = MIME; 1 = plain text
		            .BodyFormat = 0  ' HTML=0; text=1
		            .Body = header & M & footer
		            .Send
	            end with
	            set objmail = nothing
            end if ' server location
            Response.Write("<p><b>The following message was sent to " & email & ":</b></p>")
            Response.Write("<table bgcolor='fffacd' border=1 cellspacing=0 cellpadding=4 width=650><tr><td>" & M & "</td></tr></table>") 
		end if  ' party already in database
	end if ' errors in form, errn>0
end if   '  form submitted

function makePassword(byVal maxLen)
	' Random password generator
	' by rob@aspfree.com = Robert Chartier 
	' Usage: newpw = makePassword(10) for a pw of length 10 chars.
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
	    whatsNext = Int((1 - 0 + 1) * Rnd + 0)
		If whatsNext = 0 Then
		'character
		     upper = 90
		     lower = 65
		Else
		     upper = 57
		     lower = 48
		End If
		strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
    makePassword = strNewPass
end function

Function CheckMail(email)
	' Function to check email addresses
	' Usage: CheckMail(email,redirectpath)
	' Many thanks to http://www.aspsmith.com/re
	Dim objRegExp, blnValid
	'create a new instance of the RegExp object
	' note we do not need Server.CreateObject("")
	Set objRegExp = New RegExp
	'this is the pattern we check:
	objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
	'store the result either true or false in blnValid
	'blnValid = objRegExp.Test(email)
	CheckMail = objRegExp.Test(email)
	'If Not blnValid Then
		'do this if it is an invalid email address 
		'Response.Redirect("addreviewer.asp?id=" & loginid)
	'End If 
End Function
%>
<br>

<p><A href="menu.asp?id=<%=loginid%>">Return to Admin menu page</A></p>

</BODY>
</HTML>


