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
<TITLE>Add an Editor</TITLE>
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
	<H4>Enter Editor Information</H4>
	<P><I><STRONG>All fields marked with an asterisk (*) must be entered.</STRONG></I></P>
	<FORM method="post" action="addeditor.asp?id=<%=loginid%>">
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
		    <TD align="right">Notes:</TD>
		    <TD><textarea name="notes" rows=4 cols=60></textarea>
		    *</TD>
	  </TR>
	  <TR>
		<td><INPUT TYPE="submit" VALUE="Submit Data">
		</TD>
		<td valign="bottom">
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
	' trim and replace quotes:
	dim prefix,fname,lname,phone,fullname,pw,entrydate,expertise,notes 
	prefix = left(Request.Form("Prefix"),10)
	fname = left(Replace(Request.Form("FName"),"'","''"),20)
	lname = left(Replace(Request.Form("LName"),"'","''"),30)
	fullname = prefix & " " & fname & " " & lname   
	phone = left(Replace(Request.Form("Phone"),"'","''"),20) 		
	email = Lcase(left(Replace(Request.Form("Email"),"'","''"),50)) 
	expertise = "Editor"
	notes = left(Replace(Request.Form("Notes"),"'","''"),255)
	'  Check for errors in data entry
	dim errn
	errn=0
	if fullname = "" then 
		errn=errn+1
		Response.Write("<br>Name fields are blank.")
	end if
	if phone="" then
		errn=errn+1
		Response.Write("<br>Phone field is blank.")
	end if
	if email="" then
		errn=errn+1
		Response.Write("<br>Email field is blank.")
	end if
	if not CheckMail(email) then
		errn=errn+1
		Response.Write("<br>Email address is not valid.")
	end if
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
		dim refdate,quot,partyid,role
		entrydate= Now()
		pw = makePassword(6)
		role=1   '  Editors and Administrators=1 - numeric value
		dim rs,i,Q
		' Before storing data, verify that the party & role is not already stored:
		Q = "SELECT * FROM [tblParty] WHERE [Role]= 1"
		Q = Q & " AND [Email] = '" & email & "'"
		set rs=server.CreateObject("ADODB.recordset")
		rs.Open Q,LoginDSN,1,3
		if not rs.EOF then
			Response.Write("<p><b>This Editor is already in the database!</b></p>")
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
				' Read the data back and display:
				'Q = "SELECT * FROM [tblParty] WHERE [role]= 1"
				Q = "SELECT TOP 1 PartyID,Prefix,Fname,Lname,Email,Phone,Notes FROM tblParty WHERE [role]= 1 ORDER BY 1 DESC "
				rs.Open Q,LoginDSN,1,3
				'rs.MoveLast   ' Is this a valid method for finding the latest record?
				partyid=rs("PartyID")
				Response.Write("<table bgcolor='#ffffff' border=1 width=650><tr><td>")
				Response.Write("<small>ID=" & partyid & "</small><br>")
				Response.Write ("<p>Name: <b>" & rs("Prefix") & " " & rs("Fname") & " " & rs("Lname") & "</b></p>")
				' Have to remove all extra apostrophes to make it work. 
				Response.Write ("Email: <a href='mailto:" & rs("Email") & "'>" & rs("Email") & "</a><br>")
				Response.Write ("Phone: " & rs("Phone") & "<br>")
				Response.Write ("Notes: " & rs("Notes"))
				Response.Write("</td></tr></table>")
				set rs = nothing
				Response.Write ("<p>Please review data for accuracy.</b><br>If incorrect, <a href='confirm.asp?id=" & loginid & "&partyid=" & partyid & "'>delete</a> and re-enter.</p>")
			end if   ' diagnostic
			' Send an email notification of the addition to the new party.
	        M = "Your name has been included in the list of editors " & Chr(13)
	        M = M & "for the American Scientific Affiliation. " & Chr(13)
	        M = M & "You may log in to the ASA Editor's login page at " & Chr(13) 
	        M = M & "http://www.asaeditor.org/" & Chr(13)
	        M = M & "using your email address " & email & " and your password: " & Chr(13)
	        M = M & pw & Chr(13) & Chr(13)
	        M = M & "You may change your password if desired. " & Chr(13)
	        M = M & "Thanks for your support of ASA!" & Chr(13) & Chr(13)
	        M = M & "Sincerely," & Chr(13) & Chr(13)
	        M = M & "Dr. Roman Miller, Editor " & Chr(13)
	        M = M & "Perspectives in Science and Christian Faith" & Chr(13)
	        M = M & "American Scientific Affiliation" & Chr(13)
	        dim serverpath
	        serverpath = Server.MapPath("addeditor.asp")
	        if left(serverpath,2)="c:" or diagnostic=true then
		        ' local web site; CDONTS not available
	        else
		        ' Send a welcome email to the new member:
		        ' Use CDONTS
		        dim subject,M,cc,bcc,importance,objMail
		        subject="Welcome to the ASA Editors Team!" 
		        set objMail = CreateObject("CDONTS.NewMail")
		        with objMail
			        .From = "paul@arveson.com"
			        .To = email
		            .Cc = ""
		            .Bcc = ""
			        .Subject = subject
			        .Importance = 2   ' High=2; normal=1
			        .MailFormat = 1  ' 0 = MIME; 1 = plain text
			        .BodyFormat = 1  ' HTML=0; text=1
			        .Body = M
			        .Send
		        end with
		        set objmail = nothing
	        end if ' serversite
	        Response.Write("<p><b>The following message was sent to " & email & ":</b></p>")
	        Response.Write("<p>" & M & "</p>")
		end if  ' party already in database
	end if ' errors in form, errn>0
end if   '  form submitted

function makePassword(byVal maxLen)
	' Random password generator
	' by <A href="mailto:rob@aspfree.com">Robert Chartier</A>
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


