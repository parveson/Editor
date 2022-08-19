<%@Language="VBScript"%>
<%option explicit
Response.expires=5
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="test2.asp"  -->
<!-- The following variables are dimensioned in test2.asp:        -->
<!-- dim loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<html>
<head>
<title>Edit Member Data</title>
<link rel="stylesheet" type="text/css" href="../../../EntryStyle.css">
<script LANGUAGE="JavaScript1.1" SRC="../../FormChek3.js"></script>
</head>
<body BGCOLOR="#ffffff">
<h2>Edit Member Data</h2>
<%
'  Same as add.asp except a large number of edits is allowed, for testing.
dim diagnostic,refid
diagnostic=false
refid = Request.QueryString("refid")
If Request.Form("submitted")<>"True" then
	dim rs,i,q,quot
	dim email,occ,occlabel,dom,role,rolabel,fuser,luser
	dim status,stat,entrydate,country,occupation
	' select the requested record:
	q = "SELECT * FROM [tblLogins] WHERE [ID] = " & refid
	if diagnostic then
		Response.Write "<p><small>" & q & "</small></p>"
	end if
	if not diagnostic then
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.Open q,LoginDSN,3,3
		' Display most fields in the record:
		if not rs.EOF then
			Response.Write("<h5>Current Data:</h5>")
			refid = rs("ID")
			email=rs("Email")
			fuser = rs("Fuser") 
			luser = rs("Luser")
			occ=rs("Occupation")
			country=rs("Country")
			dom=rs("Dom")
			role = rs("Role")
			entrydate=rs("Entrydate")
			status=rs("Status")
			rs.Close
			set rs=nothing
		end if
	end if
	dim employment(7)
	employment(1)="student"
	employment(2)="private sector"
	employment(3)="public sector"
	employment(4)="nonprofit"
	employment(5)="consultant"
	employment(6)="other"
	employment(7)="unknown"
	dim rolename(5)
	rolename(1)="System Administrator"
	rolename(2)="Associate"
	rolename(3)="Course Attendee"
	rolename(4)="Vendor"
	rolename(5)="Member"
	' Form data have not been submitted
	%>
	<p><i>Fields marked with an asterisk (*) must be entered.</i></p>
	<form method="POST" action="aedit.asp?id=<%=loginid%>">
		<table border="0">
		<tr>
		    <td align="right">First Name:</td>
		    <td><input NAME="fuser" size="20" maxlength="20" value="<%=fuser%>">*</td>
		</tr>
		<tr>
		    <td align="right">Last Name:</td>
		    <td><input NAME="luser" size="20" maxlength="20" value="<%=luser%>">*</td>
		</tr>
		<tr>
		    <td align="right">Email:</td>
		    <td><input NAME="email" size="50" maxlength="50" value="<%=email%>">*</td>
		</tr>
		<tr>
		    <td align="right">Occupation: </td>
		    <td><select NAME="occ">
		    <%
		    for i=1 to 7
				if occ = i then
					Response.Write("<option value='" & i & "' selected>" & employment(i) & "</option>")
				else
					Response.Write("<option value='" & i & "'>" & employment(i) & "</option>")
				end if
			next
			%>
			</select>&nbsp;&nbsp;(originally <%=employment(occ)%>)
		</tr>
		<tr>
		    <td align="right">Country:</td>
		    <td><input name="country" size="50" value="<%=country%>">*</td>
		</tr>
		<tr>
		    <td align="right">Domain:</td>
		    <td><input NAME="Dom" size="20" maxlength="20" value="<%=dom%>">*</td>
		</tr>
	<tr><td align="right">Membership Role: </td>
		<td><select name="role">
		<%
		for i=1 to 5
			if role = i then
				Response.Write("<option value='" & i & "' selected>" & rolename(i) & "</option>")
			else
				Response.Write("<option value='" & i & "'>" & rolename(i) & "</option>")
			end if
		next
		%>
		</select>&nbsp;&nbsp;(originally <%=rolename(role)%>)
		</td>
	</tr>
	<tr>
	    <td>
			<input type="hidden" name="submitted" value="True">
			<input type="hidden" name="entrydate" value="<%=entrydate%>">
			<input type="hidden" name="refid" value="<%=refid%>">
			<input TYPE="submit" VALUE="Submit Data">
		</td><td>
			&nbsp;&nbsp;&nbsp;&nbsp;
			<input TYPE="reset" VALUE="Reset All">
		</td>
	</tr>
	</table>
	</form>
<%
else
	' Data have been submitted in the form.
	' This starts the second page, where the data will be updated.
	' Order of fields is not assumed.
	' List all the form fields for diagnostic:
	dim item,Q2,fullname
	if diagnostic then
	  Response.Write("<p>Number of form variables submitted = " & Request.Form.Count & "<br>")
		For each Item in Request.Form
			Response.Write("Element '" & item & "' Value = '" & Request.Form(Item) & "'<br>")
		Next
	end if
	' trim and replace quotes:  
	refid=Cint(Request.Form("refid"))
	fuser = Left(Replace(Request.Form("Fuser"),"'","''"),20)
	luser = Left(Replace(Request.Form("Luser"),"'","''"),30)
	email = Left(Request.Form("Email"),50)
	entrydate = Left(Request.Form("entrydate"),50) ' Preserve original date
	dom = Left(Request.Form("Dom"),5)  ' email domain
	occupation = Cint(Left(Request.Form("occ"),1))
	country = Left(Replace(Request.Form("Country"),"'","''"),20)
	role=Cint(Left(Request.Form("Role"),3))
	' NOTE: Variables Pw, Comment, Entrytime and Logs will not be changed here.
	set rs=Server.CreateObject("ADODB.Recordset")		
	' Do not do a delete and then an insert, because this will delete the old ID and create 
	' a new one that is not known.  Then you would have to find it somehow.
	' Instead, do an update on the existing record.  
	Q = "UPDATE [tblLogins] SET "
	Q = Q & " [Fuser] = '" & fuser & "',"
	Q = Q & " [Luser] = '" & luser & "',"
	Q = Q & " [Email] = '" & email & "',"
	Q = Q & " [Dom] = '" & dom & "',"
	Q = Q & " [Occupation] = " & occupation & ","
	Q = Q & " [Country] = '" & country & "',"
	Q = Q & " [Entrydate] = '" & entrydate & "',"
	Q = Q & " [Role] = " & role 
	Q = Q & " WHERE (([ID]) = " & refid & ")"
	if diagnostic then 
		Response.Write ("<p>" & Q & "</p>")
	else
		' This stores the new record and closes the recordset
		rs.Open Q, LoginDSN,3,3
		Response.Write "<p><b>New data have been inserted into the database.</b></p>"
		' Now read back the data in the record:
		Q2 = "SELECT * FROM [tblLogins] WHERE [ID] = " & refid
		' Read the data back and display for the user:
		' Display most fields in the record:
		rs.Open Q2, LoginDSN,3,3
		if not rs.EOF then
			Response.Write("<h5>Current Member Data:</h5>")
			refid = rs("ID")
			email=rs("Email")
			fullname = rs("Fuser") & " " & rs("Luser")
			occ=rs("Occupation")
			country=rs("Country")
			dom=rs("Dom")
			role = rs("Role")
			entrydate=rs("Entrydate")
			status=rs("Status")
			rs.Close
			set rs=nothing
		end if
		Response.Write "<p>ID = <b>" & refid & "</b>"
		Response.Write("<br>Member Name: <b>" & fullname & "</b>")
		Response.Write("<br>Email: <b>" & email & "</b>")
		select case occ
			case 1
				occlabel="student"   ' User did not enter an occupation
			case 2
				occlabel="private"
			case 3
				occlabel="public"
			case 4
				occlabel="nonprofit"
			case 5
				occlabel="consultant"
			case 6
				occlabel="other"
			case 7
				occlabel="?"
			case else
				occlabel="??"   ' This should not occur
		end select
		Response.Write ("<br>Occupation: <b>" & occlabel & " </b>")		
		Response.Write ("<br>Country: <b>" & country & " </b>")
		Response.Write ("<br>Domain: <b>" & dom & " </b>")
		select case role
			case 0
				rolabel="other"  ' no role defined
			case 1
				rolabel="administrator"
			case 2
				rolabel="associate"
			case 3
				rolabel="partner"
			case 4
				rolabel="vendor"
			case 5
				rolabel="member"   
			case else
				rolabel="??"   ' This should not occur
		end select
		Response.Write ("<br>Role: <b>" & rolabel & " </b>")	
		Response.Write ("<br>Join date: <b>" & entrydate & " </b>")
		if status=true then
			stat="blocked"
		else
			stat="ok"
		end if
		Response.Write("<br>Status: <b>" & stat & "</b></P>")
		Response.Write("<br>Last Update: " & entrydate & "<br /></p>")
	end if   ' diagnostic
	Response.Write("<p><a href='aedit.asp?id=" & loginid & "&refid=" & refid & "'><b>Edit again.</b></a></p>")
end if   '  form submitted
%>

<p><a href="memberlist.asp?id=<%=loginid%>"><b>View the members list.</b></a></p>

</body>
</html>
