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
<title>Add an Editor</title>
<link rel="stylesheet" type="text/css" href="../StyleSheet1.css">
<script LANGUAGE="JavaScript1.1" SRC="../../FormChek.js"></script>
<script LANGUAGE="JavaScript" type="text/JavaScript">
	<!--
function validate(form)
{   return (
      checkString(form.elements["Lname"],sWorldLastName,false) &&
      checkString(form.elements["Fname"],sWorldFirstName,false) &&
      checkInternationalPhone(form.elements["Phone"]) &&
      checkEmail(form.elements["Email"],false);
      form.submit();
    )
}
function printWindow() {
	bV = parseInt(navigator.appVersion);
	if (bV >= 4) window.print();
}
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=700,width=800,left=20,top=10,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
// Usage:<A href="javascript:win('http://www.ibm./com/');">IBM</a>
//  End -->
</script>
</head>
<body BGCOLOR="#ffffff">
<p class="small">Administration</p>
<h3>Add an Administrator</h3>
<p><i>NOTE: Administrators have access to all the files and text data.  Administrators 
will log in with their email address and password.</i></p>
<%
dim diagnostic
diagnostic=false
If Request.Form("submitted")<>"True" then
	' Form data have not been submitted
	%>
	<h4>Enter New Administrator Information</h4>
	<p><i><strong>Fields marked with an asterisk (*) must be entered.</strong></i></p>
	<form method="post" action="addadmin.asp?id=<%=loginid%>">
		<table border="0">
		<tr>
		    <td align="right">First Name:</td>
		    <td><input NAME="Fname" maxlength="20" onFocus="promptEntry(sWorldFirstName)" onChange="checkString(this,sWorldFirstName)">*</td>
		</tr>
		<tr>
		    <td align="right">Last Name:</td>
		    <td><input NAME="Lname" size="30" maxlength="30" onFocus="promptEntry(sWorldLastName)" onChange="checkString(this,sWorldLastName)">*</td>
		</tr>
		<tr>
		    <td align="right">Email:</td>
		    <td><input NAME="Email" size="50" maxlength="50" onFocus="promptEntry(pEmail)" onChange="checkEmail(this, true)">*</td>
		</tr>
		<tr>
		    <td align="right">Phone:</td>
		    <td><input NAME="Phone" maxlength="20" onFocus="promptEntry(pWorldPhone)" onChange="checkInternationalPhone(this)">*</td>
	  </tr>
		<tr>
		    <td align="right">Password: (Please make a note of it!)</td>
		    <td><input type="password" NAME="Pw1" size="15" maxlength="15">*</td>
		</tr>
		<tr>
		    <td align="right">Password: enter again to confirm:</td>
		    <td><input type="password" NAME="Pw2" size="15" maxlength="15">*</td>
		</tr>
	<tr>
		<td align="right">Comment:</td>
			<td><textarea NAME="Comment" rows="4" cols="60"></textarea>
		</td>
	</tr>

	  <tr>
		<td><input TYPE="submit" VALUE="Submit Data" onClick="validate(this.form)">
		</td>
		<td>
			<br><br><br>
			&nbsp;&nbsp;&nbsp;&nbsp;
			<input TYPE="reset" VALUE="Reset All">
			<input type="hidden" name="submitted" value="True">
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
	dim item
	if diagnostic then
	  Response.Write("<p>Number of form variables submitted = " & Request.Form.Count & "<br>")
		For each Item in Request.Form
			Response.Write("Element '" & item & "' Value = '" & Request.Form(Item) & "'<br>")
		Next
	end if
	' trim and replace quotes:
	dim fname,lname,phone,fullname,email,errn,pw1,pw2,comment
	fname = left(Replace(Request.Form("Fname"),"'","''"),20)
	lname = left(Replace(Request.Form("Lname"),"'","''"),30)
	fullname = fname & " " & lname   
	phone = left(Replace(Request.Form("Phone"),"'","''"),20) 		
	email = Lcase(left(Replace(Request.Form("Email"),"'","''"),50))
	Pw1 = left(Request.Form("Pw1"),15)
	Pw2 = left(Request.Form("Pw2"),15) 
	comment=left(Replace(Request.Form("Comment"),"'","''"),255)
	'  Check for errors in data entry
	errn=0
	Response.Write("<p>")
	if fullname = "" then 
		errn=errn+1
		Response.Write("<br>Name fields are blank.")
	end if
	if email="" then
		errn=errn+1
		Response.Write("<br>Email field is blank.")
	end if
	if phone="" then
		errn=errn+1
		Response.Write("<br>Phone field is blank.")
	end if
	if Pw1="" or Pw2="" then
		errn=errn+1
		Response.Write("<br>Password field is blank.")
	end if
	if not Pw1=Pw2 then
		errn=errn+1
		Response.Write("<br>Passwords do not match.")
	end if
	' If any errors, return to form
	if errn>0 then
		Response.Write("<p><b>There are missing entries in your form.</b><br>")
		Response.Write("If you don't know the data, you may insert a dash (-).</p>")
		Response.Write("<p><b>Please <a href='Javascript:history.back();'>TRY AGAIN</a></b></p>")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br />")
		Response.Write("<br /></p>")
	' Else store data:
	else
		' Define additional fields and data not contained on form:
		dim refdate,quot,refid,entrytime,entrydate,role,logs
		entrytime="0"
		entrydate= Now()
		role=1  '  Role in process - not used
		logs=0  '  Number of times new administrator logged in
		dim rs,i,Q
		dim sDSN   ' other variables are defined in test2.asp
		datapath = Server.MapPath("\database")
		datapath = datapath & "\"
		DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath
		DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
		sDSN = DSN1 & "Reviews.mdb" & DSN2
		'sDSN=Application("cnSQL_ConnectionString")
		'  Construct SQL statement; here is the table into which the data is being inserted
		Q = "INSERT INTO [tblEditors] ([Fuser],[Luser],[Pw],[Email],[Phone],[Entrytime],"
		Q = Q & "[Entrydate],[Role],[Comment],[Logs]) VALUES ('" & fname & "','" & lname & "','" & Pw1 & "','"
		Q = Q & email & "','" & phone & "','" & entrytime & "','" & entrydate & "'," 
		Q = Q & role & ",'" & comment & "'," & logs & ")"					
		if diagnostic then 
			Response.Write ("<p>" & Q & "</p>")
		else
			' Open the database to insert the new record:
			set rs=server.CreateObject("ADODB.recordset")
			rs.Open Q,sDSN,1,3
			' This stores the new record and closes the recordset
			Response.Write "<p><b>New data have been stored.</b></p>"
			Response.Write("<h5>Current data in this record:</h5>")
			' Read the data back and display for the user:
			Q = "SELECT * FROM [tblEditors]"
			rs.Open Q,sDSN,1,3
			rs.MoveLast   ' Is this a valid method for finding the latest record?
			refid=rs("EdID")
			Response.Write("<table bgcolor='#ffffff' border=1 width=650><tr><td>")
			Response.Write("<small>Record ID=" & refid & "</small><br>")
			Response.Write ("<p>Name: <b>" & rs("Fuser") & " " & rs("Luser") & "</b></p>")
			' Have to remove all extra apostrophes to make it work. 
			Response.Write ("Email: <a href='mailto:" & rs("Email") & "'>" & rs("Email") & "</a><br>")
			Response.Write ("Phone: " & rs("Phone") & "<br>")
			Response.Write ("Comments: " & rs("Comment"))
			Response.Write("</td></tr></table>")
			set rs = nothing
			Response.Write ("<p>Please review data for accuracy.</b></p>")
		end if   ' diagnostic
		' Display the detail data of the user:
		Response.Write ("<p>If you wish to edit these data, <a href='aedit.asp?id=" & loginid & "&refid=" & refid & "'>CLICK HERE</a><br>") 
		' Send an email notification of the new addition to the administrators.
	end if ' errors in form
end if   '  form submitted
%>
<br>

<p><a href="../menu.asp?id=<%=loginid%>">Return to Admin page</a></p>

</body>
</html>


