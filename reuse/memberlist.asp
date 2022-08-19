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
<title>Editor's Review Status Board</title>
<link rel="stylesheet" type="text/css" href="../StyleSheet1.css">
</head>
<body>
<h3 align="center">Editor's Review Status Board</h3>


</tr>

<% 
' Memberlist.asp: This page lists all members records in one long page, allowing
' the administrator to change status of each if desired. 
' The following variables are defined in test2.asp:      
' loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN 
dim rs,i,k,q,count,ck,refid,block
dim diagnostic
diagnostic=false
dim formvalue,occ,country,domain,role,email
dim lname,fname,fullname,entrydate
dim occlabel,colabel,dolabel,rolabel 
if not Request.Form("submitted")then %> 
	<p><i><b>Click to view the Detail data.</b></i></p>
	<table border="1" cellspacing="0" cellpadding="3" ALIGN="CENTER">
	<form method="post" action="memberlist.asp?id=<%=loginid%>">
	<table align="center" border="1" cellpadding="2" cellspacing="0">
	<tr>
	<th>No.</th>
	<th>MS Title</th>
	<th>Author</th>
	<th>MS</th>
	<th>Rev. 1</th>
	<th>Rev. 2</th>
	<th>Rev. 3</th>
	<th>Close?</th>
	</tr>
	<% 
	' Select all records:
	q = "SELECT * FROM [tblManuscripts] " 
	'q = "SELECT [ID],[Fuser],[Luser],[Email],[Dom],[Occupation],[Country],"
	'q = q & "[Entrytime],[Entrydate],[Role],[Status],[Comment],[Logs]"
	'q = q & " FROM [tblLogins]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open q,LoginDSN,3,3
	count=0
	rs.MoveFirst
	do while not rs.EOF
		count = count + 1
		refid = rs("ID")
		fname=rs("Fuser")
		lname=rs("Luser")
		email=rs("Email")
		country=rs("Country")
		entrydate=rs("Entrydate")
		domain=rs("Dom")
		role = rs("Role")
		occ=rs("Occupation")
		block = rs("status") 
		fullname = fname & " " & lname
		if fullname=" " then    ' User did not enter name
			fullname = email
		end if
		select case occ
			case 1
				occlabel="?"   ' User did not enter an occupation
			case 2
				occlabel="private"
			case 3
				occlabel="public"
			case 4
				occlabel="nonprof"
			case 5
				occlabel="consult"
			case 6
				occlabel="student"
			case 7
				occlabel="unknown"
			case else
				occlabel="??"   ' This should not occur
		end select
		select case role
			case 1
				rolabel="adm"
			case 2
				rolabel="aso"
			case 3
				rolabel="prt"
			case 4
				rolabel="vnd"
			case 5
				rolabel="mem"
			case else
				rolabel="?"   ' This should not occur
		end select
				if count mod 2 = 1 then
			Response.Write ("<tr bgcolor='linen'>")
		else
			Response.Write ("<tr bgcolor='aliceblue'>")
		end if 
		Response.Write ("<td><small><a href='memberdetails.asp?id=" & loginid & "&refid=" & refid & "'>" & fullname & "</a></small></td>")
		Response.Write ("<td><small>" & email & "</small></td>")
		Response.Write ("<td><small>" & occlabel & " </small></td>")		
		Response.Write ("<td><small>" & country & " </small></td>")
		Response.Write ("<td><small>" & domain & " </small></td>")
		Response.Write ("<td><small>" & rolabel & " </small></td>")		
		Response.Write ("<td><small>" & entrydate & " </small></td>")
		Response.Write ("<td><input type=checkbox name=status" & refid)
		if block = true then
			Response.Write (" checked>")
		else
			Response.Write (" >")
		end if
		Response.Write ("</td></tr>")
		Response.Write("<input type=hidden name=refid value=" & refid & ">")		
		rs.MoveNext
	loop
	rs.Close
	set rs=nothing
	%>
	</table><p><%=count%> records found.</p>
	<input type="hidden" name="submitted" value="true">
	<input type="submit" value="Update Blocked Members">
	<br>                                                    
	</form>
	<%
else
	' Update database with approved records set viewable:
	dim item,slen,formname,intloop,id,vid
	if diagnostic then
	  Response.Write("<p>Number of form variables submitted = " & Request.Form.Count & "<br>")
		For each Item in Request.Form
			Response.Write("Element '" & item & "' Value = '" & Request.Form(Item) & "'<br>")
		Next
		Response.Write ("<p>----</p>")
		For each Item in Request.Form
			count = Request.Form(Item).Count 
			If count > 1 then
				Response.Write(Item & ":<br>")
				For intloop = 1 to count
					Response.Write ("Subkey " & intloop & " value = " & Request.Form(Item)(intloop) & "<br>")
				Next
			else
				Response.Write (Item & " = " & Request.Form(Item) & "<br>")
			end if
		Next	
		Response.Write ("<p>----</p>")
	end if
	' First find each ID field and set the corresponding checkbox field to False 
	set rs=Server.CreateObject("ADODB.Recordset")
	for each Item in Request.Form
		formname = Cstr(Left(Item,5))
		if formname = "refid" then  
			count=Request.Form(Item).Count
			for intloop=1 to count
				refid = CInt(Request.Form(Item)(intloop))   ' value
				Q = "UPDATE [tblLogins] SET [Status] = " & Cint(0) & " WHERE [ID] = " & refid
				if diagnostic then Response.Write("<p>" & Q & "</p>")
				rs.Open Q, LoginDSN,3,3
			next
		end if
	next 
	' Now loop through the records and set status = True when form checkbox is on:
	for each Item in Request.Form
		formname = Cstr(Left(Item,6))
		if formname = "status" then  '  All the "on" checkboxes only
			slen = len(Item)
			refid = Cstr(Right(Item,slen-6))   ' ID number, string variable
			Response.Write("<p>Block item: " & refid & "</p>")
			Q = "UPDATE [tblLogins] SET [Status] = " & Cint(1) & " WHERE [ID] = " & refid
			rs.Open Q, LoginDSN,3,3
		end if
	next
	set rs = nothing
	Response.Write("<p><strong>All records have been updated.</strong></p>")
	Response.Write("<p><a href='memberlist.asp?id=" & loginid & "'>Examine list again</a></p>")
	set rs = nothing
end if
%>
<p><a href="../scraps/admin.asp?id=<%=loginid%>">Return to admin page</a></p>
</body>
</html>
