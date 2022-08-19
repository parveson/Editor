<%@Language="VBScript"%>
<%option explicit
'Response.Expires=30
Response.AddHeader "pragma","no-cache"  '  Must not cache form data
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!-- #include file="test2.asp"  -->
<!-- The following variables are dimensioned in test2.asp:        -->
<!-- dim loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>View one Record</title>
<link rel="stylesheet" type="text/css" href="../../../EntryStyle.css">
</head>
<body>
<%
'  This page shows one record.
'  This page is called by admin/memberlist.asp.
dim diagnostic,refid
refid=CInt(Request.Querystring("refid"))  '  Record id of selected member
diagnostic=false
If not isEmpty(refid) then
	dim rs,i,q,quot
	dim email,occ,occlabel,dom,role,rolabel,fullname,fixname
	dim status,stat,entrydate,country
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
				rolabel="other"
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
				rolabel="??"  ' This should not occur
		end select
		Response.Write ("<br>Role: <b>" & rolabel & " </b>")	
		Response.Write ("<br>Join date: <b>" & entrydate & " </b>")
		if status=true then
			stat="blocked"
		else
			stat="ok"
		end if
		Response.Write("<br>Status: <b>" & stat & "</b></P>")
	end if  ' diagnostic
else
	Response.Write "<p>ID = " & refid & "</p>"
	Response.Write "<p>The ID (refid) does not exist.  There is an error in the code.</p>"
end if
Response.Write ("<p><b><a href='aedit.asp?id=" & loginid & "&refid=" & refid & "'>Edit this record</a></b></p>")
Response.Write ("<p><b><a href='confirm.asp?id=" & loginid & "&refid=" & refid & "'>Delete this record</a></b> (you will be asked to confirm)</p>")
Response.Write ("<p><b><a href='resetpw.asp?id=" & loginid & "&refid=" & refid & "'>Reset user's password</a></b></p>")

Response.Write ("<p><b><a href='admin.asp?id=" & loginid & "'>Return to admin page</a></b></p>")

%>

</body>
</html>
