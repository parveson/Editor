<%@ Language=VBScript%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = false 
%>
<HTML>
<HEAD>
<title>Export All Member Data</title>
</HEAD>
<BODY>
<!-- #include file="test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional //EN">
<html>
<head>
</HEAD>
<BODY>
<% 
' This page lists all members records in one long page, allowing the
' administrator to change status of each if desired. 
dim diagnostic
diagnostic=false
dim rs,i,k,q,count,ck,refid,block
dim formvalue,recordcount,occ,country,dom,role
dim lname,fname,fullname,email,entrydate,sent
dim occlabel,rolabel  
sent=""
' Select all records:
q = "SELECT * FROM [tblLogins] " 
set rs=Server.CreateObject("ADODB.Recordset")
LoginDSN=Application("cnSQL_ConnectionString")
rs.Open q,LoginDSN,3,3
recordcount=rs.RecordCount
count=0
Response.Write("<p><font size=1>")
rs.MoveFirst
do while not rs.EOF
	block=rs("Status")   
	if block=false then   ' Skip blocked members
		count = count + 1
		email=LCase(rs("Email"))
		fname=rs("Fuser")
		lname=rs("Luser")
		dom=LCase(rs("Dom"))
		occ=rs("Occupation")
		entrydate=rs("Entrydate")
		role=rs("Role")
		Response.Write(email & "," & fname & ", " & lname & "," & dom & ",")
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
				occlabel="other"
			case else
				occlabel=""   ' This should not occur
		end select
		select case role
			case 0
				rolabel="oth"
			case 1
				rolabel="adm"
			case 2
				rolabel="aso"
			case 3
				rolabel="prt"
			case 4
				rolabel="vnd"
			case else
				rolabel="?"   ' This should not occur
		end select
		Response.Write(occlabel & "," & entrydate & "," & sent & "," & role)
		Response.Write("<BR>")
	end if		
	rs.MoveNext
loop
rs.Close
set rs=nothing
%>
</font></p>
	
</BODY>
</HTML>
