<%@ Language=VBScript %>
<%option explicit
' Process the data and store in Login2 database:
dim first,last,occupation,country,role,status,email,pw,entrytime,entrydate
dim datapath,LoginDSN,DSN1,DSN2,Q,rs,dom,d3,diagnostic,comment
diagnostic=false
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Thank you!</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<link rel="stylesheet" type="text/css" href="../../sheet2.css">
</head>
<body>
<div align="center">
<img height="80" alt="Balanced Scorecard Institute" src="../../images/bsci_logo.JPG" width="131"> 
<br>
<%
'datapath = Server.MapPath("\database")
'datapath = datapath & "\"
'DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath
'DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
'LoginDSN = DSN1 & "Login2.mdb" & DSN2
LoginDSN=Application("cnSQL_ConnectionString")
first = Left(Replace(Request.Form("FirstName"),"'","''"),20)
if first="" then 
	first=" " 
end if
last = Left(Replace(Request.Form("LastName"),"'","''"),30)
if last="" then
	last=" "
end if
email = LCase(Left(Request.Form("Email"),50))
occupation = Left(Request.Form("occupation"),1)
country = UCase(Left(Replace(Request.Form("Country"),"'","''"),20))
if country="" then
	country=" "
end if
pw = Left(Request.Form("Pw"),15)
entrytime = Left(Request.Form("entrytime"),20)
' If any errors, return to form
'Response.Write("<p>String input: " & email & "</p>")
'  Find org. domain:
dim at,domain,semail,dot,length,dotpos,cdom
length = len(email)
at = InStr(email,"@")
' Find country domain:
dot = InStrRev(email,".")
cdom = Right(email,length - dot + 1)
if len(cdom) > 3 then
	dom=cdom
else
	' country domains, 2 characters long
	semail=Left(email,length-3)
	dot = InStrRev(semail,".")
	dom = Right(semail,length-3-dot)
end if
' Check to see if user is already in the database.
dim recordcount,count
Q = "SELECT * FROM [tblLogins] WHERE [email] = '" & email & "'"
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open Q,LoginDSN,3,3
if not rs.EOF then 
	' Member email address is already in the database.
	Response.Redirect ("already.asp")
end if
entrydate = Now()
role=5   ' Member=default; roles can be changed by the administrator later
status=Cint(1)  ' -1 = initially presume member in good standing; 0=otherwise
comment = "v. beta 1"    '  255 characters text string for admin comments
' Store new member password and other data:
Q = "INSERT INTO [tblLogins] ([Fuser],[Luser],[Pw],[Email],[Dom],[Occupation],"
Q = Q & "[Country],[Entrytime],[Entrydate],[Role],[Status],[Comment]) "
Q = Q & " VALUES ('" & first & "','" & last & "','" & pw & "','" & email & "','" 
Q = Q & dom & "'," & occupation & ",'" & country & "','" & entrytime & "','"
Q = Q & entrydate & "'," & role & ",'" & status & "','" & comment & "');"					
if diagnostic then 
	Response.Write ("<p>" & Q & "</p>")
else
	' Open the database to insert the new record:
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open Q, LoginDSN,1,3
end if
'  Send an email to the user:
dim serverpath
serverpath = Server.MapPath("thankyou.asp")
if left(serverpath,2)="c:" or diagnostic=true then
	' local web site; CDONTS not available
else
	' Send a new user alert to the administrators:
	' Use CDONTS
	dim subject,recipient,M,cc,bcc,importance,objMail
	subject="Member joined"
	' Define recipients where contact form data should go:
	'recipient="julie@arveson.com;jam@howardrohm.com;paul@arveson.com;hrohm@mindspring.com"
	recipient="paul@arveson.com;paul@balancedscorecard.org"
	' Concatenate several variables into the message:
	M = "A new member entry was submitted by " & first & " " & last & Chr(13)
	M = M & "(" & email & ") on " & entrydate & Chr(13)
	set objMail = CreateObject("CDONTS.NewMail")
	with objMail
		.From = "system@balancedscorecard.org"
		.To = recipient
	    .Cc = ""
	    .Bcc = ""
		.Subject = subject
		.Importance = 2   ' High=2; normal=1
		.BodyFormat = 1  ' HTML=0; text=1
		.Body = M
		.Send
	end with
	set objMail=nothing
	' Send a welcome email to the new member:
	subject="Welcome to the Balance Scorecard Inst. Members Network" 
	M = "We received your request to join the Member's Network of the " & Chr(13)
	M = M & "Balanced Scorecard Institute. " & Chr(13)
	M = M & "You may log in to the Member's web site at http://www.balancedscorecard.org/members/" & Chr(13)
	M = M & " at any time using your email address and your password: " & Chr(13)
	M = M & pw & Chr(13) & Chr(13)
	M = M & "You may log in and change your password if desired. " & Chr(13)
	M = M & "If this email reached you in error, or you would like to unsubscribe, " & Chr(13) 
	M = M & "please reply to this email with 'unsubscribe' in the subject line." & Chr(13) & Chr(13)  
	M = M & "Best wishes for your success," & Chr(13) & Chr(13)
	M = M & "Balanced Scorecard Institute " & Chr(13)
	M = M & "975 Walnut St." & Chr(13)
	M = M & "Cary, NC 27511" & Chr(13)
	M = M & "(919) 420-8180" 
	set objMail = CreateObject("CDONTS.NewMail")
	with objMail
		.From = "info@balancedscorecard.org"
		.To = email
	    .Cc = ""
	    .Bcc = ""
		.Subject = subject
		.Importance = 2   ' High=2; normal=1
		.BodyFormat = 1  ' HTML=0; text=1
		.Body = M
		.Send
	end with
	set objmail = nothing
end if ' serversite
if first="" and last="" then
	Response.Write("<p><strong>Thank you for joining the Global Network of the Balanced Scorecard Institute!</strong></p>")
else
	Response.Write("<p><strong>Thank you, " & first & " " & last & ", for joining the Global Network of the Balanced Scorecard Institute!</strong></p>")
end if
%>
<p><em>We will notify you of upcoming events, new resources and other information. <br></em></p>

<p><strong><a href="../welcome.asp?id=<%=entrytime%>&amp;email=<%=email%>">Enter members site</a></p>

</div>

<br>
<br>
<br>
<br>
<br>
<br>

<p><a href="../../default.html"><strong>Return to home page</strong></a></p>

<!-- #include file="../footer.htm" --> 
</body>
</html>
