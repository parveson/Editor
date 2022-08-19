<%
'  Include file for testing logins.
'  Insert this include statement before <HTML> in all secure pages:
' <!-- #include file="test2.asp"  -->
dim loginid,rsLogintest,qtest
dim datapath,DSN1,DSN2,LoginDSN
'datapath = Server.MapPath("\database")
'DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath & "\"
'DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
'LoginDSN = DSN1 & "Login2.mdb" & DSN2
LoginDSN=Application("cnSQL_ConnectionString")
loginid=Left(Request.QueryString("id"),15)   '  Login id
if isEmpty(loginid) or loginid="" then
   '  Page was entered by bypassing the form.
	 Response.Redirect "../default.asp?no_id"
else
	qtest = "SELECT [Entrytime] FROM tblLogins"
	qtest = qtest & " WHERE Entrytime = '" & loginid & "'"
	set rsLogintest=Server.CreateObject("ADODB.Recordset")
	rsLogintest.Open qtest,LoginDSN
	if rsLogintest.EOF then
			rsLogintest.Close
			set rsLogintest = nothing
			'  A match was not found.
		  Response.Redirect "../default.asp?no_match"
	end if
end if
'  Login test passed - proceed.
%>