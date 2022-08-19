<%@Language=VBScript %>
<%option explicit
' Distribute request to appropriate users pages depending on role. 
' Login page determines the role of the user.  
' This is the only path into the secure pages.
' NO WRITING ON THIS PAGE!
dim loginid,diagnostic
dim rsAdmin,q,q2,email,pw,partyid,role,nopw
diagnostic=false
loginid = Trim(Left(Request.Form("timex"),20))
' Open the user database to check access list and find user permission type:
email = Lcase(Trim(Left(Request.Form("Email"),50)))
role = Request.Form("role")
CheckMail(email)
nopw=Cint(Request.Form("nopw"))  ' Valid user checked "send my password"
if nopw=1 then
	Response.Redirect("sendpw.asp?email=" & email & "&role=" & role)
	'Response.Write("<p>email=" & email & "</p>")
	'Response.Write("<form action='test.htm'><input type=submit name=continue></form>")
end if
pw = Trim(Left(Request.Form("Pw"),10))
if pw="" then
	Response.Redirect("default.htm?no_pwd")
end if
role = Cint(Request.Form("Role"))
set rsAdmin=Server.CreateObject("ADODB.Recordset")
dim datapath,LoginDSN,DSN1,DSN2
datapath = Server.MapPath("\database")
datapath = datapath & "\"
DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath
DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
LoginDSN = DSN1 & "Reviews.mdb" & DSN2
'LoginDSN=Application("cnSQL_ConnectionString")
q = "SELECT * FROM tblParty "
q = q & " WHERE [Email]='" & email & "' AND [Pw]='" & pw & "' AND [Role]=" & role
Response.Write ("<p>" & q & "</p>")
rsAdmin.Open q,LoginDSN
if rsAdmin.EOF then
	rsAdmin.Close
	set rsAdmin=nothing
	'  A match was not found.  Login failed.
    '  Response.Redirect "default.htm?" & q   ' Diagnostic
	Response.Redirect "default.htm?no_match2"     
end if
partyid = rsAdmin("PartyID")
rsAdmin.Close
'  Create a unique login ID for this user/session to use in query string:
'  In this simple case, the entry time (ms) is used.
'  (For a better id, a hash with the user IP address should be used.)
'  Store the session ID in the database:
'  (The last session time for each user will remain in database).
q2 = "UPDATE [tblParty] SET [Entrytime] = '" & loginid & "' "
q2 = q2 & "WHERE [PartyID] = " & partyid
if not diagnostic then
	rsAdmin.Open q2,LoginDSN,3,3
	set rsAdmin=nothing
	' Depending on the role, send the user to the appropriate menu page.
	select case role
	case 1   ' Administrators and editors
		Response.Redirect("menu.asp?id=" & loginid)
	case 2    ' Authors
		Response.Redirect("authors/menu.asp?id=" & loginid)
	case 3     ' Reviewers
		Response.Redirect("reviews/menu.asp?id=" & loginid)
	case else
		Response.Redirect "default.htm?unknown_id=" & loginid
	end select
end if

Function CheckMail(email)
	'our function to check email addresses
	' many thanks to http://www.aspsmith.com/re
	Dim objRegExp, blnValid
	'create a new instance of the RegExp object
	' note we do not need Server.CreateObject("")
	Set objRegExp = New RegExp
	'this is the pattern we check:
	objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
	'store the result either true or false in blnValid
	blnValid = objRegExp.Test(email)
	If Not blnValid Then
		'do this if it is an invalid email address 
		Response.Redirect "default.htm?invalid_email"
	End If 
End Function
%>

		