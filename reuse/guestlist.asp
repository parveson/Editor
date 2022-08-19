<%@Language="VBScript"%>
<%option explicit
Response.expires=30
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = false %>
<!-- #include file="test2.asp" -->
<!-- Note: the following variables were previously defined in test2.asp:  -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN  -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional //EN">
<html>
<head>
<title>List of Guest Vendors</title>
<link rel="stylesheet" type="text/css" href="../../../EntryStyle.css">
</head>
<body>
<% 
' Guestlist.asp - This page lists all members in the Vendors database in one long table, 
' and allows the administrator to approve some for public viewing.
dim sDSN
'datapath = Server.MapPath("\database")
'DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath & "\"
'DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
'sDSN = DSN1 & "Vendors.mdb" & DSN2
sDSN=Application("cnSQL_ConnectionString")
dim rs,i,j,k,q,count,ck,approval,formvalue,recordcount
dim diagnostic
diagnostic=false
if not Request.Form("submitted") then
	%>
	<h3 align="center">Select a company to examine:</h3>
	<form method="post" action="guestlist.asp?id=<%=loginid%>">
	<table align="center" border="1" cellpadding="2" cellspacing="0">
	<tr>
		<th>Approved?</th>
		<th>Company</th>
		<th>State</th>
		<th>Country</th>
		<th>Last Updated</th>
	</tr>
	<% 
	' Select all records:
	' Thanks to http://www.asp101.com/samples/db_getrows.asp
	Dim cnnGetRows   ' ADO connection
	Dim strDBPath    ' Path to our Access DB (*.mdb) file
	Dim arrDBData    ' Array that we dump all the data into
	dim refid(1000)
	Dim iRecFirst, iRecLast
	Dim iFieldFirst, iFieldLast
	recordcount = Request.Form("recordcount")
	set cnnGetRows = Server.CreateObject("ADODB.Connection")
	cnnGetRows.Open sDSN
	set rs=cnnGetRows.Execute("SELECT * FROM [tblVendors]")
	arrDBData = rs.GetRows(-1,0,Array("ID","ViewRecord","Company","State","Country","RefDate"))
	rs.Close
	cnnGetRows.Close
	set cnnGetRows=nothing			
	iRecFirst   = LBound(arrDBData, 2)
	iRecLast    = UBound(arrDBData, 2)
	iFieldFirst = LBound(arrDBData, 1)   '  = 0
	iFieldLast  = UBound(arrDBData, 1)
	' Display a table of the data in the array.
	' We loop through the array displaying the values.
	' Loop through the records (second dimension of the array)
	recordcount=0
	For I = iRecFirst To iRecLast
		recordcount = recordcount + 1
		' A table row for each record
		if recordcount mod 2 = 1 then
			Response.Write ("<tr bgcolor='linen'>")
		else
			Response.Write ("<tr bgcolor='aliceblue'>")
		end if 					
		' Columns: Loop through the fields (first dimension of the array)
		For J = 0 To iFieldLast
			if J=0 then
				' Don't tabulate; just get ID number for record
				refid(I) = arrDBData(J,I)
				Response.Write("<input type=hidden name=refid value=" & refid(I) & ">")
			else
				' A table cell for each field
				if J=1 then
					Response.Write ("<td><input type=checkbox name=viewrecord" & refid(I))
					if arrDBData(J,I) = "True" then 
						Response.Write (" checked></td>")
					else
						Response.Write (" ></td>")
					end if
				else
					if J=2 then
						' Company name and hyperlink
						Response.Write ("<td><a href='guestdetails.asp?id=" & loginid & "&refid=" & refid(I) & "'><b>" & arrDBData(J, I) & "</b></a></td>")
					else 
						' J > 2
						Response.Write ("<td><small>" & arrDBData(J, I) & "</small></td>")
					end if
				end if
			end if
		Next ' J				
		Response.Write "</tr>" 
	Next ' I
	Response.Write("</table>")
	%>
	</table><p><%=recordcount%> records found.</p>
	<input type="hidden" name="recordcount" value="<%=recordcount%>">
	<input type="hidden" name="submitted" value="true">
	<input type="submit" value="Approve Vendors">
	<br>                                                    
	</form>
	<%
else 
	' Update database with approved records set viewable:
	dim item,slen,formname,intloop,id,vid
	if diagnostic then
	  ' Tabulate all the form variables
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
	' First set the viewrecord fields to False for all records
	recordcount = Request.Form("recordcount")
	set rs=Server.CreateObject("ADODB.Recordset")
	for each Item in Request.Form
		formname = Cstr(Left(Item,5))
		if formname = "refid" then  
			count=Request.Form(Item).Count
			for intloop=1 to recordcount
				id = CInt(Request.Form(Item)(intloop))   ' value
				Q = "UPDATE [tblVendors] SET [ViewRecord] = " & CInt(0) & " WHERE [ID] = " & id
				if diagnostic then 
					Response.Write ("<p>Query: " & Q & "</p>")
				else
					rs.Open Q, sDSN,3,3
				end if
			next
		end if
	next
	' Update database with approved records set viewable:
	' Now loop through the records and set viewrecord = True when form checkbox is on:
	for each Item in Request.Form
		formname = Cstr(Left(Item,10))
		if formname = "viewrecord" then  '  All the "on" checkboxes only
			slen = len(Item)
			id = Cstr(Right(Item,slen-10))   ' strip off ID number from ViewRecord form data
			Q = "UPDATE [tblVendors] SET [ViewRecord] = " & CInt(1) & " WHERE [ID] = " & id
			if diagnostic then 
				Response.Write("<br>ID=" & id & " set to true.")
			end if
			rs.Open q,sDSN,3,3
		end if
	next
	set rs = nothing
	Response.Write("<p><strong>Records have been updated.</strong></p>")
	Response.Write("<p><a href='guestlist.asp?id=" & loginid & "'>Examine list again</a></p>")
end if
%>

<p><a href="../scraps/admin.asp?id=<%=loginid%>">Return to admin page</a></p>

</body>
</html>
