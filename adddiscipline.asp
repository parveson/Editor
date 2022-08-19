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
<TITLE>Add a Discipline</TITLE>
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
<h3>Add a Discipline</h3>
<%
dim diagnostic,email
diagnostic=false
If Request.Form("submitted")<>"True" then
	' Form data have not been submitted
	%>
	<H4>Enter a New Discipline</H4>
	<P><I><STRONG>Note: the Discipline Name field must be entered.</STRONG></I></P>
	<FORM method="post" action="adddiscipline.asp?id=<%=loginid%>">
		<TABLE border=0>
		<TR>
		    <TD align="right">Discipline Name:</TD>
		    <TD><INPUT NAME="Disname" size="20">*</TD>
		</TR>
		<TR>
		    <TD align="right">Description:</TD>
		    <TD><textarea name="disdesc" rows=4 cols=60></textarea>
		    </TD>
		</TR>
	  <TR>
		<td><INPUT TYPE="submit" VALUE="Submit Data">
		</TD>
		<td align="bottom">
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
	dim disname,disdesc
	disname = left(Replace(Request.Form("DisName"),"'","''"),20)
	disdesc = left(Replace(Request.Form("DisDesc"),"'","''"),255)
	'  Check for errors in data entry
	dim errn
	errn=0
	if disname = "" then 
		errn=errn+1
		Response.Write("<br>Name field is blank.")
	end if
	' If any errors, return to form
	if errn>0 then
		Response.Write("<p><b>There is a missing or invalid entry on your form.</b><br>")
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
		dim rs,i,Q,sDSN ' other variables are defined in test2.asp   
		datapath = Server.MapPath("\database")
		datapath = datapath & "\"
		DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath
		DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
		sDSN = DSN1 & "Reviews.mdb" & DSN2
		' Verify that this entry is not already in the database:
		Q = "SELECT * FROM [tblDiscipline] WHERE [DisName]= '" & disname & "'"
		set rs=server.CreateObject("ADODB.recordset")
		rs.Open Q,sDSN,1,3
		if not rs.EOF then
			Response.Write("<p><b>This Discipline is already in the database!</b></p>")
			rs.Close
			set rs=nothing
		else
			rs.Close
			Q = "INSERT INTO [tblDiscipline] ([DisName],[DisDesc]) "
			Q = Q & " VALUES ('" & disname & "','" & disdesc & "')"				
			if diagnostic then 
				Response.Write ("<p>" & Q & "</p>")
			else
				' Open the database to insert the new record:
				rs.Open Q,sDSN,1,3
				' This stores the new record and closes the recordset
				Response.Write "<p><b>New data have been stored.</b></p>"
				Response.Write("<h5>Current data in this record:</h5>")
				' Read the data back and display:
				Q = "SELECT * FROM [tblDiscpline] "
				rs.Open Q,sDSN,1,3
				rs.MoveLast   ' Is this a valid method for finding the latest record?
				Response.Write ("<p>Discipline Name: <b>" & rs("DisName") & "</b></p>")
				' Have to remove all extra apostrophes to make it work. 
				Response.Write ("<p>Description: " & rs("DisDesc") & "</p>")
				set rs = nothing
			end if   ' diagnostic
		end if  ' already in database
	end if ' errors in form, errn>0
end if  ' form submitted
%>
<br>

<p><A href="menu.asp?id=<%=loginid%>">Return to Admin menu page</A></p>

</BODY>
</HTML>


