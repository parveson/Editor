<%@Language=VBScript %>
<%option explicit
%>
<html>
<head>
<title>Document Index</title>
<link rel="stylesheet" type="text/css" href="../upload/sheet4.css">
</head>
<script language="Javascript" type="text/Javascript">
<!--
// Make a new popup window:
function win(strPath)		
{
	window.open("" + strPath,"newWindowName","height=600,width=700,left=200,top=100,toolbar=yes,location=1,directories=yes,menubar=1,scrollbars=yes,resizable=yes,status=yes");
}
function printWindow() {
bV = parseInt(navigator.appVersion);
if (bV >= 4) window.print();
}
// -->
</script>
<body>
<div align="center">
<p align="center"><strong><font color="forestgreen">Balanced Scorecard Institute</strong></font></p>
<h3>Document Management System</h3>
<p><i>This table lists files currently stored for access by <br>
<%
Response.Write ("all consultants.<br>")
Response.Write("Click on a heading to sort the data by this heading.<br>")
Response.Write("Click on any file name to download it to your computer. </i></p>")
Response.Write("Note: if you wish to upload a file, you will have to log in <a href='../upload/default.htm'>here</a>.</p>")
dim datapath,sDSN,DSN1,DSN2
datapath = Server.MapPath("\database")
DSN1="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & datapath & "\"
DSN2=";Mode=Share Deny None;Extended Properties="""";Locale Identifier=1033;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
sDSN = DSN1 & "Documents2.mdb" & DSN2
%>
<p><a href="../authors/search.asp?t=4">Search for a document</a></p>
<%'  Sort the table by one of the columns
dim t,sortby,orderby,sc1,sc2,sc3,sc4,sc5,c,bypass
sortby=1
orderby="ORDER BY tblFile.Filename"   ' Default
if not isNull(Request.Querystring("sortby")) then
	sortby=Request.Querystring("sortby")
end if
sc1="lavender"
sc2="#ffffff"
sc3="#ffffff"
sc4="#ffffff"
sc5="#ffffff"
select case sortby
	case 1
	orderby="ORDER BY tblFile.FileID"
	sc1="lavender"
	case 2
	orderby="ORDER BY tblFile.Filename"
	sc2="lavender"
	sc1="#FFFFFF"
	case 3
	orderby="ORDER BY tblFile.Title"
	sc3="lavender"
	sc1="#FFFFFF"
	case 4
	orderby="ORDER BY tblFile.Author" 
	sc4="lavender"
	sc1="#FFFFFF"
	case 5
	orderby="ORDER BY tblFile.CreateDate"
	sc5="lavender"
	sc1="#FFFFFF"
end select
' Write table headings:
%>
<a name="top">&nbsp;</a>
<table width="85%" align="center" border="1" cellpadding="4" cellspacing="0" bgcolor="white">
<tr>
<th align="center" bgcolor="<%=sc1%>">
<a href="index.asp?t=4&amp;sortby=1#top">File<br>ID</a>
</th>
<th align="center" bgcolor="<%=sc2%>">
<a href="index.asp?t=4&amp;sortby=2#top">File<br>Name</a>
</th>
<th align="center" bgcolor="<%=sc3%>">
<a href="index.asp?t=4&amp;sortby=3#top">Long<br>Title</a>
</th>
<th align="center" bgcolor="<%=sc4%>">
<a href="index.asp?t=4&amp;sortby=4#top">Author</a>
</th>
<th align="center" bgcolor="<%=sc5%>">
<a href="index.asp?t=4&amp;sortby=5#top">Create<br>Date</a>
</th>
<th>Size<br>(kB)</th>
</tr>
<%
dim q,rs,fileid,title,filename,filesize,author,createdate,empl,status,desc
dim filetype,uploader,popup
q = "SELECT * from [tblFile] "
q = q & orderby
set rs=Server.CreateObject("ADODB.Recordset")
rs.Open q,sDSN
do while not rs.EOF
	fileid = rs("FileID")
	filename=Trim(rs("Filename"))
	title=rs("Title")
	filesize=rs("Filesize") 
	filetype = rs("Filetype")
	author=rs("Author")
	createdate=rs("CreateDate")
	empl=rs("Empl")  ' File permissions affect which users can see this row
	if CInt(t)<=CInt(empl) then
		Response.Write "<tr><td><small>" & fileid & "</small></td>"	
		if filetype = "text/plain" or filetype = "text/html" or filetype = "text/xml" then
		    ' Place file in a new browser window if it is text or html:
		    %>
		    <td><a href="javascript:win('../upload/files/<%=filename%>');"><small><%=filename%></small></a> 
		    <%
		else
			' Open download dialog box if any other type of file:
			Response.Write "<td><a href='../upload/files/" & filename & "'><small>" & filename & "</small></a></td>"
		end if
		Response.Write "<td align='center'><small>" & title & "</small></td>"
		Response.Write "<td align='center'><small>" & author & "</small></td>"
		Response.Write "<td align='center'><small>" & createdate & "</small></td>"
		Response.Write "<td align='center'><small>" & filesize & "</small></td>"
	end if
	rs.MoveNext
loop
rs.Close
%>
</table>
<br>
<br>
<p><a href="index.asp#top">Back to top</a></p>
<p><a href="../authors/search.asp?t=4">Search for a document</a></p>
<br>
<br>
<div align="center">
</body>
</html>
