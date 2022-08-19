<%@ Language=VBScript %>
<%option explicit%>
<!-- #include file="../test2.asp"  -->
<html>
<head>
<title>Upload a File</title>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
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
</head>
<body>
<%
dim diagnostic,email
diagnostic=false
email = Request.QueryString("email")  ' email address of person uploading this file = uploader
Server.ScriptTimeout=1200      ' uploads can take up to 1200 sec. or 20 minutes.
'  IF YOU CHANGE THIS PLEASE CHANGE THE HTML BELOW AND CHANGE TIMEOUT ON ackupload.asp
%>
<h3 align="center">Upload a File</h3>
<P align=left>The following types of manuscripts may be submitted:</P>
<table border=1 cellspacing=0 width=400 align=center bgcolor="lemonchiffon">
<tr><th>Manuscript Category</th><th>Maximum text words</TD></tr>
 <tr><td>Regular article (Requires submission of abstract 50-150 words as separate 
  document)</td><td align=middle>6000</td></tr>
<tr><td>Communications</td><td align=middle>2700</td></tr>
<tr><td>Early Career (Requires abstract)</td><td align=middle>6000</td></tr>
<tr><td>News and Views</td><td align=middle>1500</td></tr>
<tr><td>Letter </td><td align=middle>700</td></tr>
<tr><td>Art</td><td align=middle>300</td></tr>
</table>

<P>Any common file format is acceptable. Maximum file size is 32 MB.  Large files can 
take a long time to load; a maximum of 20 minutes is allowed.
If your upload lasts longer than this it will timeout. If you have any problems please 
<A href="../contact.asp">contact us.</A>&nbsp;&nbsp;
<i><b>Note: all fields must be filled in unless labeled "optional".</b></i>
<br>
</P>
<form enctype="multipart/form-data" method="post" action="ackupload.asp?id=<%=loginid%>">

<!--<form method="post" action="ackupload.asp?id=<%=loginid%>">-->
	<table WIDTH="95%" border=0>
	<tr>
	   <td align="right"><STRONG>File:</STRONG></td>
	   <td ALIGN="left"><input TYPE="file" NAME="MyFile" size="40"></td>
	</tr>
	<tr><td align="right"><STRONG>Enter a title for this submission:</STRONG></td>
		<td><input name="Title" size="80"></td>
	</tr>
	<tr><td align="right"><STRONG>Manuscript Category:</STRONG> </td>
		<td><select name="TypeID">
			<option value=0 selected>Select one:</option>
			<option value=1>Regular article</option>
			<option value=2>Abstract</option>
			<option value=3>Communications&nbsp;</option>
			<option value=4>Early Career</option>
			<option value=5>News and Views&nbsp;</option>
			<option value=6>Letter</option>
			<option value=7>Art or Poetry</option>
			<option value=8>Other</option>
			</select>
		</td>
	</tr>
	<tr><td align="right"><STRONG>Main subject area or discipline of this file:</STRONG></td>
		<td>
		<!-- PULL THESE FROM THE DATABASE!   YOU CAN USE MORE THAN ONE IF THIS IS DEVELOPED -->
			<select name="DisID">
			<option value=0 selected>Select one:&nbsp;</option>
			<option value=1>agriculture  </option>
			<option value=2>anatomy  </option>
			<option value=3>animal behavior  </option>
			<option value=4>anthropology  </option>
			<option value=5>archaeology  </option>
			<option value=6>astrology  </option>
			<option value=7>astronomy  </option>
			<option value=8>Biblical studies  </option>
			<option value=9>biochemistry  </option>
			<option value=10>bioethics  </option>
			<option value=11>biography  </option>
			<option value=12>biology  </option>
			<option value=13>biophysics  </option>
			<option value=14>botany  </option>
			<option value=15>biotechnology  </option>
			<option value=16>brain research  </option>
			<option value=17>broadcasting  </option>
			<option value=18>business  </option>
			<option value=19>chemistry  </option>
			<option value=20>church history  </option>
			<option value=21>communications  </option>
			<option value=22>computer science  </option>
			<option value=23>cosmology  </option>
			<option value=24>counseling  </option>
			<option value=25>ecology  </option>
			<option value=26>economics  </option>
			<option value=27>education  </option>
			<option value=28>engineering  </option>
			<option value=29>environmental science  </option>
			<option value=30>ethics  </option>
			<option value=31>ethnology  </option>
			<option value=32>ethology  </option>
			<option value=33>evangelism  </option>
			<option value=34>fine arts  </option>
			<option value=35>genetics  </option>
			<option value=36>genetic engineering  </option>
			<option value=37>geography  </option>
			<option value=38>geology  </option>
			<option value=39>geophysics  </option>
			<option value=40>government  </option>
			<option value=41>hermeneutics  </option>
			<option value=42>historiography  </option>
			<option value=43>history  </option>
			<option value=44>history of science  </option>
			<option value=45>information science  </option>
			<option value=46>journalism  </option>
			<option value=47>jurisprudence  </option>
			<option value=48>law  </option>
			<option value=49>linguistics  </option>
			<option value=50>literature  </option>
			<option value=51>logic  </option>
			<option value=52>management  </option>
			<option value=53>mathematics  </option>
			<option value=54>medical ethics  </option>
			<option value=55>medicine  </option>
			<option value=56>military science  </option>
			<option value=57>ministry  </option>
			<option value=58>missions  </option>
			<option value=59>music  </option>
			<option value=60>neuroscience</option>
			<option value=61>nuclear physics  </option>
			<option value=62>nutrition  </option>
			<option value=63>optics  </option>
			<option value=64>paleoanthropology  </option>
			<option value=65>paleontology  </option>
			<option value=66>penology  </option>
			<option value=67>philosophy  </option>
			<option value=68>philosophy of science  </option>
			<option value=69>physical education  </option>
			<option value=70>physics  </option>
			<option value=71>physiology  </option>
			<option value=72>politics  preaching  </option>
			<option value=73>psychiatry  </option>
			<option value=74>psychology  </option>
			<option value=75>psychotherapy  </option>
			<option value=76>publishing  </option>
			<option value=77>radiation biology  </option>
			<option value=78>recreation  </option>
			<option value=79>religion  </option>
			<option value=80>research  </option>
			<option value=81>resource studies  </option>
			<option value=82>scholarship  </option>
			<option value=83>science  </option>
			<option value=84>social work  </option>
			<option value=85>sociology  </option>
			<option value=86>space science  </option>
			<option value=87>statistical science  </option>
			<option value=88>statistics  </option>
			<option value=89>systems analysis  </option>
			<option value=90>taxonomy  </option>
			<option value=91>teaching  </option>
			<option value=92>technology  </option>
			<option value=93>theology  </option>
			<option value=94>translation  </option>
			<option value=95>zoology  </option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right"><STRONG>Revision or Part:</STRONG></td>
		<td><select name="VersionNo">
			<option value=0 selected>Select one:</option>
			<option value=1>Initial submission</option>
			<option value=2>Major revision or update</option>
			<option value=3>Minor revisions or additions&nbsp;</option>
			<option value=4>Figure or table </option>
			<option value=5>Final for Publication </option>
			<option value=6>Other</option>
			</select>
		</td>
	</tr>
  <TR>
		<td align="right"><STRONG>Notes:</STRONG><br> 
		<FONT size=1>(e.g. what MS does this file relate to? - optional - 255 chars. max.)</FONT></td>
		<td><TEXTAREA name="AuthorNote" rows=4 cols=60></TEXTAREA>
		</td></TR></TR>
  <TR>
		<td align="right"><STRONG>Abstract:</STRONG> <BR>
		<FONT size=1>(optional)</FONT></td>
		<td><TEXTAREA name="Abstract" rows=8 cols=60></TEXTAREA>
		</td>
	</TR>
	<tr><td align="right"><br><br>
		<input type="hidden" name="email" value="<%=email%>">
	   <input TYPE="submit" VALUE="Upload File"></td>
	   <td valign="bottom"><br>&nbsp;&nbsp;<STRONG>Click once; then wait for upload 
      to finish.</STRONG>       
	   <br>&nbsp;&nbsp;Then you will be directed to the next page. </td>
	</tr>
	</table>
	</form>
	<p><A href="menu.asp?id=<%=loginid%>">Return to menu</a></p>
	<br>
	<p><A href="default.htm">Return to login page</A></p>
</body>
</html>
