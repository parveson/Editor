<%@ Language=VBScript %>
<%option explicit
Response.expires=30
'Response.addHeader "pragma","no-cache"
'Response.addHeader "cache-control","private"
'Response.CacheControl = "no-cache"
Response.Buffer = true %>
<!-- #include file="../test2.asp" -->
<!-- The following variables are defined in test2.asp:        -->
<!-- loginid,rsLogintest,qtest,datapath,DSN1,DSN2,LoginDSN    -->
<HTML>
<HEAD>
<title>Author's Menu</title>
<LINK rel="stylesheet" type="text/css" href="../StyleSheet1.css">
</HEAD>
<BODY>
<h3>Author's Menu</h3>

<p>Welcome to the ASA Author's page.  You have the following options:</p>

<p><A href="addfile.asp?id=<%=loginid%>">Upload a Manuscript file</A></p><!--<p><A href="addms_test.asp?id=<%=loginid%>">Test Upload</a></p>

<p><A href="insert_test2.asp?id=<%=loginid%>">Insert test 2</a></p>-->

<p><A href="listmyfiles.asp?id=<%=loginid%>">List my Manuscript files</A></p>

<p><A href="changepw.asp?id=<%=loginid%>">Change my password</A></p>
<br>
<p><A href="../default.htm">Return to ASA Editor's home page</A></p>

<p><a href="http://www.asa3.org/">Return to ASA main web site</a></p>
<P>
<P><A href="../contact.asp">Contact us</A></P>
<P>

<hr>

<P></P>
<P></P>
<table align="center" summary="BSCI Addresses" border="0" width="300" cellpadding="0" cellspacing="0">
  <tr>
    <td align="middle"><b><font color="darkgreen">American Scientific Affiliation</font></b>
      </td>
   </tr>
   <tr>
      <td class="small" align="middle">Editorial Office</td>
    </tr>
    <tr>
      <td class="small" align="middle">Roman J. Miller, Editor</td>
   </tr>
   <tr>
      <td class="small" align="middle">4956 Singers Glen Road</td>
  <tr>
	<td class="small" align="middle">Harrisonburg, VA 22802</td>
  </tr>
    <tr><td class="small" align="middle">(540) 432-4412</td>
  </tr>
  <tr><td class="small" align="middle"><small><strong><em><font color="darkblue">www.asaeditor.org</font></em></strong></small></td>
  </tr>
</table>

</BODY>
</HTML>
