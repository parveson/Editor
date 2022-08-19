<%@ Language=VBScript %>
<%option explicit%>
<HTML>
<HEAD>
<title>Function demo</title>
</HEAD>
<BODY>

<h3>Function demo</h3>

<% dim a,b,c,csum

a=1
b=3

csum = summit(a,b)

Response.Write("<p>sum = " & csum & "</p>")


function summit(a,b)
	summit = a + b
end function

%>
</BODY>
</HTML>
