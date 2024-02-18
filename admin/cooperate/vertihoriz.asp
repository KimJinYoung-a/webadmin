<%@ language=vbscript %>
<% option explicit %>
<%
	If Request.cookies("scmcooperatevertihoriz") = "h" Then
		Response.Cookies("scmcooperatevertihoriz") = "v"
	Else
		Response.Cookies("scmcooperatevertihoriz") = "h"
	End IF
	Response.Redirect "/admin/cooperate/popIndex.asp?mn=" & Request("mn")
%>