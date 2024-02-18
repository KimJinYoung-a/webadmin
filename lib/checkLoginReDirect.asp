<%
dim redirectpath

if (request.ServerVariables("QUERY_STRING") = "") then
        redirectpath = request.ServerVariables("URL")
else
        redirectpath = request.ServerVariables("URL") + "?" + request.ServerVariables("QUERY_STRING")
end if
redirectpath = server.URLEncode(redirectpath)

If (session("ssBctId") = "") then
	response.redirect("/index.asp?backpath=" + redirectpath)
end if
%>