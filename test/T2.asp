<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

response.write "refer="&refer
%>