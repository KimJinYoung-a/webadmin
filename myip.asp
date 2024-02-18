<%
response.write request.ServerVariables("REMOTE_ADDR") 
 
 
response.write "<BR>"&request.ServerVariables("HTTP_USER_AGENT") 
%>       
      