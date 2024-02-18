<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
If session("ssBctId") = "" then
    %><html>
    <script>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if
%>