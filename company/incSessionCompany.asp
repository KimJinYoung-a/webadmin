<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
If session("ssBctId") = "" then
    %><html>
    <script>
    alert("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End
End if
%>