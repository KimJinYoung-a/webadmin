<%
DIM CAddDetailSpliter : CAddDetailSpliter= CHR(3)&CHR(4)

dim C_ADMIN_AUTH
dim C_MngPart               '' �濵������ ����.
dim C_InspectorUser			''����

C_ADMIN_AUTH = (session("ssBctId") = "coolhas")
C_MngPart = (session("ssAdminPsn")="8")
C_InspectorUser = (session("ssBctId") = "aimcta1" )

dim iiisAdmin
iiisAdmin = (session("ssBctId") = "10x10")

if Not iiisAdmin then
  iiisAdmin = (session("ssBctId")<>"")
  iiisAdmin = iiisAdmin and ((session("ssBctDiv")<=9) or (session("ssBctDiv")=101) or (session("ssBctDiv")=111) or (session("ssBctDiv")=112) or (session("ssBctDiv")=201) or (session("ssBctDiv")=301))
end if

If (Not iiisAdmin) then
 %>
    <script>
    alert("60���� ����Ǿ� �α׾ƿ��Ǿ����ϴ�. \n�ٽ� �α��� �� ����ϽǼ� �ֽ��ϴ�.");
    top.location = "/index.asp";
    </script>
    <%
    response.End
End if
%> 