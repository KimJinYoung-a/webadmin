<%
'' �귣�� ������ �߿������� ���� ��� �α�
'' ������������ ���Ե� ������

function fnChkPartnerPageLog(bltype, refip)
    dim sqlStr
    dim scrname : scrname = Request.ServerVariables("SCRIPT_NAME")
    dim strMethod : strMethod = Request.ServerVariables("REQUEST_METHOD")
    dim qryStr
    If strMethod = "POST" Then
        qryStr = (Request.Form)  ''Server.HTMLEncode
    else
        qryStr = Request.QueryString
    end if

    sqlStr = "exec db_log.dbo.sp_TEN_ChkAllowIpLog '"&bltype&"','"&session("ssBctID")&"','"&refip&"','"&scrname&"','"&replace(qryStr,"'","")&"','"&LEFT(strMethod,1)&"'"
    dbget.Execute sqlStr
end function

dim TMP_check_PartnerIP
TMP_check_PartnerIP = request.ServerVariables("REMOTE_ADDR")

Call fnChkPartnerPageLog("P",TMP_check_PartnerIP)
%>