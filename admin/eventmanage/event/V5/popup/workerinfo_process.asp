<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : workerinfo_process.asp
' Discription : I��(������) �̺�Ʈ ����� ���� ��� ���μ���
' History : 2019.01.22 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, eDgId, edgstat
dim eMdId, ePsId, eDpId, eCCId, blnReqPublish
dim sWorkTag, strSql, fromlist
dim refer
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
eDgId = requestCheckVar(Request.Form("sDgId"),32)
edgstat = requestCheckVar(Request.Form("designerstatus"),2)

    
eMdId = requestCheckVar(Request.Form("sMdId"),32)	
ePsId = requestCheckVar(Request.Form("sPsId"),32)
eDpId = requestCheckVar(Request.Form("sDpId"),32)
eCCId = requestCheckVar(Request.Form("sCCId"),32)

sWorkTag = requestCheckVar(Request.Form("sWorkTag"),32)
blnReqPublish = requestCheckVar(Request.Form("chkReqP"),1)
fromlist = requestCheckVar(Request.Form("fromlist"),1)

IF blnReqPublish = "" THEN blnReqPublish = 0 

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.');history.back();"
	response.write "</script>"
	response.End
end if

'--------------------------------------------------------
' ������ ó��
' I : �̺�Ʈ ������, U: �������, disply���/����
'--------------------------------------------------------
select case eMode
case "WU"
	dbget.beginTrans
        '===========================================================
        '--2.disply ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET designerid='" & eDgId & "'" & vbCrlf
        strSql = strSql + ", partMDid='" & eMdId & "'" & vbCrlf
        strSql = strSql + ", publisherid='" & ePsId & "'" & vbCrlf
        strSql = strSql + ", developerid='" & eDpId & "'" & vbCrlf
        strSql = strSql + ", workTag='" & sWorkTag & "'" & vbCrlf
        strSql = strSql + ", dsn_state1='" & edgstat & "'" & vbCrlf
        strSql = strSql + ", isReqPublish=" & blnReqPublish & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
    if fromlist="Y" then
        response.write "<script type='text/javascript'>"
        response.write "    window.document.domain = ""10x10.co.kr"";"
        response.write "	opener.document.location.reload();"
        response.write "    self.close();"
        response.write "</script>"
        dbget.close()	:	response.End
    else
        response.write "<script type='text/javascript'>"
        response.write "    window.document.domain = ""10x10.co.kr"";"
        response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
        'response.write "    location.replace('" + refer + "');"
        response.write "    self.close();"
        response.write "</script>"
        dbget.close()	:	response.End
    end if
case else
	Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->