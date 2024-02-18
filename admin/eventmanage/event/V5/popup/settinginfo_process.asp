<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : listbanner_process.asp
' Discription : I��(������) �̺�Ʈ ����Ʈ ��� ��� ���μ���
' History : 2019.01.25 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, strSql, eDispCate, eBrand, eTag, nocate, eISort
dim refer
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
eDispCate = requestCheckVar(Request.Form("disp"),12)
nocate = requestCheckVar(Request.Form("nocate"),1)
eISort = requestCheckVar(Request.Form("itemsort"),4)
eBrand = Request.Form("ebrand")
eTag = html2db(requestCheckVar(Replace(Request.Form("eTag")," ",""),300))
If Right(eTag,1) = "," Then
    eTag = Left(eTag,(Len(eTag)-1))
End If

if eTag <> "" then
	if checkNotValidHTML(eTag) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.');history.back();"
	response.write "</script>"
	response.End
end if

select case eMode
case "SU"
	dbget.beginTrans
        '===========================================================
        '--1.disply ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET evt_dispCate='" & eDispCate & "'" & vbCrlf
        strSql = strSql + ", brand='" & eBrand & "'" & vbCrlf
        strSql = strSql + ", evt_tag='" & eTag & "'" & vbCrlf
        strSql = strSql + ", evt_itemsort='" & eISort & "'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
            response.End 
        end if

        '--2.theme ����
        strSql = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
        strSql = strSql + " SET nocate='" & nocate & "'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if

    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->