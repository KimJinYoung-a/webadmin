<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : listbanner_process.asp
' Discription : I형(통합형) 이벤트 리스트 배너 등록 프로세스
' History : 2019.01.25 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, strSql, etcitemban, banMoList
dim eMdId, ePsId, eDpId, eCCId, blnReqPublish
dim refer
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
etcitemban = requestCheckVar(Request.Form("etcitemban"),256)
banMoList = requestCheckVar(Request.Form("banMoList"),256)

if etcitemban <> "" then
	if checkNotValidHTML(etcitemban) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if banMoList <> "" then
	if checkNotValidHTML(banMoList) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
	response.write "</script>"
	response.End
end if

select case eMode
case "BU"
	dbget.beginTrans
        '===========================================================
        '--2.disply 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET etc_itemimg='" & etcitemban & "'" & vbCrlf
        strSql = strSql + ", evt_mo_listbanner='" & banMoList & "'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
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
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->