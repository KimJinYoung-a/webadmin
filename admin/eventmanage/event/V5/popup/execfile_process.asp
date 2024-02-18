<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : execfile_process.asp
' Discription : I형(통합형) 이벤트 멀티컨텐츠 개발파일 정보 등록 프로세스
' History : 2019.10.14 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->

<%

dim eCode, eMode, strSql, device, menuidx
dim evt_html_mo, evt_mainimg_mo, Idx, sEFP, isusing

eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
Idx = requestCheckVar(Request.Form("Idx"),10)
menuidx = requestCheckVar(Request.Form("menuidx"),10)
device = requestCheckVar(Request.Form("device"),1)

sEFP = html2db(Request.Form("sEFP"))
isusing = requestCheckVar(Request.Form("isusing"),1)

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
	response.write "</script>"
	response.End
end if

select case eMode
Case "TI"
	dbget.beginTrans
		strSql = "Insert Into db_event.dbo.tbl_event_multi_contents " &_
					" (menuidx, device , imgurl, isusing) values " &_
					" ('" & menuidx  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & sEFP &"'" &_
					" ,'" & isusing &"')"
		dbget.Execute(strSql)

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case "TU"

	strSql = "Update db_event.dbo.tbl_event_multi_contents Set " & vbCrLf
	strSql = strSql & " imgurl='" & sEFP & "'" & vbCrLf
	strSql = strSql & ", isusing='" & isusing & "'" & vbCrLf
	strSql = strSql & " Where idx='" & Idx & "';"
	dbget.execute strSql

	if Err.Number <> 0 then
		dbget.RollBackTrans 
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		response.End 
	end if

	response.write "<script type='text/javascript'>"
	response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->