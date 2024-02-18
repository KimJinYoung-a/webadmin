<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : relationevent_process.asp
' Discription : I형(통합형) 이벤트 관련 이벤트 등록 프로세스
' History : 2019.02.27 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, strSql, masterCode, idx, msg
dim arrViewIdx, arrIDX, ix
dim refer
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
masterCode = requestCheckVar(Request.Form("ecode"),10)
idx = requestCheckVar(Request.Form("idx"),10)

select case eMode
case "RI"
    if eCode<>"" then
        dbget.beginTrans
            '===========================================================
            strSql = "insert into [db_event].[dbo].[tbl_relation_event](mastercode, ecode)" & vbCrlf
            strSql = strSql + " values(" & masterCode & "," & eCode & ")" & vbCrlf
            dbget.execute strSql

            if Err.Number <> 0 then
                dbget.RollBackTrans 
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
                response.End 
            end if
            '===========================================================
        dbget.CommitTrans
        response.write "<script type='text/javascript'>"
        response.write "    window.document.domain = ""10x10.co.kr"";"
        response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(masterCode) + "&togglediv=4&viewset='+opener.document.frmEvt.viewset.value);"
        'response.write "    location.replace('" + refer + "');"
        response.write "    self.close();"
        response.write "</script>"
        dbget.close()	:	response.End
    else
        dbget.beginTrans
            '===========================================================
            for ix=1 to request.form("idx").count
                arrIDX = request.form("idx")(ix)
                arrViewIdx = request.form("viewidx")(ix)
                strSql = "UPDATE [db_event].[dbo].[tbl_relation_event]" & vbCrlf
                strSql = strSql + " SET viewidx=" & arrViewIdx & vbCrlf
                strSql = strSql + " where idx=" & arrIDX
                dbget.execute strSql
            next

            if Err.Number <> 0 then
                dbget.RollBackTrans 
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
                response.End 
            end if
            '===========================================================
        dbget.CommitTrans
        response.write "<script type='text/javascript'>"
        response.write "    window.document.domain = ""10x10.co.kr"";"
        response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(masterCode) + "&togglediv=4&viewset='+opener.document.frmEvt.viewset.value);"
        'response.write "    location.replace('" + refer + "');"
        response.write "    self.close();"
        response.write "</script>"
        dbget.close()	:	response.End
    end if
case "RD"
	dbget.beginTrans
        '===========================================================
        '--1.disply 수정
        strSql = "delete from [db_event].[dbo].[tbl_relation_event]" & vbCrlf
        strSql = strSql + " where idx=" & idx
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
            response.End 
        end if
        '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(masterCode) + "&togglediv=4&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->