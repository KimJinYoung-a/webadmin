<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : multicontentssort_process.asp
' Discription : I형(통합형) 이벤트 멀티 컨텐츠 메뉴 순서 지정 프로세스
' History : 2019.02.13 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim eCode, eMode, strSql
dim sIDX, sSortNo, ix

eCode = requestCheckVar(Request.Form("evt_code"),10)

    if eCode="" then
        response.write "<script type='text/javascript'>"
        response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
        response.write "</script>"
        response.End
    end if

	dbget.beginTrans
        '===========================================================
        '--순서 수정
        for ix=1 to request.form("idx").count
            sIDX = request.form("idx")(ix)
            sSortNo = request.form("sort")(ix)
            strSql = "UPDATE [db_event].[dbo].[tbl_event_multi_contents_master]" & vbCrlf
            strSql = strSql + " SET viewsort='" & sSortNo & "'" & vbCrlf
            strSql = strSql + " where evt_code=" & eCode
            strSql = strSql + " and idx=" & sIDX & ";"
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
	response.write "	parent.document.location.reload();"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->