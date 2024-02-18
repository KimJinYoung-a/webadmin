<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/eventmanage/event_process.asp
' Description :  이벤트 개요 데이터처리 - 등록, 수정, 삭제
' History : 2007.02.12 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim sqlStr
Dim idx : idx = requestCheckVar(Request.Form("idx"),9)

If idx <> "" Then
dbget.beginTrans
		sqlStr = " delete FROM db_event.[dbo].[tbl_event_slide_addimage] WHERE idx=" & idx
		dbget.execute sqlStr
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
		response.End 
	END IF
dbget.CommitTrans
Response.write "<script>alert('삭제 되었습니다.');parent.location.reload();</script>"
Else
Response.write "<script>alert('정보가 부정확하여 삭제가 불가능합니다.');</script>"
End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->