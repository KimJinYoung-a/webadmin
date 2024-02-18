<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_EntryList.asp
' Description :  이벤트 응모자 처리 페이지
' History : 2007.09.21 김정인
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<%

dim evtCode,arridx,selStr
'
evtCode =request("eC")
arridx =chkarray(request("arridx"))
selStr =request("selStr")

dim strSQL,msg

IF selStr = "N" THEN
'// 삭제
	strSQL =" DELETE FROM db_event.dbo.tbl_event_comment " &_
			" where evt_code='" & evtCode & "' " &_
			" and userid in( " &_
			" 	SELECT userid  " &_
			" 	FROM db_event.dbo.tbl_event_comment " &_
			" 	WHERE evtcom_idx in (" & arridx &")) "

	msg="삭제 되었습니다"

ELSEIF selStr= "S" THEN
'// 선택
	strSQL =" INSERT INTO [db_event].[dbo].[tbl_event_winner_log](evt_code,com_idx,userid) " &_
			" SELECT evt_code,com_idx,userid  " &_
			" FROM [db_event].[dbo].tbl_event_common_comment " &_
			" WHERE com_idx in (" & arridx &") " &_
			" and com_idx not in (select com_idx from db_event.dbo.tbl_event_winner_log where evt_code='" & evtCode & "')"
	msg="선택되었습니다."


END IF
	dbget.BeginTrans

	'response.write strsQL
	dbget.execute(strSQL)

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		response.write	"<script language='javascript'>"
		response.write	"	alert('" & msg & "'); parent.location.reload();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>"
		response.write	"	alert('처리중 에러가 발생했습니다.');"
		response.write	"</script>"
		dbget.close()	:	response.End

	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->