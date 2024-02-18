<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_EntryList.asp
' Description :  이벤트 응모자 공통 통계 선작업
' History : 2007.10.06 김정인
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

dim cEvtCont
dim ekind
set cEvtCont = new ClsEvent
cEvtCont.FECode = evtCode	'이벤트 코드
'이벤트 내용 가져오기
cEvtCont.fnGetEventCont

ekind =	cEvtCont.FEKind
set cEvtCont = nothing
dbget.BeginTrans
dim strSQL,msg
	'// 쇼핑 찬스 게시판
	strSQL =" INSERT into db_event.dbo.tbl_event_common_comment (evt_code ,userid ,com_txt ,com_regdate ,org_evtcom_idx) " &_
			" SELECT evt_code ,userid ,evtbbs_content ,evtbbs_regdate ,evtbbs_idx " &_
			" FROM [db_event].[dbo].[tbl_event_bbs] " &_
			" WHERE evt_code=" & evtCode & " " &_
			" and evtbbs_using='Y' and evtbbs_idx not in (SELECT org_evtcom_idx FROM db_event.dbo.tbl_event_common_comment WHERE evt_code='" & evtCode & "')"

	dbget.execute(strSQL)

	'//한줄 낙서
	strSQL =" INSERT into db_event.dbo.tbl_event_common_comment (evt_code,userid,com_txt,com_regdate,org_evtcom_idx) " &_
			" SELECT evt_code,userid,comment,regdate,idx " &_
			" FROM [db_contents].[dbo].tbl_one_comment c " &_
			" WHERE evt_code='" & evtCode & "' and isusing='Y' " &_
			" and idx not in (SELECT org_evtcom_idx FROM db_event.dbo.tbl_event_common_comment WHERE evt_code='" & evtCode & "') "
	dbget.execute(strSQL)
	'/ 문화 이벤트
	strSQL =" INSERT into db_event.dbo.tbl_event_common_comment (evt_code,userid,com_txt,com_regdate,org_evtcom_idx) " &_
			" SELECT evt_code,userid,evtcom_txt,evtcom_regdate,evtcom_idx " &_
			" FROM db_event.[dbo].tbl_event_comment " &_
			" WHERE evt_code='" & evtCode & "' and evtcom_using='Y' " &_
			" and evtcom_idx not in (SELECT org_evtcom_idx FROM db_event.dbo.tbl_event_common_comment WHERE evt_code='" & evtCode & "') " &_
			" ORDER BY evtcom_idx "
	dbget.execute(strSQL)

	strSQL =" execute db_event.dbo.ten_event_user_log_set @EVTCD='" & evtCode & "',@EVTKind ='" & ekind & "'"

	'response.write strSQL

	dbget.execute(strSQL)

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

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