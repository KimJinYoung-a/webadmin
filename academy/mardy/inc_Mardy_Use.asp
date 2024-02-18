<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
	'// 변수 선언 //
	dim Idx, mode, msg, SQL

	'// 파라메터 접수 //
	Idx = RequestCheckvar(request("Idx"),10)
	mode = RequestCheckvar(request("mode"),16)

	Select Case mode
		Case "ScrapUse"
			SQL =	"Update db_academy.dbo.tbl_mardyScrap Set " &_
					"	isusing = 'Y' " &_
					"Where scrapId=" & Idx
			msg = "사용"

		Case "ScrapDel"
			SQL =	"Update db_academy.dbo.tbl_mardyScrap Set " &_
					"	isusing = 'N' " &_
					"Where scrapId=" & Idx
			msg = "숨김"

		Case "StoryUse"
			SQL =	"Update db_academy.dbo.tbl_mardyStory Set " &_
					"	isusing = 'Y' " &_
					"Where storyId=" & Idx
			msg = "사용"

		Case "StoryDel"
			SQL =	"Update db_academy.dbo.tbl_mardyStory Set " &_
					"	isusing = 'N' " &_
					"Where storyId=" & Idx
			msg = "숨김"

		Case "TipUse"
			SQL =	"Update db_academy.dbo.tbl_mardyTip Set " &_
					"	isusing = 'Y' " &_
					"Where tipId=" & Idx
			msg = "사용"

		Case "TipDel"
			SQL =	"Update db_academy.dbo.tbl_mardyTip Set " &_
					"	isusing = 'N' " &_
					"Where tipId=" & Idx
			msg = "숨김"

		Case else
			dbget.close()	:	response.End
	End Select

	'트랜젝션 시작
	dbACADEMYget.beginTrans

	'// 사용여부 처리
	dbACADEMYget.Execute(SQL)


	'// 오류검사 및 반영
	If Err.Number = 0 Then   
		dbACADEMYget.CommitTrans				'커밋(정상)
	
		response.write	"<script language='javascript'>" &_
						"	alert('사용유무를 [" & msg & "]으로 변경하였습니다.');" &_
						"	parent.history.go(0);" &_
						"</script>"
	Else
	    dbACADEMYget.RollBackTrans				'롤백(에러발생시)
	
		response.write	"<script language='javascript'>" &_
						"	alert('처리중 에러가 발생했습니다.');" &_
						"</script>"
	
	End If
%>