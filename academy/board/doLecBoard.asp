<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode, brdId, adminid
dim ansTitle, ansCont, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= RequestCheckvar(Request("menupos"),10)
brdId		= RequestCheckvar(Request("brdId"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
ansTitle	= html2db(RequestCheckvar(Request("ansTitle"),128))
ansCont	= html2db(Request("ansCont"))
page		= RequestCheckvar(Request("page"),10)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)
adminid		= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수


'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "answer"
		'@@ 답변처리
		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	  ansTitle= '" & ansTitle & "'" &_
				"	, ansCont = '" & ansCont & "'" &_
				"	, ansUserId = '" & Session("ssBctId") & "'" &_
				"	, ansDate = getdate() " &_
				"	, isanswer = 'Y' " &_
				" Where brdId = " & brdId

		dbACADEMYget.Execute(SQL)

		msg = "답변처리하였습니다."

		'돌아갈 페이지
		retURL = "lec_board_view.asp?menupos=" & menupos & "&brdId=" & brdId & param

	Case "change"
		'@@ 구분 변경

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	commCd = '" & commCd & "'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "구분을 변경하였습니다."

		'돌아갈 페이지
		retURL = "lec_board_list.asp?menupos=" & menupos & param

	Case "delete"
		'@@ 내용 삭제

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	isusing = 'N'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "삭제하였습니다."

		'돌아갈 페이지
		retURL = "lec_board_list.asp?menupos=" & menupos & param

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->