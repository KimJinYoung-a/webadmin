<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode, brdId, lecUserId
dim qstTitle, qstCont, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= requestCheckVar(Request("menupos"),10)
brdId		= requestCheckVar(Request("brdId"),10)
mode		= requestCheckVar(Request("mode"),16)
commCd		= requestCheckVar(Request("commCd"),4)
qstTitle	= html2db(Request("qstTitle"))
qstCont	= html2db(Request("qstCont"))
page		= requestCheckVar(Request("page"),10)
searchDiv	= requestCheckVar(Request("searchDiv"),10)
searchKey	= requestCheckVar(Request("searchKey"),10)
searchString = requestCheckVar(Request("searchString"),128)
lecUserId	= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

  	if qstTitle <> "" then
		if checkNotValidHTML(qstTitle) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end If
  	if qstCont <> "" then
		if checkNotValidHTML(qstCont) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end If
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if

'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 내용 저장
		SQL =	"Insert into db_academy.dbo.tbl_lec_board (qstTitle, qstCont, lecUserId, commCd) values " &_
				"	 ('" & qstTitle & "'" &_
				"	, '" & qstCont & "'" &_
				"	, '" & lecUserId & "'" &_
				"	, '" & commCd & "')"

		dbACADEMYget.Execute(SQL)

		msg = "저장하였습니다."

		'돌아갈 페이지
		retURL = "board_list.asp?menupos=" & menupos & param

	Case "modify"
		'@@ 내용 수정

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	qstTitle = '" & qstTitle & "'" &_
				"	qstCont = '" & qstCont & "'" &_
				"	commCd = '" & commCd & "'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "내용을 수정하였습니다."

		'돌아갈 페이지
		retURL = "board_view.asp?menupos=" & menupos & "&brdId=" & brdId & param

	Case "delete"
		'@@ 내용 삭제

		SQL =	"Update db_academy.dbo.tbl_lec_board Set " &_
				"	isusing = 'N'" &_
				" Where brdId = " & brdId
		dbACADEMYget.Execute(SQL)

		msg = "삭제하였습니다."

		'돌아갈 페이지
		retURL = "board_list.asp?menupos=" & menupos & param

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
<!-- #include virtual="/lib/db/dbclose.asp" -->