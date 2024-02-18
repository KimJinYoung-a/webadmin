<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode
dim groupCd, commCd, commNm, isusing
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)
groupCd		= RequestCheckvar(Request("groupCd"),10)
commCd		= RequestCheckvar(Request("commCd"),10)
commNm		= html2db(RequestCheckvar(Request("commNm"),32))
isusing		= RequestCheckvar(Request("isusing"),2)
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(Request("searchDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수


'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 신규등록
		'중복검사
		SQL = "Select count(commCd) as cnt From db_academy.dbo.tbl_CommCd where commCd='" & commCd & "'"
		rsACADEMYget.Open sql, dbACADEMYget, 1
			if rsACADEMYget("cnt")>0 then
				response.write	"<script language='javascript'>" &_
								"	alert('중복된 코드를 입력하였습니다.');" &_
								"	history.back();" &_
								"</script>"
				dbget.close()	:	response.End
			end if
		rsACADEMYget.close

		'저장
		SQL =	"Insert into db_academy.dbo.tbl_CommCd (groupCd, commCd, commNm) " &_
				"	Values " &_
				"	( '" & groupCd & "'" &_
				"	, '" & commCd & "'" &_
				"	, '" & commNm & "') "

		dbACADEMYget.Execute(SQL)

		msg = "신규 등록하였습니다."

		'돌아갈 페이지
		retURL = "CommCd_list.asp?menupos=" & menupos & param

	Case "modify"
		'@@ 수정처리

		SQL =	"Update db_academy.dbo.tbl_CommCd Set " &_
				"	commNm = '" & commNm & "'" &_
				"	,isUsing = '" & isusing & "'" &_
				" Where CommCd = '" & CommCd & "'"
		dbACADEMYget.Execute(SQL)

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "CommCd_list.asp?menupos=" & menupos & param

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