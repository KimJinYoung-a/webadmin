<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode, prfid, userid
dim title, prfCont, groupCd, commCd, isusing
dim SQL
dim page, searchDiv, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= RequestCheckvar(Request("menupos"),10)
prfid		= RequestCheckvar(Request("prfid"),10)
mode		= RequestCheckvar(Request("mode"),16)
groupCd		= RequestCheckvar(Request("groupCd"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
prfCont	= html2db(Request("prfCont"))
isusing		= RequestCheckvar(Request("isusing"),2)
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(Request("searchDiv"),16)
searchString = RequestCheckvar(Request("searchString"),128)

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchString=" & searchString	'페이지 변수


'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 내용 저장
		SQL =	"Insert into db_academy.dbo.tbl_preface " &_
				"	(prfCont, groupCd, commCd) values " &_
				"	('" & prfCont & "'" &_
				"	,'" & groupCd & "'" &_
				"	,'" & commCd & "')"
		dbACADEMYget.Execute(SQL)

		'결과 메시지
		msg = "저장하였습니다."
		
		'돌아갈 페이지
		retURL = "Preface_list.asp?menupos=" & menupos & param


	Case "modify"
		'@@ 내용 수정
		SQL =	"Update db_academy.dbo.tbl_preface Set " &_
				"	  prfCont = '" & prfCont & "'" &_
				"	, groupCd = '" & groupCd & "'" &_
				"	, commCd = '" & commCd & "'" &_
				"	, isusing = '" & isusing & "'" &_
				" Where prfid = " & prfid

		dbACADEMYget.Execute(SQL)

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "Preface_modi.asp?menupos=" & menupos & "&prfid=" & prfid & param

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