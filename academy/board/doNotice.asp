<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode, ntcId, userid
dim title, contents, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL


'// 내용 접수 및 처리
menupos		= RequestCheckvar(Request("menupos"),10)
ntcId		= RequestCheckvar(Request("ntcId"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
title		= html2db(RequestCheckvar(Request("title"),128))
contents	= html2db(Request("contents"))
page		= RequestCheckvar(Request("page"),10)
searchDiv	= RequestCheckvar(request("searchDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)
userid		= session("ssBctId")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수


'==============================================================================
'## 내용 저장(수정) 처리

if (checkNotValidHTML(title) = true) Then
	response.write "<script>alert('공지사항 제목에는 HTML을 사용하실 수 없습니다.');history.back();</script>"
	dbget.Close
	response.End
End If

'' imgsrc / ahref 도 체크하는 이유?	checkNotValidHTML = > checkNotValidHTMLcritical
''if (checkNotValidHTMLcritical(sBrd_content) = true) Then			'// img 태그 허용으로 수정 > 검사항목 단일화
if (checkNotValidHTML(contents) = true) Then
	response.write "<script>alert('공지사항 내용에는 Script 또는 Action을 사용하실 수 없습니다.');history.back();</script>"
	dbget.Close
	response.End
End If


'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 내용 저장
		SQL =	"Insert into db_academy.dbo.tbl_Notice " &_
				"	(title, contents, commCd, userid) values " &_
				"	('" & title & "'" &_
				"	,'" & contents & "'" &_
				"	,'" & commCd & "'" &_
				"	,'" & userid & "')"
		dbACADEMYget.Execute(SQL)

		'결과 메시지
		msg = "저장하였습니다."
		
		'돌아갈 페이지
		retURL = "Notice_list.asp?menupos=" & menupos & param


	Case "modify"
		'@@ 내용 수정
		SQL =	"Update db_academy.dbo.tbl_Notice Set " &_
				"	  title= '" & title & "'" &_
				"	, contents = '" & contents & "'" &_
				"	, commCd = '" & commCd & "'" &_
				" Where ntcId = " & ntcId

		dbACADEMYget.Execute(SQL)

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "Notice_view.asp?menupos=" & menupos & "&ntcId=" & ntcId & param

	Case "delete"
		'@@ 내용 삭제

		'# 내용 삭제
		SQL =	"Update db_academy.dbo.tbl_Notice Set " &_
				"	  title= '" & title & "'" &_
				"	, isusing = 'N'" &_
				" Where ntcId = " & ntcId

		msg = "삭제하였습니다."

		'돌아갈 페이지
		retURL = "Notice_list.asp?menupos=" & menupos & param

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"</script>"
'"	self.location='" & retURL & "';" &_
Else
    dbACADEMYget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

IF application("Svr_Info") = "Dev" THEN
	Response.Redirect "http://test.thefingers.co.kr/chtml/make_index_notice.asp?retURL=http://testwebadmin.10x10.co.kr/academy/board/notice_list.asp?menupos=784"
ELSE
	Response.Redirect "http://www.thefingers.co.kr/chtml/make_index_notice.asp?retURL=http://webadmin.10x10.co.kr/academy/board/notice_list.asp?menupos=784"
END IF

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->