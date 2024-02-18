<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트
' History : 2010.09.17 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim msg, lp, menupos ,evtTitle, evtCont, evtSdate, evtEdate, isComment
dim mode, evtId ,SQL ,page, searchKey, searchString, param, retURL
	menupos		= RequestCheckvar(Request("menupos"),10)
	evtId		= RequestCheckvar(Request("evtId"),10)
	mode		= RequestCheckvar(Request("mode"),10)
	evtTitle		= html2db(RequestCheckvar(Request("evtTitle"),64))
	evtCont	= html2db(Request("evtCont"))
	page		= RequestCheckvar(Request("page"),10)
	searchKey	= RequestCheckvar(Request("searchKey"),16)
	searchString = RequestCheckvar(Request("searchString"),128)
	evtSdate = Request("syy") & "-" & Request("smm") & "-" & Request("sdd")
	evtEdate = Request("eyy") & "-" & Request("emm") & "-" & Request("edd")
	isComment = Request("isComment")

  	if evtCont <> "" then
		if checkNotValidHTML(evtCont) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
  	if isComment <> "" then
		if checkNotValidHTML(isComment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 내용 저장
		SQL =	"Insert into db_academy.dbo.tbl_eventInfo " &_
				"	(evtTitle, evtCont, evtSdate, evtEdate, isComment) values " &_
				"	('" & evtTitle & "'" &_
				"	,'" & evtCont & "'" &_
				"	,'" & evtSdate & "'" &_
				"	,'" & evtEdate & "'" &_
				"	,'" & isComment & "')"
		dbACADEMYget.Execute(SQL)

		'결과 메시지
		msg = "저장하였습니다."
		
		'돌아갈 페이지
		retURL = "Event_list.asp?menupos=" & menupos & param

	Case "modify"
		'@@ 내용 수정
		SQL =	"Update db_academy.dbo.tbl_eventInfo Set " &_
				"	  evtTitle= '" & evtTitle & "'" &_
				"	, evtCont = '" & evtCont & "'" &_
				"	, evtSdate = '" & evtSdate & "'" &_
				"	, evtEdate = '" & evtEdate & "'" &_
				"	, isComment = '" & isComment & "'" &_
				" Where evtId = " & evtId

		dbACADEMYget.Execute(SQL)

		msg = "수정하였습니다."

		'돌아갈 페이지
		retURL = "Event_view.asp?menupos=" & menupos & "&evtId=" & evtId & param

	Case "delete"
		'@@ 내용 삭제

		'# 내용 삭제
		SQL =	"Update db_academy.dbo.tbl_eventInfo Set " &_
				"	  evtTitle= '" & evtTitle & "'" &_
				"	, isusing = 'N'" &_
				" Where evtId = " & evtId

		msg = "삭제하였습니다."

		'돌아갈 페이지
		retURL = "Event_list.asp?menupos=" & menupos & param

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

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->