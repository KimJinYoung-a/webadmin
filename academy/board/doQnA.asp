<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.16 한용민 수정
'	Description : 1:1상담
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/QnA_cls.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode, qnaId, adminid
dim ansTitle, ansContents, commCd, mailOk, qstUserMail
dim SQL, mailcontent
dim page, searchDiv, searchKey, searchString, param, retURL
dim qstUserName,regdate,qstTitle , qstContents


'// 내용 접수 및 처리
qstUserName		= RequestCheckvar(Request("qstUserName"),32)
qstContents		= Request("qstContents")
regdate		= RequestCheckvar(Request("regdate"),32)
qstTitle		= RequestCheckvar(Request("qstTitle"),128)
menupos		= RequestCheckvar(Request("menupos"),10)
qnaId		= RequestCheckvar(Request("qnaId"),10)
mode		= RequestCheckvar(Request("mode"),16)
commCd		= RequestCheckvar(Request("commCd"),16)
mailOk		= RequestCheckvar(Request("mailOk"),16)
qstUserMail	= RequestCheckvar(Request("qstUserMail"),64)
ansTitle	= html2db(RequestCheckvar(Request("ansTitle"),128))
ansContents	= html2db(Request("ansContents"))
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
		SQL =	"Update db_academy.dbo.tbl_QnA Set " &_
				"	  ansTitle= '" & ansTitle & "'" &_
				"	, ansContents = '" & ansContents & "'" &_
				"	, ansDate = getdate() " &_
				"	, isanswer = 'Y' " &_
				" Where qnaId = " & qnaId

		dbACADEMYget.Execute(SQL)


		'답변 메일 발송
		if (mailOk = "수신") and (Cstr(qstUserMail)<>"") then
            '메일 템플릿 접수
            mailcontent = ReadLocalFile("tpl_fingers_qna.html", "/academy/lib/mail_templete")
            
            '내용 치환
            mailcontent = Replace(mailcontent,"#ansContents#",nl2br(ansContents))
            '내용 치환
            mailcontent = Replace(mailcontent,"#qstContents#",nl2br(qstContents))
            '내용 치환
            mailcontent = Replace(mailcontent,"#regdate#",FormatDate(regdate,"0000-00-00"))
            '내용 치환
            mailcontent = Replace(mailcontent,"#qstUserName#",qstUserName)
            '내용 치환
            mailcontent = Replace(mailcontent,"#qstTitle#",qstTitle)           
            '발송
            call Send_mail("customer@thefingers.co.kr", qstUserMail, "함께 배우는 즐거움 핑거스", mailcontent)
		end if
				
		msg = "답변처리하였습니다."

		'돌아갈 페이지
		retURL = "QnA_view.asp?menupos=" & menupos & "&qnaId=" & qnaId & param

	Case "change"
		'@@ 구분 변경

		SQL =	"Update db_academy.dbo.tbl_QnA Set " &_
				"	commCd = '" & commCd & "'" &_
				" Where qnaId = " & qnaId
		dbACADEMYget.Execute(SQL)

		msg = "구분을 변경하였습니다."

		'돌아갈 페이지
		retURL = "QnA_list.asp?menupos=" & menupos & param

	Case "delete"
		'@@ 내용 삭제

		SQL =	"Update db_academy.dbo.tbl_QnA Set " &_
				"	isusing = 'N'" &_
				" Where qnaId = " & qnaId
		dbACADEMYget.Execute(SQL)

		msg = "삭제하였습니다."

		'돌아갈 페이지
		retURL = "QnA_list.asp?menupos=" & menupos & param

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
