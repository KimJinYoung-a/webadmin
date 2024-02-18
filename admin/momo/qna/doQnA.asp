<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 qna 저장
' Hieditor : 2009.11.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// 변수 선언
dim msg, lp, menupos
dim mode, qnaId, adminid
dim ansTitle, ansContents, commCd, mailOk, qstUserMail
dim SQL, mailcontent
dim page, searchDiv, searchKey, searchString, retURL
dim qstUserName,regdate,qstTitle , qstContents


'// 내용 접수 및 처리
qstUserName		= Request("qstUserName")
qstContents		= Request("qstContents")
regdate		= Request("regdate")
qstTitle		= Request("qstTitle")
menupos		= Request("menupos")
qnaId		= Request("qnaId")
mode		= Request("mode")
commCd		= Request("commCd")
mailOk		= Request("mailOk")
qstUserMail	= Request("qstUserMail")
ansTitle	= html2db(Request("ansTitle"))
ansContents	= html2db(Request("ansContents"))
page		= Request("page")
searchKey	= Request("searchKey")
searchString = Request("searchString")
adminid		= session("ssBctId")

'==============================================================================
'## 내용 저장(수정) 처리

'트랜젝션 시작
dbget.beginTrans

Select Case mode
	Case "answer"
		'@@ 답변처리
		SQL =	"Update db_momo.dbo.tbl_QnA Set " &_
				"	  ansTitle= '" & ansTitle & "'" &_
				"	, ansContents = '" & ansContents & "'" &_
				"	, ansDate = getdate() " &_
				"	, isanswer = 'Y' " &_
				" Where qnaId = " & qnaId

		dbget.Execute(SQL)


		'답변 메일 발송
		if qstUserMail<>"" then
            
            'response.write ansContents&"a<br>"
            'response.write qstContents&"a<br>"
            'response.write regdate&"a<br>"
            'response.write qstUserName&"a<br>"
            'response.write qstTitle&"a<br>"
            'response.write now()&"a<br>"
        	
            '메일 템플릿 접수
            mailcontent = ReadLocalFile("mail_qna.html", "/admin/momo/qna")
            
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
            '내용 치환
            mailcontent = Replace(mailcontent,"#ansTitle#",ansTitle) 
            '내용 치환
            mailcontent = Replace(mailcontent,"#ansregdate#",now())                    
            '발송
            
            'response.end
            call Send_mail("snowsilver@10x10.co.kr", qstUserMail, "[텐바이텐] 문의하신 질문에 대한 답변입니다", mailcontent)
		end if
				
		msg = "답변처리하였습니다."

		'돌아갈 페이지
		retURL = "QnA_list.asp"

	Case "delete"
		'@@ 내용 삭제

		SQL =	"Update db_momo.dbo.tbl_QnA Set " &_
				"	isusing = 'N'" &_
				" Where qnaId = " & qnaId
		dbget.Execute(SQL)

		msg = "삭제하였습니다."

		'돌아갈 페이지
		retURL = "QnA_list.asp"

End Select


'오류검사 및 반영
If Err.Number = 0 Then   
	dbget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
