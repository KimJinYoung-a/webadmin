<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� qna ����
' Hieditor : 2009.11.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode, qnaId, adminid
dim ansTitle, ansContents, commCd, mailOk, qstUserMail
dim SQL, mailcontent
dim page, searchDiv, searchKey, searchString, retURL
dim qstUserName,regdate,qstTitle , qstContents


'// ���� ���� �� ó��
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
'## ���� ����(����) ó��

'Ʈ������ ����
dbget.beginTrans

Select Case mode
	Case "answer"
		'@@ �亯ó��
		SQL =	"Update db_momo.dbo.tbl_QnA Set " &_
				"	  ansTitle= '" & ansTitle & "'" &_
				"	, ansContents = '" & ansContents & "'" &_
				"	, ansDate = getdate() " &_
				"	, isanswer = 'Y' " &_
				" Where qnaId = " & qnaId

		dbget.Execute(SQL)


		'�亯 ���� �߼�
		if qstUserMail<>"" then
            
            'response.write ansContents&"a<br>"
            'response.write qstContents&"a<br>"
            'response.write regdate&"a<br>"
            'response.write qstUserName&"a<br>"
            'response.write qstTitle&"a<br>"
            'response.write now()&"a<br>"
        	
            '���� ���ø� ����
            mailcontent = ReadLocalFile("mail_qna.html", "/admin/momo/qna")
            
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#ansContents#",nl2br(ansContents))
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#qstContents#",nl2br(qstContents))
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#regdate#",FormatDate(regdate,"0000-00-00"))
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#qstUserName#",qstUserName)
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#qstTitle#",qstTitle)     
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#ansTitle#",ansTitle) 
            '���� ġȯ
            mailcontent = Replace(mailcontent,"#ansregdate#",now())                    
            '�߼�
            
            'response.end
            call Send_mail("snowsilver@10x10.co.kr", qstUserMail, "[�ٹ�����] �����Ͻ� ������ ���� �亯�Դϴ�", mailcontent)
		end if
				
		msg = "�亯ó���Ͽ����ϴ�."

		'���ư� ������
		retURL = "QnA_list.asp"

	Case "delete"
		'@@ ���� ����

		SQL =	"Update db_momo.dbo.tbl_QnA Set " &_
				"	isusing = 'N'" &_
				" Where qnaId = " & qnaId
		dbget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "QnA_list.asp"

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
