<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/QnA_cls.asp"-->
<%
'// ���� ����
'response.write "������.."
'response.end

dim msg, lp, menupos
dim mode, qnaId, adminid
dim ansTitle, ansContents, commCd, mailOk, qstUserMail
dim SQL, mailcontent
dim page, searchDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
qnaId		= RequestCheckvar(Request("qnaId"),10)
mode		= RequestCheckvar(Request("mode"),10)
commCd		= RequestCheckvar(Request("commCd"),10)
mailOk		= RequestCheckvar(Request("mailOk"),10)
qstUserMail	= RequestCheckvar(Request("qstUserMail"),64)
ansTitle	= html2db(Request("ansTitle"))
ansContents	= html2db(Request("ansContents"))
page		= RequestCheckvar(Request("page"),10)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = RequestCheckvar(Request("searchString"),128)
adminid		= session("ssBctId")
if ansTitle <> "" then
	if checkNotValidHTML(ansTitle) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end If
if ansContents <> "" then
	if checkNotValidHTML(ansContents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end if
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end if
'// ���� ���� �� ó��
dim qstUserName,qstContents, regdate,qstTitle
qstUserName		= RequestCheckvar(Request("qstUserName"),16)
qstContents		= Request("qstContents")
regdate		= RequestCheckvar(Request("regdate"),10)
qstTitle		= RequestCheckvar(Request("qstTitle"),64)
if qstContents <> "" then
	if checkNotValidHTML(qstContents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end If
if qstTitle <> "" then
	if checkNotValidHTML(qstTitle) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end if

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

'==============================================================================
'## ���� ����(����) ó��

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "answer"
		'@@ �亯ó��
		SQL =	"Update db_academy.dbo.tbl_QnA Set " &_
				"	  ansTitle= '" & ansTitle & "'" &_
				"	, ansContents = '" & ansContents & "'" &_
				"	, ansDate = getdate() " &_
				"	, isanswer = 'Y' " &_
				" Where qnaId = " & qnaId

		dbACADEMYget.Execute(SQL)


		'�亯 ���� �߼�
		if (mailOk = "����") and (Cstr(qstUserMail)<>"") then
			
			'���� ���ø� ����
            mailcontent = ReadLocalFile("tpl_fingers_qna.html", "/academy/lib/mail_templete")
            
            '���� ġȯ
            '''mailcontent = Replace(mailcontent,"#contents#",nl2br(db2html(ansContents)))
            
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
            

            '�߼�
            call Send_mail("customer@thefingers.co.kr", qstUserMail, "�Բ� ���� ��ſ� �ΰŽ�", mailcontent)
		end if

		msg = "�亯ó���Ͽ����ϴ�."

		'���ư� ������
		retURL = "QnA_view.asp?menupos=" & menupos & "&qnaId=" & qnaId & param

	Case "delete"
		'@@ ���� ����

		SQL =	"Update db_academy.dbo.tbl_QnA Set " &_
				"	isusing = 'N'" &_
				" Where qnaId = " & qnaId
		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "QnA_list.asp?menupos=" & menupos & param

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->