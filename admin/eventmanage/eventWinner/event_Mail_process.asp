<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_Mail_process.asp
' Description :  �̺�Ʈ ��÷�� ���� �߼� ó��
' History : 2007.10.01 ������
'####################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<%

'// Mail ����
function SaveMailLog(byval eCd,byval mlT,byval mlC,byval rpNm,byval rpMl,byval rgUser ,byval SendYn)
	'// eCd �̺�Ʈ �ڵ�
	'// mlT ���� Ÿ��Ʋ
	'// mlC ���� ����
	'// rpNm �������
	'// rpMl ������� �����ּ�
	'// rgUser ����Ÿ���� ������(�����ڿ�)
	'// SendYn �߼� ����

	dim fnSQL,cnt

	rguser = session("ssBctId")

	fnSQL =" SELECT count(evt_code) as cnt " &_
			" FROM db_event.dbo.tbl_event_mail_log " &_
			" WHERE evt_code='" & eCd &"'"
	rsget.open fnSQL,dbget,1

	If Not rsget.eof Then
		cnt = rsget("cnt")
	End If

	rsget.close

	if cnt>0 then
		fnSQL =" UPDATE db_event.dbo.tbl_event_mail_log " &_
				" set evt_code='" & eCd & "' " &_
				" ,mailTitle='" & mlT & "' " &_
				" ,mailCont='" & mlC & "' " &_
				" ,replyName='" & rpNm & "' " &_
				" ,replyMail='" & rpMl & "' " &_
				" ,regUser='" & rgUser & "' " &_
				" ,regDate = getdate() " &_
				" ,isSended ='" & SendYn & "'" &_
				" WHERE evt_code='" & eCd & "'"

	else
		fnSQL =" INSERT INTO db_event.dbo.tbl_event_mail_log " &_
				" (evt_code,mailTitle,mailCont,replyName,replyMail,regUser,regDate) " &_
				" values " &_
				" ('" & eCd &"','" & mlT & "','" & mlC & "','" & rpNm & "','" & rpMl & "','" & rgUser & "',getdate(),'" & SendYn & "' " &_
				" ) "
	end if
	'response.write fnSQL
	dbget.execute(fnSQL)

end function

function SendMail(byval mailfrom, byval mailto, byval mailtitle, byval mailcontent)

		dim cdoMessage,cdoConfig
        'On Error Resume Next
		Set cdoConfig = CreateObject("CDO.Configuration")

		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="210.92.223.238"
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "administrator"
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
		cdoConfig.Fields.Update

		Set cdoMessage = CreateObject("CDO.Message")

		Set cdoMessage.Configuration = cdoConfig

		cdoMessage.To 				= mailto
		cdoMessage.From 			= mailfrom
		cdoMessage.SubJect 	= mailtitle
		'���� ������ �ؽ�Ʈ�� ��� cdoMessage.TextBody, html�� ��� cdoMessage.HTMLBody
		cdoMessage.HTMLBody	= mailcontent
		cdoMessage.Send

		Set cdoMessage = nothing
		Set cdoConfig = nothing
       ' On Error Goto 0
end function


'################-- ó�� Process ���� --###################

dim evtCode,mailTitle,mailContents,replyName,replyMail,regUser,mode

	evtCode 	= request("eC")
	mailTitle 	= html2db(request("mlTitle"))
	mailContents 	= html2db(request("mlCont"))
	replyName 	= html2db(request("rpName"))
	replyMail 	= html2db(request("rpMail"))
	regUser 	= session("ssBctId")

	arridx = chkarray(request("arridx"))

	if mailContents<>"" then mailContents=Replace(mailContents, vbcrlf,"<br>")

	mode=request("mode")

'// ������ ����
dim fso,contFile,MailPath,MailForm,msg

	MailPath = server.mappath("/lib/email/email_event.htm")

	set fso = Server.Createobject("Scripting.filesystemObject")
	set contFile = fso.Opentextfile(MailPath)

	MailForm = contFile.readAll

	contFile.close

	MailForm= replace(MailForm,"$$MAILCONTENTS$$",mailContents)



dbget.begintrans

if mode="save" then
	'//��������
	SaveMailLog evtCode,mailTitle,mailContents,replyName,replyMail,regUser,"N"
	msg ="���ϳ����� �����Ͽ����ϴ�."
elseif mode="send" then

	SaveMailLog evtCode,mailTitle,mailContents,replyName,replyMail,regUser,"Y"
	fnWinnerMailSended evtCode,arridx

	arridx = split(arridx)
	dim loopCnt

	for loopcnt =0 to ubound(arridx)
		SendMail db2html(replyName &"<"& replyMail &">"), arridx(loopcnt), db2html(mailTitle), db2html(MailForm)
	next

	msg="������ �߼��Ͽ����ϴ�."
end if

'dbget.execute(strSQL)


'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "');"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"</script>"


	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->