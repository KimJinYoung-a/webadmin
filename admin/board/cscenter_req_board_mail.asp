<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp" -->
<%
public function getmail(byval uid)
	dim sql
	sql =" select top 1 IsNULL(email,'') as email "
	sql = sql + " from [db_partner].[dbo].tbl_partner"
	sql = sql + " where id='" + uid + "'" + vbcrlf

	rsget.open sql,dbget,1
	if not rsget.eof then
		getmail=rsget("email")
	end  if
	rsget.close
end function

dim user,mailfrom, mailto, mailname, mailtitle, mailcontent, content,id
dim companyrequest

user        = request("user")
mailto      = request("mailto")
mailname    = request("mailname")
mailtitle   = "10x10�Դϴ�."
content     = nl2br(request("content"))
mailfrom    = getmail(session("ssBctId"))
id          = request("id")

if (Len(mailto)<3) or (InStr(mailto,"@")<1) then
    response.write "������ �̸��� �ּҰ� �ùٸ��� �ʾ� �߼��� �� �����ϴ�. "
    response.end
end if

if (mailfrom="") then 
    response.write "�߼��� �̸��� �ּҰ� ���� �߼��� �� �����ϴ�. ȸ������ ������ �����."
    response.end
end if


set companyrequest = New CCompanyRequest
companyrequest.finish(id)

dim fs,dirPath,fileName,objFile
' ������ �ҷ��ͼ�
Set fs = Server.CreateObject("Scripting.FileSystemObject")
dirPath = server.mappath("/admin/board")
fileName = dirPath & "\\cscenter_req_board_mail.htm"

Set objFile = fs.OpenTextFile(fileName,1)
mailcontent = objFile.readall

Set objFile = Nothing
Set fs = Nothing


		'//=======  ���� �߼� =========/
		dim oMail
		dim MailHTML

		set oMail = New MailCls

		'oMail.MailType		 = 8 '���� ������ ������ (mailLib2.asp ����)
		oMail.MailTitles	= "�ٹ����� �Դϴ�."
		oMail.SenderNm 		= "�ٹ�����"
		oMail.SenderMail 	= mailfrom
		oMail.AddrType 		= "string"
		oMail.ReceiverNm 	= mailname
		oMail.ReceiverMail	= mailto

		MailHTML= mailcontent

		IF MailHTML="" Then
			SET oMail = nothing
			response.write "<script>alert('���Ϲ߼��� ���� �Ͽ����ϴ�.');history.go(-1);</script>"
			dbget.close()	:	response.End
	    End IF

		'// ���� ���Ͽ� ���� ġȯ
		MailHTML = replace(MailHTML,"$MAILNAME$", mailname) ' �޴³� �̸�
		MailHTML = replace(MailHTML,"$USER$", user) ' �����³� �̸�
		MailHTML = replace(MailHTML,"$CONTENT$",content) '���� ����

		oMail.MailConts = MailHTML

		'oMail.Send_Mailer()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		'oMail.ReceiverMail	 = "yanbest@naver.com"
		'oMail.Send_Mailer()

		SET oMail = nothing

	 response.write "<script language='JavaScript'>alert('������ �߼��Ͽ����ϴ�.');history.go(-1);</script>"
	 dbget.close()	:	response.End


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->