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
mailtitle   = "10x10입니다."
content     = nl2br(request("content"))
mailfrom    = getmail(session("ssBctId"))
id          = request("id")

if (Len(mailto)<3) or (InStr(mailto,"@")<1) then
    response.write "수신자 이메일 주소가 올바르지 않아 발송할 수 없습니다. "
    response.end
end if

if (mailfrom="") then 
    response.write "발송자 이메일 주소가 없어 발송할 수 없습니다. 회원정보 수정후 사용요망."
    response.end
end if


set companyrequest = New CCompanyRequest
companyrequest.finish(id)

dim fs,dirPath,fileName,objFile
' 파일을 불러와서
Set fs = Server.CreateObject("Scripting.FileSystemObject")
dirPath = server.mappath("/admin/board")
fileName = dirPath & "\\cscenter_req_board_mail.htm"

Set objFile = fs.OpenTextFile(fileName,1)
mailcontent = objFile.readall

Set objFile = Nothing
Set fs = Nothing


		'//=======  메일 발송 =========/
		dim oMail
		dim MailHTML

		set oMail = New MailCls

		'oMail.MailType		 = 8 '메일 종류별 고정값 (mailLib2.asp 참고)
		oMail.MailTitles	= "텐바이텐 입니다."
		oMail.SenderNm 		= "텐바이텐"
		oMail.SenderMail 	= mailfrom
		oMail.AddrType 		= "string"
		oMail.ReceiverNm 	= mailname
		oMail.ReceiverMail	= mailto

		MailHTML= mailcontent

		IF MailHTML="" Then
			SET oMail = nothing
			response.write "<script>alert('메일발송이 실패 하였습니다.');history.go(-1);</script>"
			dbget.close()	:	response.End
	    End IF

		'// 실제 메일에 정보 치환
		MailHTML = replace(MailHTML,"$MAILNAME$", mailname) ' 받는넘 이름
		MailHTML = replace(MailHTML,"$USER$", user) ' 보내는넘 이름
		MailHTML = replace(MailHTML,"$CONTENT$",content) '메일 내용

		oMail.MailConts = MailHTML

		'oMail.Send_Mailer()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		'oMail.ReceiverMail	 = "yanbest@naver.com"
		'oMail.Send_Mailer()

		SET oMail = nothing

	 response.write "<script language='JavaScript'>alert('메일을 발송하였습니다.');history.go(-1);</script>"
	 dbget.close()	:	response.End


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->