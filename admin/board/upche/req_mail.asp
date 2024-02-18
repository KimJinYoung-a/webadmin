<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 입점문의
' History : 서동석 생성
'			2020.12.10 한용민 수정(이메일발송 메일러로 이관)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/maillib2.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%
dim user,mailfrom, mailto, mailname, mailtitle, mailcontent, content,id,userid
dim companyrequest
userid=requestCheckvar(request("userid"),32)
user=requestCheckvar(request("user"),32)
mailto=requestCheckvar(request("mailto"),128)
mailname=request("mailname")
mailtitle="[텐바이텐] 입점문의에 대한 답변 입니다."
content= nl2br(request("content"))
mailfrom=getmail(user)
id=requestCheckvar(request("id"),10)
 
set companyrequest = New CCompanyRequest
companyrequest.finish(id)

dim fs,dirPath,fileName,objFile
' 파일을 불러와서
Set fs = Server.CreateObject("Scripting.FileSystemObject")
dirPath = server.mappath("/admin/board/upche")
'fileName = dirPath & "\\req_mail.htm"
fileName = dirPath & "\\req_mail.html"

Set objFile = fs.OpenTextFile(fileName,1)
mailcontent = objFile.readall

Set objFile = Nothing
Set fs = Nothing

mailcontent = replace(mailcontent,"$MAILNAME$", mailname)
mailcontent = replace(mailcontent,"$USER$", user)
mailcontent = replace(mailcontent,"$CONTENT$", content)

'response.write mailcontent
'dbget.close()	:	response.End

if C_ADMIN_AUTH then
    mailtitle= "(재발송)"&mailtitle
end if

dim oMail
set oMail = New MailCls         '' mailLib2
	oMail.ReceiverMail	= mailto
	oMail.MailTitles	= mailtitle
	oMail.MailConts 	= mailcontent
	oMail.MailerMailGubun = 13		' 메일러 자동메일 번호
	oMail.Send_TMSMailer()		'TMS메일러
SET oMail = nothing
'sendmail mailfrom, mailto, mailtitle, mailcontent

response.write "<script type='text/JavaScript'>"
response.write " 	alert('메일을 발송하였습니다.');"
response.write " 	history.go(-1);"
response.write " 	opener.location.reload();"
response.write "</script>"
dbget.close()	:	response.End

public function getmail(byval username)
	dim sql
	sql =" select top 1 usermail "
	sql = sql + " from [db_partner].[dbo].tbl_user_tenbyten"
	sql = sql + " where userid ='" + userid + "'" + vbcrlf
	
	rsget.open sql,dbget,1
	if not rsget.eof then
		getmail=rsget("usermail")			
	end  if
	rsget.close
end function

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->