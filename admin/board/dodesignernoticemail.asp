<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [업체]공지사항
' Hieditor : 서동석 생성
'			 2023.10.23 한용민 수정(이메일발송 cdo->메일러로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/maillib2.asp" -->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<%
Dim id, sqlstr, mode, refer, contents, email, title, mduserid, catecode, targt, nboardmail
dim MailSendedCount,WillMailSendCount '메일 나눠서 보내기용
	MailSendedCount = request("MailSendedCount")
	id			= request("id")
	mode		= request("mode")
	mduserid    = request("mduserid")
	catecode    = request("catecode")
	targt		= request("targt")

refer = request.ServerVariables("HTTP_REFERER")

if MailSendedCount="" then MailSendedCount=0

sqlstr = "select top 1 * from "
sqlstr = sqlstr + " [db_board].[dbo].tbl_designer_notice with (nolock)"
sqlstr = sqlstr + " where board_idx=" + CStr(id)

rsget.CursorLocation = adUseClient
rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
if not rsget.EOF then
	contents = db2html(rsget("content"))
	email = db2html(rsget("email"))
	title = db2html(rsget("title"))
end if
rsget.Close

title=stripHTML(trim(title))

if mode = "upcheall" then
	set nboardmail = new CBoard
		'nboardmail.FRectContents = nl2br(contents)
		nboardmail.FRectContents	= contents
		nboardmail.FRectEmail		= email
		nboardmail.FRectTitle		= "[텐바이텐 공지메일]" & title
		nboardmail.FRectMDid		= mduserid
		nboardmail.FRectCatCD		= catecode
		nboardmail.FRectTarget		= targt
		nboardmail.MailSendedCount	= Cint(MailSendedCount)
		nboardmail.design_notice_mail_send

	WillMailSendCount= nboardmail.WillMailSendCount

	set nboardmail = nothing
end if

if Cint(MailSendedCount) < Cint(WillMailSendCount) then
	response.write "<h1>메일 발송중입니다. 창을 닫지마세요</h1>"
	response.write "총 " & WillMailSendCount & " 페이지중 " & MailSendedCount & " 페이지 발송중 "
	response.write "<script type='text/javascript'>window.setTimeout(""document.location='/admin/board/dodesignernoticemail.asp?id=" & CStr(id) & "&mode=" & Cstr(mode) & "&mduserid=" & Cstr(mduserid) & "&catecode=" & Cstr(catecode) & "&targt=" & Cstr(targt) & "&MailSendedCount=" & Cint(MailSendedCount)+1 & "'"", 5000);</script>"
	dbget.close()	:	response.End
else
	response.write "<script type='text/javascript'>alert('발송 되었습니다..');//window.close();</script>"
	dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->


