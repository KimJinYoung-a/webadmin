<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [��ü]��������
' Hieditor : ������ ����
'			 2023.10.23 �ѿ�� ����(�̸��Ϲ߼� cdo->���Ϸ��� ����)
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
dim MailSendedCount,WillMailSendCount '���� ������ �������
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
		nboardmail.FRectTitle		= "[�ٹ����� ��������]" & title
		nboardmail.FRectMDid		= mduserid
		nboardmail.FRectCatCD		= catecode
		nboardmail.FRectTarget		= targt
		nboardmail.MailSendedCount	= Cint(MailSendedCount)
		nboardmail.design_notice_mail_send

	WillMailSendCount= nboardmail.WillMailSendCount

	set nboardmail = nothing
end if

if Cint(MailSendedCount) < Cint(WillMailSendCount) then
	response.write "<h1>���� �߼����Դϴ�. â�� ����������</h1>"
	response.write "�� " & WillMailSendCount & " �������� " & MailSendedCount & " ������ �߼��� "
	response.write "<script type='text/javascript'>window.setTimeout(""document.location='/admin/board/dodesignernoticemail.asp?id=" & CStr(id) & "&mode=" & Cstr(mode) & "&mduserid=" & Cstr(mduserid) & "&catecode=" & Cstr(catecode) & "&targt=" & Cstr(targt) & "&MailSendedCount=" & Cint(MailSendedCount)+1 & "'"", 5000);</script>"
	dbget.close()	:	response.End
else
	response.write "<script type='text/javascript'>alert('�߼� �Ǿ����ϴ�..');//window.close();</script>"
	dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->


