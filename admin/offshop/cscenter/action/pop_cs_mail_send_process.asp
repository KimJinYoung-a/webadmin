<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim mailto, title, contents , masteridx
	mailto = requestCheckVar(request.form("mailto"),128)
	title = html2db(request.form("title"))
	contents = html2db(request.form("contents"))
	masteridx = requestCheckVar(request.form("masteridx")	,10)

call sendmailCS(mailto,title,nl2br(contents))

response.write "<script type='text/javascript'>alert('���� �߼۵Ǿ����ϴ�.')</script>"
response.write "<script type='text/javascript'>window.close();</script>"
dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
