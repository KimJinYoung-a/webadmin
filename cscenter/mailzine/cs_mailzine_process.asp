<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������ ���
' History : 2007.12.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/mailzine/mailzinecls.asp"-->
<%
dim idx , sql
dim username,usermail,regdate,isusing
	idx = requestCheckVar(getNumeric(trim(request("idx"))),10)
	username = requestCheckVar(trim(request("username")),32)
	usermail = requestCheckVar(trim(request("usermail")),128)
	regdate = requestCheckVar(trim(request("regdate")),32)
	isusing = requestCheckVar(trim(request("isusing")),1)

'�űԵ��
if idx="" then
	if username <> "" and not(isnull(username)) then
		username = ReplaceBracket(username)
	end If
	if usermail <> "" and not(isnull(usermail)) then
		usermail = ReplaceBracket(usermail)
	end If

	sql = "insert into db_user.dbo.tbl_mailzine_notmember (username,usermail,isusing) values (" &vbcrlf
	sql = sql & " '"&html2db(username)&"','"&html2db(usermail)&"','"&isusing&"'" &vbcrlf
	sql = sql & " )"

	'response.write sql &"<br>"
	dbget.execute sql
	
'����	
else
	if username <> "" and not(isnull(username)) then
		username = ReplaceBracket(username)
	end If
	if usermail <> "" and not(isnull(usermail)) then
		usermail = ReplaceBracket(usermail)
	end If

	sql = "update db_user.dbo.tbl_mailzine_notmember set" &vbcrlf
	sql = sql & " username='"&html2db(username)&"'" &vbcrlf
	sql = sql & " ,usermail='"&html2db(usermail)&"'" &vbcrlf 
	sql = sql & " ,isusing='"&isusing&"'" &vbcrlf 
	sql = sql & " where idx="&idx&""
	
	'response.write sql &"<br>"
	dbget.execute sql
end if		
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script type='text/javascript'>
	alert('����Ǿ����ϴ�');
	opener.location.reload();
	self.close();
</script>