<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� USB ����
' History : 2008.06.30 �ѿ�� ���� 
'           2008.09.25 ������ ����- Key Int��char ����
'           2008.09.25 �ѿ�� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/admin_keyclass.asp" -->
<%
Dim key_idx ,  teamname, username ,username_detail , del_isusing , mode , idx
	key_idx = request("key_idx")
	teamname = request("teamname")	
	username = request("username")
	username_detail = request("username_detail")	
	del_isusing = request("del_isusing")
	mode = request("mode")
	idx = request("idx")
	'response.write mode
dim sql

'// �������
if mode = "edit" then
	sql = "update db_partner.dbo.tbl_admin_key set"	+vbcrlf
	sql = sql & " key_idx = '"& key_idx &"'"	+vbcrlf
	sql = sql & " ,teamname = '"& teamname &"'"	+vbcrlf
	sql = sql & " ,username = '"& username &"'"	+vbcrlf
	sql = sql & " ,username_detail = '"& username_detail &"'"	+vbcrlf
	sql = sql & " ,del_isusing = '"& del_isusing &"'"	+vbcrlf
	sql = sql & " where idx = '"& idx &"'"	+vbcrlf		
	
	'response.write sql
	dbget.execute sql	
%>
	<script language="javascript">
		location.href="/admin/member/admin_keylist.asp";
	</script>
<%
'// �ű��Է¸��
elseif mode = "new" then
	sql = "insert into db_partner.dbo.tbl_admin_key (key_idx,teamname,username,username_detail,del_isusing) values ("
	sql = sql & " '"& key_idx & "','" & teamname & "','" & username & "','" & username_detail & "','" & del_isusing & "')"

	'response.write sql
	dbget.execute sql
%>
	<script language="javascript">
		opener.location.reload();
		window.close();
	</script>
<%
end if
%>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
