<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008�� 11�� 6�� �ѿ�� ����
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<% 
dim mode ,  event_type , evt_code, idx_order, isusing , idx
	mode = request("mode")
	event_type = request("event_type")
	evt_code = request("evt_code")
	idx_order = request("idx_order")
	isusing = request("isusing")
	idx = request("idx")
	
if event_type = "" or evt_code = "" then
	response.write "<script>"
	response.write "alert('�̺�Ʈ�ڵ峪 �̺�ƮŸ�� �Ķ��Ÿ�������ϴ�');"
	response.write "history.go(-1);"
	response.write "</script>"		
end if	

dim sql

'//�ű��߰�
if mode = "new" then

sql = ""
sql = "insert into db_diary2009.dbo.tbl_event (evt_code , event_type , isusing , idx_order) values (" & vbcrlf
sql = sql & " '"&evt_code&"'" & vbcrlf
sql = sql & " ,'"&event_type&"'" & vbcrlf
sql = sql & " ,'"&isusing&"'" & vbcrlf
sql = sql & " ,'"&idx_order&"'" & vbcrlf
sql = sql & " )" & vbcrlf

response.write sql&"<br>"
dbget.execute sql
%>
<script language="javascript">
	opener.location.reload();
	window.close();
</script>

<%
elseif mode = "edit" then

'//�������
if idx = "" then
	response.write "<script>"
	response.write "alert('idx �Ķ��Ÿ ���̾����ϴ�');"
	response.write "history.go(-1);"
	response.write "</script>"		
end if

sql = ""
sql = "update db_diary2009.dbo.tbl_event set" & vbcrlf
sql = sql & " evt_code = '"& evt_code &"'"  & vbcrlf
sql = sql & " ,event_type = '"& event_type &"'"  & vbcrlf
sql = sql & " ,isusing = '"& isusing &"'"  & vbcrlf
sql = sql & " ,idx_order = '"& idx_order &"'" & vbcrlf
sql = sql & " where idx = "& idx &""  & vbcrlf

response.write sql&"<br>"
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