<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 �ѿ�� ����
'	Description : ���ų�����
'#######################################################
%>

<%
response.write mode	
dim idx , organizerid,mode , selectid
	selectid = request("selectid")
	organizerid = request("organizerid")
	idx = request("idx")
	mode = request("mode")

if 	organizerid = "" or idx = "" then
	response.write "<script>"
	response.write "alert('���ų������ڵ尡 ���ų� ��ǰ�ı��ȣ�� �����ϴ�');"	
	response.write "window.close();"	
	response.write "</script>"	
end if	

	dim sql

'//��ǰ�ı� ������ ���
if mode = "insert" then

		sql = "insert into db_diary2009.dbo.tbl_organizer_eval_list (Eval_idx,organizerid,isusing) values"
		sql = sql & "("
		sql = sql & ""&idx&""
		sql = sql & ", "&organizerid&""
		sql = sql & ",'Y'"		
		sql = sql & ")"
	
		response.write 	sql &"<br>"
		dbget.execute sql

%>
<script language="javascript">
	alert('����Ǿ����ϴ�');
	history.go(-1);
</script>
<%

elseif mode = "update" then

		sql = "update db_diary2009.dbo.tbl_organizer_eval_list set"
		sql = sql & " isusing = '"& selectid &"'"	
		sql = sql & " where idx = "& idx &""
	
		response.write 	sql &"<br>"
		dbget.execute sql
%>
<script language="javascript">
	alert('����Ǿ����ϴ�');
	opener.location.reload();
	window.close();
</script>
<%
end if		
%>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->