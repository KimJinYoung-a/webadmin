<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Thanks 10x10 ����  
' History : 2009.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%

dim idx,idxsum ,itemid, rectitemid,mode ,comment
dim sql
	idx = request("idx")
	mode = request("mode")
	comment = request("comment")

'// �Ǽ��� ���� �ݿ� 		
if mode = "" then
	idxsum = request("idx")
	idx = left(idxsum,len(idxsum)-1)
	
	sql = "update db_culture_station.dbo.tbl_thanks_10x10 set"
	sql = sql & " isusing_display = 'Y'"
	sql = sql & " where 1=1 and idx in ( "& idx &" )"

	'response.write sql
	dbget.execute sql

'// �亯 �ڸ�Ʈ �ű�
elseif mode = "add" then

	sql = "insert into db_culture_station.dbo.tbl_thanks_10x10_comment (idx,comment) values ("& idx &",'"& html2db(comment) &"')"
	
	'response.write sql
	dbget.execute sql

'// �亯 �ڸ�Ʈ ����
elseif mode = "edit" then	
	

	sql = "update db_culture_station.dbo.tbl_thanks_10x10_comment set comment = '"& html2db(comment) &"' where idx = "& idx &""
	
	'response.write sql
	dbget.execute sql

'//�亯 ����
elseif mode = "comment_del" then
	
	sql = "delete from db_culture_station.dbo.tbl_thanks_10x10_comment where idx = "& idx &""
	
	'response.write sql
	dbget.execute sql
	
'//���� ����(�亯������)	
elseif mode = "del" then
	
	sql = "update db_culture_station.dbo.tbl_thanks_10x10 set isusing_del = 'Y'  where idx = "& idx &""	
	
	'response.write sql
	dbget.execute sql
	sql = ""
	
	sql = "delete from db_culture_station.dbo.tbl_thanks_10x10_comment where idx = "& idx &""	
	
	'response.write sql
	dbget.execute sql
	
end if	
%>	

	<script>
	opener.location.reload();
	alert('ó�� �Ǿ����ϴ�.');
	self.close();
	</script>	
	
<!-- #include virtual="/lib/db/dbclose.asp" -->
	
