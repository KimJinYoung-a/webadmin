<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������������ �űԵ��
' History : 2007.09.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/wonga/wonga_month_class.asp"-->

<% 
dim mode,groupname,groupname1 , yyyymm,sql , sql1
mode = request("mode")
groupname = request("groupname")
yyyymm = request("yyyymm")

dim category,field
	category = cint(request("category"))
	field = cint(request("field"))
%>

<% 
if mode = "total_del" then			'�׷���ü ����
	sql = "delete from db_datamart.dbo.tbl_month_wonga where groupname='"& groupname &"'"
	'response.write sql+"<br>"			'������ ȭ�鿡 �ѷ����� 	
	dbget.execute sql
	sql1 = "delete from db_datamart.dbo.tbl_month_wonga_category where groupname='"& groupname &"'"
	'response.write sql1+"<br>"			'������ ȭ�鿡 �ѷ����� 	
	dbget.execute sql1

elseif mode = "del" then			'���ú� ����
	sql = "delete from db_datamart.dbo.tbl_month_wonga" 
	sql = sql & " where 1=1 and groupname='"& groupname &"' and category = '" & category &"' and field = '" & field &"'"
	'response.write sql+"<br>"			'������ ȭ�鿡 �ѷ����� 	
	'dbget.execute sql
	sql1 = "delete from db_datamart.dbo.tbl_month_wonga_category"
	sql1 = sql1 & " where 1=1 and groupname='"& groupname &"' and category = '" & category &"' and field = '" & field &"'"
	'response.write sql1+"<br>"			'������ ȭ�鿡 �ѷ����� 	
	'dbget.execute sql1

else		'��¥�� ����
	sql = "delete from db_datamart.dbo.tbl_month_wonga where groupname='"& groupname &"' and yyyymm='"& yyyymm &"'"
	'response.write sql+"<br>"			'������ ȭ�鿡 �ѷ����� 	
	dbget.execute sql
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<script language="javascript">
opener.location.reload();
self.close();
</script>		