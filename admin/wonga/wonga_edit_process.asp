<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������������ ���� ���� ������
' History : �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<% 
dim groupname,category_box_0,field_box_0,gijun_box_0,value_box_0,field,category,yyyymm,chulgocount
	yyyymm = request("yyyymm")					'��¥
	groupname = request("groupname")			'�׷��
	category_box_0 = request("category_box_0")	'ī�װ���
	field_box_0 = request("field_box_0")		'�ʵ��
	gijun_box_0 = request("gijun_box_0")		'���ذ�
	value_box_0 = request("value_box_0")		'��
	field = request("field")					'�ʵ屸�а�
	category = request("category")				'ī�װ����а�
	chulgocount = request("chulgocount")		'��갪
	
dim sql
	sql = "update [db_datamart].[dbo].tbl_month_wonga_category set"
	sql = sql & " gijun_value= '" & gijun_box_0 & "', field_name = '" & field_box_0 & "'"				& VbCrlf
	sql = sql & " where 1=1 and groupname='" & groupname & "' and field='" & field &"' and category='" & category &"'"
	'response.write sql&"ī�װ����̺�����<br>"
	db3_dbget.execute sql

sql = ""
	sql = "update [db_datamart].[dbo].tbl_month_wonga_category set"
	sql = sql & " category_name = '" & category_box_0 & "'"				& VbCrlf
	sql = sql & " where 1=1 and groupname='" & groupname & "' and category='" & category &"'"
	'response.write sql&"ī�װ����̺�����<br>"
	db3_dbget.execute sql

sql = ""				'ī�װ��Է��� �����Ͱ� ���� ���̺� ����Ǵ��� Ȯ���Ѵ�.
	sql = "select field from [db_datamart].[dbo].tbl_month_wonga"
	sql = sql & " where 1=1 and groupname='" & groupname & "' and field='" & field &"' and category='" & category &"'"
	db3_rsget.open sql,db3_dbget,1
	
if not db3_rsget.eof then				'���ڵ尡 �ִٸ� �޾ƿ� ���� ������Ʈ�Ѵ�.
	sql = ""
		sql = "update [db_datamart].[dbo].tbl_month_wonga set"
		sql = sql & " field_value= '" & value_box_0 & "'"				& VbCrlf
		sql = sql & " where 1=1 and groupname='" & groupname & "' and field='" & field &"' and category='" & category &"'"
		'response.write sql&"�������̺�����<br>"
		db3_dbget.execute sql
db3_rsget.close

else								'���ڵ尪�� ���ٸ� ���� �����Ѵ�.
	sql = ""	  	
	sql = "insert into [db_datamart].[dbo].tbl_month_wonga"
	sql = sql & " (yyyymm,groupname,category,field,field_value,count) values"
	sql = sql & " ('" & yyyymm & "'"				& VbCrlf
	sql = sql & ",'" & groupname &"'"
	sql = sql & ",'" & category &"'"			& VbCrlf
	sql = sql & ",'" & field & "'"					& VbCrlf			
	sql = sql & ",'" & value_box_0 & "'"					& VbCrlf
	sql = sql & ",'" & chulgocount & "')"
	'response.write sql&"�������̺�����<br>"
	db3_dbget.execute sql
db3_rsget.close
end if			
%>		
		
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

<script language="javascript">
opener.location.reload();
self.close();
</script>