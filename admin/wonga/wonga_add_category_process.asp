<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������������ ī�װ� �űԱ׷� ���� ������
' History : �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
dim category_box_0 , field_box_0 ,gubunbox ,groupname ,yyyy ,mm ,yyyymm ,countvalue  , gijun_box_0,gijuncount
dim category_box_0_re , field_box_0_re ,flag_field_re, categorycount , fieldcount ,  gijun_box_0_re
dim add_category
	category_box_0 = request("category_box_0")&","		' �޸��� �����ؼ� �迭�� �ֱ����� �޸��� ���δ�.
	field_box_0 = request("field_box_0")&","
	gubunbox = request("gubun_submit")		'�׷찪�� �޾ƿ´�
	groupname = request("groupname")
	yyyy = request("yyyy")					'���� �޾ƿ´�.
	mm = request("mm")						'���� �޾ƿ´�.
	yyyymm = yyyy&mm						'�޾ƿ� ��� ���� ��ģ��.
	add_category = request("add_category")	' �̹������ϴ±׷쿡 ī�װ� �߰��� ���ؼ� ���� ī�װ����޾ƿ�...
	countvalue = request("count")			'����� ���� ī��Ʈ ���� �޾ƿ´�.	
	gijun_box_0 = request("gijun_box_0")&","	'���ذ��� �޾ƿ´�.

category_box_0_re = split(category_box_0,",")	'�޸��� �������� ������ ¥����.
categorycount = ubound(category_box_0_re)
field_box_0_re = split(field_box_0,",")
fieldcount = ubound(field_box_0_re)
gijun_box_0_re = split(gijun_box_0,",")
gijuncount = ubound(gijun_box_0_re)

'response.write "categorycount : "&categorycount&"<br>"
'response.write "category_box_0 : "&category_box_0&"<br>"
'response.write "fieldcount : "&fieldcount&"<br>"
'response.write "field_box_0 : "&field_box_0&"<br>"
'response.write "gijuncount : "&gijuncount&"<br>"
'response.write "gijun_box_0 : "&gijun_box_0&"<br>"

dim i ,t,idx_f, sql , sql1
idx_f = 0		'�ʵ��ȣ�� 0���� �ű������ �ε�������
%>

<% if gubunbox = "" then 
'##################################################################	�׷��� ������ ī�װ��������
	
		for i = 0 to fieldcount -1	
			sql = "insert into [db_datamart].[dbo].tbl_month_wonga_category"
			sql = sql & " (groupname,category,category_name,category_isusing,field,field_name,gijun_value) values"
			sql = sql & "('" & groupname & "'"				& VbCrlf
			sql = sql & ",'" & 0 &"'"			& VbCrlf
			sql = sql & ",'" & category_box_0_re(0) & "'"					& VbCrlf			
			sql = sql & ",'y'"					& VbCrlf
			sql = sql & ",'" & idx_f &"'"			& VbCrlf
			sql = sql & ",'" & field_box_0_re(i) &"'"			& VbCrlf
			sql = sql & ",'" & gijun_box_0_re(i) & "')"
			'response.write sql&"ī�װ�"&t &"<br>"
			db3_dbget.execute sql
		idx_f = idx_f + 1
		next
		idx_f = 0
%>

<script language="javascript">
alert('ó�� �Ǿ����ϴ�. ����Ͻ� ī�װ��� �����͸������ž� ���� ó���˴ϴ�.');
location.replace('/admin/wonga/wonga_add_category.asp?gubunbox='+'<%= groupname %>');
</script>	

<% else 
'##################################################################	�׷��� ������ ī�װ��������

		for i = 0 to fieldcount -1	
			sql = "insert into [db_datamart].[dbo].tbl_month_wonga_category"
			sql = sql & " (groupname,category,category_name,category_isusing,field,field_name,gijun_value) values"
			sql = sql & "('" & gubunbox & "'"				& VbCrlf
			sql = sql & ",'" & add_category &"'"			& VbCrlf
			sql = sql & ",'" & category_box_0_re(0) & "'"					& VbCrlf			
			sql = sql & ",'y'"					& VbCrlf
			sql = sql & ",'" & idx_f &"'"			& VbCrlf
			sql = sql & ",'" & field_box_0_re(i) &"'"			& VbCrlf
			sql = sql & ",'" & gijun_box_0_re(i) & "')"
			'response.write sql&"ī�װ�"&t &"<br>"
			db3_dbget.execute sql
		idx_f = idx_f + 1
		next
		idx_f = 0
%>
<script language="javascript">
alert('ó�� �Ǿ����ϴ�. ����Ͻ� ī�װ��� �����͸������ž� ���� ó���˴ϴ�.');
location.replace('/admin/wonga/wonga_add_category.asp?gubunbox='+'<%= gubunbox %>');
</script>
<% end if %>


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->