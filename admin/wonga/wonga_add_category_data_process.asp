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
dim category_box_0,value_box_0 , gijun_box_0 ,gubunbox ,groupname ,yyyy ,mm ,yyyymm ,countvalue 
dim category_box_0_re,value_box_0_re , gijun_box_0_re , categorycount , gijuncount,valuecount
	category_box_0 = request("category_box_0")&","		' �޸��� �����ؼ� �迭�� �ֱ����� �޸��� ���δ�.
	value_box_0 = request("value_box_0")&","		' �޸��� �����ؼ� �迭�� �ֱ����� �޸��� ���δ�.
	gijun_box_0 = request("gijun_box_0")&","
	gubunbox = request("gubun_submit")		'�׷찪�� �޾ƿ´�
	groupname = request("groupname")
	yyyy = request("yyyy")					'���� �޾ƿ´�.
	mm = request("mm")						'���� �޾ƿ´�.
	yyyymm = yyyy&mm						'�޾ƿ� ��� ���� ��ģ��.
	countvalue = request("count")			'����� ���� ī��Ʈ ���� �޾ƿ´�.	

category_box_0_re = split(category_box_0,",")		'�޸��� �������� �迭�� ¥����.
categorycount = ubound(category_box_0_re)
value_box_0_re = split(value_box_0,",")
valuecount = ubound(value_box_0_re)
gijun_box_0_re = split(gijun_box_0,",")
gijuncount = ubound(gijun_box_0_re)

'response.write yyyymm&"<br>"
'response.write "countvalue : "&countvalue&"<br>"
'response.write "categorycount : "&categorycount&"<br>"
'response.write "category_box_0 : "&category_box_0&"<br>"
'response.write "gijuncount : "&valuecount&"<br>"
'response.write "value_box_0 : "&value_box_0&"<br>"
'response.write "gijuncount : "&gijuncount&"<br>"
'response.write "gijun_box_0 : "&gijun_box_0&"<br>"

dim i ,t,idx_f, sql 
idx_f = 0		'�ʵ��ȣ�� 0���� �ű������ �ε�������
%>

<%
dim sql1 ,ftotalcount

for t = 0 to categorycount - 1		'ī�װ� ������ŭ ����

	sql1 = "select"
	sql1 = sql1 & " field"
	sql1 = sql1 & " from db_datamart.dbo.tbl_month_wonga_category"
	sql1 = sql1 & " where 1=1 and groupname= '"& gubunbox &"' and category_isusing='y' and category = '"& t &"'"
	sql1 = sql1 & " group by field" 	
	
	db3_rsget.open sql1,db3_dbget,1
	'response.write sql1&"<br>"	
	ftotalcount = db3_rsget.recordcount		'������ ī�װ��ȿ� �ʵ尡 ����� ������ �޾ƿ´�.
	db3_rsget.close	
	
	for i = 0 to ftotalcount -1			'�ʵ����� ���� ���鼭 ����ȴ�.
		sql = "insert into [db_datamart].[dbo].tbl_month_wonga"
		sql = sql & " (yyyymm,groupname,category,field,field_value,count) values"
		sql = sql & "('" & yyyymm & "'"				& VbCrlf
		sql = sql & ",'" & gubunbox &"'"
		sql = sql & ",'" & t &"'"			& VbCrlf
		sql = sql & ",'" & idx_f & "'"					& VbCrlf			
		sql = sql & "," & value_box_0_re(i) & ""					& VbCrlf
		sql = sql & "," & countvalue & ")"
		'response.write sql&"�������̺�����"&t &"<br>"
		db3_dbget.execute sql
		sql = ""
		sql = "update [db_datamart].[dbo].tbl_month_wonga_category set"
		sql = sql & " gijun_value= '" & gijun_box_0_re(i) & "'"				& VbCrlf
		sql = sql & " where 1=1 and groupname='" & gubunbox & "' and field='" & idx_f &"' and category='" & t &"'"
		'response.write sql&"ī�װ����̺�����"&t &"<br>"
		db3_dbget.execute sql
	idx_f = idx_f + 1
	next
	idx_f = 0
	sql1 = ""
next
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

<script language="javascript">
opener.location.reload();
self.close();
</script>