<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월간원가보고서 카테고리 신규그룹 저장 페이지
' History : 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
dim category_box_0,value_box_0 , gijun_box_0 ,gubunbox ,groupname ,yyyy ,mm ,yyyymm ,countvalue 
dim category_box_0_re,value_box_0_re , gijun_box_0_re , categorycount , gijuncount,valuecount
	category_box_0 = request("category_box_0")&","		' 콤마로 구분해서 배열에 넣기위해 콤마를 붙인다.
	value_box_0 = request("value_box_0")&","		' 콤마로 구분해서 배열에 넣기위해 콤마를 붙인다.
	gijun_box_0 = request("gijun_box_0")&","
	gubunbox = request("gubun_submit")		'그룹값을 받아온다
	groupname = request("groupname")
	yyyy = request("yyyy")					'년을 받아온다.
	mm = request("mm")						'달을 받아온다.
	yyyymm = yyyy&mm						'받아온 년과 달을 합친다.
	countvalue = request("count")			'계산을 위한 카운트 값을 받아온다.	

category_box_0_re = split(category_box_0,",")		'콤마를 기준으로 배열로 짜른다.
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
idx_f = 0		'필드번호를 0부터 매기기위한 인덱스변수
%>

<%
dim sql1 ,ftotalcount

for t = 0 to categorycount - 1		'카테고리 갯수만큼 루프

	sql1 = "select"
	sql1 = sql1 & " field"
	sql1 = sql1 & " from db_datamart.dbo.tbl_month_wonga_category"
	sql1 = sql1 & " where 1=1 and groupname= '"& gubunbox &"' and category_isusing='y' and category = '"& t &"'"
	sql1 = sql1 & " group by field" 	
	
	db3_rsget.open sql1,db3_dbget,1
	'response.write sql1&"<br>"	
	ftotalcount = db3_rsget.recordcount		'각각의 카테고리안에 필드가 몇개인지 갯수를 받아온다.
	db3_rsget.close	
	
	for i = 0 to ftotalcount -1			'필드수대로 루프 돌면서 저장된다.
		sql = "insert into [db_datamart].[dbo].tbl_month_wonga"
		sql = sql & " (yyyymm,groupname,category,field,field_value,count) values"
		sql = sql & "('" & yyyymm & "'"				& VbCrlf
		sql = sql & ",'" & gubunbox &"'"
		sql = sql & ",'" & t &"'"			& VbCrlf
		sql = sql & ",'" & idx_f & "'"					& VbCrlf			
		sql = sql & "," & value_box_0_re(i) & ""					& VbCrlf
		sql = sql & "," & countvalue & ")"
		'response.write sql&"원가테이블저장"&t &"<br>"
		db3_dbget.execute sql
		sql = ""
		sql = "update [db_datamart].[dbo].tbl_month_wonga_category set"
		sql = sql & " gijun_value= '" & gijun_box_0_re(i) & "'"				& VbCrlf
		sql = sql & " where 1=1 and groupname='" & gubunbox & "' and field='" & idx_f &"' and category='" & t &"'"
		'response.write sql&"카테고리테이블저장"&t &"<br>"
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