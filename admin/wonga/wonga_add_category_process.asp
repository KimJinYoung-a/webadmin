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
dim category_box_0 , field_box_0 ,gubunbox ,groupname ,yyyy ,mm ,yyyymm ,countvalue  , gijun_box_0,gijuncount
dim category_box_0_re , field_box_0_re ,flag_field_re, categorycount , fieldcount ,  gijun_box_0_re
dim add_category
	category_box_0 = request("category_box_0")&","		' 콤마로 구분해서 배열에 넣기위해 콤마를 붙인다.
	field_box_0 = request("field_box_0")&","
	gubunbox = request("gubun_submit")		'그룹값을 받아온다
	groupname = request("groupname")
	yyyy = request("yyyy")					'년을 받아온다.
	mm = request("mm")						'달을 받아온다.
	yyyymm = yyyy&mm						'받아온 년과 달을 합친다.
	add_category = request("add_category")	' 이미존재하는그룹에 카테고리 추가를 위해서 이전 카테고리를받아옴...
	countvalue = request("count")			'계산을 위한 카운트 값을 받아온다.	
	gijun_box_0 = request("gijun_box_0")&","	'기준값을 받아온다.

category_box_0_re = split(category_box_0,",")	'콤마를 기준으로 베열로 짜른다.
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
idx_f = 0		'필드번호를 0부터 매기기위한 인덱스변수
%>

<% if gubunbox = "" then 
'##################################################################	그룹이 없을때 카테고리저장시작
	
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
			'response.write sql&"카테고리"&t &"<br>"
			db3_dbget.execute sql
		idx_f = idx_f + 1
		next
		idx_f = 0
%>

<script language="javascript">
alert('처리 되었습니다. 등록하신 카테고리에 데이터를넣으셔야 정상 처리됩니다.');
location.replace('/admin/wonga/wonga_add_category.asp?gubunbox='+'<%= groupname %>');
</script>	

<% else 
'##################################################################	그룹이 있을때 카테고리저장시작

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
			'response.write sql&"카테고리"&t &"<br>"
			db3_dbget.execute sql
		idx_f = idx_f + 1
		next
		idx_f = 0
%>
<script language="javascript">
alert('처리 되었습니다. 등록하신 카테고리에 데이터를넣으셔야 정상 처리됩니다.');
location.replace('/admin/wonga/wonga_add_category.asp?gubunbox='+'<%= gubunbox %>');
</script>
<% end if %>


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->