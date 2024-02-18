<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월간원가보고서 수정 저장 페이지
' History : 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<% 
dim groupname,category_box_0,field_box_0,gijun_box_0,value_box_0,field,category,yyyymm,chulgocount
	yyyymm = request("yyyymm")					'날짜
	groupname = request("groupname")			'그룹명
	category_box_0 = request("category_box_0")	'카테고리명
	field_box_0 = request("field_box_0")		'필드명
	gijun_box_0 = request("gijun_box_0")		'기준값
	value_box_0 = request("value_box_0")		'값
	field = request("field")					'필드구분값
	category = request("category")				'카테고리구분값
	chulgocount = request("chulgocount")		'계산값
	
dim sql
	sql = "update [db_datamart].[dbo].tbl_month_wonga_category set"
	sql = sql & " gijun_value= '" & gijun_box_0 & "', field_name = '" & field_box_0 & "'"				& VbCrlf
	sql = sql & " where 1=1 and groupname='" & groupname & "' and field='" & field &"' and category='" & category &"'"
	'response.write sql&"카테고리테이블저장<br>"
	db3_dbget.execute sql

sql = ""
	sql = "update [db_datamart].[dbo].tbl_month_wonga_category set"
	sql = sql & " category_name = '" & category_box_0 & "'"				& VbCrlf
	sql = sql & " where 1=1 and groupname='" & groupname & "' and category='" & category &"'"
	'response.write sql&"카테고리테이블저장<br>"
	db3_dbget.execute sql

sql = ""				'카테고리입력후 데이터가 원가 테이블에 저장되는지 확인한다.
	sql = "select field from [db_datamart].[dbo].tbl_month_wonga"
	sql = sql & " where 1=1 and groupname='" & groupname & "' and field='" & field &"' and category='" & category &"'"
	db3_rsget.open sql,db3_dbget,1
	
if not db3_rsget.eof then				'레코드가 있다면 받아온 값을 업데이트한다.
	sql = ""
		sql = "update [db_datamart].[dbo].tbl_month_wonga set"
		sql = sql & " field_value= '" & value_box_0 & "'"				& VbCrlf
		sql = sql & " where 1=1 and groupname='" & groupname & "' and field='" & field &"' and category='" & category &"'"
		'response.write sql&"원가테이블저장<br>"
		db3_dbget.execute sql
db3_rsget.close

else								'레코드값이 없다면 새로 저장한다.
	sql = ""	  	
	sql = "insert into [db_datamart].[dbo].tbl_month_wonga"
	sql = sql & " (yyyymm,groupname,category,field,field_value,count) values"
	sql = sql & " ('" & yyyymm & "'"				& VbCrlf
	sql = sql & ",'" & groupname &"'"
	sql = sql & ",'" & category &"'"			& VbCrlf
	sql = sql & ",'" & field & "'"					& VbCrlf			
	sql = sql & ",'" & value_box_0 & "'"					& VbCrlf
	sql = sql & ",'" & chulgocount & "')"
	'response.write sql&"원가테이블저장<br>"
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