<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 카테고리
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim mode
dim catecode, catename
dim shopid,itemid,itemname
	mode = requestCheckVar(request("mode"),32)
	catecode = requestCheckVar(request("catecode"),3)
	catename = requestCheckVar(html2db(request("catename")),64)
	shopid = requestCheckVar(request("shopid"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(html2db(request("itemname")),124)

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, AlreadyExists
if mode="inputcate" then
	AlreadyExists = false
	sqlStr = " select top 1 * from [db_shop].[dbo].tbl_cafe_category"
	sqlStr = sqlStr + " where catecode='" + catecode + "'"
	rsget.Open sqlStr,dbget,1
	if Not rsget.EOF then
		AlreadyExists = true
	end if
	rsget.Close

	if Not AlreadyExists then
		sqlStr = " insert into [db_shop].[dbo].tbl_cafe_category"
		sqlStr = sqlStr + " (catecode,catename)"
		sqlStr = sqlStr + " values('" + catecode + "'"
		sqlStr = sqlStr + " ,'" + catename + "'"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1
	else
		response.write "<script language=javascript>"
		response.write "alert('이미 존재하는 코드입니다.-" + catecode + "');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
elseif mode="delcate" then
	sqlStr = " delete from [db_shop].[dbo].tbl_cafe_category"
	sqlStr = sqlStr + " where catecode='" + catecode + "'"
	rsget.Open sqlStr,dbget,1

	sqlStr = " delete from  [db_shop].[dbo].tbl_cafe_category_link"
	sqlStr = sqlStr + " where catecode='" + catecode + "'"
	rsget.Open sqlStr,dbget,1

elseif mode="linkitem" then
	AlreadyExists = false
	sqlStr = " select top 1 * from [db_shop].[dbo].tbl_cafe_category_link"
	sqlStr = sqlStr + " where shopid='" + shopid + "'"
	sqlStr = sqlStr + " and itemid='" + itemid + "'"
	sqlStr = sqlStr + " and itemname='" + itemname + "'"

	rsget.Open sqlStr,dbget,1
	if Not rsget.EOF then
		AlreadyExists = true
	end if
	rsget.Close

	if Not AlreadyExists then
		sqlStr = " insert into [db_shop].[dbo].tbl_cafe_category_link"
		sqlStr = sqlStr + " (shopid,itemid,itemname,catecode,catename)"
		sqlStr = sqlStr + " values("
		sqlStr = sqlStr + " '" + shopid + "'"
		sqlStr = sqlStr + " ," + itemid + ""
		sqlStr = sqlStr + " ,'" + itemname + "'"
		sqlStr = sqlStr + " ,'" + catecode + "'"
		sqlStr = sqlStr + " ,'" + catename + "'"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1
	else
		sqlStr = " update [db_shop].[dbo].tbl_cafe_category_link"
		sqlStr = sqlStr + " set catecode='" + catecode + "'"
		sqlStr = sqlStr + " , catename='" + catename + "'"
		sqlStr = sqlStr + " where shopid='" + shopid + "'"
		sqlStr = sqlStr + " and itemid=" + itemid + ""
		sqlStr = sqlStr + " and itemname='" + itemname + "'"

		rsget.Open sqlStr,dbget,1
	end if
end if
%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->