<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% response.Charset="euc-kr" %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
'' Non XML. xmlDom Object에서 오류?.
dim mode
dim param1, param2, param3

mode = requestCheckVar(request("mode"),32)
param1 = requestCheckVar(request("param1"),10)
param2 = requestCheckVar(request("param2"),10)
param3 = requestCheckVar(request("param3"),10)

dim sqlStr

if mode="cdl" then
	sqlStr = "select top 100 * from [db_item].dbo.tbl_cate_large"
	sqlStr = sqlStr + " where display_yn='Y'"
	sqlStr = sqlStr + " order by code_large"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write rsget("code_large") & "|C|C|" & db2html(rsget("code_nm")) & "|R|R|"
		rsget.moveNext
	loop
	rsget.close

elseif mode="cdm" then
	sqlStr = "select top 100 * from [db_item].dbo.tbl_cate_mid"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by code_mid"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write rsget("code_mid") & "|C|C|" & db2html(rsget("code_nm")) & "|R|R|"
		rsget.moveNext
	loop
	rsget.close
elseif mode="cds" then
	sqlStr = "select top 100 * from [db_item].dbo.tbl_cate_small"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and code_mid='" + param2 + "'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by code_small"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write  rsget("code_small") & "|C|C|" & db2html(rsget("code_nm")) & "|R|R|"
		rsget.moveNext
	loop
	rsget.close
elseif mode="cdselect" then


end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->