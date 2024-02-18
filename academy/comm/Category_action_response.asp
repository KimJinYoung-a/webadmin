<?xml version="1.0"  encoding="euc-kr"?>
<% Response.contentType = "text/xml; charset=euc-kr" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<response>
<%
dim mode
dim param1, param2, param3

mode = RequestCheckvar(request("mode"),6)
param1 = RequestCheckvar(request("param1"),10)
param2 = RequestCheckvar(request("param2"),10)
param3 = RequestCheckvar(request("param3"),10)

dim sqlStr

if mode="cdl" then
	sqlStr = "select top 100 * from db_academy.dbo.tbl_diy_item_cate_large"
'	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by code_large"

	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	do until rsACADEMYget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsACADEMYget("code_large") + "</value1>" + VbCrlf
		response.write "<value2>" + db2html(rsACADEMYget("code_nm")) + "</value2>" + VbCrlf
		response.write "</item>" + VbCrlf
		rsACADEMYget.moveNext
	loop
	rsACADEMYget.close

elseif mode="cdm" then
	sqlStr = "select top 100 * from db_academy.dbo.tbl_diy_item_cate_mid"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
'	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by code_mid"

	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	do until rsACADEMYget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsACADEMYget("code_mid") + "</value1>" + VbCrlf
		response.write "<value2>" + db2html(rsACADEMYget("code_nm")) + "</value2>" + VbCrlf
		response.write "</item>" + VbCrlf

		rsACADEMYget.moveNext
	loop
	rsACADEMYget.close
elseif mode="cds" then
	sqlStr = "select top 100 * from db_academy.dbo.tbl_diy_item_cate_small"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and code_mid='" + param2 + "'"
'	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by code_small"

	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	do until rsACADEMYget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsACADEMYget("code_small") + "</value1>" + VbCrlf
		response.write "<value2>" + db2html(rsACADEMYget("code_nm")) + "</value2>" + VbCrlf
		response.write "</item>" + VbCrlf

		rsACADEMYget.moveNext
	loop
	rsACADEMYget.close
elseif mode="cdselect" then


end if
%>
</response>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->