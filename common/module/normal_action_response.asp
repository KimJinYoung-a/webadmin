<%@ language=vbscript %><% option explicit %><?xml version="1.0"  encoding="euc-kr"?>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% Response.contentType = "text/xml; charset=euc-kr" %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<response>
<%
dim mode
dim param1, param2, param3

mode = requestCheckVar(request("mode"),32)
param1 = requestCheckVar(request("param1"),10)
param2 = requestCheckVar(request("param2"),10)
param3 = requestCheckVar(request("param3"),10)

dim sqlStr

if mode="cdl" then
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_large"
	if (param1="999") then ''2016/03/23 분기 eastone 
	    sqlStr = sqlStr + " where 1=1"
	else
	    sqlStr = sqlStr + " where code_large<'999'"
    end if
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by orderNO, code_large"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsget("code_large") + "</value1>" + VbCrlf
		response.write "<value2><![CDATA[" + db2html(rsget("code_nm")) + "]]></value2>" + VbCrlf
		response.write "</item>" + VbCrlf
		rsget.moveNext
	loop
	rsget.close

elseif mode="cdm" then
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_mid"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by orderNO, code_mid"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsget("code_mid") + "</value1>" + VbCrlf
		response.write "<value2><![CDATA[" + db2html(rsget("code_nm")) + "]]></value2>" + VbCrlf
		response.write "</item>" + VbCrlf

		rsget.moveNext
	loop
	rsget.close

elseif mode="cds" then
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_small"
	sqlStr = sqlStr + " where code_large='" + param1 + "'"
	sqlStr = sqlStr + " and code_mid='" + param2 + "'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by orderNO, code_small"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		response.write "<item>" + VbCrlf
		response.write "<value1>" + rsget("code_small") + "</value1>" + VbCrlf
		response.write "<value2><![CDATA[" + db2html(rsget("code_nm")) + "]]></value2>" + VbCrlf
		response.write "</item>" + VbCrlf

		rsget.moveNext
	loop
	rsget.close

elseif mode="cdselect" then


end if
%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->