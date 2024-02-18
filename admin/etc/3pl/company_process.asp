<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim mode
dim companyid, companygubun, companyname, useyn
dim sqlStr

mode        	= requestCheckVar(request("mode"),32)
companyid   	= requestCheckVar(request("companyid"),32)
companygubun	= requestCheckVar(request("companygubun"),32)
companyname 	= html2db(requestCheckVar(request("companyname"),32))
useyn       	= requestCheckVar(request("useyn"),32)

select case mode
	case "modi"
		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_company] "
		sqlStr = sqlStr & " set updt = getdate(), "
		sqlStr = sqlStr & " company_name = '" & companyname & "', "
		sqlStr = sqlStr & " companygubun = IsNull(companygubun, '" & companygubun & "'), "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where companyid = '" & companyid & "'"
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "ins"
		sqlStr = ""
		sqlStr = sqlStr & " insert into [db_threepl].[dbo].[tbl_company](companyid, companygubun, company_name, useyn) "
		sqlStr = sqlStr & " values('" & companyid & "', '" & companygubun & "', '" & companyname & "', '" & useyn & "')"
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('저장 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case else
		response.write "에러"
end select

%>
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
