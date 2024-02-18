<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<%

dim mode
dim companyid, brandid, brandname, brandnameeng, useyn, companyBrandId
dim sqlStr

mode        	= requestCheckVar(request("mode"),32)
companyid   	= requestCheckVar(request("companyid"),32)
brandid   		= requestCheckVar(request("brandid"),32)
brandname 		= html2db(requestCheckVar(request("brandname"),32))
brandnameeng	= html2db(requestCheckVar(request("brandnameeng"),32))
companyBrandId  = requestCheckVar(request("companyBrandId"),32)
useyn       	= requestCheckVar(request("useyn"),32)

select case mode
	case "modi"
		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_brand] "
		sqlStr = sqlStr & " set lastupdate = getdate(), "
		sqlStr = sqlStr & " brand_name = '" & brandname & "', "
		sqlStr = sqlStr & " brand_name_eng = '" & brandnameeng & "', "
		sqlStr = sqlStr & " companyBrandId = '" & companyBrandId & "', "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where companyid = '" & companyid & "' and brandid = '" & brandid & "' "
		''rw sqlStr
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "ins"
		brandid = CreateBrandOne(companyid)
		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_brand] "
		sqlStr = sqlStr & " set lastupdate = getdate(), "
		sqlStr = sqlStr & " brand_name = '" & brandname & "', "
		sqlStr = sqlStr & " brand_name_eng = '" & brandnameeng & "', "
		sqlStr = sqlStr & " companyBrandId = '" & companyBrandId & "', "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where companyid = '" & companyid & "' and brandid = '" & brandid & "' "
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
