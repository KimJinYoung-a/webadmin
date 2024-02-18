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
dim companyid, prdcode, brandid
dim prdname, prdoptionname, itemgubun, itemid, itemoption, itemoptionname, customerprice, generalbarcode, useyn
dim sqlStr

mode        	= requestCheckVar(request("mode"),32)
companyid   	= requestCheckVar(request("companyid"),32)
prdcode   		= requestCheckVar(request("prdcode"),32)
brandid   		= requestCheckVar(request("brandid"),32)
prdname   		= html2db(requestCheckVar(request("prdname"),32))
prdoptionname   = html2db(requestCheckVar(request("prdoptionname"),32))
itemgubun  		= requestCheckVar(request("itemgubun"),32)
itemid   		= requestCheckVar(request("itemid"),32)
itemoption   	= requestCheckVar(request("itemoption"),32)
itemoptionname  = html2db(requestCheckVar(request("itemoptionname"),32))
customerprice   = requestCheckVar(request("customerprice"),32)
generalbarcode  = requestCheckVar(request("generalbarcode"),32)
useyn   		= requestCheckVar(request("useyn"),32)

select case mode
	case "modi"
		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_item] "
		sqlStr = sqlStr & " set updt = getdate(), "
		sqlStr = sqlStr & " brandid = '" & brandid & "', "
		sqlStr = sqlStr & " prdname = '" & prdname & "', "
		sqlStr = sqlStr & " prdoptionname = '" & prdoptionname & "', "
		sqlStr = sqlStr & " itemgubun = '" & itemgubun & "', "
		sqlStr = sqlStr & " itemid = '" & itemid & "', "
		sqlStr = sqlStr & " itemoption = '" & itemoption & "', "
		sqlStr = sqlStr & " itemoptionname = '" & itemoptionname & "', "
		sqlStr = sqlStr & " customerprice = '" & customerprice & "', "
		sqlStr = sqlStr & " generalbarcode = '" & generalbarcode & "', "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where companyid = '" & companyid & "' and prdcode = '" & prdcode & "' "
		''rw sqlStr
		dbget_TPL.Execute sqlStr

		response.write "<script>alert('수정 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End
	case "ins"
		prdcode = CreatePrdcodeOne(companyid, itemgubun, itemid, itemoption, itemoptionname)
		if (prdcode = "") then
			response.write "<script>alert('상품등록에 실패했습니다.');</script>"
			response.write "상품등록에 실패했습니다."
			dbget.close()	:	response.End
		end if

		sqlStr = ""
		sqlStr = sqlStr & " update [db_threepl].[dbo].[tbl_item] "
		sqlStr = sqlStr & " set updt = getdate(), "
		sqlStr = sqlStr & " brandid = '" & brandid & "', "
		sqlStr = sqlStr & " prdname = '" & prdname & "', "
		sqlStr = sqlStr & " prdoptionname = '" & prdoptionname & "', "
		sqlStr = sqlStr & " itemgubun = '" & itemgubun & "', "
		sqlStr = sqlStr & " itemid = '" & itemid & "', "
		sqlStr = sqlStr & " itemoption = '" & itemoption & "', "
		sqlStr = sqlStr & " itemoptionname = '" & itemoptionname & "', "
		sqlStr = sqlStr & " customerprice = '" & customerprice & "', "
		sqlStr = sqlStr & " generalbarcode = '" & generalbarcode & "', "
		sqlStr = sqlStr & " useyn = '" & useyn & "' "
		sqlStr = sqlStr & " where companyid = '" & companyid & "' and prdcode = '" & prdcode & "' "
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
