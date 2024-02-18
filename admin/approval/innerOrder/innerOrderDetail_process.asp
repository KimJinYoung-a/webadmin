<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%


dim mode
dim idx, detailidx, innerorderpercentage, dealdiv
dim appDate, divcd, SELLBIZSECTION_CD, BUYBIZSECTION_CD

mode = requestCheckvar(Request("mode"),32)

idx 					= requestCheckvar(Request("idx"),32)
detailidx 				= requestCheckvar(Request("detailidx"),32)
innerorderpercentage 	= requestCheckvar(Request("innerorderpercentage"),32)
dealdiv 				= requestCheckvar(Request("dealdiv"),32)

appDate 				= requestCheckvar(Request("appDate"),7)
divcd 					= requestCheckvar(Request("divcd"),32)
SELLBIZSECTION_CD 		= requestCheckvar(Request("SELLBIZSECTION_CD"),32)
BUYBIZSECTION_CD 		= requestCheckvar(Request("BUYBIZSECTION_CD"),32)

dim sqlStr

if (mode = "modifyinnerorderpercentage") then

	if (dealdiv = "") then
		response.write "작업중입니다."
		response.end
	end if

	sqlStr = " update d "
	sqlStr = sqlStr + " set d.totalSum = round((d.totalsellcash * " + CStr(innerorderpercentage) + " / 100), 0) "
	sqlStr = sqlStr + " 	, d.supplySum = Round(((d.totalsellcash * " + CStr(innerorderpercentage) + " / 100) * 10.0 / 11.0), 0) "
	sqlStr = sqlStr + " 	, d.taxSum = round((d.totalsellcash * " + CStr(innerorderpercentage) + " / 100), 0) - Round(((d.totalsellcash * " + CStr(innerorderpercentage) + " / 100) * 10.0 / 11.0), 0) "
	sqlStr = sqlStr + " 	, d.innerorderpercentage = " + CStr(innerorderpercentage) + " "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_InternalOrder m "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_InternalOrderDetail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " where d.idx = " & detailidx & " and d.masteridx = " & idx & " and m.VATYN = 'Y' "
	dbget.Execute sqlStr

	sqlStr = " update d "
	sqlStr = sqlStr + " set d.totalSum = round((d.totalsellcash * " + CStr(innerorderpercentage) + " / 100), 0) "
	sqlStr = sqlStr + " 	, d.supplySum = round((d.totalsellcash * " + CStr(innerorderpercentage) + " / 100), 0) "
	sqlStr = sqlStr + " 	, d.taxSum = 0 "
	sqlStr = sqlStr + " 	, d.innerorderpercentage = " + CStr(innerorderpercentage) + " "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_InternalOrder m "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_InternalOrderDetail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " where d.idx = " & detailidx & " and d.masteridx = " & idx & " and m.VATYN = 'N' "
	dbget.Execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr + " set m.totalSum = T.totalSum, m.supplySum = T.supplySum, m.taxSum = T.taxSum "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_InternalOrder m "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select "
	sqlStr = sqlStr + " 			masteridx, SUM(totalSum) as totalSum, SUM(supplySum) as supplySum, SUM(taxSum) as taxSum "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 			db_partner.dbo.tbl_InternalOrderDetail "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			masteridx = " & idx & " "
	sqlStr = sqlStr + " 		group by "
	sqlStr = sqlStr + " 			masteridx "
	sqlStr = sqlStr + " 	) T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = T.masteridx "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	m.idx = " & idx & " "
	dbget.Execute sqlStr

	'response.write sqlStr
	''dbget.Execute sqlStr

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popViewInnerOrderDetailNew.asp?idx=" + CStr(idx) + "'</script>"
    dbget.close()	:	response.End

elseif (mode = "updateOneDetail") then

	sqlStr = " exec [db_partner].[dbo].[usp_Ten_InnerOrder_UpdateOne] '" + CStr(appDate) + "', '" + CStr(divcd) + "', '" + CStr(SELLBIZSECTION_CD) + "', '" + CStr(BUYBIZSECTION_CD) + "', '" + CStr(session("ssBctId")) + "' "

	if (divcd = "101") or (divcd = "102") or (divcd = "201") or (divcd = "202") or (divcd = "301") or (divcd = "302") or (divcd = "303") or (divcd = "304") or (divcd = "305") or (divcd = "307") or (divcd = "501") or (divcd = "502") then
		dbget.Execute sqlStr

	    response.write "<script>alert('수정 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();</script>"
	    response.write "<script>opener.focus(); window.close();</script>"
	    dbget.close()	:	response.End
	else
		response.write sqlStr
	end if

else
	'// 에러
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
