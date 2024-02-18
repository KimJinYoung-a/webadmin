<%@ language=vbscript %>
<% option explicit %>
<%
''Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, yyyymm, yyyymm2
dim itemgubun, itemid, itemoption
dim stockPlace

mode = request("mode")
yyyymm = request("yyyymm")
stockPlace = request("stockPlace")

itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")

dim sqlStr, resultrows
if mode="monthlystockipgo" then

    sqlStr = " EXEC [db_summary].[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '" & yyyymm & "', '" + CStr(stockPlace) + "' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('작성 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
elseif (mode = "monthlystockavgipgoprice") then

    sqlStr = " EXEC [db_summary].[dbo].[sp_Ten_monthlyLogisstock_avgipgoPrice] '" & yyyymm & "', '" + CStr(stockPlace) + "' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('작성 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
elseif (mode = "setmwdiv2m") then

	'// 물류-직영점 이동이고, 물류-매장 매입구분이 모두 매입이면 미지정내역 매입설정
	yyyymm2 = Left(DateAdd("m", 1, yyyymm & "-01"), 7)
	sqlStr = " update d "
	sqlStr = sqlStr + " set d.mwgubun = 'M' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].[tbl_acount_storage_master] m "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_acount_storage_detail] d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.code = d.mastercode "
	sqlStr = sqlStr + " 	join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] s "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and s.yyyymm = convert(varchar(7), m.executedt, 121) "
	sqlStr = sqlStr + " 		and s.itemgubun = d.iitemgubun "
	sqlStr = sqlStr + " 		and s.itemid = d.itemid "
	sqlStr = sqlStr + " 		and s.itemoption = d.itemoption "
	sqlStr = sqlStr + " 		and s.lastmwdiv = 'M' "
	sqlStr = sqlStr + " 	join [db_summary].[dbo].[tbl_monthly_accumulated_shopstock_summary] ss "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and ss.shopid = m.socid "
	sqlStr = sqlStr + " 		and ss.yyyymm = convert(varchar(7), m.executedt, 121) "
	sqlStr = sqlStr + " 		and ss.itemgubun = d.iitemgubun "
	sqlStr = sqlStr + " 		and ss.itemid = d.itemid "
	sqlStr = sqlStr + " 		and ss.itemoption = d.itemoption "
	sqlStr = sqlStr + " 		and ss.lstComm_cd = 'B031' "
	sqlStr = sqlStr + " 	join [db_partner].[dbo].[tbl_partner] p "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.socid = p.id "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.executedt >= '" & yyyymm & "-01' "
	sqlStr = sqlStr + " 	and m.executedt < '" & yyyymm2 & "-01' "
	sqlStr = sqlStr + " 	and d.mwgubun = '' "
	sqlStr = sqlStr + " 	and m.deldt is NULL "
	sqlStr = sqlStr + " 	and d.deldt is NULL "
	sqlStr = sqlStr + " 	and p.userdiv <> '503' "
    dbget.execute sqlStr,resultrows

	sqlStr = " update ss "
	sqlStr = sqlStr + " set ss.lstComm_cd = 'B031' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].[tbl_acount_storage_master] m "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_acount_storage_detail] d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.code = d.mastercode "
	sqlStr = sqlStr + " 	join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] s "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and s.yyyymm = convert(varchar(7), m.executedt, 121) "
	sqlStr = sqlStr + " 		and s.itemgubun = d.iitemgubun "
	sqlStr = sqlStr + " 		and s.itemid = d.itemid "
	sqlStr = sqlStr + " 		and s.itemoption = d.itemoption "
	sqlStr = sqlStr + " 	join [db_summary].[dbo].[tbl_monthly_accumulated_shopstock_summary] ss "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and ss.shopid = m.socid "
	sqlStr = sqlStr + " 		and ss.yyyymm = convert(varchar(7), m.executedt, 121) "
	sqlStr = sqlStr + " 		and ss.itemgubun = d.iitemgubun "
	sqlStr = sqlStr + " 		and ss.itemid = d.itemid "
	sqlStr = sqlStr + " 		and ss.itemoption = d.itemoption "
	sqlStr = sqlStr + " 	join [db_partner].[dbo].[tbl_partner] p "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.socid = p.id "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.executedt >= '" & yyyymm & "-01' "
	sqlStr = sqlStr + " 	and m.executedt < '" & yyyymm2 & "-01' "
	sqlStr = sqlStr + " 	and s.lastmwdiv = 'M' "
	sqlStr = sqlStr + " 	and d.mwgubun = 'M' "
	sqlStr = sqlStr + " 	and ss.lstComm_cd is NULL "
	sqlStr = sqlStr + " 	and m.deldt is NULL "
	sqlStr = sqlStr + " 	and d.deldt is NULL "
	sqlStr = sqlStr + " 	and p.userdiv <> '503' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('작성 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
else
	response.write "mode=" + mode
	dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
