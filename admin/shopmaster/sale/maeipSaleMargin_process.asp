<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim mode
dim strSql, i, rs
dim idx, makerid, saleCode, startDate, endDate, meachulGubun, defaultMargin, saleMargin, reguserid, useyn, regdate, lastupdate


function UpdateItemList(idx)
	dim strSql

	strSql = " update d "
	strSql = strSql & " set d.saleprice = T.saleprice, d.orgBuyCash = T.orgsuplycash, d.saleBuyCash = T.saleBuyCash, d.useyn = 'Y', d.mwdiv = T.mwdiv, lastupdate = getdate() "
	strSql = strSql & " from "
	strSql = strSql & " 	( "
	strSql = strSql & " 		select "
	strSql = strSql & " 			m.idx as masteridx, i.itemid, i.itemname, i.mwdiv, i.orgprice, s.saleprice, i.orgsuplycash, round(s.saleprice*(100.0 - m.saleMargin)/100.0,0) as saleBuyCash "
	strSql = strSql & " 		from "
	strSql = strSql & " 			[db_order].[dbo].[tbl_meaipSaleMarginShare_master] m "
	strSql = strSql & " 			join [db_event].[dbo].[tbl_saleItem] s "
	strSql = strSql & " 			on "
	strSql = strSql & " 				m.saleCode = s.sale_Code "
	strSql = strSql & " 			join [db_item].[dbo].[tbl_item] i "
	strSql = strSql & " 			on "
	strSql = strSql & " 				s.itemid = i.itemid "
	strSql = strSql & " 		where "
	strSql = strSql & " 			1 = 1 "
	strSql = strSql & " 			and i.mwdiv='M' "
	strSql = strSql & " 			and s.salesupplycash = i.orgsuplycash "
	strSql = strSql & " 			and round(100.0 - 100.0*i.orgsuplycash/i.orgprice,1) = round(m.defaultMargin,1) "
	strSql = strSql & " 			and m.makerid = i.makerid "
	strSql = strSql & " 			and m.idx = " & idx
	strSql = strSql & " 	) T "
	strSql = strSql & " 	join [db_order].[dbo].[tbl_meaipSaleMarginShare_detail] d "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and d.masteridx = T.masteridx "
	strSql = strSql & " 		and d.itemid = T.itemid "
	strSql = strSql & " where "
	strSql = strSql & " 	(d.saleprice <> T.saleprice or d.orgBuyCash <> T.orgsuplycash or d.saleBuyCash <> T.saleBuyCash or d.useyn = 'N' or IsNull(d.mwdiv, '') <> IsNull(T.mwdiv, '')) "
	''rw strSql
	''response.end
	dbget.execute strSql

	strSql = " update [db_order].[dbo].[tbl_meaipSaleMarginShare_detail] "
	strSql = strSql & " set useyn = 'N' "
	strSql = strSql & " where "
	strSql = strSql & " 	1 = 1 "
	strSql = strSql & " 	and itemid not in ( "
	strSql = strSql & " 		select "
	strSql = strSql & " 			i.itemid "
	strSql = strSql & " 		from "
	strSql = strSql & " 			[db_order].[dbo].[tbl_meaipSaleMarginShare_master] m "
	strSql = strSql & " 			join [db_event].[dbo].[tbl_saleItem] s "
	strSql = strSql & " 			on "
	strSql = strSql & " 				m.saleCode = s.sale_Code "
	strSql = strSql & " 			join [db_item].[dbo].[tbl_item] i "
	strSql = strSql & " 			on "
	strSql = strSql & " 				s.itemid = i.itemid "
	strSql = strSql & " 		where "
	strSql = strSql & " 			1 = 1 "
	strSql = strSql & " 			and i.mwdiv='M' "
	strSql = strSql & " 			and s.salesupplycash = i.orgsuplycash "
	strSql = strSql & " 			and round(100.0 - 100.0*i.orgsuplycash/i.orgprice,1) = round(m.defaultMargin,1) "
	strSql = strSql & " 			and m.makerid = i.makerid "
	strSql = strSql & " 			and m.idx = " & idx
	strSql = strSql & " 	) "
	strSql = strSql & " 	and masteridx = " & idx
	''response.write strSql
	dbget.execute strSql

	strSql = " insert into [db_order].[dbo].[tbl_meaipSaleMarginShare_detail](masteridx, itemid, itemname, orgprice, saleprice, orgBuyCash, saleBuyCash, useyn, regdate, lastupdate) "
	strSql = strSql & " select T.masteridx, T.itemid, T.itemname, T.orgprice, T.saleprice, T.orgsuplycash, T.saleBuyCash, 'Y', getdate(), getdate() "
	strSql = strSql & " from "
	strSql = strSql & " 	( "
	strSql = strSql & " 		select "
	strSql = strSql & " 			m.idx as masteridx, i.itemid, i.itemname, i.mwdiv, i.orgprice, s.saleprice, i.orgsuplycash, round(s.saleprice*(100.0 - m.saleMargin)/100.0,0) as saleBuyCash "
	strSql = strSql & " 		from "
	strSql = strSql & " 			[db_order].[dbo].[tbl_meaipSaleMarginShare_master] m "
	strSql = strSql & " 			join [db_event].[dbo].[tbl_saleItem] s "
	strSql = strSql & " 			on "
	strSql = strSql & " 				m.saleCode = s.sale_Code "
	strSql = strSql & " 			join [db_item].[dbo].[tbl_item] i "
	strSql = strSql & " 			on "
	strSql = strSql & " 				s.itemid = i.itemid "
	strSql = strSql & " 		where "
	strSql = strSql & " 			1 = 1 "
	strSql = strSql & " 			and i.mwdiv='M' "
	strSql = strSql & " 			and s.salesupplycash = i.orgsuplycash "
	strSql = strSql & " 			and round(100.0 - 100.0*i.orgsuplycash/i.orgprice,1) = round(m.defaultMargin,1) "
	strSql = strSql & " 			and m.makerid = i.makerid "
	strSql = strSql & " 			and m.idx = " & idx
	strSql = strSql & " 	) T "
	strSql = strSql & " 	left join [db_order].[dbo].[tbl_meaipSaleMarginShare_detail] d "
	strSql = strSql & " 	on "
	strSql = strSql & " 		1 = 1 "
	strSql = strSql & " 		and d.masteridx = T.masteridx "
	strSql = strSql & " 		and d.itemid = T.itemid "
	strSql = strSql & " where "
	strSql = strSql & " 	d.idx is NULL "
	strSql = strSql & " order by "
	strSql = strSql & " 	T.itemid "
	''response.write strSql
	dbget.execute strSql
end function

mode     		= requestCheckVar(Request("mode"), 32)
idx     		= requestCheckVar(Request("idx"), 32)
makerid     	= requestCheckVar(Request("makerid"), 32)
saleCode     	= getNumeric(requestCheckVar(Request("saleCode"), 32))
startDate     	= requestCheckVar(Request("startDate"), 32)
endDate     	= requestCheckVar(Request("endDate"), 32)
meachulGubun    = requestCheckVar(Request("meachulGubun"), 32)
defaultMargin   = getNumeric(requestCheckVar(Request("defaultMargin"), 32))
saleMargin     	= getNumeric(requestCheckVar(Request("saleMargin"), 32))
useyn     		= requestCheckVar(Request("useyn"), 32)
reguserid     	= session("ssBctId")


Select Case mode
	Case "modi"
		strSql = " update [db_order].[dbo].[tbl_meaipSaleMarginShare_master] "
		strSql = strSql & " set "
		strSql = strSql & " 	makerid = '" & makerid & "', "
		strSql = strSql & " 	saleCode = '" & saleCode & "', "
		strSql = strSql & " 	startDate = '" & startDate & "', "
		strSql = strSql & " 	endDate = '" & endDate & "', "
		strSql = strSql & " 	meachulGubun = '" & meachulGubun & "', "
		strSql = strSql & " 	defaultMargin = '" & defaultMargin & "', "
		strSql = strSql & " 	saleMargin = '" & saleMargin & "', "
		strSql = strSql & " 	reguserid = '" & reguserid & "', "
		strSql = strSql & " 	useyn = '" & useyn & "', "
		strSql = strSql & " 	lastupdate = getdate() "
		strSql = strSql & " where "
		strSql = strSql & " 	idx = " & idx
		''rw strSql
		dbget.execute strSql

		Call UpdateItemList(idx)

		Alert_move "저장되었습니다.","maeipSaleMarginModi.asp?menupos=" & menupos & "&idx=" & idx
	Case "ins"
		strSql = " SET NOCOUNT ON; insert into [db_order].[dbo].[tbl_meaipSaleMarginShare_master](makerid, saleCode, startDate, endDate, meachulGubun, defaultMargin, saleMargin, reguserid, useyn, regdate, lastupdate)"
		strSql = strSql & " values('" & makerid & "', '" & saleCode & "', '" & startDate & "', '" & endDate & "', '" & meachulGubun & "', " & defaultMargin & ", " & saleMargin & ", '" & reguserid & "', '" & useyn & "', getdate(), getdate()); "
		strSql = strSql & " SELECT NEWID = SCOPE_IDENTITY(); SET NOCOUNT OFF; "
		''rw strSql
		set rs = dbget.execute(strSql)
		Call UpdateItemList(rs(0))
		rs.close: set rs = nothing

		Alert_move "저장되었습니다.","maeipSaleMarginList.asp?menupos="&menupos
	Case else
		response.write("Super Saturday!!!!")
End Select

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
