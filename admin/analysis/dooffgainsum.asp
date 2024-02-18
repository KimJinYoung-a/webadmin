<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, yyyymm
mode = request("mode")
yyyymm = request("yyyymm")

dim yyyymmdd1, yyyymmdd2
yyyymmdd1 = yyyymm + "-01"
yyyymmdd2 = CStr(DateSerial(Left(yyyymm,4),Right(yyyymm,2)+1,1))

dim sqlStr

if mode="MakeMonthlyBrandSellSum" then
	''''오프샾 샾별 브랜드 월별매출
	''sqlStr = "insert into [db_summary].[dbo].tbl_shop_brand_monthly_sellsum "
	''sqlStr = sqlStr + " (yyyymm,shopid,makerid,totalitemcount,orgsellsum,realsellsum) "
	''sqlStr = sqlStr + "  select convert(varchar(7),m.shopregdate,21) as yyyymm, m.shopid, d.makerid, "
	''sqlStr = sqlStr + "  count(itemno) as totno, sum(d.sellprice*itemno) as orgsellsum, "
	''sqlStr = sqlStr + "  sum(d.realsellprice*itemno) as realsellsum"
	''sqlStr = sqlStr + "  from "
	''sqlStr = sqlStr + "  [db_shop].[dbo].tbl_shopjumun_master m,"
	''sqlStr = sqlStr + "  [db_shop].[dbo].tbl_shopjumun_detail d"
	''sqlStr = sqlStr + "   where m.idx=d.masteridx"
	''sqlStr = sqlStr + "   and m.shopregdate>='" + yyyymmdd1 + "'"
	''sqlStr = sqlStr + "   and m.shopregdate<'" + yyyymmdd2 + "'"
	''sqlStr = sqlStr + "   and Left(m.shopid,10)='streetshop'"
	''sqlStr = sqlStr + "   and m.cancelyn='N'"
	''sqlStr = sqlStr + "   and d.cancelyn='N'"
	''sqlStr = sqlStr + "   group by  convert(varchar(7),m.shopregdate,21), m.shopid, d.makerid "
	''sqlStr = sqlStr + "   order by yyyymm, m.shopid, d.makerid"
	''
	''rsget.Open sqlStr, dbget, 1
	''
	''''오프샾 샾별 브랜드 월별출고
	''sqlStr = "insert into [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum"
	''sqlStr = sqlStr + " (yyyymm,shopid,makerid,totalitemcount,sellcashsum,upchebuysum,shopsuplysum)"
	''sqlStr = sqlStr + " select convert(varchar(7),m.executedt,21) as yyyymm, m.socid,"
	''sqlStr = sqlStr + " d.imakerid,"
	''sqlStr = sqlStr + " sum(itemno*-1),"
	''sqlStr = sqlStr + " sum(d.sellcash*itemno*-1), "
	''sqlStr = sqlStr + " sum(d.buycash*itemno*-1),"
	''sqlStr = sqlStr + " sum(d.suplycash*itemno*-1)"
	''
	''sqlStr = sqlStr + "  from [db_storage].[dbo].tbl_acount_storage_master m,"
	''sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	''sqlStr = sqlStr + " where m.code=d.mastercode"
	''sqlStr = sqlStr + " and m.deldt is null"
	''sqlStr = sqlStr + " and d.deldt is null"
	''sqlStr = sqlStr + " and m.ipchulflag='S'"
	''sqlStr = sqlStr + " and m.executedt>='" + yyyymmdd1 + "'"
	''sqlStr = sqlStr + " and m.executedt<'" + yyyymmdd2 + "'"
	''sqlStr = sqlStr + " group by convert(varchar(7),m.executedt,21),m.socid, d.imakerid"
	''sqlStr = sqlStr + " order by m.yyyymm, m.socid, d.imakerid"
	''
	''rsget.Open sqlStr, dbget, 1
	''
	''''+ 출고
	''sqlStr = "update [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum"
	''sqlStr = sqlStr + " set chul_totalitemcount=T.itemno"
	''sqlStr = sqlStr + " ,chul_sellcashsum=T.sellcash"
	''sqlStr = sqlStr + " ,chul_upchebuysum=T.buycash"
	''sqlStr = sqlStr + " ,chul_shopsuplysum=T.suplycash"
	''sqlStr = sqlStr + " from ("
	''sqlStr = sqlStr + " select convert(varchar(7),m.executedt,21) as yyyymm, m.socid,"
	''sqlStr = sqlStr + " d.imakerid,"
	''sqlStr = sqlStr + " sum(itemno*-1) as itemno,"
	''sqlStr = sqlStr + " sum(d.sellcash*itemno*-1) as sellcash, "
	''sqlStr = sqlStr + " sum(d.buycash*itemno*-1) as buycash,"
	''sqlStr = sqlStr + " sum(d.suplycash*itemno*-1) as suplycash"
	''sqlStr = sqlStr + "  from [db_storage].[dbo].tbl_acount_storage_master m,"
	''sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	''sqlStr = sqlStr + " where m.code=d.mastercode"
	''sqlStr = sqlStr + " and m.deldt is null"
	''sqlStr = sqlStr + " and d.deldt is null"
	''sqlStr = sqlStr + " and m.ipchulflag='S'"
	''sqlStr = sqlStr + " and m.executedt>='" + yyyymmdd1 + "'"
	''sqlStr = sqlStr + " and m.executedt<'" + yyyymmdd2 + "'"
	''sqlStr = sqlStr + " and d.itemno<0"
	''sqlStr = sqlStr + " group by convert(varchar(7),m.executedt,21),m.socid, d.imakerid"
	''sqlStr = sqlStr + " ) T"
	''sqlStr = sqlStr + " where [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum.yyyymm=T.yyyymm"
	''sqlStr = sqlStr + " and [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum.shopid=T.socid"
	''sqlStr = sqlStr + " and [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum.makerid=T.imakerid"
	''
	''rsget.Open sqlStr, dbget, 1
	''
	''
	''// - 반품
	''sqlStr = "update [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum"
	''sqlStr = sqlStr + " set re_totalitemcount=T.itemno"
	''sqlStr = sqlStr + " ,re_sellcashsum=T.sellcash"
	''sqlStr = sqlStr + " ,re_upchebuysum=T.buycash"
	''sqlStr = sqlStr + " ,re_shopsuplysum=T.suplycash"
	''sqlStr = sqlStr + " from ("
	''sqlStr = sqlStr + " select convert(varchar(7),m.executedt,21) as yyyymm, m.socid,"
	''sqlStr = sqlStr + " d.imakerid,"
	''sqlStr = sqlStr + " sum(itemno*-1) as itemno,"
	''sqlStr = sqlStr + " sum(d.sellcash*itemno*-1) as sellcash, "
	''sqlStr = sqlStr + " sum(d.buycash*itemno*-1) as buycash,"
	''sqlStr = sqlStr + " sum(d.suplycash*itemno*-1) as suplycash"
	''sqlStr = sqlStr + "  from [db_storage].[dbo].tbl_acount_storage_master m,"
	''sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	''sqlStr = sqlStr + " where m.code=d.mastercode"
	''sqlStr = sqlStr + " and m.deldt is null"
	''sqlStr = sqlStr + " and d.deldt is null"
	''sqlStr = sqlStr + " and m.ipchulflag='S'"
	''sqlStr = sqlStr + " and m.executedt>='" + yyyymmdd1 + "'"
	''sqlStr = sqlStr + " and m.executedt<'" + yyyymmdd2 + "'"
	''sqlStr = sqlStr + " and d.itemno>0"
	''sqlStr = sqlStr + " group by convert(varchar(7),m.executedt,21),m.socid, d.imakerid"
	''sqlStr = sqlStr + " ) T"
	''sqlStr = sqlStr + " where [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum.yyyymm=T.yyyymm"
	''sqlStr = sqlStr + " and [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum.shopid=T.socid"
	''sqlStr = sqlStr + " and [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum.makerid=T.imakerid"
	''
	''rsget.Open sqlStr, dbget, 1

    '' time out ?
    sqlStr = " exec db_shop.dbo.sp_ten_Create_shop_brand_monthly_Sum '"&yyyymm&"' ,NULL,NULL,'"&session("ssBctId")&"'"
    dbget.Execute sqlStr

	response.write "<script>alert('작성되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="MakeMonthlyBrandStockSum" then

    sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"' "
    dbget.Execute sqlStr

	response.write "<script>alert('작성되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="" then

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
