<%
Class CRateitem
	Public FDate
	Public FWeek
	Public FUSD
	Public FCNY
	Public FMYR
	Public FSGD
	Public FRegdate
	Public FRegUserid
	Public FLastUpdate
	Public FLastUserid

	Public FOrderserial
	Public FAuthcode
	Public FBeadaldate
	Public FSitename
	Public FMallSumprice
	Public FMallMaxShipping
	Public FMallTotprice
	Public FKRsellPrice
	Public FKRshipping
	Public FKRtotPrice
	Public FDeliverno
	public FitemWeigth
End Class

Class CRate
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	Public FRectYear
	Public FRectMonth
	Public FRectMallGubun
	Public FRectSDt
	Public FRectEDt

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Public Sub getdayRateList
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT L.solar_date, L.week "
		sqlStr = sqlStr & " , convert(varchar,R.USD) as USD, convert(varchar,R.CNY) as CNY, convert(varchar,R.MYR) as MYR, convert(varchar,R.SGD) as SGD, R.regdate, R.regUserid, R.lastUpdate, R.lastUserid "
		sqlStr = sqlStr & " from db_sitemaster.dbo.LunarToSolar as L "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_dayexchageRate as R on L.solar_date = R.yyyymmdd "
		sqlStr = sqlStr & " where left(L.solar_date,7)='" & FRectYear&"-"&FRectMonth & "'"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		If  not rsget.EOF  then
			i = 0
			Do until rsget.eof
				Set FItemList(i) = new CRateitem
					FItemList(i).FDate			= rsget("solar_date")
					FItemList(i).FWeek			= rsget("week")
					FItemList(i).FUSD			= rsget("USD")
					FItemList(i).FCNY			= rsget("CNY")
					FItemList(i).FMYR			= rsget("MYR")
					FItemList(i).FSGD			= rsget("SGD")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegUserid		= rsget("regUserid")
					FItemList(i).FLastUpdate	= rsget("lastUpdate")
					FItemList(i).FLastUserid	= rsget("lastUserid")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.close
	End Sub

	Public Sub getdayRateOrderList
		Dim sqlStr, i, addSql

		If FRectMallGubun <> "" then
			addSql = addSql & " and m.sitename = '"&FRectMallGubun&"' "
		End If
		
		If (FRectSDt <> "") AND (FRectEDt <> "") then
			addSql = addSql & "	and m.beadaldate >='"&FRectSDt&" 00:00:00' and m.beadaldate <='"&FRectEDt&" 23:59:59' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg FROM ( "
		sqlStr = sqlStr & " 	SELECT m.orderserial, m.authcode, m.beadaldate, m.sitename, R.USD, R.CNY, R.MYR, R.SGD "
		sqlStr = sqlStr & " 	,Sum(T.overseasRealPrice * T.orgOrderCNT) as mallSumprice, max(T.overseasDeliveryPrice) as mallMaxShipping, (Sum(T.overseasRealPrice * T.orgOrderCNT) + max(T.overseasDeliveryPrice)) as mallTotprice "
		sqlStr = sqlStr & " 	,CASE WHEN m.sitename = 'cnglob10x10' THEN R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'etsy' THEN R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'zilingo' THEN R.SGD * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'cnhigo' THEN R.CNY * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = '11stmy' THEN R.MYR * Sum(T.overseasRealPrice * T.orgOrderCNT) End as KRsellPrice "
		sqlStr = sqlStr & " 	,CASE WHEN m.sitename = 'cnglob10x10' THEN R.USD * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'etsy' THEN R.USD * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'zilingo' THEN R.SGD * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'cnhigo' THEN R.CNY * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = '11stmy' THEN R.MYR * max(T.overseasDeliveryPrice) End as KRshipping "
		sqlStr = sqlStr & " 	,CASE WHEN m.sitename = 'cnglob10x10' THEN (R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.USD * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'etsy' THEN (R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.USD * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'zilingo' THEN (R.SGD * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.SGD * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = 'cnhigo' THEN (R.CNY * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.CNY * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 		  WHEN m.sitename = '11stmy' THEN (R.MYR * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.MYR * max(T.overseasDeliveryPrice)) End as KRtotPrice "
		sqlStr = sqlStr & " 	, m.deliverno "
		sqlStr = sqlStr & " 	from [db_order].[dbo].tbl_order_master m  "
		sqlStr = sqlStr & " 	JOIN db_item.dbo.tbl_dayexchageRate as R on LEFT(convert(varchar,m.beadaldate,20),10) = R.yyyymmdd "
		sqlStr = sqlStr & " 	JOIN db_temp.dbo.tbl_xSite_TMPOrder as T on m.orderserial = T.orderserial "
		sqlStr = sqlStr & " 	WHERE 1 = 1 and m.ipkumdiv='8' and m.accountdiv='50' and m.cancelyn = 'N' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " 	and T.matchitemid in ( "
		sqlStr = sqlStr & " 		SELECT itemid "
		sqlStr = sqlStr & " 		FROM db_order.dbo.tbl_order_detail as d "
		sqlStr = sqlStr & " 		WHERE d.orderserial = m.orderserial  "
		sqlStr = sqlStr & " 		and d.itemid <> 0  "
		sqlStr = sqlStr & " 		and d.cancelyn <> 'Y'  "
		sqlStr = sqlStr & " 	) "
		sqlStr = sqlStr & " 	GROUP BY m.orderserial, m.authcode, m.sitename, m.beadaldate, R.USD, R.CNY, R.MYR, R.SGD, m.deliverno "
		sqlStr = sqlStr & " ) T "
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.orderserial, m.authcode, m.beadaldate, m.sitename, R.USD, R.CNY, R.MYR, R.SGD "
		sqlStr = sqlStr & " ,Sum(T.overseasRealPrice * T.orgOrderCNT) as mallSumprice "
		sqlStr = sqlStr & " ,max(T.overseasDeliveryPrice) as mallMaxShipping, (Sum(T.overseasRealPrice * T.orgOrderCNT) + max(T.overseasDeliveryPrice)) as mallTotprice "
		sqlStr = sqlStr & " ,CASE WHEN m.sitename = 'cnglob10x10' THEN R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'etsy' THEN R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'zilingo' THEN R.SGD * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'cnhigo' THEN R.CNY * Sum(T.overseasRealPrice * T.orgOrderCNT) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = '11stmy' THEN R.MYR * Sum(T.overseasRealPrice * T.orgOrderCNT) End as KRsellPrice "
		sqlStr = sqlStr & " ,CASE WHEN m.sitename = 'cnglob10x10' THEN R.USD * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'etsy' THEN R.USD * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'zilingo' THEN R.SGD * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'cnhigo' THEN R.CNY * max(T.overseasDeliveryPrice) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = '11stmy' THEN R.MYR * max(T.overseasDeliveryPrice) End as KRshipping "
		sqlStr = sqlStr & " ,CASE WHEN m.sitename = 'cnglob10x10' THEN (R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.USD * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'etsy' THEN (R.USD * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.USD * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'zilingo' THEN (R.SGD * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.SGD * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = 'cnhigo' THEN (R.CNY * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.CNY * max(T.overseasDeliveryPrice)) "
		sqlStr = sqlStr & " 	  WHEN m.sitename = '11stmy' THEN (R.MYR * Sum(T.overseasRealPrice * T.orgOrderCNT)) + (R.MYR * max(T.overseasDeliveryPrice)) End as KRtotPrice "
		sqlStr = sqlStr & " , m.deliverno, em.itemWeigth "
		sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_master m  "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_dayexchageRate as R on LEFT(convert(varchar,m.beadaldate,20),10) = R.yyyymmdd "
		sqlStr = sqlStr & " JOIN db_temp.dbo.tbl_xSite_TMPOrder as T on m.orderserial = T.orderserial "
		sqlStr = sqlStr & " left join db_order.[dbo].[tbl_ems_orderInfo] em on m.orderserial=em.orderserial"
		sqlStr = sqlStr & " WHERE 1 = 1 and m.ipkumdiv='8' and m.accountdiv='50' and m.cancelyn = 'N' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " and T.matchitemid in ( "
		sqlStr = sqlStr & " 	SELECT itemid "
		sqlStr = sqlStr & " 	FROM db_order.dbo.tbl_order_detail as d "
		sqlStr = sqlStr & " 	WHERE d.orderserial = m.orderserial  "
		sqlStr = sqlStr & " 	and d.itemid <> 0  "
		sqlStr = sqlStr & " 	and d.cancelyn <> 'Y'  "
		sqlStr = sqlStr & " ) "
		sqlStr = sqlStr & " GROUP BY m.orderserial, m.authcode, m.sitename, m.beadaldate, R.USD, R.CNY, R.MYR, R.SGD, m.deliverno, em.itemWeigth "
		sqlStr = sqlStr & " ORDER BY m.beadaldate DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CRateitem
					FItemList(i).FOrderserial		= rsget("orderserial")
					FItemList(i).FAuthcode			= rsget("authcode")
					FItemList(i).FBeadaldate		= rsget("beadaldate")
					FItemList(i).FSitename			= rsget("sitename")
					FItemList(i).FUSD				= rsget("USD")
					FItemList(i).FCNY				= rsget("CNY")
					FItemList(i).FMYR				= rsget("MYR")
					FItemList(i).FSGD				= rsget("SGD")
					FItemList(i).FMallSumprice		= rsget("mallSumprice")
					FItemList(i).FMallMaxShipping	= rsget("mallMaxShipping")
					FItemList(i).FMallTotprice		= rsget("mallTotprice")
					FItemList(i).FKRsellPrice		= rsget("KRsellPrice")
					FItemList(i).FKRshipping		= rsget("KRshipping")
					FItemList(i).FKRtotPrice		= rsget("KRtotPrice")
					FItemList(i).FDeliverno			= rsget("deliverno")
					FItemList(i).FitemWeigth			= rsget("itemWeigth")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
End Class
%>