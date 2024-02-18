<%
Class COffEventItem

End Class

Class COffEvent
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FShopName
	Public FTotalCouponCnt
	Public FUserJoinCnt
	Public FAlReadyUserCnt
	Public FDelUserCnt
	Public FOnBuyUserCnt
	Public FOnAvgDays
	Public FActiveUserCnt
	Public FTermCouponCnt
	Public FTermAlreayUserCnt
	Public FTermJoinCnt
	Public FTermDelUserCnt
	Public FTermOnBuyUserCnt
	Public FTermOnAvgDays
	Public FTermActiveUserCnt
	Public FTermActiveUserTotalSum
	Public FActiveUserTotalSum

	Public FRectSdate
	Public FRectEdate
	Public FRectEventNo
	Public FRectBuyprice
	Public FRectAppRunUser
	Public FRectAppRunDay
	Public FRectShopid

	Public FTotalVisitCount
	Public FTotalJumunCount
	Public FTermVisitCount
	Public FTermJumunCount

	Public FTotalUserCnt
	Public FTotalManCnt
	Public FTotalWomenCnt
	Public FTotalVVIPCnt
	Public FTotalVIPGOLDCnt
	Public FTotalVIPSILVERCnt
	Public FTotalBLUECnt
	Public FTotalGREENCnt
	Public FTotalYELLOWCnt
	Public FTotalORANGECnt

	Public FTermUserCnt
	Public FTermManCnt
	Public FTermWomenCnt
	Public FTermVVIPCnt
	Public FTermVIPGOLDCnt
	Public FTermVIPSILVERCnt
	Public FTermBLUECnt
	Public FTermGREENCnt
	Public FTermYELLOWCnt
	Public FTermORANGECnt

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	Public Function fnOffEventReport
		Dim strSql, addSql, addSql2, i, minDate, maxDate
		Dim TBLshopAddSql
		If FRectSdate <> "" AND FRectEdate <> "" Then
			addSql = addSql & " and convert(varchar(10), T.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), T.regdate, 120) <= '" & (FRectEdate) & "' "
			TBLshopAddSql = TBLshopAddSql & " and convert(varchar(10), visitDate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), visitDate, 120) <= '" & (FRectEdate) & "' "
		End If

		If (FRectAppRunUser <> "" AND FRectAppRunDay <> "") Then
			If FRectAppRunUser = 0 Then
				If FRectAppRunDay <> 0 Then
					addSql2 = addSql2 & " datediff(d, T.regdate, T.lastmatchrunlastupdate) > 0 AND datediff(d, T.regdate, T.lastmatchrunlastupdate) <= '"&FRectAppRunDay&"' "
				Else
					addSql2 = addSql2 & " datediff(d,T.regdate,T.lastmatchrunlastupdate) > 0 "
				End If
			Else
				If FRectAppRunDay <> 0 Then
					addSql2 = addSql2 & " datediff(d, T.regdate, G.buydate) > 0 AND datediff(d, T.regdate, G.buydate) <= '"&FRectAppRunDay&"' "
				Else
					addSql2 = addSql2 & " datediff(d,T.regdate,G.buydate) > 0 "
				End If
			End If
		End If

		strSql = ""
		strSql = strSql & " SELECT T.shopid, s.shopname, T.regdate, T.lastmatchJoindate, T.eventNo, T.lastmatchUserid, D.userid, T.lastmatchrunlastupdate "
		strSql = strSql & " INTO #TBLOFF "
		strSql = strSql & " FROM db_contents.[dbo].[tbl_app_offshop_inflow] T "
		strSql = strSql & " LEFT JOIN db_user.[dbo].[tbl_deluser] as D on T.lastmatchUserid = D.userid "
		strSql = strSql & " LEFT JOIN db_shop.dbo.tbl_shop_user s on T.shopid = s.userid  "
		strSql = strSql & " WHERE 1=1 "
		If FRectEventNo <> "" Then
			strSql = strSql & " and T.eventNo = '"&FRectEventNo&"' "
		End If
		dbget.Execute strSql
'rw strSql & "------------------"&"<br>"
		strSql = ""
		strSql = strSql & " Select TOP 1 "
		strSql = strSql & " min(convert(varchar(10), regdate, 120)) as minDate, max(convert(varchar(10), regdate, 120)) as maxDate "
		strSql = strSql & " from #TBLOFF "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			minDate = rsget("minDate")
			maxDate = rsget("maxDate")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT g.shopid,convert(varchar(10),g.yyyymmdd,21) as visitDate "
		strSql = strSql & " ,sum(round(convert(float,isnull(z1_in,0)+isnull(z1_out,0))/2,0)) + sum(round(convert(float,isnull(z2_in,0)+isnull(z2_out,0))/2,0)) as visitCount "
		strSql = strSql & " ,jumun.cnt as jumuncnt "
		strSql = strSql & " INTO #TBLshop "
		strSql = strSql & " from db_shop.dbo.tbl_shop_guestcount g "
		strSql = strSql & " LEFT JOIN "
		strSql = strSql & " ( "
		strSql = strSql & " 	SELECT m.shopid, convert(varchar(10),m.ixyyyymmdd,121) as yyyymmdd "
		strSql = strSql & " 	, count(distinct(m.idx)) as cnt "
		strSql = strSql & " 	FROM db_shop.dbo.tbl_shopjumun_master as m "
		strSql = strSql & " 	JOIN db_shop.dbo.tbl_shopjumun_detail as d on m.idx = d.masteridx "
		strSql = strSql & " 	LEFT JOIN db_partner.dbo.tbl_partner jp on m.shopid=jp.id  "
		strSql = strSql & " 	WHERE 1=1 and m.ixyyyymmdd >= '"&minDate&"' AND m.ixyyyymmdd <= '"&maxDate&"' "
		strSql = strSql & " 	and m.shopid in (SELECT shopid FROM #TBLOFF) "
		strSql = strSql & " 	group by m.shopid ,convert(varchar(10),m.ixyyyymmdd,121) "
		strSql = strSql & " ) as jumun on g.shopid = jumun.shopid and convert(varchar(10),g.yyyymmdd,121) = jumun.yyyymmdd "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and g.yyyymmdd >= '"&minDate&"' AND g.yyyymmdd <= '"&maxDate&"'  "
		strSql = strSql & " and g.shopid in (SELECT shopid FROM #TBLOFF) "
		strSql = strSql & " GROUP BY g.shopid,convert(varchar(10),g.yyyymmdd,21),jumun.cnt "
		dbget.Execute strSql
'rw strSql & "------------------"&"<br>"

		strSql = ""
		strSql = strSql & " SELECT isnull(SUM(visitCount), 0) as visitCount, isnull(SUM(jumuncnt), 0) as jumunCount FROM #TBLshop "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalVisitCount	= rsget("visitCount")
			FTotalJumunCount	= rsget("jumunCount")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT isnull(SUM(visitCount), 0) as visitCount, isnull(SUM(jumuncnt), 0) as jumunCount FROM #TBLshop WHERE 1=1 " & TBLshopAddSql
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTermVisitCount	= rsget("visitCount")
			FTermJumunCount	= rsget("jumunCount")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT userid, buyDate, convert(money, 0) as totalSum, 0 as useridCnt "
		strSql = strSql & " INTO #TBLOnLineORDER "
		strSql = strSql & " FROM ( "
		strSql = strSql & " 	SELECT LM.userid, LM.regdate as buyDate, ROW_NUMBER() OVER (PARTITION BY LM.userid ORDER BY MIN(LM.regdate) ASC) AS Row "
		strSql = strSql & " 	FROM #TBLOFF OO "
		strSql = strSql & " 	JOIN db_order.dbo.tbl_order_master LM on OO.lastmatchUserid = LM.userid and convert(varchar(10),OO.regdate,21) <= convert(varchar(10),LM.regdate,21) "
		strSql = strSql & " 	WHERE LM.cancelyn='N' "
		strSql = strSql & " 	AND LM.ipkumdiv>1 "
		strSql = strSql & " 	AND LM.jumundiv<>9 "
		If FRectBuyprice <> "" Then
			strSql = strSql & " and LM.totalsum >= '"&FRectBuyprice&"' "
		End If
		strSql = strSql & " 	group by LM.userid, LM.regdate "
		strSql = strSql & " ) LM "
		strSql = strSql & " WHERE ROW = 1 "
		dbget.Execute strSql
'rw strSql & "------------------"&"<br>"

		strSql = ""
		strSql = strSql & " update R "
		strSql = strSql & " set totalSum=isNULL(T.totalSum,0) "
		strSql = strSql & " ,useridCnt = isNULL(T.useridCnt,0) "
		strSql = strSql & " from #TBLOnLineORDER R "
		strSql = strSql & " Join ( "
		strSql = strSql & " 	SELECT LM.userid, Sum(LM.totalsum) as totalSum, count(LM.userid) as useridCnt "
		strSql = strSql & " 	FROM #TBLOFF OO  "
		strSql = strSql & " 	JOIN db_order.dbo.tbl_order_master LM on OO.lastmatchUserid = LM.userid and convert(varchar(10),OO.regdate,21) <= convert(varchar(10),LM.regdate,21)  "
		strSql = strSql & " 	JOIN #TBLOnLineORDER as O on LM.userid = O.userid "
		strSql = strSql & " 	WHERE LM.cancelyn='N' AND LM.ipkumdiv>1 AND LM.jumundiv<>9 "
		If FRectBuyprice <> "" Then
			strSql = strSql & " and LM.totalsum >= '"&FRectBuyprice&"' "
		End If
		strSql = strSql & " 	GROUP BY LM.userid "
		strSql = strSql & " ) T on R.userid=T.userid "
		dbget.Execute strSql

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as CPNCNT "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) >= convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') newUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) < convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') alReadyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(T.userid, '') <> '')  THEN 1 else 0 END), '0') DelUserCnt  "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '')  THEN 1 else 0 END), '0') OnBuyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') THEN datediff(d,T.regdate,G.buyDate) else 0 END), '0') OnAvgDays "
		strSql = strSql & " , isnull(sum(CASE WHEN (datediff(d,T.regdate,T.lastmatchrunlastupdate) > 0)  THEN 1 else 0 END), '0') activeUserCnt "
		strSql = strSql & " , isnull(sum(G.totalSum), '0') as totalSum "
		strSql = strSql & " FROM #TBLOFF T "
		strSql = strSql & " LEFT JOIN #TBLOnLineORDER G on T.lastmatchUserid = G.userid  "
		strSql = strSql & " WHERE 1 = 1 "
		rsget.Open strSql,dbget,1
			FTotalCouponCnt		= rsget("CPNCNT")
			FUserJoinCnt		= rsget("newUserCnt")
			FAlReadyUserCnt		= rsget("alReadyUserCnt")
			FDelUserCnt			= rsget("DelUserCnt")
			FOnBuyUserCnt		= rsget("OnBuyUserCnt")
			FOnAvgDays			= rsget("OnAvgDays")
			FActiveUserCnt		= rsget("activeUserCnt")
			FActiveUserTotalSum	= rsget("totalSum")
		rsget.Close
'rw strSql & "------------------"&"<br>"
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as CPNCNT "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) >= convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') newUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) < convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') alReadyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(T.userid, '') <> '')  THEN 1 else 0 END), '0') DelUserCnt  "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') "& addSql &" THEN 1 else 0 END), '0') OnBuyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') "& addSql &" THEN datediff(d,T.regdate,G.buyDate) else 0 END), '0') OnAvgDays "
		strSql = strSql & " , isnull(sum(CASE WHEN ("&addSql2&")  THEN 1 else 0 END), '0') activeUserCnt "
		strSql = strSql & " , isnull(sum(G.totalSum), '0') as totalSum "
		strSql = strSql & " FROM #TBLOFF T "
		strSql = strSql & " LEFT JOIN #TBLOnLineORDER G on T.lastmatchUserid = G.userid  "
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & addSql
		rsget.Open strSql,dbget,1
			FTermCouponCnt			= rsget("CPNCNT")
			FTermAlreayUserCnt		= rsget("alReadyUserCnt")
			FTermJoinCnt			= rsget("newUserCnt")
			FTermDelUserCnt			= rsget("DelUserCnt")
			FTermOnBuyUserCnt		= rsget("OnBuyUserCnt")
			FTermOnAvgDays			= rsget("OnAvgDays")
			FTermActiveUserCnt		= rsget("activeUserCnt")
			FTermActiveUserTotalSum	= rsget("totalSum")
		rsget.Close
'rw strSql & "------------------"&"<br>"
		strSql = ""
		strSql = strSql & " SELECT TOP 100 T.shopid, T.shopname, count(*) as CPNCNT "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) >= convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') newUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) < convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') alReadyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(T.userid, '') <> '')  THEN 1 else 0 END), '0') DelUserCnt  "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') "& addSql &" THEN 1 else 0 END), '0') OnBuyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') "& addSql &" THEN datediff(d,T.regdate,G.buyDate) else 0 END), '0') OnAvgDays "
		strSql = strSql & ", isnull(sum(CASE WHEN ("&addSql2&")  THEN 1 else 0 END), '0') activeUserCnt "
		strSql = strSql & ", isnull(Q.visitCount, 0) as visitCount "
		strSql = strSql & ", isnull(Q.jumuncnt, 0) as jumuncnt "
		strSql = strSql & " , isnull(sum(G.totalSum), '0') as totalSum "
		strSql = strSql & " FROM #TBLOFF T "
		strSql = strSql & " LEFT JOIN #TBLOnLineORDER G on T.lastmatchUserid = G.userid  "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT shopid, Sum(visitCount) as visitCount, Sum(jumuncnt) as jumuncnt "
		strSql = strSql & " 	FROM #TBLshop "
		strSql = strSql & " 	WHERE 1=1 " & TBLshopAddSql
		strSql = strSql & " 	GROUP BY shopid "
		strSql = strSql & " ) Q on T.shopid = Q.shopid and Q.shopid is not null "
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & addSql
		strSql = strSql & " GROUP BY T.shopid, T.shopname,Q.visitCount, Q.jumuncnt "
		strSql = strSql & " ORDER BY T.shopid "
	    rsget.Open strSql,dbget,1
	    If not rsget.EOF Then
	        fnOffEventReport = rsget.getRows()
	    End If
	    rsget.Close
'rw strSql & "------------------"&"<br>"
	End Function

	Public Function fnOffEventReportByShop
		Dim strSql, addSql, i, addSql2, TBLshopAddSql
		Dim minDate, maxDate

		If FRectSdate <> "" AND FRectEdate <> "" Then
			addSql = addSql & " and convert(varchar(10), T.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), T.regdate, 120) <= '" & (FRectEdate) & "' "
			TBLshopAddSql = TBLshopAddSql & " and convert(varchar(10), visitDate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), visitDate, 120) <= '" & (FRectEdate) & "' "
		End If

		strSql = ""
		strSql = strSql & " SELECT T.*, D.userid as delUserid, s.shopname "
		strSql = strSql & " INTO #TBLBUFShop "
		strSql = strSql & " FROM db_contents.[dbo].[tbl_app_offshop_inflow] T "
		strSql = strSql & " LEFT JOIN db_user.[dbo].[tbl_deluser] as D on T.lastmatchUserid = D.userid "
		strSql = strSql & " LEFT JOIN db_shop.dbo.tbl_shop_user s on T.shopid = s.userid "
		strSql = strSql & " WHERE 1 = 1 "
		If FRectShopid <> "" Then
			strSql = strSql & " and T.shopid = '"&FRectShopid&"' "
		End If
		If FRectEventNo <> "" Then
			strSql = strSql & " and T.eventNo = '"&FRectEventNo&"' "
		End If
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " Select TOP 1 "
		strSql = strSql & " min(convert(varchar(10), regdate, 120)) as minDate, max(convert(varchar(10), regdate, 120)) as maxDate "
		strSql = strSql & " from db_contents.[dbo].[tbl_app_offshop_inflow] "
		strSql = strSql & " WHERE eventNo = '"&FRectEventNo&"' "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			minDate = rsget("minDate")
			maxDate = rsget("maxDate")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT g.shopid,convert(varchar(10),g.yyyymmdd,21) as visitDate "
		strSql = strSql & " ,sum(round(convert(float,isnull(z1_in,0)+isnull(z1_out,0))/2,0)) + sum(round(convert(float,isnull(z2_in,0)+isnull(z2_out,0))/2,0)) as visitCount "
		strSql = strSql & " ,jumun.cnt as jumuncnt "
		strSql = strSql & " INTO #TBLshop2 "
		strSql = strSql & " from db_shop.dbo.tbl_shop_guestcount g "
		strSql = strSql & " LEFT JOIN "
		strSql = strSql & " ( "
		strSql = strSql & " 	SELECT m.shopid, convert(varchar(10),m.ixyyyymmdd,121) as yyyymmdd "
		strSql = strSql & " 	, count(distinct(m.idx)) as cnt "
		strSql = strSql & " 	FROM db_shop.dbo.tbl_shopjumun_master as m "
		strSql = strSql & " 	JOIN db_shop.dbo.tbl_shopjumun_detail as d on m.idx = d.masteridx "
		strSql = strSql & " 	LEFT JOIN db_partner.dbo.tbl_partner jp on m.shopid=jp.id  "
		strSql = strSql & " 	WHERE 1=1 and m.ixyyyymmdd >= '"&minDate&"' AND m.ixyyyymmdd <= '"&maxDate&"' "
		strSql = strSql & " 	and m.shopid in (SELECT shopid FROM #TBLBUFShop) "
		strSql = strSql & " 	group by m.shopid ,convert(varchar(10),m.ixyyyymmdd,121) "
		strSql = strSql & " ) as jumun on g.shopid = jumun.shopid and convert(varchar(10),g.yyyymmdd,121) = jumun.yyyymmdd "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and g.yyyymmdd >= '"&minDate&"' AND g.yyyymmdd <= '"&maxDate&"'  "
		strSql = strSql & " and g.shopid in (SELECT shopid FROM #TBLBUFShop) "
		strSql = strSql & " GROUP BY g.shopid,convert(varchar(10),g.yyyymmdd,21),jumun.cnt "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT isnull(SUM(visitCount), 0) as visitCount, isnull(SUM(jumuncnt), 0) as jumunCount FROM #TBLshop2 WHERE 1=1 " & TBLshopAddSql
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTermVisitCount	= rsget("visitCount")
			FTermJumunCount	= rsget("jumunCount")
		End If
		rsget.Close
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT userid, buyDate, convert(money, 0) as totalSum, 0 as useridCnt "
		strSql = strSql & " INTO #TBLOnLineORDER2 "
		strSql = strSql & " FROM ( "
		strSql = strSql & " 	SELECT LM.userid, LM.regdate as buyDate, ROW_NUMBER() OVER (PARTITION BY LM.userid ORDER BY MIN(LM.regdate) ASC) AS Row "
		strSql = strSql & " 	FROM #TBLBUFShop OO "
		strSql = strSql & " 	JOIN db_order.dbo.tbl_order_master LM on OO.lastmatchUserid = LM.userid and convert(varchar(10),OO.regdate,21) <= convert(varchar(10),LM.regdate,21) "
		strSql = strSql & " 	WHERE LM.cancelyn='N' "
		strSql = strSql & " 	AND LM.ipkumdiv>1 "
		strSql = strSql & " 	AND LM.jumundiv<>9 "
		If FRectBuyprice <> "" Then
			strSql = strSql & " and LM.totalsum >= '"&FRectBuyprice&"' "
		End If
		strSql = strSql & " 	group by LM.userid, LM.regdate "
		strSql = strSql & " ) LM "
		strSql = strSql & " WHERE ROW = 1 "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " update R "
		strSql = strSql & " set totalSum=isNULL(T.totalSum,0) "
		strSql = strSql & " ,useridCnt = isNULL(T.useridCnt,0) "
		strSql = strSql & " from #TBLOnLineORDER2 R "
		strSql = strSql & " Join ( "
		strSql = strSql & " 	SELECT LM.userid, Sum(LM.totalsum) as totalSum, count(LM.userid) as useridCnt "
		strSql = strSql & " 	FROM #TBLBUFShop OO  "
		strSql = strSql & " 	JOIN db_order.dbo.tbl_order_master LM on OO.lastmatchUserid = LM.userid and convert(varchar(10),OO.regdate,21) <= convert(varchar(10),LM.regdate,21)  "
		strSql = strSql & " 	JOIN #TBLOnLineORDER2 as O on LM.userid = O.userid "
		strSql = strSql & " 	WHERE LM.cancelyn='N' AND LM.ipkumdiv>1 AND LM.jumundiv <> 9 "
		If FRectBuyprice <> "" Then
			strSql = strSql & " and LM.totalsum >= '"&FRectBuyprice&"' "
		End If
		strSql = strSql & " 	GROUP BY LM.userid "
		strSql = strSql & " ) T on R.userid=T.userid "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT T.shopname, COUNT(*) as CPNCNT "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) >= convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') newUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) < convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') alReadyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(T.deluserid, '') <> '')  THEN 1 else 0 END), '0') DelUserCnt  "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') "& addSql &" THEN 1 else 0 END), '0') OnBuyUserCnt "
		strSql = strSql & " , isnull(sum(G.totalSum), '0') as totalSum "
		strSql = strSql & " FROM #TBLBUFShop T "
		strSql = strSql & " LEFT JOIN #TBLOnLineORDER2 G on T.lastmatchUserid = G.userid  "
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & addSql
		strSql = strSql & " GROUP BY T.shopname "
		rsget.Open strSql,dbget,1
			FShopName			= rsget("shopname")
			FTotalCouponCnt		= rsget("CPNCNT")
			FAlReadyUserCnt		= rsget("alReadyUserCnt")
			FUserJoinCnt		= rsget("newUserCnt")
			FDelUserCnt			= rsget("DelUserCnt")
			FOnBuyUserCnt		= rsget("OnBuyUserCnt")
			FActiveUserTotalSum	= rsget("totalSum")
		rsget.Close
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT TOP 100 T.shopname, convert(varchar(10),T.regdate,21) as regdate, count(*) as CPNCNT "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) >= convert(varchar(10),T.regdate,21) THEN 1 else 0 END), '0') newUserCnt "
		strSql = strSql & " , sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) < convert(varchar(10), T.regdate, 21)  THEN 1 else 0 END) alReadyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(T.deluserid, '') <> '')  THEN 1 else 0 END), '0') DelUserCnt  "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '') "& addSql &" THEN 1 else 0 END), '0') OnBuyUserCnt "
		strSql = strSql & " , isnull(sum(G.totalSum), '0') as totalSum "
		strSql = strSql & " , isnull(Q.visitCount, 0) as visitCount "
		strSql = strSql & " , isnull(Q.jumuncnt, 0) as jumuncnt "
		strSql = strSql & " FROM #TBLBUFShop T "
		strSql = strSql & " LEFT JOIN #TBLOnLineORDER2 G on T.lastmatchUserid = G.userid  "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT shopid, visitCount, jumuncnt,convert(varchar(10), visitDate, 120) as visitDate "
		strSql = strSql & " 	FROM #TBLshop2 "
		strSql = strSql & " 	WHERE 1=1 " & TBLshopAddSql
		strSql = strSql & " ) Q on T.shopid = Q.shopid and convert(varchar(10),T.regdate,21) = convert(varchar(10), visitDate, 120) "
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & addSql
		strSql = strSql & " GROUP BY shopname, convert(varchar(10),T.regdate,21),Q.visitCount, Q.jumuncnt "
		strSql = strSql & " ORDER BY convert(varchar(10),T.regdate,21) "
	    rsget.Open strSql,dbget,1
	    If not rsget.EOF Then
	        fnOffEventReportByShop = rsget.getRows()
	    End If
	    rsget.Close
'rw strSql
	End Function

	Public Function fnOffEventReportByTerm
		Dim strSql, addSql, i
		Dim minDate, maxDate, TBLshopAddSql
		strSql = ""
		strSql = strSql & " SELECT T.shopid, s.shopname, T.regdate, T.lastmatchJoindate, T.eventNo, T.lastmatchUserid, D.userid "
		strSql = strSql & " INTO #TBLOFFTerm "
		strSql = strSql & " FROM db_contents.[dbo].[tbl_app_offshop_inflow] T "
		strSql = strSql & " LEFT JOIN db_user.[dbo].[tbl_deluser] as D on T.lastmatchUserid = D.userid "
		strSql = strSql & " LEFT JOIN db_shop.dbo.tbl_shop_user s on T.shopid = s.userid  "
		strSql = strSql & " WHERE 1=1 "
		If FRectEventNo <> "" Then
			strSql = strSql & " and T.eventNo = '"&FRectEventNo&"' "
		End If
		If FRectSdate <> "" AND FRectEdate <> "" Then
			strSql = strSql & " and convert(varchar(10), T.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), T.regdate, 120) <= '" & (FRectEdate) & "' "
			TBLshopAddSql = TBLshopAddSql & " and convert(varchar(10), visitDate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), visitDate, 120) <= '" & (FRectEdate) & "' "
		End If

'		If (FRectAppRunUser <> "" AND FRectAppRunDay <> "") AND FRectAppRunDay <> "0" Then
'			If FRectAppRunUser = 0 Then
'				strSql = strSql & " and datediff(d, T.regdate, T.lastmatchrunlastupdate) <= '"&FRectAppRunDay&"' "
'			Else
'				strSql = strSql & " and datediff(d, T.regdate, G.buydate) <= '"&FRectAppRunDay&"' "
'			End If
'		End If
		dbget.Execute strSql
'rw strSql

		strSql = ""
		strSql = strSql & " Select TOP 1 "
		strSql = strSql & " min(convert(varchar(10), regdate, 120)) as minDate, max(convert(varchar(10), regdate, 120)) as maxDate "
		strSql = strSql & " from #TBLOFFTerm "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			minDate = rsget("minDate")
			maxDate = rsget("maxDate")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT g.shopid,convert(varchar(10),g.yyyymmdd,21) as visitDate "
		strSql = strSql & " ,sum(round(convert(float,isnull(z1_in,0)+isnull(z1_out,0))/2,0)) + sum(round(convert(float,isnull(z2_in,0)+isnull(z2_out,0))/2,0)) as visitCount "
		strSql = strSql & " ,jumun.cnt as jumuncnt "
		strSql = strSql & " INTO #TBLshop2 "
		strSql = strSql & " from db_shop.dbo.tbl_shop_guestcount g "
		strSql = strSql & " LEFT JOIN "
		strSql = strSql & " ( "
		strSql = strSql & " 	SELECT m.shopid, convert(varchar(10),m.ixyyyymmdd,121) as yyyymmdd "
		strSql = strSql & " 	, count(distinct(m.idx)) as cnt "
		strSql = strSql & " 	FROM db_shop.dbo.tbl_shopjumun_master as m "
		strSql = strSql & " 	JOIN db_shop.dbo.tbl_shopjumun_detail as d on m.idx = d.masteridx "
		strSql = strSql & " 	LEFT JOIN db_partner.dbo.tbl_partner jp on m.shopid=jp.id  "
		strSql = strSql & " 	WHERE 1=1 and m.ixyyyymmdd >= '"&minDate&"' AND m.ixyyyymmdd <= '"&maxDate&"' "
		strSql = strSql & " 	and m.shopid in (SELECT shopid FROM #TBLOFFTerm) "
		strSql = strSql & " 	group by m.shopid ,convert(varchar(10),m.ixyyyymmdd,121) "
		strSql = strSql & " ) as jumun on g.shopid = jumun.shopid and convert(varchar(10),g.yyyymmdd,121) = jumun.yyyymmdd "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and g.yyyymmdd >= '"&minDate&"' AND g.yyyymmdd <= '"&maxDate&"'  "
		strSql = strSql & " and g.shopid in (SELECT shopid FROM #TBLOFFTerm) "
		strSql = strSql & " GROUP BY g.shopid,convert(varchar(10),g.yyyymmdd,21),jumun.cnt "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT userid, buyDate, convert(money, 0) as totalSum, 0 as useridCnt "
		strSql = strSql & " INTO #TBLOnLineORDER "
		strSql = strSql & " FROM ( "
		strSql = strSql & " 	SELECT LM.userid, LM.regdate as buyDate, ROW_NUMBER() OVER (PARTITION BY LM.userid ORDER BY MIN(LM.regdate) ASC) AS Row "
		strSql = strSql & " 	FROM #TBLOFFTerm OO "
		strSql = strSql & " 	JOIN db_order.dbo.tbl_order_master LM on OO.lastmatchUserid = LM.userid and convert(varchar(10),OO.regdate,21) <= convert(varchar(10),LM.regdate,21) "
		strSql = strSql & " 	WHERE LM.cancelyn='N' "
		strSql = strSql & " 	AND LM.ipkumdiv>1 "
		strSql = strSql & " 	AND LM.jumundiv<>9 "
		If FRectBuyprice <> "" Then
			strSql = strSql & " and LM.totalsum >= '"&FRectBuyprice&"' "
		End If
		strSql = strSql & " 	group by LM.userid, LM.regdate "
		strSql = strSql & " ) LM "
		strSql = strSql & " WHERE ROW = 1 "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " update R "
		strSql = strSql & " set totalSum=isNULL(T.totalSum,0) "
		strSql = strSql & " ,useridCnt = isNULL(T.useridCnt,0) "
		strSql = strSql & " from #TBLOnLineORDER R "
		strSql = strSql & " Join ( "
		strSql = strSql & " 	SELECT LM.userid, Sum(LM.totalsum) as totalSum, count(LM.userid) as useridCnt "
		strSql = strSql & " 	FROM #TBLOFFTerm OO  "
		strSql = strSql & " 	JOIN db_order.dbo.tbl_order_master LM on OO.lastmatchUserid = LM.userid and convert(varchar(10),OO.regdate,21) <= convert(varchar(10),LM.regdate,21)  "
		strSql = strSql & " 	JOIN #TBLOnLineORDER as O on LM.userid = O.userid "
		strSql = strSql & " 	WHERE LM.cancelyn='N' AND LM.ipkumdiv>1 AND LM.jumundiv<>9 "
		If FRectBuyprice <> "" Then
			strSql = strSql & " and LM.totalsum >= '"&FRectBuyprice&"' "
		End If
		strSql = strSql & " 	GROUP BY LM.userid "
		strSql = strSql & " ) T on R.userid=T.userid "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT convert(varchar(10),T.regdate,21) as regdate, COUNT(*) as CPNCNT "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) >= convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') newUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN T.lastmatchJoindate is NOT NULL and convert(varchar(10),T.lastmatchJoindate,21) < convert(varchar(10), T.regdate, 21) THEN 1 else 0 END), '0') alReadyUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(T.userid, '') <> '')  THEN 1 else 0 END), '0') DelUserCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (isnull(G.userid, '') <> '')  THEN 1 else 0 END), '0') OnBuyUserCnt "
		strSql = strSql & " , isnull(sum(G.totalSum), '0') as totalSum "
		strSql = strSql & " , isnull(Q.visitCount, 0) as visitCount "
		strSql = strSql & " , isnull(Q.jumuncnt, 0) as jumuncnt "
		strSql = strSql & " FROM #TBLOFFTerm T  "
		strSql = strSql & " LEFT JOIN #TBLOnLineORDER G on T.lastmatchUserid = G.userid  "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & " 	SELECT Sum(visitCount) as visitCount, Sum(jumuncnt) as jumuncnt,convert(varchar(10), visitDate, 120) as visitDate  "
		strSql = strSql & " 	FROM #TBLshop2 "
		strSql = strSql & " 	WHERE 1=1 " & TBLshopAddSql
		strSql = strSql & " 	GROUP BY convert(varchar(10), visitDate, 120)"
		strSql = strSql & " ) Q on convert(varchar(10),T.regdate,21) = convert(varchar(10), visitDate, 120) "
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & " GROUP BY convert(varchar(10),T.regdate,21),Q.visitCount, Q.jumuncnt  "
		strSql = strSql & " ORDER BY convert(varchar(10),T.regdate,21) "
	    rsget.Open strSql,dbget,1
	    If not rsget.EOF Then
	        fnOffEventReportByTerm = rsget.getRows()
	    End If
	    rsget.Close
'rw strSql
	End Function

	Public Function fnOffEventUserReport
		Dim strSql, addSql1, i, addSql2

		If FRectEventNo <> "" Then
			addSql1 = addSql1 & " and T.eventNo = '"&FRectEventNo&"' "
		End If
		If FRectSdate <> "" AND FRectEdate <> "" Then
			addSql2 = addSql2 & " and convert(varchar(10), T.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), T.regdate, 120) <= '" & (FRectEdate) & "' "
		End If

		strSql = ""
		strSql = strSql & " SELECT T.lastmatchUserid, n.sexflag, l.userlevel "
		strSql = strSql & " INTO #TBLTotalUser "
		strSql = strSql & " FROM db_contents.[dbo].[tbl_app_offshop_inflow] T  "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_n as n on T.lastmatchUserid = n.userid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_logindata as l on T.lastmatchUserid = l.userid "
		strSql = strSql & " WHERE 1=1 " & addSql1
		strSql = strSql & " and T.lastmatchJoindate is NOT NULL "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT T.lastmatchUserid, n.sexflag, l.userlevel, n.birthday "
		strSql = strSql & " INTO #TBLTermUser "
		strSql = strSql & " FROM db_contents.[dbo].[tbl_app_offshop_inflow] T  "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_n as n on T.lastmatchUserid = n.userid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_logindata as l on T.lastmatchUserid = l.userid "
		strSql = strSql & " WHERE 1=1 " & addSql1 & addSql2
		strSql = strSql & " and T.lastmatchJoindate is NOT NULL "
		dbget.Execute strSql
'rw strSql
		strSql = ""
		strSql = strSql & " SELECT count(*) as cnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (sexflag % 2) = 1 THEN 1 else 0 END), '0') as ManCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (sexflag % 2) = 0 THEN 1 else 0 END), '0') as womenCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '6') THEN 1 else 0 END), '0') as VVIPCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '4') THEN 1 else 0 END), '0') as VIPGOLDCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '3') THEN 1 else 0 END), '0') as VIPSILVERCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '2') THEN 1 else 0 END), '0') as BLUECnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '1') THEN 1 else 0 END), '0') as GREENCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '0') THEN 1 else 0 END), '0') as YELLOWCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '5') THEN 1 else 0 END), '0') as ORANGECnt "
		strSql = strSql & " FROM #TBLTotalUser "
		rsget.Open strSql,dbget,1
			FTotalUserCnt		= rsget("cnt")
			FTotalManCnt		= rsget("ManCnt")
			FTotalWomenCnt		= rsget("womenCnt")
			FTotalVVIPCnt		= rsget("VVIPCnt")
			FTotalVIPGOLDCnt	= rsget("VIPGOLDCnt")
			FTotalVIPSILVERCnt	= rsget("VIPSILVERCnt")
			FTotalBLUECnt		= rsget("BLUECnt")
			FTotalGREENCnt		= rsget("GREENCnt")
			FTotalYELLOWCnt		= rsget("YELLOWCnt")
			FTotalORANGECnt		= rsget("ORANGECnt")
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT count(*) as cnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (sexflag % 2) = 1 THEN 1 else 0 END), '0') as ManCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (sexflag % 2) = 0 THEN 1 else 0 END), '0') as womenCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '6') THEN 1 else 0 END), '0') as VVIPCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '4') THEN 1 else 0 END), '0') as VIPGOLDCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '3') THEN 1 else 0 END), '0') as VIPSILVERCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '2') THEN 1 else 0 END), '0') as BLUECnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '1') THEN 1 else 0 END), '0') as GREENCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '0') THEN 1 else 0 END), '0') as YELLOWCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '5') THEN 1 else 0 END), '0') as ORANGECnt "
		strSql = strSql & " FROM #TBLTermUser "
		rsget.Open strSql,dbget,1
			FTermUserCnt		= rsget("cnt")
			FTermManCnt			= rsget("ManCnt")
			FTermWomenCnt		= rsget("womenCnt")
			FTermVVIPCnt		= rsget("VVIPCnt")
			FTermVIPGOLDCnt		= rsget("VIPGOLDCnt")
			FTermVIPSILVERCnt	= rsget("VIPSILVERCnt")
			FTermBLUECnt		= rsget("BLUECnt")
			FTermGREENCnt		= rsget("GREENCnt")
			FTermYELLOWCnt		= rsget("YELLOWCnt")
			FTermORANGECnt		= rsget("ORANGECnt")
		rsget.Close

		strSql = ""
		strSql = strSql & " select "
		strSql = strSql & " Case "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=0 and (datediff(year,birthday,getdate())+1)<20 Then 'v20' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=20 and (datediff(year,birthday,getdate())+1)<25 Then 'v24' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=25 and (datediff(year,birthday,getdate())+1)<30 Then 'v29' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=30 and (datediff(year,birthday,getdate())+1)<35 Then 'v34' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=35 and (datediff(year,birthday,getdate())+1)<40 Then 'v39' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=40 and (datediff(year,birthday,getdate())+1)<50 Then 'v49' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=50 Then 'v50' "
		strSql = strSql & " end as age "
		strSql = strSql & " , count(*) as cnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (sexflag % 2) = 1 THEN 1 else 0 END), '0') as ManCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (sexflag % 2) = 0 THEN 1 else 0 END), '0') as womenCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '6') THEN 1 else 0 END), '0') as VVIPCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '4') THEN 1 else 0 END), '0') as VIPGOLDCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '3') THEN 1 else 0 END), '0') as VIPSILVERCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '2') THEN 1 else 0 END), '0') as BLUECnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '1') THEN 1 else 0 END), '0') as GREENCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '0') THEN 1 else 0 END), '0') as YELLOWCnt "
		strSql = strSql & " , isnull(sum(CASE WHEN (userlevel = '5') THEN 1 else 0 END), '0') as ORANGECnt "
		strSql = strSql & " from #TBLTermUser "
		strSql = strSql & " group by  "
		strSql = strSql & " Case "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=0 and (datediff(year,birthday,getdate())+1)<20 Then 'v20' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=20 and (datediff(year,birthday,getdate())+1)<25 Then 'v24' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=25 and (datediff(year,birthday,getdate())+1)<30 Then 'v29' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=30 and (datediff(year,birthday,getdate())+1)<35 Then 'v34' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=35 and (datediff(year,birthday,getdate())+1)<40 Then 'v39' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=40 and (datediff(year,birthday,getdate())+1)<50 Then 'v49' "
		strSql = strSql & " 	When (datediff(year,birthday,getdate())+1)>=50 Then 'v50' "
		strSql = strSql & " end "
		strSql = strSql & " order by age asc "
'rw strSql
	    rsget.Open strSql,dbget,1
	    If not rsget.EOF Then
	        fnOffEventUserReport = rsget.getRows()
	    End If
	    rsget.Close
	End Function
End Class
%>