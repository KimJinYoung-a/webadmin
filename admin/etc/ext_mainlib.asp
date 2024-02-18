<%

CONST C_MINMARGIN = 15
CONST C_MINSTOCK = 5

Class CExtSellDiffItem
	Public Fsellsite
	Public Fcnt
	Public FtotToBeNotSell
	Public FtotToBeSell

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CExtMain
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public Sub GetExtSellDiffList
		dim sqlStr, i

		'// 대상상품
		'// 1. 등록완료 상품
		'// 2. 판매상태 상이
		'// 3. 제외 브랜드 아님
		'// 4. 제외 상품 아님
		'// 5. 최소마진(15) 이상 : 판매전환 대상 상품만
		sqlStr = " select top 1 "
		sqlStr = sqlStr + " 'ssg' as sellsite, count(i.itemid) as cnt "
		sqlStr = sqlStr + " 	, IsNull(sum(case when (i.sellyn <> 'Y' and r.ssgSellYn = 'Y') then 1 else 0 end),0) as totToBeNotSell "
		sqlStr = sqlStr + " 	, IsNull(sum(case when (i.sellyn = 'Y' and r.ssgSellYn = 'N' and (i.limityn <> 'Y' or (i.limitno-i.limitsold) > 5) and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100) >= " & C_MINMARGIN & ") then 1 else 0 end),0) as totToBeSell "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_etcmall.[dbo].[tbl_ssg_regItem] r on i.itemid = r.itemid "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].[tbl_jaehyumall_not_in_makerid] e1 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and e1.mallgubun = 'ssg' "
		sqlStr = sqlStr + " 		and e1.makerid = i.makerid "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].[tbl_jaehyumall_not_in_itemid] e2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and e2.mallgubun = 'ssg' "
		sqlStr = sqlStr + " 		and e2.itemid = i.itemid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and r.ssgStatCd = 7 "
		sqlStr = sqlStr + " 	and r.ssgGoodNo is not NULL "
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(i.sellyn = 'Y' and r.ssgSellYn = 'N' and (i.limityn <> 'Y' or (i.limitno-i.limitsold) > " & C_MINSTOCK & ")) "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(i.sellyn <> 'Y' and r.ssgSellYn = 'Y') "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and e1.makerid is NULL "
		sqlStr = sqlStr + " 	and e2.itemid is NULL "
		sqlStr = sqlStr + " union all "
		sqlStr = sqlStr + " select top 1 "
		sqlStr = sqlStr + " 'gsshop' as sellsite, count(i.itemid) as cnt "
		sqlStr = sqlStr + " 	, IsNull(sum(case when (i.sellyn <> 'Y' and r.gsshopSellYn = 'Y') then 1 else 0 end),0) as totToBeNotSell "
		sqlStr = sqlStr + " 	, IsNull(sum(case when (i.sellyn = 'Y' and r.gsshopSellYn = 'N' and (i.limityn <> 'Y' or (i.limitno-i.limitsold) > 5) and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100) >= " & C_MINMARGIN & ") then 1 else 0 end),0) as totToBeSell "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_gsshop_regitem r on i.itemid = r.itemid "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].[tbl_jaehyumall_not_in_makerid] e1 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and e1.mallgubun = 'gseshop' "
		sqlStr = sqlStr + " 		and e1.makerid = i.makerid "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].[tbl_jaehyumall_not_in_itemid] e2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and e2.mallgubun = 'gseshop' "
		sqlStr = sqlStr + " 		and e2.itemid = i.itemid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and r.gsshopStatCd = 7 "
		sqlStr = sqlStr + " 	and r.gsshopGoodNo is not NULL "
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(i.sellyn = 'Y' and r.gsshopSellYn = 'N' and (i.limityn <> 'Y' or (i.limitno-i.limitsold) > " & C_MINSTOCK & ")) "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(i.sellyn <> 'Y' and r.gsshopSellYn = 'Y') "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and e1.makerid is NULL "
		sqlStr = sqlStr + " 	and e2.itemid is NULL "
		sqlStr = sqlStr + " union all "
		sqlStr = sqlStr + " select top 1 "
		sqlStr = sqlStr + " 'auction' as sellsite, count(i.itemid) as cnt "
		sqlStr = sqlStr + " 	, IsNull(sum(case when (i.sellyn <> 'Y' and r.auctionSellYn = 'Y') then 1 else 0 end),0) as totToBeNotSell "
		sqlStr = sqlStr + " 	, IsNull(sum(case when (i.sellyn = 'Y' and r.auctionSellYn = 'N' and (i.limityn <> 'Y' or (i.limitno-i.limitsold) > 5) and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100) >= " & C_MINMARGIN & ") then 1 else 0 end),0) as totToBeSell "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_etcmall.dbo.tbl_auction_regitem r on i.itemid = r.itemid "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].[tbl_jaehyumall_not_in_makerid] e1 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and e1.mallgubun = 'auction1010' "
		sqlStr = sqlStr + " 		and e1.makerid = i.makerid "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].[tbl_jaehyumall_not_in_itemid] e2 "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and e2.mallgubun = 'auction1010' "
		sqlStr = sqlStr + " 		and e2.itemid = i.itemid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and r.auctionStatCd = 7 "
		sqlStr = sqlStr + " 	and r.auctionGoodNo is not NULL "
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(i.sellyn = 'Y' and r.auctionSellYn = 'N' and (i.limityn <> 'Y' or (i.limitno-i.limitsold) > " & C_MINSTOCK & ")) "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(i.sellyn <> 'Y' and r.auctionSellYn = 'Y') "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and e1.makerid is NULL "										'// 등록제외브랜드
		sqlStr = sqlStr + " 	and e2.itemid is NULL "											'// 등록제외상품
		sqlStr = sqlStr + " 	and i.isusing='Y' "
		sqlStr = sqlStr + " 	and i.isExtUsing='Y' "											'// 외부몰허용상품
		''sqlStr = sqlStr + " 	and c.isExtUsing='Y' "
		sqlStr = sqlStr + " 	and i.deliveryType <> 7 "										'// 업체착불
		sqlStr = sqlStr + " 	and i.itemdiv <> '21' "											'// 딜상품
		sqlStr = sqlStr + " 	and i.deliverfixday not in ('C','X') "							'// 꽃배달, 화물배달
		sqlStr = sqlStr + " 	and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
		sqlStr = sqlStr + " 	and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
		sqlStr = sqlStr + " 	and i.itemdiv < 50 "
		''sqlStr = sqlStr + " 	and i.optioncnt = 0 "

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CExtSellDiffItem
					FItemList(i).Fsellsite			= rsget("sellsite")
					FItemList(i).Fcnt				= rsget("cnt")
					FItemList(i).FtotToBeNotSell	= rsget("totToBeNotSell")
					FItemList(i).FtotToBeSell		= rsget("totToBeSell")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close

	End Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage 		= 1
		FPageSize 		= 100
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
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

End Class

'// 판매상태 상이
function GetExtSellDiff()


end function

%>
