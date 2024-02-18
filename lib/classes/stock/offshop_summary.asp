<%

class CShopItemSummaryItem
        public Fyyyymm
        public Fyyyymmdd

	public Fshopid
	public Fitemgubun
	public Fitemid
	public Fitemoption

	public Fsellno
	public Fresellno
	public Flogicsipgono
	public Flogicsreipgono
	public Fbrandipgono
	public Fbrandreipgono

	public Fsysstockno

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CShopItemSummary
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopID
	public FRectItemGubun
	public FRectItemId
	public FRectItemOption

        '샾아이템현재재고
	public function GetShopItemCurrentSummary()
		dim sqlStr, i

		sqlStr = " select top 1 c.shopid, c.itemgubun, c.itemid, c.itemoption, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary c "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		sqlStr = sqlStr + " and c.itemid = '" + FRectItemId + "' "
		sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		rsget.Open sqlStr,dbget,1

		set FOneItem = new CShopItemSummaryItem

		if (not rsget.EOF) then
		        FOneItem.Fshopid        = rsget("shopid")
		        FOneItem.Fitemgubun     = rsget("itemgubun")
		        FOneItem.Fitemid        = rsget("itemid")
		        FOneItem.Fitemoption    = rsget("itemoption")

		        FOneItem.Fsellno         = rsget("sellno")
		        FOneItem.Fresellno       = rsget("resellno")
		        FOneItem.Flogicsipgono   = rsget("logicsipgono")
		        FOneItem.Flogicsreipgono = rsget("logicsreipgono")
		        FOneItem.Fbrandipgono    = rsget("brandipgono")
		        FOneItem.Fbrandreipgono  = rsget("brandreipgono")

		        FOneItem.Fsysstockno    = rsget("sysstockno")
		end if
		rsget.Close
	end function

        '샆아이템 현재재고 목록
	public function GetShopItemCurrentSummaryList()
		dim sqlStr, i

		sqlStr = " select top 1000 c.shopid, c.itemgubun, c.itemid, c.itemoption, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary c "
		sqlStr = sqlStr + " where 1 = 1 "
		if (FRectShopID <> "") then
		        sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		end if
		if (FRectItemGubun <> "") then
		        sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and c.itemid = '" + FRectItemId + "' "
		end if
		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

        '샆아이템월별재고목록
	public function GetShopItemMonthlySummaryList()
		dim sqlStr, i
		dim month_pre_2

		month_pre_2 = Left(dateadd("m", -2, now()), 7)


		sqlStr = " select top 1000 c.shopid, c.itemgubun, c.itemid, c.itemoption, c.yyyymm, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_shopstock_summary c "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and c.yyyymm <= '" + month_pre_2 + "' "


		if (FRectShopID <> "") then
		        sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		end if
		if (FRectItemGubun <> "") then
		        sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and c.itemid = '" + FRectItemId + "' "
		end if
		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		sqlStr = sqlStr + " order by c.yyyymm, c.shopid, c.itemgubun, c.itemid, c.itemoption "
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

                                FItemList(i).Fyyyymm        = rsget("yyyymm")

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

        '샾아이템현재재고(last month)
	public function GetShopItemLastMonthSummary()
		dim sqlStr, i

		sqlStr = " select top 1 c.shopid, c.itemgubun, c.itemid, c.itemoption, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_last_monthly_shopstock c "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		sqlStr = sqlStr + " and c.itemid = '" + FRectItemId + "' "
		sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		rsget.Open sqlStr,dbget,1

		set FOneItem = new CShopItemSummaryItem

		if (not rsget.EOF) then
		        FOneItem.Fshopid        = rsget("shopid")
		        FOneItem.Fitemgubun     = rsget("itemgubun")
		        FOneItem.Fitemid        = rsget("itemid")
		        FOneItem.Fitemoption    = rsget("itemoption")

		        FOneItem.Fsellno         = rsget("sellno")
		        FOneItem.Fresellno       = rsget("resellno")
		        FOneItem.Flogicsipgono   = rsget("logicsipgono")
		        FOneItem.Flogicsreipgono = rsget("logicsreipgono")
		        FOneItem.Fbrandipgono    = rsget("brandipgono")
		        FOneItem.Fbrandreipgono  = rsget("brandreipgono")

		        FOneItem.Fsysstockno    = rsget("sysstockno")
		end if
		rsget.Close
	end function

        '샆아이템일별재고목록
	public function GetShopItemDailySummaryList()
		dim sqlStr, i

		sqlStr = " select top 1000 c.shopid, c.itemgubun, c.itemid, c.itemoption, c.yyyymmdd, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_daily_shopstock_summary c "
		sqlStr = sqlStr + " where 1 = 1 "
		if (FRectShopID <> "") then
		        sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		end if
		if (FRectItemGubun <> "") then
		        sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and c.itemid = '" + FRectItemId + "' "
		end if
		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		sqlStr = sqlStr + " order by c.yyyymmdd, c.shopid, c.itemgubun, c.itemid, c.itemoption "
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

                                FItemList(i).Fyyyymmdd      = rsget("yyyymmdd")

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 100
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>