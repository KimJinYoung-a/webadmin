<%

Class CCostPerMeachulItem
    public Fyyyymm
    public FtargetGbn
    public FstockPlace
    public Fshopid
    public Fitemgubun
	public Fitemid
	public Fmakerid
	public Fdefaultmargin

    public FbuySumPrevMonth
	public FbuySumThisMonth
	public FcustomerMeachul
	public FipgoMeaip

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class

Class CCostPerMeachul
    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectYYYYMM
	public FRectTargetGbn
	public FRectStockPlace
	public FRectShopid
	public FRectDefaultmargin

	public FRectShowShopID
	public FRectShowMakerID

	public FRectMakerid

	public FRectSortBy
	public FRectOrdrBy

	''	public FRectComm_cd
	''
	''	public FRectItemGubun
	''	public FRectItemID
	''	public FRectItemOption
	''	public FRectBarcode
	''	public FRectStartDate
	''	public FRectSearchMode

	public Sub GetCostPerMeachulList
	    Dim sqlStr, i, sqlADD, tmpSqlStr

		sqlADD = " and c.yyyymm = '" & FRectYYYYMM & "' "

		if (FRectTargetGbn <> "") then
			sqlADD = sqlADD & " and c.targetGbn = '" & FRectTargetGbn & "' "
		end if

		if (FRectStockPlace <> "") then
			sqlADD = sqlADD & " and c.stockPlace = '" & FRectStockPlace & "' "
		end if

		if (FRectShopID <> "") then
			sqlADD = sqlADD & " and c.shopid = '" & FRectShopID & "' "
		end if

		if (FRectDefaultmargin <> "") and IsNumeric(FRectDefaultmargin) then
			sqlADD = sqlADD & " and c.defaultmargin = " & FRectDefaultmargin & " "
		end if

		if (FRectMakerid <> "") then
			if (FRectMakerid = "¾øÀ½") then
				sqlADD = sqlADD & " and c.makerid = '' "
			else
				sqlADD = sqlADD & " and c.makerid = '" & FRectMakerid & "' "
			end if
		end if

	    sqlStr = " select count(*) as CNT"
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_monthly_CostPerMeachul c "
	    sqlStr = sqlStr & " where 1=1 "
	    sqlStr = sqlStr & sqlADD

	    db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
        db3_rsget.Close


        sqlStr = " select top " & (FPageSize*FCurrPage) & " "
        sqlStr = sqlStr & " 	yyyymm, "
		sqlStr = sqlStr & " 	targetGbn, "
		sqlStr = sqlStr & " 	stockPlace, "

		if (FRectShowShopID = "Y") then
			sqlStr = sqlStr & " 	shopid, "
		else
			sqlStr = sqlStr & " 	'' as shopid, "
		end if

		if (FRectShowMakerID = "Y") then
			sqlStr = sqlStr & " 	makerid, "
		else
			sqlStr = sqlStr & " 	'' as makerid, "
		end if

		if (FRectMakerid <> "") then
			sqlStr = sqlStr & " 	itemid, "
		else
			sqlStr = sqlStr & " 	'' as itemid, "
		end if

		sqlStr = sqlStr & " 	itemgubun, "
		sqlStr = sqlStr & " 	defaultmargin, "

        sqlStr = sqlStr & " 	sum(buySumPrevMonth) as buySumPrevMonth, "
        sqlStr = sqlStr & " 	sum(buySumThisMonth) as buySumThisMonth, "
        sqlStr = sqlStr & " 	sum(customerMeachul) as customerMeachul, "
        sqlStr = sqlStr & " 	sum(ipgoMeaip) as ipgoMeaip, "
		sqlStr = sqlStr & " 	(case when sum(customerMeachul) = 0 then 0 else 1.0*(sum(buySumPrevMonth)+sum(ipgoMeaip)-sum(buySumThisMonth))/sum(customerMeachul) end) as sortBy1, "
		sqlStr = sqlStr & " 	(case when sum(customerMeachul) = 0 then 0 else sum(customerMeachul)-1.0*(sum(buySumPrevMonth)+sum(ipgoMeaip)-sum(buySumThisMonth)) end) as sortBy2 "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	db_datamart.dbo.tbl_monthly_CostPerMeachul c "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "

		sqlStr = sqlStr & sqlADD

        sqlStr = sqlStr & " group by "
        sqlStr = sqlStr & " 	yyyymm,stockPlace,targetGbn "

		if (FRectShowShopID = "Y") then
			sqlStr = sqlStr & " 	,shopid "
		end if

		if (FRectShowMakerID = "Y") then
			sqlStr = sqlStr & " 	,makerid "
		end if

		if (FRectMakerid <> "") then
			sqlStr = sqlStr & " 	,itemid "
		end if


		sqlStr = sqlStr & " 	,itemgubun "
		sqlStr = sqlStr & " 	,defaultmargin "

		sqlStr = sqlStr & " order by "
		sqlStr = sqlStr & " 	yyyymm "

		if (FRectSortBy <> "") then
			if (FRectSortBy = "sortBy1") then
				sqlStr = sqlStr & " 	,sortBy1 "
				if (FRectOrdrBy = "ordrBy1") then
					sqlStr = sqlStr & " desc "
				end if
			elseif (FRectSortBy = "sortBy2") then
				sqlStr = sqlStr & " 	,sortBy2 "
				if (FRectOrdrBy = "ordrBy2") then
					sqlStr = sqlStr & " desc "
				end if
			end if
		end if

		sqlStr = sqlStr & " 	,stockPlace,targetGbn desc "


		''public FRectSortBy
		''public FRectOrdrBy

		if (FRectShowShopID = "Y") then
			sqlStr = sqlStr & " 	,shopid "
		end if

		if (FRectShowMakerID = "Y") then
			sqlStr = sqlStr & " 	,makerid "
		end if

		if (FRectMakerid <> "") then
			sqlStr = sqlStr & " 	,itemid "
		end if

		sqlStr = sqlStr & " 	,itemgubun "
		sqlStr = sqlStr & " 	,defaultmargin "

		''-- exec [db_datamart].[dbo].[usp_Ten_All_CostPerMeachulSYS] '2016-03', 'ON'
		''-- exec [db_datamart].[dbo].[usp_Ten_All_CostPerMeachulSYS] '2016-03', 'OF'
		''rw sqlStr

        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CCostPerMeachulItem
				FItemList(i).Fyyyymm     		= db3_rsget("yyyymm")
				FItemList(i).FtargetGbn     	= db3_rsget("targetGbn")
				FItemList(i).FstockPlace     	= db3_rsget("stockPlace")
				FItemList(i).Fshopid     		= db3_rsget("shopid")
				FItemList(i).Fitemgubun     	= db3_rsget("itemgubun")
				FItemList(i).Fitemid     		= db3_rsget("itemid")
				FItemList(i).Fmakerid	     	= db3_rsget("makerid")
				FItemList(i).Fdefaultmargin    	= db3_rsget("defaultmargin")

				FItemList(i).FbuySumPrevMonth   = db3_rsget("buySumPrevMonth")
				FItemList(i).FbuySumThisMonth   = db3_rsget("buySumThisMonth")
				FItemList(i).FcustomerMeachul   = db3_rsget("customerMeachul")
				FItemList(i).FipgoMeaip     	= db3_rsget("ipgoMeaip")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    End Sub

    Private Sub Class_Initialize()
        redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
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

End Class

%>
