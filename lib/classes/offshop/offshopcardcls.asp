<%

Class COffShopCardPromotionItem
	public Fidx
	public Fshopid
	public FcardPrice
	public FstartDate
	public FendDate
	public FrateGubun
	public FrateAmmount
	public Fisusing
	public Fregdate

	public function getRateGubunName()
		select case CStr(FrateGubun)
			case "1"
				getRateGubunName = "Á¤·ü"
			case "2"
				getRateGubunName = "Á¤¾×"
			case default
				getRateGubunName = FrateGubun
		end select
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopCardPromotion
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectShopID
	public FRectStartDay
	public FRectEndDay

	public FRectIdx
	public FRectCardPrice

	public sub COffShopCardPromotionList()
		dim i, sqlStr, addStr

		addStr = ""
		addStr = addStr + " and c.isusing = 'Y' "

		if (FRectShopid <> "") then
			addStr = addStr + " and c.shopid = '" & FRectShopid & "' "
		end if

		if (FRectCardPrice <> "") then
			addStr = addStr + " and c.cardPrice = " & FRectCardPrice & " "
		end if


		'// ====================================================================
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].[tbl_shop_card_promotion] c "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + addStr

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		'// ====================================================================
		sqlStr = " select top " & CStr(FPageSize*FCurrPage) & " c.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].[tbl_shop_card_promotion] c "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + addStr
		sqlStr = sqlStr + " order by idx desc "

		''response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopCardPromotionItem

				FItemList(i).Fidx          		= rsget("idx")
				FItemList(i).Fshopid          	= rsget("shopid")
				FItemList(i).FcardPrice         = rsget("cardPrice")
				FItemList(i).FstartDate         = rsget("startDate")
				FItemList(i).FendDate          	= rsget("endDate")
				FItemList(i).FrateGubun         = rsget("rateGubun")
				FItemList(i).FrateAmmount       = rsget("rateAmmount")
				FItemList(i).Fisusing          	= rsget("isusing")
				FItemList(i).Fregdate          	= rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end sub

	public sub getOneCardPromotion()
		dim i,sqlStr, addStr

		addStr = " and c.idx = " & FRectIdx

		sqlStr = " select top 1 c.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].[tbl_shop_card_promotion] c "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + addStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		set FOneItem = new COffShopCardPromotionItem
		if  not rsget.EOF  then

			FOneItem.Fidx          		= rsget("idx")
			FOneItem.Fshopid          	= rsget("shopid")
			FOneItem.FcardPrice         = rsget("cardPrice")
			FOneItem.FstartDate         = rsget("startDate")
			FOneItem.FendDate          	= rsget("endDate")
			FOneItem.FrateGubun         = rsget("rateGubun")
			FOneItem.FrateAmmount       = rsget("rateAmmount")
			FOneItem.Fisusing          	= rsget("isusing")
			FOneItem.Fregdate          	= rsget("regdate")

		end if
		rsget.close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

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

End Class
%>
