<%
Class CCardSaleItem
	Public FIdx
	Public FStartdate
	Public FEnddate
	Public FCardCode
	Public FSaleType
	Public FSalePrice
	Public FMinPrice
	Public FMaxPrice
	Public FIsUsing
	Public FRegdate
	Public FRegUserid
	Public FCardGroupName
	Public FCardName
	Public FbannerTitle
	Public FbannerView
	Public Fbgcolor
	Public FblnWeb
	Public FblnMobile
	Public FblnApp

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CCardSale
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FOneItem

	Public FRectIdx
	Public FRectIsusing

	Public Sub getCardSaleItemList
		Dim i, sqlStr, addSql

		If FRectIsusing <> "" Then
			addSql = addSql & " and s.isusing = '"& FRectIsusing &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(s.idx) as cnt, CEILING(CAST(Count(s.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_card_sale] as s "
		sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_card_sale_code] as c on s.cardCode = c.cardCode "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " s.idx, s.startdate, s.enddate, s.cardCode, s.saleType, s.salePrice, s.minPrice "
		sqlStr = sqlStr & " , s.maxPrice, s.isUsing, s.regdate, s.regUserid, c.cardGroupName, c.cardName "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_card_sale] as s "
		sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_card_sale_code] as c on s.cardCode = c.cardCode "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY s.idx DESC "

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CCardSaleItem
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FStartdate			= rsget("startdate")
					FItemList(i).FEnddate			= rsget("enddate")
					FItemList(i).FCardCode			= rsget("cardCode")
					FItemList(i).FSaleType			= rsget("saleType")
					FItemList(i).FSalePrice			= rsget("salePrice")
					FItemList(i).FMinPrice			= rsget("minPrice")
					FItemList(i).FMaxPrice			= rsget("maxPrice")
					FItemList(i).FIsUsing			= rsget("isUsing")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FCardGroupName		= rsget("cardGroupName")
					FItemList(i).FCardName			= rsget("cardName")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getCardSaleOneItem
	    Dim i, sqlStr, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 s.idx, s.startdate, s.enddate, s.cardCode, s.saleType, s.salePrice, s.minPrice"
		sqlStr = sqlStr & " , s.maxPrice, s.isUsing, s.regdate, s.regUserid, c.cardGroupName, c.cardName, s.bannerTitle, s.bannerView"
		sqlStr = sqlStr & " , s.bgcolor, s.blnWeb, s.blnMobile, s.blnApp"
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_card_sale] as s"
		sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_card_sale_code] as c on s.cardCode = c.cardCode"
		sqlStr = sqlStr & " WHERE s.idx = '"& FRectIdx &"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		set FOneItem = new CCardSaleItem

		if  not rsget.EOF  then
			FOneItem.FIdx				= rsget("idx")
			FOneItem.FStartdate			= rsget("startdate")
			FOneItem.FEnddate			= rsget("enddate")
			FOneItem.FCardCode			= rsget("cardCode")
			FOneItem.FSaleType			= rsget("saleType")
			FOneItem.FSalePrice			= rsget("salePrice")
			FOneItem.FMinPrice			= rsget("minPrice")
			FOneItem.FMaxPrice			= rsget("maxPrice")
			FOneItem.FIsUsing			= rsget("isUsing")
			FOneItem.FRegdate			= rsget("regdate")
			FOneItem.FRegUserid			= rsget("regUserid")
			FOneItem.FCardGroupName		= rsget("cardGroupName")
			FOneItem.FCardName			= rsget("cardName")
			FOneItem.FbannerTitle		= rsget("bannerTitle")
			FOneItem.FbannerView		= rsget("bannerView")
			FOneItem.Fbgcolor			= rsget("bgcolor")
			FOneItem.FblnWeb			= rsget("blnWeb")
			FOneItem.FblnMobile			= rsget("blnMobile")
			FOneItem.FblnApp			= rsget("blnApp")
		end if
		rsget.Close
	End Sub

	Public Function fnCardList
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT cardCode, cardGroupName, cardName "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_card_sale_code] "
		sqlStr = sqlStr & " ORDER BY cardCode ASC, cardGroupName ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			fnCardList = rsget.getRows()
		End If
		rsget.Close
	End Function


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
End Class
%>