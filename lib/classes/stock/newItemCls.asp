<%

Class CNewItemItem
	public Fpurchasetype
	public Fmakerid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fipgodate
	public FsellSTDate
	public Fipgocnt
	public Fonsellcnt
	public Foffsellcnt

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end Class

Class CNewItem
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectStartDT
	public FRectEndDT
	public FRectPurchaseType
	public FRectMWDiv
	public FRectMakerID

	public Sub GetNewItemList()
		dim i, sqlStr

		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_GetNewItemList] '" & FRectPurchaseType & "', '" & FRectStartDT & "', '" & FRectEndDT & "', '" & FRectMWDiv & "', '" & FRectMakerID & "' "
		''response.write sqlStr

		db3_rsget.CursorLocation = 3
		db3_rsget.pagesize = 1000
		db3_rsget.Open sqlStr, db3_dbget, 3, 1

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CNewItemItem

			FItemList(i).Fpurchasetype 	= db3_rsget("purchasetype")
			FItemList(i).Fmakerid 		= db3_rsget("makerid")
			FItemList(i).Fitemgubun 	= db3_rsget("itemgubun")
			FItemList(i).Fitemid 		= db3_rsget("itemid")
			FItemList(i).Fitemoption 	= db3_rsget("itemoption")
			FItemList(i).Fitemname 			= db2Html(db3_rsget("itemname"))
			FItemList(i).Fitemoptionname 	= db2Html(db3_rsget("itemoptionname"))
			FItemList(i).Fipgodate 		= db3_rsget("ipgodate")
			FItemList(i).FsellSTDate 	= db3_rsget("sellSTDate")
			FItemList(i).Fipgocnt 		= db3_rsget("ipgocnt")
			FItemList(i).Fonsellcnt 	= db3_rsget("onsellcnt")
			FItemList(i).Foffsellcnt 	= db3_rsget("offsellcnt")

			db3_rsget.movenext
			i=i+1
		loop
		db3_rsget.Close
	end Sub

	Private Sub Class_Initialize()
		FCurrPage = 1
		FPageSize = 200
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
