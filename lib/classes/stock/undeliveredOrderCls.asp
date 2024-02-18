<%

Class COrderItem
	public Forderserial
	public Fipkumdate
	public Fmakerid
	public Fitemid
	public Fitemoption
	public Ficnt
	public Frealstock
	public Fsellcash
	public Fbuycash
	public Fpreordernofix
	public Fgubun
	public FisDanjong
	public FrackcodeByOption
	public FsubRackcodeByOption
	public Fitemname
	public Foptionname

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class COrder
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectDate
    public FRectDate2
	public FRectMode
	public FRectStat
	public FRectDanjong


	public Sub OrderList()
		dim i, sqlStr

		sqlStr = "exec db_storage.dbo.usp_logics_order_undeliveredlist_get_V2 '" & FRectDate & "', '" & FRectDate2 & "', '" & FRectMode & "', '" & FRectStat & "', '" & FRectDanjong & "' "
		''response.write sqlStr: response.end

		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open sqlStr, dbget

		FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COrderItem

			If FRectMode="OD" then
				FItemList(i).Forderserial 			= rsget("orderserial")
				FItemList(i).Fipkumdate 			= rsget("ipkumdate")
			End if

			FItemList(i).Fmakerid 				= rsget("makerid")
			FItemList(i).Fitemid 				= rsget("itemid")
			FItemList(i).Fitemoption 			= rsget("itemoption")
			FItemList(i).Ficnt 					= rsget("icnt")
			FItemList(i).Frealstock 			= rsget("realstock")
			FItemList(i).Fsellcash 				= rsget("sellcash")
			FItemList(i).Fbuycash 				= rsget("buycash")
			FItemList(i).Fpreordernofix 		= rsget("preordernofix")
			FItemList(i).Fgubun 				= rsget("gubun")
			FItemList(i).FisDanjong 			= rsget("isDanjong")
			FItemList(i).FrackcodeByOption 		= rsget("rackcodeByOption")
			FItemList(i).FsubRackcodeByOption 	= rsget("subRackcodeByOption")
			FItemList(i).Fitemname 				= rsget("itemname")
			FItemList(i).Foptionname 			= rsget("optionname")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
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
