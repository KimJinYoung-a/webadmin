<%

Class CCSCenterDepositSummaryItem
	public Fcurrentdeposit
	public Fgaindeposit
	public Fspenddeposit

	Private Sub Class_Initialize()
		Fcurrentdeposit = 0
		Fgaindeposit = 0
		Fspenddeposit = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCSCenterDepositItem
	public Fidx
	public Fuserid
	public Fdeposit
	public Fjukyocd
	public Fjukyo
	public Fregdate
	public Forderserial
	public Fdeleteyn

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCSCenterDeposit
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectUserID
	public FRectDeleteYn
	public FRectExpireDate

	public Sub GetCSCenterDepositSummary()
		dim i,sqlStr

		sqlStr = "select top 1 d.currentdeposit, d.gaindeposit, d.spenddeposit "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_current_deposit d "
		sqlStr = sqlStr + " where d.userid='" + FRectUserID + "'"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		set FOneItem = new CCSCenterDepositSummaryItem

		if  not rsget.EOF  then
			FOneItem.Fcurrentdeposit	= rsget("currentdeposit")
			FOneItem.Fgaindeposit		= rsget("gaindeposit")
			FOneItem.Fspenddeposit		= rsget("spenddeposit")
		end if
		rsget.close
	end sub

	public Sub GetCSCenterDepositList()
		dim i,sqlStr

		sqlStr = " select top 500 d.idx, d.userid, d.deposit, d.jukyocd, d.jukyo, d.regdate, d.orderserial, d.deleteyn "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_depositlog d "
		sqlStr = sqlStr + " where d.userid='" + CStr(FRectUserID) + "' "
		if (FRectDeleteYn<>"") then
			sqlStr = sqlStr + " and d.deleteyn='" + CStr(FRectDeleteYn) + "' "
		end if
		sqlStr = sqlStr + " order by d.regdate desc "

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCSCenterDepositItem

				FItemList(i).Fidx                = rsget("idx")
				FItemList(i).Fuserid            = rsget("userid")
				FItemList(i).Fdeposit           = rsget("deposit")
				FItemList(i).Fjukyocd           = rsget("jukyocd")
				FItemList(i).Fjukyo             = rsget("jukyo")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Forderserial       = rsget("orderserial")
				FItemList(i).Fdeleteyn          = rsget("deleteyn")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		FCurrPage       = 1
		FPageSize       = 20
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0
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

end Class

%>
