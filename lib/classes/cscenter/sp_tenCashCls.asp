<%
''예치금 관련.
Class CTenCashLogItem
    public Fidx
    public Fuserid
    public Fdeposit
    public Fjukyocd
    public Fjukyo
    public Forderserial
    public Fdeleteyn
    public Freguserid
    public Fdeluserid
    public Fregdate
    public Fremain

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CTenCash
    public FItemList()
    public FOneItem
    public FRectUserID

    public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public Fcurrentdeposit
	public Fgaindeposit
	public Fspenddeposit

	public FRectShowDelete

	public Sub getTenCashLog
	    dim i, sqlStr

	    FTotalCount = 0
	    FResultCount = 0

        sqlStr = "exec [db_user].[dbo].sp_Ten_UserTenCashLogCnt '"& FRectUserID & "', '" + CStr(FRectShowDelete) + "' "
		rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	    if Not (Rsget.Eof) then
	        FTotalCount = rsget("CNT")
	    end if
	    rsget.Close

	    sqlStr = "exec [db_user].[dbo].sp_Ten_UserTenCashLog "&FPageSize&","&FCurrPage&",'"& FRectUserID & "', '" + CStr(FRectShowDelete) + "' "
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly


		FResultCount = rsget.RecordCount
	    if (FResultCount<1) then FResultCount=0
		FTotalPage = CInt(FTotalCount\FPageSize) + 1

	    redim preserve FItemList(FResultCount)
	    i = 0

	    if Not (Rsget.Eof) then
	        do until rsget.eof
	            set FItemList(i) = new CTenCashLogItem
    	        FItemList(i).Fidx         = rsget("idx")
                FItemList(i).Fuserid       = rsget("userid")
                FItemList(i).Fdeposit      = rsget("deposit")
                FItemList(i).Fjukyocd      = rsget("jukyocd")
                FItemList(i).Fjukyo        = rsget("jukyo")
                FItemList(i).Forderserial  = rsget("orderserial")
                FItemList(i).Fdeleteyn     = rsget("deleteyn")
                ''FItemList(i).Freguserid    = rsget("reguserid")
                FItemList(i).Fdeluserid    = rsget("deluserid")
                FItemList(i).Fregdate      = rsget("regdate")
                FItemList(i).Fremain        = rsget("remain")
                i=i+1
				rsget.moveNext

            loop
	    end if
    	rsget.Close
    end Sub

    public Sub getUserCurrentTenCash
        dim mile,sqlStr
		if (FRectUserID="") then exit sub

        sqlStr = "exec [db_user].[dbo].sp_Ten_UserCurrentTenCash '" & FRectUserID & "'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    	if Not (Rsget.Eof) then
    	    Fcurrentdeposit = rsget("currentdeposit")
    	    Fgaindeposit    = rsget("gaindeposit")
    	    Fspenddeposit   = rsget("spenddeposit")
    	end if
    	rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

        Fcurrentdeposit = 0
        Fgaindeposit    = 0
        Fspenddeposit   = 0

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
