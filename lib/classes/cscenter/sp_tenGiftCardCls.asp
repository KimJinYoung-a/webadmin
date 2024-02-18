<%
''예치금 관련.
Class CTenGiftCardLogItem
    public Fidx
    public Fuserid
    public FuseCash
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

Class CTenGiftCard
    public FItemList()
    public FOneItem
    public FRectUserID

    public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FcurrentCash
	public FgainCash
	public FspendCash
	public FrefundCash

	public Sub getTenGiftCardLog
	    dim i, sqlStr

	    FTotalCount = 0
	    FResultCount = 0

        sqlStr = "exec [db_user].[dbo].sp_Ten_UserTenGiftCardLogCnt '"& FRectUserID & "'"
        rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    if Not (Rsget.Eof) then
	        FTotalCount = rsget("CNT")
	    end if
	    rsget.Close

	    sqlStr = "exec [db_user].[dbo].sp_Ten_UserTenGiftCardLog "&FPageSize&","&FCurrPage&",'"& FRectUserID & "'"
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly


		FResultCount = rsget.RecordCount
	    if (FResultCount<1) then FResultCount=0
		FTotalPage = CInt(FTotalCount\FPageSize) + 1

	    redim preserve FItemList(FResultCount)
	    i = 0

	    if Not (Rsget.Eof) then
	        do until rsget.eof
	            set FItemList(i) = new CTenGiftCardLogItem
    	        FItemList(i).Fidx         = rsget("idx")
                FItemList(i).Fuserid       = rsget("userid")
                FItemList(i).FuseCash      = rsget("useCash")
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

    public Sub getUserCurrentTenGiftCard
        dim mile,sqlStr
		if (FRectUserID="") then exit sub

        sqlStr = "exec [db_user].[dbo].sp_Ten_UserCurrentTenGiftCard '" & FRectUserID & "'"

    	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    	if Not (Rsget.Eof) then
    	    FcurrentCash = rsget("currentCash")
    	    FgainCash    = rsget("gainCash")
    	    FspendCash   = rsget("spendCash")
    	    FrefundCash  = rsget("refundCash")
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

        FcurrentCash = 0
        FgainCash    = 0
        FspendCash   = 0

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