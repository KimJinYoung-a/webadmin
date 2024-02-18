<%
Class CCardPartialCancelItem

    public FclogIdx
    public Forderserial
    public Forgtid
    public Fcancelprice
    public Frepayprice
    public Fusermail
    public Fnewtid
    public Fresultcode
    public Fresultmsg
    public Fcancelrequestcount
    public Fregdate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CCardPartialCancel
    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectOrderSerial

	public Sub getCardCancelList
	    Dim sqlStr, i
	    sqlStr = "select top 100 * from db_order.dbo.tbl_card_cancel_log "
	    sqlStr = sqlStr & " where orderserial='"&FRectOrderSerial&"'"
	    sqlStr = sqlStr & " and resultCode in ('00','2001') "
	    sqlStr = sqlStr & " order by clogIdx"
	    rsget.Open sqlStr, dbget, 1

	    FResultCount = rsget.RecordCount
        FTotalCount  = FResultCount

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCardPartialCancelItem
				FItemList(i).FclogIdx              = rsget("clogIdx")
                FItemList(i).Forderserial          = rsget("orderserial")
                FItemList(i).Forgtid               = rsget("orgtid")
                FItemList(i).Fcancelprice          = rsget("cancelprice")
                FItemList(i).Frepayprice           = rsget("repayprice")
                FItemList(i).Fusermail             = rsget("usermail")
                FItemList(i).Fnewtid               = rsget("newtid")
                FItemList(i).Fresultcode           = rsget("resultcode")
                FItemList(i).Fresultmsg            = rsget("resultmsg")
                FItemList(i).Fcancelrequestcount   = rsget("cancelrequestcount")
                FItemList(i).Fregdate              = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
	    rsget.Close

    end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

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
