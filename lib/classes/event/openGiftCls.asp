<%
''openGiftCls.asp

Class CopenGiftItem
    public Fevent_code
    public FfrontOpen
    public FopenImage1
    public FopenHtml
    public FopenHtmlWeb
    public Freguser
    public Fregdate

    public Fevent_name
    public Fevt_startdate
    public Fevt_enddate
    public Fevt_state
    public Fevt_using
    public FopengiftType	''(구분: 전체(1), 다이어리(9))
    public FopengiftScope   ''(범위: 전체(1), 모바일(3), APP(5))

    public FGiftCNT
    public FALLGiftCNT

    function getOpengiftTypeName()
        if (FopengiftType=9) then
            getOpengiftTypeName = "다이어리"
        elseif (FopengiftType=1) then
            getOpengiftTypeName = "전체"
        end if
    end function

    function getOpengiftScopeName()
        if (FopengiftScope=1) or (FopengiftScope="") then
            getOpengiftScopeName = "전체"
        elseif (FopengiftScope=3) then
            getOpengiftScopeName = "모바일"
        elseif (FopengiftScope=5) then
            getOpengiftScopeName = "APP"
        end if
    end function

    public function getEventStateName()
        dim StateDesc
        if Fevt_state="9" then
            StateDesc = "종료"
        elseif Fevt_state="7" then
            StateDesc = "오픈"
        end if

        IF Fevt_state = "7" AND datediff("d",Fevt_startdate,date()) >= 0 and datediff("d",Fevt_enddate,date()) <=0 THEN
			getEventStateName = "오픈"
		ELSEIF Fevt_state ="7" AND datediff("d",Fevt_enddate,date()) > 0 THEN
			getEventStateName = "종료"
		ELSE
			getEventStateName = StateDesc
		END IF

    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class CopenGift
    public FItemList()
    public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectEventCode
	public FRectGiftKindCode

	public function getOpenGiftList()
	    dim sqlStr, i
	    sqlStr = "select count(*) as CNT"
	    sqlStr = sqlStr & " from db_event.dbo.tbl_openGift O"
	    sqlStr = sqlStr & "     Join db_event.dbo.tbl_Event E"
	    sqlStr = sqlStr & "     on O.event_code=E.evt_code"
	    sqlStr = sqlStr & " where 1=1"

	    rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("CNT")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + ""
		sqlStr = sqlStr & " O.*, E.evt_name, E.evt_startdate, E.evt_enddate, E.evt_state, E.evt_using"
		sqlStr = sqlStr & " ,(select count(*) from db_event.dbo.tbl_gift G where G.evt_code=O.event_code) as GiftCNT"
		sqlStr = sqlStr & " ,(select count(*) from db_event.dbo.tbl_gift G where G.evt_code=O.event_code and G.gift_scope=O.opengiftType) as ALLGiftCNT"
	    sqlStr = sqlStr & " from db_event.dbo.tbl_openGift O"
	    sqlStr = sqlStr & "     Join db_event.dbo.tbl_Event E"
	    sqlStr = sqlStr & "     on O.event_code=E.evt_code"
	    sqlStr = sqlStr & " where 1=1"
	    sqlStr = sqlStr & " order by O.event_code desc"
'rw 	sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CopenGiftItem

				FItemList(i).Fevent_code    = rsget("event_code")
    			FItemList(i).FfrontOpen     = rsget("frontOpen")
    			FItemList(i).FopenImage1    = rsget("openImage1")
    			FItemList(i).FopenHtml      = rsget("openHtml")
    			FItemList(i).FopenHtmlWeb   = rsget("openHtmlWeb")
    			FItemList(i).Freguser       = rsget("reguser")
    			FItemList(i).Fregdate       = rsget("regdate")

    			FItemList(i).Fevent_name    = db2Html(rsget("evt_name"))
    			FItemList(i).Fevt_startdate = rsget("evt_startdate")
                FItemList(i).Fevt_enddate   = rsget("evt_enddate")
                FItemList(i).Fevt_state     = rsget("evt_state")
                FItemList(i).Fevt_using     = rsget("evt_using")

                FItemList(i).FGiftCNT       = rsget("GiftCNT")
                FItemList(i).FALLGiftCNT    = rsget("ALLGiftCNT")
                FItemList(i).FopengiftType  = rsget("opengiftType")
                FItemList(i).FopengiftScope = rsget("opengiftScope")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

    end function

	' /admin/shopmaster/gift/giftuserdetail.asp
	public function getOneOpenGift()
	    dim sqlStr, i

	    sqlStr = "select O.*, E.evt_name, E.evt_startdate, E.evt_enddate, E.evt_state, E.evt_using "
	    sqlStr = sqlStr & " from db_event.dbo.tbl_openGift O"
	    sqlStr = sqlStr & "     Join db_event.dbo.tbl_Event E"
	    sqlStr = sqlStr & "     on O.event_code=E.evt_code"
	    sqlStr = sqlStr & " where event_code=" & FRectEventCode

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordCount
		FResultCount = rsget.recordCount

		if Not rsget.Eof then
			set FOneItem = new CopenGiftItem

			FOneItem.Fevent_code    = rsget("event_code")
			FOneItem.FfrontOpen     = rsget("frontOpen")
			FOneItem.FopenImage1    = rsget("openImage1")
			FOneItem.FopenHtml      = rsget("openHtml")
			FOneItem.FopenHtmlWeb   = rsget("openHtmlWeb")
			FOneItem.Freguser       = rsget("reguser")
			FOneItem.Fregdate       = rsget("regdate")
			FOneItem.FopengiftType  = rsget("opengiftType")
			FOneItem.FopengiftScope = rsget("opengiftScope")

			if IsNULL(FOneItem.FopengiftType) then FOneItem.FopengiftType=1
			i=i+1
			rsget.movenext
		end if
		rsget.Close

    end function


    public function IsOpenGiftUsingGiftKind()
        dim sqlStr
        IsOpenGiftUsingGiftKind = false

        sqlStr = "select count(*) as CNT from db_event.dbo.tbl_gift g"
        sqlStr = sqlStr & " Join db_event.dbo.tbl_openGift O"
        sqlStr = sqlStr & " on g.evt_code=O.event_code"
        sqlStr = sqlStr & " where giftkind_code="&FRectGiftKindCode&""

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            IsOpenGiftUsingGiftKind = true
        end if
        rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
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
End Class
%>