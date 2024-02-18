<%

Class CTicketItemDetail
    public Fitemid
    public FstDt
    public FedDt
    public FbookingStDt
    public FbookingEdDt
    public FbookingCharge
    public FticketDlvType
    public FrefundInfoType
    
    public FticketPlaceIdx
    public FticketPlaceName
    public FtPAddress
    public FtPTel
    public FtPHomeURL
    public FplaceLinkURL
    public FplaceImgURL
    public FplacecontentsImage1
    public FplacecontentsImage2
    public FplacecontentsImage3
    public FplaceContents
    
    public FtxplayTimInfo
    public FtxGenre
    public FtxGrade
    public FtxRunTime
    
    public Function getTicketDlvName
		Select Case CStr(FticketDlvType)
		Case "1"
			getTicketDlvName = "현장수령"
		Case "2"
			getTicketDlvName = "일반배송"
		Case "3"
			getTicketDlvName = "현장수령 or 일반배송 택일"
		Case "9"
			getTicketDlvName = "티켓 현장수령 및 사은품 상품은 배송"
		Case else
			getTicketDlvName = "현장수령"
		End select
	End function
	
    Private Sub Class_Initialize()
        FbookingCharge = 0
	End Sub

	Private Sub Class_Terminate()

    End Sub
    
End Class

Class CTicketItem
    public FOneItem
	public FItemList()
	
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectItemID
	
	public Sub GetOneTicketItem
	    dim sqlStr 
	    sqlStr = " exec [db_item].[dbo].[sp_Ten_getOneTicketItem] "&FRectItemID
	    
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    if Not (Rsget.Eof) then
	        FTotalCount = rsget.RecordCount
	        FResultCount = FTotalCount
	        
	        set FOneItem = new CTicketItemDetail
	        FOneItem.Fitemid                = rsget("itemid")
            FOneItem.FstDt                  = rsget("stDt")
            FOneItem.FedDt                  = rsget("edDt")
            FOneItem.FbookingStDt           = rsget("bookingStDt")
            FOneItem.FbookingEdDt           = rsget("bookingEdDt")
            FOneItem.FbookingCharge         = rsget("bookingCharge")            '' money
            FOneItem.FticketDlvType         = rsget("ticketDlvType")            '' int 
            FOneItem.FrefundInfoType        = rsget("refundInfoType")           '' 반품관련 규정 key
            
            FOneItem.FticketPlaceIdx        = rsget("ticketPlaceIdx")           '' 공연장 key
            FOneItem.FticketPlaceName       = db2HTML(rsget("ticketPlaceName"))
            FOneItem.FtPAddress             = db2HTML(rsget("tPAddress"))
            FOneItem.FtPTel                 = rsget("tPTel")
            FOneItem.FtPHomeURL             = db2HTML(rsget("tPHomeURL"))
            FOneItem.FplaceLinkURL          = db2HTML(rsget("placeLinkURL"))
            FOneItem.FplaceImgURL           = db2HTML(rsget("placeImgURL"))
            
            FOneItem.FplaceImgURL           = db2HTML(rsget("placeImgURL"))
            FOneItem.FplacecontentsImage1   = db2HTML(rsget("placecontentsImage1"))
            FOneItem.FplacecontentsImage2   = db2HTML(rsget("placecontentsImage2"))
            FOneItem.FplacecontentsImage3   = db2HTML(rsget("placecontentsImage3"))
            FOneItem.FplaceContents         = db2HTML(rsget("placeContents"))
            
            FOneItem.FtxplayTimInfo         = db2HTML(rsget("txplayTimInfo"))
            FOneItem.FtxGenre               = db2HTML(rsget("txGenre"))
            FOneItem.FtxGrade               = db2HTML(rsget("txGrade"))
            FOneItem.FtxRunTime             = db2HTML(rsget("txRunTime"))
        else
            set FOneItem = new CTicketItemDetail
	    end if
	    rsget.Close
    end Sub
    
    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class

Class CTicketPlaceDetail
    public FticketPlaceIdx
    public FticketPlaceName
    public FtPAddress
    public FtPTel
    public FtPHomeURL
    public FplaceLinkURL
    public FplaceImgURL                    ''약도 image
    public FplacecontentsImage1
    public FplacecontentsImage2
    public FplacecontentsImage3
    public FplaceContents
	public FparkingGuide
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

    End Sub
    
end Class

Class CTicketPlace
    public FOneItem
	public FItemList()
	
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FsortMethod
	
	public FRectTicketPlaceIdx
	public FRectTicketPlaceName
	
	public Sub GetOneTicketPLace
	    dim sqlStr
	    sqlStr = "exec db_item.dbo.sp_Ten_getOneTicketPLace "&FRectTicketPlaceIdx&""
	    
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    if Not (Rsget.Eof) then
	        FTotalCount = rsget.RecordCount
	        FResultCount = FTotalCount
	        
	        set FOneItem = new CTicketPlaceDetail
            
            FOneItem.FticketPlaceIdx        = rsget("ticketPlaceIdx")           '' 공연장 key
            FOneItem.FticketPlaceName       = db2HTML(rsget("ticketPlaceName"))
            FOneItem.FtPAddress             = db2HTML(rsget("tPAddress"))
            FOneItem.FtPTel                 = rsget("tPTel")
            FOneItem.FtPHomeURL             = db2HTML(rsget("tPHomeURL"))
            FOneItem.FplaceLinkURL          = db2HTML(rsget("placeLinkURL"))
            
            FOneItem.FplaceImgURL           = db2HTML(rsget("placeImgURL"))
            FOneItem.FplacecontentsImage1   = db2HTML(rsget("placecontentsImage1"))
            FOneItem.FplacecontentsImage2   = db2HTML(rsget("placecontentsImage2"))
            FOneItem.FplacecontentsImage3   = db2HTML(rsget("placecontentsImage3"))
            FOneItem.FplaceContents         = db2HTML(rsget("placeContents"))

			FOneItem.FparkingGuide         	= db2HTML(rsget("parkingGuide"))

        else
            set FOneItem = new CTicketPlaceDetail
	    end if
	    rsget.Close
	    
    end Sub
	
	public Sub getTicketPLaceList
	    dim sqlStr,i
	    sqlStr = " select count(*) as CNT  "
        sqlStr = sqlStr & " from db_item.dbo.tbl_ticket_placeInfo"
        sqlStr = sqlStr & " where ticketPlaceName like '" & FRectTicketPlaceName & "%'"
	    
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    if Not (Rsget.Eof) then
	        FTotalCount = rsget("CNT")
	    end if
	    rsget.Close
	    
	    sqlStr = " select top "& CStr(FCurrPage*FPageSize) & " * "
	    sqlStr = sqlStr & " from db_item.dbo.tbl_ticket_placeInfo"
	    sqlStr = sqlStr & " where ticketPlaceName like '" & FRectTicketPlaceName & "%'"
	    sqlStr = sqlStr & " order by ticketPlaceIdx " & chkIIF(FsortMethod="r","asc","desc")
	    
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.pagesize = FPageSize
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
	    redim preserve FItemList(FResultCount)
	    i = 0
	    
	    if Not (Rsget.Eof) then
	        rsget.absolutepage = FCurrPage
	        do until rsget.eof
	            set FItemList(i) = new CTicketPlaceDetail
    	        FItemList(i).FticketPlaceIdx         = rsget("ticketPlaceIdx")
                FItemList(i).FticketPlaceName        = rsget("ticketPlaceName")
                FItemList(i).FtPAddress              = rsget("tPAddress")
                FItemList(i).FtPTel                  = rsget("tPTel")
                FItemList(i).FtPHomeURL              = rsget("tPHomeURL")
                FItemList(i).FplaceLinkURL           = rsget("placeLinkURL")
                FItemList(i).FplaceImgURL            = rsget("placeImgURL")
                FItemList(i).FplacecontentsImage1    = rsget("placecontentsImage1")
                FItemList(i).FplacecontentsImage2    = rsget("placecontentsImage2")
                FItemList(i).FplacecontentsImage3    = rsget("placecontentsImage3")
                FItemList(i).FplaceContents          = rsget("placeContents")

                i=i+1
				rsget.moveNext
				
            loop
	    end if
    	rsget.Close
    end Sub

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
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

    End Sub
end Class

Class CTicketScheduleDetail
    public FTk_itemid
    public FTk_itemoption
    public FTk_optName
    public FTk_StSchedule
    public FTk_EdSchedule
    public FreturnExpireDate
    
    public function getScheduleDateStr()
        If IsNULL(FTk_StSchedule) then Exit Function
        
        If IsNULL(FTk_EdSchedule) then 
            getScheduleDateStr = Left(FTk_StSchedule,10)
            Exit Function
        end if
        
        If (Left(FTk_StSchedule,10)=Left(FTk_EdSchedule,10)) then 
            getScheduleDateStr = Left(FTk_StSchedule,10)
            Exit Function
        end if
        
        getScheduleDateStr = Left(FTk_StSchedule,10) & "~" & Left(FTk_EdSchedule,10)
        
    end function
    
    public function getScheduleDateTime()
        If IsNULL(FTk_StSchedule) then Exit Function
        
        if (Left(Right(FTk_StSchedule,8),5)="00:00") then
            getScheduleDateTime = "-"
            Exit Function
        end if
        
        getScheduleDateTime = Left(Right(FTk_StSchedule,8),5)
    end function

    Private Sub Class_Initialize()
		
	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class

Class CTicketSchedule
    public FOneItem
	public FItemList()
	
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectItemID
	public FRectItemOption
	
	public Sub getOneTicketSchdule()
	    dim sqlStr
	    sqlStr = "exec db_item.dbo.sp_Ten_getOneTicketSchedule "&FRectItemID&",'"&FRectItemOption&"'"
	    
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    
	    FTotalCount = rsget.RecordCount
	    FResultCount = FTotalCount
	    
	    set FOneItem = new CticketScheduleDetail
	    if Not (Rsget.Eof) then
                
            FOneItem.FTk_itemid        = rsget("Tk_itemid") 
            FOneItem.FTk_itemoption    = rsget("Tk_itemoption") 
            FOneItem.FTk_optName       = db2HTML(rsget("Tk_optName"))
            FOneItem.FTk_StSchedule    = rsget("Tk_StSchedule") 
            FOneItem.FTk_EdSchedule    = rsget("Tk_EdSchedule") 
            FOneItem.FreturnExpireDate = rsget("returnExpireDate") 
	    end IF
	    
	    rsget.Close
    end Sub
    
    
    public Sub getTicketSchduleList()
	    dim sqlStr,i
	    sqlStr = "exec db_item.dbo.sp_Ten_getTicketScheduleList "&FRectItemID&""
	    
	    rsget.CursorLocation = adUseClient                              ''' require RecordCount
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	    
	    FTotalCount = rsget.RecordCount
	    FResultCount = FTotalCount
	    
	    if (FResultCount<1) then FResultCount=0
	    redim preserve FItemList(FResultCount)
	    i = 0
	    
	    if Not (Rsget.Eof) then
	        do until rsget.eof
    	        set FItemList(i) = new CticketScheduleDetail
                
                FItemList(i).FTk_itemid        = rsget("Tk_itemid") 
                FItemList(i).FTk_itemoption    = rsget("Tk_itemoption") 
                FItemList(i).FTk_optName       = db2HTML(rsget("Tk_optName"))
                FItemList(i).FTk_StSchedule    = rsget("Tk_StSchedule") 
                FItemList(i).FTk_EdSchedule    = rsget("Tk_EdSchedule") 
                FItemList(i).FreturnExpireDate = rsget("returnExpireDate") 

                
                rsget.MoveNext
                i=i+1
            loop
	    end if
	    rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class
%>