<%
class COutBrandItem
    public Fmakerid
    public Fsubidx
    public FoutbrandStatus
    public FbatchmakeDate
    public FoutScheduledate
    public FpreSellitemNo
    public FpreRegedItemno
    public FpreSellitemNoSum
    public FpreSellCostSum
    public Fpreday
    public Freguser
    public Fregdate
    public Fdelayadddate
    public Fdelayreguser
    public Foutexecdate
    public Foutexecuser
    public FremainDate
    public Fdispcate1
    public Fdispcate1Name

    public FGroupid
    public FMxLastGroupLoginDt

    public FSocExpireStat
    public FSocCompanyNo
    public FSocClosuredate

    public function getScoExpireStatText()
        select CASE FSocExpireStat
            CASE "C":
                getScoExpireStatText = "폐업"&"("&LEFT(FSocClosuredate,10)&")"
            CASE "H":    
                getScoExpireStatText = "휴업"
        end Select
    end function

    public function getRemainDate()
        if (FremainDate<1) then
            getRemainDate = 0
        else
            getRemainDate = FremainDate
        end if
    end function

    public function IsActionDelayAvailState()
        IsActionDelayAvailState = ((FoutbrandStatus=0) or (FoutbrandStatus=3)) and (FremainDate<14)
    end function

    public function IsActionFinAvailState()
        IsActionFinAvailState = (FoutbrandStatus=0) or (FoutbrandStatus=3) 
    end function

    public function getOutbrandStatusHtml()
        Select CASE FoutbrandStatus
            CASE "0" 
                getOutbrandStatusHtml = "정리예정"
            CASE "3" 
                getOutbrandStatusHtml = "<font color='blue'>정리연장</font>"
            CASE "7" 
                getOutbrandStatusHtml = "<font color='gray'>정리완료</font>"
            CASE ELSE 
                getOutbrandStatusHtml = FoutbrandStatus
        End Select 
    end Function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COutBrand
    Public FItemList()
    public FPageSize
	public FCurrPage

	public FResultCount
	
	public FTotalCount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FOneItem

    public FMayTotalpreSellitemNo
    public FRectMakerid
    public FRectOutbrandStatus
    public FRectDispCate1
    public FRectPreDay

    public Sub getCompanyClosedAndSellitemExistsBrandList()
        Dim sqlStr
        if (FRectDispCate1="") then FRectDispCate1="NULL"

        sqlStr = "exec [db_statistics_const].[dbo].[usp_Ten_BrandService_OutBrand_CompanyClosed] "&FPageSize&","&FCurrPage&",'"&FRectMakerid&"',"&FRectDispCate1
	
    	rsSTSget.CursorLocation = adUseClient                             
	    rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly

        FResultCount =rsSTSget.RecordCount
        if (FResultCount<1) then FResultCount=0

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.EOF
				set FItemList(i) = new COutBrandItem
                ''makerid	groupid	company_no	regstat	closuredate	upddt	sellitemCNT	dispcate1	dispcate1name
                ''dona01	G05189	126-23-68505	C	2017-01-23 00:00:00	2019-04-15 14:29:48.607	24	118	뷰티

                FItemList(i).Fmakerid             = rsSTSget("makerid")
                FItemList(i).Fsubidx              = "-" 
                FItemList(i).FoutbrandStatus      = NULL
                FItemList(i).FbatchmakeDate       = rsSTSget("upddt")
                FItemList(i).FoutScheduledate     = "-"
                FItemList(i).FpreSellitemNo       = rsSTSget("sellitemCNT")
                
                FItemList(i).Fregdate             = rsSTSget("upddt")

                FItemList(i).Fdispcate1         = rsSTSget("dispcate1")
                FItemList(i).Fdispcate1Name     = rsSTSget("dispcate1Name")

                FItemList(i).FGroupid           = rsSTSget("Groupid")

                FItemList(i).FSocExpireStat     = rsSTSget("regstat")
                FItemList(i).FSocCompanyNo      = rsSTSget("company_no")
                FItemList(i).FSocClosuredate    = rsSTSget("closuredate")
				rsSTSget.movenext
				i=i+1
			loop
		end if
		rsSTSget.Close
    end Sub

    public Sub getOutBrandCheckList()
        Dim sqlStr

        if (FRectPreDay="92365") then FRectPreDay="NULL"
        if (FRectDispCate1="") then FRectDispCate1="NULL"

        sqlStr = "exec [db_statistics_const].[dbo].[usp_Ten_BrandService_OutBrand_PreDayCheck_ListCNT] '"&FRectMakerid&"',"&FRectPreDay&","&FRectDispCate1
        rsSTSget.CursorLocation = adUseClient                             
	    rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsSTSget("cnt")
            FMayTotalpreSellitemNo = rsSTSget("preSellitemNo")
		rsSTSget.Close
	
		if FTotalCount < 1 then exit sub
		
		sqlStr = "exec [db_statistics_const].[dbo].[usp_Ten_BrandService_OutBrand_PreDayCheck_List] "&FPageSize&","&FCurrPage&",'"&FRectMakerid&"',"&FRectPreDay&","&FRectDispCate1
		rsSTSget.CursorLocation = adUseClient                             
	    rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly, adLockReadOnly

        FResultCount =rsSTSget.RecordCount
        if (FResultCount<1) then FResultCount=0

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.EOF
				set FItemList(i) = new COutBrandItem

                FItemList(i).Fmakerid             = rsSTSget("makerid")
                FItemList(i).Fsubidx              = "-" ''rsSTSget("subidx")
                FItemList(i).FoutbrandStatus      = NULL ''rsSTSget("outbrandStatus")
                FItemList(i).FbatchmakeDate       = rsSTSget("yyyymmdd")
                FItemList(i).FoutScheduledate     = "-" ''rsSTSget("outScheduledate")
                FItemList(i).FpreSellitemNo       = rsSTSget("sellitemCNT")
                FItemList(i).FpreRegedItemno      = rsSTSget("regItemCNT")
                FItemList(i).FpreSellitemNoSum    = rsSTSget("itemNoSum")
                FItemList(i).FpreSellCostSum      = rsSTSget("itemCostSum")
                FItemList(i).Fpreday              = rsSTSget("preday")
                'FItemList(i).Freguser             = rsSTSget("reguser")
                FItemList(i).Fregdate             = rsSTSget("yyyymmdd")
                'FItemList(i).Fdelayadddate        = rsSTSget("delayadddate")
                'FItemList(i).Fdelayreguser        = rsSTSget("delayreguser")
                'FItemList(i).Foutexecdate     = rsSTSget("outexecdate")
                'FItemList(i).Foutexecuser     = rsSTSget("outexecuser")
                'FItemList(i).FremainDate      = rsSTSget("remainDate")

                FItemList(i).Fdispcate1       = rsSTSget("dispcate1")
                FItemList(i).Fdispcate1Name   = rsSTSget("dispcate1Name")

                FItemList(i).FGroupid               = rsSTSget("Groupid")
                FItemList(i).FMxLastGroupLoginDt    = rsSTSget("MxLastLoginDt")

				rsSTSget.movenext
				i=i+1
			loop
		end if
		rsSTSget.Close

    end Sub

    public Sub getOutBrandScheduledList()
        Dim sqlStr

        if (FRectOutbrandStatus="") then FRectOutbrandStatus="NULL"
        if (FRectDispCate1="") then FRectDispCate1="NULL"

        sqlStr = "exec db_brand.[dbo].[usp_Ten_BrandService_OutBrand_ListCNT] '"&FRectMakerid&"',"&FRectOutbrandStatus&","&FRectDispCate1
        rsget.CursorLocation = adUseClient                             
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
            FMayTotalpreSellitemNo = rsget("preSellitemNo")
		rsget.Close
	
		if FTotalCount < 1 then exit sub
		
		sqlStr = "exec db_brand.[dbo].[usp_Ten_BrandService_OutBrand_List] "&FPageSize&","&FCurrPage&",'"&FRectMakerid&"',"&FRectOutbrandStatus&","&FRectDispCate1
		rsget.CursorLocation = adUseClient                             
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FResultCount =rsget.RecordCount
        if (FResultCount<1) then FResultCount=0

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new COutBrandItem

                FItemList(i).Fmakerid             = rsget("makerid")
                FItemList(i).Fsubidx              = rsget("subidx")
                FItemList(i).FoutbrandStatus      = rsget("outbrandStatus")
                FItemList(i).FbatchmakeDate       = rsget("batchmakeDate")
                FItemList(i).FoutScheduledate     = rsget("outScheduledate")
                FItemList(i).FpreSellitemNo       = rsget("preSellitemNo")
                FItemList(i).FpreRegedItemno      = rsget("preRegedItemno")
                FItemList(i).FpreSellitemNoSum    = rsget("preSellitemNoSum")
                FItemList(i).FpreSellCostSum      = rsget("preSellCostSum")
                FItemList(i).Fpreday              = rsget("preday")
                FItemList(i).Freguser             = rsget("reguser")
                FItemList(i).Fregdate             = rsget("regdate")
                FItemList(i).Fdelayadddate        = rsget("delayadddate")
                FItemList(i).Fdelayreguser        = rsget("delayreguser")
                FItemList(i).Foutexecdate     = rsget("outexecdate")
                FItemList(i).Foutexecuser     = rsget("outexecuser")
                FItemList(i).FremainDate      = rsget("remainDate")

                FItemList(i).Fdispcate1       = rsget("dispcate1")
                FItemList(i).Fdispcate1Name   = rsget("dispcate1Name")

				rsget.movenext
				i=i+1
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
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
        FMayTotalpreSellitemNo =0
	End Sub
	Private Sub Class_Terminate()

	End Sub
End class
%>