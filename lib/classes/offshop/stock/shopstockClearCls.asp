<%

function IsRecentIpchul(iIpchuldate)
    IsRecentIpchul = false

    if isNULL(iIpchuldate) then Exit function

    if DateDiff("m",iIpchuldate,now())<2 then
        IsRecentIpchul = true
    end if
end function

Class CShopStockClearBrandItem

    public Fshopid
    public Fmakerid
    public Fcomm_cd
    public Fdefaultmargin
    public Fcomm_name
    public FItemCnt
    public FtotSellNo
    public FtotRealSellPrice

    public Ftotsysstockno
    public Ftotrealstockno
    public Ftoterrrealcheckno
	public Ftoterrsampleitemno

    public Ffirstipgodate
    public Flastipgodate



    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

end Class

Class CShopStockClearItem
    public Fitemgubun
    public Fitemid
    public Fitemoption
	public Fshopitemname
	public Fshopitemoptionname
	public Flogicsipgono
	public Flogicsreipgono
	public Fbrandipgono
	public Fbrandreipgono
	public FttlSellno
	public Frealstockno
	public Fsysstockno
	public Ferrsampleitemno
	public Ferrbaditemno
	public Ferrrealcheckno
	public FShopitemprice           ''판매가
	public Fshopsuplycash           ''(현)매입가
	public Fshopbuyprice            ''(현)매장공급가

	public FLstBuycash              ''매입가.   (재고테이블)
    public FLstSuplycash            ''매장공급가(재고테이블)
	public FipbuySum

    public FjungsanCNT
    public FjungsanSum
    public FlastjungsanYYYYMM

    public FCur_realstockno
    public FCur_sysstockno
    public FCur_errrealcheckno

    public function isIpChulNotExists()
        isIpChulNotExists = (Fsysstockno=0) and (Ferrrealcheckno=0) and (FttlSellno=0) and (Fbrandipgono=0) and (Fbrandreipgono=0) and (Flogicsipgono=0) and (Flogicsreipgono=0)
    end function

    public function IsCheckAvail()
        IsCheckAvail = (Frealstockno=0) and (Ferrrealcheckno+Fsysstockno=Frealstockno) or (Ferrrealcheckno<>0) or (Ferrsampleitemno<>0)
        IsCheckAvail = IsCheckAvail and (Not isIpChulNotExists)
    end function

    public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

end Class

Class CShopStockClear

    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

	public FRectShopID
	public FRectMakerID
	''public FRectYYYYMM
	public FRectCommCD
	public FRectLastYYYYMM
	public FRectOnlyerrExist ''오차,실사,시스템>0
	public FRectIpchulcode
	public FRectDispDiv

	public Sub GetShopStockClearBrandList()
	    dim i,sqlStr
        dim ArrList

		'// ===============================================================
		sqlStr = "[db_summary].[dbo].[usp_Ten_Shop_StockClear_BrandLIST]('" + CStr(FRectShopID) + "', '" + CStr(FRectMakerID) + "', '" + CStr(FRectDispDiv) + "', '" + CStr(FRectCommCD) + "')"
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		FTotalCount = 0
		FResultCount = 0
		IF Not (rsget.EOF OR rsget.BOF) THEN
			ArrList = rsget.getRows()
			FResultCount = UBound(ArrList,2)+1
			FTotalCount = FResultCount
		END IF
		rsget.close

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		If IsArray(ArrList) then
			For i=0 to FResultCount-1
				set FItemList(i) = new CShopStockClearBrandItem
        	    FItemList(i).Fshopid           = ArrList(0,i)
                FItemList(i).Fmakerid          = ArrList(1,i)
                FItemList(i).Fcomm_cd          = ArrList(2,i)
                FItemList(i).Fdefaultmargin    = ArrList(3,i)
                FItemList(i).Fcomm_name        = ArrList(4,i)
                FItemList(i).FItemCnt          = ArrList(5,i)
                FItemList(i).FtotSellNo        = ArrList(6,i)
                FItemList(i).FtotRealSellPrice = ArrList(7,i)
                FItemList(i).Ftotsysstockno     = ArrList(8,i)
                FItemList(i).Ftotrealstockno    = ArrList(9,i)
                FItemList(i).Ftoterrrealcheckno = ArrList(10,i)

                FItemList(i).Ffirstipgodate     	= ArrList(11,i)
                FItemList(i).Flastipgodate      	= ArrList(12,i)
				FItemList(i).Ftoterrsampleitemno	= ArrList(13,i)

            next
		end if
    END SUb

    public Sub GetShopStockClearBrandDetail()
        dim i,sqlStr
        dim ArrList

		'// ===============================================================
		if (FRectLastYYYYMM<>"") then
		    sqlStr = "[db_summary].[dbo].[usp_Ten_Shop_StockClear_BrandDetail]('" + CStr(FRectShopID) + "', '" + CStr(FRectMakerID) + "','"&FRectLastYYYYMM&"','"&FRectOnlyerrExist&"')"
		else
		    sqlStr = "[db_summary].[dbo].[usp_Ten_Shop_StockClear_BrandDetail]('" + CStr(FRectShopID) + "', '" + CStr(FRectMakerID) + "','','"&FRectOnlyerrExist&"')"
		    ''rw sqlStr
	    end if
''rw 	sqlStr
'response.end

        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		FTotalCount = 0
		FResultCount = 0
		IF Not (rsget.EOF OR rsget.BOF) THEN
			ArrList = rsget.getRows()
			FResultCount = UBound(ArrList,2)+1
			FTotalCount = FResultCount
		END IF
		rsget.close

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		If IsArray(ArrList) then
			For i=0 to FResultCount-1
				set FItemList(i) = new CShopStockClearItem

				FItemList(i).Fitemgubun         	= ArrList(0,i)
				FItemList(i).Fitemid               	= ArrList(1,i)
				FItemList(i).Fitemoption        	= ArrList(2,i)
				FItemList(i).Fshopitemname      	= db2html(ArrList(3,i))
				FItemList(i).Fshopitemoptionname	= db2html(ArrList(4,i))
				FItemList(i).Flogicsipgono         	= ArrList(5,i)
				FItemList(i).Flogicsreipgono       	= ArrList(6,i)
				FItemList(i).Fbrandipgono           = ArrList(7,i)
				FItemList(i).Fbrandreipgono     	= ArrList(8,i)
				FItemList(i).FttlSellno           	= ArrList(9,i)
				FItemList(i).Frealstockno       	= ArrList(10,i)
				FItemList(i).Fsysstockno        	= ArrList(11,i)
                FItemList(i).Ferrsampleitemno       = ArrList(12,i)
                FItemList(i).Ferrbaditemno          = ArrList(13,i)
                FItemList(i).Ferrrealcheckno        = ArrList(14,i)
                FItemList(i).FshopitemPrice         = ArrList(15,i)
                FItemList(i).Fshopsuplycash         = ArrList(16,i)
                FItemList(i).Fshopbuyprice          = ArrList(17,i)

                FItemList(i).FjungsanCNT            = ArrList(18,i)
                FItemList(i).FjungsanSum            = ArrList(19,i)
                FItemList(i).FlastjungsanYYYYMM     = ArrList(20,i)


'                FItemList(i).FLstBuycash            = ArrList(18,i)
'                FItemList(i).FLstSuplycash          = ArrList(19,i)
'                FItemList(i).FipbuySum              = ArrList(20,i)
'                FItemList(i).FCur_realstockno       = ArrList(21,i)
'                FItemList(i).FCur_sysstockno        = ArrList(22,i)
'                FItemList(i).FCur_errrealcheckno    = ArrList(23,i)
			next
		end if
    END SUb

    Private Sub Class_Initialize()
            FCurrPage       = 1
            FPageSize       = 1000
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
