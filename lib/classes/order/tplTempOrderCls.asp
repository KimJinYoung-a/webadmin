<%
Class CTplTempOrderItem

	public FOutMallOrderSeq

    public Ftplcompanyid
    public FSellSite
	public FSellSiteName
    public Fpartnercompanyname

	public Fitemgubun
	public Fitemid
	public Fitemoption
	public FItemName
	public FItemOptionName

	public Fbarcode

    public ForderItemID
    public ForderItemName
    public ForderItemOption
    public ForderItemOptionName

	public FItemOrderCount

    public FlinkItemID
    public FlinkItemName
    public FlinkItemOption
    public FlinkItemOptionName

    public FOutMallOrderSerial
	public FOrgDetailKey
    public FfoundPrdcode
    public Ftplprdcode
    public Forderserial

    public FmatchState
	public FsendState
	public Fsongjangno
	public Fsongjangdiv

	public function getmatchStateString()
		if FmatchState="I" then
			'단품
			getmatchStateString = "엑셀입력"
		elseif FmatchState="P" then
			'제외
			getmatchStateString = "상품매칭완료"
		elseif FmatchState="O" then
			'포함
			getmatchStateString = "주문입력완료"
		end if
	end function

	public function getSendStateString()
		if FsendState="Y" then
			getSendStateString = "다운로드완료"
		elseif FsendState="N" then
			getSendStateString = "다운이전"
		else
			getSendStateString = FsendState
		end if
	end function

	public function getorderItemName()

		getorderItemName = ForderItemName

	end function

	public function IsItemMatched()
		if (FmatchState="I" and Fitemid <> "") then
			IsItemMatched = true
		else
			IsItemMatched = false
		end if

	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CTplTempOrder
    public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectTPLCompanyID
	public FRectSellSite
	public FRectMatchState
	public FRectSendState
	public FRectorderserial
	public FRectoutmallorderserial
	public FRectregYYYYMMDD
	public FRectBeasongDate
	public FRectSearchField
	public FRectSearchText

	public Function IsAllMatched(iOutMallOrderSerial)
	    dim i
	    For i=LBound(FItemList) to UBound(FItemList)
	        if IsObject(FItemList(i)) then
	            if FItemList(i).FOutMallOrderSerial=iOutMallOrderSerial then
	                IF (Not FItemList(i).IsItemMatched) then
	                    IsAllMatched = false
	                    Exit function
	                End IF
	            end if
	        end if
	    Next

	    IsAllMatched = true
    end function

	public Function getOnlineTmpOrderList
	    Dim sqlStr, paramInfo, i
	    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
    			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage)	_
    			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize) _
    			,Array("@TPLCompanyID"		, adVarchar	, adParamInput		, 32	, FRectTPLCompanyID) _
    			,Array("@SellSite"	    , adInteger	, adParamInput	, 	    , FRectSellSite) _
    			,Array("@RectMatchState" , adVarchar	, adParamInput		, 10 , FRectMatchState) _
				,Array("@OrderSerial" , adVarchar	, adParamInput	, 32 	, FRectorderserial) _
				,Array("@OutMallOrderSerial" , adVarchar	, adParamInput	, 32 , FRectoutmallorderserial) _
				,Array("@regYYYYMMDD" , adVarchar	, adParamInput	, 10 	, FRectregYYYYMMDD) _
    		)
        sqlStr = "db_threepl.dbo.usp_OnlineTmpOrderList"

        Call dbtpl_fnExecSPReturnRSOutput(sqlStr,paramInfo)

        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1

        FResultCount = rsget_TPL.RecordCount
        redim preserve FItemList(FResultCount)

        if  not rsget_TPL.EOF  then
		do Until rsget_TPL.Eof

			set FItemList(i) = new CTplTempOrderItem

			FItemList(i).FOutMallOrderSeq		= rsget_TPL("OutMallOrderSeq")

			FItemList(i).Ftplcompanyid			= rsget_TPL("tplcompanyid")
			FItemList(i).FSellSite				= rsget_TPL("SellSite")
			FItemList(i).FSellSiteName			= rsget_TPL("SellSiteName")
			''FItemList(i).Fpartnercompanyname	= rsget_TPL("partnercompanyname")

			FItemList(i).Fitemgubun				= rsget_TPL("itemgubun")
			FItemList(i).Fitemid				= rsget_TPL("itemid")
			FItemList(i).Fitemoption			= rsget_TPL("itemoption")
			FItemList(i).FItemName				= rsget_TPL("ItemName")
			FItemList(i).FItemOptionName		= rsget_TPL("ItemOptionName")

			FItemList(i).FItemOrderCount		= rsget_TPL("ItemOrderCount")

			FItemList(i).Fbarcode				= rsget_TPL("barcode")

			FItemList(i).ForderItemID			= rsget_TPL("orderItemID")
			FItemList(i).ForderItemName			= rsget_TPL("orderItemName")
			FItemList(i).ForderItemOption		= rsget_TPL("orderItemOption")
			FItemList(i).ForderItemOptionName	= rsget_TPL("orderItemOptionName")

			FItemList(i).Forderserial			= rsget_TPL("orderserial")

			FItemList(i).FOutMallOrderSerial	= rsget_TPL("OutMallOrderSerial")
			FItemList(i).FOrgDetailKey			= rsget_TPL("OrgDetailKey")

			FItemList(i).FmatchState			= rsget_TPL("matchState")

			i=i+1
			rsget_TPL.movenext
		loop
        end if
		rsget_TPL.close

    end Function

	public Function getOnlineTmpOrderChulgoList
	    Dim sqlStr, paramInfo, i
	    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
    			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage)	_
    			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize) _
    			,Array("@TPLCompanyID"		, adVarchar	, adParamInput	, 32	, FRectTPLCompanyID) _
    			,Array("@SellSite"	    , adInteger	, adParamInput	, 	    , FRectSellSite) _
    			,Array("@sendState" , adVarchar	, adParamInput	, 10 , FRectSendState) _
				,Array("@OrderSerial" , adVarchar	, adParamInput	, 32 , FRectorderserial) _
				,Array("@OutMallOrderSerial" , adVarchar	, adParamInput	, 32 , FRectoutmallorderserial) _
				,Array("@BeasongDate" , adVarchar	, adParamInput	, 10 , FRectBeasongDate) _
				,Array("@SearchField" , adVarchar	, adParamInput	, 32 , FRectSearchField) _
				,Array("@SearchText" , adVarchar	, adParamInput	, 32 , FRectSearchText) _
    		)
        sqlStr = "db_threepl.dbo.usp_OnlineTmpOrderChulgoList"

        Call fnExecSPReturnRSOutput(sqlStr,paramInfo)

        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

        if  not rsget.EOF  then
		do Until rsget.Eof

			set FItemList(i) = new CTplTempOrderItem

			FItemList(i).FOutMallOrderSeq		= rsget("OutMallOrderSeq")

			FItemList(i).Ftplcompanyid			= rsget("tplcompanyid")
			FItemList(i).FSellSite				= rsget("SellSite")
			FItemList(i).FSellSiteName			= rsget("SellSiteName")
			''FItemList(i).Fpartnercompanyname	= rsget("partnercompanyname")

			FItemList(i).Fitemgubun				= rsget("itemgubun")
			FItemList(i).Fitemid				= rsget("itemid")
			FItemList(i).Fitemoption			= rsget("itemoption")
			FItemList(i).FItemName				= rsget("ItemName")
			FItemList(i).FItemOptionName		= rsget("ItemOptionName")

			FItemList(i).ForderItemID			= rsget("orderItemID")
			FItemList(i).ForderItemName			= rsget("orderItemName")
			FItemList(i).ForderItemOption		= rsget("orderItemOption")
			FItemList(i).ForderItemOptionName	= rsget("orderItemOptionName")

			FItemList(i).Forderserial			= rsget("orderserial")

			FItemList(i).FOutMallOrderSerial	= rsget("OutMallOrderSerial")
			FItemList(i).FOrgDetailKey			= rsget("OrgDetailKey")

			FItemList(i).FsendState				= rsget("sendState")
			FItemList(i).Fsongjangno			= rsget("songjangno")
			FItemList(i).Fsongjangdiv			= rsget("songjangdiv")

			i=i+1
			rsget.movenext
		loop
        end if
		rsget.close

    end Function

    public Function fnOutmallOrderGetList
		fnOutmallOrderGetList =  clsConnDB.fnExecSPReturnRS("db_agirlOrder.dbo.[usp_Back_OutMallOrder_GetList]("&FRectSellSite&","&FOrderStatus&",'"&FSDate&"','"&FEDate&"','"&FIsMatching&"')")

	End Function

	public Function fnOutmallOrderGetDetail
		fnOutmallOrderGetDetail =  clsConnDB.fnExecSPReturnRS("db_agirlOrder.dbo.[usp_Back_OutMallOrder_GetDetailList]("&FSellSite&",'"&FOutMallOrderSerial&"')")
	End Function


	public Function fnOutmallOrderGetData
		Dim arrValue
		arrValue = clsConnDB.fnExecSPReturnArr("[usp_Back_OutMallOrder_GetData]("&FSellSite&",'"&FOutMallOrderSerial&"')",20)
		IF isArray(arrValue) THEN
			FOrderSerial			= arrValue(0)
			FSellSite			    	= arrValue(1)
			FPartnerSeq		    	= arrValue(2)
			FOutMallOrderSerial     = arrValue(3)
			FOrderName                 	= arrValue(4)
			FOrderEmail		      	= arrValue(5)
			FOrderTelNo            	 = arrValue(6)
			FOrderHpNo                  = arrValue(7)
			FReceiveName		       = arrValue(8)
			FReceiveTelNo           	= arrValue(9)
			FReceiveHpNo               = arrValue(10)
			FReceiveZipCode           = arrValue(11)
			FReceiveAddr1              = arrValue(12)
			FReceiveAddr2               = arrValue(13)
			FEtcAsk                     	= arrValue(14)
			FTotSellPrice               	= arrValue(15)
			FPayDate                   	= arrValue(16)
			FDeliveryType               	= arrValue(17)
			FOptionCodeChk           	= arrValue(18)
			FWizDeliveryPay			= arrValue(19)
		END IF
	End Function


    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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
