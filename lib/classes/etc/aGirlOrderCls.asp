<%
function getbrandSeq2Makerid(ibrandSeq)
    dim sqlStr, ret
    ret =""
    sqlStr = "select IsNULL(tenMakerid,'') as tenMakerid from db_agirlOrder.dbo.tbl_TenLinkBrand"
    sqlStr = sqlStr & " where brandSEq=" & ibrandSeq

    dbagirl_dbget.CursorLocation = adUseClient
    dbagirl_rsget.Open sqlStr,dbagirl_dbget,adOpenForwardOnly, adLockReadOnly
    if not dbagirl_rsget.Eof then
	    ret = dbagirl_rsget("tenMakerid")
	end if
	dbagirl_rsget.Close

    getbrandSeq2Makerid = ret
end function

Class aGirlOrderTmpItem
    public FOrderserial

    public FTotSellPrice
    public FPayRealPrice
    public FOrderName
    public FOrderEmail
    public FOrderTelNo
    public FOrderHpNo

    public FReceiveName
    public FReceiveTelNo
    public FReceiveHpNo
    public FReceiveZipCode
	public FReceiveAddr1
	public FReceiveAddr2
	public FEtcAsk

    public ForderStatus
    public FIsCancel
    public FBaljuDate
    public FSellDate

    public FItemSeq
    public FOptionCode
    public FOrderItemSeq
	public FItemName
	public FOptionValue
	public FAddOrderInfo
	public FOrderCount
	public FConsumerPrice
	public FSellPrice
	public FRealSellPrice

    public FpartnerItemID
    public FpartnerOption
    public FOrderItemStatus
    public FBrandSeq
    public FSupplyPrice

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class aGirlOrder
    public FItemList()

	public FResultCount
	public FTotalCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

    public FRectBRandSeq

    public sub getAgirlOneOrder(iorderSerial)
        dim sqlStr,i
        sqlStr = "db_agirlOrder.[dbo].[usp_Back_LinkMall_OpenOrder_GetOneOrder_TEN]"
        paramInfo = Array(Array("@RETURN_VALUE"	, adInteger	, adParamReturnValue , , 0) _
            ,Array("@partnerSeq" , adInteger	, adParamInput , , 6) _
            ,Array("@brandSeq"  , adInteger	, adParamInput ,32  , FRectBRandSeq)	_
			,Array("@OrderSerial" , adVarchar	, adParamInput ,13 , oneAgirlOrder)	_
		)

        Call dbaGirl_fnExecSPReturnRSOutput(sqlStr, paramInfo)

        FResultCount = dbagirl_rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)
''rw "FResultCount="&FResultCount

        i = 0
        if (Not dbagirl_rsget.Eof) then
            Do until dbagirl_rsget.eof
                set FItemList(i) = new aGirlOrderTmpItem

                FItemList(i).FOrderserial = dbagirl_rsget("orderserial")
                FItemList(i).FTotSellPrice = dbagirl_rsget("TotSellPrice")
                FItemList(i).FPayRealPrice = dbagirl_rsget("PayRealPrice")
                FItemList(i).FOrderName    = dbagirl_rsget("OrderName")
                FItemList(i).FOrderEmail    = dbagirl_rsget("OrderEmail")
                FItemList(i).FOrderTelNo    = dbagirl_rsget("OrderTelNo")
                FItemList(i).FOrderHpNo    = dbagirl_rsget("OrderHpNo")

                FItemList(i).FReceiveName  = dbagirl_rsget("ReceiveName")
                FItemList(i).FReceiveTelNo  = dbagirl_rsget("ReceiveTelNo")
                FItemList(i).FReceiveHpNo  = dbagirl_rsget("ReceiveHpNo")
                FItemList(i).FReceiveZipCode  = replace(dbagirl_rsget("ReceiveZipCode"),"'","")
                FItemList(i).FReceiveAddr1  = dbagirl_rsget("ReceiveAddr1")
                FItemList(i).FReceiveAddr2  = dbagirl_rsget("ReceiveAddr2")
                FItemList(i).FEtcAsk  = dbagirl_rsget("EtcAsk")


                FItemList(i).ForderStatus  = dbagirl_rsget("orderStatus")
                FItemList(i).FIsCancel     = dbagirl_rsget("IsCancel")
                FItemList(i).FBaljuDate    = dbagirl_rsget("BaljuDate")
                FItemList(i).FSellDate     = dbagirl_rsget("SellDate")

                FItemList(i).FItemSeq       = dbagirl_rsget("ItemSeq")
                FItemList(i).FOptionCode    = dbagirl_rsget("OptionCode")
                FItemList(i).FOrderItemSeq  = dbagirl_rsget("OrderItemSeq")
            	FItemList(i).FItemName      = dbagirl_rsget("ItemName")
            	FItemList(i).FOptionValue   = dbagirl_rsget("OptionValue")
            	FItemList(i).FAddOrderInfo  = dbagirl_rsget("AddOrderInfo")
            	FItemList(i).FOrderCount    = dbagirl_rsget("OrderCount")
            	FItemList(i).FConsumerPrice = dbagirl_rsget("ConsumerPrice")
            	FItemList(i).FSellPrice     = dbagirl_rsget("SellPrice")
            	FItemList(i).FRealSellPrice = dbagirl_rsget("RealSellPrice")

                FItemList(i).FpartnerItemID = dbagirl_rsget("partnerItemID")
                FItemList(i).FpartnerOption = dbagirl_rsget("partnerOption")
                FItemList(i).FOrderItemStatus = dbagirl_rsget("OrderItemStatus")

                FItemList(i).FBrandSeq      = dbagirl_rsget("BrandSeq")
                FItemList(i).FSupplyPrice  = dbagirl_rsget("SupplyPrice")
                dbagirl_rsget.movenext
		        i=i+1
		    loop

        end if

        dbagirl_rsget.close
    end Sub

    public sub getAgirlNotRegOrderList(iorderStatus)
        dim sqlStr
        sqlStr = " db_agirlOrder.[dbo].[usp_Back_LinkMall_OpenOrder_GetList_TEN]"

        Dim paramInfo(3)
		paramInfo(0) = MakeParam("@partnerSeq",adInteger,adParamInput,,6)
		paramInfo(1) = MakeParam("@brandSeq",adInteger,adParamInput,,FRectBRandSeq)
		paramInfo(2) = MakeParam("@SDate",advarchar,adParamInput,10,"")
		paramInfo(3) = MakeParam("@OrderStatus",adInteger,adParamInput,,iorderStatus)

        Call dbaGirl_fnExecSPReturnRSOutput(sqlStr, paramInfo)

        FResultCount = dbagirl_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)

        if (Not dbagirl_rsget.Eof) then
            Do until dbagirl_rsget.eof
                set FItemList(i) = new aGirlOrderTmpItem
                FItemList(i).FOrderserial = dbagirl_rsget("orderserial")
                FItemList(i).FTotSellPrice = dbagirl_rsget("TotSellPrice")
                FItemList(i).FPayRealPrice = dbagirl_rsget("PayRealPrice")
                FItemList(i).FOrderName    = dbagirl_rsget("OrderName")
                FItemList(i).FReceiveName  = dbagirl_rsget("ReceiveName")

                FItemList(i).ForderStatus  = dbagirl_rsget("orderStatus")
                FItemList(i).FIsCancel     = dbagirl_rsget("IsCancel")
                FItemList(i).FBaljuDate    = dbagirl_rsget("BaljuDate")

                FItemList(i).FItemSeq       = dbagirl_rsget("ItemSeq")
                FItemList(i).FOptionCode    = dbagirl_rsget("OptionCode")
                FItemList(i).FOrderItemSeq  = dbagirl_rsget("OrderItemSeq")
            	FItemList(i).FItemName      = dbagirl_rsget("ItemName")
            	FItemList(i).FOptionValue   = dbagirl_rsget("OptionValue")
            	FItemList(i).FAddOrderInfo  = dbagirl_rsget("AddOrderInfo")
            	FItemList(i).FOrderCount    = dbagirl_rsget("OrderCount")
            	FItemList(i).FConsumerPrice = dbagirl_rsget("ConsumerPrice")
            	FItemList(i).FSellPrice     = dbagirl_rsget("SellPrice")
            	FItemList(i).FRealSellPrice = dbagirl_rsget("RealSellPrice")

                FItemList(i).FpartnerItemID = dbagirl_rsget("partnerItemID")
                FItemList(i).FpartnerOption = dbagirl_rsget("partnerOption")
                FItemList(i).FOrderItemStatus = dbagirl_rsget("OrderItemStatus")

                dbagirl_rsget.movenext
		        i=i+1
		    loop

        end if

        dbagirl_rsget.close
    end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
	FScrollCount = 10
		FTotalCount =0

        if (application("Svr_Info")	= "Dev") then
            FRectBRandSeq = 291
        else
            FRectBRandSeq = 314
        end if
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