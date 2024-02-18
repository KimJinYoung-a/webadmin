<%

Class CTPLTempOrderItem
	'// OutMallOrderSeq, OrderSerial, SellSite, SellSiteName, OutMallOrderSerial, SellDate, PayType, PayDate, matchItemID, matchitemoption, orderItemID, orderItemName, orderItemOption, orderItemOptionName,
	'// prdcode, locationidmaker, sellsiteUserID, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, SellPrice, RealSellPrice,
	'// vatinclude, ItemOrderCount, DeliveryType, deliveryprice, RegDate, deliverymemo, countryCode, requireDetail, matchState, orderDlvPay, OrgDetailKey, sendState, sendReqCNT, outMallGoodsNo, orderCsGbn,
	'// ref_OutMallOrderSerial, ref_CSID, etcFinUser, changeitemid, changeitemoption, orgOrderCNT, recvSendState, recvSendReqCnt, shoplinkerOrderID, tenCpnUint, mallCpnUnit, PRE_USE_UNITCOST, outMallJMonth,
	'// overseasPrice, overseasDeliveryPrice, overseasRealPrice, reserve01, beasongNum11st, requireDetail11stYN, sendSongjangNo, subSellSite, outMallOptionNo, companyid, brandid, brandname
	public FOutMallOrderSeq

	public FOrderSerial

	'// 고객사 정보
	public Fcompanyid
	public Fcompanyname
	public FSellSite
	public FSellSiteName

	'// 제휴주문번호
	public FOutMallOrderSerial
	public FOrgDetailKey

	'// 주문자/수취인 정보
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
	public Fdeliverymemo

	'// 주문상품 정보
	public Fbrandid
	public Fbrandname
	public ForderItemID
	public ForderItemName
	public ForderItemOption
	public ForderItemOptionName
	public FItemOrderCount
	public FSellPrice
	public FRealSellPrice

	'// 매칭된 상품정보
	public FmatchItemID
	public Fmatchitemoption
	public Fprdcode
	public FmatchState

	public Fprdname
	public Fprdoptionname
	public FmatchItemName
	public FmatchItemOptionName

	public FRegDate

	public FcountryCode
	public FfoundPrdcode
	public Ftplorderserial
	public FrequireDetail
	public ForderDlvPay
	public FSelldate
	public FoptionCnt
	public FDuppExists
	public FordercsGbn
	public Fbeadaldiv
	public FItemdiv

	public function getmatchStateString()
		if FmatchState="I" then
			getmatchStateString = "엑셀입력"
		elseif FmatchState="P" then
			getmatchStateString = "상품매칭완료"
		elseif FmatchState="O" then
			getmatchStateString = "주문입력완료"
		end if
	end function

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
End Class

Class CTPLTempOrder
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

	public FRectUseYN
	public FRectCompanyID
	public FRectMatchState
	public FRectOutMallOrderSerial
	public FRectSellSite

	public Sub GetTPLTempOrderList()
		dim i,sqlStr, addSql

		addSql = ""
		if (FRectCompanyID <> "") then
			addSql = addSql & " and t.companyid = '" & FRectCompanyID & "'" & vbcrlf
		end if

		if (FRectMatchState <> "") then
			addSql = addSql & " and t.matchState = '" & FRectMatchState & "'" & vbcrlf
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_xSite_TMPOrder] t" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on t.companyid = c.companyid " & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_partnercompany] p on t.companyid = p.companyid and t.sellsite = p.partnercompanyid" & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
			FTotalPage = rsget_TPL("totPg")
		rsget_TPL.Close


		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " t.*, c.company_name, p.partnercompanyname, i.prdname, i.prdoptionname " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_xSite_TMPOrder] t" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on t.companyid = c.companyid " & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_partnercompany] p on t.companyid = p.companyid and t.sellsite = p.partnercompanyid" & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_item] i "
		sqlStr = sqlStr & " on "
		sqlStr = sqlStr & " 1 = 1 "
		sqlStr = sqlStr & " 	and t.companyid = i.companyid "
		sqlStr = sqlStr & " 	and t.prdcode = i.prdcode "
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by t.regdate desc " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTPLTempOrderItem

					FItemList(i).FOutMallOrderSeq	= rsget_TPL("OutMallOrderSeq")

					FItemList(i).FOrderSerial		= rsget_TPL("OrderSerial")

					'// 고객사 정보
					FItemList(i).Fcompanyid			= rsget_TPL("companyid")
					FItemList(i).Fcompanyname		= db2html(rsget_TPL("company_name"))
					FItemList(i).FSellSite			= rsget_TPL("SellSite")
					FItemList(i).FSellSiteName		= db2html(rsget_TPL("partnercompanyname"))

					'// 제휴주문번호
					FItemList(i).FOutMallOrderSerial	= rsget_TPL("OutMallOrderSerial")
					FItemList(i).FOrgDetailKey			= rsget_TPL("OrgDetailKey")

					'// 물류주문번호
					FItemList(i).FOrderSerial			= rsget_TPL("OrderSerial")

					'// 주문자/수취인 정보
					FItemList(i).FOrderName			= db2html(rsget_TPL("OrderName"))
					FItemList(i).FOrderEmail		= db2html(rsget_TPL("OrderEmail"))
					FItemList(i).FOrderTelNo		= db2html(rsget_TPL("OrderTelNo"))
					FItemList(i).FOrderHpNo			= db2html(rsget_TPL("OrderHpNo"))
					FItemList(i).FReceiveName		= db2html(rsget_TPL("ReceiveName"))
					FItemList(i).FReceiveTelNo		= db2html(rsget_TPL("ReceiveTelNo"))
					FItemList(i).FReceiveHpNo		= db2html(rsget_TPL("ReceiveHpNo"))
					FItemList(i).FReceiveZipCode	= db2html(rsget_TPL("ReceiveZipCode"))
					FItemList(i).FReceiveAddr1		= db2html(rsget_TPL("ReceiveAddr1"))
					FItemList(i).FReceiveAddr2		= db2html(rsget_TPL("ReceiveAddr2"))

					'// 주문상품 정보
					FItemList(i).Fbrandid				= rsget_TPL("brandid")
					FItemList(i).Fbrandname				= db2html(rsget_TPL("brandname"))
					FItemList(i).ForderItemID			= rsget_TPL("orderItemID")
					FItemList(i).ForderItemName			= db2html(rsget_TPL("orderItemName"))
					FItemList(i).ForderItemOption		= rsget_TPL("orderItemOption")
					FItemList(i).ForderItemOptionName	= db2html(rsget_TPL("orderItemOptionName"))
					FItemList(i).FItemOrderCount		= rsget_TPL("ItemOrderCount")
					FItemList(i).FSellPrice				= rsget_TPL("SellPrice")
					FItemList(i).FRealSellPrice			= rsget_TPL("RealSellPrice")

					FItemList(i).Fprdname				= db2html(rsget_TPL("prdname"))
					FItemList(i).Fprdoptionname			= db2html(rsget_TPL("prdoptionname"))

					'// 매칭된 상품정보
					FItemList(i).FmatchItemID		= rsget_TPL("matchItemID")
					FItemList(i).Fmatchitemoption	= rsget_TPL("matchitemoption")
					FItemList(i).Fprdcode			= rsget_TPL("prdcode")
					FItemList(i).FmatchState		= rsget_TPL("matchState")

					FItemList(i).FRegDate			= rsget_TPL("RegDate")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
	end sub

	public Function getOnlineTmpOrderRealInputList()
	    Dim sqlStr, paramInfo, i
	    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
    			,Array("@SellSite"	    , adVarchar	, adParamInput	,32, FRectSellSite) _
    			,Array("@OutMallOrderSerial" , adVarchar	, adParamInput	, 32 , FRectOutMallOrderSerial) _
    		)
		sqlStr = "db_threepl.dbo.sp_TEN_xSiteTmpOrderRealInputList"
		'rw sqlStr
        Call dbtpl_fnExecSPReturnRSOutput(sqlStr,paramInfo)

        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1

        FResultCount = rsget_TPL.RecordCount
        redim preserve FItemList(FResultCount)

        if  not rsget_TPL.EOF  then
		do Until rsget_TPL.Eof

			set FItemList(i) = new CTPLTempOrderItem

			FItemList(i).Forderemail		= rsget_TPL("orderemail")
			FItemList(i).FcountryCode		= rsget_TPL("countryCode")
			FItemList(i).FOutMallOrderSeq		= rsget_TPL("OutMallOrderSeq")
			FItemList(i).Fcompanyid				= rsget_TPL("companyid")
			FItemList(i).FSellSite				= rsget_TPL("SellSite")
			FItemList(i).FSellSiteName	        = rsget_TPL("SellSiteName")
            FItemList(i).FmatchItemID           = rsget_TPL("matchItemID")
            FItemList(i).FmatchItemOption       = rsget_TPL("matchItemOption")
			FItemList(i).FmatchItemName			= rsget_TPL("matchItemName")
			FItemList(i).FmatchItemOptionName	= rsget_TPL("matchItemOptionName")
            FItemList(i).ForderItemID			= rsget_TPL("orderItemID")
			FItemList(i).ForderItemName			= rsget_TPL("orderItemName")
			FItemList(i).ForderItemOption		= rsget_TPL("orderItemOption")
			FItemList(i).ForderItemOptionName	= rsget_TPL("orderItemOptionName")
			FItemList(i).FOutMallOrderSerial	= rsget_TPL("OutMallOrderSerial")
			FItemList(i).Fprdcode				= rsget_TPL("prdcode")
			FItemList(i).Ftplorderserial		= rsget_TPL("tplorderserial")
			FItemList(i).FmatchState			= rsget_TPL("matchState")
            FItemList(i).FOrderName             = db2HTML(rsget_TPL("OrderName"))
            FItemList(i).FOrderTelNo            = db2HTML(rsget_TPL("OrderTelNo"))
            FItemList(i).FOrderHpNo             = db2HTML(rsget_TPL("OrderHpNo"))
            FItemList(i).FReceiveName           = db2HTML(rsget_TPL("ReceiveName"))
            FItemList(i).FReceiveTelNo          = db2HTML(rsget_TPL("ReceiveTelNo"))
            FItemList(i).FReceiveHpNo           = db2HTML(rsget_TPL("ReceiveHpNo"))
            FItemList(i).FReceiveZipCode        = db2HTML(rsget_TPL("ReceiveZipCode"))
            FItemList(i).FReceiveAddr1          = db2HTML(rsget_TPL("ReceiveAddr1"))
            FItemList(i).FReceiveAddr2          = db2HTML(rsget_TPL("ReceiveAddr2"))
            FItemList(i).Fdeliverymemo          = db2HTML(rsget_TPL("deliverymemo"))
            FItemList(i).FSellPrice             = rsget_TPL("SellPrice")
            FItemList(i).FRealSellPrice         = rsget_TPL("RealSellPrice")
            FItemList(i).FItemOrderCount        = rsget_TPL("ItemOrderCount")
            FItemList(i).FrequireDetail         = rsget_TPL("requireDetail")
            FItemList(i).ForderDlvPay           = rsget_TPL("orderDlvPay")
            FItemList(i).FSelldate              = rsget_TPL("Selldate")
            FItemList(i).Forderserial           = rsget_TPL("orderserial")

            IF IsNULL(FItemList(i).FrequireDetail) then FItemList(i).FrequireDetail=""

            FItemList(i).FoptionCnt             = rsget_TPL("optionCnt")
            FItemList(i).FDuppExists            = rsget_TPL("DuppExists")
            FItemList(i).FOrgDetailKey          = rsget_TPL("OrgDetailKey")
            FItemList(i).FordercsGbn            = rsget_TPL("ordercsGbn")
            FItemList(i).Fbeadaldiv             = rsget_TPL("beadaldiv")
            FItemList(i).FItemdiv             	= rsget_TPL("itemdiv")
			FItemList(i).Fcompanyid             = rsget_TPL("companyid")
			FItemList(i).Fbrandid             	= rsget_TPL("brandid")
			i=i+1
			rsget_TPL.movenext
		loop
        end if
		rsget_TPL.close
    end Function

	public Sub GetTPLProductOne()
		dim i,sqlStr, addSql

        sqlStr = " select top 1 " & vbcrlf
        sqlStr = sqlStr & " i.*, c.company_name, c.useyn as companyuseyn, b.brand_name, b.brand_name_eng " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_item] i" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on i.companyid = c.companyid " & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_brand] b on i.companyid = b.companyid and i.brandid = b.brandid " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & " and i.companyid = '" & FRectCompanyID & "' " & vbcrlf
		sqlStr = sqlStr & " and i.prdcode = '" & FRectPrdCode & "' " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1

		if not rsget_TPL.Eof then
	        FTotalCount = 1
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		set FOneItem = new CTPLProductItem
		If not rsget_TPL.EOF Then
			FOneItem.Fcompanyid			= rsget_TPL("companyid")
			FOneItem.Fcompanyname		= db2html(rsget_TPL("company_name"))
			FOneItem.Fbrandid			= rsget_TPL("brandid")
			FOneItem.Fbrandname			= db2html(rsget_TPL("brand_name"))
			FOneItem.FbrandnameEng		= db2html(rsget_TPL("brand_name_eng"))
			FOneItem.Fprdcode			= rsget_TPL("prdcode")
			FOneItem.Fprdname			= db2html(rsget_TPL("prdname"))
			FOneItem.Fprdoptionname		= db2html(rsget_TPL("prdoptionname"))
			FOneItem.Fitemgubun			= rsget_TPL("itemgubun")
			FOneItem.Fitemid			= rsget_TPL("itemid")
			FOneItem.Fitemoption		= rsget_TPL("itemoption")
			FOneItem.Fitemoptionname	= db2html(rsget_TPL("itemoptionname"))
			FOneItem.Fcustomerprice		= rsget_TPL("customerprice")
			FOneItem.Fgeneralbarcode	= rsget_TPL("generalbarcode")
			FOneItem.Fuseyn       		= rsget_TPL("useyn")
			FOneItem.Fcompanyuseyn  	= rsget_TPL("companyuseyn")
			FOneItem.Flastupdt      	= rsget_TPL("updt")
			FOneItem.Fregdate       	= rsget_TPL("indt")
        End If
        rsget_TPL.close
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
