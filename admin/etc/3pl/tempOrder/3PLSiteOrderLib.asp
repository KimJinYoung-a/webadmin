<%

class COrderMasterItem
	public Fcompanyid
	public FSellSite
	public FOutMallOrderSerial
	public FSellDate
	public FPayType
	public FPaydate
	public FOrderUserID
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
	public FdeliverPay

	public FUserID
	public ForderCsGbn
	public FcountryCode
	public Freserve01

	Private Sub Class_Initialize()
		ForderCsGbn = 0
		FcountryCode = "KR"
	End Sub
end class

class COrderDetail
	public FdetailSeq
	public Fbrandid
	public Fbrandname
	public FItemID
	public FItemOption
	public FOutMallItemID
	public FOutMallItemName
	public FOutMallItemOption
	public FOutMallItemOptionName
	public Fitemcost
	public FReducedPrice
	public FItemNo
	public FOutMallCouponPrice
	public FTenCouponPrice
	public FrequireDetail
	public FimageURL

	public FshoplinkerPrdCode
end class

Function RemoveNull(str)
	Dim re
	Set re = New RegExp
	re.Global = True
	re.Pattern = "[\0]"   ' should see backslash zero inside the square braces
	RemoveNull = re.Replace(str,"")
	Set re = Nothing
End Function

function GetOrderFromExtSite(companyid, partnercompanyid)
	select case partnercompanyid
		case "12"
			Call GetOrderFrom_29cm(companyid, partnercompanyid)
		case else
			response.write "잘못된 접근입니다."
		dbget.close : response.end
	end select
end function

function GetOrderFrom_29cm(companyid, partnercompanyid)
	dim apiURL, objXML, resultTxt
	dim JSON, oJSONoutput, jsonString, orders, order, orderdetail, orderdetails, obj
	dim oCOrderMasterItem, oCOrderDetail, oDetailArr, successCnt
	dim tmpStr, pos, sqlStr
	dim i, j, k

	if application("Svr_Info")="Dev" then
		apiURL = "https://apihub.29cm.co.kr/qa/external/tenten/orders/delivery/?limit=50&offset=0&delivery_status=2&multi_status=5"
	else
		apiURL = "https://apihub.29cm.co.kr/external/tenten/orders/delivery/?limit=50&offset=0&delivery_status=2&multi_status=5"
	end if

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", apiURL, false
	objXML.setRequestHeader "X-Partner-Id", "aplusb"
	if application("Svr_Info")="Dev" then
		'
	else
		objXML.Send()
		If objXML.Status = "200" Then
			''Response.Write objXML.responseText
			''Response.Write objXML.responseBody
		else
			Response.Write "통신오류"
			Response.End
		end if
	end if

	resultTxt = ""
	resultTxt = resultTxt + "{"
	resultTxt = resultTxt + "  ""count"": 13562,"
	resultTxt = resultTxt + "  ""next"": ""/external/tenten/orders/delivery/?limit=20&offset=20"","
	resultTxt = resultTxt + "  ""previous"": null,"
	resultTxt = resultTxt + "  ""results"": ["
	resultTxt = resultTxt + "    {"
	resultTxt = resultTxt + "      ""order_no"": 2120137,"
	resultTxt = resultTxt + "      ""order_serial"": ""ORD20181005-0017720"","
	resultTxt = resultTxt + "      ""delivery_amount"": 0,"
	resultTxt = resultTxt + "      ""order_name"": ""김동은"","
	resultTxt = resultTxt + "      ""order_email"": ""qw72380374@gmail.com"","
	resultTxt = resultTxt + "      ""order_phone"": ""010-7238-0374"","
	resultTxt = resultTxt + "      ""receiver_name"": ""김동은"","
	resultTxt = resultTxt + "      ""receiver_phone"": ""010-7238-0374"","
	resultTxt = resultTxt + "      ""receiver_zipcode"": ""61081"","
	resultTxt = resultTxt + "      ""receiver_address1"": ""광주광역시 북구 양산제로 30"","
	resultTxt = resultTxt + "      ""receiver_address2"": ""한국아델리움203동1907호"","
	resultTxt = resultTxt + "      ""etc_message"": ""부재 시 경비실에 맡겨 주세요."","
	resultTxt = resultTxt + "      ""total_order_item_amount"": 35400,"
	resultTxt = resultTxt + "      ""pay_amount"": ""31***"","
	resultTxt = resultTxt + "      ""pay_name"": ""김동은"","
	resultTxt = resultTxt + "      ""pay_timestamp"": ""2018-10-05 17:15:55"","
	resultTxt = resultTxt + "      ""pay_type_description"": ""무통장입금 (가상계좌)"","
	resultTxt = resultTxt + "      ""device_description"": ""mobile.ios"","
	resultTxt = resultTxt + "      ""sell_site_description"": ""29CM_온라인"","
	resultTxt = resultTxt + "      ""order_status"": 5,"
	resultTxt = resultTxt + "      ""order_status_description"": ""주문통보"","
	resultTxt = resultTxt + "      ""order_delivery_description"": ""배송전"","
	resultTxt = resultTxt + "      ""order_cancel_description"": ""취소전"","
	resultTxt = resultTxt + "      ""order_exchange_description"": ""교환안함"","
	resultTxt = resultTxt + "      ""possible_cancel"": ""T"","
	resultTxt = resultTxt + "      ""possible_return"": ""T"","
	resultTxt = resultTxt + "      ""possible_issue_receipt"": ""T"","
	resultTxt = resultTxt + "      ""possible_withdraw"": ""T"","
	resultTxt = resultTxt + "      ""issue_receipt_link"": ""https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid=INIMX_VBNKMOIAPLUSBP20181005171449794654&noMethod=1"","
	resultTxt = resultTxt + "      ""insert_timestamp"": ""2018-10-05 17:14:52"","
	resultTxt = resultTxt + "      ""manages"": ["
	resultTxt = resultTxt + "        {"
	resultTxt = resultTxt + "          ""order_item_manage_no"": 5392834,"
	resultTxt = resultTxt + "          ""order_item_no"": {"
	resultTxt = resultTxt + "            ""order_item_no"": 3903254,"
	resultTxt = resultTxt + "            ""item_no"": 250192,"
	resultTxt = resultTxt + "            ""option_no"": 1284394,"
	resultTxt = resultTxt + "            ""brand_no"": 3799,"
	resultTxt = resultTxt + "            ""brand_name"": ""HOWEAR WELOVE"","
	resultTxt = resultTxt + "            ""discount_rate"": 40,"
	resultTxt = resultTxt + "            ""consumer_price"": 59000,"
	resultTxt = resultTxt + "            ""sell_price"": 35400,"
	resultTxt = resultTxt + "            ""delivery_type"": 2,"
	resultTxt = resultTxt + "            ""option_sell_price"": 0,"
	resultTxt = resultTxt + "            ""item_name"": ""[스페셜오더] 'WELOVE 후리스 스웨트셔츠"","
	resultTxt = resultTxt + "            ""option_value"": ""[COLOR]NAVY [SIZE] L"","
	resultTxt = resultTxt + "            ""option_code"": ""9"","
	resultTxt = resultTxt + "            ""partner_item_name"": ""[스페셜오더] 'WELOVE 후리스 스웨트셔츠"","
	resultTxt = resultTxt + "            ""item_image_url"": ""/next-product/2018/09/28/7bbb3ac29d0e40d499b3c19262ce1987_20180928153214.jpg"","
	resultTxt = resultTxt + "            ""item_image_url_full"": ""https://img.29cm.co.kr/next-product/2018/09/28/7bbb3ac29d0e40d499b3c19262ce1987_20180928153214.jpg?width=200"","
	resultTxt = resultTxt + "            ""supply_type"": 1,"
	resultTxt = resultTxt + "            ""brand_free_shipping_limit"": ""0"","
	resultTxt = resultTxt + "            ""brand_delivery_pay"": ""0"","
	resultTxt = resultTxt + "            ""brand_default_supply_type"": 1,"
	resultTxt = resultTxt + "            ""brand_default_delivery_pay_plus"": ""0"","
	resultTxt = resultTxt + "            ""is_sold_out"": ""F"","
	resultTxt = resultTxt + "            ""front_brand_no"": 3338,"
	resultTxt = resultTxt + "            ""front_brand_name"": ""호웨어 위러브"""
	resultTxt = resultTxt + "          },"
	resultTxt = resultTxt + "          ""order_delivery_no"": {"
	resultTxt = resultTxt + "            ""order_delivery_no"": 3244448,"
	resultTxt = resultTxt + "            ""order_delivery_id"": ""D20181005-0038491"","
	resultTxt = resultTxt + "            ""delivery_describe"": ""29CM무료배송"","
	resultTxt = resultTxt + "            ""use_individual_delivery"": ""F"","
	resultTxt = resultTxt + "            ""return_type_description"": ""일반배송"","
	resultTxt = resultTxt + "            ""delivery_amount"": 0,"
	resultTxt = resultTxt + "            ""details"": [],"
	resultTxt = resultTxt + "            ""place_order_timestamp"": ""2018-10-05 17:20:34"","
	resultTxt = resultTxt + "            ""delayed_day"": 0,"
	resultTxt = resultTxt + "            ""combine_invoice_no"": null,"
	resultTxt = resultTxt + "            ""combine_delivery_company"": null,"
	resultTxt = resultTxt + "            ""confirm_place_order_timestamp"": null,"
	resultTxt = resultTxt + "            ""holding_timestamp"": null,"
	resultTxt = resultTxt + "            ""estimated_date_shipping"": null,"
	resultTxt = resultTxt + "            ""holding_reason"": """","
	resultTxt = resultTxt + "            ""holding_finish_message"": null,"
	resultTxt = resultTxt + "            ""is_free_delivery_grade"": ""F"","
	resultTxt = resultTxt + "            ""default_delivery_pay_plus"": ""0.00"","
	resultTxt = resultTxt + "            ""free_delivery_grade_description"": """""
	resultTxt = resultTxt + "          },"
	resultTxt = resultTxt + "          ""order_consumer_amount"": 59000,"
	resultTxt = resultTxt + "          ""order_item_unit_price"": 35400,"
	resultTxt = resultTxt + "          ""order_item_price"": 59000,"
	resultTxt = resultTxt + "          ""total_order_item_amount"": 35400,"
	resultTxt = resultTxt + "          ""request_comment"": null,"
	resultTxt = resultTxt + "          ""item_coupon_sale_amount"": 0,"
	resultTxt = resultTxt + "          ""add_sale_amount"": 14160,"
	resultTxt = resultTxt + "          ""save_mileage_amount"": 315,"
	resultTxt = resultTxt + "          ""order_count"": 1,"
	resultTxt = resultTxt + "          ""original_count"": 1,"
	resultTxt = resultTxt + "          ""return_phone"": ""02-1644-0560"","

	resultTxt = resultTxt + "          ""return_address"": ""(11154) 경기도 포천시 군내면 용정경제로2길 83 텐바이텐 물류센터"","

	resultTxt = resultTxt + "          ""possible_write_review"": ""F"","
	resultTxt = resultTxt + "          ""possible_return"": ""F"","
	resultTxt = resultTxt + "          ""possible_cancel"": ""T"","
	resultTxt = resultTxt + "          ""possible_exchange"": ""F"","
	resultTxt = resultTxt + "          ""possible_edit_request_comment"": ""T"","
	resultTxt = resultTxt + "          ""possible_change_before_ship"": ""T"","
	resultTxt = resultTxt + "          ""order_item_delivery_status"": 2,"
	resultTxt = resultTxt + "          ""estimated_date_shipping"": null,"
	resultTxt = resultTxt + "          ""holding_reason"": null,"
	resultTxt = resultTxt + "          ""holding_timestamp"": null,"
	resultTxt = resultTxt + "          ""holding_finish_message"": null,"
	resultTxt = resultTxt + "          ""is_apply_order_coupon"": ""T"","
	resultTxt = resultTxt + "          ""order_coupon_sale_amount"": 3894,"
	resultTxt = resultTxt + "          ""order_item_cancel_status_description"": ""정상"","
	resultTxt = resultTxt + "          ""order_item_delivery_status_description"": ""주문통보"","
	resultTxt = resultTxt + "          ""order_item_delivery_front_navigation"": ""결제완료"","
	resultTxt = resultTxt + "          ""cs_description"": ""주문통보"","
	resultTxt = resultTxt + "          ""md_admin_no"": 0,"
	resultTxt = resultTxt + "          ""md_admin_name"": ""엄준선"","
	resultTxt = resultTxt + "          ""cs_admin_no"": 544,"
	resultTxt = resultTxt + "          ""cs_admin_name"": ""이태군"""
	resultTxt = resultTxt + "        }"
	resultTxt = resultTxt + "      ]"
	resultTxt = resultTxt + "    }"
	resultTxt = resultTxt + "  ]"
	resultTxt = resultTxt + "}"

	if application("Svr_Info")="Dev" then
		''response.write resultTxt
	else
		resultTxt = objXML.responseText
		''response.write resultTxt
	end if

	jsonString = resultTxt
	set JSON = New JSONobject
	set oJSONoutput = JSON.Parse(jsonString)

	response.write "조회 주문수(" & oJSONoutput("count") & ")" & "<br />"

	successCnt = 0
	Set orders = oJSONoutput("results")
	for each order in orders.items
		Set oCOrderMasterItem = New COrderMasterItem

		oCOrderMasterItem.Fcompanyid 			= companyid
		oCOrderMasterItem.FSellSite 			= partnercompanyid
		oCOrderMasterItem.FOutMallOrderSerial 	= order("order_no")
		oCOrderMasterItem.FSellDate 			= Left(Now(), 10)
		oCOrderMasterItem.FPayType 				= "50"
		oCOrderMasterItem.FOrderUserID 			= ""

		oCOrderMasterItem.FOrderName 			= html2db(order("order_name"))
		oCOrderMasterItem.FOrderEmail 			= ""
		oCOrderMasterItem.FOrderTelNo 			= html2db(order("order_phone"))
		oCOrderMasterItem.FOrderHpNo 			= html2db(order("order_phone"))
		oCOrderMasterItem.FReceiveName 			= html2db(order("receiver_name"))
		oCOrderMasterItem.FReceiveTelNo 		= html2db(order("receiver_phone"))
		oCOrderMasterItem.FReceiveHpNo 			= html2db(order("receiver_phone"))
		oCOrderMasterItem.FReceiveZipCode 		= html2db(order("receiver_zipcode"))
		oCOrderMasterItem.FReceiveAddr1 		= html2db(order("receiver_address1"))
		oCOrderMasterItem.FReceiveAddr2 		= html2db(order("receiver_address2"))

		oCOrderMasterItem.Fdeliverymemo 		= html2db(order("etc_message"))
		oCOrderMasterItem.FdeliverPay 			= order("delivery_amount")

		'// 우편번호 수정
		if Len(oCOrderMasterItem.FReceiveZipCode) > 4 and InStr(oCOrderMasterItem.FReceiveZipCode, "-") = 0 then
			oCOrderMasterItem.FReceiveZipCode = Left(oCOrderMasterItem.FReceiveZipCode,3) & "-" & Mid(oCOrderMasterItem.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		if (oCOrderMasterItem.FReceiveAddr1 = oCOrderMasterItem.FReceiveAddr2) then
			oCOrderMasterItem.FReceiveAddr2 = ""
		end if
		oCOrderMasterItem.FReceiveAddr1 = TRIM(Replace(oCOrderMasterItem.FReceiveAddr1,"  "," "))
		oCOrderMasterItem.FReceiveAddr2 = TRIM(Replace(oCOrderMasterItem.FReceiveAddr2,"  "," "))
		tmpStr = oCOrderMasterItem.FReceiveAddr1 & " " & oCOrderMasterItem.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oCOrderMasterItem.FReceiveAddr1 = Left(tmpStr, pos)
			oCOrderMasterItem.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		Set orderdetails = order("manages")
		for each orderdetail in orderdetails.items
			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			Set obj = orderdetail("order_item_no")

			oDetailArr(0).FdetailSeq 			= orderdetail("order_item_manage_no")
			oDetailArr(0).FItemNo	 			= orderdetail("order_count")

			oDetailArr(0).Fbrandid 				= obj("brand_no")
			oDetailArr(0).Fbrandname 			= html2db(obj("brand_name"))

			oDetailArr(0).FItemID 				= ""
			oDetailArr(0).FItemOption 			= ""
			oDetailArr(0).FOutMallItemID 		= obj("item_no")
			oDetailArr(0).FOutMallItemName 		= html2db(obj("item_name"))
			oDetailArr(0).FOutMallItemOption 		= obj("option_code")
			oDetailArr(0).FOutMallItemOptionName 	= html2db(obj("option_value"))
			if IsNull(oDetailArr(0).FOutMallItemOption) then
				oDetailArr(0).FOutMallItemOption = "0"
				oDetailArr(0).FOutMallItemOptionName = ""
			end if
			oDetailArr(0).FItemID 				= oDetailArr(0).FOutMallItemID
			oDetailArr(0).FItemOption 			= oDetailArr(0).FOutMallItemOption

			oDetailArr(0).Fitemcost 			= obj("consumer_price")
			oDetailArr(0).FReducedPrice 		= obj("sell_price")
			oDetailArr(0).FOutMallCouponPrice 	= 0
			oDetailArr(0).FTenCouponPrice 		= 0
			oDetailArr(0).FrequireDetail 		= ""
			oDetailArr(0).FimageURL				= html2db(obj("item_image_url_full"))

			oDetailArr(0).FOutMallItemName 			= Replace(oDetailArr(0).FOutMallItemName, "'", "")
			oDetailArr(0).FOutMallItemOptionName 	= Replace(oDetailArr(0).FOutMallItemOptionName, "'", "")

			oDetailArr(0).FOutMallItemName 			= RemoveNull(oDetailArr(0).FOutMallItemName)
			oDetailArr(0).FOutMallItemOptionName 	= RemoveNull(oDetailArr(0).FOutMallItemOptionName)
			''response.write oDetailArr(0).FOutMallItemName & "<br />"

			if (SaveOrderToDB(oCOrderMasterItem, oDetailArr) = True) then
				successCnt = successCnt + 1

				'// 첫 주문디테일에만 배송비입력
				oCOrderMasterItem.FdeliverPay = 0
			end if
		next
	next

	response.write "주문입력(" & successCnt & ")" & "<br />"

	'// 매칭 가능한 상품코드 매칭
	sqlStr = " exec [db_threepl].[dbo].[sp_TEN_xSite_TmpOrder_matchPrdCode] '" & companyid & "', '" & partnercompanyid & "' "
	dbget_TPL.Execute sqlStr

	if companyid = "29cm" and partnercompanyid = "12" then
		'// 신규 브랜드정보 입력
		sqlStr = " exec [db_threepl].[dbo].[sp_TEN_xSite_TmpOrder_Insert_BrandInfo] '" & companyid & "', '" & partnercompanyid & "' "
		dbget_TPL.Execute sqlStr

		'// 신규 상품정보 입력
		sqlStr = " exec [db_threepl].[dbo].[sp_TEN_xSite_TmpOrder_Insert_PrdInfo] '" & companyid & "', '" & partnercompanyid & "' "
		dbget_TPL.Execute sqlStr

		'// 매칭 가능한 상품코드 매칭
		sqlStr = " exec [db_threepl].[dbo].[sp_TEN_xSite_TmpOrder_matchPrdCode] '" & companyid & "', '" & partnercompanyid & "' "
		dbget_TPL.Execute sqlStr
	end if

	set oJSONoutput = Nothing
	set JSON = Nothing
	Set objXML = Nothing

end function

function SaveOrderToDB(oMaster, oDetailArr)
	dim sqlStr
	dim i, j, k
	dim paramInfo, retParamInfo, RetErr, retErrStr
	dim orderDlvPay
	dim tmpStr

	SaveOrderToDB = False

	for i = 0 to UBound(oDetailArr)
		if (i = 0) then
			orderDlvPay = oMaster.FdeliverPay
		else
			orderDlvPay = 0
		end if

		paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
        	,Array("@companyid"				, adVarchar		, adParamInput		, 	32, Trim(oMaster.Fcompanyid))	_
			,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, Trim(oMaster.FSellSite))	_
			,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(oMaster.FOutMallOrderSerial)) _
			,Array("@SellDate"				, adDate		, adParamInput		,	  , Trim(oMaster.FSellDate)) _
			,Array("@PayType"				, adVarchar		, adParamInput		,   32, Trim(oMaster.FPayType)) _
			,Array("@Paydate"				, adDate		, adParamInput		,     , Trim(oMaster.FPaydate)) _
			,Array("@matchItemID"			, adInteger		, adParamInput		,     , Trim(oDetailArr(i).FItemID)) _
			,Array("@matchItemOption"		, adVarchar		, adParamInput		,    4, Trim(oDetailArr(i).FItemOption)) _
			,Array("@partnerItemID"			, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FItemID)) _
			,Array("@partnerItemName"		, adVarchar		, adParamInput		,  128, Trim(oDetailArr(i).FOutMallItemName)) _

			,Array("@partnerOption"			, adVarchar		, adParamInput		,  128, Trim(oDetailArr(i).FItemOption)) _
			,Array("@partnerOptionName"		, adVarchar		, adParamInput		, 1024, Trim(oDetailArr(i).FOutMallItemOptionName)) _
			,Array("@brandid"				, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).Fbrandid)) _
			,Array("@brandname"				, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).Fbrandname)) _
			,Array("@imageURL"				, adVarchar		, adParamInput		,  400, Trim(oDetailArr(i).FimageURL)) _
			,Array("@OrderUserID"			, adVarchar		, adParamInput		,   32, Trim(oMaster.FUserID)) _
			,Array("@OrderName"				, adVarchar		, adParamInput		,   32, Trim(oMaster.FOrderName)) _
			,Array("@OrderEmail"			, adVarchar		, adParamInput		,  100, Trim(oMaster.FOrderEmail)) _
			,Array("@OrderTelNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FOrderTelNo)) _
			,Array("@OrderHpNo"				, adVarchar		, adParamInput		,   16, Trim(oMaster.FOrderHpNo)) _

			,Array("@ReceiveName"			, adVarchar		, adParamInput		,   32, Trim(oMaster.FReceiveName)) _
			,Array("@ReceiveTelNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FReceiveTelNo)) _
			,Array("@ReceiveHpNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FReceiveHpNo)) _
			,Array("@ReceiveZipCode"		, adVarchar		, adParamInput		,   20, Trim(oMaster.FReceiveZipCode)) _
			,Array("@ReceiveAddr1"			, adVarchar		, adParamInput		,  128, Trim(oMaster.FReceiveAddr1)) _
			,Array("@ReceiveAddr2"			, adVarchar		, adParamInput		,  512, Trim(oMaster.FReceiveAddr2)) _
			,Array("@SellPrice"				, adCurrency	, adParamInput		,     , Trim(oDetailArr(i).Fitemcost)) _
			,Array("@RealSellPrice"			, adCurrency	, adParamInput		,     , Trim(oDetailArr(i).FReducedPrice)) _
			,Array("@ItemOrderCount"		, adInteger		, adParamInput		,     , Trim(oDetailArr(i).FItemNo)) _
			,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FdetailSeq)) _

			,Array("@DeliveryType"			, adInteger		, adParamInput		,     , 0) _
			,Array("@deliveryprice"			, adCurrency	, adParamInput		,     , 0) _
			,Array("@deliverymemo"			, adVarchar		, adParamInput		,  400, Trim(oMaster.Fdeliverymemo)) _
			,Array("@requireDetail"			, adVarchar		, adParamInput		, 1024, Trim(oDetailArr(i).FrequireDetail)) _
			,Array("@orderDlvPay"			, adCurrency	, adParamInput		,     , orderDlvPay) _
			,Array("@orderCsGbn"			, adInteger		, adParamInput		,     , oMaster.ForderCsGbn) _
			,Array("@countryCode"			, adVarchar		, adParamInput		,    2, oMaster.FcountryCode) _
            ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   20, Trim(oDetailArr(i).FOutMallItemID)) _
			,Array("@shoplinkerMallName" 	, adVarchar		, adParamInput		,   64, "") _
			,Array("@shoplinkerPrdCode"		, adVarchar		, adParamInput		,   16, "") _

			,Array("@shoplinkerOrderID"		, adVarchar		, adParamInput		,   16, "") _
			,Array("@shoplinkerMallID"		, adVarchar		, adParamInput		,   32, "") _
			,Array("@retErrStr"				, adVarchar		, adParamOutput		,  100, "") _
			,Array("@overseasPrice"			, adCurrency	, adParamInput		,     , 0) _
			,Array("@overseasDeliveryPrice"	, adCurrency	, adParamInput		,     , 0) _
			,Array("@overseasRealPrice"		, adCurrency	, adParamInput		,     , 0) _
			,Array("@reserve01"				, adVarchar		, adParamInput		,   32, "") _
			,Array("@beasongNum11st"		, adVarchar		, adParamInput		,   16, "") _
			,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FOutMallItemOption)) _
    	)

		if (False) then
			rw oMaster.Fcompanyid
			response.write oMaster.FSellSite & "<br />"
			response.write oMaster.FOutMallOrderSerial & "<br />"
			response.write oMaster.FSellDate & "<br />"
			response.write oMaster.FPayType & "<br />"
			response.write oMaster.FPaydate & "<br />"
			response.write oDetailArr(i).FItemID & "<br />"
			response.write oDetailArr(i).FItemOption & "<br />"
			response.write oDetailArr(i).FItemID & "<br />"
			response.write oDetailArr(i).FOutMallItemName & "<br />"
			response.write oDetailArr(i).FItemOption & "<br />"
			response.write oDetailArr(i).FOutMallItemOptionName & "<br />"
			response.write oMaster.FUserID & "<br />"
			response.write oMaster.FOrderName & "<br />"
			response.write oMaster.FOrderEmail & "<br />"
			response.write oMaster.FOrderTelNo & "<br />"
			response.write oMaster.FOrderHpNo & "<br />"
			response.write oMaster.FReceiveName & "<br />"
			response.write oMaster.FReceiveTelNo & "<br />"
			response.write oMaster.FReceiveHpNo & "<br />"
			response.write oMaster.FReceiveZipCode & "<br />"
			response.write oMaster.FReceiveAddr1 & "<br />"
			response.write oMaster.FReceiveAddr2 & "<br />"
			response.write oDetailArr(i).Fitemcost & "<br />"
			response.write oDetailArr(i).FReducedPrice & "<br />"
			response.write oDetailArr(i).FItemNo & "<br />"
			response.write oDetailArr(i).FdetailSeq & "<br />"
			response.write oMaster.Fdeliverymemo & "<br />"
			response.write oDetailArr(i).FrequireDetail & "<br />"
			response.write oMaster.FdeliverPay & "<br />"
			response.write oMaster.ForderCsGbn & "<br />"
			response.write oMaster.FcountryCode & "<br />"
			response.write oDetailArr(i).FOutMallItemID & "<br />"
			response.write oMaster.Fshoplinkermallname & "<br />"
			response.write oDetailArr(i).FshoplinkerPrdCode & "<br />"
			response.write oMaster.FshoplinkerOrderID & "<br />"
			response.write oMaster.FshoplinkerMallID & "<br />"
			response.write oMaster.FoverseasPrice & "<br />"
			response.write oMaster.FoverseasDeliveryPrice & "<br />"
			response.write oMaster.FoverseasRealPrice & "<br />"
			response.write oMaster.Freserve01 & "<br />"
			response.write oMaster.FbeasongNum11st & "<br />"
			dbget.close() : dbget_TPL.close() : response.end
		end if


		sqlStr = "db_threepl.dbo.sp_TEN_xSite_TmpOrder_Insert"

		dbget_TPL.BeginTrans

		retParamInfo = dbtpl_fnExecSPOutput(sqlStr, paramInfo)

        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
        retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

		if (RetErr<0) and (RetErr<>-1) then ''Break
			'// 에러코드 -1 은 중복입력
			dbget_TPL.rollbackTrans
			if IsAutoScript then
				response.write "ERROR["&retErr&"]"& retErrStr
			else
				response.write "ERROR["&retErr&"]"& retErrStr
				response.write "<script>alert('오류가 발생했습니다.');</script>"
			end if

			dbget.close() : dbget_TPL.close() : response.end
		elseif (RetErr <> -1) then
			SaveOrderToDB = True
		end if

		dbget_TPL.CommitTrans
	next

end function

%>
