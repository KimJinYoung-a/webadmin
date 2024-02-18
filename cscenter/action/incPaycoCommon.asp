<!--#include file="payco/payco_util.asp"-->
<%

Dim Payco_sellerKey, Payco_sellerKey_WEB, Payco_sellerKey_MOB, Payco_sellerKey_APP
Dim Payco_cpId, Payco_productId, Payco_deliveryId, Payco_deliveryReferenceKey
Dim Payco_orderMethod, Payco_payMode
Dim Payco_WebMode
Dim Payco_LogUse
dim Payco_orderCertifyKey

if (application("Svr_Info")="Dev") then
	Payco_sellerKey				= "S0FSJE"									'(필수) 가맹점 코드 - 파트너센터에서 알려주는 값으로, 초기 연동 시 PAYCO에서 쇼핑몰에 값을 전달한다.

	'// 사이트별로 키 생성
	Payco_sellerKey_WEB			= "S0FSJE"
	Payco_sellerKey_MOB			= "S0FSJE"
	Payco_sellerKey_APP			= "S0FSJE"

	Payco_cpId					= "PARTNERTEST "								'(필수) 상점ID, 30자 이내
	Payco_LogUse 				= False											' Log 사용 여부 ( True = 사용, False = 미사용 )
	Payco_orderCertifyKey		= ""

else
	Payco_sellerKey				=	""

	'// 사이트별로 키 생성
	Payco_sellerKey_WEB			= "78NUHJ"
	Payco_sellerKey_MOB			= "RR0VR3"
	Payco_sellerKey_APP			= "8973MQ"
	Payco_cpId					= ""
	Payco_LogUse				= False
	Payco_orderCertifyKey		= ""
end if

Payco_productId					=	"PROD_EASY"									'(필수) 상품ID, 50자 이내
Payco_deliveryId				=	"DELIVERY_PROD"								'(필수) 배송비상품ID, 50자 이내, EASYPAY 용
Payco_deliveryReferenceKey		=	"DV0001"									'(필수) 가맹점에서 관리하는 배송비상품 연동 키, 100자 이내, 고정, EASYPAY 용
Payco_orderMethod				=	"EASYPAY"									'(필수) 주문유형(=결재유형) - 체크아웃형 : CHECKOUT - 간편결제형+가맹점 id 로그인 : EASYPAY_F , 간편결제형+가맹점 id 비로그인(PAYCO 회원구매) : EASYPAY
Payco_payMode					=	"PAY2"										'결제모드 ( PAY1 - 결제인증, 승인통합 / PAY2 - 결제인증, 승인분리 )


'---------------------------------------------------------------------------------
' API 주소 설정 ( appMode 에 따라 테스트와 실서버로 분기됩니다. )
'---------------------------------------------------------------------------------
Dim Payco_URL_reserve, Payco_URL_approval, Payco_URL_cancel_check, Payco_URL_cancel, Payco_URL_upstatus, Payco_URL_cancelMileage, Payco_URL_checkUsability, Payco_URL_verifyPayment
dim Payco_URL_bill

if (application("Svr_Info")="Dev") then
	Payco_URL_reserve = "https://alpha-api-bill.payco.com/outseller/order/reserve"
	Payco_URL_approval = "https://alpha-api-bill.payco.com/outseller/payment/approval"
	Payco_URL_cancel_check = "https://alpha-api-bill.payco.com/outseller/order/cancel/checkAvailability"
	Payco_URL_cancel = "https://alpha-api-bill.payco.com/outseller/order/cancel"
	Payco_URL_upstatus = "https://alpha-api-bill.payco.com/outseller/order/updateOrderProductStatus"
	Payco_URL_cancelMileage = "https://alpha-api-bill.payco.com/outseller/order/cancel/partMileage"
	Payco_URL_checkUsability = "https://alpha-api-bill.payco.com/outseller/code/checkUsability"
	Payco_URL_verifyPayment = "https://alpha-api-bill.payco.com/outseller/payment/approval/getDetailForVerify"
	Payco_URL_bill = "https://alpha-bill.payco.com"
else
	Payco_URL_reserve = "https://api-bill.payco.com/outseller/order/reserve"
	Payco_URL_approval = "https://api-bill.payco.com/outseller/payment/approval"
	Payco_URL_cancel_check = "https://api-bill.payco.com/outseller/order/cancel/checkAvailability"
	Payco_URL_cancel = "https://api-bill.payco.com/outseller/order/cancel"
	Payco_URL_upstatus = "https://api-bill.payco.com/outseller/order/updateOrderProductStatus"
	Payco_URL_cancelMileage = "https://api-bill.payco.com/outseller/order/cancel/partMileage"
	Payco_URL_checkUsability = "https://api-bill.payco.com/outseller/code/checkUsability"
	Payco_URL_verifyPayment = "https://api-bill.payco.com/outseller/payment/approval/getDetailForVerify"
	Payco_URL_bill = "https://bill.payco.com"
end if


Function fnCallPaycoPartialCancel(ipaygatetid, remainAmount, cancelAmount, cancelReason, orderCertifyKey)
	Dim orderNo, sellerOrderReferenceKey, sellerOrderProductReferenceKey, cancelTotalAmt, cancelAmt
	Dim totalCancelTaxfreeAmt, totalCancelTaxableAmt, totalCancelVatAmt, totalCancelPossibleAmt, requestMemo, cancelDetailContent
	Dim cancelType
	dim tmpArr, sellerKey

	Dim resultValue			'결과 리턴용 JSON 변수 선언
	Set resultValue = New aspJSON

	sellerKey = Payco_sellerKey
	tmpArr = Split(orderCertifyKey, "|")
	if (UBound(tmpArr) = 1) then
		orderCertifyKey = tmpArr(0)
		Select Case tmpArr(1)
			Case "WEB"
				sellerKey = Payco_sellerKey_WEB
			Case "MOB"
				sellerKey = Payco_sellerKey_MOB
			Case "APP"
				sellerKey = Payco_sellerKey_APP
			Case Else
				''
		End Select
	end if

	cancelType = "PART"																' 취소 Type 받기 - ALL 또는 PART
	orderNo = ipaygatetid															' PAYCO에서 발급받은 주문서 번호
	cancelTotalAmt = cancelAmount													' 총 취소 금액
	totalCancelPossibleAmt = remainAmount											' 총 취소가능금액
	requestMemo = cancelReason														' 취소처리 요청메모
	''cancelAmt = request("cancelAmt")												' 취소 상품 금액 ( PART 취소 시 )

	Dim cancelOrder
	Dim orderQuantity, productUnitPrice, productAmt
	Dim cancelTotalFeeAmt
	Dim TotalUnitPrice, TotalProductPaymentAmt
	Dim i, ProductsList, itemCount

	'-----------------------------------------------------------------------------
	' 취소 내역을 담을 JSON OBJECT를 선언합니다.
	'-----------------------------------------------------------------------------
	Set cancelOrder = New aspJSON
	With cancelOrder.data
		'---------------------------------------------------------------------------------
		' 전체 취소 = "ALL", 부분취소 = "PART"
		'---------------------------------------------------------------------------------
		Select Case cancelType
			Case "ALL"
				'---------------------------------------------------------------------------------
				' 파라메터로 값을 받을 경우 필요가 없는 부분이며
				' 주문 키값으로만 DB에서 데이터를 불러와야 한다면 이 부분에서 작업하세요.
				'---------------------------------------------------------------------------------
			Case "PART"
				'---------------------------------------------------------------------------------
				' 체크할 것 없음.
				'---------------------------------------------------------------------------------
			Case Else
				'---------------------------------------------------------------------------------
				' 취소타입이 잘못되었음. ( ALL과 PART 가 아닐경우 )
				'---------------------------------------------------------------------------------
		End Select

		'---------------------------------------------------------------------------------
		' 설정한 주문정보 변수들로 Json String 을 작성합니다.
		'---------------------------------------------------------------------------------
		.Add "sellerKey", CStr(sellerKey)								'가맹점 코드. payco_config.asp 에 설정 (필수)
		.Add "orderCertifyKey", CStr(orderCertifyKey)					'PAYCO에서 발급받은 인증값 (필수)
		.Add "orderNo", CStr(orderNo)									'주문번호
		.Add "cancelTotalAmt", CStr(cancelTotalAmt)						'주문서의 총 금액을 입력합니다. (전체취소, 부분취소 전부다) (필수)
		.Add "totalCancelPossibleAmt", CStr(totalCancelPossibleAmt)		'총 취소가능금액(현재기준 : 취소가능금액 체크시 입력)
		.Add "requestMemo", CStr(requestMemo)							'취소처리 요청메모

		Dim Result
		'---------------------------------------------------------------------------------
		' 주문 결제 취소 API 호출 ( JSON 데이터로 호출 )
		'---------------------------------------------------------------------------------
		Result = payco_cancel(cancelOrder.JSONoutput())

		'-----------------------------------------------------------------------------
		' 결과를 호출한 쪽에 리턴
		'-----------------------------------------------------------------------------
		''response.write JSON.stringify(Result)
		set fnCallPaycoPartialCancel = JSON.parse(Result)
	End With

End Function


Function fnCallPaycoCancel(ipaygatetid, cancelAmount, cancelReason, orderCertifyKey)
	Dim orderNo, sellerOrderReferenceKey, sellerOrderProductReferenceKey, cancelTotalAmt, cancelAmt
	Dim totalCancelTaxfreeAmt, totalCancelTaxableAmt, totalCancelVatAmt, totalCancelPossibleAmt, requestMemo, cancelDetailContent
	Dim cancelType
	dim tmpArr, sellerKey

	Dim resultValue			'결과 리턴용 JSON 변수 선언
	Set resultValue = New aspJSON

	sellerKey = Payco_sellerKey
	tmpArr = Split(orderCertifyKey, "|")
	if (UBound(tmpArr) = 1) then
		orderCertifyKey = tmpArr(0)
		Select Case tmpArr(1)
			Case "WEB"
				sellerKey = Payco_sellerKey_WEB
			Case "MOB"
				sellerKey = Payco_sellerKey_MOB
			Case "APP"
				sellerKey = Payco_sellerKey_APP
			Case Else
				''
		End Select
	end if

	cancelType = "ALL"																' 취소 Type 받기 - ALL 또는 PART
	orderNo = ipaygatetid															' PAYCO에서 발급받은 주문서 번호
	cancelTotalAmt = cancelAmount													' 총 취소 금액
	totalCancelPossibleAmt = cancelAmount											' 총 취소가능금액
	requestMemo = cancelReason														' 취소처리 요청메모
	''cancelAmt = request("cancelAmt")												' 취소 상품 금액 ( PART 취소 시 )

	Dim cancelOrder
	Dim orderQuantity, productUnitPrice, productAmt
	Dim cancelTotalFeeAmt
	Dim TotalUnitPrice, TotalProductPaymentAmt
	Dim i, ProductsList, itemCount

	'-----------------------------------------------------------------------------
	' 취소 내역을 담을 JSON OBJECT를 선언합니다.
	'-----------------------------------------------------------------------------
	Set cancelOrder = New aspJSON
	With cancelOrder.data
		'---------------------------------------------------------------------------------
		' 전체 취소 = "ALL", 부분취소 = "PART"
		'---------------------------------------------------------------------------------
		Select Case cancelType
			Case "ALL"
				'---------------------------------------------------------------------------------
				' 파라메터로 값을 받을 경우 필요가 없는 부분이며
				' 주문 키값으로만 DB에서 데이터를 불러와야 한다면 이 부분에서 작업하세요.
				'---------------------------------------------------------------------------------
			Case "PART"
				'---------------------------------------------------------------------------------
				' 체크할 것 없음.
				'---------------------------------------------------------------------------------
			Case Else
				'---------------------------------------------------------------------------------
				' 취소타입이 잘못되었음. ( ALL과 PART 가 아닐경우 )
				'---------------------------------------------------------------------------------
		End Select

		'---------------------------------------------------------------------------------
		' 설정한 주문정보 변수들로 Json String 을 작성합니다.
		'---------------------------------------------------------------------------------
		.Add "sellerKey", CStr(sellerKey)								'가맹점 코드. payco_config.asp 에 설정 (필수)
		.Add "orderCertifyKey", CStr(orderCertifyKey)					'PAYCO에서 발급받은 인증값 (필수)
		.Add "orderNo", CStr(orderNo)									'주문번호
		.Add "cancelTotalAmt", CStr(cancelTotalAmt)						'주문서의 총 금액을 입력합니다. (전체취소, 부분취소 전부다) (필수)
		.Add "totalCancelPossibleAmt", CStr(totalCancelPossibleAmt)		'총 취소가능금액(현재기준 : 취소가능금액 체크시 입력)
		.Add "requestMemo", CStr(requestMemo)							'취소처리 요청메모

		Dim Result
		'---------------------------------------------------------------------------------
		' 주문 결제 취소 API 호출 ( JSON 데이터로 호출 )
		'---------------------------------------------------------------------------------
		Result = payco_cancel(cancelOrder.JSONoutput())

		'-----------------------------------------------------------------------------
		' 결과를 호출한 쪽에 리턴
		'-----------------------------------------------------------------------------
		set fnCallPaycoCancel = JSON.parse(Result)
	End With

End Function


%>
