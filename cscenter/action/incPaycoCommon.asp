<!--#include file="payco/payco_util.asp"-->
<%

Dim Payco_sellerKey, Payco_sellerKey_WEB, Payco_sellerKey_MOB, Payco_sellerKey_APP
Dim Payco_cpId, Payco_productId, Payco_deliveryId, Payco_deliveryReferenceKey
Dim Payco_orderMethod, Payco_payMode
Dim Payco_WebMode
Dim Payco_LogUse
dim Payco_orderCertifyKey

if (application("Svr_Info")="Dev") then
	Payco_sellerKey				= "S0FSJE"									'(�ʼ�) ������ �ڵ� - ��Ʈ�ʼ��Ϳ��� �˷��ִ� ������, �ʱ� ���� �� PAYCO���� ���θ��� ���� �����Ѵ�.

	'// ����Ʈ���� Ű ����
	Payco_sellerKey_WEB			= "S0FSJE"
	Payco_sellerKey_MOB			= "S0FSJE"
	Payco_sellerKey_APP			= "S0FSJE"

	Payco_cpId					= "PARTNERTEST "								'(�ʼ�) ����ID, 30�� �̳�
	Payco_LogUse 				= False											' Log ��� ���� ( True = ���, False = �̻�� )
	Payco_orderCertifyKey		= ""

else
	Payco_sellerKey				=	""

	'// ����Ʈ���� Ű ����
	Payco_sellerKey_WEB			= "78NUHJ"
	Payco_sellerKey_MOB			= "RR0VR3"
	Payco_sellerKey_APP			= "8973MQ"
	Payco_cpId					= ""
	Payco_LogUse				= False
	Payco_orderCertifyKey		= ""
end if

Payco_productId					=	"PROD_EASY"									'(�ʼ�) ��ǰID, 50�� �̳�
Payco_deliveryId				=	"DELIVERY_PROD"								'(�ʼ�) ��ۺ��ǰID, 50�� �̳�, EASYPAY ��
Payco_deliveryReferenceKey		=	"DV0001"									'(�ʼ�) ���������� �����ϴ� ��ۺ��ǰ ���� Ű, 100�� �̳�, ����, EASYPAY ��
Payco_orderMethod				=	"EASYPAY"									'(�ʼ�) �ֹ�����(=��������) - üũ�ƿ��� : CHECKOUT - ���������+������ id �α��� : EASYPAY_F , ���������+������ id ��α���(PAYCO ȸ������) : EASYPAY
Payco_payMode					=	"PAY2"										'������� ( PAY1 - ��������, �������� / PAY2 - ��������, ���κи� )


'---------------------------------------------------------------------------------
' API �ּ� ���� ( appMode �� ���� �׽�Ʈ�� �Ǽ����� �б�˴ϴ�. )
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

	Dim resultValue			'��� ���Ͽ� JSON ���� ����
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

	cancelType = "PART"																' ��� Type �ޱ� - ALL �Ǵ� PART
	orderNo = ipaygatetid															' PAYCO���� �߱޹��� �ֹ��� ��ȣ
	cancelTotalAmt = cancelAmount													' �� ��� �ݾ�
	totalCancelPossibleAmt = remainAmount											' �� ��Ұ��ɱݾ�
	requestMemo = cancelReason														' ���ó�� ��û�޸�
	''cancelAmt = request("cancelAmt")												' ��� ��ǰ �ݾ� ( PART ��� �� )

	Dim cancelOrder
	Dim orderQuantity, productUnitPrice, productAmt
	Dim cancelTotalFeeAmt
	Dim TotalUnitPrice, TotalProductPaymentAmt
	Dim i, ProductsList, itemCount

	'-----------------------------------------------------------------------------
	' ��� ������ ���� JSON OBJECT�� �����մϴ�.
	'-----------------------------------------------------------------------------
	Set cancelOrder = New aspJSON
	With cancelOrder.data
		'---------------------------------------------------------------------------------
		' ��ü ��� = "ALL", �κ���� = "PART"
		'---------------------------------------------------------------------------------
		Select Case cancelType
			Case "ALL"
				'---------------------------------------------------------------------------------
				' �Ķ���ͷ� ���� ���� ��� �ʿ䰡 ���� �κ��̸�
				' �ֹ� Ű�����θ� DB���� �����͸� �ҷ��;� �Ѵٸ� �� �κп��� �۾��ϼ���.
				'---------------------------------------------------------------------------------
			Case "PART"
				'---------------------------------------------------------------------------------
				' üũ�� �� ����.
				'---------------------------------------------------------------------------------
			Case Else
				'---------------------------------------------------------------------------------
				' ���Ÿ���� �߸��Ǿ���. ( ALL�� PART �� �ƴҰ�� )
				'---------------------------------------------------------------------------------
		End Select

		'---------------------------------------------------------------------------------
		' ������ �ֹ����� ������� Json String �� �ۼ��մϴ�.
		'---------------------------------------------------------------------------------
		.Add "sellerKey", CStr(sellerKey)								'������ �ڵ�. payco_config.asp �� ���� (�ʼ�)
		.Add "orderCertifyKey", CStr(orderCertifyKey)					'PAYCO���� �߱޹��� ������ (�ʼ�)
		.Add "orderNo", CStr(orderNo)									'�ֹ���ȣ
		.Add "cancelTotalAmt", CStr(cancelTotalAmt)						'�ֹ����� �� �ݾ��� �Է��մϴ�. (��ü���, �κ���� ���δ�) (�ʼ�)
		.Add "totalCancelPossibleAmt", CStr(totalCancelPossibleAmt)		'�� ��Ұ��ɱݾ�(������� : ��Ұ��ɱݾ� üũ�� �Է�)
		.Add "requestMemo", CStr(requestMemo)							'���ó�� ��û�޸�

		Dim Result
		'---------------------------------------------------------------------------------
		' �ֹ� ���� ��� API ȣ�� ( JSON �����ͷ� ȣ�� )
		'---------------------------------------------------------------------------------
		Result = payco_cancel(cancelOrder.JSONoutput())

		'-----------------------------------------------------------------------------
		' ����� ȣ���� �ʿ� ����
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

	Dim resultValue			'��� ���Ͽ� JSON ���� ����
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

	cancelType = "ALL"																' ��� Type �ޱ� - ALL �Ǵ� PART
	orderNo = ipaygatetid															' PAYCO���� �߱޹��� �ֹ��� ��ȣ
	cancelTotalAmt = cancelAmount													' �� ��� �ݾ�
	totalCancelPossibleAmt = cancelAmount											' �� ��Ұ��ɱݾ�
	requestMemo = cancelReason														' ���ó�� ��û�޸�
	''cancelAmt = request("cancelAmt")												' ��� ��ǰ �ݾ� ( PART ��� �� )

	Dim cancelOrder
	Dim orderQuantity, productUnitPrice, productAmt
	Dim cancelTotalFeeAmt
	Dim TotalUnitPrice, TotalProductPaymentAmt
	Dim i, ProductsList, itemCount

	'-----------------------------------------------------------------------------
	' ��� ������ ���� JSON OBJECT�� �����մϴ�.
	'-----------------------------------------------------------------------------
	Set cancelOrder = New aspJSON
	With cancelOrder.data
		'---------------------------------------------------------------------------------
		' ��ü ��� = "ALL", �κ���� = "PART"
		'---------------------------------------------------------------------------------
		Select Case cancelType
			Case "ALL"
				'---------------------------------------------------------------------------------
				' �Ķ���ͷ� ���� ���� ��� �ʿ䰡 ���� �κ��̸�
				' �ֹ� Ű�����θ� DB���� �����͸� �ҷ��;� �Ѵٸ� �� �κп��� �۾��ϼ���.
				'---------------------------------------------------------------------------------
			Case "PART"
				'---------------------------------------------------------------------------------
				' üũ�� �� ����.
				'---------------------------------------------------------------------------------
			Case Else
				'---------------------------------------------------------------------------------
				' ���Ÿ���� �߸��Ǿ���. ( ALL�� PART �� �ƴҰ�� )
				'---------------------------------------------------------------------------------
		End Select

		'---------------------------------------------------------------------------------
		' ������ �ֹ����� ������� Json String �� �ۼ��մϴ�.
		'---------------------------------------------------------------------------------
		.Add "sellerKey", CStr(sellerKey)								'������ �ڵ�. payco_config.asp �� ���� (�ʼ�)
		.Add "orderCertifyKey", CStr(orderCertifyKey)					'PAYCO���� �߱޹��� ������ (�ʼ�)
		.Add "orderNo", CStr(orderNo)									'�ֹ���ȣ
		.Add "cancelTotalAmt", CStr(cancelTotalAmt)						'�ֹ����� �� �ݾ��� �Է��մϴ�. (��ü���, �κ���� ���δ�) (�ʼ�)
		.Add "totalCancelPossibleAmt", CStr(totalCancelPossibleAmt)		'�� ��Ұ��ɱݾ�(������� : ��Ұ��ɱݾ� üũ�� �Է�)
		.Add "requestMemo", CStr(requestMemo)							'���ó�� ��û�޸�

		Dim Result
		'---------------------------------------------------------------------------------
		' �ֹ� ���� ��� API ȣ�� ( JSON �����ͷ� ȣ�� )
		'---------------------------------------------------------------------------------
		Result = payco_cancel(cancelOrder.JSONoutput())

		'-----------------------------------------------------------------------------
		' ����� ȣ���� �ʿ� ����
		'-----------------------------------------------------------------------------
		set fnCallPaycoCancel = JSON.parse(Result)
	End With

End Function


%>
