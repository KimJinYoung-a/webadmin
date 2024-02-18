<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ �κ����
' History : �̻� ����
'			2021.09.29 �ѿ�� ����(�˸��� �߼� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommon.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommonNew.asp"-->
<!-- #include virtual="/cscenter/action/incNaverpayCommon.asp"-->
<!-- #include virtual="/cscenter/action/incPaycoCommon.asp"-->
<!-- #include virtual="/cscenter/action/inctosspayCommon.asp"-->
<!-- #include virtual="/cscenter/action/incChaipayCommon.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<%

''��ҿ�û��(1:������,2:������������)
function fnGetNPayCancelRequester(iid)
    dim sqlStr , buf
    fnGetNPayCancelRequester = "2" ''�⺻ ������ ������.

    sqlStr = "select top 1 B.title" &VBCRLF
    sqlStr = sqlStr& " from db_cs.dbo.tbl_new_as_list A" &VBCRLF
    sqlStr = sqlStr& " Join  db_cs.dbo.tbl_new_as_list B" &VBCRLF
    sqlStr = sqlStr& " on A.refasid=B.id" &VBCRLF
    sqlStr = sqlStr& " where A.id="&iid
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if NOT rsget.Eof THEN
        buf = rsget("title")
    end if
    rsget.Close

    if LEFT(buf,3)="[��" then
        fnGetNPayCancelRequester = "1"
    end if
end function

''�������̿� idempotencyKey(������ �ֹ� ��ȣ (������ȣ, �ߺ� ���� ���� ��))
function fnGetChaiPayIdempotencyKey(iid)
    dim sqlStr , buf
    sqlStr = "select top 1 temp_idx" &VBCRLF
    sqlStr = sqlStr& " from db_order.dbo.tbl_order_temp" &VBCRLF
    sqlStr = sqlStr& " where orderserial='" & iid & "'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if NOT rsget.Eof THEN
        buf = rsget("temp_idx")
    end if
    rsget.Close

    fnGetChaiPayIdempotencyKey = buf
end function

''������ ���(�κ����)
function PartialCancelPayco(ipaygatetid, iremainAmount, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byRef inewTid,byref iprimaryPayRestAmount, byref inpointRestAmount, byref itotalRestAmount, byVal orderCertifyKey)
	dim Payco_Result , cancelYmdt
	Set Payco_Result = fnCallPaycoPartialCancel(ipaygatetid, iremainAmount, irefundrequire, "customerReq_P", orderCertifyKey)

	if (Payco_Result.code = 0) then
		iResultCode = "00"                      ''00 ����Ұ�.
		iResultMsg = replace(Payco_Result.message, "'", "")
		cancelYmdt = Payco_Result.result.cancelYmdt
		iCancelDate = LEFT(cancelYmdt,4)&"-"&MID(cancelYmdt,5,2)&"-"&MID(cancelYmdt,7,2)
		iCancelTime = MID(cancelYmdt,9,2)&":"&MID(cancelYmdt,11,2)&":"&MID(cancelYmdt,13,2)

		inewTid = Payco_Result.result.cancelTradeSeq	'// newTid ???

		iprimaryPayRestAmount = 0						'// �Ⱦ���
        inpointRestAmount = 0
        itotalRestAmount = 0
	else
		iResultCode = Payco_Result.code
		iResultMsg = replace(Payco_Result.message, "'", "")
	end if

	Set Payco_Result = Nothing
end function


''���̹����� ���(�κ����)
function PartialCancelNaverPay(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime, byRef inewTid,byref iprimaryPayRestAmount, byref inpointRestAmount, byref itotalRestAmount,byVal nPayCancelRequester)
    dim NPay_Result , cancelYmdt

'    Set NPay_Result = fnCallNaverPayCashAmt(ipaygatetid)
'    rw ":"&NPay_Result.body.totalCashAmount
'    rw ":"&NPay_Result.body.primaryPayMeans
'    Set NPay_Result = Nothing
'    response.end

    Set NPay_Result = fnCallNaverPayCancel(ipaygatetid,irefundrequire,"customerReq_P",nPayCancelRequester)

    if NPay_Result.code="Success" then
        iResultCode = "00"
        iResultMsg = replace(NPay_Result.message,"'","")
        if (iResultMsg="") then iResultMsg = NPay_Result.code

       'rw NPay_Result.body.paymentId
       'rw NPay_Result.body.payHistId
       'rw NPay_Result.body.primaryPayMeans

        'rw NPay_Result.body.primaryPayCancelAmount
        'rw NPay_Result.body.primaryPayRestAmount
        'rw NPay_Result.body.npointCancelAmount
        'rw NPay_Result.body.npointRestAmount
        'rw NPay_Result.body.totalRestAmount

        inewTid = NPay_Result.body.payHistId
        iprimaryPayRestAmount = NPay_Result.body.primaryPayRestAmount
        inpointRestAmount = NPay_Result.body.npointRestAmount
        itotalRestAmount = NPay_Result.body.totalRestAmount
        cancelYmdt = NPay_Result.body.cancelYmdt

        iCancelDate = LEFT(cancelYmdt,4)&"-"&MID(cancelYmdt,5,2)&"-"&MID(cancelYmdt,7,2)
        iCancelTime = MID(cancelYmdt,9,2)&":"&MID(cancelYmdt,11,2)&":"&MID(cancelYmdt,13,2)

    else
        iResultCode = NPay_Result.code
        iResultMsg = replace(NPay_Result.message,"'","")
    end if

    Set NPay_Result = Nothing

end function

''New KaKao �ſ�ī�� �κ� ���
function PartialCancelNewKakaoPay(ipaygatetid, imainAmount, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byRef inewTid)

    Dim objKMPay, cancelYmdt, Status

    Set objKMPay = fnCallKakaoPayPartialCancel(ipaygatetid, imainAmount, irefundrequire, Status)

    if Status="200" then
        iResultCode = "00"                                  ''00 ����Ұ�.
        cancelYmdt = objKMPay.canceled_at                 ''���� ��� �ð�
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = objKMPay.status                        ''�������°�
    else
        iResultCode = objKMPay.code                        ''�����ڵ�
        iResultMsg = objKMPay.message                     ''���� �޼���
    end if

    Set objKMPay = Nothing

end function

function CancelTossPay(ipaygatetid, irefundrequire, irefundNo, byref iResultCode, byref iResultMsg, byref iCancelDate, byref iCancelTime)
	''irefundNo = �ֹ���ȣ_ȯ��ASID
	dim refundData, conResult, rstJson

	refundData = "{"
	refundData = refundData &"""apiKey"":"""&CStr(TossPay_RestApi_Key)&""""
	refundData = refundData &",""payToken"":"""&CStr(ipaygatetid)&""""
	refundData = refundData &",""refundNo"":"""&CStr(irefundNo)&""""
	refundData = refundData &",""amount"":"""&CStr(irefundrequire)&""""
	refundData = refundData &"}"

	conResult = tossapi_refund(refundData)

	Set rstJson = new aspJson
	rstJson.loadJson(conResult)

	if CStr(rstJson.data("code")) = "0" then
		'// ȯ�Ҽ���
		iResultCode = "00"                                  ''00 ����Ұ�.
		iResultMsg = ""
        iCancelDate = LEFT(rstJson.data("approvalTime"),10)
        iCancelTime = RIGHT(rstJson.data("approvalTime"),8)
	else
        iResultCode = CStr(rstJson.data("code"))                        ''�����ڵ�
        iResultMsg = CStr(rstJson.data("msg"))                          ''���� �޼���
	end if
end function

''�������� �κ� ���
function PartialCancelChaiPay(ipaygatetid, idempotencyKey, imainAmount, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byRef inewTid)

    Dim objKMPay, cancelYmdt, Status

    Set objKMPay = fnCallChaiPayPartialCancel(ipaygatetid, idempotencyKey, imainAmount, irefundrequire, Status)

    if Status="200" then
        iResultCode = "00"                                  ''00 ����Ұ�.
        cancelYmdt = objKMPay.updatedAt                  ''���� ��� �ð�
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = objKMPay.status                        ''�������°�
    else
        iResultCode = objKMPay.code                        ''�����ڵ�
        iResultMsg = objKMPay.message                     ''���� �޼���
    end if

    Set objKMPay = Nothing

end function

''KaKao �ſ�ī�� ���
function PartialCancelKakaoPay(ipaygatetid, iremainAmount, irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime, byRef inewTid)

	''iremainAmount, inewTid

    Dim objKMPay

	dim otime,orgTim,diffTime
	otime = Timer()
	orgTim = otime

    '1) ��ü ����
    Set objKMPay = Server.CreateObject("LGCNS.CNSPayService.CnsPayWebConnector")
    objKMPay.RequestUrl = CNSPAY_DEAL_REQUEST_URL

    '2) �α� ����
    objKMPay.SetCnsPayLogging KMPAY_LOG_DIR, KMPAY_LOG_LEVEL	'-1:�α� ��� ����, 0:Error, 1:Info, 2:Debug

    '3) ��û ������ �Ķ���� ����
    objKMPay.AddRequestData "MID", KMPAY_MERCHANT_ID
    objKMPay.AddRequestData "TID", ipaygatetid
    objKMPay.AddRequestData "CancelAmt", irefundrequire
	objKMPay.AddRequestData "CheckRemainAmt", iremainAmount
    ''objKMPay.AddRequestData "SupplyAmt",0     ''���ް�
    ''objKMPay.AddRequestData "GoodsVat",0      ''�ΰ���
    ''objKMPay.AddRequestData "ServiceAmt",0    ''�����
    objKMPay.AddRequestData "CancelMsg","����û"
    objKMPay.AddRequestData "PartialCancelCode","1"     '' 0��ü���, 1�κ����.
    objKMPay.AddRequestData "PayMethod","CARD"

    '4) �߰� �Ķ���� ����
    objKMPay.AddRequestData "actionType", "CL0"  															' actionType : CL0 ���, PY0 ����, CI0 ��ȸ
    objKMPay.AddRequestData "CancelIP", Request.ServerVariables("LOCAL_ADDR")	' ������ ���� ip
    objKMPay.AddRequestData "CancelPwd", KMPAY_CANCEL_PWD														' ��� ��й�ȣ ����

    '5) ������Ű ���� (MID ���� Ʋ��)
    objKMPay.AddRequestData "EncodeKey", KMPAY_MERCHANT_KEY

	diffTime = FormatNumber(Timer()-otime,4)
	rw diffTime

    '6) CNSPAY Lite ���� �����Ͽ� ó��
    objKMPay.RequestAction
	rw diffTime

    '7) ��� ó��
    Dim resultCode, resultMsg, cancelAmt, cancelDate, cancelTime, payMethod, resMerchantId, tid, errorCD, errorMsg, authDate, ccPartCl, stateCD

    resultCode = objKMPay.GetResultData("ResultCode") 	' ����ڵ� (���� :2001(��Ҽ���), 2002(���������), �� �� ����)
    resultMsg = objKMPay.GetResultData("ResultMsg")   	' ����޽���
    cancelAmt = objKMPay.GetResultData("CancelAmt")   	' ��ұݾ�
    cancelDate = objKMPay.GetResultData("CancelDate") 	' �����
    cancelTime = objKMPay.GetResultData("CancelTime")   ' ��ҽð�
    payMethod = objKMPay.GetResultData("PayMethod")   	' ��� ��������
    resMerchantId = objKMPay.GetResultData("MID")     	' ������ ID
    inewTid = objKMPay.GetResultData("TID")             ' TID
    errorCD = objKMPay.GetResultData("ErrorCD")        	' �� �����ڵ�
    errorMsg = objKMPay.GetResultData("ErrorMsg")      	' �� �����޽���
    authDate = cancelDate & cancelTime					' �ŷ��ð�
    ccPartCl = objKMPay.GetResultData("CcPartCl")       ' �κ���� ���ɿ��� (0:�κ���ҺҰ�, 1:�κ���Ұ���)
    stateCD = objKMPay.GetResultData("StateCD")         ' �ŷ������ڵ� (0: ����, 1:�����, 2:�����)

    if (resultCode="2001") then
        iretval = "0"
        iResultCode = resultCode
        iResultMsg = resultMsg
        iCancelDate	= cancelDate
	    iCancelTime	= cancelTime
    else
        iResultCode = resultCode
        iResultMsg = resultMsg
    end if

    Set objKMPay = Nothing

end function

''������ �޴��� �Ǻκ����
function PartialCanCelMobileDacom(ipaygatetid, iremainAmount,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime, byRef inewTid)
    Dim CST_PLATFORM, CST_MID, LGD_MID, LGD_TID,Tradeid, LGD_CANCELREASON, LGD_CANCELREQUESTER, LGD_CANCELREQUESTERIP
	Dim LGD_CANCELAMOUNT, LGD_REMAINAMOUNT
    Dim configPath, xpay

    IF (application("Svr_Info") = "Dev") THEN                   ' LG���÷��� �������� ����(test:�׽�Ʈ, service:����)
		CST_PLATFORM = "test"
	Else
		CST_PLATFORM = "service"
	End If

    CST_MID              = "tenbyten02"                         ' LG���÷������� ���� �߱޹����� �������̵� �Է��ϼ���. //�����, ���� ����.
                                                                ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                               ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    Tradeid     = Split(ipaygatetid,"|")(0)
	LGD_TID     = Split(ipaygatetid,"|")(1)                     ' LG���÷������� ���� �������� �ŷ���ȣ(LGD_TID) : 24 byte

	LGD_CANCELAMOUNT		= irefundrequire					' �κ���� �ݾ�
	LGD_REMAINAMOUNT		= iremainAmount						' ����� ���� �ݾ�

    LGD_CANCELREASON        = "����û"                        ' ��һ���
    LGD_CANCELREQUESTER     = "��"                            ' ��ҿ�û��
    LGD_CANCELREQUESTERIP   = Request.ServerVariables("REMOTE_ADDR")     ' ��ҿ�ûIP


    configPath           = "C:/lgdacom" ''"C:/lgdacom/conf/"&CST_MID         				' LG���÷������� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.
    Set xpay = CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "PartialCancel"
	''xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID

	xpay.Set "LGD_CANCELAMOUNT", LGD_CANCELAMOUNT
    xpay.Set "LGD_REMAINAMOUNT", LGD_REMAINAMOUNT

    xpay.Set "LGD_CANCELREASON", LGD_CANCELREASON
    xpay.Set "LGD_CANCELREQUESTER", LGD_CANCELREQUESTER
    xpay.Set "LGD_CANCELREQUESTERIP", LGD_CANCELREQUESTERIP

	xpay.Set "LGD_TID", LGD_CANCELREQUESTERIP

    '/*
    ' * 1. ������� ��û ���ó��
    ' *
    ' * ��Ұ�� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
	' *
	' * [[[�߿�]]] ���翡�� ������� ó���ؾ��� �����ڵ�
	' * 1. �ſ�ī�� : 0000, AV11
	' * 2. ������ü : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (ȯ�������� ����-> ȯ�Ұ���ڵ�.xls ����)
	' * 3. ������ ���������� ��� 0000(����) �� ��Ҽ��� ó��
	' *
    ' */

    if xpay.TX() then
        '1)������Ұ�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
'Response.Write("������� ��û�� �Ϸ�Ǿ����ϴ�. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

        iretval = "0"
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg

		inewTid = ""

		'' response.write xpay.Tid

		'' rw 	xpay.Response("LGD_TID",0) & "aaaaa"

		'' '' '�Ʒ��� ������û ��� �Ķ���͸� ��� ��� �ݴϴ�.
        '' Dim itemCount
        '' Dim resCount
		'' Dim itemName
		'' Dim i, j
        '' itemCount = xpay.resNameCount
        '' resCount = xpay.resCount

        '' For i = 0 To itemCount - 1
        ''     itemName = xpay.ResponseName(i)
        ''     Response.Write(itemName & "&nbsp:&nbsp")
        ''     For j = 0 To resCount - 1
        ''         Response.Write(xpay.Response(itemName, j) & "<br>")
		'' 	Next
		'' Next
        '' Response.Write("<p>")
    else
        '2)API ��û ���� ȭ��ó��
'Response.Write("������� ��û�� �����Ͽ����ϴ�. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg
    end if

    iCancelDate	= year(now) & "�� " & month(now) & "�� " & day(now) & "��"
	iCancelTime	= hour(now) & "�� " & minute(now) & "�� " & second(now) & "��"

end function

dim id, finishuserid, msg, buyemail, force
dim orderserial, fullText, failText, itemName, itemCnt
dim orgOrderserial, jumundiv, pggubun, pAddParam
itemCnt=0
itemName = ""
id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
buyemail     = RequestCheckVar(request("buyemail"),100)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

if (msg="") and (IsAutoScript) then msg="����� �κ����"


dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo
end if

if (ocsaslist.FResultCount<1) or (orefund.FResultCount<1) then
    if (IsAutoScript) then
        response.write "S_ERR|ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�.[0]"
    else
        response.write "<script>alert('ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|���� ���°� �ƴմϴ�.[1]"
    else
        response.write "<script>alert('���� ���°� �ƴմϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


orderserial = ocsaslist.FOneItem.FOrderserial

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if (returnmethod<>"R120") and (returnmethod<>"R420") and (returnmethod<>"R022") Then
    if (IsAutoScript) then
        response.write "S_ERR|�ſ�ī��, �ǽð���ü, �޴��� �κ���Ҹ� �����մϴ�.[2]"
    else
        response.write "<script>alert('�ſ�ī��, �޴��� �κ���Ҹ� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


Dim IsDacomMobile
if (Len(orefund.FOneItem.FpaygateTid)>=46) then
	IsDacomMobile = True        ''46~49 Tradeid(23) & "|" & vTID(24)
else
	IsDacomMobile = False       ''32~35 Tradeid(23) & "|" & vTID(10)
end if

if orefund.FOneItem.Freturnmethod = "R420" and IsDacomMobile <> True then
    if (IsAutoScript) then
        response.write "S_ERR|������ �ڵ��������� ��� �����մϴ�.[3]"
    else
        response.write "<script>alert('������ �ڵ��������� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'==============================================================================
'���� �ְ������ܱݾ�
dim accountdiv : accountdiv ="100"
dim omainpayment, mainpaymentorg
dim cardcancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcode

if (returnmethod = "R400") or (returnmethod = "R420") then
	accountdiv = "400"
elseif (returnmethod = "R022") then  ''2016/06/21 �߰�.
    accountdiv = "20"
end if


set omainpayment = new COrderMaster

mainpaymentorg = 0			'// ���� �����ݾ�
cardcancelok = "N"
cardcancelerrormsg = ""
cardcancelcount = 0
cardcancelsum   = 0			'// ��ұݾ��հ�
cardcode = ""

if (orderserial<>"") then
	omainpayment.FRectOrderSerial = orderserial

	''response.write accountdiv & "<br />"
	if ((accountdiv = "100") or (accountdiv = "20")) then ''���̹����� �ǽð� ��ü �κ���� ����
		Call omainpayment.getMainPaymentInfo(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	else
		Call omainpayment.getMainPaymentInfoPhone(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	end if

	orgOrderserial = orderserial

	'// ��ȯ�ֹ�( jumundiv = 6 )�̸� ���ֹ����� �������� �����´�.
	sqlStr = " select top 1 m.jumundiv, IsNull(m.pggubun,'') as pggubun, IsNull(e.pAddParam,'') as pAddParam "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr + " 	left join db_order.[dbo].[tbl_order_PaymentEtc] e "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.orderserial = e.orderserial "
	sqlStr = sqlStr + " 		and m.accountdiv = e.acctdiv "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		jumundiv = rsget("jumundiv")
		pggubun = rsget("pggubun")
		pAddParam = rsget("pAddParam")
	end if
	rsget.close

	if (jumundiv = "6") then
		sqlStr = " select top 1 c.orgorderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_order.dbo.tbl_change_order c "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.chgorderserial = '" & orderserial & "' "
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgOrderserial = rsget("orgorderserial")
		end if
		rsget.close

		''2017/09/21 �߰�
		if (pggubun="PY") then
    		sqlStr = " select top 1 m.jumundiv, IsNull(m.pggubun,'') as pggubun, IsNull(e.pAddParam,'') as pAddParam "
        	sqlStr = sqlStr + " from "
        	sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master m "
        	sqlStr = sqlStr + " 	left join db_order.[dbo].[tbl_order_PaymentEtc] e "
        	sqlStr = sqlStr + " 	on "
        	sqlStr = sqlStr + " 		1 = 1 "
        	sqlStr = sqlStr + " 		and m.orderserial = e.orderserial "
        	sqlStr = sqlStr + " 		and m.accountdiv = e.acctdiv "
        	sqlStr = sqlStr + " where "
        	sqlStr = sqlStr + " 	1 = 1 "
        	sqlStr = sqlStr + " 	and m.orderserial = '" & orgOrderserial & "' "
        	rsget.Open sqlStr,dbget,1
        	if Not rsget.Eof then
        		jumundiv = rsget("jumundiv")
        		pggubun = rsget("pggubun")
        		pAddParam = rsget("pAddParam")
        	end if
        	rsget.close
    	end if
	end if
end if


set omainpayment = nothing

IF (cardcancelok<>"Y") then
    if (IsAutoScript) then
        response.write "S_ERR|"&cardcancelerrormsg&"[4]"
    else
        response.write cardcancelerrormsg
        response.write "<script>alert('"&TRIM(cardcancelerrormsg)&"');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
ENd IF

''�κ���� Ƚ�� ����..? ==>
''CHECK

''�κ���� ��������..

'''����� �α� ����.
'''fnINIrePay(MctID, ioldtid, iCancelPrice, iconfirm_price, ibuyeremail, byRef ioldtid, byRef ResultCode, byRef ResultMsg,byRef OldTid, byRef CancelPrice, byRef RepayPrice, byref CntRepay)

'' Pg_Mid
dim MctID
MctID = Mid(orefund.FOneItem.FpaygateTid,11,10)     '''MayBe teenxteenN or INIPayTest
'' response.write MctID

''''���ŷ�ID,��ұݾ�    , ����αݾ�,  ������email
dim ioldtid, iCancelPrice, iconfirm_price


''getMainPaymentInfo
if (orderserial="18032166613") then mainpaymentorg=72420

ioldtid        = orefund.FOneItem.FpaygateTid
iCancelPrice   = orefund.FOneItem.Frefundrequire
iconfirm_price = (mainpaymentorg-cardcancelsum) - orefund.FOneItem.Frefundrequire

if (Not IsAutoScript) then
    response.write "mainpaymentorg : " & mainpaymentorg & "<br />"
    response.write "cardcancelsum : " & cardcancelsum & "<br />"
    response.write "orefund.FOneItem.Frefundrequire : " & orefund.FOneItem.Frefundrequire & "<br />"
end if

''RW iCancelPrice
''RW iconfirm_price
''RW mainpaymentorg
''RW cardcancelsum
'
'response.end
'''��� ��.
dim Tid
dim OldTid, CancelPrice, RepayPrice, CntRepay


IF (ioldtid="") or (iCancelPrice<1) or (iconfirm_price<0) THEN
    if (IsAutoScript) then
        response.write "S_ERR|�κ� ��ұݾ� ���� �Ǵ� TID����[5]"
    else
        response.write "<script>alert('�κ� ��ұݾ� ���� �Ǵ� TID����"&iCancelPrice&"."&iconfirm_price&"');</script>"
        ''response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
END IF

''����� ����
Dim sqlStr, clogIdx

sqlStr = "select * from [db_order].[dbo].tbl_card_cancel_log where 1=0"
rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
    rsget("orderserial") = orgOrderserial
    rsget("orgtid")      = ioldtid
    rsget("cancelprice") = iCancelPrice
    rsget("repayprice")  = iconfirm_price
    rsget("usermail")    = buyemail

    rsget.update
	clogIdx = rsget("clogIdx")
rsget.close


dim INIpay, PInst
dim ResultCode, ResultMsg
dim CancelDate, CancelTime

''īī������
Dim IsKakaoPay : IsKakaoPay = (pggubun = "KA")

''New īī������
Dim IsNewKakaoPay : IsNewKakaoPay = (pggubun = "KK")

''TOSS����
Dim IsTossPay : IsTossPay = (pggubun = "TS")

''���̹�����
Dim iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount
Dim IsNaverPay : IsNaverPay = (pggubun = "NP")

''������
Dim IsPayco : IsPayco = (pggubun = "PY")

''��������
Dim IsChaiPay : IsChaiPay = (pggubun = "CH")

'' IniPay �� ��Ҹ� ����
if (pggubun <> "TS") and (pggubun <> "CH") and ((MID(orefund.FOneItem.FpaygateTid,9,2) <> "NP") and Not(Left(orefund.FOneItem.FpaygateTid,1)="T" and Len(orefund.FOneItem.FpaygateTid)>=20) and Left(orefund.FOneItem.FpaygateTid,3)<>"cns") and (Left(orefund.FOneItem.FpaygateTid,5)<>"KCTEN") and (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIMX_CARD") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIMX_ISP_") AND (Left(orefund.FOneItem.FpaygateTid,6)<>"Stdpay") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIAPICARD") AND orefund.FOneItem.Freturnmethod<>"R400" AND orefund.FOneItem.Freturnmethod<>"R420" and Not IsNumeric(orefund.FOneItem.FpaygateTid) then
    if (IsAutoScript) then
        response.write "S_ERR|�̴Ͻý� �ŷ��� ��� �����մϴ�.[6]"
    else
        response.write "<script>alert('�̴Ͻý� �ŷ��� ��� �����մϴ�..');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'############################################################## �ڵ��� ���� ��� ##############################################################
If orefund.FOneItem.Freturnmethod = "R420" Then

    Dim McashCancelObj, Mrchid, Svcid, Tradeid, Prdtprice, Mobilid, retval

	CALL PartialCanCelMobileDacom(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire,Request("rdsite"),retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'' '// ResultCode �� 4����Ʈ�� 2����Ʈ�� �����Ѵ�.
	'' if (ResultCode = "0000") then
	'' 	ResultCode = "00"
	'' elseif (Left(ResultCode, 2) = "00") then
	'' 	ResultCode = "XX"
	'' else
	'' 	ResultCode = Left(ResultCode, 2)
	'' end if

	CntRepay = 0
	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid

elseif (IsKakaoPay) then
	CALL PartialCancelKakaoPay(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum),orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsNewKakaoPay) then
	CALL PartialCancelNewKakaoPay(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
ELSEIF (IsTossPay) then
    CALL CancelTossPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,(CStr(ocsaslist.FOneItem.Forderserial) & "_" & CStr(ocsaslist.FOneItem.Fid)),ResultCode,ResultMsg,CancelDate,CancelTime)

	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsNaverPay) then
    dim nPayCancelRequester : nPayCancelRequester = fnGetNPayCancelRequester(id)
    CALL PartialCancelNaverPay(orefund.FOneItem.FpaygateTid, orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid, iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount, nPayCancelRequester)

	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsPayco) then
	CALL PartialCancelPayco(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire, "", retval, ResultCode, ResultMsg, CancelDate, CancelTime, Tid, iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount, pAddParam)

	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsChaiPay) then
	CALL PartialCancelChaiPay(orefund.FOneItem.FpaygateTid,fnGetChaiPayIdempotencyKey(ocsaslist.FOneItem.Forderserial), (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'// ���ϰ��� CcPartCl �ϴ� ����(0:�κ���ҺҰ�, 1:�κ���Ұ���)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
else

	''--------------------------------------------------------------------------------------------------

		'*******************************************************************************
		'*
		'* �̹� ���� ���ε� �ŷ����� ��Ҹ� ���ϴ� �ݾ��� �Է��Ͽ� �ٽ� ������ ���ϵ��� ��û�Ѵ�.
		'* ���ŷ� TID�� �κ���Ҹ� ��û�մϴ�.
		'* ��� �� ������̹Ƿ� �� �ŷ� ���̵� �����˴ϴ�.
		'* ���ŷ��� �ſ�ī�� ������ ��쿡�� �����մϴ�
		'* OK ĳ���� ���� ���� ���ԵǾ� �ִ� ��� �κ���� �Ұ����մϴ�
		'* �ݵ�� ����� �ݾ��� �Է��ϵ��� �ϼ���
		'*
		'* Date : 2004/11
		'* Author : ts@inicis.com
		'* Project : INIpay V4.1 for ASP
		'*
		'* http://www.inicis.com
		'* Copyright (C) 2004 Inicis, Co. All rights reserved.
		'*******************************************************************************

		'###############################################################################
		'# 1. ��ü ���� #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")

		'###############################################################################
		'# 2. �ν��Ͻ� �ʱ�ȭ #
		'###############################################################################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. �ŷ� ���� ���� #
		'###############################################################################
		INIpay.SetActionType CLng(PInst), "REPAY"

		'###############################################################################
		'# 4. ���� ���� ���� #
		'###############################################################################
		INIpay.SetField CLng(PInst), "pgid", "INIpayRPAY"  'PG ID (����)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '���� PG IP (����)
		INIpay.SetField CLng(PInst), "mid", MctID '�������̵�
		INIpay.SetField CLng(PInst), "admin", "1111" 'Ű�н�����(�������̵� ���� ����)
		INIpay.SetField CLng(PInst), "oldTid", ioldtid	'����� ���ŷ� ���̵�
		INIpay.SetField CLng(PInst), "currency", "WON" 'ȭ�����
		INIpay.SetField CLng(PInst), "price", iCancelPrice '����
		INIpay.SetField CLng(PInst), "confirm_price", iconfirm_price '����� ��û �ݾ� [�������αݾ� - ����� �ݾ�]
		INIpay.SetField CLng(PInst), "buyeremail", buyemail '�̸���
		INIpay.SetField CLng(PInst), "url", "http://www.10x10.co.kr" '���� Ȩ������ �ּ� (URL)
		INIpay.SetField CLng(PInst), "debug", "true" '�α׸��("true"�� �����ϸ� ���� �α׸� ����)


		'###############################################################################
		'# 5. ����� ��û #
		'###############################################################################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'#6. �κ� ��� ��� #
		'###############################################################################
		Tid         = INIpay.GetResult(CLng(PInst), "tid") '�ŷ���ȣ
		ResultCode  = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ���Ҽ���)
		ResultMsg   = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
		OldTid      = INIpay.GetResult(CLng(PInst), "tid_org") '���ŷ���ȣ
		CancelPrice = INIpay.GetResult(CLng(PInst), "price") '����� �Ǵ� �ݾ�
		RepayPrice  = INIpay.GetResult(CLng(PInst), "pr_remains") '��� �Ǵ� �ݾ�
		CntRepay    = INIpay.GetResult(CLng(PInst), "cnt_partcancel") '�κ����(�����) ��ûȽ��

		'###############################################################################
		'# 7. �ν��Ͻ� ���� # �߰�.
		'####################
		INIpay.Destroy CLng(PInst)

	'''--------------------------------------------------------------------------------------------------
end if

if (CntRepay="") then CntRepay=1 ''2013/06/27 �߰� :: ''Ű���� 'where' ��ó�� ������ �߸��Ǿ����ϴ�.

'// ������ҵǾ�����, CS�����Ϸ�� ��ȯ���� �ʾ��� ���, skyer9, 2017-04-11
If (id = 3402462) Then
	Tid = "INIpayRPAYteenxteen920170406100557769737"
	ResultCode = "00"
	ResultMsg = "[Card|����ó���Ǿ����ϴ�.]"
End If

sqlStr = "update [db_order].[dbo].tbl_card_cancel_log"&VbCRLF
sqlStr = sqlStr & " set newtid='"&Tid&"'"&VbCRLF
sqlStr = sqlStr & " ,resultcode='"&ResultCode&"'"&VbCRLF
sqlStr = sqlStr & " ,resultmsg='"&Replace(ResultMsg,"'","")&"'"&VbCRLF
sqlStr = sqlStr & " ,cancelrequestcount="&CntRepay&VbCRLF
sqlStr = sqlStr & " where clogIdx="&clogIdx&""

dbget.Execute sqlStr

''�κ���� ����ó����.tbl_order_paymentETc ������Ʈ.
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        sqlStr = " update db_order.dbo.tbl_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum="&iprimaryPayRestAmount&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF
        dbget.Execute sqlStr

        sqlStr = " update db_order.dbo.tbl_order_paymentETc"&VbCRLF
        sqlStr = sqlStr & " set realPayedSum="&inpointRestAmount&VbCRLF
        sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
        sqlStr = sqlStr & " and acctdiv='120'"&VbCRLF
        dbget.Execute sqlStr
    ELSE
		'// �Ʒ� ������ ����, skyer9, 2018-04-09
        sqlStr = " update db_order.dbo.tbl_order_paymentETc"&VbCRLF
		sqlStr = sqlStr & " set realPayedSum=realPayedSum-"&CancelPrice&VbCRLF
		sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
		sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF
		dbget.Execute sqlStr

		'// �Ʒ� ������ ����, skyer9, 2017-07-11
  		sqlStr = " update "
  		sqlStr = sqlStr + " 	e set e.realpayedsum = (select IsNull(sum(m.subtotalprice - m.sumpaymentetc),0) from [db_order].[dbo].tbl_order_master m where 1 = 1 and (m.orderserial = '" & orgOrderserial & "' or m.linkorderserial = '" & orgOrderserial & "') and m.cancelyn = 'N') "
		sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
  		sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_PaymentEtc e "
  		sqlStr = sqlStr + " 	on "
  		sqlStr = sqlStr + " 		m.orderserial = e.orderserial "
  		sqlStr = sqlStr + " where "
  		sqlStr = sqlStr + " 	1 = 1 "
  		sqlStr = sqlStr + " 	and m.orderserial = '" & orgOrderserial & "' "
  		sqlStr = sqlStr + " 	and m.accountdiv = e.acctdiv "
		''sqlStr = sqlStr + " 	and e.realpayedsum <> (m.subtotalprice - m.sumpaymentetc) "

  		sqlStr = " update e "
  		sqlStr = sqlStr + " set e.realpayedsum = ( "
  		sqlStr = sqlStr + " 	select IsNull(sum(m.subtotalprice - m.sumpaymentetc),0)  "
  		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m  "
  		sqlStr = sqlStr + " 	where 1 = 1  "
  		sqlStr = sqlStr + " 	and ( "
  		sqlStr = sqlStr + " 		m.orderserial = '" & orgOrderserial & "'  "
  		sqlStr = sqlStr + " 		or  "
  		sqlStr = sqlStr + " 		m.linkorderserial = '" & orgOrderserial & "' "
  		sqlStr = sqlStr + " 		or "
  		sqlStr = sqlStr + " 		m.orderserial in (select chgorderserial from [db_order].[dbo].[tbl_change_order] where orgorderserial = '" & orgOrderserial & "') "
  		sqlStr = sqlStr + " 		or "
  		sqlStr = sqlStr + " 		m.linkorderserial in (select chgorderserial from [db_order].[dbo].[tbl_change_order] where orgorderserial = '" & orgOrderserial & "') "
  		sqlStr = sqlStr + " 	)  "
  		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
  		sqlStr = sqlStr + " ) "
  		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
  		sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_PaymentEtc e "
  		sqlStr = sqlStr + " 	on "
  		sqlStr = sqlStr + " 		m.orderserial = e.orderserial "
  		sqlStr = sqlStr + " where "
  		sqlStr = sqlStr + " 	1 = 1 "
  		sqlStr = sqlStr + " 	and m.orderserial = '" & orgOrderserial & "' "
  		sqlStr = sqlStr + " 	and m.accountdiv = e.acctdiv "
  		sqlStr = sqlStr + "  "
  		''dbget.Execute sqlStr
  	END IF

end if

if (Not IsAutoScript) then
    rw "Tid="&Tid
    rw "ResultCode="&ResultCode
    rw "ResultMsg="&ResultMsg
    rw "OldTid="&OldTid
    rw "CancelPrice="&CancelPrice
    rw "RepayPrice="&RepayPrice
    rw "CntRepay="&CntRepay
end if

''EMail, SMS

set ocsaslist = Nothing
set orefund = Nothing



dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        contents_finish = contents_finish & "��ұݾ� : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "�����ݾ� : " &FormatNumber(itotalRestAmount,0) & VbCrlf
        ''contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "����� ID " & finishuserid
    else
        contents_finish = contents_finish & "��ұݾ� : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "�����ݾ� : " &FormatNumber(RepayPrice,0) & VbCrlf
        ''contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "����� ID " & finishuserid
    end if
end if

if date() >= "2023-03-27" then
' [���̹����� ��û]���̹����� ī��� ����Ʈī�� �κ� ��� ���а� ó�� ��� ����.    ' 2023.03.20 �ѿ��
' ī��� ����Ʈī���� ��� ī��� ��å�� ���� �κ� ��Ұ� �������� �ʾ� �ش� �ǿ� ���� �κ� ��� ��û �� �Ʒ� ������ ���޵ǰ� �ֽ��ϴ�.
' code : CancelNotComplete / message : "3013/����Ʈī��� �κ���Ұ� �Ұ����մϴ�."     '���������α�(�������� ����. 3013/����Ʈī��� �κ���Ұ� �Ұ����մϴ�.)
' ������ ���޵Ǵ� ����Ʈī�� ��� ���аǿ� ���ؼ� ������ ��û ���̵� ���̹����̿��� ȯ�� ó���� �����̹Ƿ� ���������� ��ü ȯ���� �� ���� ȯ���� �� �� �ֽ��ϴ�.
' ī��� ����Ʈī�� ��� ���аǿ� ���ؼ� ������ ��ü ȯ���� �������� �ʵ��� � ��Ź�帳�ϴ�.
if IsNaverPay then
    if ResultCode="ERR" then
        if instr(ResultMsg,"3013/����Ʈī��") > 0 then
            ' ���̹����������� ������ ���� ���/ȯ�� �Ѵٰ� �ϴ�, cs�� �Ϸ�ó���� �Ѵ�.
            Call FinishCSMaster(id, finishuserid, contents_finish)
        end if
    end if
end if
end if

if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    sqlStr = "select r.*, a.userid, m.orderserial, m.buyhp from "
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_as_refund_info r,"
    sqlStr = sqlStr + " [db_cs].dbo.tbl_new_as_list a"
    sqlStr = sqlStr + "     left join db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr + "     on a.orderserial=m.orderserial"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        iorderserial    = rsget("orderserial")
        ibuyhp          = rsget("buyhp")
    end if
    rsget.Close


    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " set refundresult=" + CStr(refundrequire)
    sqlStr = sqlStr + " where asid=" + CStr(id)
    dbget.Execute sqlStr

    Call AddCustomerOpenContents(id, "�κ� ��� �Ϸ�: " & FormatNumber(CancelPrice,0) & VbCRLF & "���� ���� �ݾ�: "& FormatNumber(RepayPrice,0)) '''CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    sqlStr ="select max(replace(replace(replace(replace(replace(ad.itemname,char(9),''),char(10),''),char(13),''),'""',''),'''','')) as itemname, count(ad.masterid) as itemcnt"
    sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list a with (nolock)"
    sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_list aa with (nolock)"
    sqlStr = sqlStr & " 	on a.orderserial = aa.orderserial"
    sqlStr = sqlStr & " 	and a.refasid = aa.id"
    sqlStr = sqlStr & " 	and a.deleteyn='N' and aa.deleteyn='N'"
    sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_detail ad with (nolock)"
    sqlStr = sqlStr & " 	on aa.id=ad.masterid"
    sqlStr = sqlStr & " 	and ad.itemid not in (0)"
    sqlStr = sqlStr & " where a.orderserial = '" & iorderserial & "'"		' �ֹ���ȣ
    sqlStr = sqlStr & " and a.id="& id &""	' �ֹ����

    'response.write sqlStr & "<Br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    IF not rsget.EOF THEN
        itemCnt = rsget("itemcnt")

        if itemName = "" then
            itemName = replace(db2html(rsget("itemname")),vbcrlf,"")
        end if
    END IF
    rsget.close

    if itemCnt > 1 then
        itemName = itemName & " �� " & (itemCnt - 1) & "��"
    end if

    ''���� ��� ��û SMS �߼�
    if (iorderserial<>"") and (ibuyhp<>"") then
    	'sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"+ibuyhp+"','1644-6030','[�ٹ�����]���� �κ� ��� �Ǿ����ϴ�. ������Ҿ� : " + FormatNumber(CancelPrice,0) + " �ֹ���ȣ : " + iorderserial + "'"
        'dbget.Execute sqlStr

        ' �κ����. īī���� �˸��� �߼�.   ' 2021.09.29 �ѿ�� ����
        fullText = "[10x10] ��������ȳ�" & vbCrLf & vbCrLf
        fullText = fullText & "����, �ֹ���Ұ� �Ϸ�Ǿ����ϴ�." & vbCrLf & vbCrLf
        fullText = fullText & "�� �ֹ���ȣ : "& iorderserial &"" & vbCrLf
        fullText = fullText & "�� ��ǰ�� : "& itemName &"" & vbCrLf
        fullText = fullText & "�� ��ұݾ� : "& FormatNumber(CancelPrice,0) &"��"
        failText = "[�ٹ�����]�ֹ���Ұ� �Ϸ�Ǿ����ϴ�.�ֹ���ȣ : "& iorderserial &""
        Call SendKakaoCSMsg_LINK("", ibuyhp,"1644-6030","KC-0021",fullText,"SMS","",failText,"",iorderserial,"")
    end if

    ''����
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & TRIM(ResultMsg) & "');</script>"
        'response.write "<script>opener.location.reload();</script>"
        'response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

else
    if (IsAutoScript) then
        response.write "S_ERR|"&ResultMsg&"[7]"
    else
        response.write ResultCode & "<br>"
        response.write ResultMsg & "<br>"

		response.write "<br><br>* <font color=red>�ݺ������� �κ���ҽ��� ����</font>�� �߻��ϴ� ��� �ý����� ���ǿ��<br>(�ߺ� ����� �� �ֽ��ϴ�.)"
        response.write "<br />* <font color=red>TOSS����</font> �� ��� ȸ��Ż���� �� �ֽ��ϴ�."

    end if
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
