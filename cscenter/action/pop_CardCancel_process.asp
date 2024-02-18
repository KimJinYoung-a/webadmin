<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ��ü���
' History : �̻� ����
'			2021.10.14 �ѿ�� ����(�˸��� �߼ۺ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommon.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommonNew.asp"-->
<!-- #include virtual="/cscenter/action/incNaverpayCommon.asp"-->
<!-- #include virtual="/cscenter/action/incPaycoCommon.asp"-->
<!-- #include virtual="/cscenter/action/inctosspayCommon.asp"-->
<!-- #include virtual="/cscenter/action/incChaipayCommon.asp"-->
<%
'' ��ġó�� wapi/cscenter/action/pop_CardCancel_process.asp ::2016/04 �� �۾�����.

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

''������ ���(��ü���)
function CancelPayco(ipaygatetid, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byVal orderCertifyKey)
	dim Payco_Result , cancelYmdt
	Set Payco_Result = fnCallPaycoCancel(ipaygatetid, irefundrequire, "customerReq", orderCertifyKey)

	if (Payco_Result.code = 0) then
		iResultCode = "00"                      ''00 ����Ұ�.
		iResultMsg = replace(Payco_Result.message, "'", "")
		cancelYmdt = Payco_Result.result.cancelYmdt
		iCancelDate = LEFT(cancelYmdt,4)&"-"&MID(cancelYmdt,5,2)&"-"&MID(cancelYmdt,7,2)
		iCancelTime = MID(cancelYmdt,9,2)&":"&MID(cancelYmdt,11,2)&":"&MID(cancelYmdt,13,2)
	else
		iResultCode = Payco_Result.code
		iResultMsg = replace(Payco_Result.message, "'", "")
	end if

	Set Payco_Result = Nothing
end function

''���̹����� ���(��ü���)
function CanCelNaverPay(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime,byVal nPayCancelRequester)
    dim NPay_Result , cancelYmdt

'    Set NPay_Result = fnCallNaverPayCashAmt(ipaygatetid)
'    rw ":"&NPay_Result.body.totalCashAmount
'    rw ":"&NPay_Result.body.primaryPayMeans
'    Set NPay_Result = Nothing
'    response.end

    Set NPay_Result = fnCallNaverPayCancel(ipaygatetid,irefundrequire,"customerReq",nPayCancelRequester)

    if NPay_Result.code="Success" then
        iResultCode = "00"                      ''00 ����Ұ�.
        iResultMsg = replace(NPay_Result.message,"'","")
        if (iResultMsg="") then iResultMsg = NPay_Result.code

        'rw NPay_Result.body.paymentId
        'rw NPay_Result.body.payHistId
        'rw NPay_Result.body.primaryPayMeans

		''rw NPay_Result.body.primaryPayCancelAmount
		''rw NPay_Result.body.primaryPayRestAmount
		''rw NPay_Result.body.npointCancelAmount
		''rw NPay_Result.body.npointRestAmount
		''rw NPay_Result.body.totalRestAmount

        cancelYmdt = NPay_Result.body.cancelYmdt

        iCancelDate = LEFT(cancelYmdt,4)&"-"&MID(cancelYmdt,5,2)&"-"&MID(cancelYmdt,7,2)
        iCancelTime = MID(cancelYmdt,9,2)&":"&MID(cancelYmdt,11,2)&":"&MID(cancelYmdt,13,2)

    else
        iResultCode = NPay_Result.code
        iResultMsg = replace(NPay_Result.message,"'","")
    end if

    Set NPay_Result = Nothing

end function

''KaKao �ſ�ī�� ���
function CanCelNewKakaoPay(ipaygatetid, irefundrequire, irdSite, byref iretval, byref iResultCode, byref iResultMsg, byref iCancelDate, byref iCancelTime)
    Dim objKMPay, cancelYmdt, Status

    Set objKMPay = fnCallKakaoPayCancel(ipaygatetid, irefundrequire, Status)

    if Status="200" then
        iResultCode = "00"                                  ''00 ����Ұ�.
        cancelYmdt = objKMPay.canceled_at                 ''���� ��� �ð�
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = objKMPay.status                        ''�������°�
    else
        iResultCode = objKMPay.code                        ''�����ڵ�
        iResultMsg = objKMPay.message                      ''���� �޼���
    end if

    Set objKMPay = Nothing

end function

''Toss ī��/�ǽð� ���
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

''�������� ���
function CanCelChaiPay(ipaygatetid, idempotencyKey, irefundrequire, irdSite, byref iretval, byref iResultCode, byref iResultMsg, byref iCancelDate, byref iCancelTime)
    Dim objKMPay, cancelYmdt, Status

    Set objKMPay = fnCallChaiPayCancel(ipaygatetid, idempotencyKey, irefundrequire, Status)

    if Status="200" then
        iResultCode = "00"                                  ''00 ����Ұ�.
        cancelYmdt = objKMPay.updatedAt                 ''���� ��� �ð�
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = objKMPay.status                        ''�������°�
    else
        iResultCode = objKMPay.code                        ''�����ڵ�
        iResultMsg = objKMPay.message                      ''���� �޼���
    end if

    Set objKMPay = Nothing

end function

''KaKao �ſ�ī�� ���
function CanCelKakaoPay(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime)
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

    ''objKMPay.AddRequestData "Amt", irefundrequire
	objKMPay.AddRequestData "CancelAmt", irefundrequire

    ''objKMPay.AddRequestData "SupplyAmt",0     ''���ް�
    ''objKMPay.AddRequestData "GoodsVat",0      ''�ΰ���
    ''objKMPay.AddRequestData "ServiceAmt",0    ''�����
    objKMPay.AddRequestData "CancelMsg","����û"
    objKMPay.AddRequestData "PartialCancelCode","0"     '' 0��ü���, 1�κ����.
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
    tid = objKMPay.GetResultData("TID")               	' TID
    errorCD = objKMPay.GetResultData("ErrorCD")        	' �� �����ڵ�
    errorMsg = objKMPay.GetResultData("ErrorMsg")      	' �� �����޽���
    authDate = cancelDate & cancelTime									' �ŷ��ð�
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

''������ �޴��� �����
function CanCelMobileDacom(ipaygatetid,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime)
    Dim CST_PLATFORM, CST_MID, LGD_MID, LGD_TID,Tradeid, LGD_CANCELREASON, LGD_CANCELREQUESTER, LGD_CANCELREQUESTERIP
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

    LGD_CANCELREASON        = "����û"                        ' ��һ���
    LGD_CANCELREQUESTER     = "��"                            ' ��ҿ�û��
    LGD_CANCELREQUESTERIP   = Request.ServerVariables("REMOTE_ADDR")     ' ��ҿ�ûIP


    configPath           = "C:/lgdacom" ''"C:/lgdacom/conf/"&CST_MID         				' LG���÷������� ������ ȯ������("/conf/lgdacom.conf") ��ġ ����.
    Set xpay = CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
    xpay.Set "LGD_CANCELREASON", LGD_CANCELREASON
    xpay.Set "LGD_CANCELREQUESTER", LGD_CANCELREQUESTER
    xpay.Set "LGD_CANCELREQUESTERIP", LGD_CANCELREQUESTERIP

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

'''�ſ�ī�� �κ���� R120 => �ٸ� ���������� ���� ó��.
'''�ڵ��� �κ���� R420 => �ٸ� ���������� ���� ó��.

dim id, finishuserid, msg, force
dim orgOrderSerial, chgOrderserial
dim jumundiv, accountdiv, pggubun, pAddParam

id           = RequestCheckVar(request("id"),10)
msg          = RequestCheckVar(request("msg"),50)
finishuserid = session("ssBctID")
force = RequestCheckVar(request("force"),10)

if (msg="") and (IsAutoScript) then msg="��������"

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
        response.write "S_ERR|ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�."
    else
        response.write "<script>alert('ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|���� ���°� �ƴմϴ�."
    else
        response.write "<script>alert('���� ���°� �ƴմϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

'' �ſ�ī�� ��Ҹ� ����
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('���� �ſ�ī�� �ŷ��� ��� �����մϴ�.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if Not ((returnmethod="R100") or (returnmethod="R020") or (returnmethod="R400") or (returnmethod="R150")) Then
    if (IsAutoScript) then
        response.write "S_ERR|�ſ�ī�� ��ü���, �ǽð���ü ���, �޴��� ��ü ���, �̴Ϸ�Ż ��ü ��Ҹ� ����."
    else
        response.write "<script>alert('�ſ�ī�� ��ü���, �ǽð���ü ���, �޴��� ��ü ���, �̴Ϸ�Ż ��ü ��Ҹ� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


''=============��ü��Ҹ� ������.. �κ���ҵ� ��Ҿȵ�..=============
dim sqlStr, isSameMoney
dim t_refundrequire, t_MaybeOrgPayPrice
isSameMoney = false

''���̳ʽ� �ֹ��ϰ�� ���ֹ���ȣ// ===> ���ֹ������� �����..
sqlStr = " select r.refundrequire, m.orderserial, m.jumundiv, m.linkorderserial"
sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list l"
sqlStr = sqlStr & " 	Join db_cs.dbo.tbl_as_refund_info r"
sqlStr = sqlStr & " 	on l.id=r.asid"
sqlStr = sqlStr & " 	and r.returnmethod  in ('R100','R020','R400')"
sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m"
sqlStr = sqlStr & " 	on l.orderserial=m.orderserial"
sqlStr = sqlStr & " where l.id="&id
sqlStr = sqlStr & " and l.divcd='A007'"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_refundrequire=rsget("refundrequire")
    ''if (rsget("jumundiv")="9") then
    ''    orgOrderserial = rsget("linkorderserial")
    ''else
        orgOrderserial = rsget("orderserial")
    ''end if
end if
rsget.Close


sqlStr = " select top 1 m.jumundiv, m.accountdiv, IsNull(m.pggubun,'') as pggubun, IsNull(e.pAddParam,'') as pAddParam"
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
	accountdiv = rsget("accountdiv")
	pggubun = rsget("pggubun")
	pAddParam = rsget("pAddParam")
end if
rsget.close

'// ��ȯ�ֹ�( jumundiv = 6 )�̸� ���ֹ����� �������� �����´�.
if (jumundiv = "6") then
	sqlStr = " select top 1 c.orgorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.chgorderserial = '" & orgOrderserial & "' "
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		chgOrderserial = orgOrderserial
		orgOrderserial = rsget("orgorderserial")
	end if
	rsget.close

	''2017/10/24 �߰�
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


'''2011-04 ���� tbl_order_paymentEtc ���.
sqlStr = " select Sum(acctamount) as acctamount"
sqlStr = sqlStr & " from db_order.dbo.tbl_order_paymentEtc"
sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"
sqlStr = sqlStr & " and acctdiv in ('100','110','120','20','400','150')"    ''�ſ�ī�� �� OkCashBag�� ���̰�����. (2016/07/20 120 �߰�, 2016/08/04 20(�ǽð���ü),400(�޴���) �߰�)
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    t_MaybeOrgPayPrice=rsget("acctamount")
    isSameMoney    = (t_refundrequire=(t_MaybeOrgPayPrice))
end if
rsget.Close

IF  (Not isSameMoney) THEN
    IF (force="on") then
        response.write "��ұݾװ� ���ݾ� ����<br><br>"
    ELSE
        if (IsAutoScript) then
            response.write "S_ERR|��ұݾװ� ���ݾ� ����"
        else
            response.write "<script>alert('��ұݾװ� ���ݾ� ���� - ������ ���� ���."&t_refundrequire&":"&t_MaybeOrgPayPrice&"');</script>"
            response.write "<script>window.close();</script>"
        end if
        dbget.close()	:	response.End
    End IF
END IF
'''=================================================================


'' IniPay �� ��Ҹ� ����
dim IsInicisTID : IsInicisTID = False
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="IniTechPG_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_CARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_ISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtCARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,6)="Stdpay")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="StdpayRTPY")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_RTPY")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIAPICARD")   ' ��������
''IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_AUTH")

if (pggubun <> "TS") and (pggubun <> "CH") and ((MID(orefund.FOneItem.FpaygateTid,9,2) <> "NP") and Not(Left(orefund.FOneItem.FpaygateTid,1)="T" and Len(orefund.FOneItem.FpaygateTid)>=20) and Left(orefund.FOneItem.FpaygateTid,3)<>"cns") and (Left(orefund.FOneItem.FpaygateTid,5)<>"KCTEN") and Not IsInicisTID AND orefund.FOneItem.Freturnmethod<>"R400" and Not IsNumeric(orefund.FOneItem.FpaygateTid) then
    if (IsAutoScript) then
        response.write "S_ERR|�̴Ͻý� �ŷ��� ��� �����մϴ�."
    else
        response.write "<script>alert('�̴Ͻý� �ŷ��� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'' Pg_Mid
dim MctID
MctID = Mid(orefund.FOneItem.FpaygateTid,11,10)
'' response.write MctID

dim INIpay, PInst
dim ResultCode, ResultMsg, CancelDate, CancelTime, Rcash_cancel_noappl

''�޴��� ���� �߰� 2015/04/21 IsINIMobile
Dim IsINIMobile : IsINIMobile = false
if (orefund.FOneItem.Freturnmethod = "R400") and (Len(orefund.FOneItem.FpaygateTid)=40) then
    IsINIMobile = (LEFT(orefund.FOneItem.Fpaygatetid,LEN("IniTechPG_"))="IniTechPG_") or (LEFT(orefund.FOneItem.Fpaygatetid,LEN("INIMX_HPP_"))="INIMX_HPP_") or (LEFT(orefund.FOneItem.Fpaygatetid,LEN("StdpayHPP_"))="StdpayHPP_")
end if

Dim IsDacomMobile : IsDacomMobile = false
if (orefund.FOneItem.Freturnmethod = "R400") and (NOT IsINIMobile) then
    if (Len(orefund.FOneItem.FpaygateTid)>=31) then
        IsDacomMobile = True        ''46~49 Tradeid(23) & "|" & vTID(24)  => 263055|tenby2014031117203148569 (31)
    else
        IsDacomMobile = False       ''32~35 Tradeid(23) & "|" & vTID(10)
    end if
end if

''īī������
Dim IsKakaoPay : IsKakaoPay = (pggubun = "KA") '' (orefund.FOneItem.Freturnmethod = "R100") and ((Left(orefund.FOneItem.FpaygateTid,3)="cns") or (Left(orefund.FOneItem.FpaygateTid,5)="KCTEN")) ''�ϴ�.
''New īī������
Dim IsNewKakaoPay : IsNewKakaoPay = (pggubun = "KK")

''TOSS����
Dim IsTossPay : IsTossPay = (pggubun = "TS")

''���̹�����
Dim IsNaverPay : IsNaverPay = (pggubun = "NP") ''((MID(orefund.FOneItem.FpaygateTid,9,2) = "NP") and (LEN(orefund.FOneItem.FpaygateTid)=20))

''������
Dim IsPayco : IsPayco = (pggubun = "PY")

''��������
Dim IsChaiPay : IsChaiPay = (pggubun = "CH")

'############################################################## �ڵ��� ���� ��� ##############################################################
If (orefund.FOneItem.Freturnmethod = "R400") and (NOT IsINIMobile) Then

    Dim McashCancelObj, Mrchid, Svcid, Tradeid, Prdtprice, Mobilid, retval


    IF (IsDacomMobile) then
        CALL CanCelMobileDacom(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,Request("rdsite"),retval,ResultCode,ResultMsg,CancelDate,CancelTime)
    ELSE
        '' Not Using MCash
        dim dummi : dummi=1/0
        dbget.close() : response.end


    	Set McashCancelObj = Server.CreateObject("Mcash_Cancel.Cancel.1")

    	Mrchid      = "10030289"
    	If LEFT(Request("rdsite"),6) = "mobile" Then
    		Svcid       = "100302890002"
    	Else
    		Svcid       = "100302890001"
    	End If
    	Tradeid     = Split(orefund.FOneItem.FpaygateTid,"|")(0)
    	Prdtprice   = orefund.FOneItem.Frefundrequire
    	Mobilid     = Split(orefund.FOneItem.FpaygateTid,"|")(1)

    	McashCancelObj.Mrchid			= Mrchid
    	McashCancelObj.Svcid			= Svcid
    	McashCancelObj.Tradeid			= Tradeid
    	McashCancelObj.Prdtprice		= Prdtprice
    	McashCancelObj.Mobilid	        = Mobilid

    	retval = McashCancelObj.CancelData

    	set McashCancelObj = nothing

    	If retval = "0" Then
    		ResultCode 	= "00"
    		ResultMsg	= "����ó��"
    	Else
    		ResultCode = retval
    		Select Case ResultCode
    			Case "14"
    				ResultMsg = "����"
    			Case "20"
    				ResultMsg = "�޴��� ������� ����(PG��) (LGT�� ��� ������������濡 ���� ��������)"
    			Case "41"
    				ResultMsg = "�ŷ����� ������"
    			Case "42"
    				ResultMsg = "��ұⰣ���"
    			Case "43"
    				ResultMsg = "���γ������� ( ������������ ����ġ, ���ι�ȣ ��ȿ�ð� �ʰ�( 3�� ) )"
    			Case "44"
    				ResultMsg = "�ߺ� ��� ��û"
    			Case "45"
    				ResultMsg = "��� ��û �� ��� ���� ����ġ"
    			Case "97"
    				ResultMsg = "��û�ڷ� ����"
    			Case "98"
    				ResultMsg = "��Ż� ��ſ���"
    			Case "99"
    				ResultMsg = "��Ÿ"
    			Case "11"
    				ResultMsg = "��������������� ���� ��ҺҰ�(11)"
    			Case Else
    				ResultMsg = ""
    		End Select
    	End If

    	CancelDate	= year(now) & "�� " & month(now) & "�� " & day(now) & "��"
    	CancelTime	= hour(now) & "�� " & minute(now) & "�� " & second(now) & "��"
    END IF
ELSEIF (IsKakaoPay) then
    CALL CanCelKakaoPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime)
ELSEIF (IsNewKakaoPay) then
    CALL CanCelNewKakaoPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime)
ELSEIF (IsTossPay) then
    CALL CancelTossPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,(CStr(ocsaslist.FOneItem.Forderserial) & "_" & CStr(ocsaslist.FOneItem.Fid)),ResultCode,ResultMsg,CancelDate,CancelTime)
ELSEIF (IsNaverPay) then
    dim nPayCancelRequester : nPayCancelRequester = fnGetNPayCancelRequester(id)
    CALL CanCelNaverPay(orefund.FOneItem.FpaygateTid,orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime,nPayCancelRequester)
ELSEIF (IsPayco) then
    CALL CancelPayco(orefund.FOneItem.FpaygateTid, orefund.FOneItem.Frefundrequire, "", retval, ResultCode, ResultMsg, CancelDate, CancelTime,pAddParam)
ELSEIF (IsChaiPay) then
    CALL CanCelChaiPay(orefund.FOneItem.FpaygateTid,fnGetChaiPayIdempotencyKey(ocsaslist.FOneItem.Forderserial),orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime)
Else
'############################################################## ī��, �ǽð� ���� ��� ##############################################################
		'###############################################################################
		'# 1. ��ü ���� #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")


		'###############################################################################
		'# 2. �ν��Ͻ� �ʱ�ȭ #
		'######################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. �ŷ� ���� ���� #
		'#####################
		INIpay.SetActionType CLng(PInst), "CANCEL"

		'###############################################################################
		'# 4. ���� ���� #
		'################
		INIpay.SetField CLng(PInst), "pgid", "IniTechPG_" 'PG ID (����)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '���� PG IP (����)
		INIpay.SetField CLng(PInst), "mid", MctID '�������̵�
		INIpay.SetField CLng(PInst), "admin", "1111" 'Ű�н�����(�������̵� ���� ����)
		INIpay.SetField CLng(PInst), "tid", Request("tid") '����� �ŷ���ȣ(TID)
		INIpay.SetField CLng(PInst), "msg", msg '��� ����
		INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
		INIpay.SetField CLng(PInst), "debug", "false" '�α׸��("true"�� �����ϸ� ���� �α׸� ����)
		INIpay.SetField CLng(PInst), "merchantreserved", "����" '����

		'###############################################################################
		'# 5. ��� ��û #
		'################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'# 6. ��� ��� #
		'################
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '����ڵ� ("00"�̸� ��Ҽ���)
		ResultMsg  = INIpay.GetResult(CLng(PInst), "resultmsg") '�������
		CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '�̴Ͻý� ��ҳ�¥
		CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '�̴Ͻý� ��ҽð�
		Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '���ݿ����� ��� ���ι�ȣ

		'###############################################################################
		'# 7. �ν��Ͻ� ���� #
		'####################
		INIpay.Destroy CLng(PInst)
End If



dim itemCnt, itemName, refunddepositsum, refundmileagesum, refundgiftcardsum, refundstr, tmpgubun, fullText, failText
dim refundrequire,refundresult,userid
dim iorderserial, ibuyhp
dim contents_finish

contents_finish = "��� " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
contents_finish = contents_finish & "����Ͻ� : " & CancelDate & " " & CancelTime & VbCrlf
contents_finish = contents_finish & "����� ID " & finishuserid

if ((ResultCode="00") or (ResultCode="0000")) or (IsKakaoPay and (resultCode="2001")) then

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

	'// OKĳ�ù� ������ ���, ��ǰ �� ���̳ʽ� �ֹ� �Է� �� ī�� ��ü����̸� ���̳ʽ� �ֹ��� ���������ݾ� �Է�
	if (accountdiv="110") then ''2015/08/05
        sqlStr = " exec [db_order].[dbo].[usp_Ten_AddEtcPaymentWhenCardCancel] '" + CStr(orgOrderserial) + "', '" + CStr(chgOrderserial) + "'"
        dbget.Execute sqlStr
    end if

    Call AddCustomerOpenContents(id, "ȯ��(���) �Ϸ�: " & CStr(refundrequire))


    Call FinishCSMaster(id, finishuserid, contents_finish)

    ''���� ��� ��û SMS �߼�
    if (iorderserial<>"") and (ibuyhp<>"") then
        'SendAcctCancelMsg ibuyhp, iorderserial

		itemCnt=0
		itemName=""
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

		refundresult=0
		refunddepositsum=0
		refundmileagesum=0
		refundgiftcardsum=0
		refundrequire=0
		refundstr=""
		returnmethod=""
		tmpgubun=""
		sqlStr ="select returnmethod, refundresult, refunddepositsum, refundmileagesum, refundgiftcardsum, refundrequire"
		sqlStr = sqlStr & " from db_cs.dbo.tbl_as_refund_info r with (nolock)"
		sqlStr = sqlStr & " where asid="& id &""

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		IF not rsget.EOF THEN
			returnmethod = rsget("returnmethod")
			refundresult = rsget("refundresult")
			refunddepositsum = rsget("refunddepositsum")
			refundmileagesum = rsget("refundmileagesum")
			refundgiftcardsum = rsget("refundgiftcardsum")
			refundrequire = rsget("refundrequire")
		END IF
		rsget.close

        if refunddepositsum=0 and refundmileagesum=0 and refundgiftcardsum=0 then
            refundstr=FormatNumber(refundresult,0) & "��"
        else
            refundstr=FormatNumber(refundresult,0) & "��(��ġ��ȯ�� "& refunddepositsum &"�� / ���ϸ���ȯ�� "& refundmileagesum &"pt / ����Ʈȯ�� "& refundgiftcardsum &"��)"
        end if

		' ��ü���. īī���� �˸��� �߼�.   ' 2021.10.13 �ѿ�� ����
		fullText = "[10x10] ��������ȳ�" & vbCrLf & vbCrLf
		fullText = fullText & "����, �ֹ���Ұ� �Ϸ�Ǿ����ϴ�." & vbCrLf & vbCrLf
		fullText = fullText & "�� �ֹ���ȣ : "& iorderserial &"" & vbCrLf
		fullText = fullText & "�� ��ǰ�� : "& itemName &"" & vbCrLf
		fullText = fullText & "�� ��ұݾ� : "& refundstr &""
		failText = "[�ٹ�����]�ֹ���Ұ� �Ϸ�Ǿ����ϴ�.�ֹ���ȣ : "& iorderserial &""
		Call SendKakaoCSMsg_LINK("", ibuyhp,"1644-6030","KC-0021",fullText,"SMS","",failText,"",iorderserial,"")
    end if

    ''����
    Call SendCsActionMail(id)

    if (IsAutoScript) then
        response.write "S_OK"
    else
        response.write "<script>alert('" & ResultMsg & "');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End

else
    if (IsAutoScript) then
        response.write "S_ERR|"&ResultMsg
    else
        response.write ResultCode & "<br>"
        response.write ResultMsg & "<br>"
        response.write CancelDate & "<br>"
        response.write CancelTime & "<br>"
        response.write Rcash_cancel_noappl & "<br>"
    end if
end if
%>



<%
set ocsaslist = Nothing
set orefund = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
