<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 부분취소
' History : 이상구 생성
'			2021.09.29 한용민 수정(알림톡 발송 추가)
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

''취소요청자(1:구매자,2:가맹점관리자)
function fnGetNPayCancelRequester(iid)
    dim sqlStr , buf
    fnGetNPayCancelRequester = "2" ''기본 가맹점 관리자.

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

    if LEFT(buf,3)="[고객" then
        fnGetNPayCancelRequester = "1"
    end if
end function

''차이페이용 idempotencyKey(가맹점 주문 번호 (고유번호, 중복 결제 방지 등))
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

''페이코 취소(부분취소)
function PartialCancelPayco(ipaygatetid, iremainAmount, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byRef inewTid,byref iprimaryPayRestAmount, byref inpointRestAmount, byref itotalRestAmount, byVal orderCertifyKey)
	dim Payco_Result , cancelYmdt
	Set Payco_Result = fnCallPaycoPartialCancel(ipaygatetid, iremainAmount, irefundrequire, "customerReq_P", orderCertifyKey)

	if (Payco_Result.code = 0) then
		iResultCode = "00"                      ''00 사용할것.
		iResultMsg = replace(Payco_Result.message, "'", "")
		cancelYmdt = Payco_Result.result.cancelYmdt
		iCancelDate = LEFT(cancelYmdt,4)&"-"&MID(cancelYmdt,5,2)&"-"&MID(cancelYmdt,7,2)
		iCancelTime = MID(cancelYmdt,9,2)&":"&MID(cancelYmdt,11,2)&":"&MID(cancelYmdt,13,2)

		inewTid = Payco_Result.result.cancelTradeSeq	'// newTid ???

		iprimaryPayRestAmount = 0						'// 안쓰임
        inpointRestAmount = 0
        itotalRestAmount = 0
	else
		iResultCode = Payco_Result.code
		iResultMsg = replace(Payco_Result.message, "'", "")
	end if

	Set Payco_Result = Nothing
end function


''네이버페이 취소(부분취소)
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

''New KaKao 신용카드 부분 취소
function PartialCancelNewKakaoPay(ipaygatetid, imainAmount, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byRef inewTid)

    Dim objKMPay, cancelYmdt, Status

    Set objKMPay = fnCallKakaoPayPartialCancel(ipaygatetid, imainAmount, irefundrequire, Status)

    if Status="200" then
        iResultCode = "00"                                  ''00 사용할것.
        cancelYmdt = objKMPay.canceled_at                 ''결제 취소 시각
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = objKMPay.status                        ''결제상태값
    else
        iResultCode = objKMPay.code                        ''실패코드
        iResultMsg = objKMPay.message                     ''실패 메세지
    end if

    Set objKMPay = Nothing

end function

function CancelTossPay(ipaygatetid, irefundrequire, irefundNo, byref iResultCode, byref iResultMsg, byref iCancelDate, byref iCancelTime)
	''irefundNo = 주문번호_환불ASID
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
		'// 환불성공
		iResultCode = "00"                                  ''00 사용할것.
		iResultMsg = ""
        iCancelDate = LEFT(rstJson.data("approvalTime"),10)
        iCancelTime = RIGHT(rstJson.data("approvalTime"),8)
	else
        iResultCode = CStr(rstJson.data("code"))                        ''실패코드
        iResultMsg = CStr(rstJson.data("msg"))                          ''실패 메세지
	end if
end function

''차이페이 부분 취소
function PartialCancelChaiPay(ipaygatetid, idempotencyKey, imainAmount, irefundrequire, irdSite, byREF iretval, byREF iResultCode, byREF iResultMsg, byREF iCancelDate, byREF iCancelTime, byRef inewTid)

    Dim objKMPay, cancelYmdt, Status

    Set objKMPay = fnCallChaiPayPartialCancel(ipaygatetid, idempotencyKey, imainAmount, irefundrequire, Status)

    if Status="200" then
        iResultCode = "00"                                  ''00 사용할것.
        cancelYmdt = objKMPay.updatedAt                  ''결제 취소 시각
        iCancelDate = LEFT(cancelYmdt,10)
        iCancelTime = RIGHT(cancelYmdt,8)
        iResultMsg = objKMPay.status                        ''결제상태값
    else
        iResultCode = objKMPay.code                        ''실패코드
        iResultMsg = objKMPay.message                     ''실패 메세지
    end if

    Set objKMPay = Nothing

end function

''KaKao 신용카드 취소
function PartialCancelKakaoPay(ipaygatetid, iremainAmount, irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime, byRef inewTid)

	''iremainAmount, inewTid

    Dim objKMPay

	dim otime,orgTim,diffTime
	otime = Timer()
	orgTim = otime

    '1) 객체 생성
    Set objKMPay = Server.CreateObject("LGCNS.CNSPayService.CnsPayWebConnector")
    objKMPay.RequestUrl = CNSPAY_DEAL_REQUEST_URL

    '2) 로그 정보
    objKMPay.SetCnsPayLogging KMPAY_LOG_DIR, KMPAY_LOG_LEVEL	'-1:로그 사용 안함, 0:Error, 1:Info, 2:Debug

    '3) 요청 페이지 파라메터 셋팅
    objKMPay.AddRequestData "MID", KMPAY_MERCHANT_ID
    objKMPay.AddRequestData "TID", ipaygatetid
    objKMPay.AddRequestData "CancelAmt", irefundrequire
	objKMPay.AddRequestData "CheckRemainAmt", iremainAmount
    ''objKMPay.AddRequestData "SupplyAmt",0     ''공급가
    ''objKMPay.AddRequestData "GoodsVat",0      ''부가세
    ''objKMPay.AddRequestData "ServiceAmt",0    ''봉사료
    objKMPay.AddRequestData "CancelMsg","고객요청"
    objKMPay.AddRequestData "PartialCancelCode","1"     '' 0전체취소, 1부분취소.
    objKMPay.AddRequestData "PayMethod","CARD"

    '4) 추가 파라메터 셋팅
    objKMPay.AddRequestData "actionType", "CL0"  															' actionType : CL0 취소, PY0 승인, CI0 조회
    objKMPay.AddRequestData "CancelIP", Request.ServerVariables("LOCAL_ADDR")	' 가맹점 고유 ip
    objKMPay.AddRequestData "CancelPwd", KMPAY_CANCEL_PWD														' 취소 비밀번호 설정

    '5) 가맹점키 셋팅 (MID 별로 틀림)
    objKMPay.AddRequestData "EncodeKey", KMPAY_MERCHANT_KEY

	diffTime = FormatNumber(Timer()-otime,4)
	rw diffTime

    '6) CNSPAY Lite 서버 접속하여 처리
    objKMPay.RequestAction
	rw diffTime

    '7) 결과 처리
    Dim resultCode, resultMsg, cancelAmt, cancelDate, cancelTime, payMethod, resMerchantId, tid, errorCD, errorMsg, authDate, ccPartCl, stateCD

    resultCode = objKMPay.GetResultData("ResultCode") 	' 결과코드 (정상 :2001(취소성공), 2002(취소진행중), 그 외 에러)
    resultMsg = objKMPay.GetResultData("ResultMsg")   	' 결과메시지
    cancelAmt = objKMPay.GetResultData("CancelAmt")   	' 취소금액
    cancelDate = objKMPay.GetResultData("CancelDate") 	' 취소일
    cancelTime = objKMPay.GetResultData("CancelTime")   ' 취소시간
    payMethod = objKMPay.GetResultData("PayMethod")   	' 취소 결제수단
    resMerchantId = objKMPay.GetResultData("MID")     	' 가맹점 ID
    inewTid = objKMPay.GetResultData("TID")             ' TID
    errorCD = objKMPay.GetResultData("ErrorCD")        	' 상세 에러코드
    errorMsg = objKMPay.GetResultData("ErrorMsg")      	' 상세 에러메시지
    authDate = cancelDate & cancelTime					' 거래시간
    ccPartCl = objKMPay.GetResultData("CcPartCl")       ' 부분취소 가능여부 (0:부분취소불가, 1:부분취소가능)
    stateCD = objKMPay.GetResultData("StateCD")         ' 거래상태코드 (0: 승인, 1:전취소, 2:후취소)

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

''데이콤 휴대폰 실부분취소
function PartialCanCelMobileDacom(ipaygatetid, iremainAmount,irefundrequire,irdSite,byREF iretval,byREF iResultCode,byREF iResultMsg,byREF iCancelDate,byREF iCancelTime, byRef inewTid)
    Dim CST_PLATFORM, CST_MID, LGD_MID, LGD_TID,Tradeid, LGD_CANCELREASON, LGD_CANCELREQUESTER, LGD_CANCELREQUESTERIP
	Dim LGD_CANCELAMOUNT, LGD_REMAINAMOUNT
    Dim configPath, xpay

    IF (application("Svr_Info") = "Dev") THEN                   ' LG유플러스 결제서비스 선택(test:테스트, service:서비스)
		CST_PLATFORM = "test"
	Else
		CST_PLATFORM = "service"
	End If

    CST_MID              = "tenbyten02"                         ' LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요. //모바일, 서비스 동일.
                                                                ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                               ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    Tradeid     = Split(ipaygatetid,"|")(0)
	LGD_TID     = Split(ipaygatetid,"|")(1)                     ' LG유플러스으로 부터 내려받은 거래번호(LGD_TID) : 24 byte

	LGD_CANCELAMOUNT		= irefundrequire					' 부분취소 금액
	LGD_REMAINAMOUNT		= iremainAmount						' 취소전 남은 금액

    LGD_CANCELREASON        = "고객요청"                        ' 취소사유
    LGD_CANCELREQUESTER     = "고객"                            ' 취소요청자
    LGD_CANCELREQUESTERIP   = Request.ServerVariables("REMOTE_ADDR")     ' 취소요청IP


    configPath           = "C:/lgdacom" ''"C:/lgdacom/conf/"&CST_MID         				' LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.
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
    ' * 1. 결제취소 요청 결과처리
    ' *
    ' * 취소결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
	' *
	' * [[[중요]]] 고객사에서 정상취소 처리해야할 응답코드
	' * 1. 신용카드 : 0000, AV11
	' * 2. 계좌이체 : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (환불진행중 응답-> 환불결과코드.xls 참고)
	' * 3. 나머지 결제수단의 경우 0000(성공) 만 취소성공 처리
	' *
    ' */

    if xpay.TX() then
        '1)결제취소결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
'Response.Write("결제취소 요청이 완료되었습니다. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

        iretval = "0"
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg

		inewTid = ""

		'' response.write xpay.Tid

		'' rw 	xpay.Response("LGD_TID",0) & "aaaaa"

		'' '' '아래는 결제요청 결과 파라미터를 모두 찍어 줍니다.
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
        '2)API 요청 실패 화면처리
'Response.Write("결제취소 요청이 실패하였습니다. <br>")
'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
        iResultCode = xpay.resCode
		iResultMsg	= xpay.resMsg
    end if

    iCancelDate	= year(now) & "년 " & month(now) & "월 " & day(now) & "일"
	iCancelTime	= hour(now) & "시 " & minute(now) & "분 " & second(now) & "초"

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

if (msg="") and (IsAutoScript) then msg="배송전 부분취소"


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
        response.write "S_ERR|환불내역이 없거나 유효하지 않은 내역입니다.[0]"
    else
        response.write "<script>alert('환불내역이 없거나 유효하지 않은 내역입니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    if (IsAutoScript) then
        response.write "S_ERR|접수 상태가 아닙니다.[1]"
    else
        response.write "<script>alert('접수 상태가 아닙니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


orderserial = ocsaslist.FOneItem.FOrderserial

Dim returnmethod, IsCardPartialCancel
returnmethod = orefund.FOneItem.Freturnmethod

if (returnmethod<>"R120") and (returnmethod<>"R420") and (returnmethod<>"R022") Then
    if (IsAutoScript) then
        response.write "S_ERR|신용카드, 실시간이체, 휴대폰 부분취소만 가능합니다.[2]"
    else
        response.write "<script>alert('신용카드, 휴대폰 부분취소만 가능합니다.');</script>"
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
        response.write "S_ERR|데이콤 핸드폰결제만 취소 가능합니다.[3]"
    else
        response.write "<script>alert('데이콤 핸드폰결제만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'==============================================================================
'최초 주결제수단금액
dim accountdiv : accountdiv ="100"
dim omainpayment, mainpaymentorg
dim cardcancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcode

if (returnmethod = "R400") or (returnmethod = "R420") then
	accountdiv = "400"
elseif (returnmethod = "R022") then  ''2016/06/21 추가.
    accountdiv = "20"
end if


set omainpayment = new COrderMaster

mainpaymentorg = 0			'// 최초 결제금액
cardcancelok = "N"
cardcancelerrormsg = ""
cardcancelcount = 0
cardcancelsum   = 0			'// 취소금액합계
cardcode = ""

if (orderserial<>"") then
	omainpayment.FRectOrderSerial = orderserial

	''response.write accountdiv & "<br />"
	if ((accountdiv = "100") or (accountdiv = "20")) then ''네이버페이 실시간 이체 부분취소 가능
		Call omainpayment.getMainPaymentInfo(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	else
		Call omainpayment.getMainPaymentInfoPhone(accountdiv, mainpaymentorg, cardcancelok, cardcancelerrormsg, cardcancelcount,cardcancelsum, cardcode)
	end if

	orgOrderserial = orderserial

	'// 교환주문( jumundiv = 6 )이면 원주문에서 결제정보 가져온다.
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

		''2017/09/21 추가
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

''부분취소 횟수 검토..? ==>
''CHECK

''부분취소 가능한지..

'''통신전 로그 저장.
'''fnINIrePay(MctID, ioldtid, iCancelPrice, iconfirm_price, ibuyeremail, byRef ioldtid, byRef ResultCode, byRef ResultMsg,byRef OldTid, byRef CancelPrice, byRef RepayPrice, byref CntRepay)

'' Pg_Mid
dim MctID
MctID = Mid(orefund.FOneItem.FpaygateTid,11,10)     '''MayBe teenxteenN or INIPayTest
'' response.write MctID

''''원거래ID,취소금액    , 재승인금액,  구매자email
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
'''결과 값.
dim Tid
dim OldTid, CancelPrice, RepayPrice, CntRepay


IF (ioldtid="") or (iCancelPrice<1) or (iconfirm_price<0) THEN
    if (IsAutoScript) then
        response.write "S_ERR|부분 취소금액 오류 또는 TID오류[5]"
    else
        response.write "<script>alert('부분 취소금액 오류 또는 TID오류"&iCancelPrice&"."&iconfirm_price&"');</script>"
        ''response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
END IF

''통신전 저장
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

''카카오페이
Dim IsKakaoPay : IsKakaoPay = (pggubun = "KA")

''New 카카오페이
Dim IsNewKakaoPay : IsNewKakaoPay = (pggubun = "KK")

''TOSS페이
Dim IsTossPay : IsTossPay = (pggubun = "TS")

''네이버페이
Dim iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount
Dim IsNaverPay : IsNaverPay = (pggubun = "NP")

''페이코
Dim IsPayco : IsPayco = (pggubun = "PY")

''차이페이
Dim IsChaiPay : IsChaiPay = (pggubun = "CH")

'' IniPay 만 취소만 가능
if (pggubun <> "TS") and (pggubun <> "CH") and ((MID(orefund.FOneItem.FpaygateTid,9,2) <> "NP") and Not(Left(orefund.FOneItem.FpaygateTid,1)="T" and Len(orefund.FOneItem.FpaygateTid)>=20) and Left(orefund.FOneItem.FpaygateTid,3)<>"cns") and (Left(orefund.FOneItem.FpaygateTid,5)<>"KCTEN") and (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIMX_CARD") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIMX_ISP_") AND (Left(orefund.FOneItem.FpaygateTid,6)<>"Stdpay") AND (Left(orefund.FOneItem.FpaygateTid,10)<>"INIAPICARD") AND orefund.FOneItem.Freturnmethod<>"R400" AND orefund.FOneItem.Freturnmethod<>"R420" and Not IsNumeric(orefund.FOneItem.FpaygateTid) then
    if (IsAutoScript) then
        response.write "S_ERR|이니시스 거래만 취소 가능합니다.[6]"
    else
        response.write "<script>alert('이니시스 거래만 취소 가능합니다..');</script>"
        response.write "<script>window.close();</script>"
    end if
    dbget.close()	:	response.End
end if


'############################################################## 핸드폰 결제 취소 ##############################################################
If orefund.FOneItem.Freturnmethod = "R420" Then

    Dim McashCancelObj, Mrchid, Svcid, Tradeid, Prdtprice, Mobilid, retval

	CALL PartialCanCelMobileDacom(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire,Request("rdsite"),retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'' '// ResultCode 가 4바이트라서 2바이트로 변경한다.
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

	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsNewKakaoPay) then
	CALL PartialCancelNewKakaoPay(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
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

	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsPayco) then
	CALL PartialCancelPayco(orefund.FOneItem.FpaygateTid, (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire, "", retval, ResultCode, ResultMsg, CancelDate, CancelTime, Tid, iprimaryPayRestAmount, inpointRestAmount, itotalRestAmount, pAddParam)

	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
elseif (IsChaiPay) then
	CALL PartialCancelChaiPay(orefund.FOneItem.FpaygateTid,fnGetChaiPayIdempotencyKey(ocsaslist.FOneItem.Forderserial), (mainpaymentorg-cardcancelsum), orefund.FOneItem.Frefundrequire,"",retval,ResultCode,ResultMsg,CancelDate,CancelTime, Tid)

	'// 리턴값중 CcPartCl 일단 무시(0:부분취소불가, 1:부분취소가능)
	CntRepay = "2"

	RepayPrice = iconfirm_price
	CancelPrice = orefund.FOneItem.Frefundrequire

	OldTid = ioldtid
else

	''--------------------------------------------------------------------------------------------------

		'*******************************************************************************
		'*
		'* 이미 정상 승인된 거래에서 취소를 원하는 금액을 입력하여 다시 승인을 득하도록 요청한다.
		'* 원거래 TID로 부분취소를 요청합니다.
		'* 취소 후 재승인이므로 새 거래 아이디가 생성됩니다.
		'* 원거래가 신용카드 지불인 경우에만 가능합니다
		'* OK 캐쉬백 적립 등이 포함되어 있는 경우 부분취소 불가능합니다
		'* 반드시 취소할 금액을 입력하도록 하세요
		'*
		'* Date : 2004/11
		'* Author : ts@inicis.com
		'* Project : INIpay V4.1 for ASP
		'*
		'* http://www.inicis.com
		'* Copyright (C) 2004 Inicis, Co. All rights reserved.
		'*******************************************************************************

		'###############################################################################
		'# 1. 객체 생성 #
		'################
		Set INIpay = Server.CreateObject("INItx41.INItx41.1")

		'###############################################################################
		'# 2. 인스턴스 초기화 #
		'###############################################################################
		PInst = INIpay.Initialize("")

		'###############################################################################
		'# 3. 거래 유형 설정 #
		'###############################################################################
		INIpay.SetActionType CLng(PInst), "REPAY"

		'###############################################################################
		'# 4. 지불 정보 설정 #
		'###############################################################################
		INIpay.SetField CLng(PInst), "pgid", "INIpayRPAY"  'PG ID (고정)
		INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '예비 PG IP (고정)
		INIpay.SetField CLng(PInst), "mid", MctID '상점아이디
		INIpay.SetField CLng(PInst), "admin", "1111" '키패스워드(상점아이디에 따라 변경)
		INIpay.SetField CLng(PInst), "oldTid", ioldtid	'취소할 원거래 아이디
		INIpay.SetField CLng(PInst), "currency", "WON" '화폐단위
		INIpay.SetField CLng(PInst), "price", iCancelPrice '가격
		INIpay.SetField CLng(PInst), "confirm_price", iconfirm_price '재승인 요청 금액 [이전승인금액 - 취소할 금액]
		INIpay.SetField CLng(PInst), "buyeremail", buyemail '이메일
		INIpay.SetField CLng(PInst), "url", "http://www.10x10.co.kr" '상점 홈페이지 주소 (URL)
		INIpay.SetField CLng(PInst), "debug", "true" '로그모드("true"로 설정하면 상세한 로그를 남김)


		'###############################################################################
		'# 5. 재승인 요청 #
		'###############################################################################
		INIpay.StartAction(CLng(PInst))

		'###############################################################################
		'#6. 부분 취소 결과 #
		'###############################################################################
		Tid         = INIpay.GetResult(CLng(PInst), "tid") '거래번호
		ResultCode  = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 지불성공)
		ResultMsg   = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
		OldTid      = INIpay.GetResult(CLng(PInst), "tid_org") '원거래번호
		CancelPrice = INIpay.GetResult(CLng(PInst), "price") '재결제 되는 금액
		RepayPrice  = INIpay.GetResult(CLng(PInst), "pr_remains") '취소 되는 금액
		CntRepay    = INIpay.GetResult(CLng(PInst), "cnt_partcancel") '부분취소(재승인) 요청횟수

		'###############################################################################
		'# 7. 인스턴스 해제 # 추가.
		'####################
		INIpay.Destroy CLng(PInst)

	'''--------------------------------------------------------------------------------------------------
end if

if (CntRepay="") then CntRepay=1 ''2013/06/27 추가 :: ''키워드 'where' 근처의 구문이 잘못되었습니다.

'// 승인취소되었으나, CS내역완료로 전환되지 않았을 경우, skyer9, 2017-04-11
If (id = 3402462) Then
	Tid = "INIpayRPAYteenxteen920170406100557769737"
	ResultCode = "00"
	ResultMsg = "[Card|정상처리되었습니다.]"
End If

sqlStr = "update [db_order].[dbo].tbl_card_cancel_log"&VbCRLF
sqlStr = sqlStr & " set newtid='"&Tid&"'"&VbCRLF
sqlStr = sqlStr & " ,resultcode='"&ResultCode&"'"&VbCRLF
sqlStr = sqlStr & " ,resultmsg='"&Replace(ResultMsg,"'","")&"'"&VbCRLF
sqlStr = sqlStr & " ,cancelrequestcount="&CntRepay&VbCRLF
sqlStr = sqlStr & " where clogIdx="&clogIdx&""

dbget.Execute sqlStr

''부분취소 정상처리시.tbl_order_paymentETc 업데이트.
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
		'// 아래 쿼리로 변경, skyer9, 2018-04-09
        sqlStr = " update db_order.dbo.tbl_order_paymentETc"&VbCRLF
		sqlStr = sqlStr & " set realPayedSum=realPayedSum-"&CancelPrice&VbCRLF
		sqlStr = sqlStr & " where orderserial='"&orgOrderserial&"'"&VbCRLF
		sqlStr = sqlStr & " and acctdiv='"&accountdiv&"'"&VbCRLF
		dbget.Execute sqlStr

		'// 아래 쿼리로 변경, skyer9, 2017-07-11
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

contents_finish = "결과 " & "[" & ResultCode & "]" & ResultMsg & VbCrlf
if (ResultCode="00") or (ResultCode="0000") or (IsKakaoPay and (resultCode="2001")) then
    IF (IsNaverPay) then
        contents_finish = contents_finish & "취소금액 : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "남은금액 : " &FormatNumber(itotalRestAmount,0) & VbCrlf
        ''contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "취소자 ID " & finishuserid
    else
        contents_finish = contents_finish & "취소금액 : " &FormatNumber(CancelPrice,0) & VbCrlf
        contents_finish = contents_finish & "남은금액 : " &FormatNumber(RepayPrice,0) & VbCrlf
        ''contents_finish = contents_finish & "취소일시 : " & CancelDate & " " & CancelTime & VbCrlf
        contents_finish = contents_finish & "취소자 ID " & finishuserid
    end if
end if

if date() >= "2023-03-27" then
' [네이버페이 요청]네이버페이 카드사 기프트카드 부분 취소 실패건 처리 방식 변경.    ' 2023.03.20 한용민
' 카드사 기프트카드의 경우 카드사 정책상 결제 부분 취소가 지원되지 않아 해당 건에 대해 부분 취소 요청 시 아래 응답이 전달되고 있습니다.
' code : CancelNotComplete / message : "3013/기프트카드는 부분취소가 불가능합니다."     '실제받은로그(결제인증 오류. 3013/기프트카드는 부분취소가 불가능합니다.)
' 응답이 전달되는 기프트카드 취소 실패건에 대해선 가맹점 요청 없이도 네이버페이에서 환불 처리할 예정이므로 가맹점에서 자체 환불할 시 이중 환불이 될 수 있습니다.
' 카드사 기프트카드 취소 실패건에 대해선 가맹점 자체 환불을 진행하지 않도록 운영 부탁드립니다.
if IsNaverPay then
    if ResultCode="ERR" then
        if instr(ResultMsg,"3013/기프트카드") > 0 then
            ' 네이버페이측에서 고객에게 직접 취소/환불 한다고 하니, cs건 완료처리만 한다.
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

    Call AddCustomerOpenContents(id, "부분 취소 완료: " & FormatNumber(CancelPrice,0) & VbCRLF & "남은 승인 금액: "& FormatNumber(RepayPrice,0)) '''CStr(refundrequire))


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
    sqlStr = sqlStr & " where a.orderserial = '" & iorderserial & "'"		' 주문번호
    sqlStr = sqlStr & " and a.id="& id &""	' 주문취소

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
        itemName = itemName & " 외 " & (itemCnt - 1) & "종"
    end if

    ''승인 취소 요청 SMS 발송
    if (iorderserial<>"") and (ibuyhp<>"") then
    	'sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"+ibuyhp+"','1644-6030','[텐바이텐]승인 부분 취소 되었습니다. 승인취소액 : " + FormatNumber(CancelPrice,0) + " 주문번호 : " + iorderserial + "'"
        'dbget.Execute sqlStr

        ' 부분취소. 카카오톡 알림톡 발송.   ' 2021.09.29 한용민 생성
        fullText = "[10x10] 취소접수안내" & vbCrLf & vbCrLf
        fullText = fullText & "고객님, 주문취소가 완료되었습니다." & vbCrLf & vbCrLf
        fullText = fullText & "■ 주문번호 : "& iorderserial &"" & vbCrLf
        fullText = fullText & "■ 상품명 : "& itemName &"" & vbCrLf
        fullText = fullText & "■ 취소금액 : "& FormatNumber(CancelPrice,0) &"원"
        failText = "[텐바이텐]주문취소가 완료되었습니다.주문번호 : "& iorderserial &""
        Call SendKakaoCSMsg_LINK("", ibuyhp,"1644-6030","KC-0021",fullText,"SMS","",failText,"",iorderserial,"")
    end if

    ''메일
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

		response.write "<br><br>* <font color=red>반복적으로 부분취소실패 오류</font>가 발생하는 경우 시스템팀 문의요망<br>(중복 취소일 수 있습니다.)"
        response.write "<br />* <font color=red>TOSS페이</font> 인 경우 회원탈퇴일 수 있습니다."

    end if
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
