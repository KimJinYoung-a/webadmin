<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/CSFunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%
if (C_InspectorUser = True) then
	response.write "<br><br>접근이 제한되었습니다.(접속 로그는 저장됩니다.)"
	dbget.close()
	response.end
end if

'[코드정리]
'------------------------------------------------------------------------------
'A008			주문취소
'
'A004			반품접수(업체배송)
'A010			회수신청(텐바이텐배송)
'
'A001			누락재발송
'A002			서비스발송
'
'A200			기타회수
'
'A000			맞교환출고
'A100			상품변경 맞교환출고
'
'A009			기타사항
'A006			출고시유의사항
'A700			업체기타정산
'
'A003			환불
'A005			외부몰환불요청
'A007			카드,이체,휴대폰취소요청
'
'A011			맞교환회수(텐바이텐배송)
'A012			맞교환반품(업체배송)

'A111			상품변경 맞교환회수(텐바이텐배송)
'A112			상품변경 맞교환반품(업체배송)

'[변수정리]
'------------------------------------------------------------------------------
'CSFunction.asp
'
'dim IsStatusRegister			'접수
'dim IsStatusEdit				'수정
'dim IsStatusFinishing			'처리완료 시도
'dim IsStatusFinished			'처리완료

'dim IsDisplayPreviousCSList	'이전 CS 내역
'dim IsDisplayCSMaster			'CS 마스터정보
'dim IsDisplayItemList			'상품목록
'dim IsDisplayChangeItemList	'다른상품 맞교환출고 상품목록
'dim IsDisplayRefundInfo		'환불정보
'dim IsDisplayButton			'버튼
'
'dim IsPossibleModifyCSMaster
'dim IsPossibleModifyItemList
'dim IsPossibleModifyRefundInfo

dim i, id, mode, divcd, orderserial, ckAll, sqlStr
dim IsOrderCanceled, OrderMasterState, IsTicketOrder, IsTravelOrder, IsChangeOrder, SelectedChangeOrderBrandId
dim IsMinusOrder, IsGiftingOrder, IsGiftiConOrder, IsOrderCancelDisabled, OrderCancelDisableStr, IsOutMallOrder, iPgGubun, iAccountDiv
	id			= request("id")
	divcd		= request("divcd")
	orderserial	= request("orderserial")
	mode		= request("mode")
	ckAll		= request("ckAll")

'CS접수마스터 가져오기
dim ocsaslist
set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id

	if (id<>"") then
	    ocsaslist.GetOneCSASMaster_3PL
	end if

'CS접수마스터 정보가 없을경우 신규 접수
if (ocsaslist.FResultCount<1) then
	set ocsaslist.FOneItem = new CCSASMasterItem
	ocsaslist.FOneItem.FId = 0
	ocsaslist.FOneItem.Fdivcd = divcd

	mode = "regcsas"
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    orderserial = ocsaslist.FOneItem.Forderserial

    if (ocsaslist.FOneItem.FCurrState = "B007") then
		mode = "finished"
    else
    	if (mode = "finishreginfo") then
    		'
    	else
    		mode = "editreginfo"
    	end if
    end if
end if

Call SetCSVariable(mode, divcd)

''환불정보
dim orefund
set orefund = New CCSASList
	orefund.FRectCsAsID = ocsaslist.FOneItem.FId
	orefund.GetOneRefundInfo

if (orefund.FOneItem.Fencmethod = "TBT") then
	orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
elseif (orefund.FOneItem.Fencmethod = "PH1") then
	orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
elseif (orefund.FOneItem.Fencmethod = "AE2") then
	orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
end if

function Decrypt(encstr)
	if (Not IsNull(encstr)) and (encstr <> "") then
		Decrypt = TBTDecrypt(encstr)
		exit function
	end if
	Decrypt = ""
end function

if (ocsaslist.FOneItem.FId <> 0) and ((ocsaslist.FOneITem.FDeleteyn = "Y") or (mode = "finished")) then
	if DateDiff("m", ocsaslist.FOneItem.Fregdate, Now) > 3 then
		orefund.FOneItem.Frebankaccount = ""
		orefund.FOneItem.Frebankownername = ""
		orefund.FOneItem.Frebankname = ""
	end if
end if

''주문 마스타
dim oordermaster, IsCalculateAddBeasongPayNeed
set oordermaster = new COrderMaster
	oordermaster.FRectOrderSerial = orderserial

	if Left(orderserial,1)="A" then
	    set oordermaster.FOneItem = new COrderMasterItem
	else
	    oordermaster.QuickSearchOrderMaster_3PL
	end if

'' 과거 6개월 이전 내역 검색
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if

IsOrderCanceled  = (oordermaster.FOneItem.Fcancelyn = "Y")
OrderMasterState = oordermaster.FOneItem.FIpkumDiv
IsTicketOrder    = (oordermaster.FOneItem.FjumunDiv="4")
IsTravelOrder    = (oordermaster.FOneItem.FjumunDiv="3")
IsOutMallOrder   = (oordermaster.FOneItem.FjumunDiv="5")
IsChangeOrder    = (oordermaster.FOneItem.FjumunDiv="6")
IsMinusOrder	 = (oordermaster.FOneItem.FjumunDiv="9")
IsGiftingOrder   = (oordermaster.FOneItem.Faccountdiv="550") or (oordermaster.FOneItem.FSiteName = "giftting")
IsGiftiConOrder  = (oordermaster.FOneItem.Faccountdiv="560")

iPgGubun = (oordermaster.FOneItem.Fpggubun) ''2016/07/21
iAccountDiv = (oordermaster.FOneItem.FAccountDiv) ''2016/08/05

if (IsStatusRegister and IsChangeOrder) then
	SelectedChangeOrderBrandId = GetChangeOrderBrandInfo(orderserial)
end if

'보너스쿠폰
dim IsBonusCouponExist, IsBonusCouponAvailable, ocscoupon

set ocscoupon = New CCSCenterCoupon
	IsBonusCouponExist = (oordermaster.FOneItem.FUserID <> "") and (oordermaster.FOneItem.Ftencardspend > 0) and (Not IsNull(oordermaster.FOneItem.FbCpnIdx))

	if IsBonusCouponExist then
		ocscoupon.FRectBonusCouponIdx = oordermaster.FOneItem.FbCpnIdx
		ocscoupon.GetOneCSCenterCoupon
	end if

dim curr_subtotalprice, curr_sumPaymentEtc, curr_tencardspend, curr_miletotalprice, curr_spendmembership, curr_allatdiscountprice
dim curr_depositsum, curr_giftcardsum, curr_percentcouponsum, curr_itemcostsum, curr_beasongpaysum
dim totalpreminus_itemcostsum, totalpreminus_beasongpaysum, ckbeasongpayAssignChecked
dim tmpResultArr, arrOrderPriceInfoMakerid, arrOrderPriceInfoItemCost, arrOrderPriceInfoDeliverCost, arrIsDeliverCostSelected
curr_subtotalprice = 0
curr_sumPaymentEtc = 0
curr_tencardspend = 0
curr_miletotalprice = 0
curr_spendmembership = 0
curr_allatdiscountprice = 0
curr_depositsum = 0
curr_giftcardsum = 0
curr_percentcouponsum = 0
curr_itemcostsum = 0
curr_beasongpaysum = 0
totalpreminus_itemcostsum = 0
totalpreminus_beasongpaysum = 0

if (IsDisplayRefundInfo) and (IsCSCancelInfoNeeded(divcd)) then
	'// =======================================================================
	'// 접수시 : 현재 결제정보(상품변경 맞교환 포함)
	'// 접수이후 : 저장된 결제정보
	'// =======================================================================

	''원주문건( + 교환주문) 금액 조회(취소 상품 제외)
	Call GetOrgOrderPriceInfo(orderserial, curr_subtotalprice, curr_sumPaymentEtc, curr_tencardspend, curr_miletotalprice, curr_spendmembership, curr_allatdiscountprice, curr_depositsum, curr_giftcardsum, curr_percentcouponsum)

	''TODO : 정액쿠폰 분할 환불처리
	''--update
	''dbo.tbl_as_refund_info
	''set orgcouponsum = 2400
	''where asid = 1323982

	if (IsStatusRegister) then

		curr_itemcostsum = (curr_subtotalprice + curr_tencardspend + curr_miletotalprice + curr_spendmembership + curr_allatdiscountprice)

		curr_beasongpaysum = 0

		'// 브랜드별 상품금액, 배송비
		sqlStr = " exec db_order.dbo.usp_Ten_GetOrderPriceInfoByBrand '" & orderserial & "', " + CStr(ocsaslist.FOneItem.FId) + " "

	    rsget.Open sqlStr,dbget,1
	    if Not rsget.Eof then

	    	tmpResultArr = rsget.GetRows()

			redim arrOrderPriceInfoMakerid(UBound(tmpResultArr, 2) + 1)
			redim arrOrderPriceInfoItemCost(UBound(tmpResultArr, 2) + 1)
			redim arrOrderPriceInfoDeliverCost(UBound(tmpResultArr, 2) + 1)
			redim arrIsDeliverCostSelected(UBound(tmpResultArr, 2) + 1)

			i = 0
			for i = 0 to UBound(tmpResultArr, 2)
				arrOrderPriceInfoMakerid(i) 			= tmpResultArr(0, i)
				arrOrderPriceInfoItemCost(i) 			= tmpResultArr(1, i)
				arrOrderPriceInfoDeliverCost(i) 		= tmpResultArr(2, i)
				arrIsDeliverCostSelected(i) 			= tmpResultArr(5, i)

				curr_beasongpaysum = curr_beasongpaysum + arrOrderPriceInfoDeliverCost(i)
			next

			curr_itemcostsum = curr_itemcostsum - curr_beasongpaysum

		end if
		rsget.Close


		'접수시 초기값 세팅
		if IsChangeOrder then
			'교환주문은 브랜드금액만
			for i = 0 to UBound(arrOrderPriceInfoMakerid) - 1
				if (SelectedChangeOrderBrandId = arrOrderPriceInfoMakerid(i)) then
					curr_itemcostsum = arrOrderPriceInfoItemCost(i) + curr_tencardspend + curr_miletotalprice + curr_spendmembership + curr_allatdiscountprice
					curr_beasongpaysum = arrOrderPriceInfoDeliverCost(i)
				end if
			next
		end if
		orefund.FOneItem.Forgitemcostsum 	= curr_itemcostsum
		orefund.FOneItem.Forgbeasongpay 	= curr_beasongpaysum

		orefund.FOneItem.Forgmileagesum 	= curr_miletotalprice
		orefund.FOneItem.Forgcouponsum 		= curr_tencardspend
		orefund.FOneItem.Fallatsubtractsum 	= curr_allatdiscountprice
		orefund.FOneItem.Forgdepositsum		= curr_depositsum
		orefund.FOneItem.Forggiftcardsum	= curr_giftcardsum
		orefund.FOneItem.Forgallatdiscountsum = curr_allatdiscountprice

		if (curr_tencardspend <> 0) then
			if (curr_percentcouponsum <> 0) then
				orefund.FOneItem.Forgpercentcouponsum 	= curr_tencardspend
				orefund.FOneItem.Forgfixedcouponsum		= 0
			else
				orefund.FOneItem.Forgpercentcouponsum 	= 0
				orefund.FOneItem.Forgfixedcouponsum		= curr_tencardspend
			end if
		else
			orefund.FOneItem.Forgpercentcouponsum	= 0
			orefund.FOneItem.Forgfixedcouponsum		= 0
		end if

		''orefund.FOneItem.Forgpercentcouponsum	= curr_percentcouponsum
		''orefund.FOneItem.Forgfixedcouponsum		= curr_tencardspend - curr_percentcouponsum

		orefund.FoneItem.Frefundadjustpay 	= 0
		orefund.FOneItem.Frefunddeliverypay = 0

		orefund.FOneItem.Frefundcouponsum	= 0
		orefund.FOneItem.Frefundmileagesum	= 0

        orefund.FOneItem.Frefundgiftcardsum = 0
        orefund.FOneItem.Frefunddepositsum  = 0

	end if

end if

''기환불정보
dim prevrefund, prevrefundsum, csbeasongpaysum
set prevrefund = New CCSASList
	prevrefund.FRectOrderSerial = orderserial
	prevrefundsum = prevrefund.GetPrevRefundSum

'배송비 취소 없이 배송비환불이 이루어진 금액
csbeasongpaysum = prevrefund.GetPrevRefundCSDeliveryPaySum

''기존 무통장 환불정보
dim orefundInfo, prevrefundhistorycnt

set orefundInfo = New CCSASList
orefundInfo.FCurrpage = 1
orefundInfo.FPageSize = 10
orefundInfo.FRectUserID = oordermaster.FOneItem.FUserID

if (oordermaster.FOneItem.FUserID="") then
	prevrefundhistorycnt = "없음"
else
    orefundInfo.GetHisOldRefundInfo

    prevrefundhistorycnt = orefundInfo.FTotalCount
end if

'==============================================================================
'접수시 디폴트값 설정
'==============================================================================
if (IsStatusRegister) then

	'// 구매자명=무통장환불 계좌주
	orefund.FOneItem.Frebankownername = oordermaster.FOneItem.FBuyname

	'// 기본문구 설정
	if InStr("A004,A001,A002,A200,A009,A006,A700,A000", divcd) then
		if Not IsNull(session("ssBctCname")) then
			ocsaslist.FOneItem.Fcontents_jupsu = "텐바이텐 고객센터 " + CStr(session("ssBctCname")) + " 입니다"
		end if
	end if
end if

'==============================================================================
'상품취소로 배송비 재계산을 할지(외부몰, 해외배송, 군부대배송 XX)
IsCalculateAddBeasongPayNeed = (oordermaster.FOneItem.Fjumundiv <> "5") and (oordermaster.FOneItem.FDlvcountryCode = "KR")

dim oupchebeasongpay
set oupchebeasongpay = new COrderMaster
	if (orderserial<>"") then
		oupchebeasongpay.FRectOrderSerial = orderserial
		oupchebeasongpay.getUpcheBeasongPayList
	end if

'==============================================================================
'주문 디테일
dim ocsOrderDetail
set ocsOrderDetail = new CCSASList
	ocsOrderDetail.FRectCsAsID = ocsaslist.FOneItem.FId
	ocsOrderDetail.FRectOrderSerial = orderserial

	if (oordermaster.FRectOldOrder = "on") then
	    ocsOrderDetail.FRectOldOrder = "on"
	end if

	ocsOrderDetail.GetOrderDetailByCsDetailNew_3PL

'==============================================================================
'상품변경 맞교환
dim ocsChangeOrderDetail
set ocsChangeOrderDetail = new CCSASList
	ocsChangeOrderDetail.FRectCsAsID = ocsaslist.FOneItem.FId
	ocsChangeOrderDetail.FRectOrderSerial = orderserial

	if (oordermaster.FRectOldOrder = "on") then
	    ocsChangeOrderDetail.FRectOldOrder = "on"
	end if

	if (IsDisplayChangeItemList) then
		ocsChangeOrderDetail.GetChangeOrderDetailByCsDetailNew
	end if

'==============================================================================
'// 맞교환회수
dim ioneRefas, IsRefASExist, IsRefASFinished

set ioneRefas = new CCSASList
	IsRefASExist = False
	IsRefASFinished = False

	if (Not IsStatusRegister) then
		if (divcd = "A000") or (divcd = "A100") then
			set ioneRefas = new CCSASList
			ioneRefas.FRectCsRefAsID = id
			ioneRefas.GetOneCSASMaster

			if (ioneRefas.FResultCount>0) then
				IsRefASExist = True
			    if (ioneRefas.FOneItem.Fcurrstate = "B007") then
			    	IsRefASFinished = True
			    end if
			end if
		end if
	end if

'==============================================================================
'보조결제수단
dim oetcpayment, realdepositsum, realgiftcardsum, realSubPaymentSum, orgSubPaymentSum

set oetcpayment = new COrderMaster
realdepositsum = 0
realgiftcardsum = 0
realSubPaymentSum = 0
if (orderserial<>"") then
	oetcpayment.FRectOrderSerial = orderserial
	oetcpayment.getEtcPaymentList

	'200 : 예치금
	for i = 0 to oetcpayment.FResultCount - 1
		if (CStr(oetcpayment.FItemList(i).Facctdiv) = "200") then
			realdepositsum = oetcpayment.FItemList(i).FrealPayedsum
		end if
	next

	'900 : Gift카드
	for i = 0 to oetcpayment.FResultCount - 1
		if (CStr(oetcpayment.FItemList(i).Facctdiv) = "900") then
			realgiftcardsum = oetcpayment.FItemList(i).FrealPayedsum
		end if
	next
end if
realSubPaymentSum = realdepositsum+realgiftcardsum
''orgSubPaymentSum = orgdepositsum+orggiftcardsum

'==============================================================================
'최초 주결제수단금액
dim omainpayment, mainpaymentorg, phonePartialCancelok, isThisdateReturn
dim cardPartialCancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcode, cardcodeall, installment, isThisdateCancel

set omainpayment = new COrderMaster
mainpaymentorg = 0
cardPartialCancelok = "N"
phonePartialCancelok = "N"
cardcancelerrormsg = ""
cardcancelcount = 0
cardcancelsum   = 0
cardcodeall = ""
cardcode    = ""
installment = 0
isThisdateCancel = (Left(CStr(oordermaster.FoneItem.FRegdate),10)=Left(now(),10))

isThisdateReturn = False
if Not IsNull(oordermaster.FoneItem.Fbeadaldate) then
	isThisdateReturn = (Left(CStr(oordermaster.FoneItem.Fbeadaldate),10)=Left(now(),10))
end if

if (orderserial<>"") then
	omainpayment.FRectOrderSerial = orderserial

	Call omainpayment.getMainPaymentInfo(oordermaster.FOneItem.Faccountdiv, mainpaymentorg, cardPartialCancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcodeall)

    ''할불개월수
    ''installment = Right(cardcodeall,2) 14|26|00 ==> 14|26|00|1 ''마지막 코드 부분취소 가능여부 (2011-08-25)--------
    IF Not IsNULL(cardcodeall) THEN
        cardcodeall= TRIM(cardcodeall)
        cardcodeall = LEft(cardcodeall,10)   '''모바일쪽 코드 이상함 (빈값 또는 이상한 값)
    END IF

    if (LEN(TRIM(cardcodeall))=10) then
        if (Right(cardcodeall,1)="1") then
            cardPartialCancelok = "Y"
        elseif (Right(cardcodeall,1)="0") then
            cardPartialCancelok = "N"
            if (cardcancelerrormsg="") then cardcancelerrormsg  = "부분취소 <strong>불가</strong> 거래 (충전식 카드 or 복합거래)"
        end if

        installment = Mid(cardcodeall,7,2)
    else
        installment = Right(TRIM(cardcodeall),2)
		installment = Replace(installment, "|", "")
    end if
    ''----------------------------------------------------------------------------------------------------------------

    cardcode    = Left(cardcodeall,2)
    if IsNumeric(installment) then installment=CLNG(installment)
	if (TRIM(installment)="") then installment=0


	if (oordermaster.FOneItem.Faccountdiv = "400") then
		if (Left(now(), 7) = Left(oordermaster.FOneItem.Fipkumdate, 7)) then
			phonePartialCancelok = "Y"
		end if
	end if

	if (orderserial = "14123062296") then
		installment = 0
	end if
end if

'==============================================================================
dim RefundAllowLimit
	RefundAllowLimit = GetUserRefundAuthLimit(session("ssBctId"))

'==============================================================================
'원주문 상품금액
dim orgitemcostsum, orgpercentcouponpricesum

'접수상품 합계금액(inc_cs_action_item_list.asp 에서 계산된다)
dim regitemcostsum, regpercentcouponpricesum

'디테일 id(orderdetailidx)
dim distinctid

'==============================================================================
''접수 불가시 메세지
dim JupsuInValidMsg

if (Left(orderserial,1)<>"A") and (oordermaster.FResultCount<1) then
    response.write "<br><br>!!! 과거 주문내역이거나 주문 내역이 없습니다. - 관리자 문의 요망"
    dbget.close()	:	response.End
end if

''접수 가능 여부
dim IsJupsuProcessAvail
if (oordermaster.FResultCount>0) then
	if IsChangeOrder then
		IsJupsuProcessAvail = ocsaslist.FOneItem.IsChangeAsRegAvail(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
	elseif IsMinusOrder then
		IsJupsuProcessAvail = false
		JupsuInValidMsg = "마이너스주문에 대해 CS접수할 수 없습니다."
	else
		IsJupsuProcessAvail = ocsaslist.FOneItem.IsAsRegAvail(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
	end if
else
    IsJupsuProcessAvail = false
end if

dim IsNotFinisherCancelCSExist
if IsCSCancelProcess(divcd) and IsStatusRegister and (IsJupsuProcessAvail = true) then
	IsNotFinisherCancelCSExist = CheckNotFinishedCancelCSExist(orderserial)

	if IsNotFinisherCancelCSExist then
		IsJupsuProcessAvail = false
		JupsuInValidMsg = "완료되지 않은 주문취소 접수건이 있습니다.\n먼저 접수건을 완료처리 하세요."
	end if
end if

'업체처리완료상태 여부
dim IsUpcheConfirmState
IsUpcheConfirmState = (ocsaslist.FOneItem.FCurrState="B006")

'// 택배사 전송 이전상태인지
dim IsLogicsSended
IsLogicsSended = (ocsaslist.FOneItem.FCurrState<>"B001")

''티켓Order인경우 취소 검증.
Dim mayTicketCancelChargePro : mayTicketCancelChargePro=0
Dim ticketCancelStr : ticketCancelStr =""
Dim ticketCancelDisabled : ticketCancelDisabled = false   '''티켓 취소 불가한지.

if (IsTicketOrder) and (IsStatusRegister) and (oordermaster.FOneItem.IsPayedOrder) then
    '' 당일주문건은 취소수수료가 없슴.
    if (Not isThisdateCancel) then
        call TicketOrderCheck(orderserial,mayTicketCancelChargePro,ticketCancelDisabled, ticketCancelStr)
		if (session("ssBctId") = "nownhere21") and ticketCancelDisabled = True then
			'2018-02-21, skyer9
			ticketCancelDisabled = False
		end if
    end if
end if

''취소가능한 주문인지
IsOrderCancelDisabled = False
OrderCancelDisableStr = ""
if (IsGiftingOrder or IsGiftiConOrder) and (IsCSCancelProcess(divcd) or IsCSReturnProcess(divcd)) then
	IsOrderCancelDisabled = True
	OrderCancelDisableStr = "기프팅/기프티콘 결제 주문입니다.\n\n반품접수 할 수 없습니다.[주문취소시 예치금환불만 가능]"
end if

'// 여행주문
dim travelItemInfoArr, travelItemExist
travelItemExist = False
if (IsTravelOrder) and (IsStatusRegister or IsStatusEdit) and (oordermaster.FOneItem.IsPayedOrder) and IsCSReturnProcess(divcd) then
	travelItemInfoArr = TravelOrderCheckArr(orderserial)
	if IsArray(travelItemInfoArr) then
		travelItemExist = True
	end if
end if

'==============================================================================
''완료처리 불가시 메세지
dim FinishInValidMsg

''완료처리 가능 여부
dim IsFinishProcessAvail

FinishInValidMsg = ""
IsFinishProcessAvail = False

if (IsStatusFinishing) then
	IsFinishProcessAvail = True

	if (IsRefASExist) and (IsRefASFinished = False) and (ocsaslist.FOneItem.Frequireupche = "Y") then
    	FinishInValidMsg = "업체배송의 경우 맞교환회수를 먼저 완료처리해야 맞교환출고를 완료처리할 수 있습니다."
    	IsFinishProcessAvail = False
	end if
end if

'==============================================================================
dim IsDelFinishedCSAvail : IsDelFinishedCSAvail = False
dim DelFinishedCSInValidMsg : DelFinishedCSInValidMsg = "<font color='red'>완료내역 삭제불가</font>"
dim oRefCSASList

dim HasAuthTodayDelCancelReturn : HasAuthTodayDelCancelReturn = False
dim HasAuthUpcheJungsanItemPrice : HasAuthUpcheJungsanItemPrice = False
' 사용금지
'HasAuthUpcheJungsanItemPrice = (InStr(",ilovecozie,durida22,boyishP,hasora,bseo,skyer9,coolhas,oesesang52,hrkang97,tozzinet,gomgom,zerogirl0730,heendoongi,happy799,may0816,angela919,hk9566371,seokmi1221,pray16,jjy158,rabbit1693,", ("," & session("ssBctId") & ",")) > 0)

if IsStatusFinished and (ocsaslist.FOneITem.FDeleteyn="N") then
	if (divcd="A004") or (divcd="A010") or (divcd="A008") then
		'// 반품취소 완료CS 삭제

		set oRefCSASList = new CCSASList
		oRefCSASList.FRectCsRefAsID = id
		oRefCSASList.GetOneCSASMaster

		if (oRefCSASList.FResultCount > 0) then
			if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
				if (oRefCSASList.FOneItem.Fcurrstate = "B007") then
					DelFinishedCSInValidMsg = "<font color='red'>시스템팀 문의(환불완료 상태입니다.)</font>"
				else
					DelFinishedCSInValidMsg = "먼저 관련 환불CS를 삭제하세요."
				end if
			else
				IsDelFinishedCSAvail = True
				HasAuthTodayDelCancelReturn = (InStr(",ilovecozie,durida22,boyishP,hasora,bseo,skyer9,coolhas,oesesang52,hrkang97,tozzinet,gomgom,zerogirl0730,heendoongi,happy799,may0816,angela919,hk9566371,seokmi1221,pray16,jjy158,rabbit1693,", ("," & session("ssBctId") & ",")) > 0)
				if HasAuthTodayDelCancelReturn and DateDiff("d", ocsaslist.FOneItem.Ffinishdate, Now()) <> 0 then
					HasAuthTodayDelCancelReturn = False
				end if
			end if
		else
			if ((divcd="A004") or (divcd="A010")) and (IsNull(ocsaslist.FOneItem.Frefminusorderserial) or (ocsaslist.FOneItem.Frefminusorderserial = "")) then
				DelFinishedCSInValidMsg = "<font color='red'>시스템팀 문의(마이너스 주문번호 없음)</font>"
			elseif ((divcd="A008") and (oordermaster.FOneItem.Fipkumdiv >= "4")) then
				if (ocsaslist.FOneItem.Ffinishdate < oordermaster.FOneItem.Fipkumdate) then
					DelFinishedCSInValidMsg = "취소불가. <font color='red'>결제이전 취소</font> 내역입니다."
				else
					IsDelFinishedCSAvail = True
					HasAuthTodayDelCancelReturn = (InStr(",ilovecozie,durida22,boyishP,hasora,bseo,skyer9,coolhas,oesesang52,hrkang97,tozzinet,gomgom,zerogirl0730,heendoongi,happy799,may0816,angela919,hk9566371,seokmi1221,pray16,jjy158,rabbit1693,", ("," & session("ssBctId") & ",")) > 0)
					if HasAuthTodayDelCancelReturn and DateDiff("d", ocsaslist.FOneItem.Ffinishdate, Now()) <> 0 then
						HasAuthTodayDelCancelReturn = False
					end if
				end if
			else
				IsDelFinishedCSAvail = True
				HasAuthTodayDelCancelReturn = (InStr(",ilovecozie,durida22,boyishP,hasora,bseo,skyer9,coolhas,oesesang52,hrkang97,tozzinet,gomgom,zerogirl0730,heendoongi,happy799,may0816,angela919,hk9566371,seokmi1221,pray16,jjy158,rabbit1693,", ("," & session("ssBctId") & ",")) > 0)
				if HasAuthTodayDelCancelReturn and DateDiff("d", ocsaslist.FOneItem.Ffinishdate, Now()) <> 0 then
					HasAuthTodayDelCancelReturn = False
				end if
			end if
		end if
	end if
end if

''결제완료 없이 마일리지 적립 필요한 경우
dim exceptOrderserial : exceptOrderserial = "xxxxxxxxx"

'// 임시 이벤트
'// 브랜드 : laundrymat
'// 출고금액 : 50000
'// 주문당 : 1
'// 기간 : 2016.03.07~2016.03.29
'// 입점몰 제외
dim IsTempEventAvail : IsTempEventAvail = True
dim IsTempEventAvail_Str : IsTempEventAvail_Str = ""
dim IsTempEventAvail_Makerid

IF application("Svr_Info")="Dev" THEN
	IsTempEventAvail_Makerid = "noulnabi"
else
	IsTempEventAvail_Makerid = "laundrymat"
end if

''접수
IsTempEventAvail = IsTempEventAvail and IsStatusRegister and (Not IsOutMallOrder)

''반품
IsTempEventAvail = IsTempEventAvail and (divcd = "A004")

if IsTempEventAvail then
	IsTempEventAvail = False
	for i = 0 to ocsOrderDetail.FResultCount - 1
		if (ocsOrderDetail.FItemList(i).Fmakerid = IsTempEventAvail_Makerid) then
			IF application("Svr_Info")="Dev" THEN
				IsTempEventAvail_Str = CheckFreeReturnDeliveryAvail(orderserial, IsTempEventAvail_Makerid, "2016-03-03", "2016-03-29", 3000, 1)
			else
				IsTempEventAvail_Str = CheckFreeReturnDeliveryAvail(orderserial, IsTempEventAvail_Makerid, "2016-03-07", "2016-03-29", 50000, 1)
			end if

			if (IsTempEventAvail_Str = "") then
				IsTempEventAvail = True
			end if

			exit for
		end if
	next
end if


dim oGift
dim IsDisplayGift : IsDisplayGift = False
set oGift = new COrderGift

if (oordermaster.FOneItem.Fipkumdiv>1) and (oordermaster.FOneItem.Fjumundiv<>9) and ((divcd = "A008") or (divcd = "A010") or (divcd = "A004")) then
    oGift.FRectOrderSerial = orderserial
    oGift.GetOneOrderGiftlist
	if (oGift.FResultCount > 0) then
		IsDisplayGift = True
	end if
end if

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/cscenter/js/date.format.js"></script>
<script language='javascript' SRC="/js/ajax.js"></script>
<script type="text/javascript">
var IsC_ADMIN_AUTH               = <%= LCase(C_ADMIN_AUTH) %>;
var IsCsPowerUser               = <%= LCase(C_CSPowerUser) %>;
var HasAuthUpcheJungsanItemPrice	= <%= LCase(HasAuthUpcheJungsanItemPrice) %>;	// 사용금지
var C_CSpermanentUser = <%= LCase(C_CSpermanentUser) %>;
var IsTempEventAvail            = <%= LCase(IsTempEventAvail) %>;
var IsTempEventAvail_Makerid    = "<%= LCase(IsTempEventAvail_Makerid) %>";

var OrderMasterState			= "<%= OrderMasterState %>";

var IsTicketOrder               = <%= LCase(IsTicketOrder) %>;
var IsTravelOrder               = <%= LCase(IsTravelOrder) %>;
var IsChangeOrder               = <%= LCase(IsChangeOrder) %>;

var IsGiftingOrder              = <%= LCase(IsGiftingOrder) %>;
var IsGiftiConOrder             = <%= LCase(IsGiftiConOrder) %>;

var IsLogicsSended             	= <%= LCase(IsLogicsSended) %>;

var travelItemInfoArr			= new Array();
var travelItemExist				= <%= LCase(travelItemExist) %>;
<% if travelItemExist then
	for i = 0 to UBound(travelItemInfoArr,2)
		response.write "travelItemInfoArr.push(new Array(" & travelItemInfoArr(0,i) & ", '" & travelItemInfoArr(1,i) & "', " & travelItemInfoArr(2,i) & ", '" & travelItemInfoArr(3,i) & "'));"
		if i < UBound(travelItemInfoArr,2) then
			response.write vbCrLf
		end if
	next
end if %>

var ticketCancelDisabled        = <%= LCase(ticketCancelDisabled) %>;
var travelCancelDisabled		= false;
var IsOrderCancelDisabled       = <%= LCase(IsOrderCancelDisabled) %>;

var ticketCancelStr             = '<%= ticketCancelStr %>';
var travelCancelStr             = '';
var OrderCancelDisableStr       = '<%= OrderCancelDisableStr %>';

var mayTicketCancelChargePro    = <%= mayTicketCancelChargePro %>;
var RefundAllowLimit			= <%= RefundAllowLimit %>;

var IsStatusRegister 			= <%= LCase(IsStatusRegister) %>;
var IsStatusEdit 				= <%= LCase(IsStatusEdit) %>;
var IsStatusFinishing 			= <%= LCase(IsStatusFinishing) %>;
var IsStatusFinished 			= <%= LCase(IsStatusFinished) %>;

var IsDisplayPreviousCSList 	= <%= LCase(IsDisplayPreviousCSList) %>;
var IsDisplayCSMaster 			= <%= LCase(IsDisplayCSMaster) %>;
var IsDisplayItemList 			= <%= LCase(IsDisplayItemList) %>;
var IsDisplayRefundInfo 		= <%= LCase(IsDisplayRefundInfo) %>;
var IsDisplayButton 			= <%= LCase(IsDisplayButton) %>;

var IsCSCancelInfoNeeded		= <%= LCase(IsCSCancelInfoNeeded(divcd)) %>;
var IsCSRefundNeeded			= <%= LCase(IsCSRefundNeeded(divcd, OrderMasterState)) %>;

var IsPossibleModifyCSMaster	= <%= LCase(IsPossibleModifyCSMaster) %>;
var IsPossibleModifyItemList	= <%= LCase(IsPossibleModifyItemList) %>;
var IsPossibleModifyRefundInfo	= <%= LCase(IsPossibleModifyRefundInfo) %>;

var IsCSCancelProcess			= <%= LCase(IsCSCancelProcess(divcd)) %>;
var IsCSReturnProcess			= <%= LCase(IsCSReturnProcess(divcd)) %>;
var IsCSServiceProcess			= <%= LCase(IsCSServiceProcess(divcd)) %>;

var MainPaymentOrg				= <%= mainpaymentorg %>;
var precardcancelsum            = <%= cardcancelsum %>;
var installment                 = <%= installment %>;
var cardPartialCancelok			= "<%= cardPartialCancelok %>";
var cardcode					= "<%= cardcode %>";
var isThisdateCancel            = "<%= chkIIF(isThisdateCancel,"Y","N") %>";

var phonePartialCancelok		= "<%= phonePartialCancelok %>";

// 한개의 브랜드만 선택가능한가
// 반품접수(업배), 맞교환출고, 누락재발송, 서비스발송, 기타회수, 출고시유의사항, 회수신청(텐바이텐배송), 업체기타정산
var IsOnlyOneBrandAvailable		= <%= LCase(InStr("A004,A000,A001,A002,A200,A006,A010,A700", divcd) > 0) %>;

var IsDeletedCS 				= <%= LCase(ocsaslist.FOneITem.FDeleteyn = "Y") %>;

var ERROR_MSG_TRY_MODIFY		= "<%= ERROR_MSG_TRY_MODIFY %>";

var CDEFAULTBEASONGPAY 		= 2000; // 텐바이텐 기본 배송비
var divcd 					= "<%= divcd %>";
var mode 					= "<%= mode %>";
var orderserial 			= "<%= orderserial %>";
var sitename	 			= "<%= oordermaster.FOneItem.FSiteName %>";
var pggubun                 = "<%=iPgGubun%>";      //2016/07/21
var orgaccountdiv           = "<%=iAccountDiv%>";   //2016/08/05

var IsAdminLogin 			= IsCsPowerUser; ///<%= LCase((session("ssBctId") = "icommang") or (session("ssBctId") = "iroo4") or (session("ssBctId") = "bseo")) %>;
var IsOrderFound 			= <%= LCase(oordermaster.FResultCount > 0) %>;
var IsRefundInfoFound 		= <%= LCase(orefund.FResultCount > 0) %>;

<% if (oordermaster.FResultCount > 0) then %>
var IsThisMonthJumun 		= <%= LCase(datediff("m", oordermaster.FOneItem.FRegdate, now()) <= 0) %>;
<% else %>
var IsThisMonthJumun 		= false;
<% end if %>

var arrmakerid = new Array();
var arrdefaultfreebeasonglimit = new Array();
var arrdefaultdeliverpay = new Array();

<% for i = 0 to oupchebeasongpay.FResultCount - 1 %>
	arrmakerid[<%= i %>] = "<%= LCase(oupchebeasongpay.FItemList(i).Fmakerid) %>";
	arrdefaultfreebeasonglimit[<%= i %>] = <%= oupchebeasongpay.FItemList(i).Fdefaultfreebeasonglimit %>;
	arrdefaultdeliverpay[<%= i %>] = <%= oupchebeasongpay.FItemList(i).Fdefaultdeliverpay %>;
<% next %>

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function WriteNowDateString(v) {
	var d = new Date();
	v.focus();

	// /cscenter/js/date.format.js
	v.value = v.value + "\n\n+" + d.format("yyyy-mm-dd HH:MM:ss") + "  고객센터 <%= session("ssBctCname") %>입니다.\n";
}

function TnCSTemplateGubunChanged(gubun) {

	CSTemplateFrame.location.href="/cscenter/board/cs_template_select_process.asp?mastergubun=30&gubun=" + gubun;
}

function TnCSTemplateGubunProcess(v, errMSG) {

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if(v == "") {
		//
	} else {
		document.frmaction.contents_jupsu.value = v;
		// alert(v);
	}
}

function popChkGiftItem() {
	var frm = document.frmaction;
	var IsCheckNeed = "<%= CHKIIF(divcd="A008" and oGift.FResultCount>0 and IsStatusRegister, "Y", "N") %>";
	if (IsCheckNeed == "Y") {
		document.getElementById("evt_chk_need").value = "Y";
	}
	evt_chk_need = document.getElementById("evt_chk_need");
	if (evt_chk_need.value == "N") {
		alert("체크가 필요없습니다.");
		return;
	}

	if (IsAllSelected(frm) == true) {
		alert("전체취소입니다.체크가 필요없습니다.");
		evt_chk_need.value = "N";
		return;
	}

	/*
	if (frm.gubun01.value == "") {
		alert("먼저 사유구분을 입력하세요.");
		return;
	}

	if ((frm.gubun01.value != "C004") || (frm.gubun02.value != "CD01")) {
		alert("변심취소 이외에는 체크가 필요없습니다.");
		evt_chk_need.value = "N";
		return;
	}
	*/

	var orderdetailidx, itemid, regitemno;
	var itemlist = "";
	for (var i = 0; ; i++) {
		orderdetailidx = document.getElementById("orderdetailidx_" + i);
		itemid = document.getElementById("itemid_" + i);
		regitemno = document.getElementById("regitemno_" + i);

		if (orderdetailidx == undefined) { break; }
		if (orderdetailidx.checked == false) { continue; }
		if (parseInt(itemid.value,10) == 0) { continue; }

		itemlist = itemlist + "|" + orderdetailidx.value + "," + regitemno.value
	}

	/*
	if (itemlist == "") {
		alert("선택된 상품이 없습니다.");
		return;
	}
	*/

	var popwin = window.open('pop_cs_gift_modify.asp?orderserial=' + frm.orderserial.value + '&mode=chk&itemlist=' + itemlist,'pop_cs_gift_modify','width=1200,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
<script language="javascript" SRC="/admin/etc/3pl/js/newcsas_3PL.js?v=1"></script>

<form name="popForm" action="/cscenter/ordermaster/popDeliveryTrace.asp" target="_blank">
<input type="hidden" name="traceUrl">
<input type="hidden" name="songjangNo">
</form>

<form name="frmaction" method="post" action="pop_cs_action_new_process_3PL.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="modeflag2" value="<%= mode %>">
<input type="hidden" name="orderserial" value="<%= orderserial %>" >
<input type="hidden" name="id" value="<%= ocsaslist.FOneItem.Fid %>">
<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
<input type="hidden" name="detailitemlist" value="">
<input type="hidden" name="csdetailitemlist" value="">
<input type="hidden" name="ipkumdiv" value="<%= oordermaster.FOneItem.Fipkumdiv %>">
<input type="hidden" name="copycouponinfo" value="<%= orefund.FOneItem.Fcopycouponinfo %>">

<!-- 플러스값이 저장된다. -->
<input type="hidden" name="miletotalprice" value="<%= orefund.FOneItem.Forgmileagesum %>">
<input type="hidden" name="tencardspend" value="<%= orefund.FOneItem.Forgcouponsum %>">
<input type="hidden" name="allatdiscountprice" value="<%= orefund.FOneItem.Forgallatdiscountsum %>">
<input type="hidden" name="depositsum" value="<%= orefund.FOneItem.Forgdepositsum %>">
<input type="hidden" name="giftcardsum" value="<%= orefund.FOneItem.Forggiftcardsum %>">

<!-- requireupche, requiremakerid 는 접수 이후에 수정할 수 없다. -->
<!--
requiremakerid 가 빈값이면 텐텐회수, requiremakerid 10x10logistics 이면 텐텐물류 고객반품, 기타 업체반품
-->
<input type="hidden" name="requireupche" value="<%= ocsaslist.FOneItem.Frequireupche %>">
<input type="hidden" name="requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">

<!-- ====================================================================== -->
<!-- 1. 이전 CS 내역                                                        -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_prev_cslist_3PL.asp" -->

<!-- ====================================================================== -->
<!-- 2. CS 마스터 정보                                                      -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_master_info_3PL.asp" -->

<!-- ====================================================================== -->
<!-- 3. 상품정보                                                            -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_item_list_3PL.asp" -->

<!-- ====================================================================== -->
<!-- 4. 다른상품 맞교환 출고 상품정보                                       -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_change_item_list_3PL.asp" -->

</table>

<!-- ====================================================================== -->
<!-- 5. 취소/환불/업체정산 정보                                             -->
<!-- ====================================================================== -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" width="50%" valign="top">

    </td>
    <td bgcolor="#FFFFFF" valign="top" align="left">

		<% if IsCSReturnProcess(divcd) then %>
        <br>
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="BABABA">
        <tr  bgcolor="FFFFFF" >
            <td>
            	<% '업체반품/텐텐반품 일 경우만 사용할 수 있다. %>
            	<input type="checkbox" name="ForceReturnByTen" onClick="CheckForceReturnByTen(this)" <% if (Not IsStatusRegister) or ((divcd <> "A004") and (divcd <> "A010")) then %>disabled<% end if %>>텐바이텐 물류센터로 <font color="red">업체배송 상품 회수</font> (여러 브랜드 동시접수 가능)<br>
            	<input type="checkbox" name="ForceReturnByCustomer" onClick="CheckForceReturnByCustomer(this)" <% if (Not IsStatusRegister) or ((divcd <> "A004") and (divcd <> "A010")) then %>disabled<% end if %>>텐바이텐 물류센터로 <font color="red"> 고객 직접반품</font> (여러 브랜드 동시접수 가능)
            	<% if (Not IsStatusRegister) then %>
            		<% if (divcd = "A004") then %>
            			<input class="csbutton" type="button" value="고객직접반품->회수신청 전환" onClick="ChangeDivcdToA010(frmaction)" onFocus="blur()" <% if (Not IsStatusEdit) then %>disabled<% end if %>>
            		<% elseif (divcd = "A010") then %>
            			<input class="csbutton" type="button" value="회수신청->고객직접반품 전환" onClick="ChangeDivcdToA004(frmaction)" onFocus="blur()" <% if (Not IsStatusEdit) then %>disabled<% end if %>>
            		<% end if %>
            	<% end if %>
            </td>
        </tr>
        </table>
        <% end if %>

    </td>
</tr>
</table>
<!-- ====================================================================== -->
<!-- 5. 취소/환불/업체정산 정보                                             -->
<!-- ====================================================================== -->

<!-- ====================================================================== -->
<!-- 6. 버튼                                                                -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_button_3PL.asp"   -->
<!-- ====================================================================== -->
<!-- 6. 버튼                                                                -->
<!-- ====================================================================== -->

</form>

<script type="text/javascript">

// 페이지 시작시 작동하는 스크립트
function getOnload(){

	SetForceReturnByTen(frmaction);
	SetForceReturnByCustomer(frmaction);

	<% if (IsStatusRegister) and (IsCSCancelProcess(divcd)) and (ckAll = "on") then %>
	    // 배송비 동시취소
	    CheckUpcheDeliverPay(frmaction);

		// 상품 전체 체크 된 경우 체크안된 배송비 동시체크
		CheckBeasongPayIfAllItemSelected(frmaction);

	    // 마일리지, 할인권 환원, 배송비 차감 등 체크
	    CheckMileageETC(frmaction);
	<% end if %>

	// 체크된 상품/배송비 색바꾸기
	AnCheckClickAll(frmaction);

	// 재계산
    // CheckForItemChanged();
    CalculateAndApplyItemCostSum(frmaction);

	// 선택않된 상품 안보이기
	if (IsStatusRegister != true) {
		ShowOnlySelectedItem(frmaction);
	}

	if (IsStatusFinishing && (divcd == "A007" || divcd == "A003")) {
		if ((divcd == "A003") && (!frmaction.returnmethod)) {
			alert("결제완료 이전 주문에 대해 환불할 수 없습니다.");
			if (orderserial != "<%= exceptOrderserial %>") {
				frmaction.finishbutton.disabled = true;
			}
		} else {
			if (divcd == "A007" || ((divcd == "A003") && (frmaction.returnmethod.value=="R007"))) {
				alert('이곳에서 완료처리 하여도 \n\n\n신용카드 승인취소/무통장 환불처리는 이루어 지지 않으니 유의하시기 바랍니다.!\n\n\n\n\n\n');
			}
		}
	}

	if (IsStatusFinishing == true) {
        if (frmaction.add_upchejungsandeliverypay) {
	        frmaction.add_upchejungsandeliverypay.disabled = true;
	        frmaction.add_upchejungsancause.disabled = true;
        }
	}

	if (IsDeletedCS) {
		alert('삭제된 내역입니다.');
	}

	if ((IsStatusRegister==true)&&(IsTicketOrder==true)&&(ticketCancelDisabled==true)){
	    alert('티켓 주문 취소 불가 ' + ticketCancelStr);
	}

	if ((IsStatusRegister==true)&&(IsTravelOrder==true)){
	    alert('\n\n =========== 여행상품 주문입니다 =========== \n\n');
	}

	if ((IsStatusRegister==true) && ((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true)) {
	    alert('주문 취소 불가 : ' + OrderCancelDisableStr);
	}

	<% if (Not IsStatusRegister) then %>
		if (frmaction.contents_jupsu) {
			resizeTextArea(document.getElementById("contents_jupsu"), 40);
		}

		if (frmaction.contents_finish) {
			resizeTextArea(document.getElementById("contents_finish"), 40);
		}

		if (frmaction.contents_finish1) {
			resizeTextArea(document.getElementById("contents_finish1"), 40);
		}
	<% end if %>

	if (parent && parent.frames['ifrAct'] && document.getElementById("btnFinishReturn") && document.getElementById("btnFinishReturn").disabled === false) {
		document.getElementById("btnFinishReturn").click();
	}
}

window.onload = getOnload;

</script>

<%
set oordermaster = Nothing
set ocsOrderDetail = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->