<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
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
	response.write "<br><br>������ ���ѵǾ����ϴ�.(���� �α״� ����˴ϴ�.)"
	dbget.close()
	response.end
end if

'[�ڵ�����]
'------------------------------------------------------------------------------
'A008			�ֹ����
'
'A004			��ǰ����(��ü���)
'A010			ȸ����û(�ٹ����ٹ��)
'
'A001			������߼�
'A002			���񽺹߼�
'
'A200			��Ÿȸ��
'
'A000			�±�ȯ���
'A100			��ǰ���� �±�ȯ���
'
'A009			��Ÿ����
'A006			�������ǻ���
'A700			��ü��Ÿ����
'
'A003			ȯ��
'A005			�ܺθ�ȯ�ҿ�û
'A007			ī��,��ü,�޴�����ҿ�û
'
'A011			�±�ȯȸ��(�ٹ����ٹ��)
'A012			�±�ȯ��ǰ(��ü���)

'A111			��ǰ���� �±�ȯȸ��(�ٹ����ٹ��)
'A112			��ǰ���� �±�ȯ��ǰ(��ü���)

'[��������]
'------------------------------------------------------------------------------
'CSFunction.asp
'
'dim IsStatusRegister			'����
'dim IsStatusEdit				'����
'dim IsStatusFinishing			'ó���Ϸ� �õ�
'dim IsStatusFinished			'ó���Ϸ�

'dim IsDisplayPreviousCSList	'���� CS ����
'dim IsDisplayCSMaster			'CS ����������
'dim IsDisplayItemList			'��ǰ���
'dim IsDisplayChangeItemList	'�ٸ���ǰ �±�ȯ��� ��ǰ���
'dim IsDisplayRefundInfo		'ȯ������
'dim IsDisplayButton			'��ư
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

'CS���������� ��������
dim ocsaslist
set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id

	if (id<>"") then
	    ocsaslist.GetOneCSASMaster_3PL
	end if

'CS���������� ������ ������� �ű� ����
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

''ȯ������
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

''�ֹ� ����Ÿ
dim oordermaster, IsCalculateAddBeasongPayNeed
set oordermaster = new COrderMaster
	oordermaster.FRectOrderSerial = orderserial

	if Left(orderserial,1)="A" then
	    set oordermaster.FOneItem = new COrderMasterItem
	else
	    oordermaster.QuickSearchOrderMaster_3PL
	end if

'' ���� 6���� ���� ���� �˻�
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

'���ʽ�����
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
	'// ������ : ���� ��������(��ǰ���� �±�ȯ ����)
	'// �������� : ����� ��������
	'// =======================================================================

	''���ֹ���( + ��ȯ�ֹ�) �ݾ� ��ȸ(��� ��ǰ ����)
	Call GetOrgOrderPriceInfo(orderserial, curr_subtotalprice, curr_sumPaymentEtc, curr_tencardspend, curr_miletotalprice, curr_spendmembership, curr_allatdiscountprice, curr_depositsum, curr_giftcardsum, curr_percentcouponsum)

	''TODO : �������� ���� ȯ��ó��
	''--update
	''dbo.tbl_as_refund_info
	''set orgcouponsum = 2400
	''where asid = 1323982

	if (IsStatusRegister) then

		curr_itemcostsum = (curr_subtotalprice + curr_tencardspend + curr_miletotalprice + curr_spendmembership + curr_allatdiscountprice)

		curr_beasongpaysum = 0

		'// �귣�庰 ��ǰ�ݾ�, ��ۺ�
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


		'������ �ʱⰪ ����
		if IsChangeOrder then
			'��ȯ�ֹ��� �귣��ݾ׸�
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

''��ȯ������
dim prevrefund, prevrefundsum, csbeasongpaysum
set prevrefund = New CCSASList
	prevrefund.FRectOrderSerial = orderserial
	prevrefundsum = prevrefund.GetPrevRefundSum

'��ۺ� ��� ���� ��ۺ�ȯ���� �̷���� �ݾ�
csbeasongpaysum = prevrefund.GetPrevRefundCSDeliveryPaySum

''���� ������ ȯ������
dim orefundInfo, prevrefundhistorycnt

set orefundInfo = New CCSASList
orefundInfo.FCurrpage = 1
orefundInfo.FPageSize = 10
orefundInfo.FRectUserID = oordermaster.FOneItem.FUserID

if (oordermaster.FOneItem.FUserID="") then
	prevrefundhistorycnt = "����"
else
    orefundInfo.GetHisOldRefundInfo

    prevrefundhistorycnt = orefundInfo.FTotalCount
end if

'==============================================================================
'������ ����Ʈ�� ����
'==============================================================================
if (IsStatusRegister) then

	'// �����ڸ�=������ȯ�� ������
	orefund.FOneItem.Frebankownername = oordermaster.FOneItem.FBuyname

	'// �⺻���� ����
	if InStr("A004,A001,A002,A200,A009,A006,A700,A000", divcd) then
		if Not IsNull(session("ssBctCname")) then
			ocsaslist.FOneItem.Fcontents_jupsu = "�ٹ����� ������ " + CStr(session("ssBctCname")) + " �Դϴ�"
		end if
	end if
end if

'==============================================================================
'��ǰ��ҷ� ��ۺ� ������ ����(�ܺθ�, �ؿܹ��, ���δ��� XX)
IsCalculateAddBeasongPayNeed = (oordermaster.FOneItem.Fjumundiv <> "5") and (oordermaster.FOneItem.FDlvcountryCode = "KR")

dim oupchebeasongpay
set oupchebeasongpay = new COrderMaster
	if (orderserial<>"") then
		oupchebeasongpay.FRectOrderSerial = orderserial
		oupchebeasongpay.getUpcheBeasongPayList
	end if

'==============================================================================
'�ֹ� ������
dim ocsOrderDetail
set ocsOrderDetail = new CCSASList
	ocsOrderDetail.FRectCsAsID = ocsaslist.FOneItem.FId
	ocsOrderDetail.FRectOrderSerial = orderserial

	if (oordermaster.FRectOldOrder = "on") then
	    ocsOrderDetail.FRectOldOrder = "on"
	end if

	ocsOrderDetail.GetOrderDetailByCsDetailNew_3PL

'==============================================================================
'��ǰ���� �±�ȯ
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
'// �±�ȯȸ��
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
'������������
dim oetcpayment, realdepositsum, realgiftcardsum, realSubPaymentSum, orgSubPaymentSum

set oetcpayment = new COrderMaster
realdepositsum = 0
realgiftcardsum = 0
realSubPaymentSum = 0
if (orderserial<>"") then
	oetcpayment.FRectOrderSerial = orderserial
	oetcpayment.getEtcPaymentList

	'200 : ��ġ��
	for i = 0 to oetcpayment.FResultCount - 1
		if (CStr(oetcpayment.FItemList(i).Facctdiv) = "200") then
			realdepositsum = oetcpayment.FItemList(i).FrealPayedsum
		end if
	next

	'900 : Giftī��
	for i = 0 to oetcpayment.FResultCount - 1
		if (CStr(oetcpayment.FItemList(i).Facctdiv) = "900") then
			realgiftcardsum = oetcpayment.FItemList(i).FrealPayedsum
		end if
	next
end if
realSubPaymentSum = realdepositsum+realgiftcardsum
''orgSubPaymentSum = orgdepositsum+orggiftcardsum

'==============================================================================
'���� �ְ������ܱݾ�
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

    ''�ҺҰ�����
    ''installment = Right(cardcodeall,2) 14|26|00 ==> 14|26|00|1 ''������ �ڵ� �κ���� ���ɿ��� (2011-08-25)--------
    IF Not IsNULL(cardcodeall) THEN
        cardcodeall= TRIM(cardcodeall)
        cardcodeall = LEft(cardcodeall,10)   '''������� �ڵ� �̻��� (�� �Ǵ� �̻��� ��)
    END IF

    if (LEN(TRIM(cardcodeall))=10) then
        if (Right(cardcodeall,1)="1") then
            cardPartialCancelok = "Y"
        elseif (Right(cardcodeall,1)="0") then
            cardPartialCancelok = "N"
            if (cardcancelerrormsg="") then cardcancelerrormsg  = "�κ���� <strong>�Ұ�</strong> �ŷ� (������ ī�� or ���հŷ�)"
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
'���ֹ� ��ǰ�ݾ�
dim orgitemcostsum, orgpercentcouponpricesum

'������ǰ �հ�ݾ�(inc_cs_action_item_list.asp ���� ���ȴ�)
dim regitemcostsum, regpercentcouponpricesum

'������ id(orderdetailidx)
dim distinctid

'==============================================================================
''���� �Ұ��� �޼���
dim JupsuInValidMsg

if (Left(orderserial,1)<>"A") and (oordermaster.FResultCount<1) then
    response.write "<br><br>!!! ���� �ֹ������̰ų� �ֹ� ������ �����ϴ�. - ������ ���� ���"
    dbget.close()	:	response.End
end if

''���� ���� ����
dim IsJupsuProcessAvail
if (oordermaster.FResultCount>0) then
	if IsChangeOrder then
		IsJupsuProcessAvail = ocsaslist.FOneItem.IsChangeAsRegAvail(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
	elseif IsMinusOrder then
		IsJupsuProcessAvail = false
		JupsuInValidMsg = "���̳ʽ��ֹ��� ���� CS������ �� �����ϴ�."
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
		JupsuInValidMsg = "�Ϸ���� ���� �ֹ���� �������� �ֽ��ϴ�.\n���� �������� �Ϸ�ó�� �ϼ���."
	end if
end if

'��üó���Ϸ���� ����
dim IsUpcheConfirmState
IsUpcheConfirmState = (ocsaslist.FOneItem.FCurrState="B006")

'// �ù�� ���� ������������
dim IsLogicsSended
IsLogicsSended = (ocsaslist.FOneItem.FCurrState<>"B001")

''Ƽ��Order�ΰ�� ��� ����.
Dim mayTicketCancelChargePro : mayTicketCancelChargePro=0
Dim ticketCancelStr : ticketCancelStr =""
Dim ticketCancelDisabled : ticketCancelDisabled = false   '''Ƽ�� ��� �Ұ�����.

if (IsTicketOrder) and (IsStatusRegister) and (oordermaster.FOneItem.IsPayedOrder) then
    '' �����ֹ����� ��Ҽ����ᰡ ����.
    if (Not isThisdateCancel) then
        call TicketOrderCheck(orderserial,mayTicketCancelChargePro,ticketCancelDisabled, ticketCancelStr)
		if (session("ssBctId") = "nownhere21") and ticketCancelDisabled = True then
			'2018-02-21, skyer9
			ticketCancelDisabled = False
		end if
    end if
end if

''��Ұ����� �ֹ�����
IsOrderCancelDisabled = False
OrderCancelDisableStr = ""
if (IsGiftingOrder or IsGiftiConOrder) and (IsCSCancelProcess(divcd) or IsCSReturnProcess(divcd)) then
	IsOrderCancelDisabled = True
	OrderCancelDisableStr = "������/����Ƽ�� ���� �ֹ��Դϴ�.\n\n��ǰ���� �� �� �����ϴ�.[�ֹ���ҽ� ��ġ��ȯ�Ҹ� ����]"
end if

'// �����ֹ�
dim travelItemInfoArr, travelItemExist
travelItemExist = False
if (IsTravelOrder) and (IsStatusRegister or IsStatusEdit) and (oordermaster.FOneItem.IsPayedOrder) and IsCSReturnProcess(divcd) then
	travelItemInfoArr = TravelOrderCheckArr(orderserial)
	if IsArray(travelItemInfoArr) then
		travelItemExist = True
	end if
end if

'==============================================================================
''�Ϸ�ó�� �Ұ��� �޼���
dim FinishInValidMsg

''�Ϸ�ó�� ���� ����
dim IsFinishProcessAvail

FinishInValidMsg = ""
IsFinishProcessAvail = False

if (IsStatusFinishing) then
	IsFinishProcessAvail = True

	if (IsRefASExist) and (IsRefASFinished = False) and (ocsaslist.FOneItem.Frequireupche = "Y") then
    	FinishInValidMsg = "��ü����� ��� �±�ȯȸ���� ���� �Ϸ�ó���ؾ� �±�ȯ��� �Ϸ�ó���� �� �ֽ��ϴ�."
    	IsFinishProcessAvail = False
	end if
end if

'==============================================================================
dim IsDelFinishedCSAvail : IsDelFinishedCSAvail = False
dim DelFinishedCSInValidMsg : DelFinishedCSInValidMsg = "<font color='red'>�Ϸ᳻�� �����Ұ�</font>"
dim oRefCSASList

dim HasAuthTodayDelCancelReturn : HasAuthTodayDelCancelReturn = False
dim HasAuthUpcheJungsanItemPrice : HasAuthUpcheJungsanItemPrice = False
' ������
'HasAuthUpcheJungsanItemPrice = (InStr(",ilovecozie,durida22,boyishP,hasora,bseo,skyer9,coolhas,oesesang52,hrkang97,tozzinet,gomgom,zerogirl0730,heendoongi,happy799,may0816,angela919,hk9566371,seokmi1221,pray16,jjy158,rabbit1693,", ("," & session("ssBctId") & ",")) > 0)

if IsStatusFinished and (ocsaslist.FOneITem.FDeleteyn="N") then
	if (divcd="A004") or (divcd="A010") or (divcd="A008") then
		'// ��ǰ��� �Ϸ�CS ����

		set oRefCSASList = new CCSASList
		oRefCSASList.FRectCsRefAsID = id
		oRefCSASList.GetOneCSASMaster

		if (oRefCSASList.FResultCount > 0) then
			if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
				if (oRefCSASList.FOneItem.Fcurrstate = "B007") then
					DelFinishedCSInValidMsg = "<font color='red'>�ý����� ����(ȯ�ҿϷ� �����Դϴ�.)</font>"
				else
					DelFinishedCSInValidMsg = "���� ���� ȯ��CS�� �����ϼ���."
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
				DelFinishedCSInValidMsg = "<font color='red'>�ý����� ����(���̳ʽ� �ֹ���ȣ ����)</font>"
			elseif ((divcd="A008") and (oordermaster.FOneItem.Fipkumdiv >= "4")) then
				if (ocsaslist.FOneItem.Ffinishdate < oordermaster.FOneItem.Fipkumdate) then
					DelFinishedCSInValidMsg = "��ҺҰ�. <font color='red'>�������� ���</font> �����Դϴ�."
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

''�����Ϸ� ���� ���ϸ��� ���� �ʿ��� ���
dim exceptOrderserial : exceptOrderserial = "xxxxxxxxx"

'// �ӽ� �̺�Ʈ
'// �귣�� : laundrymat
'// ���ݾ� : 50000
'// �ֹ��� : 1
'// �Ⱓ : 2016.03.07~2016.03.29
'// ������ ����
dim IsTempEventAvail : IsTempEventAvail = True
dim IsTempEventAvail_Str : IsTempEventAvail_Str = ""
dim IsTempEventAvail_Makerid

IF application("Svr_Info")="Dev" THEN
	IsTempEventAvail_Makerid = "noulnabi"
else
	IsTempEventAvail_Makerid = "laundrymat"
end if

''����
IsTempEventAvail = IsTempEventAvail and IsStatusRegister and (Not IsOutMallOrder)

''��ǰ
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
var HasAuthUpcheJungsanItemPrice	= <%= LCase(HasAuthUpcheJungsanItemPrice) %>;	// ������
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

// �Ѱ��� �귣�常 ���ð����Ѱ�
// ��ǰ����(����), �±�ȯ���, ������߼�, ���񽺹߼�, ��Ÿȸ��, �������ǻ���, ȸ����û(�ٹ����ٹ��), ��ü��Ÿ����
var IsOnlyOneBrandAvailable		= <%= LCase(InStr("A004,A000,A001,A002,A200,A006,A010,A700", divcd) > 0) %>;

var IsDeletedCS 				= <%= LCase(ocsaslist.FOneITem.FDeleteyn = "Y") %>;

var ERROR_MSG_TRY_MODIFY		= "<%= ERROR_MSG_TRY_MODIFY %>";

var CDEFAULTBEASONGPAY 		= 2000; // �ٹ����� �⺻ ��ۺ�
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
	v.value = v.value + "\n\n+" + d.format("yyyy-mm-dd HH:MM:ss") + "  ������ <%= session("ssBctCname") %>�Դϴ�.\n";
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
		alert("üũ�� �ʿ�����ϴ�.");
		return;
	}

	if (IsAllSelected(frm) == true) {
		alert("��ü����Դϴ�.üũ�� �ʿ�����ϴ�.");
		evt_chk_need.value = "N";
		return;
	}

	/*
	if (frm.gubun01.value == "") {
		alert("���� ���������� �Է��ϼ���.");
		return;
	}

	if ((frm.gubun01.value != "C004") || (frm.gubun02.value != "CD01")) {
		alert("������� �̿ܿ��� üũ�� �ʿ�����ϴ�.");
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
		alert("���õ� ��ǰ�� �����ϴ�.");
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

<!-- �÷������� ����ȴ�. -->
<input type="hidden" name="miletotalprice" value="<%= orefund.FOneItem.Forgmileagesum %>">
<input type="hidden" name="tencardspend" value="<%= orefund.FOneItem.Forgcouponsum %>">
<input type="hidden" name="allatdiscountprice" value="<%= orefund.FOneItem.Forgallatdiscountsum %>">
<input type="hidden" name="depositsum" value="<%= orefund.FOneItem.Forgdepositsum %>">
<input type="hidden" name="giftcardsum" value="<%= orefund.FOneItem.Forggiftcardsum %>">

<!-- requireupche, requiremakerid �� ���� ���Ŀ� ������ �� ����. -->
<!--
requiremakerid �� ���̸� ����ȸ��, requiremakerid 10x10logistics �̸� ���ٹ��� ����ǰ, ��Ÿ ��ü��ǰ
-->
<input type="hidden" name="requireupche" value="<%= ocsaslist.FOneItem.Frequireupche %>">
<input type="hidden" name="requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">

<!-- ====================================================================== -->
<!-- 1. ���� CS ����                                                        -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_prev_cslist_3PL.asp" -->

<!-- ====================================================================== -->
<!-- 2. CS ������ ����                                                      -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_master_info_3PL.asp" -->

<!-- ====================================================================== -->
<!-- 3. ��ǰ����                                                            -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_item_list_3PL.asp" -->

<!-- ====================================================================== -->
<!-- 4. �ٸ���ǰ �±�ȯ ��� ��ǰ����                                       -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_change_item_list_3PL.asp" -->

</table>

<!-- ====================================================================== -->
<!-- 5. ���/ȯ��/��ü���� ����                                             -->
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
            	<% '��ü��ǰ/���ٹ�ǰ �� ��츸 ����� �� �ִ�. %>
            	<input type="checkbox" name="ForceReturnByTen" onClick="CheckForceReturnByTen(this)" <% if (Not IsStatusRegister) or ((divcd <> "A004") and (divcd <> "A010")) then %>disabled<% end if %>>�ٹ����� �������ͷ� <font color="red">��ü��� ��ǰ ȸ��</font> (���� �귣�� �������� ����)<br>
            	<input type="checkbox" name="ForceReturnByCustomer" onClick="CheckForceReturnByCustomer(this)" <% if (Not IsStatusRegister) or ((divcd <> "A004") and (divcd <> "A010")) then %>disabled<% end if %>>�ٹ����� �������ͷ� <font color="red"> �� ������ǰ</font> (���� �귣�� �������� ����)
            	<% if (Not IsStatusRegister) then %>
            		<% if (divcd = "A004") then %>
            			<input class="csbutton" type="button" value="��������ǰ->ȸ����û ��ȯ" onClick="ChangeDivcdToA010(frmaction)" onFocus="blur()" <% if (Not IsStatusEdit) then %>disabled<% end if %>>
            		<% elseif (divcd = "A010") then %>
            			<input class="csbutton" type="button" value="ȸ����û->��������ǰ ��ȯ" onClick="ChangeDivcdToA004(frmaction)" onFocus="blur()" <% if (Not IsStatusEdit) then %>disabled<% end if %>>
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
<!-- 5. ���/ȯ��/��ü���� ����                                             -->
<!-- ====================================================================== -->

<!-- ====================================================================== -->
<!-- 6. ��ư                                                                -->
<!-- ====================================================================== -->
<!-- #include virtual="/admin/etc/3pl/cscenter/action/include/inc_cs_action_button_3PL.asp"   -->
<!-- ====================================================================== -->
<!-- 6. ��ư                                                                -->
<!-- ====================================================================== -->

</form>

<script type="text/javascript">

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){

	SetForceReturnByTen(frmaction);
	SetForceReturnByCustomer(frmaction);

	<% if (IsStatusRegister) and (IsCSCancelProcess(divcd)) and (ckAll = "on") then %>
	    // ��ۺ� �������
	    CheckUpcheDeliverPay(frmaction);

		// ��ǰ ��ü üũ �� ��� üũ�ȵ� ��ۺ� ����üũ
		CheckBeasongPayIfAllItemSelected(frmaction);

	    // ���ϸ���, ���α� ȯ��, ��ۺ� ���� �� üũ
	    CheckMileageETC(frmaction);
	<% end if %>

	// üũ�� ��ǰ/��ۺ� ���ٲٱ�
	AnCheckClickAll(frmaction);

	// ����
    // CheckForItemChanged();
    CalculateAndApplyItemCostSum(frmaction);

	// ���þʵ� ��ǰ �Ⱥ��̱�
	if (IsStatusRegister != true) {
		ShowOnlySelectedItem(frmaction);
	}

	if (IsStatusFinishing && (divcd == "A007" || divcd == "A003")) {
		if ((divcd == "A003") && (!frmaction.returnmethod)) {
			alert("�����Ϸ� ���� �ֹ��� ���� ȯ���� �� �����ϴ�.");
			if (orderserial != "<%= exceptOrderserial %>") {
				frmaction.finishbutton.disabled = true;
			}
		} else {
			if (divcd == "A007" || ((divcd == "A003") && (frmaction.returnmethod.value=="R007"))) {
				alert('�̰����� �Ϸ�ó�� �Ͽ��� \n\n\n�ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� ������ �����Ͻñ� �ٶ��ϴ�.!\n\n\n\n\n\n');
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
		alert('������ �����Դϴ�.');
	}

	if ((IsStatusRegister==true)&&(IsTicketOrder==true)&&(ticketCancelDisabled==true)){
	    alert('Ƽ�� �ֹ� ��� �Ұ� ' + ticketCancelStr);
	}

	if ((IsStatusRegister==true)&&(IsTravelOrder==true)){
	    alert('\n\n =========== �����ǰ �ֹ��Դϴ� =========== \n\n');
	}

	if ((IsStatusRegister==true) && ((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true)) {
	    alert('�ֹ� ��� �Ұ� : ' + OrderCancelDisableStr);
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