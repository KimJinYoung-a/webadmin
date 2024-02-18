<%
'###########################################################
' Description : cs���� �����Լ�
' History : �̻� ����
'###########################################################

dim GC_IsOldOrder
GC_IsOldOrder = false

function RegCSMasterRefundInfoBeforeCancel(asid, orderserial, reguserid, refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, byref contents_finish, byref refundmileagesum, byref refunddepositsum, byref refundgiftcardsum)
	dim sqlStr

	dim IsOldOrder

	dim returnmethod, orgsubtotalprice, orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum
	dim rebankname, rebankaccount, rebankownername, paygateTid
	dim orggiftcardsum, orgdepositsum

	dim refundbeasongpay, refunddeliverypay, refundadjustpay

	dim ipkumdiv, sumPaymentEtc

	IsOldOrder = CheckIsOldOrder(orderserial)

	returnmethod			= "R007"			'������ ȯ��

	orgsubtotalprice		= 0
	orgitemcostsum			= 0
	orgbeasongpay			= 0
	orgmileagesum			= 0
	orgcouponsum			= 0
	orgallatdiscountsum		= 0
	orggiftcardsum			= 0
	orgdepositsum			= 0

	'refundrequire			= 0
	'canceltotal			= 0
	'refunditemcostsum		= 0
	refundmileagesum		= 0
	'refundcouponsum		= 0
	'allatsubtractsum		= 0
	refundbeasongpay		= 0
	refunddeliverypay		= 0
	refundgiftcardsum		= 0
	refunddepositsum		= 0
	refundadjustpay			= 0

	rebankname				= ""
	rebankaccount			= ""
	rebankownername			= ""
	paygateTid				= ""

	refundrequire			= refundrequire*1

	'==========================================================================
	'���ֹ�����
	sqlStr = " select top 1 "
	sqlStr = sqlStr + " 	subtotalprice "
	sqlStr = sqlStr + " 	, totalsum "
	sqlStr = sqlStr + " 	, miletotalprice "
	sqlStr = sqlStr + " 	, tencardspend "
	sqlStr = sqlStr + " 	, allatdiscountprice "
	sqlStr = sqlStr + " 	, ipkumdiv "
	sqlStr = sqlStr + " 	, sumPaymentEtc "
	sqlStr = sqlStr + " from "

	if (IsOldOrder) then
		sqlStr = sqlStr + " 	db_log.dbo.tbl_old_order_master_2003 "
	else
		sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master "
	end if

	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	orderserial = '" + CStr(orderserial) + "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
		orgsubtotalprice		= rsget("subtotalprice")
		orgitemcostsum			= rsget("totalsum")
		orgbeasongpay			= 0								'���߿� ������.
		orgmileagesum			= rsget("miletotalprice")
		orgcouponsum			= rsget("tencardspend")
		orgallatdiscountsum		= rsget("allatdiscountprice")
		orggiftcardsum			= 0
		orgdepositsum			= 0

		ipkumdiv				= rsget("ipkumdiv")
		sumPaymentEtc			= rsget("sumPaymentEtc")
    end if
    rsget.close

	'==========================================================================
	'��ۺ�
	sqlStr = " select "
	sqlStr = sqlStr + " 	IsNull(sum(case when itemid = 0 then itemcost*itemno else 0 end), 0) as orgbeasongpay "
	sqlStr = sqlStr + " from "

	if (IsOldOrder) then
	    sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003"
	else
	    sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail"
	end if

	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and cancelyn <> 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
		orgitemcostsum			= orgitemcostsum - rsget("orgbeasongpay")
		orgbeasongpay			= rsget("orgbeasongpay")
    end if
    rsget.close

	'==========================================================================
	'200 : ��ġ��, 900 : Giftī��
	sqlStr = " select "
	sqlStr = sqlStr + " 	IsNull(sum(case when acctdiv = '200' then IsNull(realPayedsum, 0) else 0 end), 0) as orgdepositsum "
	sqlStr = sqlStr + " 	, IsNull(sum(case when acctdiv = '900' then IsNull(realPayedsum, 0) else 0 end), 0) as orggiftcardsum "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	orderserial = '" + CStr(orderserial) + "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
		orggiftcardsum			= rsget("orggiftcardsum")
		orgdepositsum			= rsget("orgdepositsum")
    end if
    rsget.close

	'==========================================================================
	if ((orgsubtotalprice - sumPaymentEtc) < refundrequire) and (orgmileagesum > 0) then
		if (refundrequire >= orgmileagesum) then
			refundmileagesum = orgmileagesum
			refundrequire = refundrequire - orgmileagesum
		else
			refundmileagesum = refundrequire
			refundrequire = 0
		end if

		contents_finish = contents_finish + vbCrLf + "���ϸ��� ȯ�� : " + CStr(refundmileagesum*-1)
	end if

	if ((orgsubtotalprice - sumPaymentEtc) < refundrequire) and (orgdepositsum > 0) then
		if (refundrequire >= orgdepositsum) then
			refunddepositsum = orgdepositsum
			refundrequire = refundrequire - orgdepositsum
		else
			refunddepositsum = refundrequire
			refundrequire = 0
		end if

		contents_finish = contents_finish + vbCrLf + "��ġ�� ȯ�� : " + CStr(refunddepositsum*-1)
	end if

	if ((orgsubtotalprice - sumPaymentEtc) < refundrequire) and (orggiftcardsum > 0) then
		if (refundrequire >= orggiftcardsum) then
			refundgiftcardsum = orggiftcardsum
			refundrequire = refundrequire - orggiftcardsum
		else
			refundgiftcardsum = refundrequire
			refundrequire = 0
		end if

		contents_finish = contents_finish + vbCrLf + "Giftī�� ȯ�� : " + CStr(refundgiftcardsum*-1)
	end if

	if (ipkumdiv < 4) then
		refundrequire = 0
	end if

	if (refundrequire <= 0) then
		returnmethod = "R000"
	end if

	'==========================================================================
    'CS Master ȯ�� �������� ����
    Call RegCSMasterRefundInfo(asid, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum*-1, refundcouponsum*-1, allatsubtractsum*-1, refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
    Call AddCSMasterRefundInfo(asid, orggiftcardsum, orgdepositsum, refundgiftcardsum*-1, refunddepositsum*-1)

	RegCSMasterRefundInfoBeforeCancel = 0
	if (refundrequire > 0) and (ipkumdiv >= 4) then
        'ȯ�� ������ �ִ��� üũ �� ������ȯ��/���ϸ���ȯ��/�ſ�ī����� CS ���� ���
        RegCSMasterRefundInfoBeforeCancel = CheckNRegRefund(asid, orderserial,reguserid)
	end if

end function

function EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	dim sqlStr

	dim IsOldOrder

	IsOldOrder = CheckIsOldOrder(orderserial)

	sqlStr = " update "

	if (IsOldOrder) then
		sqlStr = sqlStr + " 	db_log.dbo.tbl_old_order_master_2003 "
	else
		sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master "
	end if

    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	miletotalprice = miletotalprice - " + CStr(refundmileagesum) + " "    + VbCrlf
    sqlStr = sqlStr + " 	, tencardspend = tencardspend - " + CStr(refundcouponsum) + " "    + VbCrlf
    sqlStr = sqlStr + " 	, allatdiscountprice = allatdiscountprice - " + CStr(allatsubtractsum) + " "    + VbCrlf
    sqlStr = sqlStr + " 	, sumPaymentEtc = sumPaymentEtc - " + CStr(refunddepositsum) + " - " + CStr(refundgiftcardsum) + " "    + VbCrlf
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	orderserial = '" + CStr(orderserial) + "' "
    dbget.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc "
    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	realPayedsum = realPayedsum - " + CStr(refunddepositsum) + " "    + VbCrlf
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and acctdiv = '200' "
    dbget.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc "
    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	realPayedsum = realPayedsum - " + CStr(refundgiftcardsum) + " "    + VbCrlf
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and acctdiv = '900' "
    dbget.Execute sqlStr

end function

function CheckIsOldOrder(orderserial)
    ''���� �ֹ����� Check
    dim sqlStr

    if orderserial="" or isnull(orderserial) then
        CheckIsOldOrder=""
        exit function
    end if

	sqlStr = " select orderserial from db_order.dbo.tbl_order_master where orderserial='" & orderserial & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
	    CheckIsOldOrder = False
	else
        CheckIsOldOrder = True
    end if
    rsget.close

    if (CheckIsOldOrder) then
        sqlStr = " select orderserial from db_log.dbo.tbl_old_order_master_2003 where orderserial='" & orderserial & "'"
	    rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	    if rsget.Eof then
	        CheckIsOldOrder = False
	    end if
	    rsget.close
    end if
end function

function CheckNotFinishedCancelCSExist(orderserial)
    dim sqlStr

	sqlStr = " select top 1 orderserial from db_cs.dbo.tbl_new_as_list where divcd = 'A008' and currstate <> 'B007' and deleteyn <> 'Y' and orderserial='" & orderserial & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
	    CheckNotFinishedCancelCSExist = True
	else
        CheckNotFinishedCancelCSExist = False
    end if
    rsget.close
end function

function GetCsOrderSerial(orderserial)
    dim sqlStr, csorderserial

    csorderserial = ""
	sqlStr = " select orderserial from db_order.dbo.tbl_order_master where linkorderserial='" & orderserial & "' and sitename='10x10_cs' and jumundiv <> '9' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
	    csorderserial = rsget("orderserial")
    end if
    rsget.close

    GetCsOrderSerial = csorderserial
end function

function CheckNotFinishedCancelCSMakeridList(orderserial)
    dim sqlStr, resultStr

	sqlStr = " select distinct makerid from db_cs.dbo.tbl_new_as_list where divcd = 'A008' and currstate <> 'B007' and deleteyn <> 'Y' and orderserial='" & orderserial & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	resultStr = ""
	if Not rsget.Eof then
		resultStr = "|"
	    do until rsget.eof
			resultStr = resultStr & rsget("makerid") & "|"
			rsget.MoveNext
		loop
	end if
	rsget.Close

	CheckNotFinishedCancelCSMakeridList = resultStr
end function

function CheckCSFinished(asid)
    dim sqlStr

	sqlStr = " select top 1 currstate from [db_cs].[dbo].[tbl_new_as_list] where id = " + CStr(asid) + " "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		CheckCSFinished = (rsget("currstate") = "B007")
    rsget.close
end function

function CheckCSFinished_3PL(asid)
    dim sqlStr

	sqlStr = " select top 1 currstate from [db_threepl].[dbo].[tbl_tpl_as_list] where id = " + CStr(asid) + " "
	rsget_TPL.Open sqlStr,dbget_TPL,1
		CheckCSFinished_3PL = (rsget_TPL("currstate") = "B007")
    rsget_TPL.close
end function

function CheckFreeReturnDeliveryAvail(orderserial, makerid, startDate, endDate, reducedPriceSUM, csCnt)
    dim sqlStr, result

	If Left(Now(),10) < startDate Or Left(Now(),10) > endDate Then
		CheckFreeReturnDeliveryAvail = "�̺�Ʈ �Ⱓ �ƴ�[" & startDate & "~" & endDate & "]"
		Exit Function
	End If

	sqlStr = " select IsNull(sum(reducedPrice*itemno),0) as reducedPriceSUM "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " [db_order].[dbo].[tbl_order_detail] "
	sqlStr = sqlStr & " where "
	sqlStr = sqlStr & " 	1 = 1 "
	sqlStr = sqlStr & " 	and orderserial = '" & orderserial & "' "
	sqlStr = sqlStr & " 	and makerid = '" & makerid & "' "
	sqlStr = sqlStr & " 	and itemid not in (0,100) "
	sqlStr = sqlStr & " 	and cancelyn <> 'Y' "
	sqlStr = sqlStr & " 	and currstate = 7 "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		result = (rsget("reducedPriceSUM") >= reducedPriceSUM)
    rsget.Close

	If (result = False) Then
		CheckFreeReturnDeliveryAvail = "����ǰ �ݾ� ����[" & reducedPriceSUM & "��]"
		Exit Function
	End If

	sqlStr = " select count(*) as CNT from [db_cs].[dbo].[tbl_new_as_list] "
	sqlStr = sqlStr & " where "
	sqlStr = sqlStr & " 	1 = 1 "
	sqlStr = sqlStr & " 	and orderserial = '" & orderserial & "' "
	sqlStr = sqlStr & " 	and deleteyn = 'N' "
	sqlStr = sqlStr & " 	and gubun01 = 'C004' "
	sqlStr = sqlStr & " 	and gubun02 = 'CD11' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		result = (rsget("CNT") < csCnt)
	rsget.Close

	If (result = False) Then
		CheckFreeReturnDeliveryAvail = "�ֹ��� �ѹ��� ����"
		Exit Function
	End If

	CheckFreeReturnDeliveryAvail = ""

end function

function getCardRibonName(cardribbon)
    if IsNULL(cardribbon) then Exit Function

    if (cardribbon="1") then
        getCardRibonName  = "ī��"
    elseif (cardribbon="2") then
        getCardRibonName  = "����"
    elseif (cardribbon="3") then
        getCardRibonName  = "����"
    end if
end function

function FinishCSMaster(iAsid, finishuser, contents_finish)
    dim sqlStr
    dim IsCsErrStockUpdateRequire
    IsCsErrStockUpdateRequire = False

    sqlStr = "select divcd, finishdate, currstate"
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(iAsid)
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		'// finishdate üũ �ϴ°� ��(2014-05-30 skyer9 : ����ó���Ϸ�)
        IsCsErrStockUpdateRequire = ((rsget("divcd")="A000") or (rsget("divcd")="A011") or (rsget("divcd")="A100") or (rsget("divcd")="A111")) and (rsget("currstate")<>"B007")
    end if
    rsget.close

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"                      + VbCrlf
    sqlStr = sqlStr + " set finishuser='" + finishuser + "'"            + VbCrlf
	if (contents_finish <> "Finished_by_system") then
		sqlStr = sqlStr + " , contents_finish='" + contents_finish + "'"    + VbCrlf
	end if
    sqlStr = sqlStr + " , finishdate=getdate()"                         + VbCrlf
    sqlStr = sqlStr + " , currstate='B007'"                             + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(iAsid)

    dbget.Execute sqlStr

    ''�±�ȯȸ�� �Ϸ��ϰ�� ��������Ʈ. 2007.11.16
	'// ����� ó���ϴ� �� ����.(2017-01-31, skyer9)
    if (IsCsErrStockUpdateRequire) then
        sqlStr = " exec db_summary.dbo.ten_RealTimeStock_CsErr " & iAsid & ",'','" & finishuser & "'"
        dbget.Execute sqlStr
    end if
end function

function FinishCSMaster_3PL(iAsid, finishuser, contents_finish)
    dim sqlStr
    dim IsCsErrStockUpdateRequire
    IsCsErrStockUpdateRequire = False

    sqlStr = "select divcd, finishdate, currstate"
    sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_list]"
    sqlStr = sqlStr + " where id=" + CStr(iAsid)
    rsget_TPL.Open sqlStr,dbget_TPL,1
    if Not rsget_TPL.Eof then
		'// finishdate üũ �ϴ°� ��(2014-05-30 skyer9 : ����ó���Ϸ�)
        IsCsErrStockUpdateRequire = ((rsget_TPL("divcd")="A000") or (rsget_TPL("divcd")="A011") or (rsget_TPL("divcd")="A100") or (rsget_TPL("divcd")="A111")) and (rsget_TPL("currstate")<>"B007")
    end if
    rsget_TPL.close

    sqlStr = " update [db_threepl].[dbo].[tbl_tpl_as_list]"                      + VbCrlf
    sqlStr = sqlStr + " set finishuser='" + finishuser + "'"            + VbCrlf
    sqlStr = sqlStr + " , contents_finish='" + contents_finish + "'"    + VbCrlf
    sqlStr = sqlStr + " , finishdate=getdate()"                         + VbCrlf
    sqlStr = sqlStr + " , currstate='B007'"                             + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(iAsid)

    dbget_TPL.Execute sqlStr

    ''�±�ȯȸ�� �Ϸ��ϰ�� ��������Ʈ. 2007.11.16
	'// ����� ó���ϴ� �� ����.(2017-01-31, skyer9)
    if (IsCsErrStockUpdateRequire) then
        ''sqlStr = " exec db_summary.dbo.ten_RealTimeStock_CsErr " & iAsid & ",'','" & finishuser & "'"
        ''dbget_TPL.Execute sqlStr
    end if
end function

function SetStockOutByCsAs(iAsid)
    dim sqlStr
    dim resultCount	: resultCount = 0
    dim arrItemID

	'// �����ǰ�� ǰ�� ���

	'// =======================================================================
	sqlStr = " select IsNull(count(i.itemid), 0) as cnt " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join db_item.dbo.tbl_item i " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	d.itemid = i.itemid " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// �������� ǰ���� ��츸, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and i.sellyn = 'Y' " + VbCrLf
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        resultCount = resultCount + rsget("cnt")
    rsget.Close

	sqlStr = " select IsNull(count(o.itemid), 0) as cnt " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join [db_item].[dbo].tbl_item_option o " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid = o.itemid " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = o.itemoption " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// �������� ǰ���� ��츸, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption <> '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and o.optsellyn = 'Y' " + VbCrLf
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        resultCount = resultCount + rsget("cnt")
    rsget.Close

    '// =======================================================================
    SetStockOutByCsAs = resultCount
    if (resultCount < 1) then
        exit function
    end if

    '// =======================================================================
    '// 1. �ɼ� ���� ��ǰ(�Ͻ�ǰ�� ��ȯ)
    sqlStr = " update i " + VbCrLf
    sqlStr = sqlStr + " set i.sellyn = 'S', i.lastupdate = getdate() " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join db_item.dbo.tbl_item i " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	d.itemid = i.itemid " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// �������� ǰ���� ��츸, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and i.sellyn = 'Y' " + VbCrLf
    'response.write sqlStr
	dbget.Execute sqlStr

    '// =======================================================================
	'// 2-1. �ɼ� �ִ� ��ǰ(��ǰ�ڵ���)
	sqlStr = " select distinct o.itemid " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join [db_item].[dbo].tbl_item_option o " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid = o.itemid " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = o.itemoption " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// �������� ǰ���� ��츸, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption <> '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and o.optsellyn = 'Y' " + VbCrLf
    'response.write sqlStr
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    arrItemID = "-1"
	do until rsget.Eof
		arrItemID = arrItemID + "," + CStr(rsget("itemid"))
		rsget.MoveNext
	loop
	rsget.Close

	'// 2-2. �ɼ� �ִ� ��ǰ(ǰ����ȯ)
	sqlStr = " update o " + VbCrLf
	sqlStr = sqlStr + " set o.isusing = 'N', o.optsellyn = 'N' " + VbCrLf
    sqlStr = sqlStr + " from " + VbCrLf
    sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list m " + VbCrLf
    sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_detail d " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	m.id = d.masterid " + VbCrLf
    sqlStr = sqlStr + " join [db_item].[dbo].tbl_item_option o " + VbCrLf
    sqlStr = sqlStr + " on " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid = o.itemid " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption = o.itemoption " + VbCrLf
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// �������� ǰ���� ��츸, 2022-02-24, skyer9
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and T.orderserial = m.orderserial "
	sqlStr = sqlStr + " 	and T.detailidx = d.orderdetailidx "
    sqlStr = sqlStr + " 	and T.code = '05' "
    sqlStr = sqlStr + " where " + VbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
    sqlStr = sqlStr + " 	and m.id = " + CStr(iAsid) + " " + VbCrLf
    sqlStr = sqlStr + " 	and ((m.divcd in ('A008', 'A112')) or ((m.divcd = 'A900') and (d.confirmitemno < 0))) " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun01 = 'C004' " + VbCrLf
    sqlStr = sqlStr + " 	and d.gubun02 = 'CD05' " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemid <> 0 " + VbCrLf
    sqlStr = sqlStr + " 	and d.itemoption <> '0000' " + VbCrLf
    sqlStr = sqlStr + " 	and d.isupchebeasong = 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and m.deleteyn <> 'Y' " + VbCrLf
    sqlStr = sqlStr + " 	and o.optsellyn = 'Y' " + VbCrLf
    'response.write sqlStr
	dbget.Execute sqlStr

	'// 2-3. �ɼ� �ִ� ��ǰ(�ɼǰ���)
	sqlStr = " update i " + VbCrLf
	sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0), lastupdate = getdate() " + VbCrLf
	sqlStr = sqlStr + " from " + VbCrLf
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i " + VbCrLf
	sqlStr = sqlStr + " 	join ( " + VbCrLf
	sqlStr = sqlStr + " 		select itemid, sum(case when isusing = 'Y' then 1 else 0 end) optioncnt " + VbCrLf
	sqlStr = sqlStr + " 		from [db_item].[dbo].tbl_item_option " + VbCrLf
	sqlStr = sqlStr + " 		where itemid in ( " + VbCrLf
	sqlStr = sqlStr + " 			" + CStr(arrItemID) + " " + VbCrLf
	sqlStr = sqlStr + " 		) " + VbCrLf
	''sqlStr = sqlStr + " 		and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " 		group by itemid " + VbCrLf
	sqlStr = sqlStr + " 	) T " + VbCrLf
	sqlStr = sqlStr + " 	on " + VbCrLf
	sqlStr = sqlStr + " 		i.itemid = T.itemid " + VbCrLf
	'response.write sqlStr
	dbget.Execute sqlStr

	'// 2-4. �ɼ� �ִ� ��ǰ(�Ǹ����� �ɼ��� ������ ǰ��ó��)
    sqlStr = " update [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " set sellyn='N'"
	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
	sqlStr = sqlStr + " where itemid in (" + CStr(arrItemID) + ") "
	sqlStr = sqlStr + " and optioncnt=0"
	'response.write sqlStr
    dbget.Execute sqlStr

end function

function GetDefaultTitle(divcd, id, orderserial)
    dim opentitle, opencontents
    dim ipkumdiv, accountdiv, cancelyn, comm_name, ipkumdivName, accountdivName, pggubun, comm_cd
    dim sqlStr

    sqlStr = " select m.ipkumdiv, m.accountdiv, m.cancelyn, C.comm_name, isNULL(m.pggubun,'') as pggubun, C.comm_cd"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    end if
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list A"
    sqlStr = sqlStr + "     on A.orderserial='" + orderserial + "'"
    if (id<>"") then
        sqlStr = sqlStr + " and A.id=" + CStr(id)
    end if
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_cs_comm_code C"
    sqlStr = sqlStr + " on C.comm_cd='" + divcd + "'"

    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        ipkumdiv    = rsget("ipkumdiv")
        cancelyn    = rsget("cancelyn")
        comm_name   = rsget("comm_name")
        accountdiv  = Trim(rsget("accountdiv"))
        pggubun     = rsget("pggubun")
        comm_cd     = rsget("comm_cd")
    end if
    rsget.close

    if (ipkumdiv="2") then
        ipkumdivName = "�Ա� ���"
    elseif (ipkumdiv="4") then
        ipkumdivName = "���� �Ϸ�"
    elseif (ipkumdiv="5") then
        ipkumdivName = "��ǰ �غ�"
    elseif (ipkumdiv="6") then
        ipkumdivName = "��� �غ�"
    elseif (ipkumdiv="7") then
        ipkumdivName = "�Ϻ� ���"
    elseif (ipkumdiv="8") then
        ipkumdivName = "��� �Ϸ�"
    end if

    if (accountdiv="7") then
        accountdivName = "������"
    elseif (accountdiv="14") then
        accountdivName = "����������"
    elseif (accountdiv="100") then
        accountdivName = "�ſ�ī��"
    elseif (accountdiv="550") then
        accountdivName = "������"
    elseif (accountdiv="560") then
        accountdivName = "����Ƽ��"
    elseif (accountdiv="80") then
        accountdivName = "�ÿ�ī��"
    elseif (accountdiv="50") then
        accountdivName = "���޸�����"
    elseif (accountdiv="20") then
        accountdivName = "�ǽð���ü"
    elseif (accountdiv="150") then
        accountdivName = "�̴Ϸ�Ż"
    end if

    ''2016/08/04
    if (pggubun="NP") then
        accountdivName = "���̹�����"
        if (comm_cd="A007") then
            comm_name = "���̹����� ��ҿ�û"
        end if
    end if

    ''��Ҹ�..
    if (divcd="A007") or (divcd="A008") then
        GetDefaultTitle = accountdivName + " " + ipkumdivName + " ���� �� " + comm_name
    else
        GetDefaultTitle = comm_name
    end if
end function

function AddCsMemoWithMemoGubun(orderserial,divcd,userid,writeuser,contents_jupsu,mmgubun)
	dim sqlStr

	if divcd="1" then
        ''�Ϲݸ޸�
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate())"

        dbget.Execute sqlStr
    else
        ''ó����û�޸�
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

        dbget.Execute sqlStr
    end if
end function

function AddCsMemo(orderserial,divcd,userid,writeuser,contents_jupsu)
    dim sqlStr
    dim mmgubun ''�޸𱸺�
	dim phoneNumber, startPhoneIdx, endPhoneIdx
    if (LCase(LEFT(contents_jupsu,5))="[sms ") then
    	mmgubun = "4"
		startPhoneIdx = Len("[sms ") + 1
		endPhoneIdx = InStr(contents_jupsu, "]")
		if (endPhoneIdx > 0) and ((endPhoneIdx - startPhoneIdx) < 16) then
			phoneNumber = Mid(contents_jupsu, startPhoneIdx, (endPhoneIdx - startPhoneIdx))
		end if
    elseif (LCase(LEFT(contents_jupsu,5))="[mail") then
    	mmgubun = "5"
    else
    	mmgubun = "0"
    end if

    if divcd="1" then
        ''�Ϲݸ޸�
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate, phoneNumber)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate(), '" + CStr(phoneNumber) + "')"

        dbget.Execute sqlStr
    else
        ''ó����û�޸�
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

        dbget.Execute sqlStr
    end if

end function

function AddCsMemoRequest(orderserial, userid, qadiv, writeuser, contents_jupsu)
    dim sqlStr

    ''ó����û�޸�
	sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
    sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,qadiv,writeuser,contents_jupsu,finishyn)"
    sqlStr = sqlStr + " values('" + orderserial + "','2','" + userid + "','0','" + qadiv + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"
    dbget.Execute sqlStr

end function

function SetCustomerOpenMsg(id, opentitle, opencontents)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"        + VbCrlf
    sqlStr = sqlStr + " set opentitle='" + opentitle + "'"  + VbCrlf
    sqlStr = sqlStr + " , opencontents='" + opencontents + "'" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function

function SetDetailCurrState(orderserial)
    dim sqlStr

    sqlStr = " update d "
    sqlStr = sqlStr + " Set d.currstate = 2 "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
    sqlStr = sqlStr + " 	and m.ipkumdiv >= '4' "
    sqlStr = sqlStr + " 	and m.jumundiv <> '6' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '4' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '7' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
    sqlStr = sqlStr + " 	and d.itemid<>0 "
    sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " 	and IsNULL(d.currstate,0)=0 "

    dbget.Execute sqlStr

end function

'function AddCustomerOpenMsg(id, orderserial, addcontents)
'    dim sqlStr
'
'    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"        + VbCrlf
'    sqlStr = sqlStr + " set opentitle=opentitle + '" + VbCrlf + addcontents + "'" + VbCrlf
'    sqlStr = sqlStr + " where id=" + CStr(id)
'
'    dbget.Execute sqlStr
'
'end function

function AddCustomerOpenContents(id, ByVal addcontents)
    dim sqlStr

    if ((addcontents="") or (id="")) then Exit Function

    '// SQL ������??
    addcontents = Replace(addcontents, "--", "")

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"        + VbCrlf
    sqlStr = sqlStr + " set opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '" & addcontents & "' else '" & VbCrlf & addcontents + "' End )" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function

function RegCSMasterAddUpcheByAsid(id)
    dim sqlStr

    sqlStr = " update a " + vbCrLf
    sqlStr = sqlStr + " set a.requireupche='Y', a.makerid = T.makerid " + vbCrLf
    sqlStr = sqlStr + " from " + vbCrLf
    sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a " + vbCrLf
    sqlStr = sqlStr + " 	join ( " + vbCrLf
    sqlStr = sqlStr + " 		select a.id as asid, max(d.makerid) as makerid, count(distinct d.makerid) as makerCnt, sum(case when d.isupchebeasong = 'N' then 1 else 0 end) as tenbaeCnt " + vbCrLf
    sqlStr = sqlStr + " 		from " + vbCrLf
    sqlStr = sqlStr + " 			[db_cs].[dbo].[tbl_new_as_list] a " + vbCrLf
    sqlStr = sqlStr + " 			join [db_cs].[dbo].[tbl_new_as_detail] d on a.id = d.masterid " + vbCrLf
    sqlStr = sqlStr + " 		where " + vbCrLf
    sqlStr = sqlStr + " 			1 = 1 " + vbCrLf
    sqlStr = sqlStr + " 			and a.id = " & id & vbCrLf
    sqlStr = sqlStr + " 			and d.itemid <> 0 " + vbCrLf
    sqlStr = sqlStr + " 		group by " + vbCrLf
    sqlStr = sqlStr + " 			a.id " + vbCrLf
    sqlStr = sqlStr + " 	) T on a.id = T.asid " + vbCrLf
    sqlStr = sqlStr + " where " + vbCrLf
    sqlStr = sqlStr + " 	1 = 1 " + vbCrLf
    sqlStr = sqlStr + " 	and a.id = " & id & vbCrLf
    sqlStr = sqlStr + " 	and T.makerCnt = 1 " + vbCrLf
    sqlStr = sqlStr + " 	and T.tenbaeCnt = 0 " + vbCrLf

    dbget.Execute sqlStr
end function

function RegCSMasterAddUpche(id, imakerid)
    dim sqlStr
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"    + VbCrlf
    sqlStr = sqlStr + " set makerid='" + imakerid + "'"   + VbCrlf
    sqlStr = sqlStr + " , requireupche='Y'"               + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr
end function

function RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    '' CS Master ����
    dim sqlStr, InsertedId

    sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
    rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
        rsget("divcd")          = divcd
    	rsget("orderserial")    = orderserial
    	rsget("customername")   = ""
    	rsget("userid")         = ""
    	rsget("writeuser")      = reguserid
    	rsget("title")          = title
    	rsget("contents_jupsu") = contents_jupsu
    	rsget("gubun01")        = gubun01
    	rsget("gubun02")        = gubun02

    	rsget("currstate")      = "B001"
    	rsget("deleteyn")       = "N"

        ''''''''''''''''''''''''''''''''''
    	''rsget("requireupche")   = "N"
    	''rsget("makerid")        = ""
    	''''''''''''''''''''''''''''''''''

    rsget.update
	    InsertedId = rsget("id")
	rsget.close

	dim opentitle, opencontents
	dim IsUpdateSuccess

	opentitle = GetDefaultTitle(divcd, InsertedId, orderserial)

	sqlStr = " update [db_cs].[dbo].tbl_new_as_list"  + VbCrlf
	sqlStr = sqlStr + " set userid=T.userid"        + VbCrlf
	sqlStr = sqlStr + " , customername=T.buyname"   + VbCrlf
	sqlStr = sqlStr + " , opentitle='" + html2db(opentitle) + "'" + VbCrlf
	sqlStr = sqlStr + " , opencontents='" + html2db(opencontents) + "'" + VbCrlf
	sqlStr = sqlStr + " , extsitename=(CASE WHEN T.sitename<>'10x10' THEN T.sitename ELSE NULL END)"   + VbCrlf   ''2011-06-14 �߰�

	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 T" + VbCrlf
	else
    	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master T" + VbCrlf
    end if

	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"  + VbCrlf
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.id=" + CStr(InsertedId)
	dbget.Execute sqlStr

	IsUpdateSuccess = False
	sqlStr = " select @@rowcount as cnt "
	'response.write sqlStr

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IsUpdateSuccess = (rsget("cnt") > 0)
    rsget.Close

	''ȸ����û �����ΰ�� - �⺻ ȸ�� ����� ����
	''�±�ȯ, ���� �߼�, �����߼�, ��Ÿȸ��
	if (divcd="A010") or (divcd="A010") or (divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002") or (divcd="A200") then
	    Call RegDefaultDEliverInfo(InsertedId, orderserial)
    end if

	if (Not IsNumeric(orderserial)) and (IsUpdateSuccess = False) then
		'Giftī�� �ֹ����� Ȯ���Ѵ�
		sqlStr = " update [db_cs].[dbo].tbl_new_as_list"  + VbCrlf
		sqlStr = sqlStr + " set userid=T.userid"        + VbCrlf
		sqlStr = sqlStr + " , customername=T.buyname"   + VbCrlf
		sqlStr = sqlStr + " , opentitle='" + title + "'" + VbCrlf
		sqlStr = sqlStr + " , opencontents=''" + VbCrlf
		sqlStr = sqlStr + " , extsitename='giftcard' "   + VbCrlf
    	sqlStr = sqlStr + " from [db_order].[dbo].tbl_giftcard_order T" + VbCrlf
		sqlStr = sqlStr + " where T.giftorderserial='" + orderserial + "'"  + VbCrlf
		sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_list.id=" + CStr(InsertedId)
		dbget.Execute sqlStr

		sqlStr = " select @@rowcount as cnt "
		'response.write sqlStr

	    rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	        IsUpdateSuccess = (rsget("cnt") > 0)
	    rsget.Close
	end if

	RegCSMaster = InsertedId
end function

''�⺻ ȸ��/�±�ȯ/���񽺹߼� �ּ��� �Է� - ������ �ֹ���ȣ �⺻ �ּ����� �����. - ������ �����ϴ� Procsess
function RegDefaultDEliverInfo(AsID, orderserial)
    dim sqlStr
    sqlStr = "insert into [db_cs].[dbo].tbl_new_as_delivery"
    sqlStr = sqlStr + " (asid, reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqetcaddr)"
    ''sqlStr = sqlStr + " select " + CStr(AsID) + ",reqname, reqphone, reqhp, reqzipcode, reqaddress, reqzipaddr" ''�ٲ���.
    sqlStr = sqlStr + " select " + CStr(AsID) + ",reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqaddress"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 T" + VbCrlf
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    dbget.Execute sqlStr
end function

function EditCSMaster(AsID, modiuserid, title, contents_jupsu, gubun01, gubun02)
    '' CS Master ����
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " set writeuser='" + modiuserid + "'"
    sqlStr = sqlStr + " ,title='" + title + "'"
    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
    sqlStr = sqlStr + " ,gubun01='" + gubun01 + "'"
    sqlStr = sqlStr + " ,gubun02='" + gubun02 + "'"
    sqlStr = sqlStr + " where id=" + CStr(AsID)

    dbget.Execute sqlStr

end function

function GetCSRefundTitle(AsID, divcd, orderserial, returnmethod, orgtitle)
    '' CS ȯ�� ����
	'// GetIsCSServiceRefund() �� �Բ� �ٲܰ�
    dim tmptitle

	GetCSRefundTitle = orgtitle

    if (divcd <> "A003") and (divcd <> "A007") then
    	Exit function
    end if

    if (InStr("�ֹ� ��� ��,��ǰ ó�� ��,ȸ�� ó�� ��", Left(orgtitle, 7)) > 0) then
		GetCSRefundTitle = Left(orgtitle, 7) & " " & GetRefundMethodString(returnmethod)
    end if

	if (Left(orgtitle, 7) = "CS���� -") then
		GetCSRefundTitle = Left(orgtitle, 7) & " " & GetRefundMethodString(returnmethod)
	end if

    if ((orgtitle = "���ϸ��� ����(CS����)") or (orgtitle = "ȯ��(������)") or (orgtitle = "ȯ��(��ġ��)") or (orgtitle = "ȯ��(���ϸ���)") or (orgtitle = "��ġ�� ����(ǰ��)") or (orgtitle = "���ϸ��� ����(ǰ��)") or (orgtitle = "���ϸ��� ����(�������)")) then
    	GetCSRefundTitle = "CS���� -" & " " & GetRefundMethodString(returnmethod)
    end if

end function

function GetIsCSServiceRefund(AsID, divcd, orgtitle)
	'// GetCSRefundTitle() �� �Բ� �ٲܰ�
	GetIsCSServiceRefund = False

    if (divcd <> "A003") and (divcd <> "A007") then
    	Exit function
    end if

    if (InStr("�ֹ� ��� ��,��ǰ ó�� ��,ȸ�� ó�� ��", Left(orgtitle, 7)) > 0) then
		Exit function
    end if

	if (Left(orgtitle, 7) = "CS���� -") then
		GetIsCSServiceRefund = True
		Exit function
	end if

    if ((orgtitle = "���ϸ��� ����(CS����)") or (orgtitle = "ȯ��(������)") or (orgtitle = "ȯ��(��ġ��)") or (orgtitle = "ȯ��(���ϸ���)") or (orgtitle = "��ġ�� ����(ǰ��)") or (orgtitle = "���ϸ��� ����(ǰ��)") or (orgtitle = "���ϸ��� ����(�������)")) then
		GetIsCSServiceRefund = True
		Exit function
    end if

end function

function SetCSServiceRefund(AsID)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " set isCSServiceRefund='Y' "
    sqlStr = sqlStr + " where asid=" + CStr(AsID)

    dbget.Execute sqlStr
end function

function GetRefundMethodString(returnmethod)
	dim tmpstr

    'R007 ������ȯ��
    'R020 �ǽð���ü���
    'R050 ���������� ���
    'R080 �ÿ�ī�����
    'R100 �ſ�ī�����
    'R550 ���������
    'R560 ����Ƽ�����
    'R120 �ſ�ī��κ����
    'R400 �޴������
	'R420 �޴����κ����
    'R900 ���ϸ�����ȯ��
    'R910 ��ġ��ȯ��
    'R022 �ǽð���ü�κ����(NP)
    'R150 �̴Ϸ�Ż���

	tmpstr = ""

    if (returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420") or (returnmethod="R150") then
        if (returnmethod="R020") then
            tmpstr = "�ǽð���ü���"
        elseif (returnmethod="R022") then ''2016/07/21
            tmpstr = "�ǽð���ü�κ����"
        elseif (returnmethod="R080") then
            tmpstr = "�ÿ�ī�����"
        elseif (returnmethod="R100") then
            tmpstr = "�ſ�ī�����"
        elseif (returnmethod="R550") then
            tmpstr = "���������"
        elseif (returnmethod="R560") then
            tmpstr = "����Ƽ�����"
        elseif (returnmethod="R120") then
            tmpstr = "�ſ�ī��κ����"
		elseif (returnmethod="R400") then
            tmpstr = "�޴������"
        elseif (returnmethod="R420") then
            tmpstr = "�޴����κ����"
        elseif (returnmethod="R150") then
            tmpstr = "�̴Ϸ�Ż���"
        end if
    elseif (returnmethod="R050") then
        tmpstr = "���������� ���"
    elseif (returnmethod="R900") then
        tmpstr = "���ϸ��� ȯ��"
    elseif (returnmethod="R910") then
        tmpstr = "��ġ�� ȯ��"
    elseif (returnmethod<>"") then
        tmpstr = "������ ȯ��"
    end if

	GetRefundMethodString = tmpstr

end function

function EditCSMasterFinished(AsID, title, contents_jupsu, gubun01, gubun02, finishuserid, contents_finish)
    '' CS Master �Ϸ�� ���� ����
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " set finishuser='" + finishuserid + "'"
    sqlStr = sqlStr + " ,title='" + title + "'"
    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
    sqlStr = sqlStr + " ,contents_finish='" + contents_finish + "'"
    sqlStr = sqlStr + " ,gubun01='" + gubun01 + "'"
    sqlStr = sqlStr + " ,gubun02='" + gubun02 + "'"
    sqlStr = sqlStr + " where id=" + CStr(AsID)

    dbget.Execute sqlStr
end Function

' Arr = Array( _
' 	Array("needChkYN","F"), _
' 	Array("2","��"), _
' 	Array("3","��"), _
' 	Array("4","��") )
function EditCSMasterAddInfo(AsID, addInfoArr)
    '' CS Master �߰����� �Է�
    dim sqlStr, key, updateAvail
	updateAvail = False

	If (Not IsArray(addInfoArr)) Then
		Exit function
	End If

	If (UBound(addInfoArr) < 0) Then
		Exit function
	End If

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " set regdate = regdate"

	For Each key In addInfoArr
		If (key(1) <> "") Then
			sqlStr = sqlStr + " ," & key(0) & "='" + key(1) + "'"
			updateAvail = True
		End If
	Next

	sqlStr = sqlStr + " where id=" + CStr(AsID)

	If (updateAvail = True) Then
		dbget.Execute sqlStr
	End If
end function

function CopyWebCancelRefundInfo(FromCsId, ToCsId)
    dim sqlStr
    ''��ü ��� ȯ������ ����

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " 	, returnmethod"
    sqlStr = sqlStr + " 	, refundrequire"
    sqlStr = sqlStr + " 	, orgsubtotalprice"
    sqlStr = sqlStr + " 	, orgitemcostsum"
    sqlStr = sqlStr + " 	, orgbeasongpay"
    sqlStr = sqlStr + " 	, orgmileagesum"
    sqlStr = sqlStr + " 	, orgcouponsum"
    sqlStr = sqlStr + " 	, orgallatdiscountsum"

    ''��� ��������
    sqlStr = sqlStr + " 	, canceltotal"
    sqlStr = sqlStr + " 	, refunditemcostsum"
    sqlStr = sqlStr + " 	, refundmileagesum"
    sqlStr = sqlStr + " 	, refundcouponsum"
    sqlStr = sqlStr + " 	, allatsubtractsum"
    sqlStr = sqlStr + " 	, refundbeasongpay"
    sqlStr = sqlStr + " 	, refunddeliverypay"
    sqlStr = sqlStr + " 	, refundadjustpay"
    sqlStr = sqlStr + " 	, rebankname"
    sqlStr = sqlStr + " 	, rebankaccount"
    sqlStr = sqlStr + " 	, rebankownername"
    sqlStr = sqlStr + " 	, encmethod"
    sqlStr = sqlStr + " 	, encaccount"

    sqlStr = sqlStr + " 	, paygateTid"
    sqlStr = sqlStr + " 	, orggiftcardsum"
    sqlStr = sqlStr + " 	, orgdepositsum"
    sqlStr = sqlStr + " 	, refundgiftcardsum"
    sqlStr = sqlStr + " 	, refunddepositsum"
    sqlStr = sqlStr + " )"

    sqlStr = sqlStr + " select " + CStr(ToCsId)
    sqlStr = sqlStr + " 	, returnmethod"
    sqlStr = sqlStr + " 	, refundrequire"
    sqlStr = sqlStr + " 	, orgsubtotalprice"
    sqlStr = sqlStr + " 	, orgitemcostsum"
    sqlStr = sqlStr + " 	, orgbeasongpay"
    sqlStr = sqlStr + " 	, orgmileagesum"
    sqlStr = sqlStr + " 	, orgcouponsum"
    sqlStr = sqlStr + " 	, orgallatdiscountsum"

    ''��� ��������
    sqlStr = sqlStr + " 	, canceltotal"
    sqlStr = sqlStr + " 	, refunditemcostsum"
    sqlStr = sqlStr + " 	, refundmileagesum"
    sqlStr = sqlStr + " 	, refundcouponsum"
    sqlStr = sqlStr + " 	, allatsubtractsum"
    sqlStr = sqlStr + " 	, refundbeasongpay"
    sqlStr = sqlStr + " 	, refunddeliverypay"
    sqlStr = sqlStr + " 	, refundadjustpay"
    sqlStr = sqlStr + " 	, rebankname"
    sqlStr = sqlStr + " 	, rebankaccount"
    sqlStr = sqlStr + " 	, rebankownername"
    sqlStr = sqlStr + " 	, encmethod"
    sqlStr = sqlStr + " 	, encaccount"

    sqlStr = sqlStr + " 	, paygateTid"
    sqlStr = sqlStr + " 	, orggiftcardsum"
    sqlStr = sqlStr + " 	, orgdepositsum"
    sqlStr = sqlStr + " 	, refundgiftcardsum"
    sqlStr = sqlStr + " 	, refunddepositsum"
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr + " where asid = " & FromCsId & " "

    dbget.Execute sqlStr

    '���� CS REF KEY*******************************
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
    sqlStr = sqlStr + " set refasid = " & CStr(FromCsId) & " "
    sqlStr = sqlStr + " where id = " & CStr(ToCsId) & " "
    dbget.Execute sqlStr

end function

function RegCSMasterRefundInfo(asid, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay  , rebankname, rebankaccount, rebankownername, paygateTid)

    dim sqlStr
    if IsNULL(orgmileagesum) then orgmileagesum=0

    sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"

    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"
    sqlStr = sqlStr + " )"

	'response.write "aaaaaaaaaaa" & sqlStr

    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(asid)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(orgsubtotalprice)
    sqlStr = sqlStr + " ," + CStr(orgitemcostsum)
    sqlStr = sqlStr + " ," + CStr(orgbeasongpay)
    sqlStr = sqlStr + " ," + CStr(orgmileagesum)
    sqlStr = sqlStr + " ," + CStr(orgcouponsum)
    sqlStr = sqlStr + " ," + CStr(orgallatdiscountsum)

	'response.write "aaaaaaaaaaa" & sqlStr

    sqlStr = sqlStr + " ," + CStr(canceltotal)
    sqlStr = sqlStr + " ," + CStr(refunditemcostsum)
    sqlStr = sqlStr + " ," + CStr(refundmileagesum)
    sqlStr = sqlStr + " ," + CStr(refundcouponsum)
    sqlStr = sqlStr + " ," + CStr(allatsubtractsum)
    sqlStr = sqlStr + " ," + CStr(refundbeasongpay)
    sqlStr = sqlStr + " ," + CStr(refunddeliverypay)
    sqlStr = sqlStr + " ," + CStr(refundadjustpay)
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,'" + paygateTid + "'"
    sqlStr = sqlStr + " )"

	'response.write "aaaaaaaaaaa" & sqlStr
    dbget.Execute sqlStr
end function

function AddCSMasterRefundInfo(asid, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

    dim sqlStr

    sqlStr = " update "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	orggiftcardsum = " & CStr(orggiftcardsum) & " "
    sqlStr = sqlStr + " 	, orgdepositsum = " & CStr(orgdepositsum) & " "
    sqlStr = sqlStr + " 	, refundgiftcardsum = " & CStr(refundgiftcardsum) & " "
    sqlStr = sqlStr + " 	, refunddepositsum = " & CStr(refunddepositsum) & " "
    sqlStr = sqlStr + " where asid = " & CStr(asid) & " "

	'response.write "aaaaaaaaaaa" & sqlStr
    dbget.Execute sqlStr

end function

function EditCSMasterRefundEncInfo(asid, encmethod, bnkaccount)
    dim sqlStr

    ''2017/10/02 ��ȣȭ ��� ����
    sqlStr = "exec db_cs.[dbo].[sp_Ten_EditCSMasterRefundEncInfo] "&CStr(asid)&",'"&encmethod&"','"&bnkaccount&"'"
    dbget.Execute sqlStr
    exit function

    IF (encmethod="PH1") then
        IF (bnkaccount="") then
            sqlStr = " update [db_cs].[dbo].tbl_as_refund_info " & VbCRLF
            sqlStr = sqlStr + " set encmethod = '' " & VbCRLF
            sqlStr = sqlStr + " 	, encaccount = NULL" & VbCRLF
            sqlStr = sqlStr + " 	, rebankaccount=''" & VbCRLF
            sqlStr = sqlStr + " where asid = " & CStr(asid) & " " & VbCRLF

            dbget.Execute sqlStr
        ELSE
            sqlStr = " update [db_cs].[dbo].tbl_as_refund_info " & VbCRLF
            sqlStr = sqlStr + " set encmethod = '" & Left(CStr(encmethod), 8) & "' " & VbCRLF
            sqlStr = sqlStr + " 	, encaccount = db_cs.dbo.uf_EncAcctPH1('"&bnkaccount&"')" & VbCRLF
            sqlStr = sqlStr + " 	, rebankaccount=''" & VbCRLF
            sqlStr = sqlStr + " where asid = " & CStr(asid) & " " & VbCRLF

            dbget.Execute sqlStr
        END IF
    end IF

end function

function RegCSUpcheAddJungsanPay(iasid, iadd_upchejungsandeliverypay, iadd_upchejungsancause, buf_requiremakerid)
    dim sqlStr

    sqlStr = " insert into [db_cs].[dbo].tbl_as_upcheAddjungsan"
    sqlStr = sqlStr + " (asid, add_upchejungsandeliverypay, add_upchejungsancause)"
    sqlStr = sqlStr + " values(" &iasid
    sqlStr = sqlStr + " ," & iadd_upchejungsandeliverypay
    sqlStr = sqlStr + " ,'" & iadd_upchejungsancause & "')"

    dbget.Execute sqlStr

    ''��Ÿ ���� �߰��ΰ�츸 makerid ���� : ��ǰ����(��ü���) / �±�ȯ(��ü)�� ���� �� ������.
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" & VbCrlf
    sqlStr = sqlStr + " set makerid='" & buf_requiremakerid & "'" & VbCrlf
    sqlStr = sqlStr + " where id=" & iasid & "" & VbCrlf
    sqlStr = sqlStr + " and divcd='A700'" & VbCrlf

    dbget.Execute sqlStr

end function

function setRestoreEtcRealPayment(asid, orderserial)
    dim sqlStr

    '// TODO : ���� �ſ�ī��ݾ׸� ����
    sqlStr = " update e "
    sqlStr = sqlStr + " set e.realPayedsum = e.realPayedsum + r.refundresult "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_as_refund_info] r on a.id = r.asid "
    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_PaymentEtc e on a.orderserial = e.orderserial and e.acctdiv = '100' "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and a.id = " & asid
    sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
    sqlStr = sqlStr + " 	and a.currstate = 'B007' "
    dbget.Execute sqlStr
end function

function EditCSUpcheAddJungsanPay(iasid, iadd_upchejungsandeliverypay, iadd_upchejungsancause, buf_requiremakerid)
    dim sqlStr

    sqlStr = " IF EXISTS( select * from [db_cs].[dbo].tbl_as_upcheAddjungsan where asid=" & iasid & ")" & VbCrlf
    sqlStr = sqlStr + " BEGIN" & VbCrlf
    sqlStr = sqlStr + "     update [db_cs].[dbo].tbl_as_upcheAddjungsan" & VbCrlf
    sqlStr = sqlStr + "     set add_upchejungsandeliverypay=" & add_upchejungsandeliverypay & VbCrlf
    sqlStr = sqlStr + "     , add_upchejungsancause='" & iadd_upchejungsancause & "'" & VbCrlf
    sqlStr = sqlStr + "     where asid = " & iasid & VbCrlf
    sqlStr = sqlStr + " END" & VbCrlf
    sqlStr = sqlStr + " ELSE " & VbCrlf
    sqlStr = sqlStr + " BEGIN" & VbCrlf
    sqlStr = sqlStr + "     IF (0<>" & iadd_upchejungsandeliverypay & ")" & VbCrlf
    sqlStr = sqlStr + "     BEGIN" & VbCrlf
    sqlStr = sqlStr + "         insert into [db_cs].[dbo].tbl_as_upcheAddjungsan" & VbCrlf
    sqlStr = sqlStr + "         (asid, add_upchejungsandeliverypay, add_upchejungsancause)" & VbCrlf
    sqlStr = sqlStr + "         values(" &iasid & VbCrlf
    sqlStr = sqlStr + "         ," & iadd_upchejungsandeliverypay & VbCrlf
    sqlStr = sqlStr + "         ,'" & iadd_upchejungsancause & "')" & VbCrlf
    sqlStr = sqlStr + "     END" & VbCrlf
    sqlStr = sqlStr + " END" & VbCrlf

    dbget.Execute sqlStr

    ''��Ÿ ���� �߰��ΰ�츸 makerid ���� : ��ǰ����(��ü���) / �±�ȯ(��ü)�� ���� �� ������.
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" & VbCrlf
    sqlStr = sqlStr + " set makerid='" & buf_requiremakerid & "'" & VbCrlf
    sqlStr = sqlStr + " where id=" & iasid & "" & VbCrlf
    sqlStr = sqlStr + " and divcd='A700'" & VbCrlf
    sqlStr = sqlStr + " and IsNULL(makerid,'')<>'" & buf_requiremakerid & "'" & VbCrlf

    dbget.Execute sqlStr
end function

function DeleteAllCSDetail(id, orderserial)
	dim sqlStr

	sqlStr = " delete from [db_cs].[dbo].tbl_new_as_detail "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and masterid = " + CStr(id) + " "
	dbget.Execute sqlStr

end function

function AddCSDetailByArrStr(byval detailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno

	dim CURR_IsOLDOrder : CURR_IsOLDOrder = False

	sqlStr = " select top 1 orderserial from [db_log].[dbo].tbl_old_order_master_2003 where orderserial = '" + CStr(orderserial) + "' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		CURR_IsOLDOrder = True
	end if
	rsget.Close

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)

	        call AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
		end if
	next

	sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
	sqlStr = sqlStr + " set itemid=T.itemid"
	sqlStr = sqlStr + " , itemoption=T.itemoption"
	sqlStr = sqlStr + " , makerid=T.makerid"
	sqlStr = sqlStr + " , itemname=T.itemname"
	sqlStr = sqlStr + " , itemoptionname=T.itemoptionname"
	sqlStr = sqlStr + " , itemcost=T.itemcost"
	sqlStr = sqlStr + " , buycash=T.buycash"
	sqlStr = sqlStr + " , orderitemno=(CASE WHEN T.cancelyn='Y' THEN 0 ELSE T.itemno END)"
	sqlStr = sqlStr + " , isupchebeasong=T.isupchebeasong"
	sqlStr = sqlStr + " , regdetailstate=T.currstate"
	if (CURR_IsOLDOrder) then
	    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 T"
	else
	    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail T"
	end if
	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_detail.masterid=" + CStr(id)
	sqlStr = sqlStr + " and [db_cs].[dbo].tbl_new_as_detail.orderdetailidx=T.idx"

	dbget.Execute sqlStr

end Function

function ModiCSDetailByArrStr(byval detailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)

	        call ModiOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
		end if
	next

end function

'// �� �ֹ������� ���� ��ǰ ���(��ǰ���� �±�ȯ)
function AddCSDetailWithoutOrderDetailByArrStr(byval newdetailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno, dregitemid, dregitemoption

	dim CURR_IsOLDOrder : CURR_IsOLDOrder = False

	sqlStr = " select top 1 orderserial from [db_log].[dbo].tbl_old_order_master_2003 where orderserial = '" + CStr(orderserial) + "' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		CURR_IsOLDOrder = True
	end if
	rsget.Close

    buf = split(newdetailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)
			dregitemid      = tmp(4)
			dregitemoption  = tmp(5)

	        Call AddOneCSDetailWithoutOrderDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemid, dregitemoption, dregitemno)
		end if
	next

	'// ��ǰ���� ����
    sqlStr = "update D"
    sqlStr = sqlStr & " set makerid=i.makerid"
    sqlStr = sqlStr & " , itemname=i.itemname"
    sqlStr = sqlStr & " , itemoptionname=IsNULL(o.optionname,'')"
    sqlStr = sqlStr & " , regdetailstate=2"
    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail D"
    sqlStr = sqlStr & "     Join db_item.dbo.tbl_item i"
    sqlStr = sqlStr & "     on D.itemid=i.itemid"
    sqlStr = sqlStr & "     left join db_item.dbo.tbl_item_option o"
    sqlStr = sqlStr & "     on D.itemid=o.itemid"
    sqlStr = sqlStr & "     and D.itemOption=o.itemOption"
    sqlStr = sqlStr & " where D.masterid="&id
    sqlStr = sqlStr & " and D.reforderdetailidx is not null"
    dbget.Execute sqlStr

	''��������
	''�ɼǺ��� : ���� �ɼǰ� �����ؾ߸� �±�ȯ ����, �������� �״�� ī��
	''��ǰ���� : �귣��,���� �ǸŰ�(���ΰ�),���԰�, �������밡�� �� ��� �����ؾ��ϰ� 1:1 ���游 ����(������ ������ ����), �������� �״�� ī��
    sqlStr = "update D"
    sqlStr = sqlStr & " set itemcost=T.itemcost"
    sqlStr = sqlStr & " , buycash=T.buycash"
    sqlStr = sqlStr & " , isupchebeasong=T.isupchebeasong"
    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail D"
	if (CURR_IsOLDOrder) then
	    sqlStr = sqlStr + " join [db_log].[dbo].tbl_old_order_detail_2003 T"
	else
	    sqlStr = sqlStr + " join [db_order].[dbo].tbl_order_detail T"
	end if
	sqlStr = sqlStr & "     on D.reforderdetailidx=T.idx"
    sqlStr = sqlStr & " where D.masterid="&id
    sqlStr = sqlStr & " and D.reforderdetailidx is not null"
	dbget.Execute sqlStr

end Function

function ModiCSDetailWithoutOrderDetailByArrStr(byval newdetailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno, dregitemid, dregitemoption

    buf = split(newdetailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)
			dregitemid      = tmp(4)
			dregitemoption  = tmp(5)

			Call ModiOneCSDetailWithoutOrderDetail(id,dgubun01, dgubun02, dregitemid, dregitemoption, dregitemno)
		end if
	next

end function

function ModiCSDetailAddedItem(id, toItemId, toItemOption, SalePrice, ItemCouponPrice, BonusCouponPrice, buycash)
	dim sqlStr, tmp, buf, i

	sqlStr = "update D"
    sqlStr = sqlStr & " set itemcost=" & BonusCouponPrice
    sqlStr = sqlStr & " , buycash=" & buycash
	sqlStr = sqlStr & " , SalePrice=" & SalePrice
	sqlStr = sqlStr & " , ItemCouponPrice=" & ItemCouponPrice
    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail D"
    sqlStr = sqlStr & " where D.masterid="&id
    sqlStr = sqlStr & " and D.itemid = " & toItemId & " and D.itemoption = '" & toItemOption & "' "
	dbget.Execute sqlStr
End Function

'''���̳ʽ� �ֹ����� ���� �����ϴ°��� �´°���..
function ChangeReturnItems(orderserial, id, newasid, pitemid, pitemoption, citemid, citemoption, citemno ,byRef ScanErr)
    ''id : �±�ȯ ���id, newasid : �±�ȯ ȸ��id, pitemid :����ǰID citemid : �����ǰID

	response.write detailitemlist + "���� : �ý����� ����"
	response.end

    dim sqlStr
    dim buf_pitemid, buf_pitemoption, buf_citemid, buf_citemoption, buf_citemno, i
    dim citemExists : citemExists = false
    dim curritemcost, curritemcostsum, regedSum
    dim MinusOrderSerial

    curritemcost = 0
    curritemcostsum = 0
    regedSum = 0

    buf_pitemid = split(pitemid,",")
    buf_pitemoption = split(pitemoption,",")
    buf_citemid = split(citemid,",")
    buf_citemoption = split(citemoption,",")
    buf_citemno = split(citemno,",")

    ''���� // ������ ���� �հ�ݾװ� ��ġ �ؾ���.
    for i = 0 to UBound(buf_citemid)
		if (TRIM(buf_citemid(i)) <> "") then  ''����� ��ǰ��ȣ.
		    citemExists = true
	        sqlStr = "select (i.sellcash + IsNULL(o.optaddprice,0)) as curritemcost"
            sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
            sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item_option o"
            sqlStr = sqlStr & " 	on i.itemid=o.itemid"
            sqlStr = sqlStr & " 	and o.itemoption='"&Trim(buf_pitemoption(i))&"'"
            sqlStr = sqlStr & " where i.itemid="&Trim(buf_pitemid(i))

		    rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		    if (Not rsget.Eof) then
                curritemcost = rsget("curritemcost")
            end if
            rsget.close

            curritemcostsum = curritemcostsum + curritemcost*Trim(buf_citemno(i))
		end if
    next

    if (Not citemExists) then Exit Function

    sqlStr = " select sum(od.itemcostcouponNotApplied*regitemno) as regedSum"
    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail d"
    sqlStr = sqlStr & "     Join db_order.dbo.tbl_order_detail od"
    sqlStr = sqlStr & "     on d.orderdetailidx=od.idx"
    sqlStr = sqlStr & " where d.masterid="&id

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        regedSum = rsget("regedSum")
    end if
    rsget.close

    IF (curritemcostsum<>regedSum) then
        ScanErr = "��ȯ�ݾ��� ��ġ���� �ʽ��ϴ�.\n�����ݾ� :" &regedSum & "\n��ȯ�ݾ� :"&curritemcostsum
        Exit Function
    End IF

    ''''////MinusOrderSerial = AddMinusOrder(id,orderserial)

    ''�±�ȯ ����.
    for i = 0 to UBound(buf_citemid)
		if (TRIM(buf_citemid(i)) <> "") then
'		    sqlStr = "update [db_cs].[dbo].tbl_new_as_detail"
'		    sqlStr = sqlStr & " set itemid="&TRIM(buf_citemid(i))
'		    sqlStr = sqlStr & " ,itemoption='"&TRIM(buf_citemid(i))&"'"
'		    sqlStr = sqlStr & " where masterid="&id
'		    sqlStr = sqlStr & " and itemid="&Trim(buf_pitemid(i))
'		    sqlStr = sqlStr & " and itemoption='"&Trim(buf_pitemoption(i))&"'"

            sqlStr = " Insert Into [db_cs].[dbo].tbl_new_as_detail"
            sqlStr = sqlStr & " (masterid,orderdetailidx,gubun01,gubun02,orderserial,itemid,itemoption,makerid"
            sqlStr = sqlStr & " ,itemname,itemoptionname,regitemno,confirmitemno,orderitemno,itemcost,buycash"
            sqlStr = sqlStr & " ,isupchebeasong,regdetailstate,currstate,reforderdetailidx)"
            sqlStr = sqlStr & " select masterid,NULL,gubun01,gubun02,orderserial,"&TRIM(buf_citemid(i))&",'"&TRIM(buf_citemoption(i))&"',''"
            sqlStr = sqlStr & " ,'','',regitemno,confirmitemno,orderitemno,itemcost,buycash"
            sqlStr = sqlStr & " ,isupchebeasong,2,currstate,orderdetailidx "
            sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail"
            sqlStr = sqlStr & " where masterid="&id
            sqlStr = sqlStr & " and itemid="&Trim(buf_pitemid(i))
		    sqlStr = sqlStr & " and itemoption='"&Trim(buf_pitemoption(i))&"'"

	        dbget.Execute sqlStr

		    sqlStr = "update D"
		    sqlStr = sqlStr & " set makerid=i.makerid"
		    sqlStr = sqlStr & " , itemname=i.itemname"
		    sqlStr = sqlStr & " , itemoptionname=IsNULL(o.optionname,0)"
		    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail D"
		    sqlStr = sqlStr & "     Join db_item.dbo.tbl_item i"
		    sqlStr = sqlStr & "     on D.itemid=i.itemid"
		    sqlStr = sqlStr & "     left join db_item.dbo.tbl_item_option o"
		    sqlStr = sqlStr & "     on D.itemid=o.itemid"
		    sqlStr = sqlStr & "     and D.itemOption=o.itemOption"
		    sqlStr = sqlStr & " where D.masterid="&id
            sqlStr = sqlStr & " and D.itemid="&Trim(buf_citemid(i))
		    sqlStr = sqlStr & " and D.itemoption='"&Trim(buf_citemoption(i))&"'"

		    dbget.Execute sqlStr

		    IF (newasid<>"") THEN
    		    sqlStr = " Insert Into [db_cs].[dbo].tbl_new_as_detail"
                sqlStr = sqlStr & " (masterid,orderdetailidx,gubun01,gubun02,orderserial,itemid,itemoption,makerid"
                sqlStr = sqlStr & " ,itemname,itemoptionname,regitemno,confirmitemno,orderitemno,itemcost,buycash"
                sqlStr = sqlStr & " ,isupchebeasong,regdetailstate,currstate,reforderdetailidx)"
                sqlStr = sqlStr & " select masterid,NULL,gubun01,gubun02,orderserial,"&TRIM(buf_citemid(i))&",'"&TRIM(buf_citemoption(i))&"',''"
                sqlStr = sqlStr & " ,'','',regitemno,confirmitemno,orderitemno,itemcost,buycash"
                sqlStr = sqlStr & " ,isupchebeasong,2,currstate,orderdetailidx "
                sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail"
                sqlStr = sqlStr & " where masterid="&newasid
                sqlStr = sqlStr & " and itemid="&Trim(buf_pitemid(i))
    		    sqlStr = sqlStr & " and itemoption='"&Trim(buf_pitemoption(i))&"'"

    	        dbget.Execute sqlStr

    		    sqlStr = "update D"
    		    sqlStr = sqlStr & " set makerid=i.makerid"
    		    sqlStr = sqlStr & " , itemname=i.itemname"
    		    sqlStr = sqlStr & " , itemoptionname=IsNULL(o.optionname,0)"
    		    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail D"
    		    sqlStr = sqlStr & "     Join db_item.dbo.tbl_item i"
    		    sqlStr = sqlStr & "     on D.itemid=i.itemid"
    		    sqlStr = sqlStr & "     left join db_item.dbo.tbl_item_option o"
    		    sqlStr = sqlStr & "     on D.itemid=o.itemid"
    		    sqlStr = sqlStr & "     and D.itemOption=o.itemOption"
    		    sqlStr = sqlStr & " where D.masterid="&newasid
                sqlStr = sqlStr & " and D.itemid="&Trim(buf_citemid(i))
    		    sqlStr = sqlStr & " and D.itemoption='"&Trim(buf_citemoption(i))&"'"

    		    dbget.Execute sqlStr
		    END IF
		end if
    next
end function

function EditCSDetailByArrStr(byval detailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno, dcausecontent

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)
			dcausecontent   = tmp(4)

	        call EditOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno, dcausecontent)
		end if
	next

end function

function AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
    dim sqlStr

    sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01,gubun02"
    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno) "
    sqlStr = sqlStr + " values(" + CStr(id) + ""
    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
    sqlStr = sqlStr + " ,'" + CStr(dgubun01) + "'"
    sqlStr = sqlStr + " ,'" + CStr(dgubun02) + "'"
    sqlStr = sqlStr + " ,'" + CStr(orderserial) + "'"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " )"

    dbget.Execute sqlStr
end Function

function ModiOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " set gubun01='" + dgubun01 + "'"
    sqlStr = sqlStr + " , gubun02='" + dgubun02 + "'"
    sqlStr = sqlStr + " , regitemno=" + dregitemno + ""
    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and orderdetailidx=" + CStr(dorderdetailidx)

    dbget.Execute sqlStr
end function

function AddOneCSDetailWithoutOrderDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemid, dregitemoption, dregitemno)
    dim sqlStr

    sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " (masterid, reforderdetailidx, gubun01,gubun02"
    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno) "
    sqlStr = sqlStr + " values(" + CStr(id) + ""
    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
    sqlStr = sqlStr + " ,'" + CStr(dgubun01) + "'"
    sqlStr = sqlStr + " ,'" + CStr(dgubun02) + "'"
    sqlStr = sqlStr + " ,'" + CStr(orderserial) + "'"
    sqlStr = sqlStr + " ," + CStr(dregitemid) + ""
    sqlStr = sqlStr + " ,'" + CStr(dregitemoption) + "'"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " )"

    dbget.Execute sqlStr
end Function

function ModiOneCSDetailWithoutOrderDetail(id, dgubun01, dgubun02, dregitemid, dregitemoption, dregitemno)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " set gubun01='" + dgubun01 + "'"
    sqlStr = sqlStr + " , gubun02='" + dgubun02 + "'"
    sqlStr = sqlStr + " , regitemno=" + dregitemno + ""
    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and itemid = " & dregitemid & " and itemoption = '" & dregitemoption & "' "
    dbget.Execute sqlStr
end function

function EditOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno, dcausecontent)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " set gubun01='" + dgubun01 + "'"
    sqlStr = sqlStr + " , gubun02='" + dgubun02 + "'"
    sqlStr = sqlStr + " , regitemno=" + dregitemno + ""
    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and orderdetailidx=" + CStr(dorderdetailidx)

    dbget.Execute sqlStr
end function

function AddOneDeliveryInfoCSDetail(id, gubun01, gubun02, orderserial)
    dim sqlStr

    sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail"
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01, gubun02,"
    sqlStr = sqlStr + " orderserial, itemid, itemoption, makerid,itemname, itemoptionname,"
    sqlStr = sqlStr + " regitemno, confirmitemno, orderitemno, itemcost, buycash, isupchebeasong,regdetailstate) "
    sqlStr = sqlStr + " select top 1 "
    sqlStr = sqlStr + " " + CStr(id)
    sqlStr = sqlStr + " ,d.idx"
    sqlStr = sqlStr + " ,'" + CStr(gubun01) + "'"
    sqlStr = sqlStr + " ,'" + CStr(gubun02) + "'"
    sqlStr = sqlStr + " ,d.orderserial"
    sqlStr = sqlStr + " ,d.itemid"
    sqlStr = sqlStr + " ,d.itemoption"
    sqlStr = sqlStr + " ,IsNULL(d.makerid,'')"
    sqlStr = sqlStr + " ,IsNULL(d.itemname,'��۷�')"
    sqlStr = sqlStr + " ,IsNULL(d.itemoptionname,(case when d.itemcost=0 then '������' else '�Ϲ��ù�' end))"
    sqlStr = sqlStr + " ,d.itemno"
    sqlStr = sqlStr + " ,d.itemno"
    sqlStr = sqlStr + " ,d.itemno"
    sqlStr = sqlStr + " ,d.itemcost"
    sqlStr = sqlStr + " ,d.buycash"
    sqlStr = sqlStr + " ,d.isupchebeasong"
    sqlStr = sqlStr + " ,d.currstate"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
    end if
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid=0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    dbget.Execute sqlStr

end function

''�ٷ� �Ϸ� ó���� ���� ���� ����.
function IsDirectProceedFinish(divcd, Asid, orderserial, byRef EtcStr)
    dim sqlStr
    dim cancelyn, ipkumdiv
    IsDirectProceedFinish = false

    '' currstate:2 ��ü(����) �뺸
    if (divcd="A008") then
        ''' ��� Case
        '' ��ϵ� ��ǰ�� ��ü Ȯ���� ���°� ������ �������·� ����
        sqlStr = " select count(d.idx) as invalidcount"
        sqlStr = sqlStr + " from "
        if (GC_IsOLDOrder) then
            sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
            sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
        else
            sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
            sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
        end if
        sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
        sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
        sqlStr = sqlStr + " and m.orderserial=d.orderserial"
        sqlStr = sqlStr + " and d.itemid<>0"
        sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
        sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
        sqlStr = sqlStr + " and d.currstate>=3"
        sqlStr = sqlStr + " and d.cancelyn<>'Y'"

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            IsDirectProceedFinish = (rsget("invalidcount")=0)
        rsget.close

    else

    end if

end function

function GetAllCancelRegValidResult(Asid, orderserial)
	'���� - ��ü��� ����
	'
	' - ��ü��Ұ� �´���		- rsget("itemno") = rsget("totalcancelregno") �ΰ�
	' - �����������			- d.cancelyn = 'Y' �� ���� ���
	' - �ʰ��������			- rsget("itemno") < rsget("totalcancelregno") �ΰ�
	' - ������ ��� �Ǿ�����	- m.cancelyn = 'Y' �� ���� ���

    dim sqlStr, result
    GetAllCancelRegValidResult = ""
    result = ""

	'==========================================================================
	' - ������ ��� �Ǿ�����
	'==========================================================================
	if (IsMasterCanceled(Asid, orderserial)) then
		GetAllCancelRegValidResult = "��ҵ� �ֹ��Դϴ�."
		exit function
	end if

	'==========================================================================
	'��ü��Ұ� �´��� - ������ ��� ����(CSó���Ϸ�����) ��ü�� ���� �ܿ��ֹ������� ������
	'==========================================================================
	if (Not IsAllCancelState(Asid, orderserial)) then
		if (IsErrorCancelState(Asid, orderserial)) then
			GetAllCancelRegValidResult = "�ֹ������� �ʰ��Ͽ� ���(CS���� ����)�� ��ǰ�� �ֽ��ϴ�."
		else
			GetAllCancelRegValidResult = "����ݾ� ����ȯ���̸鼭 �������(CS���� ����)�� �ƴ� ��ǰ�� �ֽ��ϴ�."
		end if
		exit function
	end if

	'==========================================================================
	'����������� - ��ҵ� �����Ͽ� ���� ��Ұ� �ִ���
	'==========================================================================
	if (IsDoubleCancelState(Asid, orderserial)) then
		GetAllCancelRegValidResult = "��ҵ� ��ǰ�� ���� ��Ұ� �ֽ��ϴ�."
		exit function
	end if

end function

function GetPartialCancelRegValidResult(Asid, orderserial)
	'���� - �Ϻ���� ����
	'
	' - �κ��������
	' - �ʰ��������
	' - �����������
	' - ������ ��� �Ǿ�����

    dim sqlStr, result
    GetPartialCancelRegValidResult = ""
    result = ""

	'==========================================================================
	' - ������ ��� �Ǿ�����
	'==========================================================================
	if (IsMasterCanceled(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "��ҵ� �ֹ��Դϴ�."
		exit function
	end if

	'==========================================================================
	'�κ�������� - ������ ��� ����(CSó���Ϸ�����) ��ü�� ���� �ܿ��ֹ��������� �������� �ִ���
	'�ʰ�������� - ������ ��� ����(CSó���Ϸ�����) ��ü�� ���� �ܿ��ֹ��������� ū���� �ִ���
	'==========================================================================
	if (IsAllCancelState(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "�ֹ���ü ���(CS���� ����)�̸鼭 ����ݾװ� ȯ�ұݾ� �հ谡 �ٸ��ϴ�. - ���ϸ���/���α� ȯ���� üũ�ϼ���."
		exit function
	else
		if (IsErrorCancelState(Asid, orderserial)) then
			GetPartialCancelRegValidResult = "�ֹ������� �ʰ��Ͽ� ���(CS���� ����)�� ��ǰ�� �ֽ��ϴ�."
			exit function
		end if
	end if

	'==========================================================================
	'����������� - ��ҵ� �����Ͽ� ���� ��Ұ� �ִ���
	'==========================================================================
	if (IsDoubleCancelState(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "��ҵ� ��ǰ�� ���� ��Ұ� �ֽ��ϴ�."
		exit function
	end if

end function

'��ü��� ��������
function IsAllCancelState(Asid, orderserial)
    dim sqlStr, result
    IsAllCancelState = true

	'==========================================================================
	'��ü��Ұ� �´��� - ������ ��� ����(CSó���Ϸ�����) ��ü�� ���� �ܿ��ֹ������� ������
	'd.cancelyn = 'Y' �� ��ǰ�� ������ҿ���, rsget("itemno") < rsget("totalcancelregno") �� ���� �ʰ���ҿ��� üũ�Ѵ�.
	'==========================================================================
    sqlStr = " select "
    sqlStr = sqlStr + "     d.itemno "
    sqlStr = sqlStr + "     , sum(IsNULL(csd.regitemno,0)) as totalcancelregno "
    sqlStr = sqlStr + " from "

    if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_master_2003 m "
	    sqlStr = sqlStr + " 	join [db_log].[dbo].tbl_old_order_detail_2003 d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    else
	    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    end if

    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list csm "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and m.orderserial = csm.orderserial "
    sqlStr = sqlStr + " 		and csm.divcd = 'A008' "
    sqlStr = sqlStr + " 		and csm.currstate <> 'B007' "
    sqlStr = sqlStr + " 		and csm.deleteyn <> 'Y' "
    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_detail csd "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and csm.id = csd.masterid "
    sqlStr = sqlStr + " 		and csd.orderdetailidx = d.idx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and d.itemid <> 0 "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " group by "
    sqlStr = sqlStr + " 	m.idx, d.idx, d.itemno "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
		do until rsget.eof
	    	if (rsget("itemno") <> rsget("totalcancelregno")) then
	    		IsAllCancelState = false
				exit do
	    	end if
			rsget.moveNext
		loop
	end if
	rsget.close

end function

'�ʰ���� ��������
function IsErrorCancelState(Asid, orderserial)
    dim sqlStr, result
    IsErrorCancelState = false

	'==========================================================================
	'�ʰ�������� - ������ ��� ����(CSó���Ϸ�����) ��ü�� ���� �ܿ��ֹ��������� ū��
	'==========================================================================
    sqlStr = " select "
    sqlStr = sqlStr + "     d.itemno "
    sqlStr = sqlStr + "     , sum(IsNULL(csd.regitemno,0)) as totalcancelregno "
    sqlStr = sqlStr + " from "

    if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_master_2003 m "
	    sqlStr = sqlStr + " 	join [db_log].[dbo].tbl_old_order_detail_2003 d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    else
	    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    end if

    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list csm "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and m.orderserial = csm.orderserial "
    sqlStr = sqlStr + " 		and csm.divcd = 'A008' "
    sqlStr = sqlStr + " 		and csm.currstate <> 'B007' "
    sqlStr = sqlStr + " 		and csm.deleteyn <> 'Y' "
    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_detail csd "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and csm.id = csd.masterid "
    sqlStr = sqlStr + " 		and csd.orderdetailidx = d.idx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and d.itemid <> 0 "
    sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " group by "
    sqlStr = sqlStr + " 	m.idx, d.idx, d.itemno "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
		do until rsget.eof
	    	if (rsget("itemno") < rsget("totalcancelregno")) then
	    		IsErrorCancelState = true
				exit do
	    	end if
			rsget.moveNext
		loop
	end if
	rsget.close

end function

'������ ������� �ִ���
function IsDoubleCancelState(Asid, orderserial)
    dim sqlStr, result
    IsDoubleCancelState = false

	'==========================================================================
	'����������� - ��ҵ� �����Ͽ� ���� ��Ұ� �ִ���
	'==========================================================================
    sqlStr = " select top 1 "
    sqlStr = sqlStr + "     d.itemid "
    sqlStr = sqlStr + " from "

    if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_master_2003 m "
	    sqlStr = sqlStr + " 	join [db_log].[dbo].tbl_old_order_detail_2003 d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    else
	    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	    sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		m.orderserial = d.orderserial "
    end if

    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list csm "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and m.orderserial = csm.orderserial "
    sqlStr = sqlStr + " 		and csm.id = " & Asid & " "
    sqlStr = sqlStr + " 		and csm.divcd = 'A008' "
    sqlStr = sqlStr + " 		and csm.currstate <> 'B007' "
    sqlStr = sqlStr + " 		and csm.deleteyn <> 'Y' "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_detail csd "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and csm.id = csd.masterid "
    sqlStr = sqlStr + " 		and csd.orderdetailidx = d.idx "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and d.itemid <> 0 "
    sqlStr = sqlStr + " 	and d.cancelyn = 'Y' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
		IsDoubleCancelState = true
	end if
	rsget.close

end function

function IsMasterCanceled(Asid, orderserial)
    dim sqlStr, result
    IsMasterCanceled = false
    result = ""

	'==========================================================================
	' - ������ ��� �Ǿ�����
	'==========================================================================
    sqlStr = " select top 1 "
    sqlStr = sqlStr + " 	m.cancelyn as ordercancelyn "
    sqlStr = sqlStr + " from "

    if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_master_2003 m "
    else
	    sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
    end if

    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
    	if (rsget("ordercancelyn") <> "N") then
    		IsMasterCanceled = true
    	end if
    end if
    rsget.close

end function

''�ֹ� �� ������ ��� �������� üũ - ��� �Ϸ�� ������ �ִ���, �ֹ����� ��ҵȳ����� �ִ���
function IsCancelValidState(Asid, orderserial)
    dim sqlStr

    IsCancelValidState = false

    sqlStr = " select m.cancelyn, m.ipkumdiv, sum(case when d.currstate>=7 then 1 else 0 end) as invalidcount, sum(case when d.cancelyn='Y' then 1 when c.confirmitemno > d.itemno then 1 else 0 end) as detailcancelcount, sum(case when d.itemid <> 0 then 1 else 0 end) as notbeasongpaycount "
    sqlStr = sqlStr + " from "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
    else
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
    end if
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
    ''sqlStr = sqlStr + " and d.currstate>=7"
    sqlStr = sqlStr + " group by m.cancelyn, m.ipkumdiv"
	''response.write sqlStr

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        IsCancelValidState = (rsget("cancelyn")="N") and ((rsget("ipkumdiv")<=7) or ((rsget("ipkumdiv")=8) and rsget("notbeasongpaycount") = 0)) and (rsget("invalidcount")<1) and (rsget("detailcancelcount")<1)
    else
        IsCancelValidState = true
    end if
    rsget.close

end function

'// ���Ի�ǰ �ִ���
function ChkMaeipItemExist(asid)

	ChkMaeipItemExist = False

    sqlStr = " select count(*) as cnt "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_detail] d "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		a.id = d.masterid "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] od "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		d.orderdetailidx = od.idx "
    sqlStr = sqlStr + " where a.id = " & asid & " and a.divcd in ('A004', 'A010') and d.itemid <> 0 and od.omwdiv = 'M' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        ChkMaeipItemExist = (rsget("cnt") > 0)
    end if
    rsget.close

end function

function IsCancelChangeOrderValidState(changeorderserial)
    dim sqlStr

    IsCancelChangeOrderValidState = false

	sqlStr = " select count(id) as invalidcount "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" + CStr(changeorderserial) + "' "
	sqlStr = sqlStr + " 	and deleteyn <> 'Y' "
	sqlStr = sqlStr + " 	and divcd in ('A004', 'A010', 'A100', 'A111', 'A112') "
	sqlStr = sqlStr + " group by id "
	'rw sqlStr

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        IsCancelChangeOrderValidState = (rsget("invalidcount")<1)
    else
        IsCancelChangeOrderValidState = true
    end if
    rsget.close

end function

''��ǰ/ ȸ�� �������� üũ
function IsReturnRegValid(Asid, orderserial,byref ScanErr, upcheMakerid)
    ''  ��ü��۰� ��ü����� ���� �������� ����.
    ''  ��ü����� ������ ��� MakerID�� 1���� ���� �ؾ���.

    dim sqlStr
    sqlStr = " select count(d.idx) as cnt, d.isupchebeasong "
    sqlStr = sqlStr + " from "
     if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
    else
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
    end if
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and d.itemid not in (0, 100)"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
    sqlStr = sqlStr + " group by d.isupchebeasong"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        if (rsget.RecordCount>1) then
            ScanErr = "�ٹ����� ��۰� ��ü����� ���ÿ� �����Ͻ� �� �����ϴ�."
        end if
    end if
    rsget.Close

    if ScanErr<>"" then
        IsReturnRegValid = false
        exit function
    end if

    sqlStr = " select count(d.idx) as cnt, d.makerid "
    sqlStr = sqlStr + " from "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
    else
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
    end if
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and d.isupchebeasong='Y'"
    sqlStr = sqlStr + " and d.itemid not in (0, 100)"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d.idx=c.orderdetailidx"
    sqlStr = sqlStr + " group by d.makerid"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        if (rsget.RecordCount>1) then
            ScanErr = "��ü����� ��� �� �귣�� ���� �����ϼž� �մϴ�."
        else
            upcheMakerid = rsget("makerid")
        end if
    end if
    rsget.Close

    if ScanErr<>"" then
        IsReturnRegValid = false
        exit function
    end if

    IsReturnRegValid = true
end function

function IsReturnValidState(Asid, orderserial, byref iScanErr)
    dim sqlStr
    IsReturnValidState = false

    sqlStr = " select ipkumdiv, cancelyn "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        cancelyn    = rsget("cancelyn")
        ipkumdiv    = rsget("ipkumdiv")
    end if
    esget.Close

    if (cancelyn<>"N") then Exit function

    IsReturnValidState = true
end function

function setCancelMaster(Asid, orderserial)
    dim sqlStr

	if (GC_IsOLDOrder) then
		sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
	else
		sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
	end if
    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
    '' ����� �߰�
    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
	'' ������ �Է¾ȵ� ��� ������ �Է�, skyer9, 2018-02-26
	sqlStr = sqlStr + " ,baljudate=IsNULL(baljudate,getdate())" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr
end function

' ����ֹ� ����ȭ. ���󺹱�
function setRestoreCancelMaster(Asid, orderserial)
    dim sqlStr

	sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
    sqlStr = sqlStr + " set cancelyn='N'" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
    dbget.Execute sqlStr
end function

'������ ������ ��� Flag �ٸ��� ��������
'��ۺ� ���
function setCancelDetail(Asid, orderserial)
    dim sqlStr
    ''����� �߰�
	sqlStr = " update d "
	sqlStr = sqlStr + " set d.cancelyn = 'Y', d.canceldate = IsNULL(d.canceldate,getdate()) "
	sqlStr = sqlStr + " 	from "
	if (GC_IsOLDOrder) then
		sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_detail_2003 d "
	else
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
	end if
	sqlStr = sqlStr + " 		join [db_cs].[dbo].tbl_new_as_detail c "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and d.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 			and c.masterid = " + CStr(Asid) + " "
	sqlStr = sqlStr + " 			and d.idx = c.orderdetailidx "
	sqlStr = sqlStr + " 			and d.itemno = c.regitemno "
    dbget.Execute sqlStr

    '�������� - ��ǰ�Ϻ�����ΰ��
	sqlStr = " update d "
	sqlStr = sqlStr + " set d.itemno = d.itemno - c.regitemno "
	sqlStr = sqlStr + " 	from "
	if (GC_IsOLDOrder) then
		sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_detail_2003 d "
	else
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
	end if
	sqlStr = sqlStr + " 		join [db_cs].[dbo].tbl_new_as_detail c "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and d.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 			and c.masterid = " + CStr(Asid) + " "
	sqlStr = sqlStr + " 			and d.idx = c.orderdetailidx "
	sqlStr = sqlStr + " 			and d.itemno > c.regitemno "
	sqlStr = sqlStr + " 			and d.itemid <> 0 "				'// ��ۺ�� ������ �ٸ� �� ����.(������ 1��)
    dbget.Execute sqlStr

    '// ǰ������ ����
	sqlStr = " update d "
	sqlStr = sqlStr + " set d.itemlackno = d.itemlackno - c.regitemno, d.code = '03' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_mibeasong_list d "
	sqlStr = sqlStr + " 		join [db_cs].[dbo].tbl_new_as_detail c "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and d.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 			and c.masterid = " & CStr(Asid)
	sqlStr = sqlStr + " 			and d.detailidx = c.orderdetailidx "
	sqlStr = sqlStr + " 			and d.itemlackno = c.regitemno "
	sqlStr = sqlStr + " 			and d.code in ('05', '06') "
    dbget.Execute sqlStr

end function



''�ֹ� ����Ÿ ����
function recalcuOrderMaster(byVal orderserial)
	dim sqlStr

	dim CURR_IsOLDOrder : CURR_IsOLDOrder = False

	if (GC_IsOLDOrder) then
		sqlStr = " select top 1 orderserial from [db_log].[dbo].tbl_old_order_master_2003 where orderserial = '" + CStr(orderserial) + "' "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			CURR_IsOLDOrder = True
		end if
		rsget.Close
	end if

	if (CURR_IsOLDOrder) then
		sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
	else
		sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
	end if

	sqlStr = sqlStr + " set totalsum=IsNULL(T.dtotalsum,0)" + VbCrlf
	''sqlStr = sqlStr + " , totalcost=IsNULL(T.dtotalsum,0)"  + VbCrlf
	sqlStr = sqlStr + " , totalmileage=IsNULL(T.dtotalmileage,0)" + VbCrlf
	sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=IsNULL(T.dtotalitemcostCouponNotApplied,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select sum(itemcost*itemno) as dtotalsum, sum(mileage*itemno) as dtotalmileage, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied" + VbCrlf
	if (CURR_IsOLDOrder) then
		sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003" + VbCrlf
	else
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail" + VbCrlf
	end if
	sqlStr = sqlStr + "     where orderserial='" + orderserial + "'" + VbCrlf
	sqlStr = sqlStr + "     and cancelyn<>'Y'" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	if (CURR_IsOLDOrder) then
		sqlStr = sqlStr + " where [db_log].[dbo].tbl_old_order_master_2003.orderserial='" + orderserial + "'" + VbCrlf
	else
		sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.orderserial='" + orderserial + "'" + VbCrlf
	end if

	dbget.Execute sqlStr

	sqlStr = " update m " + VbCrlf
	sqlStr = sqlStr + " set " + VbCrlf
	sqlStr = sqlStr + " 	m.sumPaymentEtc = IsNull(T.realPayedsum, 0) " + VbCrlf
    if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if

	sqlStr = sqlStr + " 	left join ( " + VbCrlf
	sqlStr = sqlStr + " 		select " + VbCrlf
	sqlStr = sqlStr + " 			orderserial " + VbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(realPayedsum), 0) as realPayedsum " + VbCrlf
	sqlStr = sqlStr + " 		from " + VbCrlf
	sqlStr = sqlStr + " 			[db_order].[dbo].tbl_order_PaymentEtc " + VbCrlf
	sqlStr = sqlStr + " 		where " + VbCrlf
	sqlStr = sqlStr + " 			1 = 1 " + VbCrlf
	sqlStr = sqlStr + " 			and orderserial = '" & orderserial & "' " + VbCrlf
	sqlStr = sqlStr + " 			and acctdiv in ('200', '900') " + VbCrlf
	sqlStr = sqlStr + " 		group by " + VbCrlf
	sqlStr = sqlStr + " 			orderserial " + VbCrlf
	sqlStr = sqlStr + " 	) T " + VbCrlf
	sqlStr = sqlStr + " 	on " + VbCrlf
	sqlStr = sqlStr + " 		m.orderserial = T.orderserial " + VbCrlf
	sqlStr = sqlStr + " where " + VbCrlf
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' " + VbCrlf

	dbget.Execute sqlStr

	if (CURR_IsOLDOrder) then
		sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
	else
		sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
	end if
	sqlStr = sqlStr + " set subtotalprice=totalsum-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
	'sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=subtotalpriceCouponNotApplied-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " 	e set e.acctamount = (m.subtotalprice - m.sumpaymentetc), e.realpayedsum = (m.subtotalprice - m.sumpaymentetc) "
    if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if
	sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = e.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
	sqlStr = sqlStr + " 	and m.accountdiv = e.acctdiv "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '4' "
	sqlStr = sqlStr + " 	and m.accountdiv = '7' "

	dbget.Execute sqlStr

	'// e.acctdiv = '120' ���̹� ����Ʈ
	'// ���� �ֹ���ȣ : 16092146018
  	sqlStr = " update e set e.realPayedSum = (T.realpayedsum - T.realpayedsum120) "
  	sqlStr = sqlStr + " from "
  	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
  	sqlStr = sqlStr + " 	join ( "
  	sqlStr = sqlStr + " 		select m.orderserial, m.accountdiv, (m.subtotalprice - m.sumpaymentetc) as realpayedsum, IsNull(sum(Case when e.acctdiv = '120' then e.realpayedsum else 0 end),0) as realpayedsum120 "
	if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if
  	sqlStr = sqlStr + " 		join [db_order].[dbo].tbl_order_PaymentEtc e "
  	sqlStr = sqlStr + " 		on "
  	sqlStr = sqlStr + " 			1 = 1 "
  	sqlStr = sqlStr + " 			and m.orderserial = e.orderserial "
  	sqlStr = sqlStr + " 			and e.acctdiv in (m.accountdiv, '120') "
  	sqlStr = sqlStr + " 		where "
  	sqlStr = sqlStr + " 			m.orderserial = '" & orderserial & "' "
  	sqlStr = sqlStr + " 		group by "
  	sqlStr = sqlStr + " 			m.orderserial, m.accountdiv, (m.subtotalprice - m.sumpaymentetc) "
  	sqlStr = sqlStr + " 	) T "
  	sqlStr = sqlStr + " 	on "
  	sqlStr = sqlStr + " 		1 = 1 "
  	sqlStr = sqlStr + " 		and e.orderserial = T.orderserial "
  	sqlStr = sqlStr + " 		and e.acctdiv = T.accountdiv "
	dbget.Execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr + " set subtotalpriceCouponNotApplied = (case when T.dtotalitemcostCouponNotApplied = 0 then 0 else subtotalpriceCouponNotApplied end) "
    if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select "
	sqlStr = sqlStr + " 			orderserial, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied "
	if (CURR_IsOLDOrder) then
		sqlStr = sqlStr + " 	from [db_log].[dbo].tbl_old_order_detail_2003" + VbCrlf
	else
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail" + VbCrlf
	end if
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and orderserial = '" & orderserial & "' "
	sqlStr = sqlStr + " 			and cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and itemid <> 0 "
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	orderserial "
	sqlStr = sqlStr + " 	) T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = T.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' "

	dbget.Execute sqlStr

end function

function recalcuOrderMaster_3PL(byVal orderserial)
	dim sqlStr

	dim CURR_IsOLDOrder : CURR_IsOLDOrder = False

	sqlStr = " update [db_threepl].[dbo].[tbl_tpl_orderMaster] " + VbCrlf
	sqlStr = sqlStr + " set totalsum=IsNULL(T.dtotalsum,0)" + VbCrlf
	sqlStr = sqlStr + " , subtotalprice=IsNULL(T.reducedPrice,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select sum(itemcost*itemno) as dtotalsum, sum(IsNull(reducedPrice,0)*itemno) as reducedPrice" + VbCrlf
	sqlStr = sqlStr + " 	from [db_threepl].[dbo].[tbl_tpl_orderDetail]" + VbCrlf
	sqlStr = sqlStr + "     where orderserial='" + orderserial + "'" + VbCrlf
	sqlStr = sqlStr + "     and cancelyn<>'Y'" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	sqlStr = sqlStr + " where [db_threepl].[dbo].[tbl_tpl_orderMaster].orderserial='" + orderserial + "'" + VbCrlf
	dbget_TPL.Execute sqlStr

end function

function updateUserMileage(byVal userid)
	dim sqlStr

	'// ���ʽ�/��븶�ϸ��� ��� ����(�ű�Proc)
	sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&userid&"'"
	dbget.Execute sqlStr

	'// �ֹ����ϸ��� ��� ����(����Proc:�������)
	sqlStr = " exec [db_order].[dbo].sp_Ten_recalcuHesJumunmileage '"&userid&"'"
	dbget.Execute sqlStr

	if (GC_IsOLDOrder) then
		sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + VbCrlf
		sqlStr = sqlStr + " set [db_user].[dbo].tbl_user_current_mileage.flowerjumunmileage=IsNull(T1.totmile,0) + IsNull(T2.totmile,0) " + VbCrlf
		sqlStr = sqlStr + " from " + VbCrlf
		sqlStr = sqlStr + "     (select sum(totalmileage) as totmile " + VbCrlf
		sqlStr = sqlStr + "     from [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		sqlStr = sqlStr + "     where userid='" + CStr(userid) + "' " + VbCrlf
		sqlStr = sqlStr + "     and cancelyn='N'" + VbCrlf
		sqlStr = sqlStr + "     and ipkumdiv>3" + VbCrlf
		sqlStr = sqlStr + " ) as T1" + VbCrlf
		sqlStr = sqlStr + " join (select sum(totalmileage) as totmile" + VbCrlf
		sqlStr = sqlStr + "     from db_log.dbo.tbl_old_order_master_5YearExPired " + VbCrlf
		sqlStr = sqlStr + "     where userid='" + CStr(userid) + "' " + VbCrlf
		sqlStr = sqlStr + "     and cancelyn='N'" + VbCrlf
		sqlStr = sqlStr + "     and ipkumdiv>3" + VbCrlf
		sqlStr = sqlStr + " ) as T2" + VbCrlf
		sqlStr = sqlStr + " on 1 = 1 " + VbCrlf
		sqlStr = sqlStr + " where [db_user].[dbo].tbl_user_current_mileage.userid='" + CStr(userid) + "' " + VbCrlf

		dbget.Execute sqlStr
	end if
end function

function updateUserDeposit(byVal userid)
	dim sqlStr
	dim dataexist

	'==============================================================
	'��ġ�� ����
	sqlStr = " update c " + vbCrlf
	sqlStr = sqlStr + " set " + vbCrlf
	sqlStr = sqlStr + " 	c.currentdeposit = T.gaindeposit - T.spenddeposit " + vbCrlf
	sqlStr = sqlStr + " 	, c.gaindeposit = T.gaindeposit " + vbCrlf
	sqlStr = sqlStr + " 	, c.spenddeposit = T.spenddeposit " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	db_user.dbo.tbl_user_current_deposit c " + vbCrlf
	sqlStr = sqlStr + " 	join ( " + vbCrlf
	sqlStr = sqlStr + " 		select " + vbCrlf
	sqlStr = sqlStr + " 			'" + userid + "' as userid " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when deposit>0 then deposit else 0 end), 0) as gaindeposit " + vbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(case when deposit<0 then (deposit * -1) else 0 end), 0) as spenddeposit " + vbCrlf
	sqlStr = sqlStr + " 		from db_user.dbo.tbl_depositlog " + vbCrlf
	sqlStr = sqlStr + "     	where userid='" + userid + "'" + vbCrlf
	sqlStr = sqlStr + "     		and deleteyn='N' " + vbCrlf
	sqlStr = sqlStr + " 	) T " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		c.userid = T.userid " + vbCrlf
	'response.write sqlStr
	dbget.Execute sqlStr

	sqlStr = " select @@rowcount as cnt "
	'response.write sqlStr

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        dataexist = (rsget("cnt") > 0)
    rsget.Close

	'����Ÿ�� ������ �����Ѵ�.
	if (Not dataexist) then

		sqlStr = " if not exists (select * from db_user.dbo.tbl_user_current_deposit where userid = '" + userid + "') begin " + vbCrlf
		sqlStr = sqlStr + " 	insert into db_user.dbo.tbl_user_current_deposit(userid, currentdeposit, gaindeposit, spenddeposit) " + vbCrlf
		sqlStr = sqlStr + " 		select " + vbCrlf
		sqlStr = sqlStr + " 			'" + userid + "' " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(deposit), 0) as currentdeposit " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when deposit>0 then deposit else 0 end), 0) as gaindeposit " + vbCrlf
		sqlStr = sqlStr + " 			, IsNull(sum(case when deposit<0 then (deposit * -1) else 0 end), 0) as spenddeposit " + vbCrlf
		sqlStr = sqlStr + " 		from db_user.dbo.tbl_depositlog " + vbCrlf
		sqlStr = sqlStr + "     	where userid='" + userid + "'" + vbCrlf
		sqlStr = sqlStr + " end " + vbCrlf

		dbget.Execute sqlStr
	end if

end function

function GetDepositLogCountByAsid(orderserial, asid)
	dim sqlStr
	sqlStr = " select count(idx) as cnt from db_user.dbo.tbl_depositlog where orderserial = '" & orderserial & "' and asid = " & asid
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        GetDepositLogCountByAsid = rsget("cnt")
    rsget.Close
end function

function AddDepositCancelLogByAsid(userid, orderserial, asid)
	dim sqlStr
	sqlStr = " insert into db_user.dbo.tbl_depositlog(userid,deposit,jukyocd,jukyo,orderserial,deleteyn,asid) "
	sqlStr = sqlStr + " select userid,deposit*-1,jukyocd,jukyo+' ���',orderserial,deleteyn,asid "
	sqlStr = sqlStr + " from db_user.dbo.tbl_depositlog where userid = '" & userid & "' and orderserial = '" & orderserial & "' and asid = " & asid
	dbget.Execute sqlStr

	Call updateUserDeposit(userid)
end function

function updateUserGiftCard(byVal userid)
	dim sqlStr

	sqlStr = " exec db_cs.[dbo].[sp_Ten_ReCalcu_GiftCardSummary] '" & userid & "'"
	dbget.Execute sqlStr
end function

function ValidDeleteCS(id)
    dim sqlStr
    dim currstate

    ValidDeleteCS = false

    sqlStr = "select * from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        currstate = rsget("currstate")
    rsget.Close

    If (currstate>="B006") then Exit function

    ValidDeleteCS = true
end function

function DeleteCSProcess(id, finishuserid)
    dim sqlStr, resultCount

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
    sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " , finishuser = '" + finishuserid+ "'" + VbCrlf
    sqlStr = sqlStr + " , finishdate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)
    sqlStr = sqlStr + " and currstate<'B006'"

    dbget.Execute sqlStr, resultCount

    DeleteCSProcess = (resultCount>0)
end function

function CancelProcess(id, orderserial)
    dim IsAllCancel, IsUpdatedMile, IsUpdatedDeposit, IsUpdatedGiftCard

    dim sqlStr, userid, ipkumdiv, miletotalprice, tencardspend, allatdiscountprice

    dim refundmileagesum, refundcouponsum, allatsubtractsum
    dim refundbeasongpay, refunditemcostsum, refunddeliverypay
    dim refundadjustpay, canceltotal

    dim detailidx, orgbeasongpay, deliveritemoption, deliverbeasongpay
    dim InsureCd
    dim openMessage

    dim regDetailRows, i
    dim remaintencardspend, gubun01, gubun02

    dim orggiftcardsum, refundgiftcardsum, orgdepositsum, refunddepositsum

response.write "1" & "<br>"
    IsAllCancel = IsAllCancelRegValid(id, orderserial)

    sqlStr = " select userid, ipkumdiv, IsNULL(miletotalprice,0) as miletotalprice "
    sqlStr = sqlStr + " ,IsNULL(tencardspend,0) as tencardspend, IsNULL(allatdiscountprice,0) as allatdiscountprice" + VbCrlf
    sqlStr = sqlStr + " ,IsNULL(InsureCd,'') as InsureCd" + VbCrlf

    if (GC_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master"
    end if

    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        userid              = rsget("userid")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        allatdiscountprice  = rsget("allatdiscountprice")
        InsureCd            = rsget("InsureCd")
        ipkumdiv            = rsget("ipkumdiv")
    end if
    rsget.close

    sqlStr = " select acctdiv, IsNull(realPayedsum, 0) as realPayedsum " + VbCrlf
    sqlStr = sqlStr + " from " + VbCrlf
    sqlStr = sqlStr + " db_order.dbo.tbl_order_PaymentEtc " + VbCrlf
    sqlStr = sqlStr + " where " + VbCrlf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
    sqlStr = sqlStr + " 	and orderserial = '" + orderserial + "' " + VbCrlf
    sqlStr = sqlStr + " 	and acctdiv in ('200', '900') " + VbCrlf			'200 : ��ġ��, 900 : Giftī��

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	orgdepositsum = 0
	orggiftcardsum = 0
	do until rsget.eof
		if (CStr(rsget("acctdiv")) = "200") then
			orgdepositsum = rsget("realPayedsum")
		elseif (CStr(rsget("acctdiv")) = "900") then
			orggiftcardsum = rsget("realPayedsum")
		end if

		rsget.movenext
	loop
	rsget.close

response.write "2" & "<br>"
    sqlStr = " select r.*, a.gubun01, a.gubun02 from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"


    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")

        refundgiftcardsum   = rsget("refundgiftcardsum")
        refunddepositsum    = rsget("refunddepositsum")

        allatsubtractsum    = rsget("allatsubtractsum")

        refunditemcostsum   = rsget("refunditemcostsum")

        refundbeasongpay    = rsget("refundbeasongpay")
        refunddeliverypay   = rsget("refunddeliverypay")
        refundadjustpay     = rsget("refundadjustpay")
        canceltotal         = rsget("canceltotal")
        gubun01             = rsget("gubun01")
        gubun02             = rsget("gubun02")

    else
        refundmileagesum    = 0
        refundcouponsum     = 0
        allatsubtractsum    = 0

        refundgiftcardsum   = 0
        refunddepositsum    = 0

        refunditemcostsum   = 0

        refundbeasongpay    = 0
        refunddeliverypay   = 0
        refundadjustpay     = 0
        canceltotal         = 0
    end if
    rsget.close

'' ���ϸ��� ȯ��

    IsUpdatedMile = false
response.write "3" & "<br>"
    if (userid<>"") and (IsAllCancel) and (miletotalprice<>0) then
        '' ��ü ����ΰ�� �ֹ��� ��ҷ� jukyocd : 2 ��ǰ����, 3 : �κ���ҽ� ȯ�����ϸ���
        sqlStr = " update [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf
        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "��� ���ϸ��� ȯ�� : " & miletotalprice
        else
            openMessage = openMessage + VbCrlf + "��� ���ϸ��� ȯ�� : " & miletotalprice
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundmileagesum<>0) then
        '' �κ� ����ε� ���ϸ��� ȯ���� ���.
		if (GC_IsOLDOrder) then
			sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		else
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
		end if

        sqlStr = sqlStr + " set miletotalprice=miletotalprice + " + CStr(refundmileagesum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " insert into [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refundmileagesum*-1) + ""
        sqlStr = sqlStr + " ,'3'"
        sqlStr = sqlStr + " ,'��ǰ���� ��� ȯ��'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "��� ���ϸ��� ȯ�� : " & refundmileagesum
        else
            openMessage = openMessage + VbCrlf + "��� ���ϸ��� ȯ�� : " & refundmileagesum
        end if
    end if

'��ġ��ȯ��
	IsUpdatedDeposit = false
    if (userid<>"") and (IsAllCancel) and (orgdepositsum <> 0) then
        '' ��ü ����ΰ�� �ֹ��� ��ҷ� jukyocd : 100 ��ǰ����, 10 : �κ���ҽ� ��ġ�� ȯ��
        sqlStr = " update [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('100','10')" + VbCrlf					'100 : ��ǰ���Ż�� / 10 : �Ϻ�ȯ�� (���� : db_user.dbo.tbl_deposit_gubun)
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "��� ��ġ�� ȯ�� : " & orgdepositsum
        else
            openMessage = openMessage + VbCrlf + "��� ��ġ�� ȯ�� : " & orgdepositsum
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refunddepositsum <> 0) then
        '' �κ� ����ε� ��ġ�� ȯ���� ���.

        sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
        sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refunddepositsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and acctdiv='200'" + VbCrlf
        dbget.Execute sqlStr

		if (GC_IsOLDOrder) then
			sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		else
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
		end if

        sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refunddepositsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " insert into [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " (userid, deposit, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refunddepositsum*-1) + ""
        sqlStr = sqlStr + " ,'10'"
        sqlStr = sqlStr + " ,'��ǰ���� ��� ȯ��'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "��� ��ġ�� ȯ�� : " & refunddepositsum
        else
            openMessage = openMessage + VbCrlf + "��� ��ġ�� ȯ�� : " & refunddepositsum
        end if
    end if

'Giftī��ȯ��
	IsUpdatedGiftCard = false
    if (userid<>"") and (IsAllCancel) and (orggiftcardsum <> 0) then
        '' ��ü ����ΰ�� �ֹ��� ��ҷ� jukyocd : 200 ��ǰ����, 300 : �κ���ҽ� Giftī�� ȯ��
        sqlStr = " update [db_user].[dbo].tbl_giftcard_log " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('200','300')" + VbCrlf					'200 : ��ǰ���Ż�� / 300 : �Ϻ�ȯ�� (���� : db_user.dbo.tbl_giftcard_gubun)
        dbget.Execute sqlStr

        IsUpdatedGiftCard = true

        if openMessage="" then
            openMessage = openMessage + "��� Giftī�� ȯ�� : " & orggiftcardsum
        else
            openMessage = openMessage + VbCrlf + "��� Giftī�� ȯ�� : " & orggiftcardsum
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundgiftcardsum <> 0) then
        '' �κ� ����ε� Giftī�� ȯ���� ���.

        sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
        sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refundgiftcardsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and acctdiv='900'" + VbCrlf
        dbget.Execute sqlStr

		if (GC_IsOLDOrder) then
			sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		else
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
		end if

        sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refundgiftcardsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
        dbget.Execute sqlStr

        sqlStr = " insert into [db_user].[dbo].tbl_giftcard_log " + VbCrlf
        sqlStr = sqlStr + " (userid, useCash, jukyocd, jukyo, orderserial, deleteyn, reguserid) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refundgiftcardsum*-1) + ""
        sqlStr = sqlStr + " ,'300'"
        sqlStr = sqlStr + " ,'��ǰ���� ��� ȯ��'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " ,'" + CStr(session("ssbctid")) + "'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedGiftCard = true

        if openMessage="" then
            openMessage = openMessage + "��� Giftī�� ȯ�� : " & refundgiftcardsum
        else
            openMessage = openMessage + VbCrlf + "��� Giftī�� ȯ�� : " & refundgiftcardsum
        end if
    end if

'' ���α� ȯ��
    if (IsAllCancel) and (tencardspend<>0) then
        sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
	    sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
	    sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

	    dbget.Execute sqlStr

	    if openMessage="" then
            openMessage = openMessage + "��� ���ʽ����� ȯ��"
        else
            openMessage = openMessage + VbCrlf + "��� ���ʽ����� ȯ��"
        end if
    end if

    if (Not IsAllCancel) and (refundcouponsum<>0) then
        '' �κ� ����ΰ�� - ȯ���� ��ŭ ��..
		if (GC_IsOLDOrder) then
			sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		else
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
		end if
        sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(refundcouponsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        ''��ü ȯ���� ��츸 ������ ������
        sqlStr = "select IsNULL(tencardspend,0) as tencardspend " + VbCrlf
		if (GC_IsOLDOrder) then
			sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003"
		else
			sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master"
		end if
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            remaintencardspend = rsget("tencardspend")
        rsget.close

        ''���� ���α� ������ �ְ�, ���� ���������� ������� ��ü  ȯ��
        if (tencardspend>0) then
            if (remaintencardspend=0)   then
                sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
            	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
            	sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

            	dbget.Execute sqlStr

            	if openMessage="" then
                    openMessage = openMessage + "��� ���α�  ȯ��"
                else
                    openMessage = openMessage + VbCrlf + "��� ���α�  ȯ��"
                end if
            else
                ''(�Ǵ�, %������ ��� ����,�ܼ������� ��� �����ϰ� ȯ������./ �κ���� ) C004 CD01
                if (ipkumdiv>3) and (Not ((gubun01="C004") and (gubun02="CD01"))) then
                    sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
                	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
                	sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
                	sqlStr = sqlStr + " and coupontype=1"

                	dbget.Execute sqlStr

                	if openMessage="" then
                        openMessage = openMessage + "��� ���α�  ȯ��."
                    else
                        openMessage = openMessage + VbCrlf + "��� ���α�  ȯ��."
                    end if
                end if
            end if
        end if

    end if

    '' �ÿ�ī�� ���� ����
    if (IsAllCancel) and (allatdiscountprice<>0) then

    end if

    if (Not IsAllCancel) and (allatsubtractsum<>0) then
		if (GC_IsOLDOrder) then
			sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		else
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
		end if
        sqlStr = sqlStr + " set allatdiscountprice=allatdiscountprice + " + CStr(allatsubtractsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        if openMessage="" then
            openMessage = openMessage + "�ÿ�ī�� ���� ���� : " & allatsubtractsum
        else
            openMessage = openMessage + VbCrlf + "�ÿ�ī�� ���� ���� : " & allatsubtractsum
        end if
    end if

response.write "4" & "<br>"

	'��ۺ� ���� ��ҵȴ�. setCancelDetail()

response.write "5" & "<br>"

    if (IsAllCancel) then
	    ''��ü ����ΰ��
	    '' �ֹ�  master ��� ����
	    call setCancelMaster(id, orderserial)

	    if openMessage="" then
            openMessage = openMessage + "�ֹ���� �Ϸ�"
        else
            openMessage = openMessage + VbCrlf + "�ֹ���� �Ϸ�"
        end if
	else
	    ''�κ� ����ΰ��
	    '' �ֹ�  detail ��� ����
	    call setCancelDetail(id, orderserial)

		if (refunddeliverypay <> 0) then
			'// ��ü �߰���ۺ� �ΰ�
			Call AddBeasongpayForCancel(id, orderserial)
		end if

	    call reCalcuOrderMaster(orderserial)

	    if openMessage="" then
            openMessage = openMessage + "�ֹ��κ���� �Ϸ�"
        else
            openMessage = openMessage + VbCrlf + "�ֹ��κ���� �Ϸ�"
        end if
	end if

    ''���ϸ����� �ֹ��� ��� �� �����ؾ���.
    '��ġ��, Giftī�� ����
    if (userid<>"") then
        Call updateUserMileage(userid)

        if IsUpdatedDeposit then
        	Call updateUserDeposit(userid)
        end if

        if IsUpdatedGiftCard then
        	Call updateUserGiftCard(userid)
        end if
    end if

    ''�ֱ� �ֹ����� ���� 2015/08/12
    if (userid<>"") and (IsAllCancel) then
        sqlStr = "exec [db_order].[dbo].sp_Ten_Recalcu_His_recent_OrderCNT '" & userid & "'"
        dbget.Execute(sqlStr)
    end if

    '' ''���ں����� �߱޵� ��� ���
    '' if (InsureCd="0") then
    ''     Call UsafeCancel(orderserial)
    '' end if

    if (openMessage<>"") then
        call AddCustomerOpenContents(id, openMessage)
    end if
end function

function CancelAddBeasongpayForCancel(id)
	dim sqlStr
	dim refunddeliverypay, lastitemoption, masteridx, makerid

    sqlStr = " select r.refunddeliverypay from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    ''sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate='B007'"
	sqlStr = sqlStr + " and r.refunddeliverypay <> 0 "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
		refunddeliverypay   = rsget("refunddeliverypay")
    else
        refunddeliverypay   = 0
    end if
    rsget.close

	if (refunddeliverypay = 0) then
		exit function
	end if

	sqlStr = " select top 1 (case when d.isupchebeasong = 'Y' then d.makerid else '' end) as makerid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_detail d on a.id = d.masterid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " + CStr(id)
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
		makerid  = rsget("makerid")
    end if
    rsget.close

	sqlStr = " update d "
	sqlStr = sqlStr + " set d.cancelyn = 'Y', d.canceldate = getdate() "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_list a "

	if (GC_IsOLDOrder) then
		sqlStr = sqlStr + " 	join [db_log].[dbo].tbl_old_order_detail_2003 d on a.orderserial = d.orderserial "
	else
		sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail d on a.orderserial = d.orderserial "
	end if

	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " & id
	sqlStr = sqlStr + " 	and d.itemid = 0 "
	sqlStr = sqlStr + " 	and d.itemoption >= '8000' "
	sqlStr = sqlStr + " 	and d.itemoption < '9000' "
	sqlStr = sqlStr + " 	and d.makerid = '" & makerid & "' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
    dbget.Execute sqlStr

end function

function AddBeasongpayForCancel(id, orderserial)
	dim sqlStr
	dim refunddeliverypay, lastitemoption, masteridx, makerid

    sqlStr = " select r.*, a.gubun01, a.gubun02, m.idx as masteridx from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"

	if (GC_IsOLDOrder) then
		sqlStr = sqlStr + " , [db_log].[dbo].tbl_old_order_master_2003 m " & vbCrlf
	else
		sqlStr = sqlStr + " , [db_order].[dbo].tbl_order_master m " & vbCrlf
	end if

    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
	sqlStr = sqlStr + " and a.orderserial=m.orderserial"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
		refunddeliverypay   = rsget("refunddeliverypay")
		masteridx   		= rsget("masteridx")
    else
        refunddeliverypay   = 0
		masteridx			= ""
    end if
    rsget.close

	if (refunddeliverypay = 0) then
		exit function
	end if

	sqlStr = " select IsNull(max(itemoption), '8000') as itemoption "
	sqlStr = sqlStr + " from "
	if (GC_IsOLDOrder) then
		sqlStr = sqlStr + " 	[db_log].[dbo].tbl_old_order_detail_2003 "
	else
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail "
	end if

	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" & orderserial & "' "
	sqlStr = sqlStr + " 	and itemid = 0 "
	sqlStr = sqlStr + " 	and itemoption >= '8000' "
	sqlStr = sqlStr + " 	and itemoption < '9000' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
		lastitemoption  = rsget("itemoption")
    else
        lastitemoption  = ""
    end if
    rsget.close

	if (lastitemoption = "") then
		exit function
	end if

	lastitemoption = CStr(CLng(lastitemoption) + 1)

	sqlStr = " select top 1 (case when d.isupchebeasong = 'Y' then d.makerid else '' end) as makerid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_detail d on a.id = d.masterid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " + CStr(id)
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
		makerid  = rsget("makerid")
    end if
    rsget.close

	if (refunddeliverypay <> 0) and (Left(now, 10) >= "2019-01-01") then
		'�߰���ۺ�
		sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
		sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
		sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
		sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, currstate,beasongdate,upcheconfirmdate,itemcouponidx, bonuscouponidx)" & vbCrlf
		sqlStr = sqlStr + " select " & CStr(masteridx) & vbCrlf
		sqlStr = sqlStr + " ,'" & orderserial & "'" & vbCrlf
		sqlStr = sqlStr + " ,0" & vbCrlf
		sqlStr = sqlStr + " ,'" & lastitemoption & "'" & vbCrlf
		sqlStr = sqlStr + " ,1" & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , Round(((1.0 * " + CStr(refunddeliverypay * -1) + ") / 11.0), 0) " & vbCrlf
		sqlStr = sqlStr + " , 0 " & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , '�߰���ۺ�' " & vbCrlf
		sqlStr = sqlStr + " , (case when '" & makerid & "' <> '' then '��ü����' else '' end) " & vbCrlf
		sqlStr = sqlStr + " , '" + makerid + "' " & vbCrlf
		sqlStr = sqlStr + " , (case when '" & makerid & "' <> '' then " + CStr(refunddeliverypay * -1) + " else 0 end) " & vbCrlf
		sqlStr = sqlStr + " , 'Y' " & vbCrlf
		sqlStr = sqlStr + " , NULL " & vbCrlf
		sqlStr = sqlStr + " , 'N' " & vbCrlf
		sqlStr = sqlStr + " , '01' " & vbCrlf
		sqlStr = sqlStr + " , NULL " & vbCrlf
		sqlStr = sqlStr + " , NULL " & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , (case when '" & makerid & "' <> '' then " + CStr(refunddeliverypay * -1) + " else 0 end) " & vbCrlf
		sqlStr = sqlStr + " , 0 " & vbCrlf
		sqlStr = sqlStr + " , 0 " & vbCrlf
		sqlStr = sqlStr + " ,'0'" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
		sqlStr = sqlStr + " ,NULL, NULL " & vbCrlf
		sqlStr = sqlStr + " from " & vbCrlf

		if (GC_IsOLDOrder) then
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m " & vbCrlf
		else
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m " & vbCrlf
		end if

		sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_list a "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	m.orderserial = a.orderserial "
		sqlStr = sqlStr + " join db_cs.dbo.tbl_as_refund_info r "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	a.id = r.asid "
		sqlStr = sqlStr + " where a.id = " & CStr(id)
		dbget.Execute sqlStr

		sqlStr = " update r "
		sqlStr = sqlStr + " set r.isRefundDeliveryPayAddedToOrder = 'Y' "
		sqlStr = sqlStr + " from " & vbCrlf
		sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list a "
		sqlStr = sqlStr + " join db_cs.dbo.tbl_as_refund_info r "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	a.id = r.asid "
		sqlStr = sqlStr + " where a.id = " & CStr(id)
		dbget.Execute sqlStr
	end if
end function

'// �ֹ���� �Ϸ�� �������� ������ ��ǰ�ݾ� ������Ʈ
function UpdateCancelJupsuCSPrice(id, orderserial)
	dim sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	r.orgitemcostsum = r.orgitemcostsum - T.refunditemcostsum, "
	sqlStr = sqlStr + " 	r.orgbeasongpay = r.orgbeasongpay - T.refundbeasongpay, "
	sqlStr = sqlStr + " 	r.orgallatdiscountsum = r.orgallatdiscountsum + T.allatsubtractsum, "
	sqlStr = sqlStr + " 	r.orgcouponsum = r.orgcouponsum + T.refundcouponsum, "
	sqlStr = sqlStr + " 	r.orgmileagesum = r.orgmileagesum + T.refundmileagesum, "
	sqlStr = sqlStr + " 	r.orggiftcardsum = r.orggiftcardsum + T.refundgiftcardsum, "
	sqlStr = sqlStr + " 	r.orgdepositsum = r.orgdepositsum + T.refunddepositsum "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_as_refund_info] r on a.id = r.asid "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select top 1 refunditemcostsum, refundbeasongpay, allatsubtractsum, refundcouponsum, refundmileagesum, refundgiftcardsum, refunddepositsum "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 			[db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr + " 			join [db_cs].[dbo].[tbl_as_refund_info] r on a.id = r.asid "
	sqlStr = sqlStr + " 		where a.id = " & CStr(id) & " and a.currstate = 'B007' and a.divcd = 'A008' and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	) T "
	sqlStr = sqlStr + " 	on 1=1 "
	sqlStr = sqlStr + " where a.orderserial = '" & orderserial & "' and a.id <> " & CStr(id) & " and a.currstate < 'B007' and a.divcd = 'A008' and a.deleteyn = 'N' "
	dbget.Execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	r.orgsubtotalprice = r.orgitemcostsum + r.orgbeasongpay - r.orgallatdiscountsum - r.orgcouponsum - r.orgmileagesum - r.orggiftcardsum - r.orgdepositsum "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_as_refund_info] r on a.id = r.asid "
	sqlStr = sqlStr + " where a.orderserial = '" & orderserial & "' and a.id <> " & CStr(id) & " and a.currstate < 'B007' and a.divcd = 'A008' and a.deleteyn = 'N' "
	dbget.Execute sqlStr

end function

function CheckRefundFinish(id, orderserial,byRef RefreturnMethod,byRef Refrealrefund)
    dim sqlStr
    dim returnmethod, refundrequire, refundresult
    dim realrefund ,userid
    dim title

    realrefund = 0

    sqlStr = "select r.*, a.userid, a.title from "
    sqlStr = sqlStr + " [db_cs].[dbo].tbl_as_refund_info r,"
    sqlStr = sqlStr + " [db_cs].dbo.tbl_new_as_list a"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        title           = rsget("title")

        realrefund      = refundrequire - refundresult

        RefreturnMethod = returnmethod
        Refrealrefund   = realrefund
    end if
    rsget.Close

    ''���ϸ��� ȯ��
    if (returnmethod="R900") then
        sqlStr = "insert into [db_user].[dbo].tbl_mileagelog"
        sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn)"
        sqlStr = sqlStr + " values('" + userid + "'," + CStr(realrefund) + ",'999','" & title & "','" + orderserial + "','N')"
        dbget.Execute sqlStr

        sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call updateUserMileage(userid)

        call AddCustomerOpenContents(id, "���ϸ��� ȯ�� �Ϸ�: " & CStr(realrefund))
    elseif (returnmethod="R910") then
    	'��ġ�� ��ȯ

    	title = Replace(title, "���ϸ���", "��ġ��")
    	title = Replace(title, "������", "��ġ��")

        sqlStr = "insert into [db_user].[dbo].tbl_depositlog"
        sqlStr = sqlStr + " (userid, deposit, jukyocd, jukyo, orderserial, deleteyn)"
        sqlStr = sqlStr + " values('" + userid + "'," + CStr(realrefund) + ",'200','" & title & "','" + orderserial + "','N')"
        dbget.Execute sqlStr

        sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call updateUserDeposit(userid)

        call AddCustomerOpenContents(id, "��ġ�� ȯ�� �Ϸ�: " & CStr(realrefund))
    elseif (returnmethod<>"R000") then
        sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call AddCustomerOpenContents(id, "ȯ��(���) �Ϸ�: " & CStr(realrefund))
    end if

end function

function AddminusOrderLink(asid, minusorderserial)
    dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr & " set refminusorderserial='"&minusorderserial&"'"
    sqlStr = sqlStr & " where id="&asid
    dbget.Execute sqlStr
end function

function AddminusOrderLink_3PL(asid, minusorderserial)
    dim sqlStr

    sqlStr = " update [db_threepl].[dbo].[tbl_tpl_as_list]"
    sqlStr = sqlStr & " set refminusorderserial='"&minusorderserial&"'"
    sqlStr = sqlStr & " where id="&asid
    dbget_TPL.Execute sqlStr
end function

function AddChangeOrderLink(asid, changeorderserial)

	Call AddChangeOrderJupsuLink(asid, changeorderserial)

	'// �±�ȯȸ��(A111)�� ��ȯ�ֹ��� ��ϵǸ� �±�ȯ���(A100)���� ���� ���
	Call AddChangeOrderChulgoLink(asid, changeorderserial)

end function

function AddChangeOrderJupsuLink(asid, changeorderserial)
	dim sqlStr
	dim chulgoasid

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr & " set refchangeorderserial='"&changeorderserial&"'"
    sqlStr = sqlStr & " where id="&asid
    dbget.Execute sqlStr

end function

function AddChangeOrderChulgoLink(asid, changeorderserial)
	dim sqlStr
	dim chulgoasid

    sqlStr = "select top 1 IsNull(refasid, 0) as refasid from [db_cs].[dbo].tbl_new_as_list where id = " + CStr(asid) + " "

	chulgoasid = 0

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        chulgoasid    = rsget("refasid")
    end if
    rsget.Close

	if (chulgoasid <> 0) then
	    sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
	    sqlStr = sqlStr & " set refchangeorderserial='"&changeorderserial&"'"
	    sqlStr = sqlStr & " where id="&chulgoasid
	    dbget.Execute sqlStr
	end if

end function

function SetRefAsid(asid, refasid)
	dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
    sqlStr = sqlStr + " set refasid = " & CStr(refasid) & " "
    sqlStr = sqlStr + " where id = " & CStr(asid) & " "
    dbget.Execute sqlStr
end function

function GetRefAsid(asid)
	dim sqlStr

    sqlStr = "select top 1 refasid from [db_cs].[dbo].tbl_new_as_list where id = " + CStr(asid) + " "

	GetRefAsid = 0

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        GetRefAsid    = rsget("refasid")
    end if
    rsget.Close

end function

function CheckNRegRefund(id, orderserial, reguserid)
    '' A003 ȯ�ҿ�û , A005 �ܺθ�ȯ�ҿ�û , A007 �ſ�ī��/�ǽð���ü��ҿ�û
    '' Result -1, or newAsID
    dim divcd
    dim returnmethod, gubun01, gubun02

    dim sqlStr, RegDivCd
    dim title, contents_jupsu
    dim NewRegedID

    CheckNRegRefund = -1

    sqlStr = " select a.divcd, a.gubun01, a.gubun02"
    sqlStr = sqlStr + " , r.returnmethod "
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + "     on a.id=r.asid"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        divcd                = rsget("divcd")
        returnmethod         = rsget("returnmethod")
        gubun01              = rsget("gubun01")
        gubun02              = rsget("gubun02")

        if IsNULL(returnmethod) then returnmethod=""
    end if
    rsget.close

    'R007 ������ȯ��
    'R020 �ǽð���ü���
    'R050 ���������� ���
    'R080 �ÿ�ī�����
    'R100 �ſ�ī�����
    'R550 ���������
    'R560 ����Ƽ�����
    'R120 �ſ�ī��κ����
    'R400 �޴������
	'R420 �޴����κ����
    'R900 ���ϸ�����ȯ��
    'R910 ��ġ��ȯ��
    'R022 �ǽð���ü�κ����(NP)
    'R150 �̴Ϸ�Ż���

	title = GetRefundMethodString(returnmethod)

    if (returnmethod="R000") or (Trim(returnmethod)="") then
        Exit function
    elseif (returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420") or (returnmethod="R150") then
        RegDivCd = "A007"

        contents_jupsu = paygateTid
    elseif (returnmethod="R050") then
        RegDivCd = "A005"
    elseif (returnmethod="R900") then
        RegDivCd = "A003"
    elseif (returnmethod="R910") then
        RegDivCd = "A003"
    elseif (returnmethod<>"") then
        RegDivCd = "A003"
        contents_jupsu = ""
    end if

    if (divcd="A008") then
        title = "�ֹ� ��� �� " + title
    elseif (divcd="A004") then
        title = "��ǰ ó�� �� " + title
    elseif (divcd="A010") then
        title = "ȸ�� ó�� �� " + title
    elseif (divcd="A100") then
        title = "��ȯ ��� �� " + title
    end if

    if (RegDivCd<>"") then
        NewRegedID =  RegCSMaster(RegDivCd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

''        '''���� ���(��ǰ) ���� ����
''        sqlStr = " insert into [db_cs].[dbo].tbl_as_refund_info"
''        sqlStr = sqlStr + " (asid, returnmethod, refundrequire, rebankname, rebankaccount, "
''        sqlStr = sqlStr + " rebankownername, paygateTid, paygateresultTid, paygateresultMsg) "
''        sqlStr = sqlStr + " select " + CStr(NewRegedID)
''        sqlStr = sqlStr + " ,returnmethod, refundrequire, rebankname, rebankaccount, "
''        sqlStr = sqlStr + " rebankownername, paygateTid, paygateresultTid, paygateresultMsg "
''        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_as_refund_info"
''        sqlStr = sqlStr + " where asid=" + CStr(id)
''        dbget.Execute sqlStr
''
''		'���� CS''''''''''''''''''''*******************************
''        sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
''        sqlStr = sqlStr + " set refasid = " & CStr(id) & " "
''        sqlStr = sqlStr + " where id = " & CStr(NewRegedID) & " "
''        dbget.Execute sqlStr

        Call CopyWebCancelRefundInfo(id, NewRegedID)

        CheckNRegRefund = NewRegedID
    end if
end function

'// ===========================================================================
'// ���̳ʽ� �ֹ� �Է��� ���� ����üũ
'// ===========================================================================
''���ֹ�
''      ----> ��ȯ�ֹ�1
''                     ----> ��ȯ�ֹ�2
''  I
''  V
'' ���̳ʽ�1     I
''               I              I
''               V              I
''            ���̳ʽ�2         I
''                              I
''                              V
''                          ���̳ʽ�3
'// ===========================================================================
'// ���ֹ� + ��ȯ�ֹ� ���� >= ���ֹ�(and ��ȯ�ֹ�) �� ���� ���̳ʽ� �ֹ� ���� + CS���� ����
'// ===========================================================================
function CheckOverMinusOrderItemnoExist(id, orderserial)
	dim sqlStr

	CheckOverMinusOrderItemnoExist = False

    ''�����Ǵ� �������� ���� ���̳ʽ�+ �߰� ���̳ʽ�  �հ谡 ū�� üũ (�ߺ�����)
    if (GC_IsOLDOrder) then
        '' ���� �ֹ��� ���.. Skip
        Exit function
    else
        '// TODO : ���� ������ CS �� ī��Ʈ ���� �ʰ� �ִ�.
        sqlStr = " exec db_order.dbo.sp_Ten_MinusOrderInValidCnt " & CStr(id) & ",'" & orderserial & "'"
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if Not (rsget.Eof) then
            if (rsget("InvalidCnt") > 0) then
            	CheckOverMinusOrderItemnoExist = True
            end if
        end if
        rsget.Close
    end if
end function

function CheckOverMinusOrderItemnoExist_3PL(id, orderserial)
	dim sqlStr

	CheckOverMinusOrderItemnoExist_3PL = False

    ''�����Ǵ� �������� ���� ���̳ʽ�+ �߰� ���̳ʽ�  �հ谡 ū�� üũ (�ߺ�����)
    if (GC_IsOLDOrder) then
        '' ���� �ֹ��� ���.. Skip
        Exit function
    else
        '// TODO : ���� ������ CS �� ī��Ʈ ���� �ʰ� �ִ�.
        sqlStr = " exec db_threepl.dbo.usp_Tpl_MinusOrderInValidCnt " & CStr(id) & ",'" & orderserial & "'"
        rsget_TPL.Open sqlStr, dbget_TPL, 1
        if Not (rsget_TPL.Eof) then
            if (rsget_TPL("InvalidCnt") > 0) then
            	CheckOverMinusOrderItemnoExist_3PL = True
            end if
        end if
        rsget_TPL.Close
    end if
end function

function CheckOverChangeOrderItemnoExist(id, orderserial)
	dim sqlStr

	CheckOverChangeOrderItemnoExist = False

    ''�ʰ� ��ȯ�ֹ� üũ (�ߺ�����)
    if (GC_IsOLDOrder) then
        '' ���� �ֹ��� ���.. Skip
        Exit function
    else
        '// TODO : ���� ������ CS �� ī��Ʈ ���� �ʰ� �ִ�.
        sqlStr = " exec db_order.dbo.sp_Ten_ChangeOrderInValidCnt " & CStr(id) & ",'" & orderserial & "'"
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if Not (rsget.Eof) then
            if (rsget("InvalidCnt") > 0) then
            	CheckOverChangeOrderItemnoExist = True
            end if
        end if
        rsget.Close
    end if
end function

'// ===========================================================================
'// * ���̳ʽ� �ֹ� ��� ���� üũ
'// ===========================================================================
'// ���ֹ� + ��ȯ�ֹ� �ݾ� >= ���ֹ�(and ��ȯ�ֹ�) �� ���� ���̳ʽ� �ֹ� �ݾ�
'// ===========================================================================
'// CheckOverMinusOrderItemnoExist ����
'// ===========================================================================
function CheckOverMinusOrderPriceExist(orderserial, byref ErrStr)

	CheckOverMinusOrderPriceExist = False

end function

function IsOrderExists(orderserial)
	dim sqlStr

	IsOrderExists = True

    ''���ֹ��� ��ȸ
    sqlStr = " select idx "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master"
    end if
    sqlStr = sqlStr  + " where orderserial='" + orderserial + "'"
    sqlStr = sqlStr  + " and cancelyn='N'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if rsget.Eof then
		IsOrderExists = False
    end if
    rsget.Close

end function

function IsOrderExists_3PL(orderserial)
	dim sqlStr

	IsOrderExists_3PL = True

    ''���ֹ��� ��ȸ
    sqlStr = " select idx "
    sqlStr = sqlStr  + " from [db_threepl].[dbo].[tbl_tpl_orderMaster]"
    sqlStr = sqlStr  + " where orderserial='" + orderserial + "'"
    sqlStr = sqlStr  + " and cancelyn='N'"

    rsget_TPL.Open sqlStr,dbget_TPL,1
    if rsget_TPL.Eof then
		IsOrderExists_3PL = False
    end if
    rsget_TPL.Close

end function

function IsCSDetailExists(asid, orderserial)
	dim sqlStr

    sqlStr = " select count(*) as cnt from" & Vbcrlf
    sqlStr = sqlStr  + " [db_cs].[dbo].tbl_new_as_list a," & Vbcrlf
    sqlStr = sqlStr  + " [db_cs].[dbo].tbl_new_as_detail d" & Vbcrlf
    sqlStr = sqlStr  + " where a.id=" & CStr(asid) & Vbcrlf
    sqlStr = sqlStr  + " and a.id=d.masterid" & Vbcrlf
    sqlStr = sqlStr  + " and a.orderserial='" + orderserial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IsCSDetailExists    = rsget("cnt")>0
    rsget.Close

end function

function IsCSDetailExists_3PL(asid, orderserial)
	dim sqlStr

    sqlStr = " select count(*) as cnt from" & Vbcrlf
    sqlStr = sqlStr  + " [db_threepl].[dbo].[tbl_tpl_as_list] a," & Vbcrlf
    sqlStr = sqlStr  + " [db_threepl].[dbo].[tbl_tpl_as_detail] d" & Vbcrlf
    sqlStr = sqlStr  + " where a.id=" & CStr(asid) & Vbcrlf
    sqlStr = sqlStr  + " and a.id=d.masterid" & Vbcrlf
    sqlStr = sqlStr  + " and a.orderserial='" + orderserial + "'"

    rsget_TPL.Open sqlStr,dbget_TPL,1
        IsCSDetailExists_3PL    = rsget_TPL("cnt")>0
    rsget_TPL.Close

end function

function GetOrgOrderPriceInfo(orderserial, byref sumSubTotalPrice, byref sumPaymentEtc, byref sumTencardSpend, byref sumMileTotalPrice, byref sumSpendmembership, byref sumAllatdiscountprice, byref sumDepositPrice, byref sumGiftCardPrice, byref sumPercentCoupon)
	dim sqlStr

	sqlStr = " exec db_order.dbo.usp_Ten_GetOrderPriceInfoSUM '" & orderserial & "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then

		sumSubTotalPrice		= rsget("sumSubTotalPrice")
		sumPaymentEtc			= rsget("sumPaymentEtc")
		sumTencardSpend			= rsget("sumTencardSpend")
		sumMileTotalPrice		= rsget("sumMileTotalPrice")
		sumSpendmembership		= rsget("sumSpendmembership")
		sumAllatdiscountprice	= rsget("sumAllatdiscountprice")
		sumDepositPrice			= rsget("sumDepositPrice")
		sumGiftCardPrice		= rsget("sumGiftCardPrice")
		sumPercentCoupon		= rsget("sumPercentCoupon")

    end if
    rsget.Close

end function

function GetMinusOrderPriceInfo(orderserial, byref sumMinusSubTotalPrice, byref sumMinusPaymentEtc, byref sumMinusTencardSpend, byref sumMinusMileTotalPrice, byref sumMinusSpendmembership, byref sumMinusAllatdiscountprice, byref sumMinusDepositPrice, byref sumMinusGiftCardPrice, byref sumMinusPercentCoupon)
	dim sqlStr

	sqlStr = " exec db_order.dbo.usp_Ten_GetMinusOrderPriceInfoSUM '" & orderserial & "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then

		sumMinusSubTotalPrice		= rsget("sumSubTotalPrice")
		sumMinusPaymentEtc			= rsget("sumPaymentEtc")
		sumMinusTencardSpend		= rsget("sumTencardSpend")
		sumMinusMileTotalPrice		= rsget("sumMileTotalPrice")
		sumMinusSpendmembership		= rsget("sumSpendmembership")
		sumMinusAllatdiscountprice	= rsget("sumAllatdiscountprice")
		sumMinusDepositPrice		= rsget("sumDepositPrice")
		sumMinusGiftCardPrice		= rsget("sumGiftCardPrice")
		sumMinusPercentCoupon		= rsget("sumPercentCoupon")

    end if
    rsget.Close

end function

function CheckRefundPrice(id, orderserial, byref ErrStr)
	dim sqlStr

	ErrStr = ""
	sqlStr = " exec [db_cs].[dbo].usp_Ten_CS_Refund_Check_Price '" & orderserial & "', " & id
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		ErrStr = rsget("msg")
	end if
	rsget.Close
end function

'// ��� �� �ܿ���ǰ �ݾ��� 3���� �̸��̸� ���ϸ��� ����
function CheckRefundMileage(id, orderserial)
	dim sqlStr

	sqlStr = " exec [db_cs].[dbo].usp_Ten_CS_Refund_Check_Mileage '" & orderserial & "', " & id
    dbget.Execute sqlStr
end function

function CheckNAddMinusOrder(id, orderserial, reguserid,byref MinusOrderserial, byref ErrStr)
    dim sqlStr
    dim currjupsusum, orgidx
    dim userid, sitename
    dim AsDetailExists
    dim MinusMiletotalprice, MinusDepositPrice, MinusGiftCardPrice
    ''dim totalpreminussum
    ''totalpreminussum = 0

    dim orgsubtotalprice, orgsumPaymentEtc, orgtencardspend, orgmiletotalprice, orgspendmembership, orgallatdiscountprice, orgdepositsum, orggiftcardsum, orgpercentcouponsum
    dim totalpreminus_subtotalprice, totalpreminus_sumPaymentEtc, totalpreminus_tencardspend, totalpreminus_miletotalprice, totalpreminus_spendmembership, totalpreminus_allatdiscountprice, totalpreminus_depositsum, totalpreminus_giftcardsum, totalpreminus_percentcouponsum

    totalpreminus_subtotalprice = 0
    totalpreminus_sumPaymentEtc = 0
    totalpreminus_tencardspend = 0
    totalpreminus_miletotalprice = 0
    totalpreminus_spendmembership = 0
    totalpreminus_allatdiscountprice = 0

    orgidx           = 0
    orgsubtotalprice = 0
    currjupsusum = 0
    AsDetailExists = false
    MinusMiletotalprice = 0

	'// =======================================================================
	if (CheckOverMinusOrderItemnoExist(id, orderserial) = True) then

        CheckNAddMinusOrder = False
        ErrStr = "���̳ʽ� �ֹ� ��ǰ �հ谡 �� ��ǰ��Ÿ Ŭ �� �ֽ��ϴ�.\n(�ߺ� �����Ǿ��� �� �ֽ��ϴ�. ����� ��� �˴ϴ�.)"

        exit function

	end if

	if CheckCSFinished(id) = True then
        CheckNAddMinusOrder = False
        ErrStr = "�̹� �Ϸ�� CS�����Դϴ�."
		exit function
	end if

	if (IsOrderExists(orderserial) = False) then

        CheckNAddMinusOrder = False
        ErrStr = "�� �ֹ����� �������� �ʽ��ϴ�."

        exit function

	end if

	if (IsCSDetailExists(id, orderserial) = False) then

        CheckNAddMinusOrder = False
        ErrStr = "��ǰ �ֹ��� �󼼳����� �����ϴ�. - ������ ���ǿ��"

        exit function

	end if

	'// =======================================================================
	''���ֹ���( + ��ȯ�ֹ�) �ݾ� ��ȸ
	Call GetOrgOrderPriceInfo(orderserial, orgsubtotalprice, orgsumPaymentEtc, orgtencardspend, orgmiletotalprice, orgspendmembership, orgallatdiscountprice, orgdepositsum, orggiftcardsum, orgpercentcouponsum)

	'// =======================================================================
    MinusOrderSerial =  AddMinusOrder(id, orderserial)

    if (MinusOrderSerial="") then

        CheckNAddMinusOrder = false
        ErrStr = "��ǰ �ֹ��� ���� ���� - �ݵ��! ������ ���ǿ��."

        exit function

    end if

	Call AddminusOrderLink(id, MinusOrderserial)

	'// =======================================================================
	''���ֹ���( + ��ȯ�ֹ�) �� ���� ���̳ʽ� �ֹ� �ݾ� �հ�
	Call GetMinusOrderPriceInfo(orderserial, totalpreminus_subtotalprice, totalpreminus_sumPaymentEtc, totalpreminus_tencardspend, totalpreminus_miletotalprice, totalpreminus_spendmembership, totalpreminus_allatdiscountprice, totalpreminus_depositsum, totalpreminus_giftcardsum, totalpreminus_percentcouponsum)

	'// =======================================================================
    if (totalpreminus_subtotalprice > orgsubtotalprice) then
        CheckNAddMinusOrder = false
        ErrStr = "���̳ʽ��ֹ� �����ݾ��հ谡 ���ֹ����� Ů�ϴ�.(�ߺ� ���� : 101)"
        exit function
    end if

    if (totalpreminus_sumPaymentEtc > orgsumPaymentEtc) then
        CheckNAddMinusOrder = false
        ErrStr = "���̳ʽ��ֹ� ���������ݾ��հ谡 ���ֹ����� Ů�ϴ�.(�ߺ� ȯ�� : 102)"
        exit function
    end if

    if (totalpreminus_tencardspend > orgtencardspend) then
        CheckNAddMinusOrder = false
        ErrStr = "���̳ʽ��ֹ� �����հ谡 ���ֹ����� Ů�ϴ�.(�ߺ� ȯ�� : 103)"
        exit function
    end if

    if (totalpreminus_miletotalprice > orgmiletotalprice) then
        CheckNAddMinusOrder = false
        ErrStr = "���̳ʽ��ֹ� ���ϸ����հ谡 ���ֹ����� Ů�ϴ�.(�ߺ� ȯ�� : 104)"
        exit function
    end if

    if (totalpreminus_spendmembership > orgspendmembership) then
        CheckNAddMinusOrder = false
        ErrStr = "���̳ʽ��ֹ� �ɹ���ī���հ谡 ���ֹ����� Ů�ϴ�.(�ߺ� ȯ�� : 105)"
        exit function
    end if

    if (totalpreminus_allatdiscountprice > orgallatdiscountprice) then
        CheckNAddMinusOrder = false
        ErrStr = "���̳ʽ� ��Ÿ�����հ谡 ���ֹ����� Ů�ϴ�.(�ߺ� ���� : 106)"
        exit function
    end if

	'// =======================================================================
    ''���ֹ��� ��ȸ
    sqlStr = " select userid, sitename "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master"
    end if
    sqlStr = sqlStr  + " where orderserial='" + orderserial + "'"
    sqlStr = sqlStr  + " and cancelyn='N'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        userid              = rsget("userid")
        sitename            = rsget("sitename")
    end if
    rsget.Close

    sqlStr = " select IsNULL(miletotalprice,0) as miletotalprice from [db_order].[dbo].tbl_order_master"
    sqlStr = sqlStr + " where orderserial='" + MinusOrderSerial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
       ''��ǰ ȯ�� ���ϸ���
       MinusMiletotalprice = rsget("miletotalprice")
    end if
    rsget.Close

    sqlStr = " select IsNULL(realPayedsum,0) as realPayedsum from [db_order].[dbo].tbl_order_PaymentEtc "
    sqlStr = sqlStr + " where orderserial='" + MinusOrderSerial + "' and acctdiv = '200' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	MinusDepositPrice = 0
    if Not rsget.Eof then
       ''��ǰ ȯ�� ��ġ��
       MinusDepositPrice = rsget("realPayedsum")
    end if
    rsget.Close

    sqlStr = " select IsNULL(realPayedsum,0) as realPayedsum from [db_order].[dbo].tbl_order_PaymentEtc "
    sqlStr = sqlStr + " where orderserial='" + MinusOrderSerial + "' and acctdiv = '900' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	MinusGiftCardPrice = 0
    if Not rsget.Eof then
       ''��ǰ ȯ�� Giftī��
       MinusGiftCardPrice = rsget("realPayedsum")
    end if
    rsget.Close

    ''���ϸ���/��ġ�� ����
    '���ϸ����� ���Ż�ǰ��ǰ���� �׻� ������ �ʿ�������, ��ġ��,Giftī��� ����� �ִ� ��츸 �����Ѵ�.
    if (userid<>"") and (sitename="10x10") then

        ''��ǰ ȯ�� ���ϸ��� �߰�
        if (MinusMiletotalprice<>0) then
            sqlStr = "insert into [db_user].[dbo].tbl_mileagelog(userid,mileage,jukyocd,jukyo,orderserial)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusMiletotalprice)) + ",'02','��ǰȯ��','" + MinusOrderSerial + "')"

			dbget.Execute  sqlStr
        end if

        Call updateUserMileage(userid)

        ''��ǰ ȯ�� ��ġ�� �߰�
        if (MinusDepositPrice<>0) then
            sqlStr = "insert into [db_user].[dbo].tbl_depositlog(userid,deposit,jukyocd,jukyo,orderserial)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusDepositPrice)) + ",'100','��ǰȯ��','" + MinusOrderSerial + "')"

			dbget.Execute  sqlStr

			Call updateUserDeposit(userid)
        end if

        ''��ǰ ȯ�� Giftī�� �߰�
        if (MinusGiftCardPrice<>0) then
            sqlStr = "insert into [db_user].[dbo].tbl_giftcard_log(userid,useCash,jukyocd,jukyo,orderserial, reguserid)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusGiftCardPrice)) + ",'200','��ǰȯ��','" + MinusOrderSerial + "','" + CStr(session("ssbctid")) + "')"

			dbget.Execute  sqlStr

			Call updateUserGiftCard(userid)
        end if

    end if

    CheckNAddMinusOrder = true
end function

function CheckNAddMinusOrder_3PL(id, orderserial, reguserid,byref MinusOrderserial, byref ErrStr)
    dim sqlStr
    dim currjupsusum, orgidx
    dim userid, sitename
    dim AsDetailExists
    dim MinusMiletotalprice, MinusDepositPrice, MinusGiftCardPrice
    ''dim totalpreminussum
    ''totalpreminussum = 0

    dim orgsubtotalprice, orgsumPaymentEtc, orgtencardspend, orgmiletotalprice, orgspendmembership, orgallatdiscountprice, orgdepositsum, orggiftcardsum, orgpercentcouponsum
    dim totalpreminus_subtotalprice, totalpreminus_sumPaymentEtc, totalpreminus_tencardspend, totalpreminus_miletotalprice, totalpreminus_spendmembership, totalpreminus_allatdiscountprice, totalpreminus_depositsum, totalpreminus_giftcardsum, totalpreminus_percentcouponsum

    totalpreminus_subtotalprice = 0
    totalpreminus_sumPaymentEtc = 0
    totalpreminus_tencardspend = 0
    totalpreminus_miletotalprice = 0
    totalpreminus_spendmembership = 0
    totalpreminus_allatdiscountprice = 0

    orgidx           = 0
    orgsubtotalprice = 0
    currjupsusum = 0
    AsDetailExists = false
    MinusMiletotalprice = 0

	'// =======================================================================
	if (CheckOverMinusOrderItemnoExist_3PL(id, orderserial) = True) then

        CheckNAddMinusOrder_3PL = False
        ErrStr = "���̳ʽ� �ֹ� ��ǰ �հ谡 �� ��ǰ��Ÿ Ŭ �� �ֽ��ϴ�.\n(�ߺ� �����Ǿ��� �� �ֽ��ϴ�. ����� ��� �˴ϴ�.)"

        exit function

	end if

	if CheckCSFinished_3PL(id) = True then
        CheckNAddMinusOrder_3PL = False
        ErrStr = "�̹� �Ϸ�� CS�����Դϴ�."
		exit function
	end if

	if (IsOrderExists_3PL(orderserial) = False) then

        CheckNAddMinusOrder_3PL = False
        ErrStr = "�� �ֹ����� �������� �ʽ��ϴ�."

        exit function

	end if

	if (IsCSDetailExists_3PL(id, orderserial) = False) then

        CheckNAddMinusOrder_3PL = False
        ErrStr = "��ǰ �ֹ��� �󼼳����� �����ϴ�. - ������ ���ǿ��"

        exit function

	end if

	'// =======================================================================
    MinusOrderSerial =  AddMinusOrder_3PL(id, orderserial)

    if (MinusOrderSerial="") then

        CheckNAddMinusOrder_3PL = false
        ErrStr = "��ǰ �ֹ��� ���� ���� - �ݵ��! ������ ���ǿ��."

        exit function

    end if

	Call AddminusOrderLink_3PL(id, MinusOrderserial)

    CheckNAddMinusOrder_3PL = true
end function

function AddMinusOrder(id, orderserial)
    dim sqlStr
    dim iid
    dim rndjumunno
    dim neworderserial

    dim subtotalprice, miletotalprice, tencardspend, spendmembership, allatdiscountprice
    sqlStr = " select subtotalprice, IsNULL(miletotalprice,0) as miletotalprice,"
    sqlStr = sqlStr + " IsNULL(tencardspend,0) as tencardspend, IsNULL(spendmembership,0) as spendmembership,"
    sqlStr = sqlStr + " IsNULL(allatdiscountprice,0) as allatdiscountprice "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        subtotalprice       = rsget("subtotalprice")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        spendmembership     = rsget("spendmembership")
        allatdiscountprice  = rsget("allatdiscountprice")
    rsget.close

    dim refundmileagesum, refundcouponsum, allatsubtractsum, refunditemcostsum
    dim refundbeasongpay, refunddeliverypay, refundadjustpay, canceltotal
    dim refundgiftcardsum, refunddepositsum
	dim IsCsExists : IsCsExists = False

    ''���� ���ϸ��� ȯ�� ���
    sqlStr = " select r.*, IsNull(r.refundgiftcardsum, 0) as refundgiftcardsum, IsNull(r.refunddepositsum, 0) as refunddepositsum from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"

    'ȯ�Ҿ����� �����ص� ���ϸ��� ��ġ�� ���� ȯ���Ѵ�.
    'sqlStr = sqlStr + " and r.returnmethod<>'R000'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
		IsCsExists = True
        refundrequire       = rsget("refundrequire")
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")
        allatsubtractsum    = rsget("allatsubtractsum")

        refunditemcostsum   = rsget("refunditemcostsum")

        refundbeasongpay    = rsget("refundbeasongpay")
        refunddeliverypay   = rsget("refunddeliverypay")
        refundadjustpay     = rsget("refundadjustpay")
        canceltotal         = rsget("canceltotal")

        refundgiftcardsum   = rsget("refundgiftcardsum")
        refunddepositsum    = rsget("refunddepositsum")
    else
        refundrequire       = 0
        refundmileagesum    = 0
        refundcouponsum     = 0
        allatsubtractsum    = 0

        refunditemcostsum   = 0

        refundbeasongpay    = 0
        refunddeliverypay   = 0
        refundadjustpay     = 0
        canceltotal         = 0

        refundgiftcardsum	= 0
        refunddepositsum	= 0
    end if
    rsget.Close

    ''ȯ�� �� ������ ���� �� ����
    if (subtotalprice=refundrequire) and (Not IsCsExists) then
        refundmileagesum    = miletotalprice * -1
        refundcouponsum     = tencardspend * -1
        allatsubtractsum    = allatdiscountprice * -1

		sqlStr = " select e.acctdiv, e.realPayedsum "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and e.orderserial = '" & orderserial & "' "
		sqlStr = sqlStr + " 	and e.acctdiv in ('200', '900') "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	    if Not rsget.Eof then
	    	do until rsget.eof

		        if (CStr(rsget("acctdiv")) = "200") then
		        	refunddepositsum = rsget("realPayedsum") * -1
		        end if

		        if (CStr(rsget("acctdiv")) = "900") then
		        	refundgiftcardsum = rsget("realPayedsum") * -1
		        end if
                 rsget.MoveNext  '''�̺κ� �����־� timeOUT
			loop
	    end if
	    rsget.Close
    end if

	Randomize
	rndjumunno = CLng(Rnd * 100000) + 1
	rndjumunno = CStr(rndjumunno)

	sqlStr = "select * from [db_order].[dbo].tbl_order_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("orderserial") = rndjumunno
	rsget("jumundiv") = "9"
	rsget("userid") = ""
	rsget("accountname") = ""
	rsget("accountdiv") = "7"
	rsget("sitename") = ""
	rsget.update
	    iid = rsget("idx")
	rsget.close

	neworderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
	neworderserial = neworderserial & Format00(5,Right(CStr(iid),5))

    sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
    sqlStr = sqlStr + " set orderserial='" + neworderserial + "'" & vbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iid)

    dbget.Execute sqlStr

    sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
	sqlStr = sqlStr + " set userid=O.userid" & vbCrlf
	sqlStr = sqlStr + " ,accountname=O.accountname" & vbCrlf
	sqlStr = sqlStr + " ,accountdiv=O.accountdiv" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdiv='8'" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,regdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,beadaldiv=O.beadaldiv" & vbCrlf
	sqlStr = sqlStr + " ,beadaldate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,buyname=O.buyname" & vbCrlf
	sqlStr = sqlStr + " ,buyphone=O.buyphone" & vbCrlf
	sqlStr = sqlStr + " ,buyhp=O.buyhp" & vbCrlf
	sqlStr = sqlStr + " ,buyemail=O.buyemail" & vbCrlf
	sqlStr = sqlStr + " ,reqname=O.reqname" & vbCrlf
	sqlStr = sqlStr + " ,reqzipcode=O.reqzipcode" & vbCrlf
	sqlStr = sqlStr + " ,reqaddress=O.reqaddress" & vbCrlf
	sqlStr = sqlStr + " ,reqphone=O.reqphone" & vbCrlf
	sqlStr = sqlStr + " ,reqhp=O.reqhp" & vbCrlf
	sqlStr = sqlStr + " ,comment='���ֹ���ȣ:" + orderserial +"'" & vbCrlf
	sqlStr = sqlStr + " ,linkorderserial=O.orderserial" & vbCrlf
	sqlStr = sqlStr + " ,deliverno=''" & vbCrlf
	sqlStr = sqlStr + " ,sitename=O.sitename" & vbCrlf
	sqlStr = sqlStr + " ,discountrate=O.discountrate" & vbCrlf
	sqlStr = sqlStr + " ,subtotalprice=O.subtotalprice" & vbCrlf
	sqlStr = sqlStr + " ,miletotalprice=" & CStr(refundmileagesum) & vbCrlf
	sqlStr = sqlStr + " ,tencardspend=" & CStr(refundcouponsum) & vbCrlf
	sqlStr = sqlStr + " ,spendmembership=0" & vbCrlf
	sqlStr = sqlStr + " ,allatdiscountprice=" & CStr(allatsubtractsum) & vbCrlf
	sqlStr = sqlStr + " ,rduserid=O.rduserid" & vbCrlf
	sqlStr = sqlStr + " ,sentenceidx=O.sentenceidx" & vbCrlf
	sqlStr = sqlStr + " ,reqzipaddr=O.reqzipaddr" & vbCrlf
	sqlStr = sqlStr + " ,rdsite=O.rdsite" & vbCrlf
	sqlStr = sqlStr + " ,subtotalpriceCouponNotApplied=IsNull(O.subtotalpriceCouponNotApplied,0)" & vbCrlf
	sqlStr = sqlStr + " ,sumPaymentEtc=" & CStr(refundgiftcardsum + refunddepositsum) & " " & vbCrlf
	sqlStr = sqlStr + " ,userlevel=O.userlevel" & vbCrlf    ''20121219 �߰�
	sqlStr = sqlStr + " ,pggubun=O.pggubun" & vbCrlf    	''2015-08-25 �߰�
	if (CStr(refundcouponsum)<>"") then
        sqlStr = sqlStr + " ,bCpnIdx=O.bCpnIdx" & vbCrlf    ''20121129 �߰�
    end if
	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from (select top 1 * from [db_log].[dbo].tbl_old_order_master_2003 where orderserial='" + orderserial + "') O" & vbCrlf
	else
	    sqlStr = sqlStr + " from (select top 1 * from [db_order].[dbo].tbl_order_master where orderserial='" + orderserial + "') O" & vbCrlf
	end if
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.idx=" + CStr(iid)

	dbget.Execute sqlStr

	''����ۺ� ȯ�� �������
'	if (refundbeasongpay<>0) then
'	    sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
'	    sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
'	    sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
'        sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,currstate,beasongdate,upcheconfirmdate)" & vbCrlf
'        sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
'        sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemid" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemno*-1" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemvat" & vbCrlf
'        sqlStr = sqlStr + " ,d.mileage" & vbCrlf
'        sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemname" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
'        sqlStr = sqlStr + " ,d.makerid" & vbCrlf
'        sqlStr = sqlStr + " ,d.buycash" & vbCrlf
'        sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
'        sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
'        sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
'        sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
'        sqlStr = sqlStr + " ,'7'" & vbCrlf
'        sqlStr = sqlStr + " ,getdate()" & vbCrlf
'        sqlStr = sqlStr + " ,getdate()" & vbCrlf
'        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d" & vbCrlf
'        sqlStr = sqlStr + " where d.orderserial='" & orderserial & "'"  & vbCrlf
'        sqlStr = sqlStr + " and d.itemid=0" & vbCrlf
'        sqlStr = sqlStr + " and d.cancelyn<>'Y'"
'
'        dbget.Execute sqlStr
'	end if

	''���/��ǰ ��ǰ �󼼳���
	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
    sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
    sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, currstate,beasongdate,upcheconfirmdate,itemcouponidx, bonuscouponidx, etcDiscount)" & vbCrlf
    sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
    sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
    sqlStr = sqlStr + " ,d.itemid" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
    sqlStr = sqlStr + " ,J.confirmitemno*-1" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
    sqlStr = sqlStr + " ,d.itemvat" & vbCrlf
    sqlStr = sqlStr + " ,d.mileage" & vbCrlf
    sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
    sqlStr = sqlStr + " ,d.itemname" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
    sqlStr = sqlStr + " ,d.makerid" & vbCrlf
    sqlStr = sqlStr + " ,d.buycash" & vbCrlf
    sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
    sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
    sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.omwdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.odlvType" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.orgitemcost,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.itemcostCouponNotApplied,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.buycashCouponNotApplied,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.plussaleDiscount,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.specialShopDiscount,0)" & vbCrlf
    sqlStr = sqlStr + " ,'7'" & vbCrlf
    sqlStr = sqlStr + " ,getdate()" & vbCrlf
    sqlStr = sqlStr + " ,getdate()" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcouponidx, d.bonuscouponidx, d.etcDiscount" & vbCrlf
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_detail J" & vbCrlf
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " ,[db_log].[dbo].tbl_old_order_detail_2003 d" & vbCrlf
    else
        sqlStr = sqlStr + " ,[db_order].[dbo].tbl_order_detail d" & vbCrlf
    end if
    sqlStr = sqlStr + " where J.masterid=" & CStr(id)
    sqlStr = sqlStr + " and d.orderserial='" & orderserial & "'"  & vbCrlf
    sqlStr = sqlStr + " and J.orderdetailidx=d.idx"  & vbCrlf
   sqlStr = sqlStr + " and J.confirmitemno<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    dbget.Execute sqlStr

	if (refunddeliverypay <> 0) and (Left(now, 10) >= "2014-01-01") then
	''if (refunddeliverypay <> 0) and (Left(now, 10) >= "2013-12-01") then
		'��ǰ��ۺ�
		sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
		sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
		sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
		sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, currstate,beasongdate,upcheconfirmdate,itemcouponidx, bonuscouponidx)" & vbCrlf
		sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
		sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
		sqlStr = sqlStr + " ,0" & vbCrlf
		if (id = 5026974) then
			sqlStr = sqlStr + " ,'5001'" & vbCrlf
		else
			sqlStr = sqlStr + " ,'5000'" & vbCrlf
		end if
		sqlStr = sqlStr + " ,1" & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , Round(((1.0 * " + CStr(refunddeliverypay * -1) + ") / 11.0), 0) " & vbCrlf
		sqlStr = sqlStr + " , 0 " & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , '��ǰ��ۺ�' " & vbCrlf
		sqlStr = sqlStr + " , (case when IsNull(a.requireupche, 'Y') = 'Y' then '��ü����' else '' end) " & vbCrlf
		sqlStr = sqlStr + " , IsNull(a.makerid, '') " & vbCrlf
		sqlStr = sqlStr + " , (case when IsNull(a.requireupche, 'Y') = 'Y' and IsNull(a.makerid, '') <> '10x10logistics' then " + CStr(refunddeliverypay * -1) + " else 0 end) " & vbCrlf
		sqlStr = sqlStr + " , 'Y' " & vbCrlf
		sqlStr = sqlStr + " , NULL " & vbCrlf
		sqlStr = sqlStr + " , 'N' " & vbCrlf
		sqlStr = sqlStr + " , '01' " & vbCrlf
		sqlStr = sqlStr + " , NULL " & vbCrlf
		sqlStr = sqlStr + " , NULL " & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , " + CStr(refunddeliverypay * -1) + " " & vbCrlf
		sqlStr = sqlStr + " , (case when IsNull(a.requireupche, 'Y') = 'Y' and IsNull(a.makerid, '') <> '10x10logistics' then " + CStr(refunddeliverypay * -1) + " else 0 end) " & vbCrlf
		sqlStr = sqlStr + " , 0 " & vbCrlf
		sqlStr = sqlStr + " , 0 " & vbCrlf
		sqlStr = sqlStr + " ,'7'" & vbCrlf
		sqlStr = sqlStr + " ,getdate()" & vbCrlf
		sqlStr = sqlStr + " ,getdate()" & vbCrlf
		sqlStr = sqlStr + " ,NULL, NULL " & vbCrlf

		sqlStr = sqlStr + " from " & vbCrlf
		if (GC_IsOLDOrder) then
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m " & vbCrlf
		else
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m " & vbCrlf
		end if

		sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_list a "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	m.orderserial = a.orderserial "
		sqlStr = sqlStr + " join db_cs.dbo.tbl_as_refund_info r "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	a.id = r.asid "

		sqlStr = sqlStr + " where a.id = " & CStr(id)
		dbget.Execute sqlStr

		sqlStr = " update r "
		sqlStr = sqlStr + " set r.isRefundDeliveryPayAddedToOrder = 'Y' "
		sqlStr = sqlStr + " from " & vbCrlf
		sqlStr = sqlStr + " db_cs.dbo.tbl_new_as_list a "
		sqlStr = sqlStr + " join db_cs.dbo.tbl_as_refund_info r "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	a.id = r.asid "
		sqlStr = sqlStr + " where a.id = " & CStr(id)
		dbget.Execute sqlStr

	end if

	if (refundgiftcardsum <> 0) then
		Call InsertEtcPaymentOne(neworderserial, "900", refundgiftcardsum)
	end if

	if (refunddepositsum <> 0) then
		Call InsertEtcPaymentOne(neworderserial, "200", refunddepositsum)
	end if

    ''�ֹ��ݾ� ����
    call recalcuOrderMaster(neworderserial)

    ''���������� - ���������� ���� �ȵ�
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_minusOrder '" & neworderserial & "'"
    dbget.Execute sqlStr

    AddMinusOrder    = neworderserial
end function

function GetBankName(acctno)
	select case acctno
		case "11"
			GetBankName = "����"
		case "06"
			GetBankName = "����"
		case "20"
			GetBankName = "�츮"
		case "26"
			GetBankName = "����"
		case "81"
			GetBankName = "�ϳ�"
		case "03"
			GetBankName = "���"
		case "39"
			GetBankName = "�泲"
		case "32"
			GetBankName = "�λ�"
		case "71"
			GetBankName = "��ü��"
		case "07"
			GetBankName = "����"
		case else
			GetBankName = acctno
	end select
end function

function AddPaymentOrder(id, orderserial, additempay, addbeasongpay, payordertype, accountdiv, accountname, requiremakerid)
    dim sqlStr
    dim iid
    dim rndjumunno
    dim neworderserial

	Randomize
	rndjumunno = CLng(Rnd * 100000) + 1
	rndjumunno = CStr(rndjumunno)

	sqlStr = "select * from [db_order].[dbo].tbl_order_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("orderserial") = rndjumunno
	rsget("jumundiv") = "1"
	rsget("userid") = ""
	rsget("accountname") = ""
	rsget("accountdiv") = "7"
	rsget("sitename") = ""
	rsget.update
	    iid = rsget("idx")
	rsget.close

	neworderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
	neworderserial = neworderserial & Format00(5,Right(CStr(iid),5))

    sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
    sqlStr = sqlStr + " set orderserial='" + neworderserial + "'" & vbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iid)
    dbget.Execute sqlStr

    sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
	sqlStr = sqlStr + " set userid=O.userid" & vbCrlf
	sqlStr = sqlStr + " ,accountname='" & accountname & "'" & vbCrlf			'// �Ա��ڸ�
	sqlStr = sqlStr + " ,accountno=''" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdiv='0'" & vbCrlf
	sqlStr = sqlStr + " ,accountdiv=" & accountdiv & vbCrlf
	sqlStr = sqlStr + " ,regdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,beadaldiv=1" & vbCrlf									'// ���߰������� �ڻ���ֹ����� ����
	sqlStr = sqlStr + " ,buyname=O.buyname" & vbCrlf
	sqlStr = sqlStr + " ,buyphone=O.buyphone" & vbCrlf
	sqlStr = sqlStr + " ,buyhp=O.buyhp" & vbCrlf
	sqlStr = sqlStr + " ,buyemail=O.buyemail" & vbCrlf
	sqlStr = sqlStr + " ,reqname=O.reqname" & vbCrlf
	sqlStr = sqlStr + " ,reqzipcode=O.reqzipcode" & vbCrlf
	sqlStr = sqlStr + " ,reqaddress=O.reqaddress" & vbCrlf
	sqlStr = sqlStr + " ,reqphone=O.reqphone" & vbCrlf
	sqlStr = sqlStr + " ,reqhp=O.reqhp" & vbCrlf
	sqlStr = sqlStr + " ,reqzipaddr=O.reqzipaddr" & vbCrlf
	sqlStr = sqlStr + " ,comment='���ֹ���ȣ:" + orderserial +"'" & vbCrlf
	sqlStr = sqlStr + " ,linkorderserial=O.orderserial" & vbCrlf
	sqlStr = sqlStr + " ,deliverno=''" & vbCrlf
	sqlStr = sqlStr + " ,sitename='10x10'" & vbCrlf
	sqlStr = sqlStr + " ,discountrate=O.discountrate" & vbCrlf
	sqlStr = sqlStr + " ,subtotalprice=0" & vbCrlf
	sqlStr = sqlStr + " ,miletotalprice=0" & vbCrlf
	sqlStr = sqlStr + " ,tencardspend=0" & vbCrlf
	sqlStr = sqlStr + " ,spendmembership=0" & vbCrlf
	sqlStr = sqlStr + " ,allatdiscountprice=0" & vbCrlf
	sqlStr = sqlStr + " ,rduserid=O.rduserid" & vbCrlf
	sqlStr = sqlStr + " ,rdsite=O.rdsite" & vbCrlf
	sqlStr = sqlStr + " ,subtotalpriceCouponNotApplied=0" & vbCrlf
	sqlStr = sqlStr + " ,sumPaymentEtc=0" & vbCrlf
	sqlStr = sqlStr + " ,userlevel=O.userlevel" & vbCrlf    ''20121219 �߰�
	sqlStr = sqlStr + " ,pggubun=''" & vbCrlf    	''2015-08-25 �߰�
	if (payordertype = "A") then
		sqlStr = sqlStr + " ,baljudate=getdate()" & vbCrlf			'// ������̹Ƿ� ���־ȵǵ���
	end if
	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from (select top 1 * from [db_log].[dbo].tbl_old_order_master_2003 where orderserial='" + orderserial + "') O" & vbCrlf
	else
	    sqlStr = sqlStr + " from (select top 1 * from [db_order].[dbo].tbl_order_master where orderserial='" + orderserial + "') O" & vbCrlf
	end if
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.idx=" + CStr(iid)

	dbget.Execute sqlStr

	'// �߰���ۺ�
	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
	sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
	sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, currstate, itemcouponidx, bonuscouponidx)" & vbCrlf
	sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
	sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
	sqlStr = sqlStr + " ,0" & vbCrlf
	if (requiremakerid = "") then
		sqlStr = sqlStr + " ,'0101'" & vbCrlf
	else
		sqlStr = sqlStr + " ,'9001'" & vbCrlf
	end if
	sqlStr = sqlStr + " ,1" & vbCrlf
	sqlStr = sqlStr + " , " + CStr(addbeasongpay) + " " & vbCrlf
	sqlStr = sqlStr + " , Round(((1.0 * " + CStr(addbeasongpay) + ") / 11.0), 0) " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " , " + CStr(addbeasongpay) + " " & vbCrlf
	sqlStr = sqlStr + " , '���߰�����' " & vbCrlf
	sqlStr = sqlStr + " , '' " & vbCrlf
	sqlStr = sqlStr + " , '" & requiremakerid & "' " & vbCrlf
	sqlStr = sqlStr + " , (case when IsNull(a.requireupche, 'Y') = 'Y' then " + CStr(addbeasongpay) + " else 0 end) " & vbCrlf
	sqlStr = sqlStr + " , 'Y' " & vbCrlf
	sqlStr = sqlStr + " , NULL " & vbCrlf
	sqlStr = sqlStr + " , 'N' " & vbCrlf
	sqlStr = sqlStr + " , '01' " & vbCrlf
	sqlStr = sqlStr + " , NULL " & vbCrlf
	sqlStr = sqlStr + " , NULL " & vbCrlf
	sqlStr = sqlStr + " , " + CStr(addbeasongpay) + " " & vbCrlf
	sqlStr = sqlStr + " , " + CStr(addbeasongpay) + " " & vbCrlf
	sqlStr = sqlStr + " , (case when IsNull(a.requireupche, 'Y') = 'Y' then " + CStr(addbeasongpay) + " else 0 end) " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " ,'0'" & vbCrlf
	sqlStr = sqlStr + " ,NULL, NULL " & vbCrlf

	sqlStr = sqlStr + " from " & vbCrlf
	if (GC_IsOLDOrder) then
		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m " & vbCrlf
	else
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m " & vbCrlf
	end if

	sqlStr = sqlStr + " join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + "		m.orderserial = a.orderserial "
	sqlStr = sqlStr + " where a.id = " & CStr(id)
	dbget.Execute sqlStr

	if (payordertype = "A") or (payordertype = "N") then
		'//��������
        '// A : ��������, N : �ֹ�����
		sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
		sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
		sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
		sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType, itemcouponidx, bonuscouponidx, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, etcDiscount, currstate,beasongdate,upcheconfirmdate)" & vbCrlf
		sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
		sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
		sqlStr = sqlStr + " ,d.itemid" & vbCrlf
		sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
		sqlStr = sqlStr + " ,J.confirmitemno" & vbCrlf
		sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
		sqlStr = sqlStr + " ,d.itemvat" & vbCrlf
		sqlStr = sqlStr + " ,d.mileage" & vbCrlf
		sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
		sqlStr = sqlStr + " ,d.itemname" & vbCrlf
		sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
		sqlStr = sqlStr + " ,d.makerid" & vbCrlf
		sqlStr = sqlStr + " ,d.buycash" & vbCrlf
		sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
		sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
		sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
		sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
		sqlStr = sqlStr + " ,d.omwdiv" & vbCrlf
		sqlStr = sqlStr + " ,d.odlvType" & vbCrlf
		sqlStr = sqlStr + " ,d.itemcouponidx" & vbCrlf
		sqlStr = sqlStr + " ,d.bonuscouponidx" & vbCrlf
		sqlStr = sqlStr + " ,IsNull(d.orgitemcost,0)" & vbCrlf
		sqlStr = sqlStr + " ,IsNull(d.itemcostCouponNotApplied,0)" & vbCrlf
		sqlStr = sqlStr + " ,IsNull(d.buycashCouponNotApplied,0)" & vbCrlf
		sqlStr = sqlStr + " ,IsNull(d.plussaleDiscount,0)" & vbCrlf
		sqlStr = sqlStr + " ,IsNull(d.specialShopDiscount,0)" & vbCrlf
		sqlStr = sqlStr + " ,IsNull(d.etcDiscount,0)" & vbCrlf
		''sqlStr = sqlStr + " ,'1'" & vbCrlf									'// ����� : 0 �� ��ü�뺸�� �Ѿ�Ƿ� '1' �� ����
        if (payordertype = "A") then
            sqlStr = sqlStr + " ,'1'" & vbCrlf
        else
            sqlStr = sqlStr + " ,'0'" & vbCrlf
        end if
		sqlStr = sqlStr + " ,NULL" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_detail J" & vbCrlf
		if (GC_IsOLDOrder) then
			sqlStr = sqlStr + " ,[db_log].[dbo].tbl_old_order_detail_2003 d" & vbCrlf
		else
			sqlStr = sqlStr + " ,[db_order].[dbo].tbl_order_detail d" & vbCrlf
		end if
		sqlStr = sqlStr + " where J.masterid=" & CStr(id)
		sqlStr = sqlStr + " and d.orderserial='" & orderserial & "'"  & vbCrlf
		sqlStr = sqlStr + " and J.orderdetailidx=d.idx"  & vbCrlf
		sqlStr = sqlStr + " and J.confirmitemno<>0"
		''sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and J.orderdetailidx is not null "
		dbget.Execute sqlStr

		sqlStr = " update m "
		sqlStr = sqlStr + " set m.tencardspend = T.tencardspend, m.allatdiscountprice = T.allatdiscountprice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
		sqlStr = sqlStr + " 	join ( "
		sqlStr = sqlStr + " 		select "
		sqlStr = sqlStr + " 			d.orderserial "
		sqlStr = sqlStr + " 			, sum((d.itemcost - d.reducedPrice - IsNull(d.etcDiscount,0)) * d.itemno) as tencardspend "
		sqlStr = sqlStr + " 			, sum(IsNull(d.etcDiscount,0) * d.itemno) as allatdiscountprice "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 		[db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " 		where orderserial = '" & neworderserial & "' "
		sqlStr = sqlStr + " 		and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 		group by d.orderserial "
		sqlStr = sqlStr + " 	) T on m.orderserial = T.orderserial "
		dbget.Execute sqlStr
	end if

	Call InsertEtcPaymentOne(neworderserial, accountdiv, additempay + addbeasongpay)

    ''�ֹ��ݾ� ����
    call recalcuOrderMaster(neworderserial)

	AddPaymentOrder = neworderserial

end function

function SetPayOrderserial(asid, payorderserial)
	dim sqlStr

    sqlStr = " update [db_cs].[dbo].[tbl_as_customer_addbeasongpay_info] "
    sqlStr = sqlStr + " set payorderserial = " & CStr(payorderserial) & " "
    sqlStr = sqlStr + " where asid = " & CStr(asid) & " "
    dbget.Execute sqlStr
end function

function AddMinusOrder_3PL(id, orderserial)
    dim sqlStr
    dim iid
    dim rndjumunno
    dim neworderserial

    dim subtotalprice, miletotalprice, tencardspend, spendmembership, allatdiscountprice

    dim refundmileagesum, refundcouponsum, allatsubtractsum, refunditemcostsum
    dim refundbeasongpay, refunddeliverypay, refundadjustpay, canceltotal
    dim refundgiftcardsum, refunddepositsum
	dim IsCsExists : IsCsExists = False

	dim uuid_obj, uuid, temporderserial
	set uuid_obj = Server.CreateObject("Scriptlet.Typelib")
	uuid = uuid_obj.guid
	set uuid_obj = Nothing
	temporderserial = Mid(Replace(uuid, "-", ""), 2, 32)

	Randomize
	rndjumunno = CLng(Rnd * 100000) + 1
	rndjumunno = CStr(rndjumunno)

	sqlStr = "select * from [db_threepl].[dbo].[tbl_tpl_orderMaster] where 1=0"
	rsget.Open sqlStr,dbget_TPL,1,3
	rsget.AddNew
	rsget("tplcompanyid") = "undefined"
	rsget("orderserial") = temporderserial
	rsget("jumundiv") = "9"
	rsget("sitename") = ""
	rsget("totalsum") = 0
	rsget("subtotalprice") = 0
	rsget("ipkumdiv") = 8
	rsget.update
	    iid = rsget("idx")
	rsget.close

	neworderserial = Replace(Left(Now, 10), "-", "") & Right(Format00(9, iid), 9)

    sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderMaster]" & vbCrlf
    sqlStr = sqlStr + " set orderserial='" + neworderserial + "'" & vbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iid)

    dbget_TPL.Execute sqlStr

    sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderMaster]" & vbCrlf
	sqlStr = sqlStr + " set ipkumdiv='8'" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,tplcompanyid=O.tplcompanyid" & vbCrlf
	sqlStr = sqlStr + " ,regdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,buyname=O.buyname" & vbCrlf
	sqlStr = sqlStr + " ,buyphone=O.buyphone" & vbCrlf
	sqlStr = sqlStr + " ,buyhp=O.buyhp" & vbCrlf
	sqlStr = sqlStr + " ,buyemail=O.buyemail" & vbCrlf
	sqlStr = sqlStr + " ,reqname=O.reqname" & vbCrlf
	sqlStr = sqlStr + " ,reqzipcode=O.reqzipcode" & vbCrlf
	sqlStr = sqlStr + " ,reqaddress=O.reqaddress" & vbCrlf
	sqlStr = sqlStr + " ,reqphone=O.reqphone" & vbCrlf
	sqlStr = sqlStr + " ,reqhp=O.reqhp" & vbCrlf
	sqlStr = sqlStr + " ,comment='���ֹ���ȣ:" + orderserial +"'" & vbCrlf
	sqlStr = sqlStr + " ,linkorderserial=O.orderserial" & vbCrlf
	sqlStr = sqlStr + " ,sitename=O.sitename" & vbCrlf
	sqlStr = sqlStr + " ,totalsum=O.totalsum" & vbCrlf
	sqlStr = sqlStr + " ,subtotalprice=O.subtotalprice" & vbCrlf
	sqlStr = sqlStr + " ,reqzipaddr=O.reqzipaddr" & vbCrlf
	sqlStr = sqlStr + " from (select top 1 * from [db_threepl].[dbo].[tbl_tpl_orderMaster] where orderserial='" + orderserial + "') O" & vbCrlf
	sqlStr = sqlStr + " where [db_threepl].[dbo].[tbl_tpl_orderMaster].idx=" + CStr(iid)

	dbget_TPL.Execute sqlStr

	''���/��ǰ ��ǰ �󼼳���
	sqlStr = "insert into [db_threepl].[dbo].[tbl_tpl_orderDetail]"
	sqlStr = sqlStr + " (masteridx, orderserial,itemgubun,itemid,itemoption,itemno," & vbCrlf
    sqlStr = sqlStr + " itemcost,reducedPrice,mileage,itemname," & vbCrlf
    sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude, currstate,beasongdate)" & vbCrlf
    sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
    sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
    sqlStr = sqlStr + " ,d.itemgubun" & vbCrlf
	sqlStr = sqlStr + " ,d.itemid" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
    sqlStr = sqlStr + " ,J.confirmitemno*-1" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
    sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
	sqlStr = sqlStr + " ,d.mileage" & vbCrlf
    sqlStr = sqlStr + " ,d.itemname" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
    sqlStr = sqlStr + " ,d.makerid" & vbCrlf
    sqlStr = sqlStr + " ,d.buycash" & vbCrlf
    sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,'7'" & vbCrlf
    sqlStr = sqlStr + " ,getdate()" & vbCrlf
    sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_as_detail] J" & vbCrlf
    sqlStr = sqlStr + " ,[db_threepl].[dbo].[tbl_tpl_orderDetail] d" & vbCrlf
    sqlStr = sqlStr + " where J.masterid=" & CStr(id)
    sqlStr = sqlStr + " and d.orderserial='" & orderserial & "'"  & vbCrlf
    sqlStr = sqlStr + " and J.orderdetailidx=d.idx"  & vbCrlf
	sqlStr = sqlStr + " and J.confirmitemno<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    dbget_TPL.Execute sqlStr

    ''�ֹ��ݾ� ����
    call recalcuOrderMaster_3PL(neworderserial)

    ''���������� - ���������� ���� �ȵ�
	'// !!! Ʈ������ ������ ����.
    ''sqlStr = " exec [db_summary].[dbo].[usp_TPL_RealtimeStock_minusOrder] '" & neworderserial & "'"
    ''dbget.Execute sqlStr

    AddMinusOrder_3PL    = neworderserial
end function

'// ��ǰ���� �±�ȯ�� �� �߰���ۺ�
function SetCustomerAddBeasongPay(asid, addmethod, addbeasongpay, receiveyn, realbeasongpay)
    dim sqlStr

    if ((addbeasongpay="") or (realbeasongpay="") or (asid="")) then Exit Function

	sqlStr = " IF EXISTS " + VbCrlf
	sqlStr = sqlStr + " 	( " + VbCrlf
	sqlStr = sqlStr + " 		SELECT TOP 1 " + VbCrlf
	sqlStr = sqlStr + " 			asid " + VbCrlf
	sqlStr = sqlStr + " 		FROM " + VbCrlf
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_as_customer_addbeasongpay_info " + VbCrlf
	sqlStr = sqlStr + " 		WHERE " + VbCrlf
	sqlStr = sqlStr + " 			asid = " + CStr(asid) + " " + VbCrlf
	sqlStr = sqlStr + " 	) " + VbCrlf
	sqlStr = sqlStr + " 	BEGIN " + VbCrlf
	sqlStr = sqlStr + " 		UPDATE " + VbCrlf
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_as_customer_addbeasongpay_info " + VbCrlf
	sqlStr = sqlStr + " 		SET " + VbCrlf
	sqlStr = sqlStr + " 			addmethod = '" + CStr(addmethod) + "' " + VbCrlf
	sqlStr = sqlStr + " 			, addbeasongpay = " + CStr(addbeasongpay) + " " + VbCrlf
	sqlStr = sqlStr + " 			, receiveyn = '" + CStr(receiveyn) + "' " + VbCrlf
	sqlStr = sqlStr + " 			, realbeasongpay = " + CStr(realbeasongpay) + " " + VbCrlf
	sqlStr = sqlStr + " 		WHERE " + VbCrlf
	sqlStr = sqlStr + " 			asid = " + CStr(asid) + " " + VbCrlf
	sqlStr = sqlStr + " 	END " + VbCrlf
	sqlStr = sqlStr + " ELSE " + VbCrlf
	sqlStr = sqlStr + " 	BEGIN " + VbCrlf
	sqlStr = sqlStr + " 		INSERT INTO db_cs.dbo.tbl_as_customer_addbeasongpay_info(asid, addmethod, addbeasongpay, receiveyn, realbeasongpay) " + VbCrlf
	sqlStr = sqlStr + " 		VALUES(" + CStr(asid) + ", '" + CStr(addmethod) + "', " + CStr(addbeasongpay) + ", '" + CStr(receiveyn) + "', " + CStr(realbeasongpay) + ") " + VbCrlf
	sqlStr = sqlStr + " 	END " + VbCrlf

    dbget.Execute sqlStr

end function

function SetCustomerAddPay(asid, addmethod, additempay, additembuypay, addbeasongpay, payordertype, receiveyn, realbeasongpay)
    dim sqlStr

    if ((addbeasongpay="") or (realbeasongpay="") or (asid="")) then Exit Function

	sqlStr = " IF EXISTS " + VbCrlf
	sqlStr = sqlStr + " 	( " + VbCrlf
	sqlStr = sqlStr + " 		SELECT TOP 1 " + VbCrlf
	sqlStr = sqlStr + " 			asid " + VbCrlf
	sqlStr = sqlStr + " 		FROM " + VbCrlf
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_as_customer_addbeasongpay_info " + VbCrlf
	sqlStr = sqlStr + " 		WHERE " + VbCrlf
	sqlStr = sqlStr + " 			asid = " + CStr(asid) + " " + VbCrlf
	sqlStr = sqlStr + " 	) " + VbCrlf
	sqlStr = sqlStr + " 	BEGIN " + VbCrlf
	sqlStr = sqlStr + " 		UPDATE " + VbCrlf
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_as_customer_addbeasongpay_info " + VbCrlf
	sqlStr = sqlStr + " 		SET " + VbCrlf
	sqlStr = sqlStr + " 			addmethod = '" + CStr(addmethod) + "' " + VbCrlf
	sqlStr = sqlStr + " 			, additempay = " + CStr(additempay) + " " + VbCrlf
	sqlStr = sqlStr + " 			, additembuypay = " + CStr(additembuypay) + " " + VbCrlf
	sqlStr = sqlStr + " 			, addbeasongpay = " + CStr(addbeasongpay) + " " + VbCrlf
	sqlStr = sqlStr + " 			, payordertype = '" + CStr(payordertype) + "' " + VbCrlf
	sqlStr = sqlStr + " 			, receiveyn = '" + CStr(receiveyn) + "' " + VbCrlf
	sqlStr = sqlStr + " 			, realbeasongpay = " + CStr(realbeasongpay) + " " + VbCrlf
	sqlStr = sqlStr + " 		WHERE " + VbCrlf
	sqlStr = sqlStr + " 			asid = " + CStr(asid) + " " + VbCrlf
	sqlStr = sqlStr + " 	END " + VbCrlf
	sqlStr = sqlStr + " ELSE " + VbCrlf
	sqlStr = sqlStr + " 	BEGIN " + VbCrlf
	sqlStr = sqlStr + " 		INSERT INTO db_cs.dbo.tbl_as_customer_addbeasongpay_info(asid, addmethod, additempay, additembuypay, addbeasongpay, payordertype, receiveyn, realbeasongpay) " + VbCrlf
	sqlStr = sqlStr + " 		VALUES(" + CStr(asid) + ", '" + CStr(addmethod) + "', " + CStr(additempay) + ", " + CStr(additembuypay) + ", " + CStr(addbeasongpay) + ", '" + CStr(payordertype) + "', '" + CStr(receiveyn) + "', " + CStr(realbeasongpay) + ") " + VbCrlf
	sqlStr = sqlStr + " 	END " + VbCrlf

    dbget.Execute sqlStr

end function

function CheckNChulgoPaymentOrder(asid, byref ErrStr)
    dim sqlStr, currstate, deleteyn, payorderserial, payordertype

	sqlStr = " select top 1 a.currstate, a.deleteyn, i.payorderserial, i.payordertype "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_as_customer_addbeasongpay_info] i on a.id = i.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " & asid
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    currstate = ""
	if Not rsget.Eof then
        currstate = rsget("currstate")
        deleteyn = rsget("deleteyn")
        payorderserial = rsget("payorderserial")
        payordertype = rsget("payordertype")
	end if
	rsget.Close

    if currstate = "" then
        ErrStr = "�߸��� �����Դϴ�. : CS�� ����"
    end if

    if currstate = "B007" then
        ErrStr = "�̹� �Ϸ�� CS�����Դϴ�."
    end if

    if deleteyn = "Y" then
        ErrStr = "������ CS�����Դϴ�."
    end if

    if (payordertype <> "A") then
        '// ����� ������ �ƴ�
        exit function
    end if

	sqlStr = " 	select top 1 m.orderserial "
	sqlStr = sqlStr + " 	from "
	sqlStr = sqlStr + " 		[db_order].[dbo].[tbl_order_master] m "
	sqlStr = sqlStr + " 		join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.orderserial = '" & payorderserial & "' "
	sqlStr = sqlStr + " 		and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 		and d.itemid <> 0 "
	sqlStr = sqlStr + " 		and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " 		and m.ipkumdiv > '3' "
    sqlStr = sqlStr + " 		and m.ipkumdiv <= '8' "			'// �̹� ���Ϸ�� ��쵵 ����ó��
    sqlStr = sqlStr + " 		and d.currstate >= '1' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
        '//
    else
        ErrStr = "������ֹ��� �����ϴ�. (�����Ϸ����� �Ǵ� ��һ����Դϴ�.)"
	end if
	rsget.Close

    if ErrStr <> "" then
        exit function
    end if

	sqlStr = " update [db_order].[dbo].tbl_order_detail "
	sqlStr = sqlStr + " set currstate = '7', beasongdate = getdate(), upcheconfirmdate = getdate(), songjangdiv = '99', songjangno = '������ֹ�' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(payorderserial) + "' "
    sqlStr = sqlStr + " and itemid <> 0 "
    sqlStr = sqlStr + " and cancelyn <> 'Y' "
    sqlStr = sqlStr + " and currstate = '1' "
	dbget.Execute sqlStr

	sqlStr = " update [db_order].[dbo].tbl_order_detail "
	sqlStr = sqlStr + " set currstate = '7', beasongdate = getdate() "
	sqlStr = sqlStr + " where orderserial = '" + CStr(payorderserial) + "' "
    sqlStr = sqlStr + " and itemid = 0 "
    sqlStr = sqlStr + " and cancelyn <> 'Y' "
    sqlStr = sqlStr + " and currstate = '0' "
	dbget.Execute sqlStr

	sqlStr = " update [db_order].[dbo].tbl_order_master "
	sqlStr = sqlStr + " set ipkumdiv = '8' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(payorderserial) + "' "
    sqlStr = sqlStr + " and ipkumdiv < '8' "
	dbget.Execute sqlStr

	sqlStr = " insert into db_temp.dbo.tbl_michulgoMile_Recalcu_Que "
	sqlStr = sqlStr + " (userid) "
	sqlStr = sqlStr + " select m.userid "
	sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr + " where m.orderserial='" + CStr(payorderserial) + "' "
	sqlStr = sqlStr + " and m.userid<>'' "
    dbget.Execute sqlStr

end function

'// ===========================================================================
'// ��ǰ���� �±�ȯ ��ȯ�ֹ�
'// ===========================================================================
function CheckAndAddChangeOrder(asid, orderserial, byref ErrStr)

	'// =======================================================================
	if (CheckOverChangeOrderItemnoExist(id, orderserial) = True) then

        CheckAndAddChangeOrder = False
        ErrStr = "���ֹ�( + ��ȯ�ֹ� + ���̳ʽ��ֹ�) �� �ʰ��Ͽ� ȸ���Ǵ� ��ǰ�� �ֽ��ϴ�.\n(�ߺ� �����Ǿ��� �� �ֽ��ϴ�. �ý����� ����)"

        exit function

	end if

	CheckAndAddChangeOrder = AddChangeOrder(asid, orderserial)
end function

'// ===========================================================================
'// ��ǰ���� �±�ȯ ��ȯ�ֹ�(�ֹ����� ���·� ���)
'// ===========================================================================
function CheckAndAddChangeOrderJupsu(asid, orderserial, byref ErrStr)

	'// =======================================================================
	if (CheckOverChangeOrderItemnoExist(id, orderserial) = True) then

        CheckAndAddChangeOrderJupsu = False
        ErrStr = "���ֹ�( + ��ȯ�ֹ� + ���̳ʽ��ֹ�) �� �ʰ��Ͽ� ȸ���Ǵ� ��ǰ�� �ֽ��ϴ�.\n(�ߺ� �����Ǿ��� �� �ֽ��ϴ�. �ý����� ����)"

        exit function

	end if

	CheckAndAddChangeOrderJupsu = AddChangeOrderJupsu(asid, orderserial)
end function

function DelChangeOrder(asid)

	dim sqlStr

	DelChangeOrder = ""

end function

'// ===========================================================================
'// ��ǰ���� �±�ȯ ��ȯ�ֹ� ���(�ֹ����� ���·� ���)
'// ===========================================================================
''��������
''�ɼǺ��� : ���� �ɼǰ� �����ؾ߸� �±�ȯ ����, �������� �״�� ī��
''��ǰ���� :
''           �ǸŰ� �����ϰ�, ������ ������ ��� : �Ǻ귣��,���� �ǸŰ�(���ΰ�),���԰�, �������밡�� �� ��� �����ؾ��ϰ� 1:1 ���游 ����(������ ������ ����), �������� �״�� ī��
''           �ٸ� ���, �ǸŰ�(���ΰ�),���԰��� CS���������� �̿�, ������ ������(���ΰ�-������ �� �ٸ� ���) �������� ����
function AddChangeOrderJupsu(id, orderserial)
    dim sqlStr
    dim iid
    dim rndjumunno
    dim neworderserial
    dim ischangeorder, orgorderserial
    dim refasid

    dim subtotalprice, miletotalprice, tencardspend, spendmembership, allatdiscountprice
    sqlStr = " select 0 as subtotalprice, 0 as miletotalprice,"
    sqlStr = sqlStr + " 0 as tencardspend, 0 as spendmembership,"
    sqlStr = sqlStr + " 0 as allatdiscountprice "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        subtotalprice       = rsget("subtotalprice")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        spendmembership     = rsget("spendmembership")
        allatdiscountprice  = rsget("allatdiscountprice")
    rsget.close

    dim refundmileagesum, refundcouponsum, allatsubtractsum, refunditemcostsum
    dim refundbeasongpay, refunddeliverypay, refundadjustpay, canceltotal
    dim refundgiftcardsum, refunddepositsum

    refundrequire       = 0
    refundmileagesum    = 0
    refundcouponsum     = 0
    allatsubtractsum    = 0

    refunditemcostsum   = 0

    refundbeasongpay    = 0
    refunddeliverypay   = 0
    refundadjustpay     = 0
    canceltotal         = 0

    refundgiftcardsum	= 0
    refunddepositsum	= 0

	Randomize
	rndjumunno = CLng(Rnd * 100000) + 1
	rndjumunno = CStr(rndjumunno)

	sqlStr = "select * from [db_order].[dbo].tbl_order_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("orderserial") = rndjumunno
	rsget("jumundiv") = "6"					'// ��ȯ�ֹ�
	rsget("userid") = ""
	rsget("accountname") = ""
	rsget("accountdiv") = "7"				'// �ϴ��� ����������
	rsget("sitename") = ""
	rsget.update
	    iid = rsget("idx")
	rsget.close

	neworderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
	neworderserial = neworderserial & Format00(5,Right(CStr(iid),5))

    sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
    sqlStr = sqlStr + " set orderserial='" + neworderserial + "'" & vbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iid)
    dbget.Execute sqlStr

    sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
	sqlStr = sqlStr + " set userid=O.userid" & vbCrlf
	sqlStr = sqlStr + " ,accountname=O.accountname" & vbCrlf
	sqlStr = sqlStr + " ,accountdiv=O.accountdiv" & vbCrlf			'// �������� ����
	sqlStr = sqlStr + " ,paygatetid=O.paygatetid" & vbCrlf
	sqlStr = sqlStr + " ,authcode=O.authcode" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdiv='2'" & vbCrlf
	''sqlStr = sqlStr + " ,ipkumdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,regdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,beadaldiv=O.beadaldiv" & vbCrlf
	sqlStr = sqlStr + " ,beadaldate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,buyname=O.buyname" & vbCrlf
	sqlStr = sqlStr + " ,buyphone=O.buyphone" & vbCrlf
	sqlStr = sqlStr + " ,buyhp=O.buyhp" & vbCrlf
	sqlStr = sqlStr + " ,buyemail=O.buyemail" & vbCrlf
	sqlStr = sqlStr + " ,reqname=O.reqname" & vbCrlf
	sqlStr = sqlStr + " ,reqzipcode=O.reqzipcode" & vbCrlf
	sqlStr = sqlStr + " ,reqaddress=O.reqaddress" & vbCrlf
	sqlStr = sqlStr + " ,reqphone=O.reqphone" & vbCrlf
	sqlStr = sqlStr + " ,reqhp=O.reqhp" & vbCrlf
	sqlStr = sqlStr + " ,comment='���ֹ���ȣ:" + orderserial +"'" & vbCrlf
	sqlStr = sqlStr + " ,linkorderserial=O.orderserial" & vbCrlf
	sqlStr = sqlStr + " ,deliverno=''" & vbCrlf
	sqlStr = sqlStr + " ,sitename=O.sitename" & vbCrlf
	sqlStr = sqlStr + " ,discountrate=O.discountrate" & vbCrlf
	sqlStr = sqlStr + " ,subtotalprice=0" & vbCrlf
	sqlStr = sqlStr + " ,miletotalprice=0" & vbCrlf
	sqlStr = sqlStr + " ,tencardspend=0" & vbCrlf
	sqlStr = sqlStr + " ,spendmembership=0" & vbCrlf
	sqlStr = sqlStr + " ,allatdiscountprice=0" & vbCrlf
	sqlStr = sqlStr + " ,rduserid=O.rduserid" & vbCrlf
	sqlStr = sqlStr + " ,sentenceidx=O.sentenceidx" & vbCrlf
	sqlStr = sqlStr + " ,reqzipaddr=O.reqzipaddr" & vbCrlf
	sqlStr = sqlStr + " ,rdsite=O.rdsite" & vbCrlf
	sqlStr = sqlStr + " ,subtotalpriceCouponNotApplied=0" & vbCrlf
	sqlStr = sqlStr + " ,sumPaymentEtc=0 " & vbCrlf
	sqlStr = sqlStr + " ,pggubun=O.pggubun" & vbCrlf

	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from (select top 1 * from [db_log].[dbo].tbl_old_order_master_2003 where orderserial='" + orderserial + "') O" & vbCrlf
	else
	    sqlStr = sqlStr + " from (select top 1 * from [db_order].[dbo].tbl_order_master where orderserial='" + orderserial + "') O" & vbCrlf
	end if
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.idx=" + CStr(iid)
	dbget.Execute sqlStr

	'�󼼳���(�±�ȯȸ�� ��ǰ����)
	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
    sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
    sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType, itemcouponidx, bonuscouponidx, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, etcDiscount, currstate,beasongdate,upcheconfirmdate)" & vbCrlf
    sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
    sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
    sqlStr = sqlStr + " ,d.itemid" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
    sqlStr = sqlStr + " ,-1*J.confirmitemno" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
    sqlStr = sqlStr + " ,d.itemvat" & vbCrlf
    sqlStr = sqlStr + " ,d.mileage" & vbCrlf
    sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
    sqlStr = sqlStr + " ,d.itemname" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
    sqlStr = sqlStr + " ,d.makerid" & vbCrlf
    sqlStr = sqlStr + " ,d.buycash" & vbCrlf
    sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
    sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
    sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.omwdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.odlvType" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcouponidx" & vbCrlf
    sqlStr = sqlStr + " ,d.bonuscouponidx" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.orgitemcost,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.itemcostCouponNotApplied,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.buycashCouponNotApplied,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.plussaleDiscount,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.specialShopDiscount,0)" & vbCrlf
	sqlStr = sqlStr + " ,IsNull(d.etcDiscount,0)" & vbCrlf
    sqlStr = sqlStr + " ,'0'" & vbCrlf
    sqlStr = sqlStr + " ,NULL" & vbCrlf
    sqlStr = sqlStr + " ,NULL" & vbCrlf
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_detail J" & vbCrlf
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " ,[db_log].[dbo].tbl_old_order_detail_2003 d" & vbCrlf
    else
        sqlStr = sqlStr + " ,[db_order].[dbo].tbl_order_detail d" & vbCrlf
    end if
    sqlStr = sqlStr + " where J.masterid=" & CStr(id)
    sqlStr = sqlStr + " and d.orderserial='" & orderserial & "'"  & vbCrlf
    sqlStr = sqlStr + " and J.orderdetailidx=d.idx"  & vbCrlf
	sqlStr = sqlStr + " and J.confirmitemno<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " and J.orderdetailidx is not null "
    dbget.Execute sqlStr

	refasid = GetRefAsid(id)

	'// ��ǰ���� ��ȯ��� ��ǰ(A100)
	Call AddOrderDetail(refasid, orderserial, neworderserial, iid, "0")

    ''�ֹ��ݾ� ����
    call recalcuOrderMaster(neworderserial)

	'' ����� --> ���Ϸ�� �Ѵ�.
	'' �������� --> CS������ �̹� �ݿ��ߴ�.

    sqlStr = " select jumundiv "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    	if (rsget("jumundiv") = "6") then
    		ischangeorder = True
    	else
    		ischangeorder = False
    	end if
    rsget.close

	if (ischangeorder = True) then

	    sqlStr = " select orgorderserial "
        sqlStr = sqlStr + " from [db_order].[dbo].[tbl_change_order] "
	    sqlStr = sqlStr + " where chgorderserial='" + orderserial + "'"

	    rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	    	if Not rsget.Eof then
	    		orgorderserial = rsget("orgorderserial")
	    	end if
	    rsget.close

	else
		orgorderserial = orderserial
	end if

	sqlStr = " insert into [db_order].[dbo].[tbl_change_order](orgorderserial, chgorderserial) "
	sqlStr = sqlStr + " values('" + CStr(orgorderserial) + "', '" + CStr(neworderserial) + "') "
    dbget.Execute sqlStr

    AddChangeOrderJupsu    = neworderserial
end Function

Function AddOrderDetail(refasid, orderserial, neworderserial, neworderserialidx, currstate)
	'// ��ǰ���� ��ȯ��� ��ǰ(A100)
	Dim sqlStr

	''��������
	''�ɼǺ��� : ���� �ɼǰ� �����ؾ߸� �±�ȯ ����, �������� �״�� ī��
	''��ǰ���� :
	''           �ǸŰ� �����ϰ�, ������ ������ ��� : �Ǻ귣��,���� �ǸŰ�(���ΰ�),���԰�, �������밡�� �� ��� �����ؾ��ϰ� 1:1 ���游 ����(������ ������ ����), �������� �״�� ī��
	''           �ٸ� ���, �ǸŰ�(���ΰ�),���԰��� CS���������� �̿�, ������ ������(���ΰ�-������ �� �ٸ� ���) �������� ����

	'�󼼳���(�±�ȯ��� ��ǰ����) : �ǸŰ� �����ϰ�, ������ ������ ���
	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
    sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
    sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType, itemcouponidx, bonuscouponidx, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, etcDiscount, currstate,beasongdate,upcheconfirmdate)" & vbCrlf
    sqlStr = sqlStr + " select " & CStr(neworderserialidx) & vbCrlf
    sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
    sqlStr = sqlStr + " ,J.itemid" & vbCrlf
    sqlStr = sqlStr + " ,J.itemoption" & vbCrlf
    sqlStr = sqlStr + " ,J.confirmitemno" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
    sqlStr = sqlStr + " ,d.itemvat" & vbCrlf
    sqlStr = sqlStr + " ,d.mileage" & vbCrlf
    sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
    sqlStr = sqlStr + " ,J.itemname" & vbCrlf
    sqlStr = sqlStr + " ,J.itemoptionname" & vbCrlf
    sqlStr = sqlStr + " ,J.makerid" & vbCrlf
    sqlStr = sqlStr + " ,d.buycash" & vbCrlf
    sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
    sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
    sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.omwdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.odlvType" & vbCrlf
    sqlStr = sqlStr + " ,d.itemcouponidx" & vbCrlf
    sqlStr = sqlStr + " ,d.bonuscouponidx" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.orgitemcost,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.itemcostCouponNotApplied,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.buycashCouponNotApplied,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.plussaleDiscount,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(d.specialShopDiscount,0)" & vbCrlf
	sqlStr = sqlStr + " ,IsNull(d.etcDiscount,0)" & vbCrlf

	If (currstate = "0") Then
		sqlStr = sqlStr + " ,'0'" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
	Else
		sqlStr = sqlStr + " ,'7'" & vbCrlf
		sqlStr = sqlStr + " ,getdate()" & vbCrlf
		sqlStr = sqlStr + " ,getdate()" & vbCrlf
	End If

    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_detail J" & vbCrlf
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " ,[db_log].[dbo].tbl_old_order_detail_2003 d" & vbCrlf
    else
        sqlStr = sqlStr + " ,[db_order].[dbo].tbl_order_detail d" & vbCrlf
    end if
    sqlStr = sqlStr + " where J.masterid=" & CStr(refasid)
    sqlStr = sqlStr + " and d.orderserial='" & orderserial & "'"  & vbCrlf
    sqlStr = sqlStr + " and J.reforderdetailidx=d.idx"  & vbCrlf
	sqlStr = sqlStr + " and J.confirmitemno<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " and J.orderdetailidx is null "
    sqlStr = sqlStr + " and J.reforderdetailidx <> 0 "
	sqlStr = sqlStr + " and J.SalePrice is NULL "			'// ���� �����ϰ�, �ǸŰ� ������ ���
    dbget.Execute sqlStr

	'�󼼳���(�±�ȯ��� ��ǰ����) : �ǸŰ� �ٸ� ��� or ���� �ٸ� ���
	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
    sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
    sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType, itemcouponidx, bonuscouponidx, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, currstate,beasongdate,upcheconfirmdate)" & vbCrlf
	sqlStr = sqlStr + " select " & CStr(neworderserialidx) & vbCrlf
    sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
    sqlStr = sqlStr + " ,J.itemid" & vbCrlf
    sqlStr = sqlStr + " ,J.itemoption" & vbCrlf
    sqlStr = sqlStr + " ,J.confirmitemno" & vbCrlf
    sqlStr = sqlStr + " , J.SalePrice " & vbCrlf
    sqlStr = sqlStr + " , Round(J.SalePrice/11, 0)" & vbCrlf
    sqlStr = sqlStr + " ,i.mileage" & vbCrlf
    sqlStr = sqlStr + " ,J.itemcost " & vbCrlf
    sqlStr = sqlStr + " ,J.itemname" & vbCrlf
    sqlStr = sqlStr + " ,J.itemoptionname" & vbCrlf
    sqlStr = sqlStr + " ,J.makerid" & vbCrlf
    sqlStr = sqlStr + " ,J.buycash" & vbCrlf
    sqlStr = sqlStr + " ,i.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,J.isupchebeasong" & vbCrlf
    sqlStr = sqlStr + " ,i.sailyn" & vbCrlf
    sqlStr = sqlStr + " ,i.itemdiv" & vbCrlf
    sqlStr = sqlStr + " ,i.mwdiv" & vbCrlf
    sqlStr = sqlStr + " ,i.deliverytype" & vbCrlf
    sqlStr = sqlStr + " ,NULL" & vbCrlf
    sqlStr = sqlStr + " ,NULL" & vbCrlf
    sqlStr = sqlStr + " , i.orgprice + IsNull(v.optaddprice,0) " & vbCrlf
    sqlStr = sqlStr + " , i.sellcash + IsNull(v.optaddprice,0)" & vbCrlf
    sqlStr = sqlStr + " ,IsNull(i.buycash,0)" & vbCrlf
    sqlStr = sqlStr + " ,0" & vbCrlf
    sqlStr = sqlStr + " ,0" & vbCrlf

	If (currstate = "0") Then
		sqlStr = sqlStr + " ,'0'" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
		sqlStr = sqlStr + " ,NULL" & vbCrlf
	Else
		sqlStr = sqlStr + " ,'7'" & vbCrlf
		sqlStr = sqlStr + " ,getdate()" & vbCrlf
		sqlStr = sqlStr + " ,getdate()" & vbCrlf
	End If

    sqlStr = sqlStr + " from" & vbCrlf
	sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_detail J " & vbCrlf
	sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i " & vbCrlf
	sqlStr = sqlStr + " 	on " & vbCrlf
	sqlStr = sqlStr + " 		J.itemid = i.itemid " & vbCrlf
	sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option v " & vbCrlf
	sqlStr = sqlStr + " 	on " & vbCrlf
	sqlStr = sqlStr + " 		1 = 1 " & vbCrlf
	sqlStr = sqlStr + " 		and i.itemid = v.itemid " & vbCrlf
	sqlStr = sqlStr + " 		and J.itemoption = v.itemoption " & vbCrlf
    sqlStr = sqlStr + " where J.masterid=" & CStr(refasid)
    sqlStr = sqlStr + " and J.orderserial='" & orderserial & "'"  & vbCrlf
	sqlStr = sqlStr + " and J.confirmitemno<>0"
    sqlStr = sqlStr + " and J.orderdetailidx is null "
    sqlStr = sqlStr + " and J.reforderdetailidx <> 0 "
	sqlStr = sqlStr + " and J.SalePrice is not NULL "			'// �ǸŰ� �ٸ� ��� or ���� �ٸ� ���
    dbget.Execute sqlStr
End Function

function FinishChangeOrder(changeorderserial)
	dim sqlStr

	sqlStr = " update [db_order].[dbo].tbl_order_detail "
	sqlStr = sqlStr + " set currstate = '7', beasongdate = getdate(), upcheconfirmdate = getdate() "
	sqlStr = sqlStr + " where orderserial = '" + CStr(changeorderserial) + "' "
	dbget.Execute sqlStr

	sqlStr = " update [db_order].[dbo].tbl_order_master "
	sqlStr = sqlStr + " set ipkumdiv = '8', ipkumdate = getdate(), baljudate = getdate() "
	sqlStr = sqlStr + " where orderserial = '" + CStr(changeorderserial) + "' "
	dbget.Execute sqlStr

	'// ���Ϸ�� ��� ����
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_changeOrder '" & changeorderserial & "'"
    dbget.Execute sqlStr

end function

function AddChangeOrder(id, orderserial)
    dim changeorderserial

    changeorderserial = AddChangeOrderJupsu(id, orderserial)

    Call FinishChangeOrder(changeorderserial)

    AddChangeOrder = changeorderserial
end function

'// ��ȯ�ֹ� ã��
function GetChangeOrderInfo(asid, byref changeorderserial, byref changeorderstate,  byref errMsg)
    dim sqlStr

	changeorderserial = ""
	changeorderstate = ""
	errMsg = ""

	'// =======================================================================
    sqlStr = " select top 1 a.divcd, LTrim(IsNull(m.orderserial, '')) as changeorderserial, LTrim(IsNull(m.ipkumdiv, '')) as changeorderstate "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.refchangeorderserial = m.orderserial "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and a.id = " + CStr(asid) + " "
    sqlStr = sqlStr + " 	and a.divcd in ('A111', 'A112') "
	'rw sqlStr

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        changeorderserial = rsget("changeorderserial")
        changeorderstate = rsget("changeorderstate")
    else
    	errMsg = "�߸��� �����Դϴ�."
    end if
    rsget.Close

end function

'// ��ȯ�ֹ� �귣������
function GetChangeOrderBrandInfo(orderserial)
    dim sqlStr

    sqlStr = " select top 1 isupchebeasong, makerid "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	db_order.dbo.tbl_order_detail "
    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and cancelyn <> 'Y' "

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        if rsget("isupchebeasong") = "Y" then
        	GetChangeOrderBrandInfo = rsget("makerid")
        else
        	GetChangeOrderBrandInfo = ""
        end if
    end if
    rsget.Close

end function

function InsertEtcPaymentOne(orderserial, acctdiv, acctamount)
	dim sqlStr

    sqlStr = " insert into [db_order].[dbo].tbl_order_PaymentEtc( " + VbCrlf
    sqlStr = sqlStr + " 	orderserial " + VbCrlf
    sqlStr = sqlStr + " 	, acctdiv " + VbCrlf
    sqlStr = sqlStr + " 	, acctamount " + VbCrlf
    sqlStr = sqlStr + " 	, realPayedsum " + VbCrlf
    sqlStr = sqlStr + " ) " + VbCrlf
    sqlStr = sqlStr + " values( " + VbCrlf
    sqlStr = sqlStr + " 	'" & orderserial & "' " + VbCrlf
    sqlStr = sqlStr + " 	, '" & acctdiv & "' " + VbCrlf
    sqlStr = sqlStr + " 	, " & CStr(acctamount) & " " + VbCrlf
    sqlStr = sqlStr + " 	, " & CStr(acctamount) & " " + VbCrlf
    sqlStr = sqlStr + " ) " + VbCrlf

    dbget.Execute sqlStr

end function

function CheckNEditRefundInfo(asid, returnmethod, rebankaccount, rebankownername, rebankname, paygateTid , refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay  )
    dim sqlStr
    dim refundInfoExists, oldrefundrequire
    refundInfoExists     = false
    CheckNEditRefundInfo = false

    if ((returnmethod="") ) then Exit function
    if ((Not IsNumeric(refundrequire)) or (refundrequire="")) then Exit function

    sqlStr = " select * from [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " where asid=" + CStr(asid)

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        refundInfoExists = True
        oldrefundrequire = rsget("refundrequire")
    end if
    rsget.Close

    if (Not refundInfoExists) then Exit function

    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"                             + VbCrlf
    sqlStr = sqlStr + " set returnmethod='" + returnmethod + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankaccount='" + rebankaccount + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankownername='" + rebankownername + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankname='" + rebankname + "'"                          + VbCrlf
    sqlStr = sqlStr + " , paygateTid='" + paygateTid + "'"                          + VbCrlf

    sqlStr = sqlStr + " , orgsubtotalprice=" & orgsubtotalprice & VbCrlf
    sqlStr = sqlStr + " , orgitemcostsum=" & orgitemcostsum & VbCrlf
    sqlStr = sqlStr + " , orgbeasongpay=" & orgbeasongpay & VbCrlf
    sqlStr = sqlStr + " , orgmileagesum=" & orgmileagesum & VbCrlf
    sqlStr = sqlStr + " , orgcouponsum=" & orgcouponsum & VbCrlf
    sqlStr = sqlStr + " , orgallatdiscountsum=" & orgallatdiscountsum & VbCrlf
    sqlStr = sqlStr + " , canceltotal=" & canceltotal & VbCrlf
    sqlStr = sqlStr + " , refunditemcostsum=" & refunditemcostsum & VbCrlf
    sqlStr = sqlStr + " , refundmileagesum=" & refundmileagesum & VbCrlf
    sqlStr = sqlStr + " , refundcouponsum=" & refundcouponsum & VbCrlf
    sqlStr = sqlStr + " , allatsubtractsum=" & allatsubtractsum & VbCrlf
    sqlStr = sqlStr + " , refundbeasongpay=" & refundbeasongpay & VbCrlf
    sqlStr = sqlStr + " , refunddeliverypay=" & refunddeliverypay & VbCrlf
    sqlStr = sqlStr + " , refundadjustpay=" & refundadjustpay & VbCrlf

    ''�������̳� ���ϸ��� ȯ���̳� ��ġ����ȯ�� ��츸 ���� ���� ����
    ''if ((returnmethod="R007") or (returnmethod="R900") or (returnmethod="R910") or (returnmethod="R000")) and (refundrequire<>oldrefundrequire) then
    if (refundrequire<>oldrefundrequire) then
        sqlStr = sqlStr + " , refundrequire=" + CStr(refundrequire)                     + VbCrlf
        '''sqlStr = sqlStr + " , refundadjustpay=" + CStr(refundrequire) + "-canceltotal"  + VbCrlf
    end if
    sqlStr = sqlStr + " where asid=" + CStr(asid)

'response.write   sqlStr
    dbget.Execute sqlStr

    CheckNEditRefundInfo = true
end function

function CheckNEditRefundInfo_OLD(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire)
    dim sqlStr
    dim refundInfoExists, oldrefundrequire
    refundInfoExists     = false
    CheckNEditRefundInfo_OLD = false

    if ((returnmethod="") or (returnmethod="R000")) then Exit function
    if ((Not IsNumeric(refundrequire)) or (refundrequire="")) then Exit function


    sqlStr = " select * from [db_cs].[dbo].tbl_as_refund_info"
    sqlStr = sqlStr + " where asid=" + CStr(id)

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        refundInfoExists = True
        oldrefundrequire = rsget("refundrequire")
    end if
    rsget.Close

    if (Not refundInfoExists) then Exit function

    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"                             + VbCrlf
    sqlStr = sqlStr + " set returnmethod='" + returnmethod + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankaccount='" + rebankaccount + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankownername='" + rebankownername + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankname='" + rebankname + "'"                          + VbCrlf
    sqlStr = sqlStr + " , paygateTid='" + paygateTid + "'"                          + VbCrlf

    ''�������̳� ���ϸ��� ȯ���̳� ��ġ����ȯ�� ��츸 ���� ���� ����
    if ((returnmethod="R007") or (returnmethod="R900") or (returnmethod="R910")) and (refundrequire<>oldrefundrequire) then
        sqlStr = sqlStr + " , refundrequire=" + CStr(refundrequire)                     + VbCrlf
        '''sqlStr = sqlStr + " , refundadjustpay=" + CStr(refundrequire) + "-canceltotal"  + VbCrlf
    end if
    sqlStr = sqlStr + " where asid=" + CStr(id)

'response.write   sqlStr
    dbget.Execute sqlStr

    CheckNEditRefundInfo_OLD = true
end function

function LimitItemRecover(byval orderserial)
    dim sqlStr
    On Error Resume Next
        ''�������� ���� -
        sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
        sqlStr = sqlStr + " set limitsold=(case when 0>limitsold - T.itemno then 0 else limitsold - T.itemno end)" + vbCrlf
        sqlStr = sqlStr + " from " + vbCrlf
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " 	select d.itemid, d.itemno" + vbCrlf
        sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d" + vbCrlf
        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemid<>0 "
        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " ) as T" + vbCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.Itemid"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.limityn='Y'"

        dbget.Execute(sqlStr)

        ''�ɼ��ִ»�ǰ
        sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
        sqlStr = sqlStr + " set optlimitsold=(case when 0 >optlimitsold - T.itemno then 0 else optlimitsold - T.itemno end)" + vbCrlf
        sqlStr = sqlStr + " from " + vbCrlf
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " 	select d.itemid, d.itemoption, d.itemno" + vbCrlf
        sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d " + vbCrlf
        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemid<>0" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemoption<>'0000'" + vbCrlf
        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " ) as T" + vbCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

        dbget.Execute(sqlStr)
    On Error Goto 0
end function

'// CS �±�ȯ���(���ϻ�ǰ, ��ǰ���� - A000, A100) ������ ���Ǵ� ��ǰ ��������
function ApplyLimitItemByCS(asid)
    dim sqlStr
    dim divcd, currstate

    divcd = ""

    sqlStr = " select top 1 m.divcd "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.id = " + CStr(asid) + " "
    sqlStr = sqlStr + " 	and m.divcd in ('A000', 'A100') "
    sqlStr = sqlStr + " 	and m.currstate = 'B001' "
    sqlStr = sqlStr + " 	and m.deleteyn = 'N' "

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        divcd = rsget("divcd")
    end if
    rsget.Close

	if (divcd = "")  then Exit function


    ''�������� ���� -
    sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
    sqlStr = sqlStr + " set limitsold=(case when limitno <= (limitsold + T.itemno) then limitno else (limitsold + T.itemno) end)" + vbCrlf
    sqlStr = sqlStr + " from " + vbCrlf
    sqlStr = sqlStr + " ("
    sqlStr = sqlStr + " select d.itemid, d.regitemno as itemno " + vbCrlf
    sqlStr = sqlStr + " from " + vbCrlf
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_detail d " + vbCrlf
    sqlStr = sqlStr + " where " + vbCrlf
    sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
    sqlStr = sqlStr + " 	and d.masterid = " + CStr(asid) + " " + vbCrlf
    sqlStr = sqlStr + " 	and d.regitemno > 0 " + vbCrlf			'// ��ǰ���� �±�ȯ����ǰ +, ȸ����ǰ -
    sqlStr = sqlStr + " ) as T" + vbCrlf
    sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.Itemid"
    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.limityn='Y'"
    dbget.Execute(sqlStr)

	''�ɼ��ִ»�ǰ
    sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
    sqlStr = sqlStr + " set optlimitsold=(case when optlimitno <= (optlimitsold + T.itemno) then optlimitno else (optlimitsold + T.itemno) end)" + vbCrlf
    sqlStr = sqlStr + " from " + vbCrlf
    sqlStr = sqlStr + " ("
    sqlStr = sqlStr + " select d.itemid, d.itemoption, d.regitemno as itemno " + vbCrlf
    sqlStr = sqlStr + " from " + vbCrlf
    sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_detail d " + vbCrlf
    sqlStr = sqlStr + " where " + vbCrlf
    sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
    sqlStr = sqlStr + " 	and d.masterid = " + CStr(asid) + " " + vbCrlf
    sqlStr = sqlStr + " 	and d.regitemno > 0 " + vbCrlf			'// ��ǰ���� �±�ȯ����ǰ +, ȸ����ǰ -
    sqlStr = sqlStr + " 	and d.itemoption<>'0000'" + vbCrlf
    sqlStr = sqlStr + " ) as T" + vbCrlf
    sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"
	dbget.Execute(sqlStr)

end function


function IsExtSiteOrder(orderserial)
    dim sqlStr

    sqlStr = " select IsNULL(sitename,'') as sitename from [db_order].[dbo].tbl_order_master"
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        IsExtSiteOrder = (rsget("sitename")<>"10x10")
    else
        IsExtSiteOrder = False
    end if
    rsget.close

end function

function CheckNUsafeCancel(byval orderserial)
    dim sqlStr, result
	dim InsureCd

	result = False

    sqlStr = " select IsNULL(InsureCd,'') as InsureCd from [db_order].[dbo].tbl_order_master"
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        result = (rsget("InsureCd")="0")
    end if
    rsget.close

	if (result = True) then
		Call UsafeCancel(orderserial)
	end if
end function

sub UsafeCancel(byval orderserial)
    '// ���ں������� ������ ������ ��� ��û (2006.06.15; ������� ������)
    dim objUsafe, result, result_code, result_msg
    On Error Resume Next
    	Set objUsafe = CreateObject( "USafeCom.guarantee.1"  )

    '	Test�� ��
    '	objUsafe.Port = 80
    '	objUsafe.Url = "gateway2.usafe.co.kr"
    '	objUsafe.CallForm = "/esafe/guartrn.asp"

        ' Real�� ��
        objUsafe.Port = 80
        objUsafe.Url = "gateway.usafe.co.kr"
        objUsafe.CallForm = "/esafe/guartrn.asp"

    	objUsafe.gubun	= "B0"				'// �������� (A0:�űԹ߱�, B0:���������, C0:�Ա�Ȯ��)
    	objUsafe.EncKey	= ""			'�ΰ��� ��� ��ȣȭ �ȵ�
    	objUsafe.mallId	= "ZZcube1010"		'// ���θ�ID
    	objUsafe.oId	= CStr(orderserial)	'// �ֹ���ȣ

    	'ó�� ����!
    	result = objUsafe.cancelInsurance

    	result_code	= Left( result , 1 )
    	result_msg	= Mid( result , 3 )

    	Set objUsafe = Nothing
    On Error Goto 0
end Sub

function GetUserRefundAuthLimit(userid)
    dim sqlStr

	GetUserRefundAuthLimit = 0

	sqlStr = " select top 1 defaultCSRefundLimit "
	sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_refund_user "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and useyn = 'Y' "
	sqlStr = sqlStr + " 	and userid = '" + CStr(userid) + "' "
	sqlStr = sqlStr + " order by idx "

	''rw sqlStr
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			if (rsget("defaultCSRefundLimit") > 0) then
				GetUserRefundAuthLimit = rsget("defaultCSRefundLimit")
			end if
		end if
    rsget.close

end function

''����. ��ü ��� �´���.
''��ü��������� Ȯ���ϰ� �ٸ� ���� ���Ѵ�.
function IsAllCancelRegValid(Asid, orderserial)
    dim sqlStr
    IsAllCancelRegValid = false

    sqlStr = "select count(d.idx) as cnt"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
    end if
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + "     on c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + "     and c.orderdetailidx=d.idx"
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " and d.itemno<>IsNULL(c.regitemno,0)"
''rw sqlStr
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IsAllCancelRegValid = (rsget("cnt")=0)
    rsget.close

end function

function GetIsLastReturnBrand(orderserial)
    dim sqlStr
    GetIsLastReturnBrand = false

    sqlStr = " select count(T.makerid) as cnt "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	( "
    sqlStr = sqlStr + " 		select (case when d.isupchebeasong = 'N' then '' else lower(d.makerid) end) as makerid "
    sqlStr = sqlStr + " 		from "
    sqlStr = sqlStr + " 			[db_order].[dbo].[tbl_order_detail] d "
    sqlStr = sqlStr + " 			left join ( "
    sqlStr = sqlStr + " 				select d.itemid, d.itemoption, sum(d.confirmitemno) as itemCnt "
    sqlStr = sqlStr + " 				from "
    sqlStr = sqlStr + " 					[db_cs].[dbo].[tbl_new_as_list] a "
    sqlStr = sqlStr + " 					join [db_cs].[dbo].[tbl_new_as_detail] d on a.id = d.masterid "
    sqlStr = sqlStr + " 				where "
    sqlStr = sqlStr + " 					1 = 1 "
    sqlStr = sqlStr + " 					and a.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 					and a.divcd in ('A010', 'A004') "
    sqlStr = sqlStr + " 					and a.deleteyn = 'N' "
    sqlStr = sqlStr + " 				group by "
    sqlStr = sqlStr + " 					d.itemid, d.itemoption "
    sqlStr = sqlStr + " 			) T "
    sqlStr = sqlStr + " 			on "
    sqlStr = sqlStr + " 				1 = 1 "
    sqlStr = sqlStr + " 				and d.itemid = T.itemid "
    sqlStr = sqlStr + " 				and d.itemoption = T.itemoption "
    sqlStr = sqlStr + " 		where "
    sqlStr = sqlStr + " 			1 = 1 "
    sqlStr = sqlStr + " 			and d.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 			and d.itemid not in (0, 100) "
    sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
    sqlStr = sqlStr + " 			and d.itemno <> IsNull(T.itemCnt, 0) "
    sqlStr = sqlStr + " 		group by "
    sqlStr = sqlStr + " 			(case when d.isupchebeasong = 'N' then '' else lower(d.makerid) end) "
    sqlStr = sqlStr + " 	) T "
    ''rw sqlStr
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        GetIsLastReturnBrand = (rsget("cnt") = 1)
    rsget.close

end function

'���̴� ���� ������ �����Ƿ� ���ܵд�.
function IsPartialCancelRegValid(Asid, orderserial)
    dim sqlStr
    IsPartialCancelRegValid = false

    sqlStr = "select count(d.idx) as cnt, sum(case when d.itemno=IsNULL(c.regitemno,0) then 1 else 0 end) as Matchcount"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
    else
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
    end if
    sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_detail c"
    sqlStr = sqlStr + "     on c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + "     and c.orderdetailidx=d.idx"
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        IsPartialCancelRegValid = Not (rsget("cnt")=rsget("Matchcount"))
    rsget.close
end function

function SaveCSListHistory(asid)
    dim sqlStr

	'// ���� ó���� ���̵� ����
	sqlStr = " exec [db_log].[dbo].[usp_Ten_SaveCSHistory] " + CStr(asid) + " "
	dbget.Execute(sqlStr)

end function

' �ֹ� ��ǰ���� ����� üũ		'2023.10.19 �ѿ�� ����
function ItemCouponCount(asid, couponGubun, userid)
    dim sqlStr, returnCount
    returnCount=0

	if asid="" or isnull(asid) then
        ItemCouponCount=returnCount
        exit function
    end if
    asid = trim(asid)
	if userid="" or isnull(userid) then
        ItemCouponCount=returnCount
        exit function
    end if
    userid = trim(userid)

    sqlStr = " select"
    sqlStr = sqlStr & " count(t.itemcouponidx) as itemCouponCount"
    sqlStr = sqlStr & " from ("
    sqlStr = sqlStr & " 	select"
    sqlStr = sqlStr & " 	d.itemcouponidx"
    sqlStr = sqlStr & " 	, isnull((select count(cc.couponidx)"
    sqlStr = sqlStr & "  		from db_item.dbo.tbl_user_item_coupon cc with (nolock)"
    sqlStr = sqlStr & "  		where cc.itemcouponidx = c.itemcouponidx"
    sqlStr = sqlStr & "  		and cc.userid = c.userid"
    sqlStr = sqlStr & "  		and cc.couponidx <> c.couponidx),0) as prevCopiedItemCouponCount"
    sqlStr = sqlStr & "  		, rank() over (partition by c.userid, c.itemcouponidx order by c.couponidx desc) as rk"
    sqlStr = sqlStr & " 	from db_cs.dbo.tbl_new_as_detail ad with (nolock)"
    sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail d with (nolock)"
    sqlStr = sqlStr & " 		on ad.orderdetailidx=d.idx"
    sqlStr = sqlStr & " 		and ad.orderserial=d.orderserial"
    sqlStr = sqlStr & " 	join db_item.dbo.tbl_user_item_coupon c with (nolock)"
    sqlStr = sqlStr & " 		on d.itemcouponidx=c.itemcouponidx"
    sqlStr = sqlStr & " 		and ad.orderserial=c.orderserial"
    sqlStr = sqlStr & " 		and c.itemcouponexpiredate>getdate()"	' ��ȿ�Ⱓüũ

    if couponGubun<>"" then
        sqlStr = sqlStr & " and c.couponGubun='"& couponGubun &"'"
    end if
    if userid<>"" then
        sqlStr = sqlStr & " and c.userid='"& userid &"'"
    end if
    
    sqlStr = sqlStr & " 	where ad.masterid='" & asid & "'"
    sqlStr = sqlStr & " ) as t"
    sqlStr = sqlStr & " where t.rk=1"
    sqlStr = sqlStr & " and t.prevCopiedItemCouponCount=0"    ' ��߱�üũ

	'response.write sqlStr & "<Br>"
    'response.end
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        returnCount=rsget("itemCouponCount")
    else
        returnCount=0
    end if
    rsget.Close

    ItemCouponCount=returnCount
end function

' ��ǰ���� ����߱�     ' 2023.10.19 �ѿ�� ����
function CheckAndCopyItemCoupon(asid, reguserid, couponGubun, userid)
	dim orderserial, copyitemcouponinfo, sqlStr, excuteRowCount
    excuteRowCount=0

	if asid="" or isnull(asid) then
        CheckAndCopyItemCoupon = False
        exit function
    end if
    asid = trim(asid)

	sqlStr = " select top 1"
    sqlStr = sqlStr & " a.orderserial, IsNull(r.copyitemcouponinfo, 'N') as copyitemcouponinfo"
	sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_list a with (nolock)"
	sqlStr = sqlStr & " join [db_cs].[dbo].tbl_as_refund_info r with (nolock)"
	sqlStr = sqlStr & " 	on a.id = r.asid "
	sqlStr = sqlStr & " where a.id = "& asid &""
    sqlStr = sqlStr & " and a.divcd in ('A008', 'A004', 'A010')"    ' A008 �ֹ���� / A004 ��ǰ����(��ü���) / A010 ȸ����û(�ٹ����ٹ��)

	orderserial = ""
	copyitemcouponinfo = "N"
    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        orderserial    	= rsget("orderserial")
		copyitemcouponinfo  = rsget("copyitemcouponinfo")
    end if
    rsget.Close

	if (orderserial = "") or (copyitemcouponinfo <> "Y") then
		CheckAndCopyItemCoupon = False
		exit function
	end if

    sqlStr = "insert into db_item.dbo.tbl_user_item_coupon("
    sqlStr = sqlStr & " userid, itemcouponidx, issuedno, itemcoupontype, itemcouponvalue"
    sqlStr = sqlStr & " , itemcouponstartdate, itemcouponexpiredate"
    sqlStr = sqlStr & " , itemcouponname, itemcouponimage, regdate, usedyn, orderserial, couponGubun, csorderserial"
    sqlStr = sqlStr & " )"
    sqlStr = sqlStr & "     select"
    sqlStr = sqlStr & "     t.userid, t.itemcouponidx, t.issuedno, t.itemcoupontype, t.itemcouponvalue"
    sqlStr = sqlStr & "     , t.itemcouponstartdate, t.itemcouponexpiredate"
    sqlStr = sqlStr & "     , t.itemcouponname, t.itemcouponimage, t.regdate, t.usedyn"
    sqlStr = sqlStr & "     , t.orderserial, t.couponGubun, t.csorderserial"
    sqlStr = sqlStr & "     from ("
    sqlStr = sqlStr & " 	    select"
    sqlStr = sqlStr & "         c.userid, c.itemcouponidx, c.issuedno, c.itemcoupontype, c.itemcouponvalue"
    sqlStr = sqlStr & "         , c.itemcouponstartdate, c.itemcouponexpiredate"
    sqlStr = sqlStr & "         , c.itemcouponname, c.itemcouponimage, getdate() as regdate"
    sqlStr = sqlStr & "         , 'N' as usedyn, NULL as orderserial, c.couponGubun, c.orderserial as csorderserial"
    sqlStr = sqlStr & " 	    , isnull((select count(cc.couponidx)"
    sqlStr = sqlStr & "  	    	from db_item.dbo.tbl_user_item_coupon cc with (nolock)"
    sqlStr = sqlStr & "  	    	where cc.itemcouponidx = c.itemcouponidx"
    sqlStr = sqlStr & "  	    	and cc.userid = c.userid"
    sqlStr = sqlStr & "  	    	and cc.couponidx <> c.couponidx),0) as prevCopiedItemCouponCount"
    sqlStr = sqlStr & "  		, rank() over (partition by c.userid, c.itemcouponidx order by c.couponidx desc) as rk"
    sqlStr = sqlStr & " 	    from db_cs.dbo.tbl_new_as_detail ad with (nolock)"
    sqlStr = sqlStr & " 	    join db_order.dbo.tbl_order_detail d with (nolock)"
    sqlStr = sqlStr & " 	    	on ad.orderdetailidx=d.idx"
    sqlStr = sqlStr & " 	    	and ad.orderserial=d.orderserial"
    sqlStr = sqlStr & " 	    join db_item.dbo.tbl_user_item_coupon c with (nolock)"
    sqlStr = sqlStr & " 	    	on d.itemcouponidx=c.itemcouponidx"
    sqlStr = sqlStr & " 	    	and ad.orderserial=c.orderserial"
    sqlStr = sqlStr & " 	    	and c.itemcouponexpiredate>getdate()"	' ��ȿ�Ⱓüũ

    if couponGubun<>"" then
        sqlStr = sqlStr & "     and c.couponGubun='"& couponGubun &"'"
    end if
    if userid<>"" then
        sqlStr = sqlStr & "     and c.userid='"& userid &"'"
    end if
    
    sqlStr = sqlStr & " 	    where ad.masterid='" & asid & "'"
    sqlStr = sqlStr & "     ) as t"
    sqlStr = sqlStr & "     where t.rk=1"
    sqlStr = sqlStr & "     and t.prevCopiedItemCouponCount=0"    ' ��߱�üũ

    'response.write sqlStr & "<br>"
	dbget.Execute sqlStr, excuteRowCount

    if excuteRowCount>0 then
	    CheckAndCopyItemCoupon = True
    else
        ' ��ǰ���� ��߱� ���� �� �Ϸ�ó�� ���̿� ���� ��ȿ�Ⱓ�� ������� ����࿩�� N���� �ٲ۴�.
        Call EditCSCopyItemCouponInfo(asid, "N")

        CheckAndCopyItemCoupon = false
    end if
end function

function CheckAndCopyBonusCoupon(asid, reguserid)
	dim sqlStr
	dim orderserial, copycouponinfo, bCpnIdx, prevCopyCouponExist

	sqlStr = " select top 1 a.orderserial, IsNull(r.copycouponinfo, 'N') as copycouponinfo "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where a.id = " + CStr(asid) + " and a.divcd in ('A008', 'A004', 'A010') "

	orderserial = ""
	copycouponinfo = "N"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        orderserial    	= rsget("orderserial")
		copycouponinfo  = rsget("copycouponinfo")
    end if
    rsget.Close

	if (orderserial = "") or (copycouponinfo = "N") then
		CheckAndCopyBonusCoupon = False
		exit function
	end if

	sqlStr = " select "
	sqlStr = sqlStr + " 	m.bCpnIdx "
	sqlStr = sqlStr + " 	, ( "
	sqlStr = sqlStr + " 		select count(*) "
	sqlStr = sqlStr + " 		from [db_user].[dbo].tbl_user_coupon chk "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and chk.userid = c.userid "
	sqlStr = sqlStr + " 			and chk.masteridx = c.masteridx "
	sqlStr = sqlStr + " 			and chk.deleteyn <> 'Y' "
	sqlStr = sqlStr + " 			and chk.csorderserial = c.orderserial "
	sqlStr = sqlStr + " 			and chk.masteridx <> 287 "
	sqlStr = sqlStr + " 	) as cnt "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr + " 	join [db_user].[dbo].tbl_user_coupon c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.bCpnIdx = c.idx "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	m.orderserial = '" + CStr(orderserial) + "' "

	prevCopyCouponExist = True
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		prevCopyCouponExist = (rsget("cnt") > 0)
		bCpnIdx  = rsget("bCpnIdx")
    end if
    rsget.Close

	if prevCopyCouponExist = True then
		CheckAndCopyBonusCoupon = False
		exit function
	end if

	sqlStr = "insert into [db_user].[dbo].tbl_user_coupon(reguserid, isusing, masteridx, userid, coupontype, couponvalue, couponname, minbuyprice, targetitemlist, couponimage, startdate, expiredate, deleteyn, exitemid, validsitename, notvalid10x10, couponmeaipprice, ssnkey, scratchcouponidx, evtprize_code, useLevel, csorderserial, targetCpnType  , targetCpnSource, mxCpnDiscount) " + vbCrlf
	sqlStr = sqlStr + " select top 1 '" + CStr(reguserid) + "', 'N', masteridx, userid, coupontype, couponvalue, couponname, minbuyprice, targetitemlist, couponimage, startdate, expiredate, deleteyn, exitemid, validsitename, notvalid10x10, couponmeaipprice, ssnkey, scratchcouponidx, evtprize_code, useLevel, orderserial, targetCpnType  , targetCpnSource, mxCpnDiscount " + vbCrlf
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon " + vbCrlf
	sqlStr = sqlStr + " where idx = '" + CStr(bCpnIdx) + "' " + vbCrlf
	dbget.Execute sqlStr

	CheckAndCopyBonusCoupon = True

end function

function EditCSCopyCouponInfo(asid, copycouponinfo)
	dim sqlStr

    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr + " set copycouponinfo = '" & CStr(copycouponinfo) & "' "
    sqlStr = sqlStr + " where asid = " & CStr(asid) & " "
    dbget.Execute sqlStr
end function

' ��ǰ���� ���翩��     ' 2023.10.19 �ѿ�� ����
function EditCSCopyItemCouponInfo(asid, copyitemcouponinfo)
	dim sqlStr

	if asid="" or isnull(asid) or copyitemcouponinfo="" or isnull(copyitemcouponinfo) then exit function
    asid = trim(asid)

    sqlStr = " update [db_cs].[dbo].tbl_as_refund_info "
    sqlStr = sqlStr & " set copyitemcouponinfo = '" & copyitemcouponinfo & "' "
    sqlStr = sqlStr & " where asid = " & asid & " "
    
    'response.write sqlStr & "<br>"
    dbget.Execute sqlStr
end function

function SetNeedCheckToY(asid)
	dim sqlStr

    sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
    sqlStr = sqlStr + " set needChkYn = 'Y' "
    sqlStr = sqlStr + " where id = " & CStr(asid) & " "
    dbget.Execute sqlStr
end function

function CheckJungsanExists(orderserial)
    dim sqlStr

	CheckJungsanExists = False

    sqlStr = "select top 1 * from db_order.dbo.tbl_order_detail od "
    sqlStr = sqlStr + " Join db_jungsan.dbo.tbl_designer_jungsan_detail jd "
    sqlStr = sqlStr + " on od.idx=jd.detailidx "
    sqlStr = sqlStr + " where od.orderserial='" + CStr(orderserial) + "' "

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        CheckJungsanExists = True
    end if
    rsget.Close
end function

' ���곻�� üũ ' �̻� ����
function CheckJungsanExistsByAsid(asid)
    dim sqlStr

	CheckJungsanExistsByAsid = False

    sqlStr = "select top 1 *"
    sqlStr = sqlStr & " from [db_cs].[dbo].[tbl_new_as_detail] od with (nolock)"
    sqlStr = sqlStr & " Join db_jungsan.dbo.tbl_designer_jungsan_detail jd with (nolock)"
    sqlStr = sqlStr & "     on od.orderdetailidx=jd.detailidx "
    sqlStr = sqlStr & " where od.masterid = " & asid

    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        CheckJungsanExistsByAsid = True
    end if
    rsget.Close
end function

function CheckRestoreCancelOKByAsid(asid)
    dim sqlStr

	'// �߰���ۺ� �ִ� ���, ������ ��ҰǺ��� �����ؾ� ��.
	CheckRestoreCancelOKByAsid = False

    sqlStr = "select max(a.id) as maxid, isnull(sum(case when r.isRefundDeliveryPayAddedToOrder = 'Y' then 1 else 0 end),0) as cnt "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] m "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a on m.orderserial = a.orderserial "
    sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_as_refund_info r on a.id = r.asid "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and m.id = " & asid
    sqlStr = sqlStr + " 	and a.currstate >= 'B001' "
    sqlStr = sqlStr + " 	and a.divcd = 'A008' "
    sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		if (CLng(asid) = rsget("maxid")) or (rsget("cnt") = 0) then
			CheckRestoreCancelOKByAsid = True
		end if
    end if
    rsget.Close
end function

' ���� ������ü� ����       ' 2021.03.31 �ѿ�� ����
function Del_logicschulgodata(byval id, byval orderserial)
    dim sqlStr, resultCount, baljucount, tendb

    if orderserial="" or isnull(orderserial) then exit function
    baljucount = 0
    resultCount = 0
    IF application("Svr_Info")="Dev" THEN
        tendb="tendb."
    end if

    sqlStr = " select top 1 id "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_cs].[dbo].[tbl_new_as_list] a "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and a.id = " & id
    sqlStr = sqlStr & " 	and a.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr & " 	and a.divcd = 'A008' "
    sqlStr = sqlStr & " 	and a.requireupche = 'Y' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		'// �ٹ� �ֹ� �ƴ�
        Del_logicschulgodata = 0
        rsget.Close
        exit function
    end if
    rsget.Close

    sqlStr = " select top 1 a.id "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_cs].[dbo].[tbl_new_as_list] a "
    sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_detail] d on a.orderserial = d.orderserial "
    sqlStr = sqlStr & " 	left join [db_cs].[dbo].[tbl_new_as_detail] ad "
    sqlStr = sqlStr & " 	on "
    sqlStr = sqlStr & " 		1 = 1 "
    sqlStr = sqlStr & " 		and a.id = ad.masterid "
    sqlStr = sqlStr & " 		and d.idx = ad.orderdetailidx "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and a.id = " & id
    sqlStr = sqlStr & " 	and a.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr & " 	and d.itemid not in (0, 100) "
    sqlStr = sqlStr & " 	and d.itemno <> IsNull(ad.regitemno, 0) "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		'// �ֹ� ��ü ��� �� �ֹ����� �ƴ�
        Del_logicschulgodata = 0
        rsget.Close
        exit function
    end if
    rsget.Close

    sqlStr = "delete from db_order.[dbo].[tbl_baljudetail] where orderserial = '"& orderserial &"'"
    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr, resultCount

    sqlStr = "update [db_order].[dbo].[tbl_order_master] set baljudate=NULL where orderserial = '"& orderserial &"'"
    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    sqlStr = "update [db_order].[dbo].[tbl_order_detail] set currstate=0, upcheconfirmdate=NULL where orderserial = '"& orderserial &"' and isupchebeasong='N'"
    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    sqlStr = "delete from "&tendb&"[db_aLogistics].[dbo].[tbl_Logistics_order_detail] where orderserial = '"& orderserial &"'"
    'response.write sqlStr & "<Br>"
    dbget_Logistics.Execute sqlStr

    sqlStr = "delete from "&tendb&"[db_aLogistics].[dbo].[tbl_Logistics_order_master] where orderserial = '"& orderserial &"'"
    'response.write sqlStr & "<Br>"
    dbget_Logistics.Execute sqlStr

    sqlStr = "delete from "&tendb&"[db_aLogistics].[dbo].[tbl_Logistics_order_gift] where orderserial = '"& orderserial &"'"
    'response.write sqlStr & "<Br>"
    dbget_Logistics.Execute sqlStr

    sqlStr = "delete from "&tendb&"[db_aLogistics].[dbo].[tbl_Logistics_baljudetail] where orderserial = '"& orderserial &"'"
    'response.write sqlStr & "<Br>"
    dbget_Logistics.Execute sqlStr

    Del_logicschulgodata = resultCount
end function

function DeleteFinishedCSProcess(id)
    dim sqlStr, resultCount

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
    sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " , deletedate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)
    sqlStr = sqlStr + " and currstate='B007'"
	sqlStr = sqlStr + " and divcd in ('A004', 'A010', 'A008')"		'// ��ǰ, ȸ��, ���
	sqlStr = sqlStr + " and deleteyn = 'N' "

    dbget.Execute sqlStr, resultCount

    DeleteFinishedCSProcess = (resultCount>0)
end function

function DeleteFinishedCSForce(id)
    dim sqlStr, resultCount
    dim IsCsErrStockUpdateRequire
    IsCsErrStockUpdateRequire = False

    sqlStr = "select divcd, finishdate, currstate, deleteyn"
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		'// finishdate üũ �ϴ°� ��(2014-05-30 skyer9 : ����ó���Ϸ�)
        IsCsErrStockUpdateRequire = ((rsget("divcd")="A000") or (rsget("divcd")="A011") or (rsget("divcd")="A100") or (rsget("divcd")="A111")) and (rsget("currstate")="B007") and (rsget("deleteyn")="N")
    end if
    rsget.close

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
    sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " , deletedate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)
    sqlStr = sqlStr + " and currstate='B007'"
	''sqlStr = sqlStr + " and divcd in ('A004', 'A010', 'A008')"		'// ��ǰ, ȸ��, ���
	sqlStr = sqlStr + " and deleteyn = 'N' "

    dbget.Execute sqlStr, resultCount

    DeleteFinishedCSForce = (resultCount>0)

    if (IsCsErrStockUpdateRequire) then
        sqlStr = " exec db_summary.dbo.ten_RealTimeStock_CsErr " & id & ",'DEL','" & session("ssBctId") & "'"
        dbget.Execute sqlStr
    end if
end function

function CancelMinusOrderProcess(minusorderserial)
    dim sqlStr, resultCount

	CancelMinusOrderProcess = True

    sqlStr = " update db_order.dbo.tbl_order_master " + VbCrlf
    sqlStr = sqlStr + " set cancelyn = 'D'" + VbCrlf
    sqlStr = sqlStr + " , canceldate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where orderserial = '" + CStr(minusorderserial) + "' "
    sqlStr = sqlStr + " and cancelyn='N'"
	sqlStr = sqlStr + " and jumundiv = '9' "
    dbget.Execute sqlStr, resultCount

	if (resultCount < 1) then
		CancelMinusOrderProcess = False
		exit function
	end if

	dim userid, sitename

	userid = ""
	sitename = ""

    sqlStr = " select userid, sitename from [db_order].[dbo].tbl_order_master"
    sqlStr = sqlStr + " where orderserial='" + CStr(minusorderserial) + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		userid = rsget("userid")
		sitename = rsget("sitename")
    end if
    rsget.Close

	if (userid <> "") and (sitename = "10x10") then
		'
		sqlStr = " update [db_user].[dbo].tbl_mileagelog "
		sqlStr = sqlStr + " set deleteyn = 'Y' "
		sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and deleteyn = 'N' and orderserial = '" + CStr(minusorderserial) + "' "
		dbget.Execute sqlStr, resultCount

		Call updateUserMileage(userid)

		sqlStr = " update [db_user].[dbo].tbl_depositlog "
		sqlStr = sqlStr + " set deleteyn = 'Y' "
		sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and deleteyn = 'N' and orderserial = '" + CStr(minusorderserial) + "' "
		dbget.Execute sqlStr, resultCount

		Call updateUserDeposit(userid)

		sqlStr = " update [db_user].[dbo].tbl_giftcard_log "
		sqlStr = sqlStr + " set deleteyn = 'Y' "
		sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and deleteyn = 'N' and orderserial = '" + CStr(minusorderserial) + "' "
		dbget.Execute sqlStr, resultCount

		Call updateUserGiftCard(userid)

	end if

    ''���������� - ���������� ���� �ȵ�
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_minusOrder '" & minusorderserial & "'"
    dbget.Execute sqlStr

end function

' ��ǰ���� �±�ȯȸ�� ���� ó��     ' 2019.10.18 �ѿ�� ����
function CancelChangeOrderProcess(changeorderserial)
    dim sqlStr, resultCount

	CancelChangeOrderProcess = True

    sqlStr = " update db_order.dbo.tbl_order_master " + VbCrlf
    sqlStr = sqlStr + " set cancelyn = 'D'" + VbCrlf
    sqlStr = sqlStr + " , canceldate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where orderserial = '" + CStr(changeorderserial) + "' "
    sqlStr = sqlStr + " and cancelyn='N'"
    sqlStr = sqlStr + " and jumundiv = '6' "
    dbget.Execute sqlStr, resultCount

	if (resultCount < 1) then
		CancelChangeOrderProcess = False
		exit function
	end if

	dim userid, sitename

	userid = ""
	sitename = ""

    sqlStr = " select userid, sitename from [db_order].[dbo].tbl_order_master"
    sqlStr = sqlStr + " where orderserial='" + CStr(changeorderserial) + "'"

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
		userid = rsget("userid")
		sitename = rsget("sitename")
    end if
    rsget.Close

	if (userid <> "") and (sitename = "10x10") then
		'
		sqlStr = " update [db_user].[dbo].tbl_mileagelog "
		sqlStr = sqlStr + " set deleteyn = 'Y' "
		sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and deleteyn = 'N' and orderserial = '" + CStr(changeorderserial) + "' "
		dbget.Execute sqlStr, resultCount

		Call updateUserMileage(userid)

		sqlStr = " update [db_user].[dbo].tbl_depositlog "
		sqlStr = sqlStr + " set deleteyn = 'Y' "
		sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and deleteyn = 'N' and orderserial = '" + CStr(changeorderserial) + "' "
		dbget.Execute sqlStr, resultCount

		Call updateUserDeposit(userid)

		sqlStr = " update [db_user].[dbo].tbl_giftcard_log "
		sqlStr = sqlStr + " set deleteyn = 'Y' "
		sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and deleteyn = 'N' and orderserial = '" + CStr(changeorderserial) + "' "
		dbget.Execute sqlStr, resultCount

		Call updateUserGiftCard(userid)

	end if

    ''���������� - ���������� ���� �ȵ�
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_minusOrder '" & changeorderserial & "'"
    dbget.Execute sqlStr

end function

function RestoreCancelProcess(asid, orderserial)
	dim sqlStr, resultCount

	RestoreCancelProcess = False
	if (RestoreCancelValid(asid, orderserial) <> True) then
		exit function
	end if

	dim userid, couponsum, mileagesum, allatdiscountsum
	dim depositsum, giftcardsum
	dim refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum

    sqlStr = " select userid, IsNULL(tencardspend,0) as couponsum, IsNULL(miletotalprice,0) as mileagesum, IsNULL(allatdiscountprice,0) as allatdiscountsum " + vbCrLf
	sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master" + vbCrLf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        userid              = rsget("userid")
        couponsum      		= rsget("couponsum")
        mileagesum        	= rsget("mileagesum")
        allatdiscountsum  	= rsget("allatdiscountsum")
    end if
    rsget.close

    sqlStr = " select acctdiv, IsNull(realPayedsum, 0) as realPayedsum " + VbCrlf
    sqlStr = sqlStr + " from " + VbCrlf
    sqlStr = sqlStr + " db_order.dbo.tbl_order_PaymentEtc " + VbCrlf
    sqlStr = sqlStr + " where " + VbCrlf
    sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
    sqlStr = sqlStr + " 	and orderserial = '" + orderserial + "' " + VbCrlf
    sqlStr = sqlStr + " 	and acctdiv in ('200', '900') " + VbCrlf			'200 : ��ġ��, 900 : Giftī��
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	depositsum = 0
	giftcardsum = 0
	do until rsget.eof
		if (CStr(rsget("acctdiv")) = "200") then
			depositsum = rsget("realPayedsum")
		elseif (CStr(rsget("acctdiv")) = "900") then
			giftcardsum = rsget("realPayedsum")
		end if

		rsget.movenext
	loop
	rsget.close

    sqlStr = " select r.*, a.gubun01, a.gubun02 from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(asid)
    sqlStr = sqlStr + " and a.id=r.asid"
    ''sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate='B007'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    if Not rsget.Eof then
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")
		allatsubtractsum    = rsget("allatsubtractsum")
        refundgiftcardsum   = rsget("refundgiftcardsum")
        refunddepositsum    = rsget("refunddepositsum")
    else
        refundmileagesum    = 0
        refundcouponsum     = 0
		allatsubtractsum	= 0
        refundgiftcardsum   = 0
        refunddepositsum    = 0
    end if
    rsget.close

	dim IsUpdatedDeposit, IsUpdatedGiftCard
	IsUpdatedDeposit = False
	IsUpdatedGiftCard = False

	''RestoreCancelProcess = True
	''exit function

	if IsMasterCanceled(asid, orderserial) = True then
		''userid
		''mileagesum, depositsum, giftcardsum
		''refundmileagesum, refunddepositsum, refundgiftcardsum
		''refundcouponsum
		''couponsum
		'// ====================================================================
		if (userid <> "") and (mileagesum <> 0) then
			'' ��ü ����ΰ�� �ֹ��� ��ҷ� jukyocd : 2 ��ǰ����, 3 : �κ���ҽ� ȯ�����ϸ���
			sqlStr = " update [db_user].[dbo].tbl_mileagelog " + VbCrlf
			sqlStr = sqlStr + " set deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
			sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and deleteyn <> 'N'" + VbCrlf
			sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf
			dbget.Execute sqlStr
		end if

		if (userid <> "") and (depositsum <> 0) then
			'' ��ü ����ΰ�� �ֹ��� ��ҷ� jukyocd : 100 ��ǰ����, 10 : �κ���ҽ� ��ġ�� ȯ��
			sqlStr = " update [db_user].[dbo].tbl_depositlog " + VbCrlf
			sqlStr = sqlStr + " set deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
			sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and deleteyn <> 'N' " + VbCrlf
			sqlStr = sqlStr + " and jukyocd in ('100','10')" + VbCrlf					'100 : ��ǰ���Ż�� / 10 : �Ϻ�ȯ�� (���� : db_user.dbo.tbl_deposit_gubun)
			dbget.Execute sqlStr

			IsUpdatedDeposit = True
		end if

		if (userid <> "") and (giftcardsum <> 0) then
			'' ��ü ����ΰ�� �ֹ��� ��ҷ� jukyocd : 200 ��ǰ����, 300 : �κ���ҽ� Giftī�� ȯ��
			sqlStr = " update [db_user].[dbo].tbl_giftcard_log " + VbCrlf
			sqlStr = sqlStr + " set deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
			sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and deleteyn <> 'N' " + VbCrlf
			sqlStr = sqlStr + " and jukyocd in ('200','300')" + VbCrlf					'200 : ��ǰ���Ż�� / 300 : �Ϻ�ȯ�� (���� : db_user.dbo.tbl_giftcard_gubun)
			dbget.Execute sqlStr

			IsUpdatedGiftCard = True
		end if

		if (couponsum <> 0) then
			sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
			sqlStr = sqlStr + " set isusing='Y' "                   + VbCrlf
			sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
			dbget.Execute sqlStr
		end if

		'// ��� �ֹ� ����ȭ
		Call setRestoreCancelMaster(asid, orderserial)
	else
		'// ====================================================================
		if (userid <> "") and (refundmileagesum <> 0) then
			'' �κ� ����ε� ���ϸ��� ȯ���� ���.
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
			sqlStr = sqlStr + " set miletotalprice=miletotalprice + " + CStr(refundmileagesum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			dbget.Execute sqlStr

			sqlStr = " insert into [db_user].[dbo].tbl_mileagelog " + VbCrlf
			sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
			sqlStr = sqlStr + " values ("
			sqlStr = sqlStr + " '" + userid + "'"
			sqlStr = sqlStr + " ," + CStr(refundmileagesum) + ""
			sqlStr = sqlStr + " ,'2'"
			sqlStr = sqlStr + " ,'��ǰ���� ��� öȸ' "
			sqlStr = sqlStr + " ,'" + orderserial + "'"
			sqlStr = sqlStr + " ,'N'"
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr
		end if

		if (userid <> "") and (refunddepositsum <> 0) then
			'' �κ� ����ε� ��ġ�� ȯ���� ���.
			sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
			sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refunddepositsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and acctdiv='200'" + VbCrlf
			dbget.Execute sqlStr

			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
			sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refunddepositsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			dbget.Execute sqlStr

			sqlStr = " insert into [db_user].[dbo].tbl_depositlog " + VbCrlf
			sqlStr = sqlStr + " (userid, deposit, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
			sqlStr = sqlStr + " values ("
			sqlStr = sqlStr + " '" + userid + "'"
			sqlStr = sqlStr + " ," + CStr(refunddepositsum) + ""
			sqlStr = sqlStr + " ,'10'"
			sqlStr = sqlStr + " ,'��ǰ���� ��� öȸ'"
			sqlStr = sqlStr + " ,'" + orderserial + "'"
			sqlStr = sqlStr + " ,'N'"
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr

			IsUpdatedDeposit = True
		end if

		if (userid <> "") and (refundgiftcardsum <> 0) then
			'' �κ� ����ε� Giftī�� ȯ���� ���.
			sqlStr = " update [db_order].[dbo].tbl_order_PaymentEtc" + VbCrlf
			sqlStr = sqlStr + " set realPayedsum=realPayedsum + " + CStr(refundgiftcardsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and acctdiv='900'" + VbCrlf
			dbget.Execute sqlStr

			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
			sqlStr = sqlStr + " set sumPaymentETC=sumPaymentETC + " + CStr(refundgiftcardsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			dbget.Execute sqlStr

			sqlStr = " insert into [db_user].[dbo].tbl_giftcard_log " + VbCrlf
			sqlStr = sqlStr + " (userid, useCash, jukyocd, jukyo, orderserial, deleteyn, reguserid) " + VbCrlf
			sqlStr = sqlStr + " values ("
			sqlStr = sqlStr + " '" + userid + "'"
			sqlStr = sqlStr + " ," + CStr(refundgiftcardsum) + ""
			sqlStr = sqlStr + " ,'300'"
			sqlStr = sqlStr + " ,'��ǰ���� ��� öȸ'"
			sqlStr = sqlStr + " ,'" + orderserial + "'"
			sqlStr = sqlStr + " ,'N'"
			sqlStr = sqlStr + " ,'" + CStr(session("ssbctid")) + "'"
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr

			IsUpdatedGiftCard = True
		end if

		if (refundcouponsum <> 0) then
			'' �κ� ����ΰ�� - ȯ���� ��ŭ ��..
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
			sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(refundcouponsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			''response.write sqlStr
			dbget.Execute sqlStr

			sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
            sqlStr = sqlStr + " set isusing='Y' "                   + VbCrlf
            sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
            dbget.Execute sqlStr
		end if

		if (allatsubtractsum <> 0) then
			'' �κ� ����ΰ�� - ȯ���� ��ŭ ��..
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
			sqlStr = sqlStr + " set allatdiscountprice=allatdiscountprice + " + CStr(allatsubtractsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			''response.write sqlStr
			dbget.Execute sqlStr
		end if

		'// ��� ��ǰ ����ȭ
		Call setRestoreCancelDetail(asid, orderserial)

		'// �߰���ۺ� ������ ���
		Call CancelAddBeasongpayForCancel(asid)

		Call reCalcuOrderMaster(orderserial)

	end if

    ''���ϸ����� �ֹ��� ��� �� �����ؾ���.
    '��ġ��, Giftī�� ����
    if (userid<>"") then
        Call updateUserMileage(userid)

        if IsUpdatedDeposit then
        	Call updateUserDeposit(userid)
        end if

        if IsUpdatedGiftCard then
        	Call updateUserGiftCard(userid)
        end if
    end if

	RestoreCancelProcess = True

end function

function RestoreCancelValid(asid, orderserial)
	dim sqlStr, resultCount
	dim ipkumdiv

	RestoreCancelValid = True

	if (GC_IsOLDOrder = True) then
		response.write "ERROR : 6���� ���� �ֹ� ó���Ұ�"
		RestoreCancelValid = False
		exit function
	end if

    sqlStr = " select m.ipkumdiv "
    sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	ipkumdiv = 8
    if Not rsget.Eof then
		ipkumdiv = rsget("ipkumdiv")
    end if
    rsget.close

	if (ipkumdiv > 7) then
		sqlStr = " update [db_order].[dbo].tbl_order_master "
		sqlStr = sqlStr + " set ipkumdiv = '7' "
		sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
		dbget.Execute sqlStr

		''response.write "ERROR : ���Ϸ� �ֹ�"
		''RestoreCancelValid = False
		''exit function
	end if

end function

'������ ������ ��� Flag �ٸ��� ��������
'��ۺ� ���
function setRestoreCancelDetail(Asid, orderserial)
    dim sqlStr

    '�������� - ��ǰ�Ϻ�����ΰ��
	sqlStr = " update d "
	sqlStr = sqlStr + " set d.itemno = d.itemno + c.regitemno "
	sqlStr = sqlStr + " 	from "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 		join [db_cs].[dbo].tbl_new_as_detail c "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and d.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 			and c.masterid = " + CStr(Asid) + " "
	sqlStr = sqlStr + " 			and d.idx = c.orderdetailidx "
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and d.itemid <> 0 "				'// ��ۺ�� ������ �ٸ� �� ����.(������ 1��)
    dbget.Execute sqlStr

	sqlStr = " update d "
	sqlStr = sqlStr + " set d.cancelyn = 'N' "
	sqlStr = sqlStr + " 	from "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 		join [db_cs].[dbo].tbl_new_as_detail c "
	sqlStr = sqlStr + " 		on "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and d.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 			and c.masterid = " + CStr(Asid) + " "
	sqlStr = sqlStr + " 			and d.idx = c.orderdetailidx "
	sqlStr = sqlStr + " 			and d.itemno = c.regitemno "
	sqlStr = sqlStr + " 			and d.cancelyn = 'Y' "
    dbget.Execute sqlStr

end function

'//�ٹ����� ��� ��ǰ cs ���� ó��		'/2016.07.18 �ѿ�� ����
function setItemLimitcs(Asid, orderserial, updowngubun)
    dim sqlStr, divcd

	if updowngubun="" or Asid="" or orderserial="" then exit function

    divcd = ""

    sqlStr = " select top 1"
    sqlStr = sqlStr & " m.divcd"
    sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_list m"
    sqlStr = sqlStr & " where m.id=" & asid & ""
    sqlStr = sqlStr & " and m.divcd in ('A010', 'A011', 'A111')"
    sqlStr = sqlStr & " and m.deleteyn='N'"

	'response.write sqlStr & "<Br>"
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.Eof) then
        divcd = rsget("divcd")
    end if
    rsget.Close

	if divcd = ""  then Exit function

	'/��ǰ ���� ó��
	sqlStr = "update i set" & vbcrlf

	if updowngubun="UP" then
		sqlStr = sqlStr & " i.limitsold=(case when 0>i.limitsold - T.itemno then 0 else i.limitsold - T.itemno end)" & vbcrlf
	elseif updowngubun="DOWN" then
		sqlStr = sqlStr & " i.limitsold=(case when 0>i.limitsold + T.itemno then 0 else i.limitsold + T.itemno end)" & vbcrlf
	end if

	sqlStr = sqlStr & " from db_item.dbo.tbl_item i" & vbCrLf
	sqlStr = sqlStr & " join (" & vbCrLf
	sqlStr = sqlStr & " 	select" & vbCrLf
	sqlStr = sqlStr & " 	nd.itemid, sum(isnull(nd.confirmitemno,0)) as itemno" & vbCrLf
	sqlStr = sqlStr & " 	from [db_cs].[dbo].[tbl_new_as_detail] nd" & vbCrLf
	sqlStr = sqlStr & " 	where nd.masterid="& Asid &"" & vbcrlf
	sqlStr = sqlStr & " 	and nd.itemid not in (0,100)" & vbcrlf	' ��ۺ�, ����� ����

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " 	group by nd.itemid " & vbCrLf
	sqlStr = sqlStr & " ) T" & vbCrLf
	sqlStr = sqlStr & " 	on i.itemid = T.itemid " & vbCrLf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/�ɼ� �ִ� ��ǰ ���� ó��
	sqlStr = "update o set" & vbcrlf

	if updowngubun="UP" then
		sqlStr = sqlStr & " o.optlimitsold=(case when 0>o.optlimitsold - nd.confirmitemno then 0 else o.optlimitsold - nd.confirmitemno end)" & vbcrlf
	elseif updowngubun="DOWN" then
		sqlStr = sqlStr & " o.optlimitsold=(case when 0>o.optlimitsold + nd.confirmitemno then 0 else o.optlimitsold + nd.confirmitemno end)" & vbcrlf
	end if

	sqlStr = sqlStr & " from [db_cs].[dbo].[tbl_new_as_detail] nd" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item_option o" & vbcrlf
	sqlStr = sqlStr & " 	on nd.itemid=o.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and nd.itemoption = o.itemoption" & vbcrlf
	sqlStr = sqlStr & " where nd.masterid="& Asid &"" & vbcrlf

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " and nd.itemid not in (0,100)" & vbcrlf	' ��ۺ�, ����� ����
	sqlStr = sqlStr & " and nd.itemoption<>'0000'" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/��ǰ ǰ��ó�� �Ǹ����� ��ǰ�� ����ǰ�� �� ��ǰ �Ͻ� ǰ���� ����
	sqlStr = "update i" & vbcrlf
	sqlStr = sqlStr & " set i.sellyn='S' , i.lastupdate=getdate()" & vbcrlf
	sqlStr = sqlStr & " from (" & vbcrlf
	sqlStr = sqlStr & " 	select nd.itemid" & vbcrlf
	sqlStr = sqlStr & " 	from [db_cs].[dbo].[tbl_new_as_detail] nd" & vbcrlf
	sqlStr = sqlStr & " 	where nd.masterid="& Asid &"" & vbcrlf

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " 	and nd.itemid not in (0,100)" & vbcrlf	' ��ۺ�, ����� ����
	sqlStr = sqlStr & " 	group by nd.itemid" & vbcrlf
	sqlStr = sqlStr & " ) as t" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i" & vbcrlf
	sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.sellyn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and i.limityn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and (i.limitno-i.limitSold<1)" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/�Ͻ� ǰ���̳� ��������>0 ��� �Ǹŷ� ����
	sqlStr = "update i" & vbcrlf
	sqlStr = sqlStr & " set i.sellyn='Y' , i.lastupdate=getdate()" & vbcrlf
	sqlStr = sqlStr & " from (" & vbcrlf
	sqlStr = sqlStr & " 	select nd.itemid" & vbcrlf
	sqlStr = sqlStr & " 	from [db_cs].[dbo].[tbl_new_as_detail] nd" & vbcrlf
	sqlStr = sqlStr & " 	where nd.masterid="& Asid &"" & vbcrlf

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " 	and nd.itemid not in (0,100)" & vbcrlf	' ��ۺ�, ����� ����
	sqlStr = sqlStr & " 	group by nd.itemid" & vbcrlf
	sqlStr = sqlStr & " ) as t" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i" & vbcrlf
	sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.sellyn='S'" & vbcrlf
	sqlStr = sqlStr & " 	and i.limityn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and (i.limitno-i.limitSold>0)" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr
end function

%>
