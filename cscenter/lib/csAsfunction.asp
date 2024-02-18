<%
'###########################################################
' Description : cs센터 공용함수
' History : 이상구 생성
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

	returnmethod			= "R007"			'무통장 환불

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
	'원주문내역
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
		orgbeasongpay			= 0								'나중에 나눈다.
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
	'배송비
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
	'200 : 예치금, 900 : Gift카드
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

		contents_finish = contents_finish + vbCrLf + "마일리지 환급 : " + CStr(refundmileagesum*-1)
	end if

	if ((orgsubtotalprice - sumPaymentEtc) < refundrequire) and (orgdepositsum > 0) then
		if (refundrequire >= orgdepositsum) then
			refunddepositsum = orgdepositsum
			refundrequire = refundrequire - orgdepositsum
		else
			refunddepositsum = refundrequire
			refundrequire = 0
		end if

		contents_finish = contents_finish + vbCrLf + "예치금 환급 : " + CStr(refunddepositsum*-1)
	end if

	if ((orgsubtotalprice - sumPaymentEtc) < refundrequire) and (orggiftcardsum > 0) then
		if (refundrequire >= orggiftcardsum) then
			refundgiftcardsum = orggiftcardsum
			refundrequire = refundrequire - orggiftcardsum
		else
			refundgiftcardsum = refundrequire
			refundrequire = 0
		end if

		contents_finish = contents_finish + vbCrLf + "Gift카드 환급 : " + CStr(refundgiftcardsum*-1)
	end if

	if (ipkumdiv < 4) then
		refundrequire = 0
	end if

	if (refundrequire <= 0) then
		returnmethod = "R000"
	end if

	'==========================================================================
    'CS Master 환불 관련정보 저장
    Call RegCSMasterRefundInfo(asid, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum*-1, refundcouponsum*-1, allatsubtractsum*-1, refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
    Call AddCSMasterRefundInfo(asid, orggiftcardsum, orgdepositsum, refundgiftcardsum*-1, refunddepositsum*-1)

	RegCSMasterRefundInfoBeforeCancel = 0
	if (refundrequire > 0) and (ipkumdiv >= 4) then
        '환불 정보가 있는지 체크 후 무통장환불/마일리지환불/신용카드취소 CS 접수 등록
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
    ''과거 주문인지 Check
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
		CheckFreeReturnDeliveryAvail = "이벤트 기간 아님[" & startDate & "~" & endDate & "]"
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
		CheckFreeReturnDeliveryAvail = "출고상품 금액 부족[" & reducedPriceSUM & "원]"
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
		CheckFreeReturnDeliveryAvail = "주문당 한번만 가능"
		Exit Function
	End If

	CheckFreeReturnDeliveryAvail = ""

end function

function getCardRibonName(cardribbon)
    if IsNULL(cardribbon) then Exit Function

    if (cardribbon="1") then
        getCardRibonName  = "카드"
    elseif (cardribbon="2") then
        getCardRibonName  = "리본"
    elseif (cardribbon="3") then
        getCardRibonName  = "없음"
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
		'// finishdate 체크 하는것 뺌(2014-05-30 skyer9 : 물류처리완료)
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

    ''맞교환회수 완료일경우 재고없데이트. 2007.11.16
	'// 업배는 처리하는 것 없음.(2017-01-31, skyer9)
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
		'// finishdate 체크 하는것 뺌(2014-05-30 skyer9 : 물류처리완료)
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

    ''맞교환회수 완료일경우 재고없데이트. 2007.11.16
	'// 업배는 처리하는 것 없음.(2017-01-31, skyer9)
    if (IsCsErrStockUpdateRequire) then
        ''sqlStr = " exec db_summary.dbo.ten_RealTimeStock_CsErr " & iAsid & ",'','" & finishuser & "'"
        ''dbget_TPL.Execute sqlStr
    end if
end function

function SetStockOutByCsAs(iAsid)
    dim sqlStr
    dim resultCount	: resultCount = 0
    dim arrItemID

	'// 업배상품만 품절 등록

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
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
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
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
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
    '// 1. 옵션 없는 상품(일시품절 전환)
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
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
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
	'// 2-1. 옵션 있는 상품(상품코드목록)
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
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
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

	'// 2-2. 옵션 있는 상품(품절전환)
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
	sqlStr = sqlStr + " join [db_temp].[dbo].[tbl_mibeasong_list] T "		'// 미출고사유 품절일 경우만, 2022-02-24, skyer9
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

	'// 2-3. 옵션 있는 상품(옵션갯수)
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

	'// 2-4. 옵션 있는 상품(판매중인 옵션이 없으면 품절처리)
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
        ipkumdivName = "입금 대기"
    elseif (ipkumdiv="4") then
        ipkumdivName = "결제 완료"
    elseif (ipkumdiv="5") then
        ipkumdivName = "상품 준비"
    elseif (ipkumdiv="6") then
        ipkumdivName = "출고 준비"
    elseif (ipkumdiv="7") then
        ipkumdivName = "일부 출고"
    elseif (ipkumdiv="8") then
        ipkumdivName = "출고 완료"
    end if

    if (accountdiv="7") then
        accountdivName = "무통장"
    elseif (accountdiv="14") then
        accountdivName = "편의점결제"
    elseif (accountdiv="100") then
        accountdivName = "신용카드"
    elseif (accountdiv="550") then
        accountdivName = "기프팅"
    elseif (accountdiv="560") then
        accountdivName = "기프티콘"
    elseif (accountdiv="80") then
        accountdivName = "올엣카드"
    elseif (accountdiv="50") then
        accountdivName = "제휴몰결제"
    elseif (accountdiv="20") then
        accountdivName = "실시간이체"
    elseif (accountdiv="150") then
        accountdivName = "이니렌탈"
    end if

    ''2016/08/04
    if (pggubun="NP") then
        accountdivName = "네이버페이"
        if (comm_cd="A007") then
            comm_name = "네이버페이 취소요청"
        end if
    end if

    ''취소만..
    if (divcd="A007") or (divcd="A008") then
        GetDefaultTitle = accountdivName + " " + ipkumdivName + " 상태 중 " + comm_name
    else
        GetDefaultTitle = comm_name
    end if
end function

function AddCsMemoWithMemoGubun(orderserial,divcd,userid,writeuser,contents_jupsu,mmgubun)
	dim sqlStr

	if divcd="1" then
        ''일반메모
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate())"

        dbget.Execute sqlStr
    else
        ''처리요청메모
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

        dbget.Execute sqlStr
    end if
end function

function AddCsMemo(orderserial,divcd,userid,writeuser,contents_jupsu)
    dim sqlStr
    dim mmgubun ''메모구분
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
        ''일반메모
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate, phoneNumber)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate(), '" + CStr(phoneNumber) + "')"

        dbget.Execute sqlStr
    else
        ''처리요청메모
        sqlStr = "insert into [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

        dbget.Execute sqlStr
    end if

end function

function AddCsMemoRequest(orderserial, userid, qadiv, writeuser, contents_jupsu)
    dim sqlStr

    ''처리요청메모
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

    '// SQL 인젝션??
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
    '' CS Master 저장
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
	sqlStr = sqlStr + " , extsitename=(CASE WHEN T.sitename<>'10x10' THEN T.sitename ELSE NULL END)"   + VbCrlf   ''2011-06-14 추가

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

	''회수신청 접수인경우 - 기본 회수 배송지 저장
	''맞교환, 서비스 발송, 누락발송, 기타회수
	if (divcd="A010") or (divcd="A010") or (divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002") or (divcd="A200") then
	    Call RegDefaultDEliverInfo(InsertedId, orderserial)
    end if

	if (Not IsNumeric(orderserial)) and (IsUpdateSuccess = False) then
		'Gift카드 주문인지 확인한다
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

''기본 회수/맞교환/서비스발송 주소지 입력 - 접수시 주문번호 기본 주소지로 저장됨. - 저장후 수정하는 Procsess
function RegDefaultDEliverInfo(AsID, orderserial)
    dim sqlStr
    sqlStr = "insert into [db_cs].[dbo].tbl_new_as_delivery"
    sqlStr = sqlStr + " (asid, reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqetcaddr)"
    ''sqlStr = sqlStr + " select " + CStr(AsID) + ",reqname, reqphone, reqhp, reqzipcode, reqaddress, reqzipaddr" ''바꼈음.
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
    '' CS Master 수정
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
    '' CS 환불 제목
	'// GetIsCSServiceRefund() 와 함께 바꿀것
    dim tmptitle

	GetCSRefundTitle = orgtitle

    if (divcd <> "A003") and (divcd <> "A007") then
    	Exit function
    end if

    if (InStr("주문 취소 후,반품 처리 후,회수 처리 후", Left(orgtitle, 7)) > 0) then
		GetCSRefundTitle = Left(orgtitle, 7) & " " & GetRefundMethodString(returnmethod)
    end if

	if (Left(orgtitle, 7) = "CS서비스 -") then
		GetCSRefundTitle = Left(orgtitle, 7) & " " & GetRefundMethodString(returnmethod)
	end if

    if ((orgtitle = "마일리지 적립(CS서비스)") or (orgtitle = "환불(무통장)") or (orgtitle = "환불(예치금)") or (orgtitle = "환불(마일리지)") or (orgtitle = "예치금 적립(품절)") or (orgtitle = "마일리지 적립(품절)") or (orgtitle = "마일리지 적립(배송지연)")) then
    	GetCSRefundTitle = "CS서비스 -" & " " & GetRefundMethodString(returnmethod)
    end if

end function

function GetIsCSServiceRefund(AsID, divcd, orgtitle)
	'// GetCSRefundTitle() 와 함께 바꿀것
	GetIsCSServiceRefund = False

    if (divcd <> "A003") and (divcd <> "A007") then
    	Exit function
    end if

    if (InStr("주문 취소 후,반품 처리 후,회수 처리 후", Left(orgtitle, 7)) > 0) then
		Exit function
    end if

	if (Left(orgtitle, 7) = "CS서비스 -") then
		GetIsCSServiceRefund = True
		Exit function
	end if

    if ((orgtitle = "마일리지 적립(CS서비스)") or (orgtitle = "환불(무통장)") or (orgtitle = "환불(예치금)") or (orgtitle = "환불(마일리지)") or (orgtitle = "예치금 적립(품절)") or (orgtitle = "마일리지 적립(품절)") or (orgtitle = "마일리지 적립(배송지연)")) then
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

    'R007 무통장환불
    'R020 실시간이체취소
    'R050 입점몰결제 취소
    'R080 올엣카드취소
    'R100 신용카드취소
    'R550 기프팅취소
    'R560 기프티콘취소
    'R120 신용카드부분취소
    'R400 휴대폰취소
	'R420 휴대폰부분취소
    'R900 마일리지로환불
    'R910 예치금환불
    'R022 실시간이체부분취소(NP)
    'R150 이니렌탈취소

	tmpstr = ""

    if (returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420") or (returnmethod="R150") then
        if (returnmethod="R020") then
            tmpstr = "실시간이체취소"
        elseif (returnmethod="R022") then ''2016/07/21
            tmpstr = "실시간이체부분취소"
        elseif (returnmethod="R080") then
            tmpstr = "올엣카드취소"
        elseif (returnmethod="R100") then
            tmpstr = "신용카드취소"
        elseif (returnmethod="R550") then
            tmpstr = "기프팅취소"
        elseif (returnmethod="R560") then
            tmpstr = "기프티콘취소"
        elseif (returnmethod="R120") then
            tmpstr = "신용카드부분취소"
		elseif (returnmethod="R400") then
            tmpstr = "휴대폰취소"
        elseif (returnmethod="R420") then
            tmpstr = "휴대폰부분취소"
        elseif (returnmethod="R150") then
            tmpstr = "이니렌탈취소"
        end if
    elseif (returnmethod="R050") then
        tmpstr = "입점몰결제 취소"
    elseif (returnmethod="R900") then
        tmpstr = "마일리지 환불"
    elseif (returnmethod="R910") then
        tmpstr = "예치금 환불"
    elseif (returnmethod<>"") then
        tmpstr = "무통장 환불"
    end if

	GetRefundMethodString = tmpstr

end function

function EditCSMasterFinished(AsID, title, contents_jupsu, gubun01, gubun02, finishuserid, contents_finish)
    '' CS Master 완료된 내역 수정
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
' 	Array("2","이"), _
' 	Array("3","삼"), _
' 	Array("4","사") )
function EditCSMasterAddInfo(AsID, addInfoArr)
    '' CS Master 추가정보 입력
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
    ''전체 취소 환불정보 복사

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

    ''취소 관련정보
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

    ''취소 관련정보
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

    '관련 CS REF KEY*******************************
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

    ''2017/10/02 암호화 방식 변경
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

    ''기타 정산 추가인경우만 makerid 지정 : 반품접수(업체배송) / 맞교환(업체)인 경우는 기 지정됨.
    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" & VbCrlf
    sqlStr = sqlStr + " set makerid='" & buf_requiremakerid & "'" & VbCrlf
    sqlStr = sqlStr + " where id=" & iasid & "" & VbCrlf
    sqlStr = sqlStr + " and divcd='A700'" & VbCrlf

    dbget.Execute sqlStr

end function

function setRestoreEtcRealPayment(asid, orderserial)
    dim sqlStr

    '// TODO : 현재 신용카드금액만 복구
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

    ''기타 정산 추가인경우만 makerid 지정 : 반품접수(업체배송) / 맞교환(업체)인 경우는 기 지정됨.
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

'// 원 주문내역에 없는 상품 등록(상품변경 맞교환)
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

	'// 상품정보 복사
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

	''가격정보
	''옵션변경 : 현재 옵션가 동일해야만 맞교환 가능, 가격정보 그대로 카피
	''상품변경 : 브랜드,현재 판매가(할인가),매입가, 쿠폰적용가능 등 모두 동일해야하고 1:1 변경만 가능(수량은 여러개 가능), 가격정보 그대로 카피
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

'''마이너스 주문건을 따고 접수하는것이 맞는건지..
function ChangeReturnItems(orderserial, id, newasid, pitemid, pitemoption, citemid, citemoption, citemno ,byRef ScanErr)
    ''id : 맞교환 출고id, newasid : 맞교환 회수id, pitemid :원상품ID citemid : 변경상품ID

	response.write detailitemlist + "에러 : 시스템팀 문의"
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

    ''검토 // 접수된 수량 합계금액과 일치 해야함.
    for i = 0 to UBound(buf_citemid)
		if (TRIM(buf_citemid(i)) <> "") then  ''변경된 상품번호.
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
        ScanErr = "교환금액이 일치하지 않습니다.\n접수금액 :" &regedSum & "\n교환금액 :"&curritemcostsum
        Exit Function
    End IF

    ''''////MinusOrderSerial = AddMinusOrder(id,orderserial)

    ''맞교환 출고건.
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
    sqlStr = sqlStr + " ,IsNULL(d.itemname,'배송료')"
    sqlStr = sqlStr + " ,IsNULL(d.itemoptionname,(case when d.itemcost=0 then '무료배송' else '일반택배' end))"
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

''바로 완료 처리로 진행 할지 여부.
function IsDirectProceedFinish(divcd, Asid, orderserial, byRef EtcStr)
    dim sqlStr
    dim cancelyn, ipkumdiv
    IsDirectProceedFinish = false

    '' currstate:2 업체(물류) 통보
    if (divcd="A008") then
        ''' 취소 Case
        '' 등록된 상품중 업체 확인중 상태가 있으면 접수상태로 진행
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
	'검증 - 전체취소 접수
	'
	' - 전체취소가 맞는지		- rsget("itemno") = rsget("totalcancelregno") 인가
	' - 이중취소인지			- d.cancelyn = 'Y' 에 대한 취소
	' - 초과취소인지			- rsget("itemno") < rsget("totalcancelregno") 인가
	' - 마스터 취소 되었는지	- m.cancelyn = 'Y' 에 대한 취소

    dim sqlStr, result
    GetAllCancelRegValidResult = ""
    result = ""

	'==========================================================================
	' - 마스터 취소 되었는지
	'==========================================================================
	if (IsMasterCanceled(Asid, orderserial)) then
		GetAllCancelRegValidResult = "취소된 주문입니다."
		exit function
	end if

	'==========================================================================
	'전체취소가 맞는지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량과 같은지
	'==========================================================================
	if (Not IsAllCancelState(Asid, orderserial)) then
		if (IsErrorCancelState(Asid, orderserial)) then
			GetAllCancelRegValidResult = "주문수량을 초과하여 취소(CS접수 포함)된 상품이 있습니다."
		else
			GetAllCancelRegValidResult = "결재금액 전부환불이면서 전부취소(CS접수 포함)가 아닌 상품이 있습니다."
		end if
		exit function
	end if

	'==========================================================================
	'이중취소인지 - 취소된 디테일에 대한 취소가 있는지
	'==========================================================================
	if (IsDoubleCancelState(Asid, orderserial)) then
		GetAllCancelRegValidResult = "취소된 상품에 대한 취소가 있습니다."
		exit function
	end if

end function

function GetPartialCancelRegValidResult(Asid, orderserial)
	'검증 - 일부취소 접수
	'
	' - 부분취소인지
	' - 초과취소인지
	' - 이중취소인지
	' - 마스터 취소 되었는지

    dim sqlStr, result
    GetPartialCancelRegValidResult = ""
    result = ""

	'==========================================================================
	' - 마스터 취소 되었는지
	'==========================================================================
	if (IsMasterCanceled(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "취소된 주문입니다."
		exit function
	end if

	'==========================================================================
	'부분취소인지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량보다 작은것이 있는지
	'초과취소인지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량보다 큰것이 있는지
	'==========================================================================
	if (IsAllCancelState(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "주문전체 취소(CS접수 포함)이면서 결재금액과 환불금액 합계가 다릅니다. - 마일리지/할인권 환원을 체크하세요."
		exit function
	else
		if (IsErrorCancelState(Asid, orderserial)) then
			GetPartialCancelRegValidResult = "주문수량을 초과하여 취소(CS접수 포함)된 상품이 있습니다."
			exit function
		end if
	end if

	'==========================================================================
	'이중취소인지 - 취소된 디테일에 대한 취소가 있는지
	'==========================================================================
	if (IsDoubleCancelState(Asid, orderserial)) then
		GetPartialCancelRegValidResult = "취소된 상품에 대한 취소가 있습니다."
		exit function
	end if

end function

'전체취소 상태인지
function IsAllCancelState(Asid, orderserial)
    dim sqlStr, result
    IsAllCancelState = true

	'==========================================================================
	'전체취소가 맞는지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량과 같은지
	'd.cancelyn = 'Y' 인 상품은 이중취소에서, rsget("itemno") < rsget("totalcancelregno") 인 것은 초과취소에서 체크한다.
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

'초과취소 상태인지
function IsErrorCancelState(Asid, orderserial)
    dim sqlStr, result
    IsErrorCancelState = false

	'==========================================================================
	'초과취소인지 - 디테일 취소 접수(CS처리완료제외) 전체의 합이 잔여주문수량보다 큰지
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

'디테일 이중취소 있는지
function IsDoubleCancelState(Asid, orderserial)
    dim sqlStr, result
    IsDoubleCancelState = false

	'==========================================================================
	'이중취소인지 - 취소된 디테일에 대한 취소가 있는지
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
	' - 마스터 취소 되었는지
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

''주문 상세 내역이 취소 가능한지 체크 - 출고 완료된 내역이 있는지, 주문건이 취소된내역이 있는지
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

'// 매입상품 있는지
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

''반품/ 회수 접수내역 체크
function IsReturnRegValid(Asid, orderserial,byref ScanErr, upcheMakerid)
    ''  업체배송과 자체배송을 같이 접수하지 못함.
    ''  업체배송이 존재할 경우 MakerID가 1개만 존재 해야함.

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
            ScanErr = "텐바이텐 배송과 업체배송을 동시에 접수하실 수 없습니다."
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
            ScanErr = "업체배송의 경우 각 브랜드 별로 접수하셔야 합니다."
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
    '' 취소일 추가
    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
	'' 발주일 입력안된 경우 발주일 입력, skyer9, 2018-02-26
	sqlStr = sqlStr + " ,baljudate=IsNULL(baljudate,getdate())" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr
end function

' 취소주문 정상화. 원상복구
function setRestoreCancelMaster(Asid, orderserial)
    dim sqlStr

	sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
    sqlStr = sqlStr + " set cancelyn='N'" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
    dbget.Execute sqlStr
end function

'수량이 같으면 취소 Flag 다르면 수량변경
'배송비도 취소
function setCancelDetail(Asid, orderserial)
    dim sqlStr
    ''취소일 추가
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

    '수량변경 - 상품일부취소인경우
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
	sqlStr = sqlStr + " 			and d.itemid <> 0 "				'// 배송비는 수량이 다를 수 없다.(언제나 1개)
    dbget.Execute sqlStr

    '// 품절수량 차감
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



''주문 마스타 재계산
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

	'// e.acctdiv = '120' 네이버 포인트
	'// 참조 주문번호 : 16092146018
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

	'// 보너스/사용마일리지 요약 재계산(신규Proc)
	sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&userid&"'"
	dbget.Execute sqlStr

	'// 주문마일리지 요약 재계산(기존Proc:변경없음)
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
	'예치금 재계산
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

	'데이타가 없으면 생성한다.
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
	sqlStr = sqlStr + " select userid,deposit*-1,jukyocd,jukyo+' 취소',orderserial,deleteyn,asid "
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
    sqlStr = sqlStr + " 	and acctdiv in ('200', '900') " + VbCrlf			'200 : 예치금, 900 : Gift카드

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

'' 마일리지 환원

    IsUpdatedMile = false
response.write "3" & "<br>"
    if (userid<>"") and (IsAllCancel) and (miletotalprice<>0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 2 상품구매, 3 : 부분취소시 환원마일리지
        sqlStr = " update [db_user].[dbo].tbl_mileagelog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf
        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환원 : " & miletotalprice
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환원 : " & miletotalprice
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundmileagesum<>0) then
        '' 부분 취소인데 마일리지 환원할 경우.
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
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환원 : " & refundmileagesum
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환원 : " & refundmileagesum
        end if
    end if

'예치금환원
	IsUpdatedDeposit = false
    if (userid<>"") and (IsAllCancel) and (orgdepositsum <> 0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 100 상품구매, 10 : 부분취소시 예치금 환원
        sqlStr = " update [db_user].[dbo].tbl_depositlog " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('100','10')" + VbCrlf					'100 : 상품구매사용 / 10 : 일부환원 (참고 : db_user.dbo.tbl_deposit_gubun)
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "사용 예치금 환원 : " & orgdepositsum
        else
            openMessage = openMessage + VbCrlf + "사용 예치금 환원 : " & orgdepositsum
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refunddepositsum <> 0) then
        '' 부분 취소인데 예치금 환원할 경우.

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
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedDeposit = true

        if openMessage="" then
            openMessage = openMessage + "사용 예치금 환원 : " & refunddepositsum
        else
            openMessage = openMessage + VbCrlf + "사용 예치금 환원 : " & refunddepositsum
        end if
    end if

'Gift카드환원
	IsUpdatedGiftCard = false
    if (userid<>"") and (IsAllCancel) and (orggiftcardsum <> 0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 200 상품구매, 300 : 부분취소시 Gift카드 환원
        sqlStr = " update [db_user].[dbo].tbl_giftcard_log " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('200','300')" + VbCrlf					'200 : 상품구매사용 / 300 : 일부환원 (참고 : db_user.dbo.tbl_giftcard_gubun)
        dbget.Execute sqlStr

        IsUpdatedGiftCard = true

        if openMessage="" then
            openMessage = openMessage + "사용 Gift카드 환원 : " & orggiftcardsum
        else
            openMessage = openMessage + VbCrlf + "사용 Gift카드 환원 : " & orggiftcardsum
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundgiftcardsum <> 0) then
        '' 부분 취소인데 Gift카드 환원할 경우.

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
        sqlStr = sqlStr + " ,'상품구매 취소 환원'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " ,'" + CStr(session("ssbctid")) + "'"
        sqlStr = sqlStr + " )"
        dbget.Execute sqlStr

        IsUpdatedGiftCard = true

        if openMessage="" then
            openMessage = openMessage + "사용 Gift카드 환원 : " & refundgiftcardsum
        else
            openMessage = openMessage + VbCrlf + "사용 Gift카드 환원 : " & refundgiftcardsum
        end if
    end if

'' 할인권 환급
    if (IsAllCancel) and (tencardspend<>0) then
        sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
	    sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
	    sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

	    dbget.Execute sqlStr

	    if openMessage="" then
            openMessage = openMessage + "사용 보너스쿠폰 환급"
        else
            openMessage = openMessage + VbCrlf + "사용 보너스쿠폰 환급"
        end if
    end if

    if (Not IsAllCancel) and (refundcouponsum<>0) then
        '' 부분 취소인경우 - 환급한 만큼 깜..
		if (GC_IsOLDOrder) then
			sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
		else
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
		end if
        sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(refundcouponsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        ''전체 환급인 경우만 쿠폰을 돌려줌
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

        ''원래 할인권 사용액이 있고, 남은 쿠폰사용액이 없을경우 전체  환급
        if (tencardspend>0) then
            if (remaintencardspend=0)   then
                sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
            	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
            	sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

            	dbget.Execute sqlStr

            	if openMessage="" then
                    openMessage = openMessage + "사용 할인권  환급"
                else
                    openMessage = openMessage + VbCrlf + "사용 할인권  환급"
                end if
            else
                ''(또는, %쿠폰인 경우 공통,단순변심인 경우 제외하고 환급해줌./ 부분취소 ) C004 CD01
                if (ipkumdiv>3) and (Not ((gubun01="C004") and (gubun02="CD01"))) then
                    sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
                	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
                	sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
                	sqlStr = sqlStr + " and coupontype=1"

                	dbget.Execute sqlStr

                	if openMessage="" then
                        openMessage = openMessage + "사용 할인권  환급."
                    else
                        openMessage = openMessage + VbCrlf + "사용 할인권  환급."
                    end if
                end if
            end if
        end if

    end if

    '' 올엣카드 할인 차감
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
            openMessage = openMessage + "올엣카드 할인 차감 : " & allatsubtractsum
        else
            openMessage = openMessage + VbCrlf + "올엣카드 할인 차감 : " & allatsubtractsum
        end if
    end if

response.write "4" & "<br>"

	'배송비도 같이 취소된다. setCancelDetail()

response.write "5" & "<br>"

    if (IsAllCancel) then
	    ''전체 취소인경우
	    '' 주문  master 취소 변경
	    call setCancelMaster(id, orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문취소 완료"
        end if
	else
	    ''부분 취소인경우
	    '' 주문  detail 취소 변경
	    call setCancelDetail(id, orderserial)

		if (refunddeliverypay <> 0) then
			'// 업체 추가배송비 부과
			Call AddBeasongpayForCancel(id, orderserial)
		end if

	    call reCalcuOrderMaster(orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문부분취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문부분취소 완료"
        end if
	end if

    ''마일리지는 주문건 취소 후 재계산해야함.
    '예치금, Gift카드 재계산
    if (userid<>"") then
        Call updateUserMileage(userid)

        if IsUpdatedDeposit then
        	Call updateUserDeposit(userid)
        end if

        if IsUpdatedGiftCard then
        	Call updateUserGiftCard(userid)
        end if
    end if

    ''최근 주문수량 조정 2015/08/12
    if (userid<>"") and (IsAllCancel) then
        sqlStr = "exec [db_order].[dbo].sp_Ten_Recalcu_His_recent_OrderCNT '" & userid & "'"
        dbget.Execute(sqlStr)
    end if

    '' ''전자보증서 발급된 경우 취소
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
		'추가배송비
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
		sqlStr = sqlStr + " , '추가배송비' " & vbCrlf
		sqlStr = sqlStr + " , (case when '" & makerid & "' <> '' then '업체개별' else '' end) " & vbCrlf
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

'// 주문취소 완료시 접수중인 내역의 상품금액 업데이트
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

    ''마일리지 환급
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

        call AddCustomerOpenContents(id, "마일리지 환불 완료: " & CStr(realrefund))
    elseif (returnmethod="R910") then
    	'예치금 전환

    	title = Replace(title, "마일리지", "예치금")
    	title = Replace(title, "무통장", "예치금")

        sqlStr = "insert into [db_user].[dbo].tbl_depositlog"
        sqlStr = sqlStr + " (userid, deposit, jukyocd, jukyo, orderserial, deleteyn)"
        sqlStr = sqlStr + " values('" + userid + "'," + CStr(realrefund) + ",'200','" & title & "','" + orderserial + "','N')"
        dbget.Execute sqlStr

        sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call updateUserDeposit(userid)

        call AddCustomerOpenContents(id, "예치금 환불 완료: " & CStr(realrefund))
    elseif (returnmethod<>"R000") then
        sqlStr = " update [db_cs].[dbo].tbl_as_refund_info"
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call AddCustomerOpenContents(id, "환불(취소) 완료: " & CStr(realrefund))
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

	'// 맞교환회수(A111)에 교환주문이 등록되면 맞교환출고(A100)에도 같이 등록
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
    '' A003 환불요청 , A005 외부몰환불요청 , A007 신용카드/실시간이체취소요청
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

    'R007 무통장환불
    'R020 실시간이체취소
    'R050 입점몰결제 취소
    'R080 올엣카드취소
    'R100 신용카드취소
    'R550 기프팅취소
    'R560 기프티콘취소
    'R120 신용카드부분취소
    'R400 휴대폰취소
	'R420 휴대폰부분취소
    'R900 마일리지로환불
    'R910 예치금환불
    'R022 실시간이체부분취소(NP)
    'R150 이니렌탈취소

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
        title = "주문 취소 후 " + title
    elseif (divcd="A004") then
        title = "반품 처리 후 " + title
    elseif (divcd="A010") then
        title = "회수 처리 후 " + title
    elseif (divcd="A100") then
        title = "교환 출고 후 " + title
    end if

    if (RegDivCd<>"") then
        NewRegedID =  RegCSMaster(RegDivCd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

''        '''기존 취소(반품) 내역 복사
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
''		'관련 CS''''''''''''''''''''*******************************
''        sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
''        sqlStr = sqlStr + " set refasid = " & CStr(id) & " "
''        sqlStr = sqlStr + " where id = " & CStr(NewRegedID) & " "
''        dbget.Execute sqlStr

        Call CopyWebCancelRefundInfo(id, NewRegedID)

        CheckNRegRefund = NewRegedID
    end if
end function

'// ===========================================================================
'// 마이너스 주문 입력을 위해 수량체크
'// ===========================================================================
''원주문
''      ----> 교환주문1
''                     ----> 교환주문2
''  I
''  V
'' 마이너스1     I
''               I              I
''               V              I
''            마이너스2         I
''                              I
''                              V
''                          마이너스3
'// ===========================================================================
'// 원주문 + 교환주문 수량 >= 원주문(and 교환주문) 에 대한 마이너스 주문 수량 + CS접수 수량
'// ===========================================================================
function CheckOverMinusOrderItemnoExist(id, orderserial)
	dim sqlStr

	CheckOverMinusOrderItemnoExist = False

    ''접수되는 내역보다 기존 마이너스+ 추가 마이너스  합계가 큰지 체크 (중복접수)
    if (GC_IsOLDOrder) then
        '' 과거 주문인 경우.. Skip
        Exit function
    else
        '// TODO : 기존 접수된 CS 는 카운트 되지 않고 있다.
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

    ''접수되는 내역보다 기존 마이너스+ 추가 마이너스  합계가 큰지 체크 (중복접수)
    if (GC_IsOLDOrder) then
        '' 과거 주문인 경우.. Skip
        Exit function
    else
        '// TODO : 기존 접수된 CS 는 카운트 되지 않고 있다.
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

    ''초과 교환주문 체크 (중복접수)
    if (GC_IsOLDOrder) then
        '' 과거 주문인 경우.. Skip
        Exit function
    else
        '// TODO : 기존 접수된 CS 는 카운트 되지 않고 있다.
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
'// * 마이너스 주문 등록 이후 체크
'// ===========================================================================
'// 원주문 + 교환주문 금액 >= 원주문(and 교환주문) 에 대한 마이너스 주문 금액
'// ===========================================================================
'// CheckOverMinusOrderItemnoExist 참조
'// ===========================================================================
function CheckOverMinusOrderPriceExist(orderserial, byref ErrStr)

	CheckOverMinusOrderPriceExist = False

end function

function IsOrderExists(orderserial)
	dim sqlStr

	IsOrderExists = True

    ''원주문건 조회
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

    ''원주문건 조회
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

'// 취소 후 잔여상품 금액이 3만원 미만이면 마일리지 차감
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
        ErrStr = "마이너스 주문 상품 합계가 원 상품보타 클 수 있습니다.\n(중복 접수되었을 수 있습니다. 명령이 취소 됩니다.)"

        exit function

	end if

	if CheckCSFinished(id) = True then
        CheckNAddMinusOrder = False
        ErrStr = "이미 완료된 CS내역입니다."
		exit function
	end if

	if (IsOrderExists(orderserial) = False) then

        CheckNAddMinusOrder = False
        ErrStr = "원 주문건이 존재하지 않습니다."

        exit function

	end if

	if (IsCSDetailExists(id, orderserial) = False) then

        CheckNAddMinusOrder = False
        ErrStr = "반품 주문건 상세내역이 없습니다. - 관리자 문의요망"

        exit function

	end if

	'// =======================================================================
	''원주문건( + 교환주문) 금액 조회
	Call GetOrgOrderPriceInfo(orderserial, orgsubtotalprice, orgsumPaymentEtc, orgtencardspend, orgmiletotalprice, orgspendmembership, orgallatdiscountprice, orgdepositsum, orggiftcardsum, orgpercentcouponsum)

	'// =======================================================================
    MinusOrderSerial =  AddMinusOrder(id, orderserial)

    if (MinusOrderSerial="") then

        CheckNAddMinusOrder = false
        ErrStr = "반품 주문건 생성 실패 - 반드시! 관리자 문의요망."

        exit function

    end if

	Call AddminusOrderLink(id, MinusOrderserial)

	'// =======================================================================
	''원주문건( + 교환주문) 에 대한 마이너스 주문 금액 합계
	Call GetMinusOrderPriceInfo(orderserial, totalpreminus_subtotalprice, totalpreminus_sumPaymentEtc, totalpreminus_tencardspend, totalpreminus_miletotalprice, totalpreminus_spendmembership, totalpreminus_allatdiscountprice, totalpreminus_depositsum, totalpreminus_giftcardsum, totalpreminus_percentcouponsum)

	'// =======================================================================
    if (totalpreminus_subtotalprice > orgsubtotalprice) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스주문 결제금액합계가 원주문보다 큽니다.(중복 접수 : 101)"
        exit function
    end if

    if (totalpreminus_sumPaymentEtc > orgsumPaymentEtc) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스주문 보조결제금액합계가 원주문보다 큽니다.(중복 환원 : 102)"
        exit function
    end if

    if (totalpreminus_tencardspend > orgtencardspend) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스주문 쿠폰합계가 원주문보다 큽니다.(중복 환원 : 103)"
        exit function
    end if

    if (totalpreminus_miletotalprice > orgmiletotalprice) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스주문 마일리지합계가 원주문보다 큽니다.(중복 환원 : 104)"
        exit function
    end if

    if (totalpreminus_spendmembership > orgspendmembership) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스주문 맴버쉽카드합계가 원주문보다 큽니다.(중복 환원 : 105)"
        exit function
    end if

    if (totalpreminus_allatdiscountprice > orgallatdiscountprice) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스 기타할인합계가 원주문보다 큽니다.(중복 접수 : 106)"
        exit function
    end if

	'// =======================================================================
    ''원주문건 조회
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
       ''반품 환급 마일리지
       MinusMiletotalprice = rsget("miletotalprice")
    end if
    rsget.Close

    sqlStr = " select IsNULL(realPayedsum,0) as realPayedsum from [db_order].[dbo].tbl_order_PaymentEtc "
    sqlStr = sqlStr + " where orderserial='" + MinusOrderSerial + "' and acctdiv = '200' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	MinusDepositPrice = 0
    if Not rsget.Eof then
       ''반품 환급 예치금
       MinusDepositPrice = rsget("realPayedsum")
    end if
    rsget.Close

    sqlStr = " select IsNULL(realPayedsum,0) as realPayedsum from [db_order].[dbo].tbl_order_PaymentEtc "
    sqlStr = sqlStr + " where orderserial='" + MinusOrderSerial + "' and acctdiv = '900' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	MinusGiftCardPrice = 0
    if Not rsget.Eof then
       ''반품 환급 Gift카드
       MinusGiftCardPrice = rsget("realPayedsum")
    end if
    rsget.Close

    ''마일리지/예치금 재계산
    '마일리지는 구매상품반품으로 항상 재계산이 필요하지만, 예치금,Gift카드는 사용이 있는 경우만 재계산한다.
    if (userid<>"") and (sitename="10x10") then

        ''반품 환급 마일리지 추가
        if (MinusMiletotalprice<>0) then
            sqlStr = "insert into [db_user].[dbo].tbl_mileagelog(userid,mileage,jukyocd,jukyo,orderserial)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusMiletotalprice)) + ",'02','반품환급','" + MinusOrderSerial + "')"

			dbget.Execute  sqlStr
        end if

        Call updateUserMileage(userid)

        ''반품 환급 예치금 추가
        if (MinusDepositPrice<>0) then
            sqlStr = "insert into [db_user].[dbo].tbl_depositlog(userid,deposit,jukyocd,jukyo,orderserial)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusDepositPrice)) + ",'100','반품환급','" + MinusOrderSerial + "')"

			dbget.Execute  sqlStr

			Call updateUserDeposit(userid)
        end if

        ''반품 환급 Gift카드 추가
        if (MinusGiftCardPrice<>0) then
            sqlStr = "insert into [db_user].[dbo].tbl_giftcard_log(userid,useCash,jukyocd,jukyo,orderserial, reguserid)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusGiftCardPrice)) + ",'200','반품환급','" + MinusOrderSerial + "','" + CStr(session("ssbctid")) + "')"

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
        ErrStr = "마이너스 주문 상품 합계가 원 상품보타 클 수 있습니다.\n(중복 접수되었을 수 있습니다. 명령이 취소 됩니다.)"

        exit function

	end if

	if CheckCSFinished_3PL(id) = True then
        CheckNAddMinusOrder_3PL = False
        ErrStr = "이미 완료된 CS내역입니다."
		exit function
	end if

	if (IsOrderExists_3PL(orderserial) = False) then

        CheckNAddMinusOrder_3PL = False
        ErrStr = "원 주문건이 존재하지 않습니다."

        exit function

	end if

	if (IsCSDetailExists_3PL(id, orderserial) = False) then

        CheckNAddMinusOrder_3PL = False
        ErrStr = "반품 주문건 상세내역이 없습니다. - 관리자 문의요망"

        exit function

	end if

	'// =======================================================================
    MinusOrderSerial =  AddMinusOrder_3PL(id, orderserial)

    if (MinusOrderSerial="") then

        CheckNAddMinusOrder_3PL = false
        ErrStr = "반품 주문건 생성 실패 - 반드시! 관리자 문의요망."

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

    ''쿠폰 마일리지 환급 계산
    sqlStr = " select r.*, IsNull(r.refundgiftcardsum, 0) as refundgiftcardsum, IsNull(r.refunddepositsum, 0) as refunddepositsum from [db_cs].[dbo].tbl_new_as_list a"
    sqlStr = sqlStr + " , [db_cs].[dbo].tbl_as_refund_info r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"

    '환불없음을 선택해도 마일리지 예치금 등은 환급한다.
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

    ''환불 상세 내역이 없을 수 있음
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
                 rsget.MoveNext  '''이부분 빠져있어 timeOUT
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
	sqlStr = sqlStr + " ,comment='원주문번호:" + orderserial +"'" & vbCrlf
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
	sqlStr = sqlStr + " ,userlevel=O.userlevel" & vbCrlf    ''20121219 추가
	sqlStr = sqlStr + " ,pggubun=O.pggubun" & vbCrlf    	''2015-08-25 추가
	if (CStr(refundcouponsum)<>"") then
        sqlStr = sqlStr + " ,bCpnIdx=O.bCpnIdx" & vbCrlf    ''20121129 추가
    end if
	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from (select top 1 * from [db_log].[dbo].tbl_old_order_master_2003 where orderserial='" + orderserial + "') O" & vbCrlf
	else
	    sqlStr = sqlStr + " from (select top 1 * from [db_order].[dbo].tbl_order_master where orderserial='" + orderserial + "') O" & vbCrlf
	end if
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.idx=" + CStr(iid)

	dbget.Execute sqlStr

	''원배송비 환급 있을경우
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

	''취소/반품 상품 상세내역
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
		'반품배송비
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
		sqlStr = sqlStr + " , '반품배송비' " & vbCrlf
		sqlStr = sqlStr + " , (case when IsNull(a.requireupche, 'Y') = 'Y' then '업체개별' else '' end) " & vbCrlf
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

    ''주문금액 재계산
    call recalcuOrderMaster(neworderserial)

    ''재고수량조정 - 한정수량은 조정 안됨
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_minusOrder '" & neworderserial & "'"
    dbget.Execute sqlStr

    AddMinusOrder    = neworderserial
end function

function GetBankName(acctno)
	select case acctno
		case "11"
			GetBankName = "농협"
		case "06"
			GetBankName = "국민"
		case "20"
			GetBankName = "우리"
		case "26"
			GetBankName = "신한"
		case "81"
			GetBankName = "하나"
		case "03"
			GetBankName = "기업"
		case "39"
			GetBankName = "경남"
		case "32"
			GetBankName = "부산"
		case "71"
			GetBankName = "우체국"
		case "07"
			GetBankName = "수협"
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
	sqlStr = sqlStr + " ,accountname='" & accountname & "'" & vbCrlf			'// 입금자명
	sqlStr = sqlStr + " ,accountno=''" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdiv='0'" & vbCrlf
	sqlStr = sqlStr + " ,accountdiv=" & accountdiv & vbCrlf
	sqlStr = sqlStr + " ,regdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,beadaldiv=1" & vbCrlf									'// 고객추가결제건 자사몰주문으로 변경
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
	sqlStr = sqlStr + " ,comment='원주문번호:" + orderserial +"'" & vbCrlf
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
	sqlStr = sqlStr + " ,userlevel=O.userlevel" & vbCrlf    ''20121219 추가
	sqlStr = sqlStr + " ,pggubun=''" & vbCrlf    	''2015-08-25 추가
	if (payordertype = "A") then
		sqlStr = sqlStr + " ,baljudate=getdate()" & vbCrlf			'// 기출고이므로 발주안되도록
	end if
	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from (select top 1 * from [db_log].[dbo].tbl_old_order_master_2003 where orderserial='" + orderserial + "') O" & vbCrlf
	else
	    sqlStr = sqlStr + " from (select top 1 * from [db_order].[dbo].tbl_order_master where orderserial='" + orderserial + "') O" & vbCrlf
	end if
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.idx=" + CStr(iid)

	dbget.Execute sqlStr

	'// 추가배송비
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
	sqlStr = sqlStr + " , '고객추가결제' " & vbCrlf
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
		'//기출고결제
        '// A : 기출고결제, N : 주문접수
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
		''sqlStr = sqlStr + " ,'1'" & vbCrlf									'// 기출고 : 0 은 업체통보로 넘어가므로 '1' 로 설정
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

    ''주문금액 재계산
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
	sqlStr = sqlStr + " ,comment='원주문번호:" + orderserial +"'" & vbCrlf
	sqlStr = sqlStr + " ,linkorderserial=O.orderserial" & vbCrlf
	sqlStr = sqlStr + " ,sitename=O.sitename" & vbCrlf
	sqlStr = sqlStr + " ,totalsum=O.totalsum" & vbCrlf
	sqlStr = sqlStr + " ,subtotalprice=O.subtotalprice" & vbCrlf
	sqlStr = sqlStr + " ,reqzipaddr=O.reqzipaddr" & vbCrlf
	sqlStr = sqlStr + " from (select top 1 * from [db_threepl].[dbo].[tbl_tpl_orderMaster] where orderserial='" + orderserial + "') O" & vbCrlf
	sqlStr = sqlStr + " where [db_threepl].[dbo].[tbl_tpl_orderMaster].idx=" + CStr(iid)

	dbget_TPL.Execute sqlStr

	''취소/반품 상품 상세내역
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

    ''주문금액 재계산
    call recalcuOrderMaster_3PL(neworderserial)

    ''재고수량조정 - 한정수량은 조정 안됨
	'// !!! 트랜젝션 밖으로 뺀다.
    ''sqlStr = " exec [db_summary].[dbo].[usp_TPL_RealtimeStock_minusOrder] '" & neworderserial & "'"
    ''dbget.Execute sqlStr

    AddMinusOrder_3PL    = neworderserial
end function

'// 상품변경 맞교환시 고객 추가배송비
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
        ErrStr = "잘못된 접근입니다. : CS건 없음"
    end if

    if currstate = "B007" then
        ErrStr = "이미 완료된 CS내역입니다."
    end if

    if deleteyn = "Y" then
        ErrStr = "삭제된 CS내역입니다."
    end if

    if (payordertype <> "A") then
        '// 기출고 결제건 아님
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
    sqlStr = sqlStr + " 		and m.ipkumdiv <= '8' "			'// 이미 출고완료된 경우도 정상처리
    sqlStr = sqlStr + " 		and d.currstate >= '1' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
        '//
    else
        ErrStr = "기출고주문이 없습니다. (결제완료이전 또는 취소상태입니다.)"
	end if
	rsget.Close

    if ErrStr <> "" then
        exit function
    end if

	sqlStr = " update [db_order].[dbo].tbl_order_detail "
	sqlStr = sqlStr + " set currstate = '7', beasongdate = getdate(), upcheconfirmdate = getdate(), songjangdiv = '99', songjangno = '기출고주문' "
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
'// 상품변경 맞교환 교환주문
'// ===========================================================================
function CheckAndAddChangeOrder(asid, orderserial, byref ErrStr)

	'// =======================================================================
	if (CheckOverChangeOrderItemnoExist(id, orderserial) = True) then

        CheckAndAddChangeOrder = False
        ErrStr = "원주문( + 교환주문 + 마이너스주문) 을 초과하여 회수되는 상품이 있습니다.\n(중복 접수되었을 수 있습니다. 시스템팀 문의)"

        exit function

	end if

	CheckAndAddChangeOrder = AddChangeOrder(asid, orderserial)
end function

'// ===========================================================================
'// 상품변경 맞교환 교환주문(주문접수 상태로 등록)
'// ===========================================================================
function CheckAndAddChangeOrderJupsu(asid, orderserial, byref ErrStr)

	'// =======================================================================
	if (CheckOverChangeOrderItemnoExist(id, orderserial) = True) then

        CheckAndAddChangeOrderJupsu = False
        ErrStr = "원주문( + 교환주문 + 마이너스주문) 을 초과하여 회수되는 상품이 있습니다.\n(중복 접수되었을 수 있습니다. 시스템팀 문의)"

        exit function

	end if

	CheckAndAddChangeOrderJupsu = AddChangeOrderJupsu(asid, orderserial)
end function

function DelChangeOrder(asid)

	dim sqlStr

	DelChangeOrder = ""

end function

'// ===========================================================================
'// 상품변경 맞교환 교환주문 등록(주문접수 상태로 등록)
'// ===========================================================================
''가격정보
''옵션변경 : 현재 옵션가 동일해야만 맞교환 가능, 가격정보 그대로 카피
''상품변경 :
''           판매가 동일하고, 수량도 동일한 경우 : 판브랜드,현재 판매가(할인가),매입가, 쿠폰적용가능 등 모두 동일해야하고 1:1 변경만 가능(수량은 여러개 가능), 가격정보 그대로 카피
''           다른 경우, 판매가(할인가),매입가는 CS디테일정보 이용, 쿠폰가 있으면(할인가-쿠폰가 가 다른 경우) 쿠폰정보 복사
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
	rsget("jumundiv") = "6"					'// 교환주문
	rsget("userid") = ""
	rsget("accountname") = ""
	rsget("accountdiv") = "7"				'// 일단은 무통장으로
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
	sqlStr = sqlStr + " ,accountdiv=O.accountdiv" & vbCrlf			'// 결제정보 복사
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
	sqlStr = sqlStr + " ,comment='원주문번호:" + orderserial +"'" & vbCrlf
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

	'상세내역(맞교환회수 상품정보)
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

	'// 상품변경 교환출고 상품(A100)
	Call AddOrderDetail(refasid, orderserial, neworderserial, iid, "0")

    ''주문금액 재계산
    call recalcuOrderMaster(neworderserial)

	'' 재고보정 --> 출고완료시 한다.
	'' 한정수량 --> CS접수시 이미 반영했다.

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
	'// 상품변경 교환출고 상품(A100)
	Dim sqlStr

	''가격정보
	''옵션변경 : 현재 옵션가 동일해야만 맞교환 가능, 가격정보 그대로 카피
	''상품변경 :
	''           판매가 동일하고, 수량도 동일한 경우 : 판브랜드,현재 판매가(할인가),매입가, 쿠폰적용가능 등 모두 동일해야하고 1:1 변경만 가능(수량은 여러개 가능), 가격정보 그대로 카피
	''           다른 경우, 판매가(할인가),매입가는 CS디테일정보 이용, 쿠폰가 있으면(할인가-쿠폰가 가 다른 경우) 쿠폰정보 복사

	'상세내역(맞교환출고 상품정보) : 판매가 동일하고, 수량도 동일한 경우
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
	sqlStr = sqlStr + " and J.SalePrice is NULL "			'// 수량 동일하고, 판매가 동일한 경우
    dbget.Execute sqlStr

	'상세내역(맞교환출고 상품정보) : 판매가 다른 경우 or 수량 다른 경우
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
	sqlStr = sqlStr + " and J.SalePrice is not NULL "			'// 판매가 다른 경우 or 수량 다른 경우
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

	'// 출고완료시 재고 보정
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_changeOrder '" & changeorderserial & "'"
    dbget.Execute sqlStr

end function

function AddChangeOrder(id, orderserial)
    dim changeorderserial

    changeorderserial = AddChangeOrderJupsu(id, orderserial)

    Call FinishChangeOrder(changeorderserial)

    AddChangeOrder = changeorderserial
end function

'// 교환주문 찾기
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
    	errMsg = "잘못된 구분입니다."
    end if
    rsget.Close

end function

'// 교환주문 브랜드정보
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

    ''무통장이나 마일리지 환불이나 예치금전환인 경우만 수기 수정 가능
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

    ''무통장이나 마일리지 환불이나 예치금전환인 경우만 수기 수정 가능
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
        ''한정수량 조정 -
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

        ''옵션있는상품
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

'// CS 맞교환출고(동일상품, 상품변경 - A000, A100) 접수시 출고되는 상품 한정차감
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


    ''한정수량 조정 -
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
    sqlStr = sqlStr + " 	and d.regitemno > 0 " + vbCrlf			'// 상품변경 맞교환출고상품 +, 회수상품 -
    sqlStr = sqlStr + " ) as T" + vbCrlf
    sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.Itemid"
    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.limityn='Y'"
    dbget.Execute(sqlStr)

	''옵션있는상품
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
    sqlStr = sqlStr + " 	and d.regitemno > 0 " + vbCrlf			'// 상품변경 맞교환출고상품 +, 회수상품 -
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
    '// 전자보증서가 있으면 보증서 취소 요청 (2006.06.15; 운영관리팀 허진원)
    dim objUsafe, result, result_code, result_msg
    On Error Resume Next
    	Set objUsafe = CreateObject( "USafeCom.guarantee.1"  )

    '	Test일 때
    '	objUsafe.Port = 80
    '	objUsafe.Url = "gateway2.usafe.co.kr"
    '	objUsafe.CallForm = "/esafe/guartrn.asp"

        ' Real일 때
        objUsafe.Port = 80
        objUsafe.Url = "gateway.usafe.co.kr"
        objUsafe.CallForm = "/esafe/guartrn.asp"

    	objUsafe.gubun	= "B0"				'// 전문구분 (A0:신규발급, B0:보증서취소, C0:입금확인)
    	objUsafe.EncKey	= ""			'널값인 경우 암호화 안됨
    	objUsafe.mallId	= "ZZcube1010"		'// 쇼핑몰ID
    	objUsafe.oId	= CStr(orderserial)	'// 주문번호

    	'처리 실행!
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

''검증. 전체 취소 맞는지.
''전체취소인지만 확인하고 다른 검증 안한다.
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

'쓰이는 곳이 있을수 있으므로 남겨둔다.
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

	'// 이전 처리자 아이디 저장
	sqlStr = " exec [db_log].[dbo].[usp_Ten_SaveCSHistory] " + CStr(asid) + " "
	dbget.Execute(sqlStr)

end function

' 주문 상품쿠폰 적용수 체크		'2023.10.19 한용민 생성
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
    sqlStr = sqlStr & " 		and c.itemcouponexpiredate>getdate()"	' 유효기간체크

    if couponGubun<>"" then
        sqlStr = sqlStr & " and c.couponGubun='"& couponGubun &"'"
    end if
    if userid<>"" then
        sqlStr = sqlStr & " and c.userid='"& userid &"'"
    end if
    
    sqlStr = sqlStr & " 	where ad.masterid='" & asid & "'"
    sqlStr = sqlStr & " ) as t"
    sqlStr = sqlStr & " where t.rk=1"
    sqlStr = sqlStr & " and t.prevCopiedItemCouponCount=0"    ' 재발급체크

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

' 상품쿠폰 복사발급     ' 2023.10.19 한용민 생성
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
    sqlStr = sqlStr & " and a.divcd in ('A008', 'A004', 'A010')"    ' A008 주문취소 / A004 반품접수(업체배송) / A010 회수신청(텐바이텐배송)

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
    sqlStr = sqlStr & " 	    	and c.itemcouponexpiredate>getdate()"	' 유효기간체크

    if couponGubun<>"" then
        sqlStr = sqlStr & "     and c.couponGubun='"& couponGubun &"'"
    end if
    if userid<>"" then
        sqlStr = sqlStr & "     and c.userid='"& userid &"'"
    end if
    
    sqlStr = sqlStr & " 	    where ad.masterid='" & asid & "'"
    sqlStr = sqlStr & "     ) as t"
    sqlStr = sqlStr & "     where t.rk=1"
    sqlStr = sqlStr & "     and t.prevCopiedItemCouponCount=0"    ' 재발급체크

    'response.write sqlStr & "<br>"
	dbget.Execute sqlStr, excuteRowCount

    if excuteRowCount>0 then
	    CheckAndCopyItemCoupon = True
    else
        ' 상품쿠폰 재발급 접수 와 완료처리 사이에 쿠폰 유효기간이 지난경우 재발행여부 N으로 바꾼다.
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

' 상품쿠폰 복사여부     ' 2023.10.19 한용민 생성
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

' 정산내역 체크 ' 이상구 생성
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

	'// 추가배송비 있는 경우, 마지막 취소건부터 복구해야 함.
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

' 물류 출고지시서 삭제       ' 2021.03.31 한용민 생성
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
		'// 텐배 주문 아님
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
		'// 주문 전체 취소 후 주문복구 아님
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
	sqlStr = sqlStr + " and divcd in ('A004', 'A010', 'A008')"		'// 반품, 회수, 취소
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
		'// finishdate 체크 하는것 뺌(2014-05-30 skyer9 : 물류처리완료)
        IsCsErrStockUpdateRequire = ((rsget("divcd")="A000") or (rsget("divcd")="A011") or (rsget("divcd")="A100") or (rsget("divcd")="A111")) and (rsget("currstate")="B007") and (rsget("deleteyn")="N")
    end if
    rsget.close

    sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
    sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " , deletedate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)
    sqlStr = sqlStr + " and currstate='B007'"
	''sqlStr = sqlStr + " and divcd in ('A004', 'A010', 'A008')"		'// 반품, 회수, 취소
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

    ''재고수량조정 - 한정수량은 조정 안됨
    sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_minusOrder '" & minusorderserial & "'"
    dbget.Execute sqlStr

end function

' 상품변경 맞교환회수 삭제 처리     ' 2019.10.18 한용민 생성
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

    ''재고수량조정 - 한정수량은 조정 안됨
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
    sqlStr = sqlStr + " 	and acctdiv in ('200', '900') " + VbCrlf			'200 : 예치금, 900 : Gift카드
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
			'' 전체 취소인경우 주문건 취소로 jukyocd : 2 상품구매, 3 : 부분취소시 환원마일리지
			sqlStr = " update [db_user].[dbo].tbl_mileagelog " + VbCrlf
			sqlStr = sqlStr + " set deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
			sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and deleteyn <> 'N'" + VbCrlf
			sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf
			dbget.Execute sqlStr
		end if

		if (userid <> "") and (depositsum <> 0) then
			'' 전체 취소인경우 주문건 취소로 jukyocd : 100 상품구매, 10 : 부분취소시 예치금 환원
			sqlStr = " update [db_user].[dbo].tbl_depositlog " + VbCrlf
			sqlStr = sqlStr + " set deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
			sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and deleteyn <> 'N' " + VbCrlf
			sqlStr = sqlStr + " and jukyocd in ('100','10')" + VbCrlf					'100 : 상품구매사용 / 10 : 일부환원 (참고 : db_user.dbo.tbl_deposit_gubun)
			dbget.Execute sqlStr

			IsUpdatedDeposit = True
		end if

		if (userid <> "") and (giftcardsum <> 0) then
			'' 전체 취소인경우 주문건 취소로 jukyocd : 200 상품구매, 300 : 부분취소시 Gift카드 환원
			sqlStr = " update [db_user].[dbo].tbl_giftcard_log " + VbCrlf
			sqlStr = sqlStr + " set deleteyn='N' " + VbCrlf
			sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
			sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
			sqlStr = sqlStr + " and deleteyn <> 'N' " + VbCrlf
			sqlStr = sqlStr + " and jukyocd in ('200','300')" + VbCrlf					'200 : 상품구매사용 / 300 : 일부환원 (참고 : db_user.dbo.tbl_giftcard_gubun)
			dbget.Execute sqlStr

			IsUpdatedGiftCard = True
		end if

		if (couponsum <> 0) then
			sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
			sqlStr = sqlStr + " set isusing='Y' "                   + VbCrlf
			sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
			dbget.Execute sqlStr
		end if

		'// 취소 주문 정상화
		Call setRestoreCancelMaster(asid, orderserial)
	else
		'// ====================================================================
		if (userid <> "") and (refundmileagesum <> 0) then
			'' 부분 취소인데 마일리지 환원할 경우.
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
			sqlStr = sqlStr + " ,'상품구매 취소 철회' "
			sqlStr = sqlStr + " ,'" + orderserial + "'"
			sqlStr = sqlStr + " ,'N'"
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr
		end if

		if (userid <> "") and (refunddepositsum <> 0) then
			'' 부분 취소인데 예치금 환원할 경우.
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
			sqlStr = sqlStr + " ,'상품구매 취소 철회'"
			sqlStr = sqlStr + " ,'" + orderserial + "'"
			sqlStr = sqlStr + " ,'N'"
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr

			IsUpdatedDeposit = True
		end if

		if (userid <> "") and (refundgiftcardsum <> 0) then
			'' 부분 취소인데 Gift카드 환원할 경우.
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
			sqlStr = sqlStr + " ,'상품구매 취소 철회'"
			sqlStr = sqlStr + " ,'" + orderserial + "'"
			sqlStr = sqlStr + " ,'N'"
			sqlStr = sqlStr + " ,'" + CStr(session("ssbctid")) + "'"
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr

			IsUpdatedGiftCard = True
		end if

		if (refundcouponsum <> 0) then
			'' 부분 취소인경우 - 환급한 만큼 깜..
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
			'' 부분 취소인경우 - 환급한 만큼 깜..
			sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
			sqlStr = sqlStr + " set allatdiscountprice=allatdiscountprice + " + CStr(allatsubtractsum*-1) + VbCrlf
			sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf
			''response.write sqlStr
			dbget.Execute sqlStr
		end if

		'// 취소 상품 정상화
		Call setRestoreCancelDetail(asid, orderserial)

		'// 추가배송비 있으면 취소
		Call CancelAddBeasongpayForCancel(asid)

		Call reCalcuOrderMaster(orderserial)

	end if

    ''마일리지는 주문건 취소 후 재계산해야함.
    '예치금, Gift카드 재계산
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
		response.write "ERROR : 6개월 이전 주문 처리불가"
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

		''response.write "ERROR : 출고완료 주문"
		''RestoreCancelValid = False
		''exit function
	end if

end function

'수량이 같으면 취소 Flag 다르면 수량변경
'배송비도 취소
function setRestoreCancelDetail(Asid, orderserial)
    dim sqlStr

    '수량변경 - 상품일부취소인경우
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
	sqlStr = sqlStr + " 			and d.itemid <> 0 "				'// 배송비는 수량이 다를 수 없다.(언제나 1개)
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

'//텐바이텐 배송 상품 cs 한정 처리		'/2016.07.18 한용민 생성
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

	'/상품 한정 처리
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
	sqlStr = sqlStr & " 	and nd.itemid not in (0,100)" & vbcrlf	' 배송비, 포장비 제외

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " 	group by nd.itemid " & vbCrLf
	sqlStr = sqlStr & " ) T" & vbCrLf
	sqlStr = sqlStr & " 	on i.itemid = T.itemid " & vbCrLf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/옵션 있는 상품 한정 처리
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

	sqlStr = sqlStr & " and nd.itemid not in (0,100)" & vbcrlf	' 배송비, 포장비 제외
	sqlStr = sqlStr & " and nd.itemoption<>'0000'" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/상품 품절처리 판매중인 상품중 한정품절 된 상품 일시 품절로 변경
	sqlStr = "update i" & vbcrlf
	sqlStr = sqlStr & " set i.sellyn='S' , i.lastupdate=getdate()" & vbcrlf
	sqlStr = sqlStr & " from (" & vbcrlf
	sqlStr = sqlStr & " 	select nd.itemid" & vbcrlf
	sqlStr = sqlStr & " 	from [db_cs].[dbo].[tbl_new_as_detail] nd" & vbcrlf
	sqlStr = sqlStr & " 	where nd.masterid="& Asid &"" & vbcrlf

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " 	and nd.itemid not in (0,100)" & vbcrlf	' 배송비, 포장비 제외
	sqlStr = sqlStr & " 	group by nd.itemid" & vbcrlf
	sqlStr = sqlStr & " ) as t" & vbcrlf
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i" & vbcrlf
	sqlStr = sqlStr & " 	on t.itemid = i.itemid" & vbcrlf
	sqlStr = sqlStr & " 	and i.sellyn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and i.limityn='Y'" & vbcrlf
	sqlStr = sqlStr & " 	and (i.limitno-i.limitSold<1)" & vbcrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	'/일시 품절이나 한정수량>0 경우 판매로 변경
	sqlStr = "update i" & vbcrlf
	sqlStr = sqlStr & " set i.sellyn='Y' , i.lastupdate=getdate()" & vbcrlf
	sqlStr = sqlStr & " from (" & vbcrlf
	sqlStr = sqlStr & " 	select nd.itemid" & vbcrlf
	sqlStr = sqlStr & " 	from [db_cs].[dbo].[tbl_new_as_detail] nd" & vbcrlf
	sqlStr = sqlStr & " 	where nd.masterid="& Asid &"" & vbcrlf

	if orderserial<>"" then sqlStr = sqlStr & " and nd.orderserial='"& orderserial &"'"

	sqlStr = sqlStr & " 	and nd.itemid not in (0,100)" & vbcrlf	' 배송비, 포장비 제외
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
