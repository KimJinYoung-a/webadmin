<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%

dim mode, targetdetailidx, targetregitemno

dim orderserial

dim gubun01, gubun02
dim divcd, title, contents_jupsu, contents_finish
dim regUserID, finishuser

dim add_customeraddbeasongpay
dim add_customeraddmethod

dim add_itemid, add_itemoption

dim detailitemlist, newdetailitemlist, contents_itemlist
dim orderdetailidx

dim jungsanExists
dim fromDetailState, toDetailState
dim refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum
dim add_SalePrice, add_ItemCouponPrice, add_BonusCouponPrice, add_buycash
dim iscouponapplied, itemcouponidxapplied, bonuscouponidxapplied

dim itemname, itemoptionname

'품절취소 상품정보 저장
dim modifyitemstockoutyn
dim ResultCount

dim result
dim strSql, iAsID, newasid
dim i, j


'// ===========================================================================
Function GetItemName(itemid)
	dim sqlStr

	GetItemName = ""

	sqlStr = " select "
	sqlStr = sqlStr + " i.itemname "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " WHERE 1=1 "
	sqlStr = sqlStr + " and i.itemid=" & itemid & ""
	'response.write sqlStr

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		GetItemName = rsget("itemname")
	End If
	rsget.close()
end Function

Function GetItemOptionName(itemid, itemoption)
	dim sqlStr

	GetItemOptionName = ""

	sqlStr = " select "
	sqlStr = sqlStr + " v.optionname "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
	sqlStr = sqlStr + " on i.itemid=v.itemid "
	sqlStr = sqlStr + " WHERE 1=1 "
	sqlStr = sqlStr + " and i.itemid=" & itemid & ""
	sqlStr = sqlStr + " and v.itemoption = '" & itemoption & "' "
	'response.write sqlStr

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		GetItemOptionName = rsget("optionname")
	End If
	rsget.close()
end Function


'// ===========================================================================
mode       				= request("mode")
targetdetailidx       	= request("targetdetailidx")
targetregitemno       	= request("targetregitemno")

orderserial       		= request("orderserial")

gubun01       			= request("gubun01")
gubun02       			= request("gubun02")
divcd       			= request("divcd")
title       			= request("title")
contents_jupsu       	= request("contents_jupsu")

''regUserID				= session("ssBctID")
''finishuser				= session("ssBctID")

add_customeraddbeasongpay	= request("add_customeraddbeasongpay")
add_customeraddmethod       = request("add_customeraddmethod")

add_itemid       		= request("add_itemid")
add_itemoption       	= request("add_itemoption")

refunditemcostsum		= request("refunditemcostsum")
refundcouponsum			= request("refundcouponsum")
allatsubtractsum		= request("refundallatsubtractsum")
refundmileagesum		= 0
refunddepositsum		= 0
refundgiftcardsum		= 0
canceltotal				= request("canceltotal")
refundrequire			= request("refundrequire")

add_SalePrice			= request("add_SalePrice")
add_ItemCouponPrice		= request("add_ItemCouponPrice")
add_BonusCouponPrice	= request("add_BonusCouponPrice")
add_buycash				= request("add_buycash")

iscouponapplied			= request("iscouponapplied")
itemcouponidxapplied	= request("itemcouponidxapplied")
bonuscouponidxapplied	= request("bonuscouponidxapplied")

modifyitemstockoutyn	= request("modifyitemstockoutyn")


if (gubun01 = "") then
	gubun01		= "C004"
	gubun02		= "CD99"
end if


'==============================================================================
dim oorderdetail

set oorderdetail = new COrderMaster
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

'' 과거 6개월 이전 내역 검색
if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oorderdetail.FRectOldOrder = "on"
    oorderdetail.QuickSearchOrderDetail
end if

dim fromItemId, fromItemOption, fromItemName, fromItemOptionName
dim isupchebeasong, requiremakerid

for i = 0 to oorderdetail.FResultCount - 1
	if (CStr(targetdetailidx) = CStr(oorderdetail.FItemList(i).Fidx)) then

		fromItemId 			= oorderdetail.FItemList(i).Fitemid
		fromItemOption 		= oorderdetail.FItemList(i).Fitemoption

		fromItemName 		= oorderdetail.FItemList(i).Fitemname
		fromItemOptionName 	= oorderdetail.FItemList(i).Fitemoptionname

		isupchebeasong = oorderdetail.FItemList(i).Fisupchebeasong
		if (isupchebeasong = "Y") then
			requiremakerid = oorderdetail.FItemList(i).Fmakerid
		end if

	end if
next



if (mode="regmodifyorder") then
	'// ===========================================================================
	'// 주문변경(상품변경)

	'==============================================================================
	jungsanExists = false
	'// 취소되는 상품
	strSql = "select top 1 * from db_order.dbo.tbl_order_detail od"
	strSql = strSql & " Join db_jungsan.dbo.tbl_designer_jungsan_detail jd"
	strSql = strSql & " on od.idx=jd.detailidx"
	strSql = strSql & " where od.orderserial='" & orderserial & "' and od.idx = " & targetdetailidx & " "

	rsget.Open strSql,dbget,1
	if Not rsget.Eof then
	    jungsanExists = true
	end if
	rsget.Close

	if (jungsanExists) then
	    response.write "<script language='javascript'>alert('에러 : " & "취소(회수) 되는 상품에 정산 내역이 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	'// 추가되는 상품
	strSql = "select top 1 * from db_order.dbo.tbl_order_detail od"
	strSql = strSql & " Join db_jungsan.dbo.tbl_designer_jungsan_detail jd"
	strSql = strSql & " on od.idx=jd.detailidx"
	strSql = strSql & " where od.orderserial='" & orderserial & "' and od.itemid = " & add_itemid & " and od.itemoption = '" & add_itemoption & "' "

	rsget.Open strSql,dbget,1
	if Not rsget.Eof then
	    jungsanExists = true
	end if
	rsget.Close

	if (jungsanExists) then
	    response.write "<script language='javascript'>alert('에러 : " & "추가되는 상품코드에 정산 내역이 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	'==============================================================================
	fromDetailState = ""
	toDetailState = ""

	'// 추가되는 상품이 이미 디테일에 있는 경우 상태체크(출고완료 상품과 이전 상품을 합칠 수 없다.)
	strSql = "select top 1 IsNull(currstate, '2') as currstate from db_order.dbo.tbl_order_detail od"
	strSql = strSql & " where od.orderserial='" & orderserial & "' and od.itemid = " & add_itemid & " and od.itemoption = '" & add_itemoption & "' "

	rsget.Open strSql,dbget,1
	if Not rsget.Eof then
	    toDetailState = rsget("currstate")
	end if
	rsget.Close

	if toDetailState <> "" then

		strSql = "select top 1 IsNull(currstate, '2') as currstate from db_order.dbo.tbl_order_detail od"
		strSql = strSql & " where od.orderserial='" & orderserial & "' and od.idx = " & targetdetailidx & " "

		rsget.Open strSql,dbget,1
		if Not rsget.Eof then
		    fromDetailState = rsget("currstate")
		end if
		rsget.Close

		if ((CStr(fromDetailState) = "7") and (CStr(toDetailState) <> "7")) or ((CStr(fromDetailState) <> "7") and (CStr(toDetailState) = "7")) then
		    response.write "<script language='javascript'>alert('에러 : " & "출고완료 상품과 이전 상품을 합칠 수 없습니다. 변경할 수 없습니다." & "');history.back();</script>"
		    dbget.close()	:	response.End
		end if

	end if


	'==========================================================================
	contents_finish = "상품변경이 정상적으로 처리되었습니다."

	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	'// 추가될 상품(한가지만 가능)
	contents_itemlist	= contents_itemlist & GetItemName(add_itemid) & vbCrLf & "[" & add_itemoption & "] " & GetItemOptionName(add_itemid, add_itemoption) & " " & targetregitemno & "개 추가" & vbCrLf & vbCrLf

	'// 취소될 상품(한가지만 가능)
	contents_itemlist	= contents_itemlist & fromItemName & vbCrLf & "[" & fromItemOption & "] " & fromItemOptionName & " " & CStr(targetregitemno) & "개 취소" & vbCrLf

	'// 접수내용에 추가
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu


	'==========================================================================
	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'환불정보
	newAsId = 0
	newAsId = RegCSMasterRefundInfoBeforeCancel(iAsID, orderserial, reguserid, refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, contents_finish, refundmileagesum, refunddepositsum, refundgiftcardsum)

	result = CSOrderChangeItemForce(orderserial, fromItemId, add_itemid, fromItemOption, add_itemoption, targetregitemno)

	'금액 이외 정보
	Call CSOrderCopyItemInfoPart(orderserial, fromItemId, add_itemid, fromItemOption, add_itemoption)

	'금액정보
	Call CSOrderSetItemPriceInfo(orderserial, add_itemid, add_itemoption, add_SalePrice, add_ItemCouponPrice, add_BonusCouponPrice, add_buycash)

	if (iscouponapplied = "Y") then
		if (itemcouponidxapplied <> "") then
			'상품쿠폰
			Call CSOrderSetItemCouponInfo(orderserial, add_itemid, add_itemoption, itemcouponidxapplied)
		else
			'보너스쿠폰
			Call CSOrderCopyBonusCouponInfo(orderserial, fromItemId, add_itemid, fromItemOption, add_itemoption)
		end if
	end if


	Call EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	CSOrderRecalculateOrder orderserial,false

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, add_itemid, add_itemoption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
	itemname = CS_ORDER_ITEM_ITEMNAME
	itemoptionname = CS_ORDER_ITEM_OPTIONNAME

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(targetregitemno) & Chr(9)

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, fromItemId, fromItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
	itemname = CS_ORDER_ITEM_ITEMNAME
	itemoptionname = CS_ORDER_ITEM_OPTIONNAME

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*targetregitemno) & Chr(9)

	'==========================================================================
	' CS 마스타 AS 수정
	Call EditCSMaster(iAsID, reguserid, title, html2db(contents_jupsu), gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	'// 업배상품 품절정보 저장
	if (modifyitemstockoutyn = "Y") then
        ResultCount   = SetStockOutByCsAs(iAsid)
	end if

	response.write "<script>" & vbCrLf
	response.write "	alert('수정 되었습니다.');" & vbCrLf
	response.write "	opener.parent.location.reload();" & vbCrLf

	if (newAsId <> 0) then
		response.write "	var a = window.open('/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + CStr(newAsId) + "&mode=editreginfo','pop_cs_action_reg_','width=1200 height=600 scrollbars=yes resizable=yes');" & vbCrLf
	else
		response.write "	var a = window.open('/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + CStr(iAsID) + "&mode=editreginfo','pop_cs_action_reg_','width=1200 height=600 scrollbars=yes resizable=yes');" & vbCrLf
	end if

	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	response.end


elseif (mode="regchangeorder") then

	regUserID	= session("ssBctID")


	'// 출고될 상품(한가지만 가능)
    newdetailitemlist = newdetailitemlist & "|" & targetdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & targetregitemno & Chr(9) & Trim(add_itemid) & Chr(9) & add_itemoption & Chr(9)
	contents_itemlist	= contents_itemlist & GetItemName(add_itemid) & vbCrLf & "[" & add_itemoption & "] " & GetItemOptionName(add_itemid, add_itemoption) & " " & targetregitemno & "개 출고" & vbCrLf & vbCrLf


	'// 회수될 상품(한가지만 가능)
    detailitemlist = detailitemlist & "|" & targetdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(targetregitemno) & Chr(9)
	contents_itemlist	= contents_itemlist & fromItemName & vbCrLf & "[" & fromItemOption & "] " & fromItemOptionName & " " & CStr(targetregitemno) & "개 회수" & vbCrLf


	'// 접수내용에 추가
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu


	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'//  CS Detail(관련상품목록) 등록
	'// 맞교환출고에는 출고되는 상품만 등록한다.
	Call AddCSDetailWithoutOrderDetailByArrStr(newdetailitemlist, iAsID, orderserial)

	'// CS 맞교환출고(동일상품, 상품변경 - A000, A100) 접수시 출고되는 상품 한정차감
	Call ApplyLimitItemByCS(iAsID)

    if (isupchebeasong = "Y") then
        '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        call RegCSMasterAddUpche(iAsID, requiremakerid)

		'// 고객 추가배송비(교환출고에 등록)
		Call SetCustomerAddBeasongPay(iAsID, add_customeraddmethod, add_customeraddbeasongpay, "N", 0)

        '업체배송인 경우 상품변경 맞교환 회수 접수
        newasid = RegCSMaster("A112", orderserial, reguserid, "교환회수(상품변경,업체배송)", contents_jupsu, gubun01, gubun02)

		''맞교환회수에는 회수되는 상품만 등록한다.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

        '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        call RegCSMasterAddUpche(newasid, requiremakerid)

		'// asid 연결
		Call SetRefAsid(newasid, iAsID)

		'// 업배상품 품절정보 저장
		if (modifyitemstockoutyn = "Y") then
	        ResultCount   = SetStockOutByCsAs(newasid)
		end if

		response.write "<script>alert('상품변경 맞교환 접수완료 - 업체배송');</script>"

    else
        '텐바이텐 배송의 경우 상품변경 맞교환 회수 접수
        newasid = RegCSMaster("A111", orderserial, reguserid, "교환회수(상품변경)", contents_jupsu, gubun01, gubun02)

		''맞교환회수에는 회수되는 상품만 등록한다.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// 고객 추가배송비(교환회수에 등록)
		Call SetCustomerAddBeasongPay(newasid, add_customeraddmethod, add_customeraddbeasongpay, "N", 0)

		'// asid 연결
		Call SetRefAsid(newasid, iAsID)

        response.write "<script>alert('옵션변경 맞교환 출고 접수 및 회수접수완료 - 텐바이텐 배송');</script>"
    end if


	response.write "<script>opener.parent.location.reload();</script>"
	response.write "<script>window.resizeTo(1200,600)</script>"

	if (requiremakerid<>"") then
		response.write "<script>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(iAsID) + "&mode=editreginfo'</script>"
	else
		'// 텐배는 맞교환 회수창으로 이동
		response.write "<script>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(newasid) + "&mode=editreginfo'</script>"
	end if
	response.end

else

	response.write "정의되지 않았습니다."
	response.end

end if


response.write "<script>alert('수정 되었습니다.');</script>"
response.write "<script>location.replace('/cscenter/ordermaster/orderdetail_editoption.asp?idx=" + detailidx + "');</script>"



%>

<!-- #include virtual="/lib/db/dbclose.asp" -->