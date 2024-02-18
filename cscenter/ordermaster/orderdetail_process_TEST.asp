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
dim orderserial, detailidx, mode
dim buycash, isupchebeasong, songjangdiv, songjangno
dim beasongdate, currstate, upcheconfirmdate
dim requiredetail, itemno, omwdiv, odlvType

dim itemId, preItemOption, itemOption, forceedit, ItemOptionName, preItemOptionName

dim fromItemId, fromItemOption, toItemId, toItemOption, itemnocancel, copysaleinfo, itemcouponinfo, itemcouponidx, bonuscouponidx
dim SalePrice, ItemCouponPrice, BonusCouponPrice
dim itemname
dim strsql

dim jungsanExists

dim newAsId

dim refundrequire, canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum, refundbeasongpay, refunddeliverypay, refundgiftcardsum, refunddepositsum, refundadjustpay

dim requiremakerid

dim arrFromItemId, arrFromItemOption, arrFromItemNo
dim arrToItemId, arrToItemOption, arrToItemNo, arrToItemCouponIdx
dim arrToSalePrice, arrToItemCouponPrice, arrToBonusCouponPrice, arrToBuyCash
dim toSaleMethod, toBonusCouponIdx
dim arrFromDetailIdx

1
사용안함

function recalcuOrderMasterCouponInfo(byVal orderserial)
	dim sqlStr

	sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
	sqlStr = sqlStr + " set tencardspend=IsNULL(T.tencardspend,0)" + VbCrlf
	''sqlStr = sqlStr + " , totalcost=IsNULL(T.dtotalsum,0)"  + VbCrlf
	sqlStr = sqlStr + " , totalmileage=IsNULL(T.dtotalmileage,0)" + VbCrlf
	sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=IsNULL(T.dtotalitemcostCouponNotApplied,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select sum((itemcost - reducedPrice)*itemno) as tencardspend, sum(mileage*itemno) as dtotalmileage, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied" + VbCrlf
	sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_detail" + VbCrlf
	sqlStr = sqlStr + "     where orderserial='" + orderserial + "'" + VbCrlf
	sqlStr = sqlStr + "     and cancelyn<>'Y'" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.orderserial='" + orderserial + "' and tencardspend <> 0" + VbCrlf

	dbget.Execute sqlStr
end function

recalcuOrderMaster("13090247696")
response.write "TEST"
response.end

orderserial     = request("orderserial")
detailidx       = request("detailidx")
mode            = request("mode")
buycash         = request("buycash")
isupchebeasong  = request("isupchebeasong")
songjangdiv     = request("songjangdiv")
songjangno      = request("songjangno")

currstate       = request("currstate")
upcheconfirmdate = request("upcheconfirmdate")
beasongdate     = request("beasongdate")
requiredetail   = html2db(request("requiredetail"))
itemno          = request("itemno")
omwdiv          = request("omwdiv")
odlvType        = request("odlvType")

forceedit       = request("forceedit")
itemId        	= request("itemId")
preItemOption   = request("preItemOption")
preItemOptionName  = request("preItemOptionName")
itemOption      = request("itemOption")
ItemOptionName  = request("ItemOptionName")

requiremakerid  = request("requiremakerid")



dim tmp


Dim divCd, regUserID, finishUser, title, contents_jupsu, gubun01, gubun02
Dim iAsID, contents_finish
dim itemoptioncode, itemoptionno, totalcountchanged
dim detailitemlist, newdetailitemlist, orderdetailidx

dim result, i

title = request("title")
contents_jupsu = request("contents_jupsu")
contents_finish = request("contents_finish")
itemoptioncode = request("itemoptioncode")
itemoptionno = request("itemoptionno")

gubun01 = request("gubun01")
gubun02 = request("gubun02")

if (mode="itemOption") then
	'상품옵션변경

	if (forceedit = "Y") then
		result = CSOrderModifyItemOptionForce(orderserial, itemId, preItemOption, itemOption)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderModifyItemOption(orderserial, itemId, preItemOption, itemOption)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' 주문내역변경
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="RestoreCancel") then

	'부분취소 정상화
	if (forceedit = "Y") then
		result = CSOrderRestoreCanceledItemForce(orderserial, itemId, preItemOption)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderRestoreCanceledItem(orderserial, itemId, preItemOption)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' 주문내역변경
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="Cancel") then

	'부분취소
	if (forceedit = "Y") then
		result = CSOrderCancelItemForce(orderserial, itemId, preItemOption)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderCancelItem(orderserial, itemId, preItemOption)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' 주문내역변경
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="EditItemNo") then

	'수량변경
	if (forceedit = "Y") then
		result = CSOrderModifyItemNoForce(orderserial, itemId, preItemOption, itemno)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderModifyItemNo(orderserial, itemId, preItemOption, itemno)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' 주문내역변경
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="EditItemNoPart") then

	itemoptioncode = SPlit(itemoptioncode, ",")
	itemoptionno = SPlit(itemoptionno, ",")
	ItemOptionName = SPlit(ItemOptionName, ",")

	totalcountchanged = 0
	detailitemlist = ""
	contents_jupsu = ""
	contents_finish = "상품옵션변경이 정상적으로 처리되었습니다."

	divCd = "A900"	' 주문내역변경
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	for i = 0 to UBound(itemoptionno)
		if ((preItemOption <> Trim(itemoptioncode(i))) and (CInt(itemoptionno(i)) > 0)) then

			if (forceedit = "Y") then

				'response.write "aaaaaaaaaaaaaaaaaaaaaaaaa"
				'response.end

				result = CSOrderModifyItemOptionForce(orderserial, itemId, preItemOption, Trim(itemoptioncode(i)), Trim(itemoptionno(i)))
				CSOrderRecalculateOrder orderserial,false
			else
				result = CSOrderModifyItemOption(orderserial, itemId, preItemOption, Trim(itemoptioncode(i)), Trim(itemoptionno(i)))
			end if

			if (CS_ORDER_FUNCTION_RESULT <> "") then
			    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if

			totalcountchanged = totalcountchanged + CInt(itemoptionno(i))

			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			result = CSOrderGetItemState(orderserial, itemId, Trim(itemoptioncode(i)))
			orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX

			ResetGlobalVarible()
			'--------------------------------------------------------------------------

	        detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & Trim(itemoptionno(i)) & Chr(9)
			contents_jupsu	= contents_jupsu & "[" & Trim(itemoptioncode(i)) & "] " & Trim(ItemOptionName(i)) & " " & Trim(itemoptionno(i)) & "개 추가" & vbCrLf

		end if
	next



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, itemId, preItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*totalcountchanged) & Chr(9)
	contents_jupsu	= contents_jupsu & "[" & preItemOption & "] " & preItemOptionName & " " & CStr(totalcountchanged) & "개 취소" & vbCrLf


	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

    ''2011-07-20 서동석 추가. //작업 후에 등록하여 원수량이 맞지 않아서. 추가====================
    strSql = " update D" & VbCRLF
    strSql = strSql & " set orderitemno=orderitemno-confirmitemno" & VbCRLF
    strSql = strSql & " from  db_cs.dbo.tbl_new_as_list A" & VbCRLF
    strSql = strSql & " 	Join db_cs.dbo.tbl_new_as_detail D" & VbCRLF
    strSql = strSql & " 	on A.id=D.masterid" & VbCRLF
    strSql = strSql & " where A.id="&iAsID

    dbget.Execute strSql
    '''==========================================================================================

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))


	response.write "<script>alert('수정 되었습니다.');</script>"
	response.write "<script>opener.parent.location.reload();</script>"
	response.write "<script>opener.focus();</script>"
	response.write "<script>window.close();</script>"
	response.end


	'response.write divcd & "<br>"
	'response.write reguserid & "<br>"
	'response.write title & "<br>"
	'response.write contents_jupsu & "<br>"
	'response.write gubun01 & "<br>"
	'response.write gubun02 & "<br>"
	'response.write contents_finish & "<br>"
	'response.write finishuser & "<br>"


elseif (mode="ChangeEditItemNoPart") then

	itemoptioncode = SPlit(itemoptioncode, ",")
	itemoptionno = SPlit(itemoptionno, ",")
	ItemOptionName = SPlit(ItemOptionName, ",")

	totalcountchanged = 0
	detailitemlist = ""
	newdetailitemlist = ""
	contents_jupsu = ""
	contents_finish = ""

	divCd = "A100"	' 상품변경 맞교환출고
	regUserID	= session("ssBctID")
	''finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	'--------------------------------------------------------------------------
	dim tenbeasongpay, upchebeasongpay, add_upchejungsandeliverypay, add_upchejungsancause, add_upchejungsancauseText

	tenbeasongpay = 2000			'// 텐바이텐 기본 배송비
	upchebeasongpay = 0
	add_upchejungsandeliverypay = 0
	add_upchejungsancause = ""
	add_upchejungsancauseText = ""

	'--------------------------------------------------------------------------
	dim oupchebeasongpay

	set oupchebeasongpay = new COrderMaster

	if (orderserial<>"") and (requiremakerid<>"") then
		oupchebeasongpay.FRectOrderSerial = orderserial
		oupchebeasongpay.getUpcheBeasongPayList

		for i = 0 to oupchebeasongpay.FResultCount - 1
			if (oupchebeasongpay.FItemList(i).Fmakerid = requiremakerid) then
				'// 업체배송이면 업체 기본배송비 가져오기
				upchebeasongpay = oupchebeasongpay.FItemList(i).Fdefaultdeliverpay
			end if
		next
	end if

	'--------------------------------------------------------------------------
	'// 단순변심 2배, 그 이외 0원
	if (gubun01 = "C004") and (gubun02 = "CD01") then
		tenbeasongpay = tenbeasongpay * 2
		upchebeasongpay = upchebeasongpay * 2

		if (orderserial<>"") and (requiremakerid<>"") then

			if (upchebeasongpay = 0) then
				'// XXXX 업체무료배송이면 텐텐배송비로 설정
				'기본배송비 설정 않되어 있으면 2500원(since 2012-06-18)
				upchebeasongpay = 2500
			end if

		end if
	else
		tenbeasongpay = 0
		upchebeasongpay = 0
	end if

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, itemId, preItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	for i = 0 to UBound(itemoptionno)
		if ((preItemOption <> Trim(itemoptioncode(i))) and (CInt(itemoptionno(i)) > 0)) then

			totalcountchanged = totalcountchanged + CInt(itemoptionno(i))

			'// 출고될 상품(한가지 이상일 수 있다.)
	        newdetailitemlist = newdetailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & Trim(itemoptionno(i)) & Chr(9) & Trim(itemId) & Chr(9) & Trim(itemoptioncode(i)) & Chr(9)
			contents_jupsu	= contents_jupsu & "[" & Trim(itemoptioncode(i)) & "] " & Trim(ItemOptionName(i)) & " " & Trim(itemoptionno(i)) & "개 출고" & vbCrLf

		end if
	next

	'// 회수될 상품(한가지만 가능)
    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(totalcountchanged) & Chr(9)
	contents_jupsu	= contents_jupsu & "[" & preItemOption & "] " & preItemOptionName & " " & CStr(totalcountchanged) & "개 회수" & vbCrLf & vbCrLf

	if (Not IsNull(session("ssBctCname"))) then
		contents_jupsu	= contents_jupsu & "텐바이텐 고객센터 " + CStr(session("ssBctCname")) + " 입니다" & vbCrLf
	end if


	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'//  CS Detail(관련상품목록) 등록
	'// 맞교환출고에는 출고되는 상품만 등록한다.
	Call AddCSDetailWithoutOrderDetailByArrStr(newdetailitemlist, iAsID, orderserial)

	'// CS 맞교환출고(동일상품, 상품변경 - A000, A100) 접수시 출고되는 상품 한정차감
	Call ApplyLimitItemByCS(iAsID)


    if (requiremakerid<>"") then
        '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        call RegCSMasterAddUpche(iAsID, requiremakerid)

		'// 고객 추가배송비
		Call SetCustomerAddBeasongPay(iAsID, "1", upchebeasongpay, "N", 0)			'// 1 = 박스동봉, N = 수령이전

		if (add_upchejungsandeliverypay <> 0) then
			Call RegCSUpcheAddJungsanPay(iAsID, add_upchejungsandeliverypay, add_upchejungsancause, requiremakerid)
		end if

        '업체배송인 경우 상품변경 맞교환 반품 접수
        newasid = RegCSMaster("A112", orderserial, reguserid, "교환회수(옵션변경,업체배송)", contents_jupsu, gubun01, gubun02)

		''맞교환반품에는 반품되는 상품만 등록한다.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

        '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        call RegCSMasterAddUpche(newasid, requiremakerid)

		'// asid 연결
		Call SetRefAsid(newasid, iAsID)

		response.write "<script>alert('옵션변경 맞교환 접수완료 - 업체배송');</script>"

    else
        '텐바이텐 배송의 경우 상품변경 맞교환 회수 접수
        newasid = RegCSMaster("A111", orderserial, reguserid, "교환회수(옵션변경)", contents_jupsu, gubun01, gubun02)

		''맞교환회수에는 회수되는 상품만 등록한다.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// 고객 추가배송비
		Call SetCustomerAddBeasongPay(newasid, "1", tenbeasongpay, "N", 0)			'// 1 = 박스동봉, N = 수령이전

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

elseif (mode="itemChange") then

	fromItemId     	= request("fromItemId")
	fromItemOption  = request("fromItemOption")
	toItemId     	= request("toItemId")
	toItemOption    = request("toItemOption")
	itemnocancel    = request("itemnocancel")
	copysaleinfo    = request("applyToAddItem")
	''itemcouponinfo  = request("itemcouponinfo")

	SalePrice  			= request("toSalePrice")
	ItemCouponPrice  	= request("toItemCouponPrice")
	BonusCouponPrice  	= request("toBonusCouponPrice")
	buycash  			= request("toAddBuycash")

	if (copysaleinfo = "Y") then
		itemcouponidx = request("fromItemCouponIdx")
		''bonuscouponidx = request("fromBonusCouponIdx")
	end if

	refundrequire		= request("refundrequire")
	canceltotal			= request("canceltotal")
	refunditemcostsum	= request("refunditemcostsum")
	refundcouponsum		= request("refundcouponsum")
	allatsubtractsum	= request("allatsubtractsum")

	detailitemlist = ""
	contents_jupsu = ""
	contents_finish = request("contents_finish")

	divCd 			= "A900"	' 주문내역변경
	regUserID		= session("ssBctID")
	finishuser		= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if


	'===========================================================================
    jungsanExists = false
    strSql = "select top 1 * from db_order.dbo.tbl_order_detail od"
    strSql = strSql & " Join db_jungsan.dbo.tbl_designer_jungsan_detail jd"
    strSql = strSql & " on od.idx=jd.detailidx"
    strSql = strSql & " where od.orderserial='" & orderserial & "' and od.idx = " & detailidx & " "

    rsget.Open strSql,dbget,1
    if Not rsget.Eof then
        jungsanExists = true
    end if
    rsget.Close

    if (jungsanExists) then
	    response.write "<script language='javascript'>alert('에러 : " & "정산 내역이 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
	    dbget.close()	:	response.End
    end if

	'===========================================================================
	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'환불정보
	newAsId = 0
	newAsId = RegCSMasterRefundInfoBeforeCancel(iAsID, orderserial, reguserid, refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, contents_finish, refundmileagesum, refunddepositsum, refundgiftcardsum)

	'==========================================================================
	if (forceedit = "Y") then
		result = CSOrderChangeItemForce(orderserial, fromItemId, toItemId, fromItemOption, toItemOption, itemnocancel)
	else
		result = CSOrderChangeItem(orderserial, fromItemId, toItemId, fromItemOption, toItemOption, itemnocancel)
	end if

	'금액 이외 정보
	Call CSOrderCopyItemInfoPart(orderserial, fromItemId, toItemId, fromItemOption, toItemOption)

	'금액정보
	Call CSOrderSetItemPriceInfo(orderserial, toItemId, toItemOption, SalePrice, ItemCouponPrice, BonusCouponPrice, buycash)

	if (copysaleinfo = "Y") then
		if (itemcouponidx <> "") then
			'상품쿠폰
			Call CSOrderSetItemCouponInfo(orderserial, toItemId, toItemOption, itemcouponidx)
		end if

		'보너스쿠폰
		Call CSOrderCopyBonusCouponInfo(orderserial, fromItemId, toItemId, fromItemOption, toItemOption)
	end if

	Call EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	CSOrderRecalculateOrder orderserial,false

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, toItemId, toItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
	itemname = CS_ORDER_ITEM_ITEMNAME
	itemoptionname = CS_ORDER_ITEM_OPTIONNAME

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(itemnocancel) & Chr(9)
	contents_jupsu	= contents_jupsu & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "개 추가" & vbCrLf



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, fromItemId, fromItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
	itemname = CS_ORDER_ITEM_ITEMNAME
	itemoptionname = CS_ORDER_ITEM_OPTIONNAME

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*itemnocancel) & Chr(9)
	contents_jupsu	= contents_jupsu & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "개 취소" & vbCrLf

	'==========================================================================
	' CS 마스타 AS 수정
	''html2db 사용하지 말것.
	Call EditCSMaster(iAsID, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	response.write "<script>" & vbCrLf
	response.write "	alert('수정 되었습니다.');" & vbCrLf
	response.write "	opener.parent.location.reload();" & vbCrLf

	if (newAsId <> 0) then
		response.write "	var a = window.open('/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + CStr(newAsId) + "&mode=editreginfo','pop_cs_action_reg_','width=1200 height=600 scrollbars=yes resizable=yes');" & vbCrLf
	else
		response.write "	window.blur();" & vbCrLf
	end if

	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	response.end

elseif (mode="itemChangeArray") then

	arrFromItemId			= request("arrFromItemId")
	arrFromItemOption		= request("arrFromItemOption")
	arrFromItemNo 			= request("arrFromItemNo")

	arrToItemId				= request("arrToItemId")
	arrToItemOption			= request("arrToItemOption")
	arrToItemNo				= request("arrToItemNo")
	arrToItemCouponIdx		= request("arrToItemCouponIdx")

	arrToSalePrice			= request("arrToSalePrice")
	arrToItemCouponPrice	= request("arrToItemCouponPrice")
	arrToBonusCouponPrice	= request("arrToBonusCouponPrice")
	arrToBuyCash			= request("arrToBuyCash")

	toSaleMethod			= request("toSaleMethod")
	toBonusCouponIdx		= request("toBonusCouponIdx")
	arrFromDetailIdx		= request("arrFromDetailIdx")

	refundrequire			= request("refundrequire")
	canceltotal				= request("canceltotal")
	refunditemcostsum		= request("refunditemcostsum")
	refundcouponsum			= request("refundcouponsum")
	allatsubtractsum		= request("allatsubtractsum")



	detailitemlist = ""
	contents_jupsu = ""
	contents_finish = request("contents_finish")

	divCd 			= "A900"	' 주문내역변경
	regUserID		= session("ssBctID")
	finishuser		= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	detailidx = "0" & Replace(arrFromDetailIdx, "|", ",")

    jungsanExists = false
    strSql = "select top 1 * from db_order.dbo.tbl_order_detail od"
    strSql = strSql & " Join db_jungsan.dbo.tbl_designer_jungsan_detail jd"
    strSql = strSql & " on od.idx=jd.detailidx"
    strSql = strSql & " where od.orderserial='" & orderserial & "' and od.idx in (" & detailidx & ") "

    rsget.Open strSql,dbget,1
    if Not rsget.Eof then
        jungsanExists = true
    end if
    rsget.Close

    if (jungsanExists) then
	    response.write "<script language='javascript'>alert('에러 : " & "정산 내역이 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
	    dbget.close()	:	response.End
    end if

	'==========================================================================
	' CS 마스타 AS 등록
	''html2db 사용하지 말것.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'환불정보
	newAsId = 0
	newAsId = RegCSMasterRefundInfoBeforeCancel(iAsID, orderserial, reguserid, refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, contents_finish, refundmileagesum, refunddepositsum, refundgiftcardsum)

	'==========================================================================
	if (forceedit = "Y") then
		result = CSOrderChangeItemArrayForce(orderserial, arrFromItemId, arrToItemId, arrFromItemOption, arrToItemOption, arrFromItemNo, arrToItemNo)
	else
		result = CSOrderChangeItemArray(orderserial, arrFromItemId, arrToItemId, arrFromItemOption, arrToItemOption, arrFromItemNo, arrToItemNo)
	end if


	arrFromItemId		= Split(arrFromItemId, "|")
	arrFromItemOption	= Split(arrFromItemOption, "|")
	arrFromItemNo		= Split(arrFromItemNo, "|")

	for i = 0 to UBound(arrFromItemId)
		if (Trim(arrFromItemId(i)) <> "") then
			fromItemId = Trim(arrFromItemId(i))
			fromItemOption = Trim(arrFromItemOption(i))
			itemnocancel = Trim(arrFromItemNo(i))

			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			result = CSOrderGetItemState(orderserial, fromItemId, fromItemOption)
			orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
			itemname = CS_ORDER_ITEM_ITEMNAME
			itemoptionname = CS_ORDER_ITEM_OPTIONNAME

			ResetGlobalVarible()
			'--------------------------------------------------------------------------

		    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*itemnocancel) & Chr(9)
			contents_jupsu	= contents_jupsu & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "개 취소" & vbCrLf
		end if
	next

	arrToItemId		= Split(arrToItemId, "|")
	arrToItemOption	= Split(arrToItemOption, "|")
	arrToItemNo		= Split(arrToItemNo, "|")

	arrToSalePrice			= Split(arrToSalePrice, "|")
	arrToItemCouponPrice	= Split(arrToItemCouponPrice, "|")
	arrToBonusCouponPrice	= Split(arrToBonusCouponPrice, "|")
	arrToBuyCash			= Split(arrToBuyCash, "|")

	arrToItemCouponIdx		= Split(arrToItemCouponIdx, "|")

	for i = 0 to UBound(arrToItemId)
		if (Trim(arrToItemId(i)) <> "") then
			toItemId = Trim(arrToItemId(i))
			toItemOption = Trim(arrToItemOption(i))
			itemnocancel = Trim(arrToItemNo(i))

			SalePrice = Trim(arrToSalePrice(i))
			ItemCouponPrice = Trim(arrToItemCouponPrice(i))
			BonusCouponPrice = Trim(arrToBonusCouponPrice(i))
			buycash = Trim(arrToBuyCash(i))

			'금액 이외 정보(첫번째 취소상품정보로 일괄복사)
			Call CSOrderCopyItemInfoPart(orderserial, fromItemId, toItemId, fromItemOption, toItemOption)

			'금액정보
			Call CSOrderSetItemPriceInfo(orderserial, toItemId, toItemOption, SalePrice, ItemCouponPrice, BonusCouponPrice, buycash)

			if (Trim(arrToItemCouponIdx(i)) <> "") then
				if (Trim(arrToItemCouponIdx(i)) <> "0") then
					itemcouponidx = Trim(arrToItemCouponIdx(i))

					'상품쿠폰
					Call CSOrderSetItemCouponInfo(orderserial, toItemId, toItemOption, itemcouponidx)
				end if
			end if

			if (ItemCouponPrice <> BonusCouponPrice) then
				if (toBonusCouponIdx <> "") and (toBonusCouponIdx <> "0") then
					'보너스쿠폰
					Call CSOrderSetBonusCouponInfo(orderserial, toItemId, toItemOption, toBonusCouponIdx)
				end if
			end if

			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			result = CSOrderGetItemState(orderserial, toItemId, toItemOption)
			orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
			itemname = CS_ORDER_ITEM_ITEMNAME
			itemoptionname = CS_ORDER_ITEM_OPTIONNAME

			ResetGlobalVarible()
			'--------------------------------------------------------------------------

		    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(itemnocancel) & Chr(9)
			contents_jupsu	= contents_jupsu & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "개 추가" & vbCrLf
		end if
	next

	Call EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	CSOrderRecalculateOrder orderserial,false

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	'==========================================================================
	' CS 마스타 AS 수정
	''html2db 사용하지 말것.
	Call EditCSMaster(iAsID, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(관련상품목록) 등록
	'옵션변경에서는 상세정보 생략
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	response.write "<script>" & vbCrLf
	response.write "	alert('수정 되었습니다.');" & vbCrLf
	response.write "	opener.parent.location.reload();" & vbCrLf

	if (newAsId <> 0) then
		response.write "	var a = window.open('/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + CStr(newAsId) + "&mode=editreginfo','pop_cs_action_reg_','width=1200 height=600 scrollbars=yes resizable=yes');" & vbCrLf
	else
		response.write "	window.blur();" & vbCrLf
	end if

	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	response.end

end if


response.write "<script>alert('수정 되었습니다.');</script>"
response.write "<script>location.replace('/cscenter/ordermaster/orderdetail_editoption.asp?idx=" + detailidx + "');</script>"



%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
