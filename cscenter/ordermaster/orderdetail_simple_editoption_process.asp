<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs센터 옵션교환
' History : 이상구 생성
'			2023.09.05 한용민 수정(6개월이전 주문도 교환 가능하게 처리)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%
dim detailidx, mode, orderserial, itemid, orgitemoptionname, orgitemoptioncode, orgitemoptionno, contents_finish
dim arritemoptioncode, arritemoptionno, arritemoptionnames, itemoptioncode, itemoptionno, itemoptionname
dim gubun01, gubun02, divcd, title, contents_jupsu, contents_itemlist, isupchebeasong, requiremakerid, upchebeasongpay
dim add_customeraddbeasongpay, add_customeraddmethod, detailitemlist, newdetailitemlist, totalchangeno
dim regUserID, finishuser, result, orderdetailidx, jungsanExists, itemChulgoFinished, csCancelExists
dim modifyitemstockoutyn, ResultCount, strSql, iAsID, newasid, i, j
	detailidx       	= request("detailidx")
	mode            	= request("mode")
	orderserial			= request("orderserial")
	itemid				= request("itemid")
	orgitemoptionno		= request("orgitemoptionno")
	arritemoptioncode	= request("itemoptioncode")
	arritemoptionno		= request("itemoptionno")
	gubun01				= request("gubun01")
	gubun02				= request("gubun02")
	divcd				= request("divcd")
	title				= request("title")
	contents_jupsu		= request("contents_jupsu")
	isupchebeasong		= request("isupchebeasong")
	requiremakerid		= request("requiremakerid")
	upchebeasongpay		= request("upchebeasongpay")
	add_customeraddbeasongpay	= request("add_customeraddbeasongpay")
	add_customeraddmethod		= request("add_customeraddmethod")
	modifyitemstockoutyn		= request("modifyitemstockoutyn")

arritemoptioncode = Split(arritemoptioncode, ",")
arritemoptionno = Split(arritemoptionno, ",")

orgitemoptioncode = arritemoptioncode(0)
totalchangeno = CLng(orgitemoptionno) - CLng(arritemoptionno(0))

if (gubun01 = "") then
	gubun01		= "C004"
	gubun02		= "CD99"
end if

''과거 주문 내역인지 Check
GC_IsOLDOrder = CheckIsOldOrder(orderserial)

arritemoptionnames = GetItemOptionNames(itemid)

orgitemoptionname = ""
For j = 0 To UBound(arritemoptionnames,2)
	if (arritemoptionnames(0, j) = orgitemoptioncode) then
		orgitemoptionname = arritemoptionnames(1, j)
	end if
next

if (mode="EditItemNoPart") then
	'// ===========================================================================
	'// 주문변경(상품옵션변경)
	jungsanExists = false
	'// 취소되는 상품
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
	    response.write "<script type='text/javascript'>alert('에러 : " & "취소(회수) 되는 상품에 정산 내역이 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	csCancelExists = False

	'// 취소 CS
	''strSql = "select top 1 orderserial"
	''strSql = strSql & " from"
	''strSql = strSql & " [db_cs].[dbo].[tbl_new_as_list]"
	''strSql = strSql & " where orderserial = '" & orderserial & "' and divcd = 'A008' and deleteyn = 'N'"

	''rsget.Open strSql,dbget,1
	''if Not rsget.Eof then
	''    csCancelExists = true
	''end if
	''rsget.Close

	''if (csCancelExists) then
	''    response.write "<script type='text/javascript'>alert('에러 : " & "취소CS가 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
	''    dbget.close()	:	response.End
	''end if

	'// 추가되는 상품
	for i = 0 to UBound(arritemoptionno)

		itemChulgoFinished = False
		itemoptioncode = Trim(arritemoptioncode(i))
		itemoptionno = CLng(arritemoptionno(i))

		if (orgitemoptioncode <> itemoptioncode) and (itemoptionno > 0) then
			strSql = "select top 1 *, jd.detailidx as jungsandetailidx, od.currstate as orderdetailstate from db_order.dbo.tbl_order_detail od"
			strSql = strSql & " Left Join db_jungsan.dbo.tbl_designer_jungsan_detail jd"
			strSql = strSql & " on od.idx=jd.detailidx"
			strSql = strSql & " where od.orderserial='" & orderserial & "' and od.itemid = " & itemid & " and od.itemoption = '" & itemoptioncode & "' "

			rsget.Open strSql,dbget,1
			if Not rsget.Eof then
			    if Not IsNull(rsget("jungsandetailidx")) then
			    	jungsanExists = true
			    end if

			    if rsget("orderdetailstate") = "7" then
			    	itemChulgoFinished = true
			    end if
			end if
			rsget.Close

			if (jungsanExists) then
			    response.write "<script type='text/javascript'>alert('에러 : " & "추가되는 상품에 정산 내역이 존재합니다. 변경할 수 없습니다." & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if

			if (itemChulgoFinished) then
			    response.write "<script type='text/javascript'>alert('에러 : " & "추가되는 상품이 출고완료입니다. 변경할 수 없습니다." & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if
		end if
	next

	contents_finish = "상품옵션변경이 정상적으로 처리되었습니다."

	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	for i = 0 to UBound(arritemoptionno)

		itemoptioncode = Trim(arritemoptioncode(i))
		itemoptionno = CLng(arritemoptionno(i))
		itemoptionname = ""

		For j = 0 To UBound(arritemoptionnames,2)
			if (arritemoptionnames(0, j) = itemoptioncode) then
				itemoptionname = arritemoptionnames(1, j)
			end if
		next

		if (orgitemoptioncode <> itemoptioncode) and (itemoptionno > 0) then

			'// 옵션변경
			result = CSOrderModifyItemOption(orderserial, itemId, orgitemoptioncode, itemoptioncode, itemoptionno)

			if (CS_ORDER_FUNCTION_RESULT <> "") then
			    response.write "<script type='text/javascript'>alert('에러 : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if

			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			result = CSOrderGetItemState(orderserial, itemId, itemoptioncode)
			orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX

			ResetGlobalVarible()
			'--------------------------------------------------------------------------

	        detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & itemoptionno & Chr(9)
			contents_itemlist	= contents_itemlist & "[" & itemoptioncode & "] " & itemoptionname & " " & itemoptionno & "개 추가" & vbCrLf
		end if

	next

    detailitemlist = detailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*totalchangeno) & Chr(9)
	contents_itemlist	= contents_itemlist & "[" & orgitemoptioncode & "] " & orgitemoptionname & " " & CStr(totalchangeno) & "개 취소" & vbCrLf

	'// 접수내용에 추가
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu

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

    '// 제휴몰 주문에 변경된 상품코드 입력, 2021-03-22, skyer9
    strSql = " exec [db_cs].[dbo].[usp_Ten_CsAs_ChangeItem2ExtSite] " & iAsID
    dbget.Execute strSql

    '''==========================================================================================

	' CS 마스타 AS완료
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' 내용변경
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	'// 업배상품 품절정보 저장
	if (modifyitemstockoutyn = "Y") then
        ResultCount   = SetStockOutByCsAs(iAsid)
	end if

	'// 업배인 경우 브랜드지정
	Call RegCSMasterAddUpcheByAsid(iAsID)

	response.write "<script type='text/javascript'>alert('수정 되었습니다.');</script>"
	response.write "<script type='text/javascript'>opener.parent.location.reload();</script>"
	response.write "<script type='text/javascript'>opener.focus();</script>"
	response.write "<script type='text/javascript'>window.close();</script>"
	dbget.close() : response.end

	'response.write divcd & "<br>"
	'response.write reguserid & "<br>"
	'response.write title & "<br>"
	'response.write contents_jupsu & "<br>"
	'response.write detailitemlist & "<br>"
	'response.write gubun01 & "<br>"
	'response.write gubun02 & "<br>"
	'response.write contents_finish & "<br>"
	'response.write finishuser & "<br>"
	'response.end

elseif (mode="ChangeEditItemNoPart") then
	''상품변경 맞교환

	regUserID	= session("ssBctID")

	for i = 0 to UBound(arritemoptionno)

		itemoptioncode = Trim(arritemoptioncode(i))
		itemoptionno = CLng(arritemoptionno(i))
		itemoptionname = ""

		For j = 0 To UBound(arritemoptionnames,2)
			if (arritemoptionnames(0, j) = itemoptioncode) then
				itemoptionname = arritemoptionnames(1, j)
			end if
		next

		if (orgitemoptioncode <> itemoptioncode) and (itemoptionno > 0) then

			'// 출고될 상품(한가지 이상일 수 있다.)
	        newdetailitemlist = newdetailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & itemoptionno & Chr(9) & Trim(itemId) & Chr(9) & itemoptioncode & Chr(9)
			contents_itemlist	= contents_itemlist & "[" & itemoptioncode & "] " & itemoptionname & " " & itemoptionno & "개 출고" & vbCrLf

		end if
	next

	'// 회수될 상품(한가지만 가능)
    detailitemlist = detailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(totalchangeno) & Chr(9)
	contents_itemlist	= contents_itemlist & "[" & orgitemoptioncode & "] " & orgitemoptionname & " " & CStr(totalchangeno) & "개 회수" & vbCrLf

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
        newasid = RegCSMaster("A112", orderserial, reguserid, "교환회수(옵션변경,업체배송)", contents_jupsu, gubun01, gubun02)

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

		response.write "<script type='text/javascript'>alert('옵션변경 맞교환 접수완료 - 업체배송');</script>"
    else
        '텐바이텐 배송의 경우 상품변경 맞교환 회수 접수
        newasid = RegCSMaster("A111", orderserial, reguserid, "교환회수(옵션변경)", contents_jupsu, gubun01, gubun02)

		''맞교환회수에는 회수되는 상품만 등록한다.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// 고객 추가배송비(교환회수에 등록)
		Call SetCustomerAddBeasongPay(newasid, add_customeraddmethod, add_customeraddbeasongpay, "N", 0)

		'// asid 연결
		Call SetRefAsid(newasid, iAsID)

        response.write "<script type='text/javascript'>alert('옵션변경 맞교환 출고 접수 및 회수접수완료 - 텐바이텐 배송');</script>"
    end if

	response.write "<script type='text/javascript'>opener.parent.location.reload();</script>"
	'response.write "<script type='text/javascript'>window.resizeTo(1400,800)</script>"

	if (requiremakerid<>"") then
		response.write "<script type='text/javascript'>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(iAsID) + "&mode=editreginfo'</script>"
	else
		'// 텐배는 맞교환 회수창으로 이동
		response.write "<script type='text/javascript'>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(newasid) + "&mode=editreginfo'</script>"
	end if
	dbget.close() : response.end
else
	response.write "정의되지 않았습니다."
	dbget.close() : response.end
end if

response.write "<script type='text/javascript'>alert('수정 되었습니다.');</script>"
response.write "<script type='text/javascript'>location.replace('/cscenter/ordermaster/orderdetail_editoption.asp?idx=" + detailidx + "');</script>"

Function GetItemOptionNames(itemid)
	dim sqlStr

	if itemid="" or isnull(itemid) then
		GetItemOptionNames = ""
		exit Function
	end if

	GetItemOptionNames = ""

	sqlStr = " select "
	sqlStr = sqlStr + " v.itemoption "
	sqlStr = sqlStr + " , v.optionname "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i with (nolock)"
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v with (nolock)"
	sqlStr = sqlStr + " 	on i.itemid=v.itemid "
	sqlStr = sqlStr + " WHERE i.itemid=" & itemid & ""
	sqlStr = sqlStr + " order by i.itemid desc, v.itemoption"

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not rsget.EOF Then
		GetItemOptionNames = rsget.getrows
	End If
	rsget.close()
end Function
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
