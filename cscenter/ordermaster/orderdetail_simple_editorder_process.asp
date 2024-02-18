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

'ǰ����� ��ǰ���� ����
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

'' ���� 6���� ���� ���� �˻�
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
	'// �ֹ�����(��ǰ����)

	'==============================================================================
	jungsanExists = false
	'// ��ҵǴ� ��ǰ
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
	    response.write "<script language='javascript'>alert('���� : " & "���(ȸ��) �Ǵ� ��ǰ�� ���� ������ �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	'// �߰��Ǵ� ��ǰ
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
	    response.write "<script language='javascript'>alert('���� : " & "�߰��Ǵ� ��ǰ�ڵ忡 ���� ������ �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	'==============================================================================
	fromDetailState = ""
	toDetailState = ""

	'// �߰��Ǵ� ��ǰ�� �̹� �����Ͽ� �ִ� ��� ����üũ(���Ϸ� ��ǰ�� ���� ��ǰ�� ��ĥ �� ����.)
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
		    response.write "<script language='javascript'>alert('���� : " & "���Ϸ� ��ǰ�� ���� ��ǰ�� ��ĥ �� �����ϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
		    dbget.close()	:	response.End
		end if

	end if


	'==========================================================================
	contents_finish = "��ǰ������ ���������� ó���Ǿ����ϴ�."

	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	'// �߰��� ��ǰ(�Ѱ����� ����)
	contents_itemlist	= contents_itemlist & GetItemName(add_itemid) & vbCrLf & "[" & add_itemoption & "] " & GetItemOptionName(add_itemid, add_itemoption) & " " & targetregitemno & "�� �߰�" & vbCrLf & vbCrLf

	'// ��ҵ� ��ǰ(�Ѱ����� ����)
	contents_itemlist	= contents_itemlist & fromItemName & vbCrLf & "[" & fromItemOption & "] " & fromItemOptionName & " " & CStr(targetregitemno) & "�� ���" & vbCrLf

	'// �������뿡 �߰�
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu


	'==========================================================================
	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'ȯ������
	newAsId = 0
	newAsId = RegCSMasterRefundInfoBeforeCancel(iAsID, orderserial, reguserid, refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, contents_finish, refundmileagesum, refunddepositsum, refundgiftcardsum)

	result = CSOrderChangeItemForce(orderserial, fromItemId, add_itemid, fromItemOption, add_itemoption, targetregitemno)

	'�ݾ� �̿� ����
	Call CSOrderCopyItemInfoPart(orderserial, fromItemId, add_itemid, fromItemOption, add_itemoption)

	'�ݾ�����
	Call CSOrderSetItemPriceInfo(orderserial, add_itemid, add_itemoption, add_SalePrice, add_ItemCouponPrice, add_BonusCouponPrice, add_buycash)

	if (iscouponapplied = "Y") then
		if (itemcouponidxapplied <> "") then
			'��ǰ����
			Call CSOrderSetItemCouponInfo(orderserial, add_itemid, add_itemoption, itemcouponidxapplied)
		else
			'���ʽ�����
			Call CSOrderCopyBonusCouponInfo(orderserial, fromItemId, add_itemid, fromItemOption, add_itemoption)
		end if
	end if


	Call EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	CSOrderRecalculateOrder orderserial,false

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
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
	' CS ����Ÿ AS ����
	Call EditCSMaster(iAsID, reguserid, title, html2db(contents_jupsu), gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	'// �����ǰ ǰ������ ����
	if (modifyitemstockoutyn = "Y") then
        ResultCount   = SetStockOutByCsAs(iAsid)
	end if

	response.write "<script>" & vbCrLf
	response.write "	alert('���� �Ǿ����ϴ�.');" & vbCrLf
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


	'// ���� ��ǰ(�Ѱ����� ����)
    newdetailitemlist = newdetailitemlist & "|" & targetdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & targetregitemno & Chr(9) & Trim(add_itemid) & Chr(9) & add_itemoption & Chr(9)
	contents_itemlist	= contents_itemlist & GetItemName(add_itemid) & vbCrLf & "[" & add_itemoption & "] " & GetItemOptionName(add_itemid, add_itemoption) & " " & targetregitemno & "�� ���" & vbCrLf & vbCrLf


	'// ȸ���� ��ǰ(�Ѱ����� ����)
    detailitemlist = detailitemlist & "|" & targetdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(targetregitemno) & Chr(9)
	contents_itemlist	= contents_itemlist & fromItemName & vbCrLf & "[" & fromItemOption & "] " & fromItemOptionName & " " & CStr(targetregitemno) & "�� ȸ��" & vbCrLf


	'// �������뿡 �߰�
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu


	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'//  CS Detail(���û�ǰ���) ���
	'// �±�ȯ����� ���Ǵ� ��ǰ�� ����Ѵ�.
	Call AddCSDetailWithoutOrderDetailByArrStr(newdetailitemlist, iAsID, orderserial)

	'// CS �±�ȯ���(���ϻ�ǰ, ��ǰ���� - A000, A100) ������ ���Ǵ� ��ǰ ��������
	Call ApplyLimitItemByCS(iAsID)

    if (isupchebeasong = "Y") then
        '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        call RegCSMasterAddUpche(iAsID, requiremakerid)

		'// �� �߰���ۺ�(��ȯ��� ���)
		Call SetCustomerAddBeasongPay(iAsID, add_customeraddmethod, add_customeraddbeasongpay, "N", 0)

        '��ü����� ��� ��ǰ���� �±�ȯ ȸ�� ����
        newasid = RegCSMaster("A112", orderserial, reguserid, "��ȯȸ��(��ǰ����,��ü���)", contents_jupsu, gubun01, gubun02)

		''�±�ȯȸ������ ȸ���Ǵ� ��ǰ�� ����Ѵ�.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

        '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        call RegCSMasterAddUpche(newasid, requiremakerid)

		'// asid ����
		Call SetRefAsid(newasid, iAsID)

		'// �����ǰ ǰ������ ����
		if (modifyitemstockoutyn = "Y") then
	        ResultCount   = SetStockOutByCsAs(newasid)
		end if

		response.write "<script>alert('��ǰ���� �±�ȯ �����Ϸ� - ��ü���');</script>"

    else
        '�ٹ����� ����� ��� ��ǰ���� �±�ȯ ȸ�� ����
        newasid = RegCSMaster("A111", orderserial, reguserid, "��ȯȸ��(��ǰ����)", contents_jupsu, gubun01, gubun02)

		''�±�ȯȸ������ ȸ���Ǵ� ��ǰ�� ����Ѵ�.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// �� �߰���ۺ�(��ȯȸ���� ���)
		Call SetCustomerAddBeasongPay(newasid, add_customeraddmethod, add_customeraddbeasongpay, "N", 0)

		'// asid ����
		Call SetRefAsid(newasid, iAsID)

        response.write "<script>alert('�ɼǺ��� �±�ȯ ��� ���� �� ȸ�������Ϸ� - �ٹ����� ���');</script>"
    end if


	response.write "<script>opener.parent.location.reload();</script>"
	response.write "<script>window.resizeTo(1200,600)</script>"

	if (requiremakerid<>"") then
		response.write "<script>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(iAsID) + "&mode=editreginfo'</script>"
	else
		'// �ٹ�� �±�ȯ ȸ��â���� �̵�
		response.write "<script>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(newasid) + "&mode=editreginfo'</script>"
	end if
	response.end

else

	response.write "���ǵ��� �ʾҽ��ϴ�."
	response.end

end if


response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
response.write "<script>location.replace('/cscenter/ordermaster/orderdetail_editoption.asp?idx=" + detailidx + "');</script>"



%>

<!-- #include virtual="/lib/db/dbclose.asp" -->