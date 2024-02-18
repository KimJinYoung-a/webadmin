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

Function GetItemOptionNames(fromItemId, fromItemOption, toItemId, toItemOption)
	dim sqlStr

	GetItemOptionNames = ""

	sqlStr = " select top 2 i.itemid, IsNull(v.itemoption, '0000') as itemoption, i.itemname, IsNull(v.optionname, '') "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option v "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		i.itemid=v.itemid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and i.itemid in (" & fromItemId & "," & toItemId & ") "
	sqlStr = sqlStr + " 	and ( "
	sqlStr = sqlStr + " 		(i.itemid = " & fromItemId & " and IsNull(v.itemoption, '0000') = '" & fromItemOption & "') "
	sqlStr = sqlStr + " 		or "
	sqlStr = sqlStr + " 		(i.itemid = " & toItemId & " and IsNull(v.itemoption, '0000') = '" & toItemOption & "') "
	sqlStr = sqlStr + " 	) "
	'response.write sqlStr

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		GetItemOptionNames = rsget.getrows
	End If
	rsget.close()
end Function


dim orderserial, detailidx, mode
dim buycash, isupchebeasong, songjangdiv, songjangno
dim beasongdate, currstate, upcheconfirmdate
dim requiredetail, itemno, omwdiv, odlvType

dim itemId, preItemOption, itemOption, forceedit, ItemOptionName, preItemOptionName

dim fromItemId, fromItemOption, toItemId, toItemOption, itemnocancel, itemnoadd, copysaleinfo, itemcouponinfo, itemcouponidx, bonuscouponidx
dim fromItemName, fromItemOptionName, toItemName, toItemOptionName
dim SalePrice, ItemCouponPrice, BonusCouponPrice, EtcDiscountPrice
dim itemname
dim strsql, sqlStr
dim arrItemInfo

dim jungsanExists, errExists

dim newAsId

dim refundrequire, canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum, refundbeasongpay, refunddeliverypay, refundgiftcardsum, refunddepositsum, refundadjustpay

dim requiremakerid

dim arrFromItemId, arrFromItemOption, arrFromItemNo
dim arrToItemId, arrToItemOption, arrToItemNo, arrToItemCouponIdx
dim arrToSalePrice, arrToItemCouponPrice, arrToBonusCouponPrice, arrToEtcDiscountPrice, arrToBuyCash
dim toSaleMethod, toBonusCouponIdx
dim arrFromDetailIdx
dim reducedPrice, OutMallOrderSeq


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
reducedPrice  	= request("reducedPrice")


dim tmp
Dim divCd, regUserID, finishUser, title, contents_jupsu, gubun01, gubun02, modifyitemstockoutyn, ResultCount
Dim iAsID, contents_finish
dim itemoptioncode, itemoptionno, totalcountchanged
dim detailitemlist, newdetailitemlist, orderdetailidx, contents_itemlist

dim result, i

title = request("title")
contents_jupsu = request("contents_jupsu")
contents_finish = request("contents_finish")
itemoptioncode = request("itemoptioncode")
itemoptionno = request("itemoptionno")

gubun01 = request("gubun01")
gubun02 = request("gubun02")
modifyitemstockoutyn = request("modifyitemstockoutyn")

if (mode="itemOption") then
	'��ǰ�ɼǺ���

	response.write "�ý����� ����"
	dbget.close : response.end

	if (forceedit = "Y") then
		result = CSOrderModifyItemOptionForce(orderserial, itemId, preItemOption, itemOption)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderModifyItemOption(orderserial, itemId, preItemOption, itemOption)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' �ֹ���������
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="RestoreCancel") then

	response.write "�ý����� ����"
	dbget.close : response.end

	'�κ���� ����ȭ
	if (forceedit = "Y") then
		result = CSOrderRestoreCanceledItemForce(orderserial, itemId, preItemOption)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderRestoreCanceledItem(orderserial, itemId, preItemOption)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' �ֹ���������
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="Cancel") then

	response.write "�ý����� ����"
	dbget.close : response.end

	'�κ����
	if (forceedit = "Y") then
		result = CSOrderCancelItemForce(orderserial, itemId, preItemOption)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderCancelItem(orderserial, itemId, preItemOption)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' �ֹ���������
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="EditItemNo") then

	response.write "�ý����� ����"
	dbget.close : response.end

	'��������
	if (forceedit = "Y") then
		result = CSOrderModifyItemNoForce(orderserial, itemId, preItemOption, itemno)
		CSOrderRecalculateOrder orderserial,false
	else
		result = CSOrderModifyItemNo(orderserial, itemId, preItemOption, itemno)
	end if

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	divCd = "A900"	' �ֹ���������
	regUserID	= session("ssBctID")
	finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	'Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

elseif (mode="EditItemNoPart") then

	response.write "�ý����� ����"
	dbget.close : response.end

	itemoptioncode = SPlit(itemoptioncode, ",")
	itemoptionno = SPlit(itemoptionno, ",")
	ItemOptionName = SPlit(ItemOptionName, ",")

	totalcountchanged = 0
	detailitemlist = ""
	contents_jupsu = ""
	contents_finish = "��ǰ�ɼǺ����� ���������� ó���Ǿ����ϴ�."

	divCd = "A900"	' �ֹ���������
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
			    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
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
			contents_jupsu	= contents_jupsu & "[" & Trim(itemoptioncode(i)) & "] " & Trim(ItemOptionName(i)) & " " & Trim(itemoptionno(i)) & "�� �߰�" & vbCrLf

		end if
	next



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, itemId, preItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*totalcountchanged) & Chr(9)
	contents_jupsu	= contents_jupsu & "[" & preItemOption & "] " & preItemOptionName & " " & CStr(totalcountchanged) & "�� ���" & vbCrLf


	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

    ''2011-07-20 ������ �߰�. //�۾� �Ŀ� ����Ͽ� �������� ���� �ʾƼ�. �߰�====================
    strSql = " update D" & VbCRLF
    strSql = strSql & " set orderitemno=orderitemno-confirmitemno" & VbCRLF
    strSql = strSql & " from  db_cs.dbo.tbl_new_as_list A" & VbCRLF
    strSql = strSql & " 	Join db_cs.dbo.tbl_new_as_detail D" & VbCRLF
    strSql = strSql & " 	on A.id=D.masterid" & VbCRLF
    strSql = strSql & " where A.id="&iAsID

    dbget.Execute strSql
    '''==========================================================================================

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))


	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
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

	response.write "�ý����� ����"
	dbget.close : response.end

	itemoptioncode = SPlit(itemoptioncode, ",")
	itemoptionno = SPlit(itemoptionno, ",")
	ItemOptionName = SPlit(ItemOptionName, ",")

	totalcountchanged = 0
	detailitemlist = ""
	newdetailitemlist = ""
	contents_jupsu = ""
	contents_finish = ""

	divCd = "A100"	' ��ǰ���� �±�ȯ���
	regUserID	= session("ssBctID")
	''finishuser	= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end if

	'--------------------------------------------------------------------------
	dim tenbeasongpay, upchebeasongpay, add_upchejungsandeliverypay, add_upchejungsancause, add_upchejungsancauseText

	tenbeasongpay = getDefaultBeasongPayByDate(Left(Now, 10))			'// ��ۺ�
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
				'// ��ü����̸� ��ü �⺻��ۺ� ��������
				upchebeasongpay = oupchebeasongpay.FItemList(i).Fdefaultdeliverpay
			end if
		next
	end if

	'--------------------------------------------------------------------------
	'// �ܼ����� 2��, �� �̿� 0��
	if (gubun01 = "C004") and (gubun02 = "CD01") then
		tenbeasongpay = tenbeasongpay * 2
		upchebeasongpay = upchebeasongpay * 2

		if (orderserial<>"") and (requiremakerid<>"") then

			if (upchebeasongpay = 0) then
				'// XXXX ��ü�������̸� ���ٹ�ۺ�� ����
				'�⺻��ۺ� ���� �ʵǾ� ������ 2500��(since 2012-06-18)
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

			'// ���� ��ǰ(�Ѱ��� �̻��� �� �ִ�.)
	        newdetailitemlist = newdetailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & Trim(itemoptionno(i)) & Chr(9) & Trim(itemId) & Chr(9) & Trim(itemoptioncode(i)) & Chr(9)
			contents_jupsu	= contents_jupsu & "[" & Trim(itemoptioncode(i)) & "] " & Trim(ItemOptionName(i)) & " " & Trim(itemoptionno(i)) & "�� ���" & vbCrLf

		end if
	next

	'// ȸ���� ��ǰ(�Ѱ����� ����)
    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(totalcountchanged) & Chr(9)
	contents_jupsu	= contents_jupsu & "[" & preItemOption & "] " & preItemOptionName & " " & CStr(totalcountchanged) & "�� ȸ��" & vbCrLf & vbCrLf

	if (Not IsNull(session("ssBctCname"))) then
		contents_jupsu	= contents_jupsu & "�ٹ����� ������ " + CStr(session("ssBctCname")) + " �Դϴ�" & vbCrLf
	end if


	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'//  CS Detail(���û�ǰ���) ���
	'// �±�ȯ����� ���Ǵ� ��ǰ�� ����Ѵ�.
	Call AddCSDetailWithoutOrderDetailByArrStr(newdetailitemlist, iAsID, orderserial)

	'// CS �±�ȯ���(���ϻ�ǰ, ��ǰ���� - A000, A100) ������ ���Ǵ� ��ǰ ��������
	Call ApplyLimitItemByCS(iAsID)


    if (requiremakerid<>"") then
        '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        call RegCSMasterAddUpche(iAsID, requiremakerid)

		'// �� �߰���ۺ�
		Call SetCustomerAddBeasongPay(iAsID, "1", upchebeasongpay, "N", 0)			'// 1 = �ڽ�����, N = ��������

		if (add_upchejungsandeliverypay <> 0) then
			Call RegCSUpcheAddJungsanPay(iAsID, add_upchejungsandeliverypay, add_upchejungsancause, requiremakerid)
		end if

        '��ü����� ��� ��ǰ���� �±�ȯ ��ǰ ����
        newasid = RegCSMaster("A112", orderserial, reguserid, "��ȯȸ��(�ɼǺ���,��ü���)", contents_jupsu, gubun01, gubun02)

		''�±�ȯ��ǰ���� ��ǰ�Ǵ� ��ǰ�� ����Ѵ�.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

        '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        call RegCSMasterAddUpche(newasid, requiremakerid)

		'// asid ����
		Call SetRefAsid(newasid, iAsID)

		response.write "<script>alert('�ɼǺ��� �±�ȯ �����Ϸ� - ��ü���');</script>"

    else
        '�ٹ����� ����� ��� ��ǰ���� �±�ȯ ȸ�� ����
        newasid = RegCSMaster("A111", orderserial, reguserid, "��ȯȸ��(�ɼǺ���)", contents_jupsu, gubun01, gubun02)

		''�±�ȯȸ������ ȸ���Ǵ� ��ǰ�� ����Ѵ�.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// �� �߰���ۺ�
		Call SetCustomerAddBeasongPay(newasid, "1", tenbeasongpay, "N", 0)			'// 1 = �ڽ�����, N = ��������

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

elseif (mode="itemChange") then

	fromItemId     	= request("fromItemId")
	fromItemOption  = request("fromItemOption")
	toItemId     	= request("toItemId")
	toItemOption    = request("toItemOption")
	itemnocancel    = request("itemnocancel")
	itemnoadd    	= request("itemnoadd")
	copysaleinfo    = request("applyToAddItem")
	''itemcouponinfo  = request("itemcouponinfo")

	SalePrice  			= request("toSalePrice")
	ItemCouponPrice  	= request("toItemCouponPrice")
	BonusCouponPrice  	= request("toBonusCouponPrice")
	EtcDiscountPrice  	= request("toEtcDiscountPrice")
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
	'' contents_jupsu = ""
	contents_finish = request("contents_finish")

	divCd 			= "A900"	' �ֹ���������
	regUserID		= session("ssBctID")
	finishuser		= session("ssBctID")

	if (gubun01 = "") then
		gubun01		= "C004"
		gubun02		= "CD99"
	end If


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
	    response.write "<script language='javascript'>alert('���� : " & "���� ������ �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
	    dbget.close()	:	response.End
    end if

	'===========================================================================
	'// ����߰� ��ǰ�� ��ۻ��� üũ
	strSql = " select sum(case when od.currstate = 7 then 1 else 0 end) as chulgoCnt, count(idx) as itemCnt from db_order.dbo.tbl_order_detail od "
    strSql = strSql & " where "
    strSql = strSql & " 	1 = 1 "
    strSql = strSql & " 	and od.orderserial='" & orderserial & "' "
    strSql = strSql & " 	and (od.idx = " & detailidx & " or (od.itemid = " & toItemId & " and od.itemoption = '" & toItemOption & "')) "

	errExists = False
    rsget.Open strSql,dbget,1
    if Not rsget.Eof then
		if rsget("itemCnt") > 1 and rsget("chulgoCnt") = 1 then
			errExists = true
		end if
    end if
    rsget.Close

    if (errExists) then
	    response.write "<script language='javascript'>alert('���� : ����߰���ǰ�� ��ۻ��°� ���� �ٸ��ϴ�.');history.back();</script>"
	    dbget.close()	:	response.End
    end if

	'===========================================================================
	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'ȯ������
	newAsId = 0
	newAsId = RegCSMasterRefundInfoBeforeCancel(iAsID, orderserial, reguserid, refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum, contents_finish, refundmileagesum, refunddepositsum, refundgiftcardsum)

	'==========================================================================
	if (forceedit = "Y") then
		result = CSOrderChangeItemForceNEW(orderserial, fromItemId, toItemId, fromItemOption, toItemOption, itemnocancel, itemnoadd)
	else
		result = CSOrderChangeItem(orderserial, fromItemId, toItemId, fromItemOption, toItemOption, itemnocancel)
	end if

	'�ݾ� �̿� ����
	Call CSOrderCopyItemInfoPart(orderserial, fromItemId, toItemId, fromItemOption, toItemOption)

	'�ݾ�����
	Call CSOrderSetItemPriceInfo(orderserial, toItemId, toItemOption, SalePrice, ItemCouponPrice, BonusCouponPrice, EtcDiscountPrice, buycash)

	if (copysaleinfo = "Y") then
		if (itemcouponidx <> "") then
			'��ǰ����
			Call CSOrderSetItemCouponInfo(orderserial, toItemId, toItemOption, itemcouponidx)
		end if

		'���ʽ�����
		Call CSOrderCopyBonusCouponInfo(orderserial, fromItemId, toItemId, fromItemOption, toItemOption)

		'// �ǸŰ�(���ΰ�) �����ϸ� ���Ÿ��ϸ��� ����(2014-04-29, skyer9)
		Call CSOrderUpdateBuyMileage(orderserial, fromItemId, fromItemOption, toItemId, toItemOption)
	end if

	Call EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	CSOrderRecalculateOrder orderserial,false

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
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

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(itemnoadd) & Chr(9)
	contents_itemlist	= contents_itemlist & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnoadd) & "�� �߰�" & vbCrLf



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	result = CSOrderGetItemState(orderserial, fromItemId, fromItemOption)
	orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX
	itemname = CS_ORDER_ITEM_ITEMNAME
	itemoptionname = CS_ORDER_ITEM_OPTIONNAME

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*itemnocancel) & Chr(9)
	contents_itemlist	= contents_itemlist & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "�� ���" & vbCrLf

	'// �������뿡 �߰�
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu
	'==========================================================================
	' CS ����Ÿ AS ����
	''html2db ������� ����.
	Call EditCSMaster(iAsID, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

    '// ���޸� �ֹ��� ����� ��ǰ�ڵ� �Է�, 2021-03-22, skyer9
    strSql = " exec [db_cs].[dbo].[usp_Ten_CsAs_ChangeItem2ExtSite] " & iAsID
    dbget.Execute strSql

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	'// �����ǰ ǰ������ ����
	if (modifyitemstockoutyn = "Y") then
        ResultCount   = SetStockOutByCsAs(iAsid)
	end if

	'// ������ ��� �귣������
	Call RegCSMasterAddUpcheByAsid(iAsID)

	response.write "<script>" & vbCrLf
	response.write "	alert('���� �Ǿ����ϴ�.');" & vbCrLf
	response.write "	opener.parent.location.reload();" & vbCrLf

	if (newAsId <> 0) then
		response.write "	var a = window.open('/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + CStr(newAsId) + "&mode=editreginfo','pop_cs_action_reg_','width=1200 height=600 scrollbars=yes resizable=yes');" & vbCrLf
	else
		response.write "	window.blur();" & vbCrLf
	end if

	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	dbget.close()
	response.end

elseif (mode="itemChangeArray") then

	response.write "�ý����� ����"
	dbget.close : response.end

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

	divCd 			= "A900"	' �ֹ���������
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
	    response.write "<script language='javascript'>alert('���� : " & "���� ������ �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
	    dbget.close()	:	response.End
    end if

	'==========================================================================
	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'ȯ������
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
			contents_jupsu	= contents_jupsu & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "�� ���" & vbCrLf
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

			'�ݾ� �̿� ����(ù��° ��һ�ǰ������ �ϰ�����)
			Call CSOrderCopyItemInfoPart(orderserial, fromItemId, toItemId, fromItemOption, toItemOption)

			'�ݾ�����
			Call CSOrderSetItemPriceInfo(orderserial, toItemId, toItemOption, SalePrice, ItemCouponPrice, BonusCouponPrice, buycash)

			if (Trim(arrToItemCouponIdx(i)) <> "") then
				if (Trim(arrToItemCouponIdx(i)) <> "0") then
					itemcouponidx = Trim(arrToItemCouponIdx(i))

					'��ǰ����
					Call CSOrderSetItemCouponInfo(orderserial, toItemId, toItemOption, itemcouponidx)
				end if
			end if

			if (ItemCouponPrice <> BonusCouponPrice) then
				if (toBonusCouponIdx <> "") and (toBonusCouponIdx <> "0") then
					'���ʽ�����
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
			contents_jupsu	= contents_jupsu & "[" & html2db(itemname) & "] " & html2db(itemoptionname) & " " & CStr(itemnocancel) & "�� �߰�" & vbCrLf
		end if
	next

	Call EditOrderMasterRefundInfo(orderserial, refundcouponsum, allatsubtractsum, refundmileagesum, refunddepositsum, refundgiftcardsum)
	CSOrderRecalculateOrder orderserial,false

	if (CS_ORDER_FUNCTION_RESULT <> "") then
	    response.write "<script language='javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	'==========================================================================
	' CS ����Ÿ AS ����
	''html2db ������� ����.
	Call EditCSMaster(iAsID, reguserid, title, contents_jupsu, gubun01, gubun02)

	'' CS Detail(���û�ǰ���) ���
	'�ɼǺ��濡���� ������ ����
	Call AddCSDetailByArrStr(detailitemlist, iAsID, orderserial)

	' CS ����Ÿ AS�Ϸ�
	Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

	' ���뺯��
	Call AddCustomerOpenContents(iAsid, html2db(contents_finish))

	response.write "<script>" & vbCrLf
	response.write "	alert('���� �Ǿ����ϴ�.');" & vbCrLf
	response.write "	opener.parent.location.reload();" & vbCrLf

	if (newAsId <> 0) then
		response.write "	var a = window.open('/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + CStr(newAsId) + "&mode=editreginfo','pop_cs_action_reg_','width=1200 height=600 scrollbars=yes resizable=yes');" & vbCrLf
	else
		response.write "	window.blur();" & vbCrLf
	end if

	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	response.end

elseif (mode="orderChange") then

	regUserID	= session("ssBctID")

	fromItemId     	= request("fromItemId")
	fromItemOption  = request("fromItemOption")
	toItemId     	= request("toItemId")
	toItemOption    = request("toItemOption")
	itemnocancel    = request("itemnocancel")
	itemnoadd    	= request("itemnoadd")

	SalePrice  			= request("toSalePrice")
	ItemCouponPrice  	= request("toItemCouponPrice")
	BonusCouponPrice  	= request("toBonusCouponPrice")
	EtcDiscountPrice  	= request("toEtcDiscountPrice")
	buycash  			= request("toAddBuycash")

	if (copysaleinfo = "Y") then
		itemcouponidx = request("fromItemCouponIdx")
		''bonuscouponidx = request("fromBonusCouponIdx")
	end If

	refundrequire		= request("refundrequire")
	canceltotal			= request("canceltotal")
	refunditemcostsum	= request("refunditemcostsum")
	refundcouponsum		= request("refundcouponsum")
	allatsubtractsum	= request("allatsubtractsum")

	arrItemInfo = GetItemOptionNames(fromItemId, fromItemOption, toItemId, toItemOption)

	For i = 0 To UBound(arrItemInfo,2)
		if CStr(arrItemInfo(0, i)) = fromItemId and arrItemInfo(1, i) = fromItemOption then
			fromItemName = arrItemInfo(2, i)
			fromItemOptionName = arrItemInfo(3, i)
		end if

		if CStr(arrItemInfo(0, i)) = toItemId and arrItemInfo(1, i) = toItemOption then
			toItemName = arrItemInfo(2, i)
			toItemOptionName = arrItemInfo(3, i)
		end if
	next


	'// ���� ��ǰ
	newdetailitemlist = newdetailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & itemnoadd & Chr(9) & Trim(toItemId) & Chr(9) & toItemOption & Chr(9)
	if (toItemOption = "0000") then
		contents_itemlist	= contents_itemlist & "[" & toItemId & "-" & toItemOption & "] " & toItemName & " " & itemnoadd & "�� ���" & vbCrLf
	else
		contents_itemlist	= contents_itemlist & "[" & toItemId & "-" & toItemOption & "] " & toItemName & "[" & toItemOptionName & "] " & itemnoadd & "�� ���" & vbCrLf
	end if

	'// ȸ���� ��ǰ
    detailitemlist = detailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(itemnocancel) & Chr(9) & Trim(fromItemId) & Chr(9) & fromItemOption & Chr(9)
	if (fromItemOption = "0000") then
		contents_itemlist	= contents_itemlist & "[" & fromItemId & "-" & fromItemOption & "] " & fromItemName & " " & itemnocancel & "�� ȸ��" & vbCrLf
	else
		contents_itemlist	= contents_itemlist & "[" & fromItemId & "-" & fromItemOption & "] " & fromItemName & "[" & fromItemOptionName & "] " & itemnocancel & "�� ȸ��" & vbCrLf
	end If


	If (refundrequire < 0) Then
		Response.Write "����!!"
		Response.end
	End If

	''Response.Write itemnoadd & "<br />"
	''Response.Write itemnocancel & "<br />"
	''Response.end

	'// �������뿡 �߰�
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu

	divCd 			= "A100"	' ��ȯ���(��ǰ����)

	' CS ����Ÿ AS ���
	''html2db ������� ����.
	iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

	'//  CS Detail(���û�ǰ���) ���
	'// �±�ȯ����� ���Ǵ� ��ǰ�� ����Ѵ�.
	Call AddCSDetailWithoutOrderDetailByArrStr(newdetailitemlist, iAsID, orderserial)

	If (refundrequire <> 0) Or (itemnoadd <> itemnocancel) Then
		'���ϻ�ǰ �ƴ�
		Call ModiCSDetailAddedItem(iAsID, toItemId, toItemOption, SalePrice, ItemCouponPrice, BonusCouponPrice, buycash)

		Call RegCSMasterRefundInfo(iAsID, "R007", refundrequire , refunditemcostsum, refunditemcostsum, 0 , 0, refundcouponsum, allatsubtractsum, canceltotal, refunditemcostsum, 0, refundcouponsum, allatsubtractsum, 0, 0, 0  , "", "", "", "")
	End If

	'// CS �±�ȯ���(���ϻ�ǰ, ��ǰ���� - A000, A100) ������ ���Ǵ� ��ǰ ��������
	Call ApplyLimitItemByCS(iAsID)

    if (isupchebeasong = "Y") then
        '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        call RegCSMasterAddUpche(iAsID, requiremakerid)

		'// �� �߰���ۺ�(��ȯ��� ���)
		Call SetCustomerAddBeasongPay(iAsID, "", 0, "N", 0)

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
		end If

		response.write "<script>alert('��ȯ���(��ǰ����) �����Ϸ� - ��ü���');</script>"

    else
        '�ٹ����� ����� ��� ��ǰ���� �±�ȯ ȸ�� ����
        newasid = RegCSMaster("A111", orderserial, reguserid, "��ȯȸ��(��ǰ����)", contents_jupsu, gubun01, gubun02)

		''�±�ȯȸ������ ȸ���Ǵ� ��ǰ�� ����Ѵ�.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// �� �߰���ۺ�(��ȯȸ���� ���)
		Call SetCustomerAddBeasongPay(newasid, "", 0, "N", 0)

		'// asid ����
		Call SetRefAsid(newasid, iAsID)

        response.write "<script>alert('��ȯ���(��ǰ����) ���� �� ȸ�������Ϸ� - �ٹ����� ���');</script>"
    end if

	response.write "<script>opener.parent.location.reload();</script>"
	response.write "<script>window.resizeTo(1200,600)</script>"

	if (requiremakerid<>"") then
		response.write "<script>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(iAsID) + "&mode=editreginfo'</script>"
	else
		'// �ٹ�� �±�ȯ ȸ��â���� �̵�
		response.write "<script>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(newasid) + "&mode=editreginfo'</script>"
	end if

	dbget.close()
	response.end

elseif (mode="modistate2") then

	'// �ֹ��뺸�� ��ȯ
	SetDetailCurrState(orderserial)

	call AddCsMemo(orderserial,"1", "", session("ssBctId"), "���� ��û���� �ֹ��뺸 ��ȯ")

	response.write "<script>alert('�ֹ��뺸 ��ȯ�Ǿ����ϴ�.'); history.back();</script>"

	dbget.close()
	response.end

elseif (mode="chgReducedPrice") then

    sqlStr = " select t.OutMallOrderSeq "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " [db_temp].[dbo].[tbl_xSite_TMPOrder] t "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and t.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr + " 	and t.matchItemID = '" & itemid & "' "
    sqlStr = sqlStr + " 	and t.matchitemoption = '" & itemoption & "' "
    sqlStr = sqlStr + " 	and t.ItemOrderCount > 0 "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	''rw sqlStr

    '// �������� ���� �� �ִ�.
    OutMallOrderSeq = ""
	if not rsget.Eof then
		do until rsget.eof
            if (OutMallOrderSeq = "") then
                OutMallOrderSeq = rsget("OutMallOrderSeq")
            else
                OutMallOrderSeq = OutMallOrderSeq + "," + rsget("OutMallOrderSeq")
            end if

			rsget.movenext
		loop
	end if
	rsget.Close

    if OutMallOrderSeq <> "" then
        OutMallOrderSeq = Split(OutMallOrderSeq, ",")

        for i = 0 to UBound(OutMallOrderSeq)
            if Trim(OutMallOrderSeq(i)) <> "" then
                sqlStr = " EXEC  [db_jungsan].[dbo].[usp_Ten_OUTAMLL_XSiteOrderTmp_ChangVal] 'chgrealsellprice','" & Trim(OutMallOrderSeq(i)) & "','" & reducedPrice & "'"
                rw sqlStr
	            ''dbget.Execute sqlStr
            end if
        next
    end if

    ''response.write "<script>alert('����Ǿ����ϴ�.'); history.back();</script>"

	dbget.close() : response.end

end if


response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
response.write "<script>location.replace('/cscenter/ordermaster/orderdetail_editoption.asp?idx=" + detailidx + "');</script>"



%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
