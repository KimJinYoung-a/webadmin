<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs���� �ɼǱ�ȯ
' History : �̻� ����
'			2023.09.05 �ѿ�� ����(6�������� �ֹ��� ��ȯ �����ϰ� ó��)
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

''���� �ֹ� �������� Check
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
	'// �ֹ�����(��ǰ�ɼǺ���)
	jungsanExists = false
	'// ��ҵǴ� ��ǰ
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
	    response.write "<script type='text/javascript'>alert('���� : " & "���(ȸ��) �Ǵ� ��ǰ�� ���� ������ �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
	    dbget.close()	:	response.End
	end if

	csCancelExists = False

	'// ��� CS
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
	''    response.write "<script type='text/javascript'>alert('���� : " & "���CS�� �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
	''    dbget.close()	:	response.End
	''end if

	'// �߰��Ǵ� ��ǰ
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
			    response.write "<script type='text/javascript'>alert('���� : " & "�߰��Ǵ� ��ǰ�� ���� ������ �����մϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if

			if (itemChulgoFinished) then
			    response.write "<script type='text/javascript'>alert('���� : " & "�߰��Ǵ� ��ǰ�� ���Ϸ��Դϴ�. ������ �� �����ϴ�." & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if
		end if
	next

	contents_finish = "��ǰ�ɼǺ����� ���������� ó���Ǿ����ϴ�."

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

			'// �ɼǺ���
			result = CSOrderModifyItemOption(orderserial, itemId, orgitemoptioncode, itemoptioncode, itemoptionno)

			if (CS_ORDER_FUNCTION_RESULT <> "") then
			    response.write "<script type='text/javascript'>alert('���� : " & CS_ORDER_FUNCTION_RESULT & "');history.back();</script>"
			    dbget.close()	:	response.End
			end if

			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			result = CSOrderGetItemState(orderserial, itemId, itemoptioncode)
			orderdetailidx = CS_ORDER_ITEM_ORDERDETAILIDX

			ResetGlobalVarible()
			'--------------------------------------------------------------------------

	        detailitemlist = detailitemlist & "|" & orderdetailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & itemoptionno & Chr(9)
			contents_itemlist	= contents_itemlist & "[" & itemoptioncode & "] " & itemoptionname & " " & itemoptionno & "�� �߰�" & vbCrLf
		end if

	next

    detailitemlist = detailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(-1*totalchangeno) & Chr(9)
	contents_itemlist	= contents_itemlist & "[" & orgitemoptioncode & "] " & orgitemoptionname & " " & CStr(totalchangeno) & "�� ���" & vbCrLf

	'// �������뿡 �߰�
	contents_jupsu = contents_itemlist & vbCrLf & vbCrLf & contents_jupsu

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

    '// ���޸� �ֹ��� ����� ��ǰ�ڵ� �Է�, 2021-03-22, skyer9
    strSql = " exec [db_cs].[dbo].[usp_Ten_CsAs_ChangeItem2ExtSite] " & iAsID
    dbget.Execute strSql

    '''==========================================================================================

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

	response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
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
	''��ǰ���� �±�ȯ

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

			'// ���� ��ǰ(�Ѱ��� �̻��� �� �ִ�.)
	        newdetailitemlist = newdetailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & itemoptionno & Chr(9) & Trim(itemId) & Chr(9) & itemoptioncode & Chr(9)
			contents_itemlist	= contents_itemlist & "[" & itemoptioncode & "] " & itemoptionname & " " & itemoptionno & "�� ���" & vbCrLf

		end if
	next

	'// ȸ���� ��ǰ(�Ѱ����� ����)
    detailitemlist = detailitemlist & "|" & detailidx & Chr(9) & gubun01 & Chr(9) & gubun02 & Chr(9) & CStr(totalchangeno) & Chr(9)
	contents_itemlist	= contents_itemlist & "[" & orgitemoptioncode & "] " & orgitemoptionname & " " & CStr(totalchangeno) & "�� ȸ��" & vbCrLf

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
        newasid = RegCSMaster("A112", orderserial, reguserid, "��ȯȸ��(�ɼǺ���,��ü���)", contents_jupsu, gubun01, gubun02)

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

		response.write "<script type='text/javascript'>alert('�ɼǺ��� �±�ȯ �����Ϸ� - ��ü���');</script>"
    else
        '�ٹ����� ����� ��� ��ǰ���� �±�ȯ ȸ�� ����
        newasid = RegCSMaster("A111", orderserial, reguserid, "��ȯȸ��(�ɼǺ���)", contents_jupsu, gubun01, gubun02)

		''�±�ȯȸ������ ȸ���Ǵ� ��ǰ�� ����Ѵ�.
        Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

		'// �� �߰���ۺ�(��ȯȸ���� ���)
		Call SetCustomerAddBeasongPay(newasid, add_customeraddmethod, add_customeraddbeasongpay, "N", 0)

		'// asid ����
		Call SetRefAsid(newasid, iAsID)

        response.write "<script type='text/javascript'>alert('�ɼǺ��� �±�ȯ ��� ���� �� ȸ�������Ϸ� - �ٹ����� ���');</script>"
    end if

	response.write "<script type='text/javascript'>opener.parent.location.reload();</script>"
	'response.write "<script type='text/javascript'>window.resizeTo(1400,800)</script>"

	if (requiremakerid<>"") then
		response.write "<script type='text/javascript'>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(iAsID) + "&mode=editreginfo'</script>"
	else
		'// �ٹ�� �±�ȯ ȸ��â���� �̵�
		response.write "<script type='text/javascript'>location.href = '/cscenter/action/pop_cs_action_new.asp?id=" + CStr(newasid) + "&mode=editreginfo'</script>"
	end if
	dbget.close() : response.end
else
	response.write "���ǵ��� �ʾҽ��ϴ�."
	dbget.close() : response.end
end if

response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
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
