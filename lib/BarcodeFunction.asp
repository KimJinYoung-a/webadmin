<%

''/lib/BarcodeFunction.asp
'' SCM �� ������ ��� ������ �����̾�� �Ѵ�.

'// ============================================================================
'// �ٹ����� ���ڵ� ��������
'// function BF_IsMaybeTenBarcode(barcode)

'// itemgubun ���ϱ�
'// function BF_GetItemGubun(barcode)

'// itemid ���ϱ�
'// function BF_GetItemId(barcode)

'// itemoption ���ϱ�
'// function BF_GetItemOption(barcode)

'// �ٹ����� ���ڵ� ���ϱ�
'// function BF_MakeTenBarcode(itemgubun, itemid, itemoption)

'// ������ ���� itemid ���ϱ�(011111, 01000000)
'// function BF_GetFormattedItemId(itemid)

'// ============================================================================
function BF_GetFormattedItemId(itemid)
	dim tmpStr

	if isnull(itemid) or trim(itemid)="" then exit function
	itemid=trim(itemid)

	if Len(itemid) < 7 then
		tmpStr = Right(CStr(1000000 + itemid), 6)
	else
		tmpStr = Right(CStr(100000000 + itemid), 8)
	end if

	BF_GetFormattedItemid = tmpStr
end function

'// ============================================================================
function BF_MakeTenBarcode(itemgubun, itemid, itemoption)
	BF_MakeTenBarcode = itemgubun & BF_GetFormattedItemId(itemid) & itemoption

end function

'// ============================================================================
function BF_IsMaybeTenBarcode(barcode)
	dim itemid

	BF_IsMaybeTenBarcode = False

	if IsNull(barcode) then
		exit function
	end if

	itemid = BF_GetItemId(barcode)
	if (itemid = "") then
		exit function
	end if

	if IsNumeric(itemid) then
		BF_IsMaybeTenBarcode = True
	end if

end function


'// ============================================================================
function BF_IsAvailPublicBarcode(barcode)
	dim itemid

	BF_IsAvailPublicBarcode = False

	'// 12 �Ǵ� 14 �ڸ� ���ڷ� ������ ���ڵ�� �ٹ����� ���ڵ�� �г� �� �ִ�.
	'// ���� ��ϺҰ�
	itemid = BF_GetItemId(barcode)
	if (itemid = "") then
		BF_IsAvailPublicBarcode = True
	end if

end function


'// ============================================================================
function BF_GetItemGubun(barcode)
	BF_GetItemGubun = ""

	if IsNull(barcode) then
		exit function
	end if

	barcode = Trim(barcode)

	BF_GetItemGubun = Left(barcode, 2)

end function


'// ============================================================================
function BF_GetItemId(barcode)
	dim tmpitemid

	BF_GetItemId = ""

	if IsNull(barcode) or barcode="" then
		exit function
	end if

	barcode = Trim(barcode)

	'if Not IsNumeric(barcode) then
	'	exit function
	'end if

    ''2014/01/06�߰�===================================
    if Not (Len(barcode)=12 or Len(barcode)=14) then
        exit function
    end if

    dim iitemgubun
    iitemgubun = left(barcode,2)

    if Not (iitemgubun="10" or iitemgubun="70" or iitemgubun="80" or iitemgubun="55" or iitemgubun="85" or iitemgubun="90")  then
        exit function
    end if
    ''2014/01/06�߰�===================================

	'// 12�ڸ� �Ǵ� 14�ڸ� ���ڵ常 ��ȿ�ϴ�.
	select case Len(barcode)
		case 12
			tmpitemid = mid(barcode,3,6)
		''case 13
		''	tmpitemid = mid(barcode,3,7)
		case 14
			tmpitemid = mid(barcode,3,8)
	end select

	if isNumeric(tmpitemid) then
		BF_GetItemId= CStr(CLng(tmpitemid))
	else
		BF_GetItemId=""
	end if
end function


'// ============================================================================
function BF_GetItemOption(barcode)
	BF_GetItemOption = ""

	if IsNull(barcode) then
		exit function
	end if

	barcode = Trim(barcode)

	BF_GetItemOption = Right(barcode, 4)

end function

%>
