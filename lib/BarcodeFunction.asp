<%

''/lib/BarcodeFunction.asp
'' SCM 과 로직스 모두 동일한 내용이어야 한다.

'// ============================================================================
'// 텐바이텐 바코드 형식인지
'// function BF_IsMaybeTenBarcode(barcode)

'// itemgubun 구하기
'// function BF_GetItemGubun(barcode)

'// itemid 구하기
'// function BF_GetItemId(barcode)

'// itemoption 구하기
'// function BF_GetItemOption(barcode)

'// 텐바이텐 바코드 구하기
'// function BF_MakeTenBarcode(itemgubun, itemid, itemoption)

'// 형식을 갖춘 itemid 구하기(011111, 01000000)
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

	'// 12 또는 14 자리 숫자로 구성된 바코드는 텐바이텐 바코드와 쫑날 수 있다.
	'// 따라서 등록불가
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

    ''2014/01/06추가===================================
    if Not (Len(barcode)=12 or Len(barcode)=14) then
        exit function
    end if

    dim iitemgubun
    iitemgubun = left(barcode,2)

    if Not (iitemgubun="10" or iitemgubun="70" or iitemgubun="80" or iitemgubun="55" or iitemgubun="85" or iitemgubun="90")  then
        exit function
    end if
    ''2014/01/06추가===================================

	'// 12자리 또는 14자리 바코드만 유효하다.
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
