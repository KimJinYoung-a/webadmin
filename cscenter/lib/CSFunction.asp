<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'###########################################################

dim IsStatusRegister			'����
dim IsStatusEdit				'����
dim IsStatusFinishing			'ó���Ϸ� �õ�
dim IsStatusFinished			'ó���Ϸ�

dim IsDisplayPreviousCSList		'���� CS ����
dim IsDisplayCSMaster			'CS ����������
dim IsDisplayItemList			'��ǰ���
dim IsDisplayChangeItemList		'�ٸ���ǰ �±�ȯ��� ��ǰ���
dim IsDisplayRefundInfo			'ȯ������
dim IsDisplayButton				'��ư

dim IsPossibleModifyCSMaster
dim IsPossibleModifyItemList
dim IsPossibleModifyRefundInfo

dim ARR_ERROR_MSG()
dim MAX_ERROR_MSG_COUNT

MAX_ERROR_MSG_COUNT = 10

ReDim Preserve ARR_ERROR_MSG(MAX_ERROR_MSG_COUNT)

dim ERROR_MSG_TRY_MODIFY
dim itemCouponRefundYN
	itemCouponRefundYN="Y"	' ��ǰ����ȯ�޿���

'���� ����
function SetCSVariable(mode, divcd)

	IsStatusRegister 			= false
	IsStatusEdit 				= false
	IsStatusFinishing 			= false
	IsStatusFinished 			= false

	IsDisplayPreviousCSList 	= true
	IsDisplayCSMaster 			= true
	IsDisplayItemList 			= true
	IsDisplayChangeItemList		= true
	IsDisplayRefundInfo 		= true
	IsDisplayButton 			= true

	IsPossibleModifyCSMaster	= true
	IsPossibleModifyItemList	= true
	IsPossibleModifyRefundInfo	= true

	IsDisplayItemList = IsCSItemListNeeded(divcd)
	IsDisplayChangeItemList = IsCSChangeItemListNeeded(divcd)

    if (mode = "regcsas") then
    	'----------------------------------------------------------------------
    	'CS ����
    	IsStatusRegister 	= true

    elseif (mode = "editreginfo") then
    	'----------------------------------------------------------------------
    	'CS ����
    	IsStatusEdit 		= true

    elseif (mode = "finishreginfo") then
    	'----------------------------------------------------------------------
    	'�Ϸ�õ�
    	IsStatusFinishing 	= true

		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

		ERROR_MSG_TRY_MODIFY = "CS �Ϸ�ó�� �ܰ迡���� ó�������Է� �� ������ �� �����ϴ�. CS ���������� �̿��ϼ���."

    elseif (mode = "finished") then
    	'----------------------------------------------------------------------
    	'�Ϸ�� ����
    	IsStatusFinished 	= true

		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

    	IsDisplayButton 	= false

    	ERROR_MSG_TRY_MODIFY = "�Ϸ�� ������ ������ �� �����ϴ�."
    else
    	'ERROR
    end if

end function

'���� ����
function SetCSErrorMessage(msg)

	dim i

	ARR_ERROR_MSG(MAX_ERROR_MSG_COUNT)

	for i = 0 to MAX_ERROR_MSG_COUNT - 1
		if (ARR_ERROR_MSG(i) = "") then
			ARR_ERROR_MSG(i) = msg
			exit for
		end if
	next

end function


''CsAction ������ ��ǰ�� üũ ���ɿ���

'masterstate, mastercancelyn, divcd, itemdetailstate

public function IsPossibleCheckItem(divcd, ismastercanceled, isdetailcanceled, masterstate, itemdetailstate, isupchebeasong)

	IsPossibleCheckItem = false

	if (ismastercanceled) then
		exit function
	end if

	if (isdetailcanceled) then
		exit function
	end if

	if (IsCSCancelProcess(divcd)) then
		IsPossibleCheckItem = true
		if (CStr(itemdetailstate) >= "7") then
			IsPossibleCheckItem = false
		end if

	elseif (IsCSReturnProcess(divcd) = true) or (IsCSExchangeProcess(divcd) = True) then
		IsPossibleCheckItem = false
		if (CStr(itemdetailstate) >= "7") then
			if _
				((divcd = "A011") and (Not isupchebeasong)) _
				or _
				(divcd = "A000") _
				or _
				(divcd = "A004") _
				or _
				(divcd = "A010") _
				or _
				(divcd = "A100") _
				or _
				((divcd = "A111") and (Not isupchebeasong)) _
			then
				'�±�ȯȸ��(�ٹ����ٹ��)
				'�±�ȯ
				'��ǰ����(��ü���)
				'ȸ����û(�ٹ����ٹ��)
				'��ǰ���� �±�ȯ���
				'��ǰ���� �±�ȯȸ��(�ٹ�)
				IsPossibleCheckItem = true
			end if
		end if
	else
		'��Ÿ
		IsPossibleCheckItem = true

		if (CStr(itemdetailstate) < "7") then
			if (divcd = "A001") then
				'// ���� ��߼�
				IsPossibleCheckItem = false
			end if
		end if
	end if

end function

public function IsCSCancelProcess(divcd)

	'�ֹ����
	if (divcd = "A008") then
		IsCSCancelProcess = true
	else
		IsCSCancelProcess = false
	end if

end function

public function IsCSReturnProcess(divcd)

	'��ǰ����(��ü���), ȸ����û(�ٹ����ٹ��)
	if ((divcd = "A004") or (divcd = "A010")) then
		IsCSReturnProcess = true
	else
		IsCSReturnProcess = false
	end if

end function

public function IsCSExchangeProcess(divcd)

	'�±�ȯ���, �±�ȯȸ��(�ٹ����ٹ��), �±�ȯ��ǰ(��ü���), ��ǰ���� �±�ȯȸ��(�ٹ����ٹ��), ��ǰ���� �±�ȯ��ǰ(��ü���)
	if ((divcd = "A000") or (divcd = "A011") or (divcd = "A012") or (divcd = "A111") or (divcd = "A112")) then
		IsCSExchangeProcess = true
	else
		IsCSExchangeProcess = false
	end if

end function

public function IsCSServiceProcess(divcd)

	'�����߼�, ���񽺹߼�  ���μ���
	if ((divcd = "A000") or (divcd = "A001") or (divcd = "A002")) then
		IsCSServiceProcess = true
	else
		IsCSServiceProcess = false
	end if

end function

public function IsCSCancelInfoNeeded(divcd)

	'�ֹ����, ��ǰ����(��ü���), ȸ����û(�ٹ����ٹ��)
	'// �ֹ����������� ������ �߻��ص� �ڵ����� ȯ�� CS �� �����ϱ⿡ ���� ��������� ǥ���� �ʿ䰡 ����.
	if ((divcd = "A008") or (divcd = "A004") or (divcd = "A010")) then
		IsCSCancelInfoNeeded = true
	else
		IsCSCancelInfoNeeded = false
	end if

end function

public function IsCSRefundNeeded(divcd, masterstate)

	if (CStr(masterstate) < "4") then
		IsCSRefundNeeded = false
		exit function
	end if

	'�ֹ����, ��ǰ����(��ü���), ȸ����û(�ٹ����ٹ��), ȯ��, �ܺθ�ȯ�ҿ�û, ī��/��ü/�޴�����ҿ�û
	'// �ֹ����������� ������ �߻��ص� �ڵ����� ȯ�� CS �� �����ϱ⿡ ���� ȯ�������� ǥ���� �ʿ䰡 ����.
	if ((divcd = "A008") or (divcd = "A004") or (divcd = "A010") or (divcd = "A003") or (divcd = "A005") or (divcd = "A007") or (divcd = "A100")) then
		IsCSRefundNeeded = true
	else
		IsCSRefundNeeded = false
	end if

end function

public function IsCSUpcheJungsanNeeded(divcd)

	'��ǰ����(��ü���), �±�ȯ���, ��ü��Ÿ����, ��ǰ���� �±�ȯ���, ������߼�, ���񽺹߼�, ��Ÿȸ��, ���߰�����
	if ((divcd = "A004") or (divcd = "A000") or (divcd = "A700") or (divcd = "A100") or (divcd = "A001") or (divcd = "A002") or (divcd = "A200") or (divcd = "A999")) then
		IsCSUpcheJungsanNeeded = true
	else
		IsCSUpcheJungsanNeeded = false
	end if

end function

'// �±�ȯ ȸ������
public function IsCSItemExchangeReceiveInfoNeeded(divcd)

	'�±�ȯ���, ��ǰ���� �±�ȯ��� = ���踸
	if (divcd = "A000") or (divcd = "A100") then
		IsCSItemExchangeReceiveInfoNeeded = true
	else
		IsCSItemExchangeReceiveInfoNeeded = false
	end if

end function

'// ���߰���ۺ�(��ǰ���� �±�ȯ)
public function IsCSItemExchangeCustomerBeasongPayNeeded(divcd)

	' ��ǰ���� �±�ȯȸ��(�ٹ�����), ��ǰ���� �±�ȯ���(��ü���)
	if (divcd = "A111") or (divcd = "A100") then
		IsCSItemExchangeCustomerBeasongPayNeeded = true
	else
		IsCSItemExchangeCustomerBeasongPayNeeded = false
	end if

end function

public function IsCSItemListNeeded(divcd)

	'ȯ��, ī��,��ü,�޴�����ҿ�û, �ܺθ�ȯ�ҿ�û, ��ǰ���� �±�ȯ���, ��ǰ���� �±�ȯȸ��(�ٹ�), ��ǰ���� �±�ȯ��ǰ(����)
	if (divcd <> "A003") and (divcd <> "A007") and (divcd <> "A005") and (divcd <> "A100") and (divcd <> "A111") and (divcd <> "A112") then
		IsCSItemListNeeded = true
	else
		IsCSItemListNeeded = false
	end if

end function

public function IsCSChangeItemListNeeded(divcd)

	'��ǰ���� �±�ȯ���, ��ǰ���� �±�ȯȸ��(�ٹ�), ��ǰ���� �±�ȯ��ǰ(����)
	if (divcd = "A100") or (divcd = "A111") or (divcd = "A112") then
		IsCSChangeItemListNeeded = true
	else
		IsCSChangeItemListNeeded = false
	end if

end function

%>
