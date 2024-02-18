<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'###########################################################

dim IsStatusRegister			'접수
dim IsStatusEdit				'수정
dim IsStatusFinishing			'처리완료 시도
dim IsStatusFinished			'처리완료

dim IsDisplayPreviousCSList		'이전 CS 내역
dim IsDisplayCSMaster			'CS 마스터정보
dim IsDisplayItemList			'상품목록
dim IsDisplayChangeItemList		'다른상품 맞교환출고 상품목록
dim IsDisplayRefundInfo			'환불정보
dim IsDisplayButton				'버튼

dim IsPossibleModifyCSMaster
dim IsPossibleModifyItemList
dim IsPossibleModifyRefundInfo

dim ARR_ERROR_MSG()
dim MAX_ERROR_MSG_COUNT

MAX_ERROR_MSG_COUNT = 10

ReDim Preserve ARR_ERROR_MSG(MAX_ERROR_MSG_COUNT)

dim ERROR_MSG_TRY_MODIFY
dim itemCouponRefundYN
	itemCouponRefundYN="Y"	' 상품쿠폰환급여부

'변수 설정
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
    	'CS 접수
    	IsStatusRegister 	= true

    elseif (mode = "editreginfo") then
    	'----------------------------------------------------------------------
    	'CS 수정
    	IsStatusEdit 		= true

    elseif (mode = "finishreginfo") then
    	'----------------------------------------------------------------------
    	'완료시도
    	IsStatusFinishing 	= true

		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

		ERROR_MSG_TRY_MODIFY = "CS 완료처리 단계에서는 처리내용입력 외 수정할 수 없습니다. CS 정보수정을 이용하세요."

    elseif (mode = "finished") then
    	'----------------------------------------------------------------------
    	'완료된 내역
    	IsStatusFinished 	= true

		IsPossibleModifyCSMaster	= false
		IsPossibleModifyItemList	= false
		IsPossibleModifyRefundInfo	= false

    	IsDisplayButton 	= false

    	ERROR_MSG_TRY_MODIFY = "완료된 내역은 수정할 수 없습니다."
    else
    	'ERROR
    end if

end function

'변수 설정
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


''CsAction 접수시 상품별 체크 가능여부

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
				'맞교환회수(텐바이텐배송)
				'맞교환
				'반품접수(업체배송)
				'회수신청(텐바이텐배송)
				'상품변경 맞교환출고
				'상품변경 맞교환회수(텐배)
				IsPossibleCheckItem = true
			end if
		end if
	else
		'기타
		IsPossibleCheckItem = true

		if (CStr(itemdetailstate) < "7") then
			if (divcd = "A001") then
				'// 누략 재발송
				IsPossibleCheckItem = false
			end if
		end if
	end if

end function

public function IsCSCancelProcess(divcd)

	'주문취소
	if (divcd = "A008") then
		IsCSCancelProcess = true
	else
		IsCSCancelProcess = false
	end if

end function

public function IsCSReturnProcess(divcd)

	'반품접수(업체배송), 회수신청(텐바이텐배송)
	if ((divcd = "A004") or (divcd = "A010")) then
		IsCSReturnProcess = true
	else
		IsCSReturnProcess = false
	end if

end function

public function IsCSExchangeProcess(divcd)

	'맞교환출고, 맞교환회수(텐바이텐배송), 맞교환반품(업체배송), 상품변경 맞교환회수(텐바이텐배송), 상품변경 맞교환반품(업체배송)
	if ((divcd = "A000") or (divcd = "A011") or (divcd = "A012") or (divcd = "A111") or (divcd = "A112")) then
		IsCSExchangeProcess = true
	else
		IsCSExchangeProcess = false
	end if

end function

public function IsCSServiceProcess(divcd)

	'누락발송, 서비스발송  프로세스
	if ((divcd = "A000") or (divcd = "A001") or (divcd = "A002")) then
		IsCSServiceProcess = true
	else
		IsCSServiceProcess = false
	end if

end function

public function IsCSCancelInfoNeeded(divcd)

	'주문취소, 반품접수(업체배송), 회수신청(텐바이텐배송)
	'// 주문내역변경은 차액이 발생해도 자동으로 환불 CS 를 생성하기에 따로 취소정보를 표시할 필요가 없다.
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

	'주문취소, 반품접수(업체배송), 회수신청(텐바이텐배송), 환불, 외부몰환불요청, 카드/이체/휴대폰취소요청
	'// 주문내역변경은 차액이 발생해도 자동으로 환불 CS 를 생성하기에 따로 환불정보를 표시할 필요가 없다.
	if ((divcd = "A008") or (divcd = "A004") or (divcd = "A010") or (divcd = "A003") or (divcd = "A005") or (divcd = "A007") or (divcd = "A100")) then
		IsCSRefundNeeded = true
	else
		IsCSRefundNeeded = false
	end if

end function

public function IsCSUpcheJungsanNeeded(divcd)

	'반품접수(업체배송), 맞교환출고, 업체기타정산, 상품변경 맞교환출고, 누락재발송, 서비스발송, 기타회수, 고객추가결제
	if ((divcd = "A004") or (divcd = "A000") or (divcd = "A700") or (divcd = "A100") or (divcd = "A001") or (divcd = "A002") or (divcd = "A200") or (divcd = "A999")) then
		IsCSUpcheJungsanNeeded = true
	else
		IsCSUpcheJungsanNeeded = false
	end if

end function

'// 맞교환 회수상태
public function IsCSItemExchangeReceiveInfoNeeded(divcd)

	'맞교환출고, 상품변경 맞교환출고 = 업배만
	if (divcd = "A000") or (divcd = "A100") then
		IsCSItemExchangeReceiveInfoNeeded = true
	else
		IsCSItemExchangeReceiveInfoNeeded = false
	end if

end function

'// 고객추가배송비(상품변경 맞교환)
public function IsCSItemExchangeCustomerBeasongPayNeeded(divcd)

	' 상품변경 맞교환회수(텐바이텐), 상품변경 맞교환출고(업체배송)
	if (divcd = "A111") or (divcd = "A100") then
		IsCSItemExchangeCustomerBeasongPayNeeded = true
	else
		IsCSItemExchangeCustomerBeasongPayNeeded = false
	end if

end function

public function IsCSItemListNeeded(divcd)

	'환불, 카드,이체,휴대폰취소요청, 외부몰환불요청, 상품변경 맞교환출고, 상품변경 맞교환회수(텐배), 상품변경 맞교환반품(업배)
	if (divcd <> "A003") and (divcd <> "A007") and (divcd <> "A005") and (divcd <> "A100") and (divcd <> "A111") and (divcd <> "A112") then
		IsCSItemListNeeded = true
	else
		IsCSItemListNeeded = false
	end if

end function

public function IsCSChangeItemListNeeded(divcd)

	'상품변경 맞교환출고, 상품변경 맞교환회수(텐배), 상품변경 맞교환반품(업배)
	if (divcd = "A100") or (divcd = "A111") or (divcd = "A112") then
		IsCSChangeItemListNeeded = true
	else
		IsCSChangeItemListNeeded = false
	end if

end function

%>
