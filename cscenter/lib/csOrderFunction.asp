<%

'배송비 계산 프로세스를 함수단위로 한번만 계산하도록 변경 필요

'==============================================================================
'주문상세 부분취소 : 기존에 있던 상품 한가지 취소
'==============================================================================
'function CSOrderCancelItem(byval orderserial, byval itemid, byval itemoption)

'==============================================================================
'주문상세 부분취소 정상화 : 기존에 있던 취소된 상품 한가지를 정상화
'==============================================================================
'function CSOrderRestoreCanceledItem(byval orderserial, byval itemid, byval itemoption)

'==============================================================================
'신규 상품추가 : 기존에 없는 상품 주문 디테일에 추가.
'==============================================================================
'function CSOrderAddNewItem(byval orderserial, byval itemid, byval itemoption, byval itemno)

'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션중 일부 변경
'==============================================================================
'function CSOrderModifyItemOption(byval orderserial, byval itemid, byval itemoptionfrom, itemoptionto, byval itemno)

'==============================================================================
'기존 상품수량변경 : 기존에 있던 상품 수량 변경
'==============================================================================
'function CSOrderModifyItemNo(byval orderserial, byval itemid, byval itemoption, byval itemno)

'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션 변경
'상품정보 복사는 따로 해주어야 한다.
'==============================================================================
'function CSOrderChangeItemForce(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto, byval itemno)

'==============================================================================
'기존 상품변경 : 기존에 있던 상품중 일부 변경
'==============================================================================
'function CSOrderChangeItem(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto, byval itemno)


'체크 : 상품쿠폰 적용상품 교환불가 체크
'체크 : 플러스세일 적용상품 교환불가 체크

'체크 : 비율쿠폰 교환가능 체크

'TODO : 결재상태별 체크...



dim CS_ORDER_FUNCTION_RESULT

dim CS_ORDER_ITEM_ORDERDETAILIDX
dim CS_ORDER_ITEM_CANCELYN
dim CS_ORDER_ITEM_CURRSTATE
dim CS_ORDER_ITEM_ISUPCHEBEASONG
dim CS_ORDER_ITEM_NO
dim CS_ORDER_ITEM_ITEMCOST
dim CS_ORDER_ITEM_MAKERID

dim CS_ORDER_ITEM_SELLCASH
dim CS_ORDER_ITEM_OPTADDPRICE

dim CS_ORDER_ITEM_ITEMNAME
dim CS_ORDER_ITEM_OPTIONNAME



sub ResetGlobalVarible()
	CS_ORDER_ITEM_CANCELYN = ""
	CS_ORDER_ITEM_CURRSTATE = ""
	CS_ORDER_ITEM_ISUPCHEBEASONG = ""
	CS_ORDER_ITEM_NO = ""
	CS_ORDER_ITEM_ITEMCOST = ""
	CS_ORDER_ITEM_MAKERID = ""

	CS_ORDER_ITEM_SELLCASH = ""
	CS_ORDER_ITEM_OPTADDPRICE = ""

	CS_ORDER_ITEM_ITEMNAME = ""
	CS_ORDER_ITEM_OPTIONNAME = ""
end sub



'==============================================================================
'신규 상품추가 : 기존에 없는 상품 주문 디테일에 추가.
'TODO : 주문마스터 결재완료인가 체크(추가 입금 발생시 추가 불가 : 입금확인 전에 출고될 수 있음)
'TODO : 주문마스터 출고완료인가 체크
'TODO : 기존에 배송업체에 속한 상품인가 체크(추가배송비 발생시 추가 불가)
'TODO : 기존 배송업체가 상품준비중인가 체크(추가불가)
'==============================================================================
function CSOrderAddNewItem(byval orderserial, byval itemid, byval itemoption, byval itemno)

	dim strSQL, result

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 추가할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품이 있습니다. 취소를 정상화하세요."
		else
			CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다."
		end if
		exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	'함수안에서 재고디비 반영
	result = CSOrderAddNewItemForce(orderserial, itemid, itemoption, itemno)

	CSOrderRecalculateOrder orderserial,false

end function



'==============================================================================
'기존 상품수량변경 : 기존에 있던 상품 수량 변경
'==============================================================================
function CSOrderModifyItemNo(byval orderserial, byval itemid, byval itemoption, byval itemno)

	dim strSQL, result

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 수량 변경할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "취소된 상품입니다."
		exit function
	end if

	'CS 에서는 출고이전에 변경할 수 있다.
	'if ((CS_ORDER_ITEM_ISUPCHEBEASONG = "Y") and ((CS_ORDER_ITEM_CURRSTATE = "3") or (CS_ORDER_ITEM_CURRSTATE = "7"))) then
	'	CS_ORDER_FUNCTION_RESULT = "업체배송의 경우 상품준비 이전에만 취소가 가능합니다."
	'	exit function
	'end if

	if (CS_ORDER_ITEM_CURRSTATE = "7") then
		CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
		exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	'// 함수 안에서 재고디비 반영
	result = CSOrderModifyItemNoForce(orderserial, itemid, itemoption, itemno)

	CSOrderRecalculateOrder orderserial, false

end function



'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션중 일부 변경
'==============================================================================
function CSOrderModifyItemOption(byval orderserial, byval itemid, byval itemoptionfrom, byval itemoptionto, byval itemno)

	dim strSQL, result

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 옵션을 변경할 수 없습니다."
		exit function
	end if

	if ((itemoptionfrom = "0000") or (itemoptionto = "0000")) then
		CS_ORDER_FUNCTION_RESULT = "옵션이 없는 상품입니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoptionfrom)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	if (CS_ORDER_ITEM_NO < CInt(itemno)) then
		CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다..(" & CStr(itemno) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
		exit function
	end if

	'CS 에서는 출고이전에 변경할 수 있다.
	'if ((CS_ORDER_ITEM_ISUPCHEBEASONG = "Y") and ((CS_ORDER_ITEM_CURRSTATE = "3") or (CS_ORDER_ITEM_CURRSTATE = "7"))) then
	'	CS_ORDER_FUNCTION_RESULT = "업체배송의 경우 상품준비 이전에만 변경이 가능합니다."
	'	exit function
	'end if

	if (CS_ORDER_ITEM_CURRSTATE = "7") then
		CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
		exit function
	end if



	'주문당시 두가지 이상의 옵션을 이미 주문한 경우, 주문당시 옵션가격이 상이하면 변경불가.
	dim itemcostfrom, itemcostto

	result = CSOrderGetItemState(orderserial, itemid, itemoptionfrom)
	itemcostfrom = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemid, itemoptionto)
	itemcostto = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemid, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			if (itemcostfrom <> itemcostto) then
				CS_ORDER_FUNCTION_RESULT = "주문당시 옵션가격이 달라 변경할 수 없습니다."
				exit function
			end if

			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
		end if
		'exit function
	end if



	'옵션변경의 경우, 옵션가격만 비교해서 같을 경우 옵션을 변경하고 기존상품의 가격정보(할인 등)를 모두 복사해온다.
	'옵션가격이 변동이 있더라도 동일하게 변동되면 여전히 변경가능.
	dim itemoptaddpricefrom, itemoptaddpriceto

	result = CSOrderGetItemOptionDeliveryPay(itemid, itemoptionfrom)
	itemoptaddpricefrom = CS_ORDER_ITEM_OPTADDPRICE

	result = CSOrderGetItemOptionDeliveryPay(itemid, itemoptionto)
	itemoptaddpriceto = CS_ORDER_ITEM_OPTADDPRICE

	if (itemoptaddpricefrom <> itemoptaddpriceto) then
		CS_ORDER_FUNCTION_RESULT = "옵션가격이 다른 경우 옵션을 변경할 수 없습니다."
		exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	result = CSOrderModifyItemOptionForce(orderserial, itemid, itemoptionfrom, itemoptionto, itemno)

	CSOrderRecalculateOrder orderserial,false

end function



'==============================================================================
'기존 상품변경 : 기존에 있던 상품중 일부 변경
'==============================================================================
function CSOrderChangeItem(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto, byval itemno)

	dim strSQL, result

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemidfrom = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 옵션을 변경할 수 없습니다."
		exit function
	end if

	if (itemidfrom = itemidto) then
		'' 2015-09-23, skyer9
		''CS_ORDER_FUNCTION_RESULT = "동일한 상품입니다. 옵션변경을 이용하세요."
		''exit function
	end if

	result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	if (CS_ORDER_ITEM_NO < CInt(itemno)) then
		CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다..(" & CStr(itemno) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
		exit function
	end if

	if (CS_ORDER_ITEM_CURRSTATE = "7") then
		CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
		exit function
	end if



	'주문당시 변경전후 상품을 모두 주문한 경우, 주문당시 옵션가격이 상이하면 변경불가.
	dim itemcostfrom, itemcostto

	result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)
	itemcostfrom = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	itemcostto = CS_ORDER_ITEM_ITEMCOST

	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			if (itemcostfrom <> itemcostto) then
				CS_ORDER_FUNCTION_RESULT = "상품가격이 달라 변경할 수 없습니다."
				exit function
			end if

			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
		end if
		'exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	result = CSOrderChangeItemForce(orderserial, itemidfrom, itemidto, itemoptionfrom, itemoptionto, itemno)

	CSOrderRecalculateOrder orderserial,false

end function



'==============================================================================
'기존 상품변경 : 기존에 있던 상품중 일부 변경
'==============================================================================
function CSOrderChangeItemArray(byval orderserial, byval arrFromItemId, byval arrToItemId, byval arrFromItemOption, byval arrToItemOption, byval arrFromItemNo, byval arrToItemNo)

	dim strSQL, result
	dim tmparrFromItemId, tmparrFromItemOption, tmparrFromItemNo
	dim itemidfrom, itemoptionfrom, itemnofrom

	CS_ORDER_FUNCTION_RESULT = ""

	tmparrFromItemId		= Split(arrFromItemId, "|")
	tmparrFromItemOption	= Split(arrFromItemOption, "|")
	tmparrFromItemNo		= Split(arrFromItemNo, "|")

	for i = 0 to UBound(tmparrFromItemId)
		if (Trim(tmparrFromItemId(i)) <> "") then
			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			itemidfrom = Trim(tmparrFromItemId(i))
			itemoptionfrom = Trim(tmparrFromItemOption(i))
			itemnofrom = Trim(tmparrFromItemNo(i))

			result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)

			if IsNull(CS_ORDER_ITEM_CANCELYN) then
				CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
				exit function
			end if

			if CS_ORDER_ITEM_CANCELYN = "Y" then
				CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
				exit function
			end if

			if (CS_ORDER_ITEM_NO < CInt(itemnofrom)) then
				CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다..(" & CStr(itemnofrom) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
				exit function
			end if

			if (CS_ORDER_ITEM_CURRSTATE = "7") then
				CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
				exit function
			end if

			ResetGlobalVarible()
			'--------------------------------------------------------------------------
		end if
	next

	result = CSOrderChangeItemArrayForce(orderserial, arrFromItemId, arrToItemId, arrFromItemOption, arrToItemOption, arrFromItemNo, arrToItemNo)

	CSOrderRecalculateOrder orderserial,false

end function



'==============================================================================
'주문상세 부분취소 : 기존에 있던 상품 한가지 취소
'==============================================================================
function CSOrderCancelItem(byval orderserial, byval itemid, byval itemoption)

	dim strSQL, result

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 취소할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	'CS 에서는 출고이전에 변경할 수 있다.
	'if ((CS_ORDER_ITEM_ISUPCHEBEASONG = "Y") and ((CS_ORDER_ITEM_CURRSTATE = "3") or (CS_ORDER_ITEM_CURRSTATE = "7"))) then
	'	CS_ORDER_FUNCTION_RESULT = "업체배송의 경우 상품준비 이전에만 취소가 가능합니다."
	'	exit function
	'end if

	if (CS_ORDER_ITEM_CURRSTATE = "7") then
		CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
		exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	'// 함수 안에서 재고디비 반영 : 0인경우 원수량전체
	result = CSOrderCancelItemForce(orderserial, itemid, itemoption)

	CSOrderRecalculateOrder orderserial, false

end function



'==============================================================================
'주문상세 부분취소 정상화 : 기존에 있던 취소된 상품 한가지를 정상화
'==============================================================================
function CSOrderRestoreCanceledItem(byval orderserial, byval itemid, byval itemoption)

	dim strSQL, result

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 처리할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "N" then
		CS_ORDER_FUNCTION_RESULT = "이미 정상 상품입니다."
		exit function
	end if

	'CS 에서는 출고이전에 변경할 수 있다.
	'if ((CS_ORDER_ITEM_ISUPCHEBEASONG = "Y") and ((CS_ORDER_ITEM_CURRSTATE = "3") or (CS_ORDER_ITEM_CURRSTATE = "7"))) then
	'	CS_ORDER_FUNCTION_RESULT = "업체배송의 경우 상품준비 이전에만 정상화가 가능합니다."
	'	exit function
	'end if

	if (CS_ORDER_ITEM_CURRSTATE = "7") then
		CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
		exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	result = CSOrderRestoreCanceledItemForce(orderserial, itemid, itemoption)

	CSOrderRecalculateOrder orderserial,false

end function



'==============================================================================
'상품상태 확인
'==============================================================================
function CSOrderGetItemState(byval orderserial, byval itemid, byval itemoption)

	dim strSQL

	strSQL = " select idx, itemname, itemoptionname, cancelyn, currstate, isupchebeasong, itemno, itemcost, makerid " + vbCrlf
	strSQL = strSQL & " from [db_order].[dbo].tbl_order_detail " + vbCrlf
	strSQL = strSQL & " where 1 = 1 " + vbCrlf
	strSQL = strSQL & " and orderserial='" & orderserial & "' " + vbCrlf
	strSQL = strSQL & " and itemid=" & itemid  & " " + vbCrlf
	strSQL = strSQL & " and itemoption='" & itemoption & "' " + vbCrlf
	'response.write "-------------" & strSQL



	rsget.Open strSQL,dbget,1
	if Not rsget.Eof then

		CS_ORDER_ITEM_ORDERDETAILIDX	= rsget("idx")
		CS_ORDER_ITEM_CANCELYN 			= rsget("cancelyn")
		CS_ORDER_ITEM_CURRSTATE 		= rsget("currstate")
		CS_ORDER_ITEM_ISUPCHEBEASONG 	= rsget("isupchebeasong")
		CS_ORDER_ITEM_NO 				= rsget("itemno")
		CS_ORDER_ITEM_ITEMCOST 			= rsget("itemcost")
		CS_ORDER_ITEM_MAKERID			= rsget("makerid")

		CS_ORDER_ITEM_ITEMNAME 			= rsget("itemname")
		CS_ORDER_ITEM_OPTIONNAME 		= rsget("itemoptionname")

		'CS_ORDER_ITEM_SELLCASH 			= rsget("sellcash")
		'CS_ORDER_ITEM_OPTADDPRICE		= rsget("optaddprice")
	else
		CS_ORDER_ITEM_ORDERDETAILIDX	= Null
		CS_ORDER_ITEM_CANCELYN 			= Null
		CS_ORDER_ITEM_CURRSTATE 		= Null
		CS_ORDER_ITEM_ISUPCHEBEASONG 	= Null
		CS_ORDER_ITEM_NO 				= Null
		CS_ORDER_ITEM_ITEMCOST 			= Null
		CS_ORDER_ITEM_MAKERID 			= Null

		CS_ORDER_ITEM_ITEMNAME 			= Null
		CS_ORDER_ITEM_OPTIONNAME 		= Null

		'CS_ORDER_ITEM_SELLCASH			= Null
		'CS_ORDER_ITEM_OPTADDPRICE		= Null
	end if
	rsget.close

end function



'==============================================================================
'옵션별 배송비 확인
'==============================================================================
function CSOrderGetItemOptionDeliveryPay(byval itemid, byval itemoption)

	dim strSQL

	strSQL = " select i.sellcash, IsNull(v.optaddprice,0) as optaddprice " + vbCrlf
	strSQL = strSQL & " from " + vbCrlf
	strSQL = strSQL & " 	[db_item].[dbo].tbl_item i " + vbCrlf
	strSQL = strSQL & " 	left join [db_item].[dbo].tbl_item_option v " + vbCrlf
	strSQL = strSQL & " 	on " + vbCrlf
	strSQL = strSQL & " 		1 = 1 " + vbCrlf
	strSQL = strSQL & " 		and i.itemid = v.itemid " + vbCrlf
	strSQL = strSQL & " where " + vbCrlf
	strSQL = strSQL & " 	1 = 1 " + vbCrlf
	strSQL = strSQL & " 	and i.itemid=" & itemid  & " " + vbCrlf
	strSQL = strSQL & " 	and IsNull(v.itemoption, '0000') = '" & itemoption & "' " + vbCrlf

	rsget.Open strSQL,dbget,1
	if Not rsget.Eof then
		CS_ORDER_ITEM_SELLCASH 			= rsget("sellcash")
		CS_ORDER_ITEM_OPTADDPRICE		= rsget("optaddprice")
	else
		CS_ORDER_ITEM_SELLCASH			= Null
		CS_ORDER_ITEM_OPTADDPRICE		= Null
	end if
	rsget.close

end function



'==============================================================================
''주문 Master 재계산
'==============================================================================
sub CSOrderRecalculateOrder(byVal orderserial, byVal isMinusjumun)
	dim sqlStr

	dim CURR_IsOLDOrder : CURR_IsOLDOrder = False

	if (GC_IsOLDOrder) then
		sqlStr = " select top 1 orderserial from [db_log].[dbo].tbl_old_order_master_2003 where orderserial = '" + CStr(orderserial) + "' "

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			CURR_IsOLDOrder = True
		end if
		rsget.Close
	end if

	if (CURR_IsOLDOrder) then
		sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
	else
		sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
	end if

	sqlStr = sqlStr + " set totalsum=IsNULL(T.dtotalsum,0)" + VbCrlf
	''sqlStr = sqlStr + " , totalcost=IsNULL(T.dtotalsum,0)"  + VbCrlf
	sqlStr = sqlStr + " , totalmileage=IsNULL(T.dtotalmileage,0)" + VbCrlf
	sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=IsNULL(T.dtotalitemcostCouponNotApplied,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select sum(itemcost*itemno) as dtotalsum, sum(mileage*itemno) as dtotalmileage, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied" + VbCrlf
	if (CURR_IsOLDOrder) then
		sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003" + VbCrlf
	else
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail" + VbCrlf
	end if
	sqlStr = sqlStr + "     where orderserial='" + orderserial + "'" + VbCrlf
	sqlStr = sqlStr + "     and cancelyn<>'Y'" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	if (CURR_IsOLDOrder) then
		sqlStr = sqlStr + " where [db_log].[dbo].tbl_old_order_master_2003.orderserial='" + orderserial + "'" + VbCrlf
	else
		sqlStr = sqlStr + " where [db_order].[dbo].tbl_order_master.orderserial='" + orderserial + "'" + VbCrlf
	end if

	dbget.Execute sqlStr

	sqlStr = " update m " + VbCrlf
	sqlStr = sqlStr + " set " + VbCrlf
	sqlStr = sqlStr + " 	m.sumPaymentEtc = IsNull(T.realPayedsum, 0) " + VbCrlf
    if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if

	sqlStr = sqlStr + " 	left join ( " + VbCrlf
	sqlStr = sqlStr + " 		select " + VbCrlf
	sqlStr = sqlStr + " 			orderserial " + VbCrlf
	sqlStr = sqlStr + " 			, IsNull(sum(realPayedsum), 0) as realPayedsum " + VbCrlf
	sqlStr = sqlStr + " 		from " + VbCrlf
	sqlStr = sqlStr + " 			[db_order].[dbo].tbl_order_PaymentEtc " + VbCrlf
	sqlStr = sqlStr + " 		where " + VbCrlf
	sqlStr = sqlStr + " 			1 = 1 " + VbCrlf
	sqlStr = sqlStr + " 			and orderserial = '" & orderserial & "' " + VbCrlf
	sqlStr = sqlStr + " 			and acctdiv in ('200', '900') " + VbCrlf
	sqlStr = sqlStr + " 		group by " + VbCrlf
	sqlStr = sqlStr + " 			orderserial " + VbCrlf
	sqlStr = sqlStr + " 	) T " + VbCrlf
	sqlStr = sqlStr + " 	on " + VbCrlf
	sqlStr = sqlStr + " 		m.orderserial = T.orderserial " + VbCrlf
	sqlStr = sqlStr + " where " + VbCrlf
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' " + VbCrlf

	dbget.Execute sqlStr

	if (CURR_IsOLDOrder) then
		sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" + VbCrlf
	else
		sqlStr = " update [db_order].[dbo].tbl_order_master" + VbCrlf
	end if
	sqlStr = sqlStr + " set subtotalprice=totalsum-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
	'sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=subtotalpriceCouponNotApplied-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " 	e set e.acctamount = (m.subtotalprice - m.sumpaymentetc), e.realpayedsum = (m.subtotalprice - m.sumpaymentetc) "
    if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if
	sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = e.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
	sqlStr = sqlStr + " 	and m.accountdiv = e.acctdiv "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '4' "
	sqlStr = sqlStr + " 	and m.accountdiv = '7' "

	dbget.Execute sqlStr

	'// e.acctdiv = '120' 네이버 포인트
	'// 참조 주문번호 : 16092146018
  	sqlStr = " update e set e.realPayedSum = (T.realpayedsum - T.realpayedsum120) "
  	sqlStr = sqlStr + " from "
  	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
  	sqlStr = sqlStr + " 	join ( "
  	sqlStr = sqlStr + " 		select m.orderserial, m.accountdiv, (m.subtotalprice - m.sumpaymentetc) as realpayedsum, IsNull(sum(Case when e.acctdiv = '120' then e.realpayedsum else 0 end),0) as realpayedsum120 "
	if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if
  	sqlStr = sqlStr + " 		join [db_order].[dbo].tbl_order_PaymentEtc e "
  	sqlStr = sqlStr + " 		on "
  	sqlStr = sqlStr + " 			1 = 1 "
  	sqlStr = sqlStr + " 			and m.orderserial = e.orderserial "
  	sqlStr = sqlStr + " 			and e.acctdiv in (m.accountdiv, '120') "
  	sqlStr = sqlStr + " 		where "
  	sqlStr = sqlStr + " 			m.orderserial = '" & orderserial & "' "
  	sqlStr = sqlStr + " 		group by "
  	sqlStr = sqlStr + " 			m.orderserial, m.accountdiv, (m.subtotalprice - m.sumpaymentetc) "
  	sqlStr = sqlStr + " 	) T "
  	sqlStr = sqlStr + " 	on "
  	sqlStr = sqlStr + " 		1 = 1 "
  	sqlStr = sqlStr + " 		and e.orderserial = T.orderserial "
  	sqlStr = sqlStr + " 		and e.acctdiv = T.accountdiv "
	dbget.Execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr + " set subtotalpriceCouponNotApplied = (case when T.dtotalitemcostCouponNotApplied = 0 then 0 else subtotalpriceCouponNotApplied end) "
    if (CURR_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr  + " from [db_order].[dbo].tbl_order_master m"
    end if
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select "
	sqlStr = sqlStr + " 			orderserial, sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalitemcostCouponNotApplied "
	if (CURR_IsOLDOrder) then
		sqlStr = sqlStr + " 	from [db_log].[dbo].tbl_old_order_detail_2003" + VbCrlf
	else
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail" + VbCrlf
	end if
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and orderserial = '" & orderserial & "' "
	sqlStr = sqlStr + " 			and cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and itemid <> 0 "
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	orderserial "
	sqlStr = sqlStr + " 	) T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = T.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	m.orderserial = '" & orderserial & "' "

	dbget.Execute sqlStr

end sub



'==============================================================================
'신규 상품추가 : 기존에 없는 상품 주문 디테일에 추가.
'==============================================================================
function CSOrderAddNewItemForce(byval orderserial, byval itemid, byval itemoption, byval itemno)

	dim strSQL, result, makeridforadd

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 추가할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품이 있습니다. 취소를 정상화하세요."
		else
			CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다."
		end if
		exit function
	end if

	makeridforadd = CS_ORDER_ITEM_MAKERID

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

    dim tmpjumundiv, tmpipkumdiv, iid, tmpuserlevel
	dim t_orgitemcost
	dim t_itemcost,t_itemvat,t_mileage
	dim t_itemname, t_itemoptionname, t_makerid
	dim t_buycash, t_buyvat, t_vatinclude, t_upchebeasong
	dim t_sailyn, t_itemdiv, t_mwdiv, t_deliverytype

	strSQL = "select top 1 idx, jumundiv, ipkumdiv, userlevel from [db_order].[dbo].tbl_order_master " & vbCrlf
	strSQL = strSQL & " where orderserial='" + orderserial + "'" & vbCrlf
	rsget.Open strSQL,dbget,1
	if rsget.Eof then
		iid = 0
	else
		iid = rsget("idx")
		tmpjumundiv = rsget("jumundiv")
		tmpipkumdiv = rsget("ipkumdiv")
		tmpuserlevel = rsget("userlevel")
	end if
	rsget.Close

	strSQL = "select top 1 i.orgprice, i.sellcash, isnull(i.mileage,0) as mileage, i.makerid ,i.buycash, " & vbCrlf
	strSQL = strSQL + " i.itemname, i.vatinclude, i.deliverytype, i.sailyn, i.itemdiv, i.mwdiv, IsNull(v.optionname,'') as codeview, IsNull(v.optaddprice,0) as optaddprice, IsNull(v.optaddbuyprice,0) as optaddbuyprice" & vbCrlf
	strSQL = strSQL + " from [db_item].[dbo].tbl_item i" & vbCrlf
	strSQL = strSQL + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid and IsNull(v.itemoption, '0000')='" + itemoption + "'" & vbCrlf
	strSQL = strSQL + " where i.itemid=" & itemid & vbCrlf
	rsget.Open strSQL,dbget,1

	if Not rsget.Eof then
		t_orgitemcost = rsget("orgprice") + rsget("optaddprice")
		t_itemcost = rsget("sellcash") + rsget("optaddprice")

		t_mileage = rsget("mileage")
		if date() >= "2018-08-01" then
			' vip, vip gold, staff, family	' 기본이 상품가격에 0.5% 적립. x2 해서 1%로 만듬
			if tmpuserlevel="2" or tmpuserlevel="3" or tmpuserlevel="7" or tmpuserlevel="8" then
				t_mileage = clng(clng(rsget("mileage"))*2)

			' vvip	' 기본이 상품가격에 0.5% 적립. x2.6 해서 1.3%로 만듬
			elseif tmpuserlevel="4" then
				t_mileage = clng(clng(rsget("mileage"))*2.6)

			' BIZ
			elseif tmpuserlevel="9" then
				t_mileage = clng(clng(rsget("mileage"))*2)
			end if
		end if

		t_itemname  = rsget("itemname")
		t_itemoptionname = rsget("codeview")
		t_makerid = rsget("makerid")
		t_buycash = rsget("buycash")    + rsget("optaddbuyprice")
		t_vatinclude	= rsget("vatinclude")
		t_upchebeasong  = rsget("deliverytype")
		t_deliverytype  = rsget("deliverytype")
		t_sailyn  = rsget("sailyn")
		t_itemdiv  = rsget("itemdiv")
        t_mwdiv    = rsget("mwdiv")

		if (t_upchebeasong="2") or (t_upchebeasong="5") or (t_upchebeasong="9") or (t_upchebeasong="7") then
			t_upchebeasong="Y"
		else
			t_upchebeasong="N"
		end if
	end if
	rsget.close

	strSQL = "insert into [db_order].[dbo].tbl_order_detail"
	strSQL = strSQL & "(masteridx, orderserial,itemid,itemoption,itemno,itemcost,itemvat,mileage,"
	strSQL = strSQL & "reducedPrice, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, cancelyn,itemname,itemoptionname,makerid,buycash,vatinclude,"
	strSQL = strSQL & "isupchebeasong,issailitem,oitemdiv,omwdiv,odlvtype" & vbCrlf

	if (tmpjumundiv="9") then
	    '' 마이너스 주문인 경우.
	    strSQL = strSQL & ",currstate,beasongdate,upcheconfirmdate " & vbCrlf
	elseif (tmpipkumdiv="4") then
	    '' 주문상태가 결제완료인 경우
	elseif (tmpipkumdiv="5") or (tmpipkumdiv="6") or (tmpipkumdiv="7") then
	    '' 주문상태가 상품준비/일부출고인경우
	    strSQL = strSQL & ",currstate " & vbCrlf
	elseif (tmpipkumdiv="8") then
	    '' 주문상태가 출고 완료인경우
	    strSQL = strSQL & ",currstate " & vbCrlf
	end if

	strSQL = strSQL & " ) " & vbCrlf
	strSQL = strSQL & " values(" + CStr(iid) + "," & vbCrlf
	strSQL = strSQL & "'" & orderserial & "'," & vbCrlf
	strSQL = strSQL & itemid & "," & vbCrlf
	strSQL = strSQL & "'" & itemoption & "'," & vbCrlf
	strSQL = strSQL & itemno & "," & vbCrlf
	strSQL = strSQL & t_itemcost & "," & vbCrlf
	strSQL = strSQL & CLng(t_itemcost*1/11) & "," & vbCrlf
	strSQL = strSQL & t_mileage & ","
	strSQL = strSQL & t_itemcost & "," & vbCrlf
	strSQL = strSQL & t_orgitemcost & "," & vbCrlf
	strSQL = strSQL & t_itemcost & "," & vbCrlf
	strSQL = strSQL & t_buycash & "," & vbCrlf
	strSQL = strSQL & "'A'," & vbCrlf
	strSQL = strSQL & "'" & NewHtml2db(t_itemname) & "'," & vbCrlf
	strSQL = strSQL & "'" & NewHtml2db(t_itemoptionname) & "'," & vbCrlf
	strSQL = strSQL & "'" & t_makerid & "'," & vbCrlf
	strSQL = strSQL & "" & t_buycash & "," & vbCrlf
	strSQL = strSQL & "'" & t_vatinclude & "'," & vbCrlf
	strSQL = strSQL & "'" & t_upchebeasong & "'," & vbCrlf
	strSQL = strSQL & "'" & t_sailyn & "'," & vbCrlf
	strSQL = strSQL & "'" & t_itemdiv & "'," & vbCrlf
	strSQL = strSQL & "'" & t_mwdiv & "'," & vbCrlf
	strSQL = strSQL & "'" & t_deliverytype & "'" & vbCrlf

	if (tmpjumundiv="9") then
	    '' 마이너스 주문인 경우.
	    strSQL = strSQL & ",'7'" & vbCrlf
		strSQL = strSQL & ",getdate()" & vbCrlf
		strSQL = strSQL & ",getdate()" & vbCrlf
	elseif (tmpipkumdiv="4") then
	    '' 주문상태가 결제완료인 경우

	elseif (tmpipkumdiv="5") or (tmpipkumdiv="6") or (tmpipkumdiv="7") then
	    '' 주문상태가 상품준비/일부출고인경우
	    strSQL = strSQL & ",'2'" & vbCrlf
	elseif (tmpipkumdiv="8") then
	    '' 주문상태가 출고 완료인경우
	    strSQL = strSQL & ",'2'" & vbCrlf
	end if

	strSQL = strSQL & ")"

	'response.write strSQL
	rsget.Open strSQL,dbget,1


    ''재고디비 반영
    '// itemno = itemno*-1
    '// strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr(itemno)
    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr(itemno)
    dbget.Execute strSQL

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	'상품이 새로 추가된 경우가 있어서 여기서 구해야 한다.
	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	makeridforadd = CS_ORDER_ITEM_MAKERID

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

end function



'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션 변경
'==============================================================================
function CSOrderModifyItemOptionForce(byval orderserial, byval itemid, byval itemoptionfrom, byval itemoptionto, byval itemno)

	dim strSQL, result, itemnoforadd, isitemoptiontoexist, iscancelall, itemnoafterchange, makeridforadd

	CS_ORDER_FUNCTION_RESULT = ""



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 옵션을 변경할 수 없습니다."
		exit function
	end if

	if ((itemoptionfrom = "0000") or (itemoptionto = "0000")) then
		CS_ORDER_FUNCTION_RESULT = "옵션이 없는 상품입니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoptionfrom)

	makeridforadd = CS_ORDER_ITEM_MAKERID

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	if CS_ORDER_ITEM_NO < CInt(itemno) then
		CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다.(" & CStr(itemno) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
		exit function
	end if

	if (CS_ORDER_ITEM_NO = CInt(itemno)) then
		iscancelall = "Y"
		itemnoafterchange = 0
	else
		itemnoafterchange = CS_ORDER_ITEM_NO - CInt(itemno)
	end if



	itemnoforadd = itemno

	result = CSOrderGetItemOptionDeliveryPay(itemid, itemoptionto)

	result = CSOrderGetItemState(orderserial, itemid, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		isitemoptiontoexist = "Y"
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
			itemnoforadd = itemnoforadd + CS_ORDER_ITEM_NO
		end if
		'exit function
	end if



	'주문당시 두가지 이상의 옵션을 이미 주문한 경우, 주문당시 옵션가격이 상이하면 변경불가.
	dim itemcostfrom, itemcostto

	result = CSOrderGetItemState(orderserial, itemid, itemoptionfrom)
	itemcostfrom = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemid, itemoptionto)
	itemcostto = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemid, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			if (itemcostfrom <> itemcostto) then
				CS_ORDER_FUNCTION_RESULT = "주문당시 옵션가격이 달라 변경할 수 없습니다."
				exit function
			end if

			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
		end if
		'exit function
	end if



	'옵션변경의 경우, 옵션가격만 비교해서 같을 경우 옵션을 변경하고 기존상품의 가격정보(할인 등)를 모두 복사해온다.
	'옵션가격이 변동이 있더라도 동일하게 변동되면 여전히 변경가능.
	dim itemoptaddpricefrom, itemoptaddpriceto

	result = CSOrderGetItemOptionDeliveryPay(itemid, itemoptionfrom)
	itemoptaddpricefrom = CS_ORDER_ITEM_OPTADDPRICE

	result = CSOrderGetItemOptionDeliveryPay(itemid, itemoptionto)
	itemoptaddpriceto = CS_ORDER_ITEM_OPTADDPRICE

	if (itemoptaddpricefrom <> itemoptaddpriceto) then
		CS_ORDER_FUNCTION_RESULT = "옵션가격이 다른 경우 옵션을 변경할 수 없습니다."
		exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------



	'변경전 옵션 취소

	if (iscancelall = "Y") then

		'// 함수 안에서 재고디비 반영 : 0인경우 원수량전체
		result = CSOrderCancelItemForce(orderserial, itemid, itemoptionfrom)
		'response.write "aaaaaaaaaaaaaaaa" & CS_ORDER_FUNCTION_RESULT

	else

		'// 함수 안에서 재고디비 반영
		result = CSOrderModifyItemNoForce(orderserial, itemid, itemoptionfrom, itemnoafterchange)
		'response.write "bbbbbbbbbbbbbbbbb" & CS_ORDER_FUNCTION_RESULT

	end if



	'--------------------------------------------------------------------------
	'변경후 옵션 추가
    dim tmpjumundiv, tmpipkumdiv, iid
	dim t_itemcost,t_itemvat,t_mileage
	dim t_itemname, t_itemoptionname, t_makerid
	dim t_buycash, t_buyvat, t_vatinclude, t_upchebeasong
	dim t_sailyn, t_itemdiv, t_mwdiv, t_deliverytype


	if (isitemoptiontoexist <> "Y") then

		'변경후 옵션 추가 - 기존에 취소된 옵션이 없는경우
		'함수안에서 재고디비 반영
		result = CSOrderAddNewItemForce(orderserial, itemid, itemoptionto, itemnoforadd)

	else

		'취소된경우 취소 정상화후 수량설정, 취소가 안된경우 기존재하는 값에 추가.
		strSQL = "update	 [db_order].[dbo].tbl_order_detail "
		strSQL = strSQL & " set cancelyn='N'" + vbCrlf
		''strSQL = strSQL & " ,orderdate=getdate()" + vbCrlf
		strSQL = strSQL & " ,itemno = " & CStr(itemnoforadd) & " " + vbCrlf
		strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
		strSQL = strSQL & " and itemid=" & itemid  & vbCrlf
		strSQL = strSQL & " and itemoption='" & itemoptionto & "'" & vbCrlf
		rsget.Open strSQL,dbget,1

	    ''재고디비 반영
	    '// itemno = itemno*-1
	    '// strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoptionto & "'," & CStr(itemno)
	    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoptionto & "'," & CStr(itemno)
	    'response.write "aaa" & strSQL
	    dbget.Execute strSQL

	    result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

	end if



	'--------------------------------------------------------------------------
	'변경전 디테일정보 가져와서 복사

	Call CSOrderCopyItemInfo(orderserial, itemid, itemid, itemoptionfrom, itemoptionto)

end function



'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션 변경
'상품정보 복사는 따로 해주어야 한다.
'==============================================================================
function CSOrderChangeItemForce(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto, byval itemno)

	dim strSQL, result, itemnoforadd, isitemoptiontoexist, iscancelall, itemnoafterchange, makeridforadd

	CS_ORDER_FUNCTION_RESULT = ""



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemidfrom = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 옵션을 변경할 수 없습니다."
		exit function
	end if

	if (itemidfrom = itemidto) then
		'강제변경가능
		'CS_ORDER_FUNCTION_RESULT = "동일한 상품입니다. 옵션변경을 이용하세요."
		'exit function
	end if

	result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)

	makeridforadd = CS_ORDER_ITEM_MAKERID

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	if CS_ORDER_ITEM_NO < CInt(itemno) then
		CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다.(" & CStr(itemno) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
		exit function
	end if

	if (CS_ORDER_ITEM_NO = CInt(itemno)) then
		iscancelall = "Y"
		itemnoafterchange = 0
	else
		itemnoafterchange = CS_ORDER_ITEM_NO - CInt(itemno)
	end if



	itemnoforadd = itemno

	result = CSOrderGetItemOptionDeliveryPay(itemidto, itemoptionto)

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		isitemoptiontoexist = "Y"
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
			itemnoforadd = itemnoforadd + CS_ORDER_ITEM_NO
		end if
		'exit function
	end if



	'주문당시 두가지 이상의 옵션을 이미 주문한 경우, 주문당시 옵션가격이 상이하면 변경불가.
	dim itemcostfrom, itemcostto

	result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)
	itemcostfrom = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	itemcostto = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			if (itemcostfrom <> itemcostto) then
				'CS_ORDER_FUNCTION_RESULT = "주문당시 옵션가격이 달라 변경할 수 없습니다."
				'exit function
			end if

			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
		end if
		'exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------


	'변경전 옵션 취소

	if (iscancelall = "Y") then

		'// 함수 안에서 재고디비 반영 : 0인경우 원수량전체
		result = CSOrderCancelItemForce(orderserial, itemidfrom, itemoptionfrom)
		'response.write "aaaaaaaaaaaaaaaa" & CS_ORDER_FUNCTION_RESULT

	else

		'// 함수 안에서 재고디비 반영
		result = CSOrderModifyItemNoForce(orderserial, itemidfrom, itemoptionfrom, itemnoafterchange)
		'response.write "bbbbbbbbbbbbbbbbb" & CS_ORDER_FUNCTION_RESULT

	end if


	'--------------------------------------------------------------------------
	'변경후 옵션 추가
    dim tmpjumundiv, tmpipkumdiv, iid
	dim t_itemcost,t_itemvat,t_mileage
	dim t_itemname, t_itemoptionname, t_makerid
	dim t_buycash, t_buyvat, t_vatinclude, t_upchebeasong
	dim t_sailyn, t_itemdiv, t_mwdiv, t_deliverytype


	if (isitemoptiontoexist <> "Y") then

		'변경후 옵션 추가 - 기존에 취소된 옵션이 없는경우
		'함수안에서 재고디비 반영
		result = CSOrderAddNewItemForce(orderserial, itemidto, itemoptionto, itemnoforadd)

		'// 판매가(할인가) 동일하면 구매마일리지 복사
		Call CSOrderUpdateBuyMileage(orderserial, itemidfrom, itemoptionfrom, itemidto, itemoptionto)
	else

		'취소된경우 취소 정상화후 수량설정, 취소가 안된경우 기존재하는 값에 추가.
		strSQL = "update	 [db_order].[dbo].tbl_order_detail "
		strSQL = strSQL & " set cancelyn='N'" + vbCrlf
		''strSQL = strSQL & " ,orderdate=getdate()" + vbCrlf
		strSQL = strSQL & " ,itemno = " & CStr(itemnoforadd) & " " + vbCrlf
		strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
		strSQL = strSQL & " and itemid=" & itemidto  & vbCrlf
		strSQL = strSQL & " and itemoption='" & itemoptionto & "'" & vbCrlf
		rsget.Open strSQL,dbget,1

	    ''재고디비 반영
	    '// itemno = itemno*-1
	    '// strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemidto & ",'" & itemoptionto & "'," & CStr(itemno)
	    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemidto & ",'" & itemoptionto & "'," & CStr(itemno)
	    dbget.Execute strSQL

	    result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

	end if

end function


'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션 변경
'상품정보 복사는 따로 해주어야 한다.
'==============================================================================
function CSOrderChangeItemForceNEW(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto, byval itemnocancel, byval itemnoadd)

	dim strSQL, result, itemnoforadd, isitemoptiontoexist, iscancelall, itemnoafterchange, makeridforadd

	CS_ORDER_FUNCTION_RESULT = ""



	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemidfrom = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 옵션을 변경할 수 없습니다."
		exit function
	end if

	if (itemidfrom = itemidto) then
		'강제변경가능
		'CS_ORDER_FUNCTION_RESULT = "동일한 상품입니다. 옵션변경을 이용하세요."
		'exit function
	end if

	result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)

	makeridforadd = CS_ORDER_ITEM_MAKERID

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	if CS_ORDER_ITEM_NO < CLng(itemnocancel) then
		CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다.(" & CStr(itemnocancel) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
		exit function
	end if

	if (CS_ORDER_ITEM_NO = CLng(itemnocancel)) then
		iscancelall = "Y"
		itemnoafterchange = 0
	else
		itemnoafterchange = CS_ORDER_ITEM_NO - CInt(itemnocancel)
	end if



	itemnoforadd = itemnoadd

	result = CSOrderGetItemOptionDeliveryPay(itemidto, itemoptionto)

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		isitemoptiontoexist = "Y"
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
			itemnoforadd = itemnoforadd + CS_ORDER_ITEM_NO
		end if
		'exit function
	end if



	'주문당시 두가지 이상의 옵션을 이미 주문한 경우, 주문당시 옵션가격이 상이하면 변경불가.
	dim itemcostfrom, itemcostto

	result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)
	itemcostfrom = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	itemcostto = CS_ORDER_ITEM_ITEMCOST

	result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
	if not IsNull(CS_ORDER_ITEM_CANCELYN) then
		if (CS_ORDER_ITEM_CANCELYN = "Y") then
			'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
		else
			if (itemcostfrom <> itemcostto) then
				'CS_ORDER_FUNCTION_RESULT = "주문당시 옵션가격이 달라 변경할 수 없습니다."
				'exit function
			end if

			'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
			'exit function
		end if
		'exit function
	end if

	ResetGlobalVarible()
	'--------------------------------------------------------------------------


	'변경전 옵션 취소

	if (iscancelall = "Y") then

		'// 함수 안에서 재고디비 반영 : 0인경우 원수량전체
		result = CSOrderCancelItemForce(orderserial, itemidfrom, itemoptionfrom)
		'response.write "aaaaaaaaaaaaaaaa" & CS_ORDER_FUNCTION_RESULT

	else

		'// 함수 안에서 재고디비 반영
		result = CSOrderModifyItemNoForce(orderserial, itemidfrom, itemoptionfrom, itemnoafterchange)
		'response.write "bbbbbbbbbbbbbbbbb" & CS_ORDER_FUNCTION_RESULT

	end if


	'--------------------------------------------------------------------------
	'변경후 옵션 추가
    dim tmpjumundiv, tmpipkumdiv, iid
	dim t_itemcost,t_itemvat,t_mileage
	dim t_itemname, t_itemoptionname, t_makerid
	dim t_buycash, t_buyvat, t_vatinclude, t_upchebeasong
	dim t_sailyn, t_itemdiv, t_mwdiv, t_deliverytype


	if (isitemoptiontoexist <> "Y") then

		'변경후 옵션 추가 - 기존에 취소된 옵션이 없는경우
		'함수안에서 재고디비 반영
		result = CSOrderAddNewItemForce(orderserial, itemidto, itemoptionto, itemnoforadd)

		'// 판매가(할인가) 동일하면 구매마일리지 복사
		Call CSOrderUpdateBuyMileage(orderserial, itemidfrom, itemoptionfrom, itemidto, itemoptionto)
	else

		'취소된경우 취소 정상화후 수량설정, 취소가 안된경우 기존재하는 값에 추가.
		strSQL = "update	 [db_order].[dbo].tbl_order_detail "
		strSQL = strSQL & " set cancelyn='N'" + vbCrlf
		''strSQL = strSQL & " ,orderdate=getdate()" + vbCrlf
		strSQL = strSQL & " ,itemno = " & CStr(itemnoforadd) & " " + vbCrlf
		strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
		strSQL = strSQL & " and itemid=" & itemidto  & vbCrlf
		strSQL = strSQL & " and itemoption='" & itemoptionto & "'" & vbCrlf
		rsget.Open strSQL,dbget,1

	    ''재고디비 반영
	    '// itemno = itemno*-1
	    '// strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemidto & ",'" & itemoptionto & "'," & CStr(itemno)
	    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemidto & ",'" & itemoptionto & "'," & CStr(itemnoadd)
	    dbget.Execute strSQL

	    result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

	end if

end function


function CSOrderUpdateBuyMileage(orderserial, itemidfrom, itemoptionfrom, itemidto, itemoptionto)
	dim strSQL

	'// 판매가(할인가) 동일하면 구매마일리지 복사
	strSQL = " update " & vbCrlf
	strSQL = strSQL & " a set a.mileage = b.mileage " & vbCrlf
	strSQL = strSQL & " from " & vbCrlf
	strSQL = strSQL & " 	[db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " 	join [db_order].[dbo].tbl_order_detail b " & vbCrlf
	strSQL = strSQL & " 	on " & vbCrlf
	strSQL = strSQL & " 		1 = 1 " & vbCrlf
	strSQL = strSQL & " 		and a.orderserial = '" + CStr(orderserial) + "' " & vbCrlf
	strSQL = strSQL & " 		and b.orderserial = '" + CStr(orderserial) + "' " & vbCrlf
	strSQL = strSQL & " 		and a.itemid = " + CStr(itemidto) + " " & vbCrlf
	strSQL = strSQL & " 		and a.itemoption = '" + CStr(itemoptionto) + "' " & vbCrlf
	strSQL = strSQL & " 		and b.itemid = " + CStr(itemidfrom) + " " & vbCrlf
	strSQL = strSQL & " 		and b.itemoption = '" + CStr(itemoptionfrom) + "' " & vbCrlf
	strSQL = strSQL & " where " & vbCrlf
	strSQL = strSQL & " 	a.itemcost = b.itemcost " & vbCrlf

	rsget.Open strSQL,dbget,1

end function


'==============================================================================
'기존 상품옵션변경 : 기존에 있던 상품 옵션 변경
'상품정보 복사는 따로 해주어야 한다.
'==============================================================================
function CSOrderChangeItemArrayForce(orderserial, arrFromItemId, arrToItemId, arrFromItemOption, arrToItemOption, arrFromItemNo, arrToItemNo)
	dim strSQL, result, itemnoforadd, isitemoptiontoexist, iscancelall, itemnoafterchange, makeridfrom, makeridto
	dim tmparrFromItemId, tmparrFromItemOption, tmparrFromItemNo
	dim tmparrToItemId, tmparrToItemOption, tmparrToItemNo
	dim itemidfrom, itemoptionfrom, itemnofrom
	dim itemidto, itemoptionto, itemnoto

	CS_ORDER_FUNCTION_RESULT = ""

	tmparrFromItemId		= Split(arrFromItemId, "|")
	tmparrFromItemOption	= Split(arrFromItemOption, "|")
	tmparrFromItemNo		= Split(arrFromItemNo, "|")

	'상품취소
	for i = 0 to UBound(tmparrFromItemId)
		if (Trim(tmparrFromItemId(i)) <> "") then
			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			itemidfrom = Trim(tmparrFromItemId(i))
			itemoptionfrom = Trim(tmparrFromItemOption(i))
			itemnofrom = Trim(tmparrFromItemNo(i))

			makeridfrom = ""
			iscancelall = ""
			itemnoafterchange = ""

			result = CSOrderGetItemState(orderserial, itemidfrom, itemoptionfrom)

			if IsNull(CS_ORDER_ITEM_CANCELYN) then
				CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
				exit function
			end if

			if CS_ORDER_ITEM_CANCELYN = "Y" then
				CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
				exit function
			end if

			if (CS_ORDER_ITEM_NO < CInt(itemnofrom)) then
				CS_ORDER_FUNCTION_RESULT = "변경할 옵션의 수량이 부족합니다..(" & CStr(itemnofrom) & "/" & CStr(CS_ORDER_ITEM_NO) & ")"
				exit function
			end if

			'if (CS_ORDER_ITEM_CURRSTATE = "7") then
			'	CS_ORDER_FUNCTION_RESULT = "이미 출고 완료된 상품입니다."
			'	exit function
			'end if

			makeridfrom = CS_ORDER_ITEM_MAKERID

			if (CS_ORDER_ITEM_NO = CInt(itemnofrom)) then
				iscancelall = "Y"
				itemnoafterchange = 0
			else
				iscancelall = "N"
				itemnoafterchange = CS_ORDER_ITEM_NO - CInt(itemnofrom)
			end if

			if (iscancelall = "Y") then
				'// 함수 안에서 재고디비 반영 : 0인경우 원수량전체
				result = CSOrderCancelItemForce(orderserial, itemidfrom, itemoptionfrom)
			else
				'// 함수 안에서 재고디비 반영
				result = CSOrderModifyItemNoForce(orderserial, itemidfrom, itemoptionfrom, itemnoafterchange)
			end if

		    result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridfrom)

			ResetGlobalVarible()
			'--------------------------------------------------------------------------
		end if
	next

	tmparrToItemId		= Split(arrToItemId, "|")
	tmparrToItemOption	= Split(arrToItemOption, "|")
	tmparrToItemNo		= Split(arrToItemNo, "|")

	'상품추가
	for i = 0 to UBound(tmparrToItemId)
		if (Trim(tmparrToItemId(i)) <> "") then
			'--------------------------------------------------------------------------
			ResetGlobalVarible()

			itemidto = Trim(tmparrToItemId(i))
			itemoptionto = Trim(tmparrToItemOption(i))
			itemnoto = Trim(tmparrToItemNo(i))

			itemnoforadd = itemnoto

			result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)
			if not IsNull(CS_ORDER_ITEM_CANCELYN) then
				isitemoptiontoexist = "Y"
				if (CS_ORDER_ITEM_CANCELYN = "Y") then
					'CS_ORDER_FUNCTION_RESULT = "변경후 옵션이 취소상태입니다. 취소된 상품을 정상화하세요."
				else
					'CS_ORDER_FUNCTION_RESULT = "이미 상품이 있습니다. 상품취소 후 수량변경하세요."
					'exit function
					itemnoforadd = itemnoforadd + CS_ORDER_ITEM_NO
				end if
				'exit function
			end if

			if (isitemoptiontoexist <> "Y") then

				'변경후 옵션 추가 - 기존에 취소된 옵션이 없는경우
				'함수안에서 재고디비 반영
				'배송비계산도 함수안에서 한다.
				result = CSOrderAddNewItemForce(orderserial, itemidto, itemoptionto, itemnoforadd)

			else

				'취소된경우 취소 정상화후 수량설정, 취소가 안된경우 기존재하는 값에 추가.
				strSQL = "update	 [db_order].[dbo].tbl_order_detail "
				strSQL = strSQL & " set cancelyn='N'" + vbCrlf
				''strSQL = strSQL & " ,orderdate=getdate()" + vbCrlf
				strSQL = strSQL & " ,itemno = " & CStr(itemnoforadd) & " " + vbCrlf
				strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
				strSQL = strSQL & " and itemid=" & itemidto  & vbCrlf
				strSQL = strSQL & " and itemoption='" & itemoptionto & "'" & vbCrlf
				rsget.Open strSQL,dbget,1

			    ''재고디비 반영
			    '// itemnoforadd = itemnoforadd*-1
			    '// strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemidto & ",'" & itemoptionto & "'," & CStr(itemnoforadd)
			    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemidto & ",'" & itemoptionto & "'," & CStr(itemnoforadd)
			    dbget.Execute strSQL

				ResetGlobalVarible()

				result = CSOrderGetItemState(orderserial, itemidto, itemoptionto)

				makeridto = CS_ORDER_ITEM_MAKERID

			    result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridto)

			end if

			ResetGlobalVarible()
			'--------------------------------------------------------------------------
		end if
	next

end function

'==============================================================================
'상품정보 복사
'==============================================================================
function CSOrderCopyItemInfo(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto)
	dim strSQL

	strSQL = "update a " & vbCrlf
	strSQL = strSQL & " set " & vbCrlf
	strSQL = strSQL & " 	a.itemcost = b.itemcost " & vbCrlf
	strSQL = strSQL & " 	, a.reducedPrice = b.reducedPrice " & vbCrlf
	strSQL = strSQL & " 	, a.mileage = b.mileage " & vbCrlf					'// 부여 마일리지 복사
	strSQL = strSQL & " 	, a.currstate = b.currstate " & vbCrlf
	strSQL = strSQL & " 	, a.songjangno = b.songjangno " & vbCrlf
	strSQL = strSQL & " 	, a.songjangdiv = b.songjangdiv " & vbCrlf
	strSQL = strSQL & " 	, a.buycash = b.buycash " & vbCrlf
	strSQL = strSQL & " 	, a.itemvat = b.itemvat " & vbCrlf
	strSQL = strSQL & " 	, a.vatinclude = b.vatinclude " & vbCrlf
	strSQL = strSQL & " 	, a.beasongdate = b.beasongdate " & vbCrlf
	strSQL = strSQL & " 	, a.isupchebeasong = b.isupchebeasong " & vbCrlf
	strSQL = strSQL & " 	, a.omwdiv = b.omwdiv " & vbCrlf
	strSQL = strSQL & " 	, a.odlvType = b.odlvType " & vbCrlf
	strSQL = strSQL & " 	, a.issailitem = b.issailitem " & vbCrlf
	strSQL = strSQL & " 	, a.upcheconfirmdate = b.upcheconfirmdate " & vbCrlf
	strSQL = strSQL & " 	, a.oitemdiv = b.oitemdiv " & vbCrlf
	strSQL = strSQL & " 	, a.requiredetail = b.requiredetail " & vbCrlf
	strSQL = strSQL & " 	, a.itemcouponidx = b.itemcouponidx " & vbCrlf
	strSQL = strSQL & " 	, a.bonuscouponidx = b.bonuscouponidx " & vbCrlf
	strSQL = strSQL & " 	, a.passday = b.passday " & vbCrlf
	strSQL = strSQL & " 	, a.orgitemcost = b.orgitemcost " & vbCrlf
	strSQL = strSQL & " 	, a.itemcostCouponNotApplied = b.itemcostCouponNotApplied " & vbCrlf
	strSQL = strSQL & " 	, a.buycashCouponNotApplied = b.buycashCouponNotApplied " & vbCrlf
	strSQL = strSQL & " 	, a.odlvfixday = b.odlvfixday " & vbCrlf
	strSQL = strSQL & " 	, a.plusSaleDiscount = b.plusSaleDiscount " & vbCrlf
	strSQL = strSQL & " 	, a.specialshopDiscount = b.specialshopDiscount " & vbCrlf
	strSQL = strSQL & " 	, a.etcDiscount = b.etcDiscount " & vbCrlf
	strSQL = strSQL & " from [db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " 	, ( " & vbCrlf
	strSQL = strSQL & " 		select top 1" & vbCrlf
	strSQL = strSQL & " 		d.itemcost, d.reducedPrice, d.mileage, d.currstate, d.songjangno, d.songjangdiv, d.buycash, d.itemvat, d.vatinclude, d.beasongdate" & vbCrlf
	strSQL = strSQL & " 		, d.isupchebeasong, d.omwdiv, d.odlvType, d.issailitem, d.upcheconfirmdate, d.oitemdiv, d.requiredetail, d.itemcouponidx, d.bonuscouponidx" & vbCrlf
	strSQL = strSQL & " 		, d.passday, d.orgitemcost, d.itemcostCouponNotApplied, d.buycashCouponNotApplied, d.odlvfixday, d.plusSaleDiscount, d.specialshopDiscount" & vbCrlf
	strSQL = strSQL & " 		, d.etcDiscount" & vbCrlf
	strSQL = strSQL & " 		from [db_order].[dbo].tbl_order_detail d" & vbCrlf
	strSQL = strSQL & " 		where 1 = 1 " & vbCrlf
	strSQL = strSQL & " 		and d.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 		and d.itemid = " & itemidfrom & " " & vbCrlf
	strSQL = strSQL & " 		and d.itemoption = '" & itemoptionfrom & "' " & vbCrlf
	strSQL = strSQL & " 	) b " & vbCrlf
	strSQL = strSQL & " where 1 = 1 " & vbCrlf
	strSQL = strSQL & " and a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " and a.itemid = " & itemidto & " " & vbCrlf
	strSQL = strSQL & " and a.itemoption = '" & itemoptionto & "' " & vbCrlf
	rsget.Open strSQL,dbget,1

	' 주문제작문구 입력		'2019.03.27 한용민
	strSQL = "if exists(" & VbCrlf
	strSQL = strSQL & " 	select top 1 isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail" & VbCrlf
	strSQL = strSQL & " 	from [db_order].[dbo].tbl_order_detail d" & vbCrlf
	strSQL = strSQL & " 	JOIN db_order.dbo.tbl_order_require dd" & vbCrlf
	strSQL = strSQL & " 		ON d.idx = dd.detailidx" & vbCrlf
	strSQL = strSQL & " 	where d.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 	and d.itemid = " & itemidto & " " & vbCrlf
	strSQL = strSQL & " 	and d.itemoption = '" & itemoptionto & "' " & vbCrlf
	strSQL = strSQL & " )" & VbCrlf
	strSQL = strSQL & "     begin" & VbCrlf
	strSQL = strSQL & " 	update b set b.requiredetailUTF8=t.requiredetail" & VbCrlf
	strSQL = strSQL & " 	from [db_order].[dbo].tbl_order_detail a" & vbCrlf
	strSQL = strSQL & " 	JOIN db_order.dbo.tbl_order_require b" & vbCrlf
	strSQL = strSQL & " 		ON a.idx = b.detailidx" & vbCrlf
	strSQL = strSQL & " 	, ( " & vbCrlf
	strSQL = strSQL & " 		select top 1" & vbCrlf
	strSQL = strSQL & " 		isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail" & vbCrlf
	strSQL = strSQL & " 		from [db_order].[dbo].tbl_order_detail d" & vbCrlf
	strSQL = strSQL & " 		left JOIN db_order.dbo.tbl_order_require dd" & vbCrlf
	strSQL = strSQL & " 			ON d.idx = dd.detailidx" & vbCrlf
	strSQL = strSQL & " 		where d.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 		and d.itemid = " & itemidfrom & " " & vbCrlf
	strSQL = strSQL & " 		and d.itemoption = '" & itemoptionfrom & "' " & vbCrlf
	strSQL = strSQL & " 	) t " & vbCrlf
	strSQL = strSQL & " 	where a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 	and a.itemid = " & itemidto & " " & vbCrlf
	strSQL = strSQL & " 	and a.itemoption = '" & itemoptionto & "' " & vbCrlf
	strSQL = strSQL & " 	and t.requiredetail is null" & vbCrlf		' 주문제작문구가 있는것만 없어침
	strSQL = strSQL & "     end" & VbCrlf
	strSQL = strSQL & " else" & VbCrlf
	strSQL = strSQL & "     begin" & VbCrlf
	strSQL = strSQL & "     insert into [db_order].[dbo].tbl_order_require (detailidx, requiredetailUTF8, regdate, lastupdate)" & VbCrlf
	strSQL = strSQL & " 		select top 1 a.idx, t.requiredetail, getdate(), getdate()" & vbCrlf
	strSQL = strSQL & " 		from [db_order].[dbo].tbl_order_detail a" & vbCrlf
	strSQL = strSQL & " 		left JOIN db_order.dbo.tbl_order_require b" & vbCrlf
	strSQL = strSQL & " 			ON a.idx = b.detailidx" & vbCrlf
	strSQL = strSQL & " 		, ( " & vbCrlf
	strSQL = strSQL & " 			select top 1" & vbCrlf
	strSQL = strSQL & " 			isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail" & vbCrlf
	strSQL = strSQL & " 			from [db_order].[dbo].tbl_order_detail d" & vbCrlf
	strSQL = strSQL & " 			left JOIN db_order.dbo.tbl_order_require dd" & vbCrlf
	strSQL = strSQL & " 				ON d.idx = dd.detailidx" & vbCrlf
	strSQL = strSQL & " 			where d.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 			and d.itemid = " & itemidfrom & " " & vbCrlf
	strSQL = strSQL & " 			and d.itemoption = '" & itemoptionfrom & "' " & vbCrlf
	strSQL = strSQL & " 		) t " & vbCrlf
	strSQL = strSQL & " 		where a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 		and a.itemid = " & itemidto & " " & vbCrlf
	strSQL = strSQL & " 		and a.itemoption = '" & itemoptionto & "' " & vbCrlf
	strSQL = strSQL & " 		and b.detailidx is null" & vbCrlf
	strSQL = strSQL & " 		and t.requiredetail is not null" & vbCrlf		' 주문제작문구가 없는건 제낌.
	strSQL = strSQL & "     end" & VbCrlf

	'response.write strSQL & "<br>"
	dbget.Execute strSQL
end function



'==============================================================================
'상품정보 복사(금액정보 제외)
'==============================================================================
function CSOrderCopyItemInfoPart(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto)
	dim strSQL

	strSQL = "update a " & vbCrlf
	strSQL = strSQL & " set " & vbCrlf
	strSQL = strSQL & " 	a.currstate = b.currstate " & vbCrlf
	strSQL = strSQL & " 	, a.songjangno = b.songjangno " & vbCrlf
	strSQL = strSQL & " 	, a.songjangdiv = b.songjangdiv " & vbCrlf
	strSQL = strSQL & " 	, a.beasongdate = b.beasongdate " & vbCrlf
	strSQL = strSQL & " 	, a.upcheconfirmdate = b.upcheconfirmdate " & vbCrlf
	strSQL = strSQL & " 	, a.passday = b.passday " & vbCrlf
	strSQL = strSQL & " 	, a.odlvfixday = b.odlvfixday " & vbCrlf
	strSQL = strSQL & " from " & vbCrlf
	strSQL = strSQL & " 	[db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " 	, ( " & vbCrlf
	strSQL = strSQL & " 		select top 1 * " & vbCrlf
	strSQL = strSQL & " 		from [db_order].[dbo].tbl_order_detail " & vbCrlf
	strSQL = strSQL & " 		where 1 = 1 " & vbCrlf
	strSQL = strSQL & " 		and orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 		and itemid = " & itemidfrom & " " & vbCrlf
	strSQL = strSQL & " 		and itemoption = '" & itemoptionfrom & "' " & vbCrlf
	strSQL = strSQL & " 	) b " & vbCrlf
	strSQL = strSQL & " where 1 = 1 " & vbCrlf
	strSQL = strSQL & " and a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " and a.itemid = " & itemidto & " " & vbCrlf
	strSQL = strSQL & " and a.itemoption = '" & itemoptionto & "' " & vbCrlf
	rsget.Open strSQL,dbget,1

end function

function CSOrderSetItemPriceInfo(byval orderserial, byval itemid, byval itemoption, byval SalePrice, byval ItemCouponPrice, byval BonusCouponPrice, byVal EtcDiscountPrice, byval buycash)
	dim strSQL

	strSQL = "update a " & vbCrlf
	strSQL = strSQL & " set " & vbCrlf
	strSQL = strSQL & " 	a.itemcostCouponNotApplied = " + CStr(SalePrice) + " " & vbCrlf
	strSQL = strSQL & " 	, a.itemcost = " + CStr(ItemCouponPrice) + " " & vbCrlf
	strSQL = strSQL & " 	, a.reducedPrice = " + CStr(EtcDiscountPrice) + " " & vbCrlf
	strSQL = strSQL & " 	, a.etcDiscount = " + CStr(BonusCouponPrice-EtcDiscountPrice) + " " & vbCrlf
	strSQL = strSQL & " 	, a.itemvat = Round((" + CStr(BonusCouponPrice) + "/11.0), 0) " & vbCrlf
	strSQL = strSQL & " 	, a.buycash = " + CStr(buycash) + " " & vbCrlf
	strSQL = strSQL & " from " & vbCrlf
	strSQL = strSQL & " 	[db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " where 1 = 1 " & vbCrlf
	strSQL = strSQL & " and a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " and a.itemid = " & itemid & " " & vbCrlf
	strSQL = strSQL & " and a.itemoption = '" & itemoption & "' " & vbCrlf
	rsget.Open strSQL,dbget,1

end function

function CSOrderSetItemCouponInfo(byval orderserial, byval itemid, byval itemoption, byval itemcouponidx)
	dim strSQL

	strSQL = "update a " & vbCrlf
	strSQL = strSQL & " set " & vbCrlf
	strSQL = strSQL & " 	a.itemcouponidx = " + CStr(itemcouponidx) + " " & vbCrlf
	strSQL = strSQL & " from " & vbCrlf
	strSQL = strSQL & " 	[db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " where 1 = 1 " & vbCrlf
	strSQL = strSQL & " and a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " and a.itemid = " & itemid & " " & vbCrlf
	strSQL = strSQL & " and a.itemoption = '" & itemoption & "' " & vbCrlf
	rsget.Open strSQL,dbget,1

end function

function CSOrderSetBonusCouponInfo(byval orderserial, byval itemid, byval itemoption, byval bonuscouponidx)
	dim strSQL

	strSQL = "update a " & vbCrlf
	strSQL = strSQL & " set " & vbCrlf
	strSQL = strSQL & " 	a.bonuscouponidx = " + CStr(bonuscouponidx) + " " & vbCrlf
	strSQL = strSQL & " from " & vbCrlf
	strSQL = strSQL & " 	[db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " where 1 = 1 " & vbCrlf
	strSQL = strSQL & " and a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " and a.itemid = " & itemid & " " & vbCrlf
	strSQL = strSQL & " and a.itemoption = '" & itemoption & "' " & vbCrlf
	rsget.Open strSQL,dbget,1

end function

function CSOrderCopyBonusCouponInfo(byval orderserial, byval itemidfrom, byval itemidto, byval itemoptionfrom, byval itemoptionto)
	dim strSQL

	strSQL = "update a " & vbCrlf
	strSQL = strSQL & " set " & vbCrlf
	strSQL = strSQL & " 	a.bonuscouponidx = b.bonuscouponidx " & vbCrlf
	strSQL = strSQL & " 	, a.etcDiscount = b.etcDiscount " & vbCrlf
	strSQL = strSQL & " from " & vbCrlf
	strSQL = strSQL & " 	[db_order].[dbo].tbl_order_detail a " & vbCrlf
	strSQL = strSQL & " 	, ( " & vbCrlf
	strSQL = strSQL & " 		select top 1 * " & vbCrlf
	strSQL = strSQL & " 		from [db_order].[dbo].tbl_order_detail " & vbCrlf
	strSQL = strSQL & " 		where 1 = 1 " & vbCrlf
	strSQL = strSQL & " 		and orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " 		and itemid = " & itemidfrom & " " & vbCrlf
	strSQL = strSQL & " 		and itemoption = '" & itemoptionfrom & "' " & vbCrlf
	strSQL = strSQL & " 	) b " & vbCrlf
	strSQL = strSQL & " where 1 = 1 " & vbCrlf
	strSQL = strSQL & " and a.orderserial = '" & orderserial & "' " & vbCrlf
	strSQL = strSQL & " and a.itemid = " & itemidto & " " & vbCrlf
	strSQL = strSQL & " and a.itemoption = '" & itemoptionto & "' " & vbCrlf
	rsget.Open strSQL,dbget,1

end function



'==============================================================================
'주문상세 부분취소 : 기존에 있던 상품 한가지 취소
'==============================================================================
function CSOrderCancelItemForce(byval orderserial, byval itemid, byval itemoption)

	dim strSQL, result, makeridforadd

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 취소할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "이미 취소된 상품입니다."
		exit function
	end if

	makeridforadd = CS_ORDER_ITEM_MAKERID

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	strSQL = "update	 [db_order].[dbo].tbl_order_detail"
	strSQL = strSQL & " set cancelyn='Y'" + vbCrlf
	strSQL = strSQL & " ,canceldate=IsNULL(canceldate,getdate())" + vbCrlf
	strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
	strSQL = strSQL & " and itemid=" & itemid  & vbCrlf
	strSQL = strSQL & " and itemoption='" & itemoption & "'" & vbCrlf
	rsget.Open strSQL,dbget,1


    '재고디비 반영 : 0인경우 원수량전체
    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & "0"
    dbget.Execute strSQL

	result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

end function



'==============================================================================
'주문상세 부분취소 정상화 : 기존에 있던 취소된 상품 한가지를 정상화
'==============================================================================
function CSOrderRestoreCanceledItemForce(byval orderserial, byval itemid, byval itemoption)

	dim strSQL, result, itemnoforadd, makeridforadd

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 처리할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "N" then
		CS_ORDER_FUNCTION_RESULT = "이미 정상 상품입니다."
		exit function
	end if

	itemnoforadd = CS_ORDER_ITEM_NO

	makeridforadd = CS_ORDER_ITEM_MAKERID

	ResetGlobalVarible()
	'--------------------------------------------------------------------------

	strSQL = "update	 [db_order].[dbo].tbl_order_detail "
	strSQL = strSQL & " set cancelyn='N'" + vbCrlf
	''strSQL = strSQL & " ,orderdate=getdate()" + vbCrlf
	strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
	strSQL = strSQL & " and itemid=" & itemid  & vbCrlf
	strSQL = strSQL & " and itemoption='" & itemoption & "'" & vbCrlf
	rsget.Open strSQL,dbget,1

    ''재고디비 반영 *-1
    '// itemnoforadd = itemnoforadd*-1
    '// strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr(itemnoforadd)
    strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr(itemnoforadd)
    'response.write "aaa" & strSQL
    dbget.Execute strSQL

    result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

end function




'==============================================================================
'주문상세 수량변경 : 기존에 있던 상품 한가지 수량변경
'==============================================================================
function CSOrderModifyItemNoForce(byval orderserial, byval itemid, byval itemoption, byval itemnoto)

	dim strSQL, result, makeridforadd, itemnofrom

	CS_ORDER_FUNCTION_RESULT = ""

	'--------------------------------------------------------------------------
	ResetGlobalVarible()

	if (itemid = "0") then
		CS_ORDER_FUNCTION_RESULT = "배송비는 취소할 수 없습니다."
		exit function
	end if

	result = CSOrderGetItemState(orderserial, itemid, itemoption)

	if IsNull(CS_ORDER_ITEM_CANCELYN) then
		CS_ORDER_FUNCTION_RESULT = "상품이 없습니다."
		exit function
	end if

	if CS_ORDER_ITEM_CANCELYN = "Y" then
		CS_ORDER_FUNCTION_RESULT = "취소된 상품입니다."
		exit function
	end if

	makeridforadd = CS_ORDER_ITEM_MAKERID
	itemnofrom = CS_ORDER_ITEM_NO

	ResetGlobalVarible()
	'--------------------------------------------------------------------------




	strSQL = "update	 [db_order].[dbo].tbl_order_detail " & vbCrlf
	strSQL = strSQL & " set itemno=" & itemnoto & vbCrlf
	'strSQL = strSQL & " ,orderdate=getdate()" & vbCrlf
	strSQL = strSQL & " where orderserial='" & orderserial & "'" & vbCrlf
	strSQL = strSQL & " and itemid=" & itemid & vbCrlf
	strSQL = strSQL & " and itemoption='" & itemoption & "'" & vbCrlf
	rsget.Open strSQL,dbget,1

	''재고디비 반영 변경수량
    '// itemnofrom = (itemnofrom-itemnoto)
    '// strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr(itemnofrom)
    if ((itemnofrom-itemnoto) > 0) then

    	strSQL = " exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr(itemnofrom-itemnoto)
    	'response.write "aa" & strSQL & itemnofrom & " --- " & itemnoto
    	dbget.Execute strSQL

    elseif ((itemnofrom-itemnoto) < 0) then

    	strSQL = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_regOrderPartial '" & orderserial & "'," & itemid & ",'" & itemoption & "'," & CStr((itemnofrom-itemnoto) * -1)
    	'response.write "aabb" & strSQL & itemnofrom & " --- " & itemnoto
    	dbget.Execute strSQL

    else

    	'// 수량변경 없으면 아무것도 않한다.

    end if


	result = CSOrderRecaculateBrandDeliveryPay(orderserial, makeridforadd)

end function



'==============================================================================
'업체 배송비 추가 삭제 : 업체조건배송인 경우, 금액에 따라 배송비를 추가하거나 삭제
'==============================================================================
function CSOrderRecaculateBrandDeliveryPay(byval orderserial, byval brandid)

	if (CStr(orderserial) <> "0") then
		'배송비 재계산 안한다. CS페이지에서와 같이 업체조건배송 고려안한다.
		exit function
	end if

	dim strSQL, result
	dim defaultfreebeasonglimit, defaultdeliverpay
	dim userlevel, jumundiv, DlvcountryCode, reducedprice

	strSQL = "select top 1 defaultfreebeasonglimit, defaultdeliverpay " + vbCrlf
	strSQL = strSQL & " from db_user.dbo.tbl_user_c " + vbCrlf
	strSQL = strSQL & " where 1 = 1 " + vbCrlf
	strSQL = strSQL & " and userid = '" & brandid & "' " + vbCrlf
	rsget.Open strSQL,dbget,1

	if Not rsget.Eof then
		defaultfreebeasonglimit = rsget("defaultfreebeasonglimit")
		defaultdeliverpay = rsget("defaultdeliverpay")
	end if
	rsget.close

	strSQL = "select top 1 IsNull(userlevel, 0) as userlevel, jumundiv, IsNull(DlvcountryCode, 'KR') as DlvcountryCode " + vbCrlf
	strSQL = strSQL & " from [db_order].[dbo].tbl_order_master " + vbCrlf
	strSQL = strSQL & " where 1 = 1 " + vbCrlf
	strSQL = strSQL & " and orderserial = '" & orderserial & "' " + vbCrlf
	rsget.Open strSQL,dbget,1

	if Not rsget.Eof then
		userlevel = rsget("userlevel")
		jumundiv = rsget("jumundiv")
		DlvcountryCode = rsget("DlvcountryCode")
	end if
	rsget.close

	if (CStr(jumundiv) = "5") then
		'외부몰 결제 : 배송비 입력 안함(정산시 입력)
		exit function
	end if

	if (DlvcountryCode <> "KR") then
		'해외, 군부대배송 : 재계산 않한다.
		exit function
	end if

	'==========================================================================
	'텐텐배송 고객회원등급에 따른 혜택적용(텐배상품기준)
	'7 : STAFF, 4 : GOLD, 3 : SILVER, 2 : BLUE
	'==========================================================================
	if (brandid = "") then
		defaultdeliverpay = getDefaultBeasongPayByDate(Left(Now, 10))
		if (CStr(userlevel) = "7") or (CStr(userlevel) = "4") then
			' 7, 4 : 무료배송
			defaultfreebeasonglimit = 0
			defaultdeliverpay = 0
		elseif (CStr(userlevel) = "3") then
			' 3 : 1만원이상 무료배송
			defaultfreebeasonglimit = 10000
		elseif (CStr(userlevel) = "2") then
			' 2 : 2만원이상 무료배송
			defaultfreebeasonglimit = 20000
		else
			defaultfreebeasonglimit = 30000
		end if
	end if



	if (defaultfreebeasonglimit = 0) then
		'업체무료배송
		exit function
	end if



	dim havedeliverypay, deliverypayyn, totalprice, maxdeliveryoption
	dim newtotalprice, tentotalprice, newtentotalprice, freebeasongitemexist

	'무료배송조건이 되는지 확인(기준금액 : 쿠폰적용안된 판매가(할인가))
	strSQL = "select " + vbCrlf
	strSQL = strSQL & " 	 sum(case when makerid = '" & brandid & "' and itemid = 0                      then 1 else 0 end) as havedeliverypay " + vbCrlf
	strSQL = strSQL & " 	,sum(case when makerid = '" & brandid & "' and itemid = 0 and cancelyn <> 'Y'  then 1 else 0 end) as deliverypayyn " + vbCrlf
	strSQL = strSQL & " 	,sum(case when makerid = '" & brandid & "' and itemid <> 0 and cancelyn <> 'Y' then (itemcost*itemno) else 0 end) as totalprice " + vbCrlf
	strSQL = strSQL & " 	,sum(case when makerid = '" & brandid & "' and itemid <> 0 and cancelyn <> 'Y' then (IsNull(itemcostCouponNotApplied,0)*itemno) else 0 end) as newtotalprice " + vbCrlf
	strSQL = strSQL & " 	,sum(case when isupchebeasong = 'N' and itemid <> 0 and cancelyn <> 'Y' then (itemcost*itemno) else 0 end) as tentotalprice " + vbCrlf
	strSQL = strSQL & " 	,sum(case when isupchebeasong = 'N' and itemid <> 0 and cancelyn <> 'Y' then (IsNull(itemcostCouponNotApplied,0)*itemno) else 0 end) as newtentotalprice " + vbCrlf
	strSQL = strSQL & " 	,max(case when itemid = 0 then itemoption else '9000' end) as maxdeliveryoption " + vbCrlf
	strSQL = strSQL & " 	,sum(case when makerid = '" & brandid & "' and itemid <> 0 and (odlvType in ('2', '4', '5')) then 1 else 0 end) as freebeasongitemexist " + vbCrlf

	strSQL = strSQL & " from [db_order].[dbo].tbl_order_detail " + vbCrlf
	strSQL = strSQL & " where 1 = 1 " + vbCrlf
	strSQL = strSQL & " and orderserial = '" & orderserial & "' " + vbCrlf
	rsget.Open strSQL,dbget,1

	if Not rsget.Eof then
		havedeliverypay = rsget("havedeliverypay")
		deliverypayyn = rsget("deliverypayyn")
		totalprice = rsget("totalprice")
		maxdeliveryoption = rsget("maxdeliveryoption")

		newtotalprice = rsget("newtotalprice")
		tentotalprice = rsget("tentotalprice")
		newtentotalprice = rsget("newtentotalprice")
		freebeasongitemexist = rsget("freebeasongitemexist")
	end if
	rsget.close

	'배송비 기준금액은 쿠폰적용안된 금액이다. 그러나 과거 내역은 쿠폰미적용금액이 없으므로 상품쿠폰 적용가를 기준으로 배송비를 넣는다.
	if (newtotalprice > 0) then
		totalprice = newtotalprice
	end if

	if (newtentotalprice > 0) then
		tentotalprice = newtentotalprice
	end if

	if (brandid = "") then
		totalprice = tentotalprice
	end if

	if ((totalprice >= defaultfreebeasonglimit) or (totalprice = 0) or (freebeasongitemexist > 0)) then

		'무료배송이면 배송비가 있는지 확인 후 취소
		if (deliverypayyn <> 0) then
			strSQL = "update	 [db_order].[dbo].tbl_order_detail"
			strSQL = strSQL & " set cancelyn='Y'" + vbCrlf
			strSQL = strSQL & " ,canceldate=IsNULL(canceldate,getdate())" + vbCrlf
			strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
			strSQL = strSQL & " and itemid=0" & vbCrlf
			strSQL = strSQL & " and makerid='" & brandid & "'" & vbCrlf
			rsget.Open strSQL,dbget,1
		end if

	else

		'무료배송이 아니면 배송비가 없는지 확인 후 추가
		if (havedeliverypay = 0) then
			'없으면 추가
			result = CSOrderAddNewDeliveryPay(orderserial, brandid, CStr(maxdeliveryoption + 1), defaultdeliverpay)
		end if

		if (havedeliverypay <> 0) and (deliverypayyn = 0) then
			'삭제된 내역이 있으면 정상화
			strSQL = "update	 [db_order].[dbo].tbl_order_detail "
			strSQL = strSQL & " set cancelyn='N'" + vbCrlf
			''strSQL = strSQL & " ,orderdate=getdate()" + vbCrlf
			strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
			strSQL = strSQL & " and itemid=0"  & vbCrlf
			strSQL = strSQL & " and makerid='" & brandid & "'" & vbCrlf
			rsget.Open strSQL,dbget,1
		end if

		if (havedeliverypay <> 0) then
			'삭제않된 내역이 있으면 금액확인
			strSQL = "select top 1 reducedprice " + vbCrlf
			strSQL = strSQL & " from [db_order].[dbo].tbl_order_detail " + vbCrlf
			strSQL = strSQL & " where 1 = 1 " + vbCrlf
			strSQL = strSQL & " and orderserial = '" & orderserial & "' " + vbCrlf
			strSQL = strSQL & " and makerid = '" & brandid & "' " + vbCrlf
			strSQL = strSQL & " and itemid=0"  & vbCrlf
			rsget.Open strSQL,dbget,1

			if Not rsget.Eof then
				reducedprice = rsget("reducedprice")
			end if
			rsget.close

			if (reducedprice = 0) then
				if (brandid = "") then
					'텐텐배송비
					strSQL = "update	 [db_order].[dbo].tbl_order_detail "
					strSQL = strSQL & " set itemcost = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , reducedprice = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , orgitemcost = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , itemcostCouponNotApplied = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
					strSQL = strSQL & " and itemid=0"  & vbCrlf
					strSQL = strSQL & " and makerid='" & brandid & "'" & vbCrlf
					rsget.Open strSQL,dbget,1
				else
					'업체배송
					strSQL = "update	 [db_order].[dbo].tbl_order_detail "
					strSQL = strSQL & " set itemcost = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , reducedprice = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , orgitemcost = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , itemcostCouponNotApplied = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , buycash = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " , buycashCouponNotApplied = " & CStr(defaultdeliverpay) & " " + vbCrlf
					strSQL = strSQL & " where orderserial='" & orderserial & "'" + vbCrlf
					strSQL = strSQL & " and itemid=0"  & vbCrlf
					strSQL = strSQL & " and makerid='" & brandid & "'" & vbCrlf
					rsget.Open strSQL,dbget,1
				end if
			end if

		end if

	end if

end function



'==============================================================================
'배송비 추가 : itemid, itemno 가 아니라 brandid, itemcost 이다.
'==============================================================================
function CSOrderAddNewDeliveryPay(byval orderserial, byval brandid, byval itemoption, byval itemcost)

	dim sqlStr, result
	dim iid
	dim ParticleBeasongCode

	sqlStr = "select top 1 idx from [db_order].[dbo].tbl_order_master " & vbCrlf
	sqlStr = sqlStr & " where orderserial='" + orderserial + "'" & vbCrlf
	rsget.Open sqlStr,dbget,1
	if rsget.Eof then
		iid = 0
	else
		iid = rsget("idx")
	end if
	rsget.Close

	if (brandid = "") then
		'텐텐배송비 입력
    	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
    	sqlStr = sqlStr & " (masteridx, orderserial, itemid, itemoption, makerid, itemno, itemname, itemoptionname,"
    	sqlStr = sqlStr & " itemcost, buycash, mileage, reducedprice, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, itemcouponidx, bonuscouponidx)" + vbCrlf
    	sqlStr = sqlStr & " values(" + CStr(iid)
    	sqlStr = sqlStr & " ,'" & orderserial & "'"
    	sqlStr = sqlStr & " , 0"
		sqlStr = sqlStr & " , '1000'"                           '''텐배송
    	sqlStr = sqlStr & " , ''"
    	sqlStr = sqlStr & " , 1"
    	sqlStr = sqlStr & " , '배송비'"                                  ''' 배송비 (명)
    	sqlStr = sqlStr & " , ''"
    	sqlStr = sqlStr & " , " & CStr(itemcost) & " "  				 ''' 상품쿠폰 적용금액(itemcost) : 기존
    	sqlStr = sqlStr & " , " & CStr(0)                                ''' 매입가
    	sqlStr = sqlStr & " , 0"
		sqlStr = sqlStr & " , " & CStr(itemcost) & " "					 ' 배송비
    	sqlStr = sqlStr & " , " & CStr(itemcost) & " "                   ''' 소비자가(orgitemcost)
    	sqlStr = sqlStr & " , " & CStr(itemcost) & " " 					 ''' 판매가 = 상품쿠폰 적용안한금액(itemcostCouponNotApplied)
    	sqlStr = sqlStr & " , 0 "                                		 ''' 매입가 (buycashCouponNotApplied)
		sqlStr = sqlStr & " , NULL"
		sqlStr = sqlStr & " , NULL"
    	sqlStr = sqlStr & ")"

    	dbget.Execute sqlStr
	else
		'업체배송비 입력
		ParticleBeasongCode = "9" & Left(CStr(itemoption), 3)

	    sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
    	sqlStr = sqlStr & " (masteridx, orderserial, itemid, itemoption, makerid, itemno, itemname, itemoptionname,"
    	sqlStr = sqlStr & " itemcost, buycash, mileage, reducedprice, orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, itemcouponidx, bonuscouponidx)" + vbCrlf
    	sqlStr = sqlStr & " values(" + CStr(iid)
    	sqlStr = sqlStr & " ,'" & orderserial & "'"
    	sqlStr = sqlStr & " , 0"
    	sqlStr = sqlStr & " , '" & ParticleBeasongCode & "'"
    	sqlStr = sqlStr & " , '" & brandid & "'"
    	sqlStr = sqlStr & " , 1"
    	sqlStr = sqlStr & " , '배송비'"
    	sqlStr = sqlStr & " , '업체개별'"                        '' or 업체 착불
    	sqlStr = sqlStr & " , " & CStr(itemcost)     ''  itemcost
    	sqlStr = sqlStr & " , " & CStr(itemcost)  ''  배송비 정산액
    	sqlStr = sqlStr & " , 0"                                 ''  마일리지
    	sqlStr = sqlStr & " , " & CStr(itemcost)     ''' 환불시 적용금액(reducedprice)
    	sqlStr = sqlStr & " , " & CStr(itemcost)      ''' 소비자가(orgitemcost)
	    sqlStr = sqlStr & " , " & CStr(itemcost)     ''' 상품쿠폰 적용안한금액(itemcostCouponNotApplied)  ''업체개별배송은 상품쿠폰 없음.
    	sqlStr = sqlStr & " , " & CStr(itemcost)  ''' 쿠폰 적용 안한 매입가.
    	sqlStr = sqlStr & " , NULL"         ''상품쿠폰번호(업체 조건배송인경우.. 추가작업 필요)
    	sqlStr = sqlStr & " , NULL"         ''보너스쿠폰번호(업체 조건배송은 없음)
    	sqlStr = sqlStr & " )"

    	dbget.Execute sqlStr
	end if

	'response.write "aaaaaaaaaaaaaaaa" & sqlStr

end function

%>
