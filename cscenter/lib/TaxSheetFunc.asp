<%


function CheckTaxSheepWithOrderserial(orderserial, totalsum, byRef orderidx, byRef errMSG, byval nocheckVal)

	dim strSql
	dim ordertotalsum, ordertotalvat, orderipkumdiv, ordercancelyn

	orderidx = ""
	errMSG = ""

	if (orderserial = "수정세금계산서") then
		'// 체크 않함
		orderidx = "0"
		exit function
	end if

	if (orderserial = "") then
		errMSG = "주문번호가 입력되지 않았습니다."
		exit function
	end if

	strSql = "select top 1 idx, subtotalprice, totalvat, ipkumdiv, cancelyn, accountdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc " & VbCRLF
	strSql = strSql & "from [db_order].[dbo].tbl_order_master " & VbCRLF
	strSql = strSql & "where 1 = 1 " & VbCRLF
	strSql = strSql & "and orderserial = '" & CStr(orderserial) & "' " & VbCRLF
	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		orderidx = rsget("idx")

		if (TRIM(CStr(rsget("accountdiv"))) = "7") or (TRIM(CStr(rsget("accountdiv"))) = "20") then
			'무통장, 실시간이체 : 전체금액
			ordertotalsum = rsget("subtotalprice")
		else
			'보조결제금액만
			ordertotalsum = rsget("sumPaymentEtc")
		end if

		ordertotalvat = rsget("totalvat")
		orderipkumdiv = rsget("ipkumdiv")
		ordercancelyn = rsget("cancelyn")
	end if
	rsget.Close

	'' 과거 6개월 이전 내역 검색
	if (orderidx = "") and (Len(orderserial)=11) and (IsNumeric(orderserial)) then

		strSql =	"select top 1 idx, subtotalprice, totalvat, ipkumdiv, cancelyn, accountdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc " & VbCRLF
		strSql = strSql & "from [db_log].[dbo].tbl_old_order_master_2003 " & VbCRLF
		strSql = strSql & "where 1 = 1 " & VbCRLF
		strSql = strSql & "and orderserial = '" & CStr(orderserial) & "' " & VbCRLF
		rsget.Open strSql,dbget,1
		if  not rsget.EOF  then
			orderidx = rsget("idx")
			'//rw orderidx
			if (TRIM(CStr(rsget("accountdiv"))) = "7") or (TRIM(CStr(rsget("accountdiv"))) = "20") then
				'무통장, 실시간이체 : 전체금액
				ordertotalsum = rsget("subtotalprice")
				'//rw ordertotalsum
			else
				'보조결제금액만
				ordertotalsum = rsget("sumPaymentEtc")
			end if

			ordertotalvat = rsget("totalvat")
			orderipkumdiv = rsget("ipkumdiv")
			ordercancelyn = rsget("cancelyn")
		end if
		rsget.Close
		'//rw strSql
	end if

	if (orderidx = "") then
		errMSG = "잘못된 주문번호입니다."
		exit function
	end if

    if (nocheckVal="on") then

    else
    	if (CLng(totalsum) <> ordertotalsum) then
    		errMSG = "구매금액과 발행금액이 다릅니다." & totalsum & " : " & ordertotalsum
    		exit function
    	end if
    end if

	'if (orderipkumdiv < 4) then
	'	errMSG = "결재완료 이전 주문입니다."
	'	exit function
	'end if

	if (ordercancelyn = "Y") then
		errMSG = "취소된 주문입니다."
		exit function
	end if

	strSql = " select top 1 taxidx " & vbCrLf
	strSql = strSql & " from " & vbCrLf
	strSql = strSql & " db_order.dbo.tbl_taxSheet " & vbCrLf
	strSql = strSql & " where orderserial = '" & orderserial & "' and delyn <> 'Y' " & vbCrLf
	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		orderidx = 0
		errMSG = "이미 발행된 세금계산서가 존재합니다."
	end if
	rsget.Close

end function

function CheckTaxSheepWithChulgoCode(chulgocode, totalsum, byRef orderidx, byRef errMSG)

	dim strSql
	dim ordertotalsum, ordertotalvat, orderipkumdiv, ordercancelyn

	orderidx = 0
	errMSG = ""

	strSql =	"select top 1 id as idx, (IsNULL(totalsuplycash,0) * -1) as subtotalprice, (case when deldt is null then 'N' else 'Y' end) as cancelyn " & VbCRLF
	strSql = strSql & "from [db_storage].[dbo].tbl_acount_storage_master " & VbCRLF
	strSql = strSql & "where 1 = 1 " & VbCRLF
	strSql = strSql & "and code = '" & CStr(chulgocode) & "' " & VbCRLF
	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		orderidx = rsget("idx")
		ordertotalsum = rsget("subtotalprice")		'결재금액/출고가
		ordercancelyn = rsget("cancelyn")
	end if
	rsget.Close

	if (orderidx = 0) then
		errMSG = "NOT EXIST"
		exit function
	end if

	if (CLng(totalsum) <> ordertotalsum) then
		errMSG = "구매금액과 발행금액이 다릅니다." & totalsum & " : " & ordertotalsum
		exit function
	end if

	if (ordercancelyn = "Y") then
		errMSG = "취소된 출고내역입니다."
		exit function
	end if

end function

function CheckTaxSheepWithEtcMeachulCode(etcmeachulcode, totalsum, byRef errMSG)

	dim strSql
	dim ordertotalsum, ordertotalvat, orderipkumdiv, ordercancelyn
	dim IsMultiMeachulCode, arretcmeachulcode
	dim orderidx

	orderidx = ""
	errMSG = ""

	if (etcmeachulcode = "") then
		errMSG = "매출코드가 입력되지 않았습니다."
		exit function
	end if

	arretcmeachulcode = Split(etcmeachulcode, ",")
	IsMultiMeachulCode = (UBound(arretcmeachulcode) > 0)

	if IsMultiMeachulCode then
		strSql =	"select top 1 0 as idx, sum(totalsum) as totalsum " & VbCRLF
		strSql = strSql & "from db_shop.dbo.tbl_fran_meachuljungsan_master " & VbCRLF
		strSql = strSql & "where 1 = 1 " & VbCRLF
		strSql = strSql & "and idx in (" & CStr(etcmeachulcode) & ") " & VbCRLF
	else
		strSql =	"select top 1 idx, totalsum " & VbCRLF
		strSql = strSql & "from db_shop.dbo.tbl_fran_meachuljungsan_master " & VbCRLF
		strSql = strSql & "where 1 = 1 " & VbCRLF
		strSql = strSql & "and idx = '" & CStr(etcmeachulcode) & "' " & VbCRLF
	end if

	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		orderidx = rsget("idx")
		ordertotalsum = rsget("totalsum")
	end if
	rsget.Close

	if (orderidx = "") then
		errMSG = "매출내역이 없습니다."
		exit function
	end if

	if (CLng(totalsum) <> ordertotalsum) then
		errMSG = "매출금액과 발행금액이 다릅니다." & totalsum & " : " & ordertotalsum
		exit function
	end if

	strSql = " select top 1 t.taxidx " & vbCrLf
	strSql = strSql & " from " & vbCrLf
	strSql = strSql & " db_order.dbo.tbl_taxSheet t " & vbCrLf
	strSql = strSql & " left join db_order.dbo.tbl_taxSheet_Match m " & vbCrLf
	strSql = strSql & " on " & vbCrLf
	strSql = strSql & " 	t.taxidx = m.taxidx " & vbCrLf
	strSql = strSql & " where " & vbCrLf
	strSql = strSql & " 	1 = 1 " & vbCrLf
	strSql = strSql & " 	and t.delyn <> 'Y' " & vbCrLf
	strSql = strSql & " 	and ( " & vbCrLf
	strSql = strSql & " 		(t.orderserial = '' and t.orderidx in (" & CStr(etcmeachulcode) & ")) " & vbCrLf
	strSql = strSql & " 		or " & vbCrLf
	strSql = strSql & " 		(t.orderserial = '' and t.orderidx = 0 and m.matchtype = 'E' and m.useyn <> 'N' and m.matchlinkkey in (" & CStr(etcmeachulcode) & ")) " & vbCrLf
	strSql = strSql & " 	) " & vbCrLf
	''response.write strSql
	''response.end
	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		orderidx = 0
		errMSG = "이미 발행된 세금계산서가 존재합니다.(계산서IDX : " & rsget("taxidx") & ")"
	end if
	rsget.Close

end function

function AddTaxMasterInfo(userid, orderIdx, repName, repEmail, repTel, supplyBusiIdx, busiIdx, orderserial, itemname, totalPrice, totalTax, isuedate, billdiv, taxtype, sellBizCd, selltype, taxissuetype, consignYN, issueMethod)
	dim strSql

	AddTaxMasterInfo = -1

	strSql = " Insert into db_order.[dbo].tbl_taxSheet "
	strSql = strSql + " (userid, orderIdx, repName, repEmail, repTel, supplyBusiIdx, busiIdx, orderserial, itemname, totalPrice, totalTax, isuedate, billdiv, taxtype, sellBizCd, selltype, taxissuetype, consignYN, issueMethod) "
	strSql = strSql + " values('" + CStr(userid) + "' "
	strSql = strSql + " , " + CStr(orderIdx) + " "
	strSql = strSql + " , '" + CStr(repName) + "' "
	strSql = strSql + " , '" + CStr(repEmail) + "' "
	strSql = strSql + " , '" + CStr(repTel) + "' "

	strSql = strSql + " , '" + CStr(supplyBusiIdx) + "' "
	strSql = strSql + " , " + CStr(busiIdx) + " "

	strSql = strSql + " , '" + CStr(orderserial) + "' "
	strSql = strSql + " , '" + CStr(itemname) + "' "
	strSql = strSql + " , '" + CStr(totalPrice) + "' "
	strSql = strSql + " , '" + CStr(totalTax) + "' "
	strSql = strSql + " , '" + CStr(isuedate) + "' "
	strSql = strSql + " , '" + CStr(billdiv) + "' "
	strSql = strSql + " , '" + CStr(taxtype) + "' "
	strSql = strSql + " , '" + CStr(sellBizCd) + "' "
	strSql = strSql + " , '" + CStr(selltype) + "' "
	strSql = strSql + " , '" + CStr(taxissuetype) + "' "
	strSql = strSql + " , '" + CStr(consignYN) + "' "
	strSql = strSql + " , '" + CStr(issueMethod) + "' "
	strSql = strSql + " ) "
	''response.write strSql
	''response.end
	rsget.Open strSql, dbget, 1

	strSql = " select top 1 t.taxidx " & vbCrLf
	strSql = strSql & " from " & vbCrLf
	strSql = strSql & " db_order.dbo.tbl_taxSheet t " & vbCrLf
	strSql = strSql & " where " & vbCrLf
	strSql = strSql & " 	1 = 1 " & vbCrLf
	strSql = strSql & " 	and t.userid = '" + CStr(userid) + "' " & vbCrLf
	strSql = strSql & " order by t.taxidx desc " & vbCrLf

	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		AddTaxMasterInfo = rsget("taxidx")
	end if
	rsget.Close
end function

'// 공급업체 계산서 정보 있는지
function CheckSupplyTaxSheepInfoExists(taxIdx)
	dim strSql

	strSql = " select supplyBusiIdx " & VbCRLF
	strSql = strSql + " from db_order.[dbo].tbl_taxSheet " & VbCRLF
	strSql = strSql + " where taxIdx = " + CStr(taxIdx) + " and supplyBusiIdx is not NULL " & VbCRLF
	rsget.Open strSql, dbget, 1
	if  not rsget.EOF  then
		CheckSupplyTaxSheepInfoExists = True
	else
		CheckSupplyTaxSheepInfoExists = False
	end if
	rsget.close
end function

'// 공급업체 계산서IDX 입력
function UpdateSupplyTaxSheepBusiIdx(taxIdx, supplyBusiIdx)
	dim strSql

	strSql = " update db_order.[dbo].tbl_taxSheet " & VbCrLf
	strSql = strSql & " set supplyBusiIdx = '" & supplyBusiIdx & "' " & VbCrLf
	strSql = strSql & " where taxIdx = " & CStr(taxIdx) & " " & VbCrLf
	rsget.Open strSql, dbget, 1
end function

'// 공급받는 업체 계산서IDX 입력
function UpdateTaxSheepBusiIdx(taxIdx, busiIdx)
	dim strSql

	strSql = " update db_order.[dbo].tbl_taxSheet " & VbCrLf
	strSql = strSql & " set busiIdx = '" & busiIdx & "' " & VbCrLf
	strSql = strSql & " where taxIdx = " & CStr(taxIdx) & " " & VbCrLf
	rsget.Open strSql, dbget, 1
end function

'// 계산서 정보 입력
function AddTaxSheepInfo(userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn)
	dim strSql

	strSql = "Insert into db_order.[dbo].tbl_busiInfo " & VbCRLF
	strSql = strSql & "	(userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn) " & VbCRLF
	strSql = strSql & " values " & VbCRLF
	strSql = strSql & "	('" & userid & "','" & busiNo & "','" & busiSubNo & "'" & VbCRLF
	strSql = strSql & "	,'" & busiName & "'" & VbCRLF
	strSql = strSql & "	,'" & busiCEOName & "'" & VbCRLF
	strSql = strSql & "	,'" & busiAddr & "'" & VbCRLF
	strSql = strSql & "	,'" & busiType & "'" & VbCRLF              ''
	strSql = strSql & "	,'" & busiItem & "'" & VbCRLF               ''
	strSql = strSql & "	,'" & repName & "'" & VbCRLF
	strSql = strSql & "	,'" & repEmail & "'" & VbCRLF
	strSql = strSql & "	,'" & repTel & "'" & VbCRLF
	strSql = strSql & "	,'" & confirmYn & "')"
	rsget.Open strSql, dbget, 1

	strSql = " select top 1 busiIdx " & VbCRLF
	strSql = strSql + " from db_order.[dbo].tbl_busiinfo " & VbCRLF
	strSql = strSql + " where busiNo='" + busiNo + "' " & VbCRLF
	strSql = strSql + " order by busiIdx desc "
	rsget.Open strSql, dbget, 1
	if  not rsget.EOF  then
		AddTaxSheepInfo = rsget("busiIdx")
	else
		AddTaxSheepInfo = -1
	end if
	rsget.close
end function

'// 공급업체 계산서 정보 수정
function ModifySupplyTaxSheepInfo(taxIdx, userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn)
	dim strSql

	strSql = " update b "
	strSql = strSql + " set "
	strSql = strSql + " 	userid = '" + CStr(userid) + "' "
	strSql = strSql + " 	, busiNo = '" + CStr(busiNo) + "' "
	strSql = strSql + " 	, busiSubNo = '" + CStr(busiSubNo) + "' "
	strSql = strSql + " 	, busiName = '" + CStr(busiName) + "' "
	strSql = strSql + " 	, busiCEOName = '" + CStr(busiCEOName) + "' "
	strSql = strSql + " 	, busiAddr = '" + CStr(busiAddr) + "' "
	strSql = strSql + " 	, busiType = '" + CStr(busiType) + "' "
	strSql = strSql + " 	, busiItem = '" + CStr(busiItem) + "' "
	strSql = strSql + " 	, repName = '" + CStr(repName) + "' "
	strSql = strSql + " 	, repEmail = '" + CStr(repEmail) + "' "
	strSql = strSql + " 	, repTel = '" + CStr(repTel) + "' "
	strSql = strSql + " 	, confirmYn = '" + CStr(confirmYn) + "' "
	strSql = strSql + " from "
	strSql = strSql + " 	db_order.[dbo].tbl_taxSheet t "
	strSql = strSql + " 	join db_order.[dbo].tbl_busiInfo b "
	strSql = strSql + " 	on "
	strSql = strSql + " 		t.supplyBusiIdx = b.busiIdx "
	strSql = strSql + " where "
	strSql = strSql + " 	t.taxIdx = " + CStr(taxIdx) + " "
	rsget.Open strSql, dbget, 1
end function

'// 공급받는 업체 계산서 정보 수정
function ModifyTaxSheepInfo(taxIdx, userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn)
	dim strSql

	strSql = " update b "
	strSql = strSql + " set "
	strSql = strSql + " 	userid = '" + CStr(userid) + "' "
	strSql = strSql + " 	, busiNo = '" + CStr(busiNo) + "' "
	strSql = strSql + " 	, busiSubNo = '" + CStr(busiSubNo) + "' "
	strSql = strSql + " 	, busiName = '" + CStr(busiName) + "' "
	strSql = strSql + " 	, busiCEOName = '" + CStr(busiCEOName) + "' "
	strSql = strSql + " 	, busiAddr = '" + CStr(busiAddr) + "' "
	strSql = strSql + " 	, busiType = '" + CStr(busiType) + "' "
	strSql = strSql + " 	, busiItem = '" + CStr(busiItem) + "' "
	strSql = strSql + " 	, repName = '" + CStr(repName) + "' "
	strSql = strSql + " 	, repEmail = '" + CStr(repEmail) + "' "
	strSql = strSql + " 	, repTel = '" + CStr(repTel) + "' "
	strSql = strSql + " 	, confirmYn = '" + CStr(confirmYn) + "' "
	strSql = strSql + " from "
	strSql = strSql + " 	db_order.[dbo].tbl_taxSheet t "
	strSql = strSql + " 	join db_order.[dbo].tbl_busiInfo b "
	strSql = strSql + " 	on "
	strSql = strSql + " 		t.busiIdx = b.busiIdx "
	strSql = strSql + " where "
	strSql = strSql + " 	t.taxIdx = " + CStr(taxIdx) + " "
	rsget.Open strSql, dbget, 1
end function

function ModifyTaxMasterInfo(taxIdx, userid, orderIdx, repName, repEmail, repTel, supplyBusiIdx, busiIdx, orderserial, itemname, totalPrice, totalTax, isuedate, billdiv, taxtype, sellBizCd, selltype, taxissuetype, consignYN, issueMethod)
	dim strSql

	strSql = " update db_order.[dbo].tbl_taxSheet "
	strSql = strSql + " set "
	strSql = strSql + " 	userid = '" + CStr(userid) + "', "
	''strSql = strSql + " 	orderIdx = '" + CStr(orderIdx) + "', "
	strSql = strSql + " 	repName = '" + CStr(repName) + "', "
	strSql = strSql + " 	repEmail = '" + CStr(repEmail) + "', "
	strSql = strSql + " 	repTel = '" + CStr(repTel) + "', "
	''strSql = strSql + " 	orderserial = '" + CStr(orderserial) + "', "
	strSql = strSql + " 	itemname = '" + CStr(itemname) + "', "
	strSql = strSql + " 	totalPrice = '" + CStr(totalPrice) + "', "
	strSql = strSql + " 	totalTax = '" + CStr(totalTax) + "', "
	strSql = strSql + " 	isuedate = '" + CStr(isuedate) + "', "
	strSql = strSql + " 	billdiv = '" + CStr(billdiv) + "', "
	strSql = strSql + " 	taxtype = '" + CStr(taxtype) + "', "
	strSql = strSql + " 	sellBizCd = '" + CStr(sellBizCd) + "', "
	strSql = strSql + " 	selltype = '" + CStr(selltype) + "', "
	strSql = strSql + " 	taxissuetype = '" + CStr(taxissuetype) + "', "
	strSql = strSql + " 	consignYN = '" + CStr(consignYN) + "', "
	strSql = strSql + " 	issueMethod = '" + CStr(issueMethod) + "' "

	strSql = strSql + " where "
	strSql = strSql + " taxidx = " & taxidx

	''response.write strSql
	''response.end
	rsget.Open strSql, dbget, 1

end function

function GetTaxSheetIssueType(taxidx)
	dim strSql

	strSql =	" select top 1 t.taxidx, o.orderserial, a.code, e.idx, me.idx as midx " & VbCrLf
	strSql = strSql & " from " & VbCrLf
	strSql = strSql & " 	db_order.[dbo].tbl_taxSheet t " & VbCrLf
	strSql = strSql & " 	left join [db_order].[dbo].tbl_order_master o " & VbCrLf
	strSql = strSql & " 	on " & VbCrLf
	strSql = strSql & " 		t.orderidx = o.idx and t.orderserial = o.orderserial " & VbCrLf
	strSql = strSql & " 	left join [db_storage].[dbo].tbl_acount_storage_master a " & VbCrLf
	strSql = strSql & " 	on " & VbCrLf
	strSql = strSql & " 		t.orderidx = a.id and t.orderserial = a.code " & VbCrLf
	strSql = strSql & " 	left join db_shop.dbo.tbl_fran_meachuljungsan_master e " & VbCrLf
	strSql = strSql & " 	on " & VbCrLf
	strSql = strSql & " 		t.orderidx = e.idx and t.orderserial = '' " & VbCrLf
	strSql = strSql & " 	left join db_order.dbo.tbl_taxSheet_Match m " & VbCrLf
	strSql = strSql & " 	on " & VbCrLf
	strSql = strSql & " 		t.taxidx = m.taxidx and t.orderserial = '' and m.matchtype = 'E' " & VbCrLf
	strSql = strSql & " 	left join db_shop.dbo.tbl_fran_meachuljungsan_master me " & VbCrLf
	strSql = strSql & " 	on " & VbCrLf
	strSql = strSql & " 		m.matchlinkkey = me.idx " & VbCrLf
	strSql = strSql & " where " & VbCrLf
	strSql = strSql & " 	1 = 1 " & VbCrLf
	strSql = strSql & " 	and t.taxidx = " & taxidx & " " & VbCrLf

	rsget.Open strSql,dbget,1
	if  not rsget.EOF  then
		if (Not IsNull(rsget("orderserial"))) then
			'// 주문번호
			GetTaxSheetIssueType = "orderserial"
		elseif (Not IsNull(rsget("code"))) then
			'// 출고코드
			GetTaxSheetIssueType = "chulgocode"
		elseif (Not IsNull(rsget("idx"))) or (Not IsNull(rsget("midx"))) then
			'// 기타매출
			GetTaxSheetIssueType = "etcmeachul"
		else
			GetTaxSheetIssueType = "ETC"
		end if
	else
		GetTaxSheetIssueType = "ERROR"
	end if
	rsget.Close

end function

function SaveTaxSheepInfo(taxIdx, userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn)
	dim strSql

	'// 사용안함
	response.write "사용안함 : SaveTaxSheepInfo"
	response.end

	if (bisiIdx = "") then
		strSql =	"Insert into db_order.[dbo].tbl_busiInfo " & VbCRLF
		strSql = strSql & "	(userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn) " & VbCRLF
		strSql = strSql & " values " & VbCRLF
		strSql = strSql & "	('" & userid & "','" & busiNo & "','" & busiSubNo & "'" & VbCRLF
		strSql = strSql & "	,'" & busiName & "'" & VbCRLF
		strSql = strSql & "	,'" & busiCEOName & "'" & VbCRLF
		strSql = strSql & "	,'" & busiAddr & "'" & VbCRLF
		strSql = strSql & "	,'" & busiType & "'" & VbCRLF              ''
		strSql = strSql & "	,'" & busiItem & "'" & VbCRLF               ''
		strSql = strSql & "	,'" & repName & "'" & VbCRLF
		strSql = strSql & "	,'" & repEmail & "'" & VbCRLF
		strSql = strSql & "	,'" & repTel & "'" & VbCRLF
		strSql = strSql & "	,'" & confirmYn & "')"
	    rsget.Open strSql, dbget, 1

		strSql = " select top 1 busiIdx " & VbCRLF
		strSql = strSql + " from db_order.[dbo].tbl_busiinfo " & VbCRLF
		strSql = strSql + " where busiNo='" + busiNo + "' " & VbCRLF
		strSql = strSql + " order by busiIdx desc "
		rsget.Open strSql, dbget, 1
		if  not rsget.EOF  then
			bisiIdx = rsget("busiIdx")
		else
			bisiIdx = ""
		end if
		rsget.close
	else
		strSql =	" update db_order.[dbo].tbl_busiInfo " & VbCrLf
		strSql = strSql & " set " & VbCrLf
		strSql = strSql & " 	userid = '" & userid & "' " & VbCrLf
		strSql = strSql & " , busiNo = '" & busiNo & "' " & VbCrLf
		strSql = strSql & " , busiSubNo = '" & busiSubNo & "' " & VbCrLf
		strSql = strSql & " , busiName = '" & busiName & "' " & VbCrLf
		strSql = strSql & " , busiCEOName = '" & busiCEOName & "' " & VbCrLf
		strSql = strSql & " , busiAddr = '" & busiAddr & "' " & VbCrLf
		strSql = strSql & " , busiType = '" & busiType & "' " & VbCrLf
		strSql = strSql & " , busiItem = '" & busiItem & "' " & VbCrLf
		strSql = strSql & " , repName = '" & repName & "' " & VbCrLf
		strSql = strSql & " , repEmail = '" & repEmail & "' " & VbCrLf
		strSql = strSql & " , repTel = '" & repTel & "' " & VbCrLf
		strSql = strSql & " , confirmYn = '" & confirmYn & "' " & VbCrLf
		strSql = strSql & " where busiIdx = " & CStr(bisiIdx) & " " & VbCrLf
		rsget.Open strSql, dbget, 1
	end if

	SaveTaxSheepInfo = bisiIdx
end function








%>
