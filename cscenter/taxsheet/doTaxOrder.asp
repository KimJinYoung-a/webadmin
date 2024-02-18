<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/cscenter/lib/TaxSheetFunc.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetReqCls.asp"-->
<%

dim strSql, errMSG
dim i

dim mode
dim taxIdx, selltype, taxissuetype, sellBizCd
dim billdiv

dim reg_socno, reg_socname, reg_ceoname, reg_socaddr, reg_socstatus, reg_socevent, reg_subsocno, reg_managername, reg_managerphone, reg_managermail
dim socno, socname, ceoname, socaddr, socstatus, socevent, subsocno, managername, managerphone, managermail

dim reg_busiidx, busiidx

dim yyyymmdd, taxtype, totaltaxprice, totalprice
dim orderidx, itemname

dim userid
dim orderserial, consignYN, issueMethod, issueReqIdx

dim arrOrderIdx, arrOrderIdxString
dim IsMeachulCodeExist, IsMultiMeachulCode
dim taxSheetIssueType
dim oTax
dim nocheckVal

Dim oCTaxRequest
dim IsTaxIdxExist

dim tplcompanyid


mode = request("mode")
menupos = request("menupos")
taxIdx = request("taxIdx")
selltype = request("selltype")
taxissuetype = request("taxissuetype")
sellBizCd = request("sellBizCd")
billdiv = request("billdiv")

reg_socno = html2db(request("reg_socno"))
reg_socname = html2db(request("reg_socname"))
reg_ceoname = html2db(request("reg_ceoname"))
reg_socaddr = html2db(request("reg_socaddr"))
reg_socstatus = html2db(request("reg_socstatus"))
reg_socevent = html2db(request("reg_socevent"))
reg_subsocno = html2db(request("reg_subsocno"))
reg_managername = html2db(request("reg_managername"))
reg_managerphone = html2db(request("reg_managerphone"))
reg_managermail = html2db(request("reg_managermail"))

socno = html2db(request("socno"))
socname = html2db(request("socname"))
ceoname = html2db(request("ceoname"))
socaddr = html2db(request("socaddr"))
socstatus = html2db(request("socstatus"))
socevent = html2db(request("socevent"))
subsocno = html2db(request("subsocno"))
managername = html2db(request("managername"))
managerphone = html2db(request("managerphone"))
managermail = html2db(request("managermail"))

yyyymmdd = request("yyyymmdd")
taxtype = request("taxtype")
totaltaxprice = request("totaltaxprice")
totalprice = request("totalprice")
orderidx = request("orderidx")
itemname = html2db(request("itemname"))

consignYN = request("consignYN")
issueMethod = request("issueMethod")
issueReqIdx = request("issueReqIdx")

tplcompanyid = request("tplcompanyid")

userid = session("ssBctId")

''response.write consignYN
''response.write "taxissuetype : " & taxissuetype
''response.end

if (mode = "tax_register_new") then
	'// ------------------------------------------------------------------------
	'// 세금계산서 발행요청

	'// 공급자
	reg_busiidx = AddTaxSheepInfo(userid, reg_socno, reg_subsocno, reg_socname, reg_ceoname, reg_socaddr, reg_socstatus, reg_socevent, reg_managername, reg_managermail, reg_managerphone, "Y")

	'// 공급받는자
	busiidx = AddTaxSheepInfo(userid, socno, subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managermail, managerphone, "Y")

	if (taxissuetype = "E") then

		'// 기타매출
		Call CheckTaxSheepWithEtcMeachulCode(orderidx, totalprice, errMSG)

	elseif (taxissuetype = "C") then

		'// 온라인매출
		if (billdiv = "01") then
			'// 2013년 출고
			orderserial = orderidx

			Call CheckTaxSheepWithOrderserial(orderserial, totalprice, orderidx, errMSG, nocheckVal)  ''nocheckVal 추가 2013/10/29
		else
			'// 2014년 이후 출고
			response.write "작업중"
			response.end
		end if

	elseif (taxissuetype = "S") then

		'// 출고리스트
		response.write "<script>alert('기타매출을 등록하여 계산서를 발행하세요');  history.back();</script>"
		response.end

	elseif (taxissuetype = "X") then

		'// 내역없음
		orderidx = "0"
		orderserial = ""

	else
		response.write "<script>alert('잘못된 접근입니다.');  history.back();</script>"
		response.end
	end if

	if (errMSG <> "") then
		response.write "<script>alert('" & errMSG & "');</script>"
		response.write errMSG
		response.end
	end if

	IsMeachulCodeExist = False
	IsMultiMeachulCode = False

	if (orderIdx <> "") then
		IsMeachulCodeExist = True

		arrOrderIdx = Split(orderIdx, ",")
		IsMultiMeachulCode = (UBound(arrOrderIdx) > 0)

		if (IsMultiMeachulCode = True) then
			orderIdx = "0"
		end if

		arrOrderIdxString = orderIdx
	end if

	'// 마스터 정보
	taxIdx = AddTaxMasterInfo(userid, orderIdx, managername, managermail, managerphone, reg_busiidx, busiIdx, orderserial, itemname, totalprice, totaltaxprice, yyyymmdd, billdiv, taxtype, sellBizCd, selltype, taxissuetype, consignYN, issueMethod)

	if (tplcompanyid <> "") then
		strSql =	" Update db_order.[dbo].tbl_taxSheet Set " &_
					"	tplcompanyid = '" + CStr(tplcompanyid) + "' " &_
					" Where taxIdx=" & taxIdx
		dbget.Execute(strSql)
	end if

	if (IsMultiMeachulCode = True) then

		for i = 0 to UBound(arrOrderIdx)
			strSql = " insert into db_order.dbo.tbl_taxSheet_Match(taxIdx, matchtype, matchlinkkey, reguserid) " & VbCRLF
			strSql = strSql & " values( " & VbCRLF
			strSql = strSql & " '" & taxIdx & "', " & VbCRLF
			strSql = strSql & " '" & taxissuetype & "', " & VbCRLF
			strSql = strSql & " '" & arrOrderIdx(i) & "', " & VbCRLF
			strSql = strSql & " '" & userid & "' " & VbCRLF
			strSql = strSql & " ) " & VbCRLF
			rsget.Open strSql, dbget, 1

			arrOrderIdxString = arrOrderIdxString & "," & arrOrderIdx(i)
		next

	end if

	if taxissuetype = "C" then
		''플래그 업데이트
		if (billdiv = "01") then
			strSql = " update  [db_order].[dbo].tbl_order_master" & VbCRLF
			strSql = strSql & "set cashreceiptreq='T'" & VbCRLF
			strSql = strSql & " where orderserial='"& CStr(orderserial)& "'" & VbCRLF
			rsget.Open strSql, dbget, 1
		else
			'
		end if
	elseif (taxissuetype = "E") then
		strSql = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
		strSql = strSql + " set issuestatecd = '0' " + vbCrlf
		strSql = strSql + " where idx in (" & arrOrderIdxString & ")"
		rsget.Open strSql, dbget, 1
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); location.href = '/cscenter/taxsheet/tax_list.asp'; " &_
					"</script>"

elseif (mode = "tax_register_new_2014") then

	orderserial = orderidx

	'// 공급받는자
	busiidx = AddTaxSheepInfo(userid, socno, subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managermail, managerphone, "Y")

	strSql = " insert into db_log.[dbo].[tbl_tax_issue_request](orderserial, userid, busiIdx, reguserid) "
	strSql = strSql + " select m.orderserial, m.userid, " + CStr(busiidx) + ", '" + CStr(userid) + "' "
	strSql = strSql + " from db_order.dbo.tbl_order_master m "
	strSql = strSql + " 	left join db_log.[dbo].[tbl_tax_issue_request] r "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and m.orderserial = r.orderserial "
	strSql = strSql + " 		and r.useYN = 'Y' "
	strSql = strSql + " where m.orderserial = '" + CStr(orderserial) + "' and r.idx is NULL "
	''response.write strSql
	rsget.Open strSql, dbget, 1

	strSql = " update  [db_order].[dbo].tbl_order_master" & VbCRLF
	strSql = strSql & "set cashreceiptreq='T'" & VbCRLF
	strSql = strSql & " where orderserial='"& CStr(orderserial)& "'" & VbCRLF
	''response.write strSql
	rsget.Open strSql, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); location.href = '/cscenter/taxsheet/popCashReceipt.asp?orderserial=" + CStr(orderserial) + "'; " &_
					"</script>"

elseif (mode = "tax_register_new_2014_upche") then
	'// 세금계산서 발행요청(업체별)

	'// 공급자
	reg_busiidx = AddTaxSheepInfo(userid, reg_socno, reg_subsocno, reg_socname, reg_ceoname, reg_socaddr, reg_socstatus, reg_socevent, reg_managername, reg_managermail, reg_managerphone, "Y")

	'// 공급받는자
	busiidx = AddTaxSheepInfo(userid, socno, subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managermail, managerphone, "Y")

	orderserial = orderidx
	orderIdx = "0"

	set oCTaxRequest = new CTaxRequest
	oCTaxRequest.FRectOrderserial = orderSerial
	oCTaxRequest.FPageSize = 100
	oCTaxRequest.FRectOrderserial = orderSerial
	oCTaxRequest.GetTaxRequestOneOrder

	'// 발행요청 계산서 있는지
	IsTaxIdxExist = False

	for i = 0 to oCTaxRequest.FResultCount - 1
		if oCTaxRequest.FTaxList(i).FbusiNO = reg_socno then
			if CLng(oCTaxRequest.FTaxList(i).FtaxIdx) > 0 then
				IsTaxIdxExist = True
			end if
		end if
	next

	if (IsTaxIdxExist = True) then
		response.write "기발행요청된 계산서가 있습니다.(중복발행)"
		response.write	"<script language='javascript'>" &_
		"	alert('기발행요청된 계산서가 있습니다.(중복발행)'); " &_
		"</script>"
		response.end
	end if

	'// 마스터 정보
	taxIdx = AddTaxMasterInfo(userid, orderIdx, managername, managermail, managerphone, reg_busiidx, busiIdx, orderserial, itemname, totalprice, totaltaxprice, yyyymmdd, billdiv, taxtype, sellBizCd, selltype, taxissuetype, consignYN, issueMethod)

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); location.href = '/cscenter/taxsheet/tax_list.asp'; " &_
					"</script>"

elseif (mode = "delIssueReq") then

	orderserial = request("orderserial")

	set oCTaxRequest = new CTaxRequest
	oCTaxRequest.FRectUseYN = "Y"
	oCTaxRequest.FPageSize = 100
	oCTaxRequest.FRectOrderserial = orderSerial
	oCTaxRequest.GetTaxRequestOneOrder

	'// 발행요청 계산서 있는지
	IsTaxIdxExist = False

	for i = 0 to oCTaxRequest.FResultCount - 1
		if CLng(oCTaxRequest.FTaxList(i).FtaxIdx) > 0 then
			IsTaxIdxExist = True
		end if
	next

	if (IsTaxIdxExist = True) then
		response.write "발행요청된 계산서가 있습니다. 먼저 계산서를 삭제하세요."
		response.write	"<script language='javascript'>" &_
				"	alert('발행요청된 계산서가 있습니다. 먼저 계산서를 삭제하세요.'); " &_
				"</script>"
		response.end
	else
		strSql =	" Update db_log.[dbo].[tbl_tax_issue_request] Set " &_
					"	useYN = 'N' " &_
					" Where orderSerial = '" & orderSerial & "' and useYN = 'Y' "
		''response.write strSql
		dbget.Execute(strSql)

		strSql = " update [db_order].[dbo].tbl_order_master" & VbCrlf
		strSql = strSql & " set " & VbCrlf
		strSql = strSql & " authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " + VbCrlf
		strSql = strSql & " , cashreceiptreq = NULL " + VbCrlf
		strSql = strSql & " where orderserial='" & orderserial & "'"
		''response.write strSql
		dbget.Execute strSql

		response.write	"<script language='javascript'>" &_
						"	alert('삭제되었습니다.'); history.back(); " &_
						"</script>"
	end if

elseif (mode = "finishIssueReq") then

	orderserial = request("orderserial")

	strSql =	" Update db_log.[dbo].[tbl_tax_issue_request] Set " &_
				"	finishYN = 'Y' " &_
				" Where orderSerial = '" & orderSerial & "' and useYN = 'Y' "
	''response.write strSql
	dbget.Execute(strSql)

	response.write	"<script language='javascript'>" &_
					"	alert('완료처리 되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

elseif (mode = "sheetDel") then

	strSql =	" Update db_order.[dbo].tbl_taxSheet Set " &_
				"	delYn = 'Y' " &_
				" Where taxIdx=" & taxIdx
	dbget.Execute(strSql)


	set oTax = new CTax
	oTax.FRecttaxIdx = taxIdx

	oTax.GetTaxRead

	taxSheetIssueType = GetTaxSheetIssueType(taxIdx)

	if (taxSheetIssueType = "orderserial") then
		if (oTax.FOneItem.Fbilldiv = "01") then
			strSql = " update [db_order].[dbo].tbl_order_master" & VbCrlf
			strSql = strSql & " set " & VbCrlf
			strSql = strSql & " authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " + VbCrlf
			strSql = strSql & " , cashreceiptreq = NULL " + VbCrlf
			strSql = strSql & " where orderserial='" & oTax.FOneItem.Forderserial & "'"
			dbget.Execute strSql
		else
			'//
			''TODO : 작업중
		end if
	elseif (taxSheetIssueType = "etcmeachul") then
		strSql = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
		strSql = strSql + " set issuestatecd = NULL, neotaxno = NULL, taxlinkidx = NULL, eserotaxkey = NULL " + vbCrlf
		strSql = strSql + " where idx=" & oTax.FOneItem.Forderidx
		rsget.Open strSql, dbget, 1

		if CStr(oTax.FOneItem.Forderidx) = "0" then
			strSql = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
			strSql = strSql + " set issuestatecd = NULL, neotaxno = NULL, taxlinkidx = NULL, eserotaxkey = NULL " + vbCrlf
			strSql = strSql + " where idx in (select matchlinkkey from db_order.dbo.tbl_taxSheet_Match where matchtype = 'E' and taxidx = " & oTax.FOneItem.Ftaxidx & ") "
			rsget.Open strSql, dbget, 1
		end if
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('삭제되었습니다.'); location.href = '/cscenter/taxsheet/tax_list.asp'; " &_
					"</script>"

elseif (mode = "tax_modify") then

	'// 공급자
	Call ModifySupplyTaxSheepInfo(taxIdx, userid, reg_socno, reg_subsocno, reg_socname, reg_ceoname, reg_socaddr, reg_socstatus, reg_socevent, reg_managername, reg_managermail, reg_managerphone, "Y")

	'// 공급받는자
	Call ModifyTaxSheepInfo(taxIdx, userid, socno, subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managermail, managerphone, "Y")

	'// 마스터 정보
	Call ModifyTaxMasterInfo(taxIdx, userid, orderIdx, managername, managermail, managerphone, reg_busiidx, busiIdx, orderserial, itemname, totalprice, totaltaxprice, yyyymmdd, billdiv, taxtype, sellBizCd, selltype, taxissuetype, consignYN, issueMethod)

	response.write	"<script language='javascript'>" &_
					"	alert('수정되었습니다.'); history.back(); " &_
					"</script>"

elseif (mode = "finishSheetByESero") then

	strSql =	" Update db_order.[dbo].tbl_taxSheet Set " &_
				"	isueYN = 'Y' " &_
				" Where taxIdx=" & taxIdx
	dbget.Execute(strSql)

	response.write	"<script language='javascript'>" &_
					"	alert('발행완료 전환되었습니다.'); location.href = '/cscenter/taxsheet/tax_list.asp'; " &_
					"</script>"

else

	response.write	"<script language='javascript'>" &_
					"	alert('잘못된 접근입니다. - " & mode & "'); history.back(); " &_
					"</script>"
	response.end

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
