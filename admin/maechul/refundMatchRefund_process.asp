<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제로그(주문별)
' Hieditor : 2011.04.22 이상구 생성
'			 2020.07.24 한용민 수정(결제로그매칭추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sqlStr, mode, orderserial, suborderserial, orgorderserial, chgorderserial, asid, paydivcode, payreqprice
dim matchorderserial, matchsuborderserial, pggubun, pgkey, pgcskey, appprice, appdate, addsuborderserial
dim ipkumdate, yyyymmdd, daypart, errMSG, startdate, enddate, dateGubun
	mode = requestCheckVar(request("mode"), 32)
	orderserial = requestCheckVar(request("orderserial"), 32)
	suborderserial = requestCheckVar(request("suborderserial"), 32)
	orgorderserial = requestCheckVar(request("orgorderserial"), 32)
	chgorderserial = requestCheckVar(request("chgorderserial"), 32)
	asid = requestCheckVar(request("asid"), 32)
	paydivcode = requestCheckVar(request("paydivcode"), 32)
	payreqprice = requestCheckVar(request("payreqprice"), 32)
	pggubun = requestCheckVar(request("pggubun"), 64)
	pgkey = requestCheckVar(request("pgkey"), 64)
	pgcskey = requestCheckVar(request("pgcskey"), 64)
	appprice = requestCheckVar(request("appprice"), 64)
	appdate = requestCheckVar(request("appdate"), 64)
	yyyymmdd = requestCheckVar(request("yyyymmdd"), 64)
	daypart = requestCheckVar(getNumeric(request("daypart")), 1)

dim refer
refer = request.ServerVariables("HTTP_REFERER")

if (mode = "matchRefundByPGdataOn") then

	'' 사용안함
	response.end

	sqlStr = " select top 1 a.orderserial, l.pgkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	left join db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.pgkey = l.pgkey "
	sqlStr = sqlStr + " 		and a.pgCSkey = l.pgCSkey "
	sqlStr = sqlStr + " 		and a.appPrice = l.realPayPrice "
	sqlStr = sqlStr + " 		and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "

	if (orgorderserial <> "") then
		sqlStr = sqlStr + " 	and a.orderserial = '" + CStr(orgorderserial) + "' "
	else
		sqlStr = sqlStr + " 	and a.orderserial = '" + CStr(orderserial) + "' "
	end if

	sqlStr = sqlStr + " 	and a.csasid = " + CStr(asid) + " "
	''response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		if Not IsNull(db3_rsget("pgkey")) then
			errMSG = "ERROR : 기매칭 승인내역(" + db3_rsget("pgkey") + ")"
		end if
	else
		errMSG = "ERROR : 승인정보 없음"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = a.appmethod "
	sqlStr = sqlStr + " 	, l.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	, l.PGuserid = a.PGuserid "
	sqlStr = sqlStr + " 	, l.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	, l.PGCSkey = a.PGCSkey "
	sqlStr = sqlStr + " 	, l.realPayPrice = a.appprice "
	sqlStr = sqlStr + " 	, l.payDate = a.appdate "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.suborderserial = '" + CStr(suborderserial) + "' "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "

	if (orgorderserial <> "") then
		sqlStr = sqlStr + " 	and a.orderserial = '" + CStr(orgorderserial) + "' "
	else
		sqlStr = sqlStr + " 	and a.orderserial = '" + CStr(orderserial) + "' "
	end if

	sqlStr = sqlStr + " 		and a.csasid = " + CStr(asid) + " "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchByPGdataOn") then

	sqlStr = " select top 1 a.orderserial, l.pgkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	left join db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.pgkey = l.pgkey "
	sqlStr = sqlStr + " 		and a.pgCSkey = l.pgCSkey "
	sqlStr = sqlStr + " 		and a.appPrice = l.realPayPrice "
	sqlStr = sqlStr + " 		and l.targetGbn = 'ON' "
	''sqlStr = sqlStr + " 		and l.orderserial = 'ON' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.pgkey = '" + CStr(pgkey) + "' "
	sqlStr = sqlStr + " 	and a.pgcskey = '" + CStr(pgcskey) + "' "
	sqlStr = sqlStr + " 	and a.appPrice = " + CStr(appprice) + " "

	''if (chgorderserial <> "") then
		''		sqlStr = sqlStr + " 	and a.orderserial in ('" + CStr(orgorderserial) + "', '" + CStr(chgorderserial) + "') "
	''	elseif (orgorderserial <> "") then
		''		sqlStr = sqlStr + " 	and a.orderserial = '" + CStr(orgorderserial) + "' "
	''	else
		''		sqlStr = sqlStr + " 	and a.orderserial = '" + CStr(orderserial) + "' "
	''	end if

	''response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		if Not IsNull(db3_rsget("pgkey")) then
			errMSG = "ERROR : 기매칭 승인내역(" + db3_rsget("pgkey") + ")"
		end if
	else
		errMSG = "ERROR : 승인정보 없음"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = a.appmethod "
	sqlStr = sqlStr + " 	, l.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	, l.PGuserid = a.PGuserid "
	sqlStr = sqlStr + " 	, l.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	, l.PGCSkey = a.PGCSkey "
	sqlStr = sqlStr + " 	, l.realPayPrice = a.appprice "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = a.commPrice "
	sqlStr = sqlStr + " 	, l.jungsanPrice = a.jungsanPrice "
	sqlStr = sqlStr + " 	, l.mayIpkumDate = a.ipkumdate " ''2013/01/06 추가
	sqlStr = sqlStr + " 	, l.maeipDate = a.pgmeachuldate " ''2013/01/06 추가
	sqlStr = sqlStr + " 	, l.paydate = (CASE WHEN a.appDivCode='A' THEN isNULL(a.appdate,a.canceldate) ELSE isNULL(a.canceldate,a.appdate) END )" ''2013/01/06 추가 //sqlStr = sqlStr + " 	, l.payDate = a.appdate "

	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.targetGbn in ('ON','AC') " ''2014/01/10 추가
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.suborderserial = '" + CStr(suborderserial) + "' "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "							'// 잘못 매칭된것 고칠수 있도록, skyer9, 2015-10-05
	sqlStr = sqlStr + " 		and a.pgkey = '" + CStr(pgkey) + "' "
	sqlStr = sqlStr + " 		and a.pgcskey = '" + CStr(pgcskey) + "' "
	sqlStr = sqlStr + " 		and a.appPrice = " + CStr(appprice) + " "
'response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchByPGdataOff") then

	sqlStr = " select top 1 IsNull(a.orderserial, '') as orderserial, IsNull(a.ipkumdate,'') as ipkumdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopjumun_cardApp_log a "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.pggubun = '" + CStr(pggubun) + "' "
	sqlStr = sqlStr + " 	and a.pgkey = '" + CStr(pgkey) + "' "
	response.write sqlStr

	errMSG = ""

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		if Len(rsget("ipkumdate")) <> 10 then
			errMSG = "ERROR : 입금예정일 오류(" + rsget("ipkumdate") + ")"
		elseif rsget("orderserial") <> "" and rsget("orderserial") <> (CStr(orderserial & "-" & Format00(3, suborderserial))) then
			errMSG = "ERROR : 기매칭 승인내역(" + rsget("orderserial") + ")"
		end If
		ipkumdate = rsget("ipkumdate")
	else
		errMSG = "ERROR : 승인정보 없음"
	end if
	rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		dbget.close()
		response.end
	end if

	sqlStr = " update db_shop.dbo.tbl_shopjumun_cardApp_log "
	sqlStr = sqlStr + " set orderserial = '" + CStr(orderserial & "-" & Format00(3, suborderserial)) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and pggubun = '" + CStr(pggubun) + "' "
	sqlStr = sqlStr + " 	and pgkey = '" + CStr(pgkey) + "' "
	''response.write sqlStr
	dbget.Execute sqlStr

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = '100' "		'// 신용카드
	sqlStr = sqlStr + " 	, l.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	, l.PGuserid = '' "
	sqlStr = sqlStr + " 	, l.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	, l.PGCSkey = '' "
	sqlStr = sqlStr + " 	, l.realPayPrice = a.cardPrice "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = a.cardChargePrice * -1 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = a.ipkumPrice "
	sqlStr = sqlStr + " 	, l.mayIpkumDate = '" & ipkumdate & "' " ''2013/01/06 추가
	''sqlStr = sqlStr + " 	, l.maeipDate = a.pgmeachuldate " ''2013/01/06 추가
	sqlStr = sqlStr + " 	, l.paydate = a.appdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopjumun_cardApp_log a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.targetGbn in ('ON','AC') "
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.suborderserial = '" + CStr(suborderserial) + "' "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 		and a.pggubun = '" + CStr(pggubun) + "' "
	sqlStr = sqlStr + " 		and a.pgkey = '" + CStr(pgkey) + "' "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchRefundByBankOn") then

	sqlStr = " select top 1 l.orderserial, l.pgkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " 	and l.PGCSkey = '" + CStr(asid) + "' "
	''response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		errMSG = "ERROR : 기매칭 환불내역(" + db3_rsget("pgkey") + ")"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " SELECT top 1 l.targetGbn "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and ( "
	sqlStr = sqlStr + " 			l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 			or "
	sqlStr = sqlStr + " 			IsNull(l.chgorderserial, l.orgorderserial) = a.orderserial "
	sqlStr = sqlStr + " 		) "
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 		and a.currstate = 'B007' "
	sqlStr = sqlStr + " 		and a.divcd = 'A003' "
	sqlStr = sqlStr + " 	join db_log.dbo.tbl_IBK_ERP_ICHE_DATA i "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and i.ten_csid = a.id "
	sqlStr = sqlStr + " 		and i.proc_yn = 'Y' "
	sqlStr = sqlStr + " 		and i.TEN_STATUS = 1 "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " + CStr(asid) + " "

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		'
	else
		errMSG = "ERROR : 환불내역 없음"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = '77' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'transfer' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'transfer' "
	sqlStr = sqlStr + " 	, l.PGkey = l.orderserial "
	sqlStr = sqlStr + " 	, l.PGCSkey = a.id "
	sqlStr = sqlStr + " 	, l.realPayPrice = i.tran_amt * -1 "
	sqlStr = sqlStr + " 	, l.payDate = CONVERT(DATETIME,STUFF(STUFF(STUFF(i.upd_date,9,0,' '),12,0,':'),15,0,':')) "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = 0 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = i.tran_amt * -1 "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and ( "
	sqlStr = sqlStr + " 			l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 			or "
	sqlStr = sqlStr + " 			IsNull(l.chgorderserial, l.orgorderserial) = a.orderserial "
	sqlStr = sqlStr + " 		) "
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 		and a.currstate = 'B007' "
	sqlStr = sqlStr + " 		and a.divcd = 'A003' "
	sqlStr = sqlStr + " 	join db_log.dbo.tbl_IBK_ERP_ICHE_DATA i "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and i.ten_csid = a.id "
	sqlStr = sqlStr + " 		and i.proc_yn = 'Y' "
	sqlStr = sqlStr + " 		and i.TEN_STATUS = 1 "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " + CStr(asid) + " "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchRefundByDepositOn") then

	sqlStr = " select top 1 l.orderserial, l.pgkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " 	and l.PGCSkey = '" + CStr(asid) + "' "
	''response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		errMSG = "ERROR : 기매칭 환불내역(" + db3_rsget("pgkey") + ")"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = 'rde' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'balance' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'balance' "
	sqlStr = sqlStr + " 	, l.PGkey = a.orderserial "
	sqlStr = sqlStr + " 	, l.PGCSkey = a.id "
	sqlStr = sqlStr + " 	, l.realPayPrice = " + CStr(appprice) + " * -1 "
	sqlStr = sqlStr + " 	, l.payDate = a.finishdate "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = 0 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = " + CStr(appprice) + " * -1 "
	sqlStr = sqlStr + " 	, l.mayIpkumDate = a.finishdate "
	sqlStr = sqlStr + " 	, l.maeipDate = a.finishdate "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and ( "
	sqlStr = sqlStr + " 			l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 			or "
	sqlStr = sqlStr + " 			IsNull(l.chgorderserial, l.orgorderserial) = a.orderserial "
	sqlStr = sqlStr + " 		) "
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.suborderserial = " + CStr(suborderserial) + " "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 		and a.currstate = 'B007' "
	sqlStr = sqlStr + " 		and a.divcd = 'A003' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " + CStr(asid) + " "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchRefundByReBankOn") then

	response.write "사용안함 : 무통장환불은 기존방식으로 매칭가능"
	response.end

	sqlStr = " select top 1 l.orderserial, l.pgkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " 	and l.PGCSkey = '" + CStr(asid) + "' "
	''response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		errMSG = "ERROR : 기매칭 환불내역(" + db3_rsget("pgkey") + ")"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = 'rde' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'balance' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'balance' "
	sqlStr = sqlStr + " 	, l.PGkey = a.orderserial "
	sqlStr = sqlStr + " 	, l.PGCSkey = a.id "
	sqlStr = sqlStr + " 	, l.realPayPrice = " + CStr(appprice) + " * -1 "
	sqlStr = sqlStr + " 	, l.payDate = a.finishdate "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = 0 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = " + CStr(appprice) + " * -1 "
	sqlStr = sqlStr + " 	, l.mayIpkumDate = a.finishdate "
	sqlStr = sqlStr + " 	, l.maeipDate = a.finishdate "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and ( "
	sqlStr = sqlStr + " 			l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 			or "
	sqlStr = sqlStr + " 			IsNull(l.chgorderserial, l.orgorderserial) = a.orderserial "
	sqlStr = sqlStr + " 		) "
	sqlStr = sqlStr + " 		and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 		and l.suborderserial = " + CStr(suborderserial) + " "
	sqlStr = sqlStr + " 		and l.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 		and a.currstate = 'B007' "
	sqlStr = sqlStr + " 		and a.divcd = 'A003' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.id = " + CStr(asid) + " "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchNoRefund") then

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = '0' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'nopayment' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'nopayment' "
	sqlStr = sqlStr + " 	, l.PGkey = l.orderserial "
	sqlStr = sqlStr + " 	, l.PGCSkey = " + CStr(asid) + " "
	sqlStr = sqlStr + " 	, l.realPayPrice = 0 "
	sqlStr = sqlStr + " 	, l.payDate = l.payReqDate "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = 0 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = 0 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " 	and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and l.suborderserial = '" + CStr(suborderserial) + "' "
	sqlStr = sqlStr + " 	and l.payDivCode = 'XXX' "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); opener.location.reload(); window.close();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchCancel") then

	sqlStr = " select top 1 p2.suborderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log p1 "
	sqlStr = sqlStr + " 	join db_datamart.dbo.tbl_order_payment_log p2 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and p1.orderserial = p2.orderserial "
	sqlStr = sqlStr + " 		and abs(p1.suborderserial - p2.suborderserial) = 1 "
	sqlStr = sqlStr + " 		and p1.payDivCode = p2.payDivCode "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p1.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and p1.suborderserial = " + CStr(suborderserial) + " "
	sqlStr = sqlStr + "     and p1.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 	and p1.payReqPrice = p2.payReqPrice*-1 "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	p2.suborderserial "
	response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
        matchsuborderserial = db3_rsget("suborderserial")
    else
        errMSG = "ERROR : 매칭오류"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = '0' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'nopayment' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'nopayment' "
	sqlStr = sqlStr + " 	, l.PGkey = l.orderserial "
	sqlStr = sqlStr + " 	, l.PGCSkey = l.suborderserial "
	sqlStr = sqlStr + " 	, l.realPayPrice = 0 "
	sqlStr = sqlStr + " 	, l.payDate = l.payReqDate "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = 0 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = 0 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	''sqlStr = sqlStr + " 	and l.targetGbn = 'AC' "
	sqlStr = sqlStr + " 	and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and l.suborderserial in ('" + CStr(suborderserial) + "', '" + CStr(matchsuborderserial) + "') "
	sqlStr = sqlStr + " 	and l.payDivCode = 'XXX' "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); history.back();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchReturn") then

	sqlStr = " select top 1 p1.orgorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log p1 "
	sqlStr = sqlStr + " 	join db_datamart.dbo.tbl_order_payment_log p2 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and p2.orderserial = p1.orgorderserial "
	sqlStr = sqlStr + " 		and p1.payDivCode = p2.payDivCode "
	sqlStr = sqlStr + " 		and p1.payReqPrice = p2.payReqPrice*-1 "
    sqlStr = sqlStr + "     	and p1.payDivCode = 'XXX' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p1.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and p1.suborderserial = " + CStr(suborderserial) + " "
    sqlStr = sqlStr + " 	and p2.suborderserial = 0 "
	''response.write sqlStr

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
        orgorderserial = db3_rsget("orgorderserial")
    else
        errMSG = "ERROR : 매칭오류"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = '0' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'nopayment' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'nopayment' "
	sqlStr = sqlStr + " 	, l.PGkey = l.orderserial "
	sqlStr = sqlStr + " 	, l.PGCSkey = l.suborderserial "
	sqlStr = sqlStr + " 	, l.realPayPrice = 0 "
	sqlStr = sqlStr + " 	, l.payDate = l.payReqDate "
	sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
	sqlStr = sqlStr + " 	, l.commPrice = 0 "
	sqlStr = sqlStr + " 	, l.jungsanPrice = 0 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and ((l.orderserial = '" + CStr(orderserial) + "' and l.suborderserial = '" + CStr(suborderserial) + "') or (l.orderserial = '" + CStr(orgorderserial) + "' and l.suborderserial = 0)) "
	sqlStr = sqlStr + " 	and l.payDivCode = 'XXX' "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); history.back();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "addpaylog") then

	'// 추가 결제로그
	addsuborderserial = 900

	sqlStr = " select top 1 l.orderserial, (l.suborderserial + 1) as suborderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and l.suborderserial >=  900 "
	sqlStr = sqlStr + " order by l.suborderserial desc "
	''response.write sqlStr

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		addsuborderserial = db3_rsget("suborderserial")
	end if
	db3_rsget.Close

	sqlStr = " insert into db_datamart.dbo.tbl_order_payment_log(orderserial, suborderserial, payDivCode, PGgubun, PGuserid, payReqPrice, payReqDate, targetGbn, matchMethod, orgorderserial, chgorderserial) "
	sqlStr = sqlStr + " select top 1 orderserial, " + CStr(addsuborderserial) + ", 'XXX', 'XXX', 'XXX', 0, payReqDate, targetGbn, 'X', orgorderserial, chgorderserial "
	sqlStr = sqlStr + " from db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and suborderserial = " + CStr(suborderserial) + " "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('추가되었습니다.'); history.back();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "matchRefundProc") then

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.matchMethod = 'R' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.targetGbn = 'ON' "
	sqlStr = sqlStr + " 	and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and l.suborderserial = '" + CStr(suborderserial) + "' "
	sqlStr = sqlStr + " 	and l.payDivCode = 'XXX' "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('수정되었습니다.'); history.back();</script>"
	db3_dbget.close()
	response.end

' 결제로그매칭하기
elseif (mode = "matchByDay") then
	if replace(yyyymmdd,"-","")="" or isnull(yyyymmdd) then
		response.write	"<script type='text/javascript'>"
		response.write	"	alert('구분자가 없습니다.');"
		response.write	"	location.replace('" + CStr(refer) + "');"
		response.write	"</script>"
	end if

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MatchOrderPaymentLogPGData_ON] '" & yyyymmdd & "', '" & yyyymmdd & "' "

	'response.write sqlStr & "<br>"
	db3_dbget.CommandTimeout = 60*5   ' 5분
	db3_dbget.execute sqlStr

	response.write	"<script type='text/javascript'>"
	response.write	"	alert('매칭되었습니다.');"
	response.write	"	location.replace('" + CStr(refer) + "');"
	response.write	"</script>"

' 결제로그매칭하기. 일괄
elseif (mode = "matchByDaydaypart") then
	if daypart="" or isnull(daypart) then
		response.write	"<script type='text/javascript'>"
		response.write	"	alert('구분자가 없습니다.');"
		response.write	"	location.replace('" + CStr(refer) + "');"
		response.write	"</script>"
	end if

	if daypart="1" then
		startdate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-01"
		enddate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-05"
	elseif daypart="2" then
		startdate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-06"
		enddate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-10"
	elseif daypart="3" then
		startdate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-11"
		enddate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-15"
	elseif daypart="4" then
		startdate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-16"
		enddate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-20"
	elseif daypart="5" then
		startdate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-21"
		enddate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-25"
	elseif daypart="6" then
		startdate = year(yyyymmdd) & "-" & Format00(2,month(yyyymmdd)) & "-26"
		enddate = dateadd("d",-1,dateadd("m",+1,DateSerial(year(yyyymmdd),Format00(2,month(yyyymmdd)),01)))
	end if

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MatchOrderPaymentLogPGData_ON] '" & startdate & "', '" & enddate & "' "

	'response.write sqlStr & "<br>"
	db3_dbget.CommandTimeout = 60*5   ' 5분
	db3_dbget.execute sqlStr

	response.write	"<script type='text/javascript'>"
	response.write	"	alert('매칭되었습니다.');"
	response.write	"	location.replace('" + CStr(refer) + "');"
	response.write	"</script>"

elseif (mode = "delmatch") then

	sqlStr = " update l "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	l.payDivCode = 'XXX' "
	sqlStr = sqlStr + " 	, l.PGgubun = 'XXX' "
	sqlStr = sqlStr + " 	, l.PGuserid = 'XXX' "
	sqlStr = sqlStr + " 	, l.PGkey = NULL "
	sqlStr = sqlStr + " 	, l.PGCSkey = NULL "
	sqlStr = sqlStr + " 	, l.realPayPrice = NULL "
	sqlStr = sqlStr + " 	, l.payDate = NULL "
	sqlStr = sqlStr + " 	, l.commPrice = NULL "
	sqlStr = sqlStr + " 	, l.jungsanPrice = NULL "
	sqlStr = sqlStr + " 	, l.maeipDate = NULL "
	sqlStr = sqlStr + " 	, l.mayIpkumDate = NULL "
	sqlStr = sqlStr + " 	, l.matchMethod = 'X' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.targetGbn in ('ON', 'AC') "
	sqlStr = sqlStr + " 	and l.orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and l.suborderserial = '" + CStr(suborderserial) + "' "
	sqlStr = sqlStr + " 	and l.payDivCode = '" + CStr(paydivcode) + "' "
	''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭삭제 되었습니다.'); history.back();</script>"
	db3_dbget.close()
	response.end

elseif (mode = "normalizematch") then

	sqlStr = " select top 1 orderserial, suborderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial = '" + CStr(orderserial) + "' "
	sqlStr = sqlStr + " 	and suborderserial < " + CStr(suborderserial) + " "
	sqlStr = sqlStr + " 	and payDivCode = 'XXX' "
	sqlStr = sqlStr + " 	and payReqPrice = " + CStr(payreqprice) + " * -1 "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	suborderserial desc "

	errMSG = ""

    db3_rsget.Open sqlStr,db3_dbget,1
	if Not(db3_rsget.EOF or db3_rsget.BOF) then
		''matchorderserial = db3_rsget("orderserial")
		matchsuborderserial = db3_rsget("suborderserial")
	else
		errMSG = "ERROR : 매칭주문 없음(잘못 매칭되어 있음)"
	end if
	db3_rsget.Close

	if (errMSG <> "") then
		response.write "<script>alert('" + errMSG + "');</script>"
		response.write errMSG
		db3_dbget.close()
		response.end
	else
		sqlStr = " update l "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	l.payDivCode = '0' "
		sqlStr = sqlStr + " 	, l.PGgubun = 'nopayment' "
		sqlStr = sqlStr + " 	, l.PGuserid = 'nopayment' "
		sqlStr = sqlStr + " 	, l.PGkey = l.orderserial "
		sqlStr = sqlStr + " 	, l.PGCSkey = (case when l.suborderserial = " + CStr(suborderserial) + " then " + CStr(matchsuborderserial) + " else " + CStr(suborderserial) + " end) "
		sqlStr = sqlStr + " 	, l.realPayPrice = 0 "
		sqlStr = sqlStr + " 	, l.payDate = l.payReqDate "
		sqlStr = sqlStr + " 	, l.matchMethod = 'H' "
		sqlStr = sqlStr + " 	, l.commPrice = 0 "
		sqlStr = sqlStr + " 	, l.jungsanPrice = 0 "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log l "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.targetGbn in ('ON', 'AC') "
		sqlStr = sqlStr + " 	and l.orderserial = '" + CStr(orderserial) + "' "
		sqlStr = sqlStr + " 	and l.suborderserial in ('" + CStr(suborderserial) + "', '" + CStr(matchsuborderserial) + "') "
		sqlStr = sqlStr + " 	and l.payDivCode = 'XXX' "
		''response.write sqlStr
		db3_dbget.Execute sqlStr

		response.write "<script>alert('매칭되었습니다.'); history.back();</script>"
		db3_dbget.close()
		response.end
	end if

elseif (mode = "setrefunding") then

    dateGubun = requestCheckVar(request("dateGubun"), 64)
    startdate = requestCheckVar(request("startDt"), 10)
    enddate = requestCheckVar(request("endDt"), 10)

	sqlStr = " update p "
	sqlStr = sqlStr + " set p.matchmethod = 'R' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_datamart].[dbo].[tbl_order_payment_log] p "
	sqlStr = sqlStr + " 	join [db_datamart].[dbo].[tbl_order_master_log] m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and p.orderserial = m.orderserial "
	sqlStr = sqlStr + " 		and p.suborderserial = m.suborderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "

    if (dateGubun = "paydate") then
	    sqlStr = sqlStr + " 	and p.paydate >= '" & startdate & "' "
	    sqlStr = sqlStr + " 	and p.paydate < '" & enddate & "' "
    elseif (dateGubun = "payreqdate") then
        sqlStr = sqlStr + " 	and p.payreqdate >= '" & startdate & "' "
	    sqlStr = sqlStr + " 	and p.payreqdate < '" & enddate & "' "
    else
        response.write "잘못된 접근입니다."
        response.end
    end if

	sqlStr = sqlStr + " 	and p.matchmethod = 'X' "
	sqlStr = sqlStr + " 	and m.actDivCode in ('C', 'M') "
	sqlStr = sqlStr + " 	and p.payReqPrice <> 0 "
    ''response.write sqlStr
	db3_dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); history.back();</script>"
	db3_dbget.close()
    response.end
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
