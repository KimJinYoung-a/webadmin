<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%

Dim StopWatch(19)

sub StartTimer(x)
	StopWatch(x) = Timer
end Sub

function StopTimer(x)
	Dim EndTime

	EndTime = Timer

	if EndTime < StopWatch(x) Then
		EndTime = EndTime + (86400)
	end if

	StopTimer = EndTime - StopWatch(x)
end function

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, reguserid
dim logidx, orderno, matchidx
dim IsMatched

dim objData, objXML, xmlURL, objLine
dim PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate
dim appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid
dim lastipkumdate, searchipkumdate
dim prevPGkey, prevPrevPGkey, prevAppDivCode, prevPrevAppDivCode, IsDuplicate
dim yyyymmdd
dim subPgKey
dim tmpStr, arrOrderSerial, orderserial
dim reasonGubun
dim searchipkumdateMAX

dim objFSO, objOpenedFile
Dim targetFileName
Const ForReading = 1

dim yyyymm


mode = requestCheckvar(request("mode"),64)
logidx = requestCheckvar(request("logidx"),32)
orderno = requestCheckvar(request("orderno"),32)
yyyymmdd = requestCheckvar(request("yyyymmdd"),32)
reasonGubun = requestCheckvar(request("reasonGubun"),32)

yyyymm = requestCheckvar(request("yyyymm"),7)

reguserid = session("ssBctId")

dim sqlStr
dim i, j, k


if (mode="matchoneorder") then

    sqlStr = " select isNULL(orderserial,'') as orderserial " & VbCRLF
    sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_cardApp_log " & VbCRLF
    sqlStr = sqlStr & " where idx="&logidx&VbCRLF

	IsMatched = True

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    IsMatched = Not (rsget("orderserial") = "")
	end if
	rsget.Close

	if IsMatched then
		response.write "<script>alert('이미 매칭된 내역입니다.');</script>"
		response.write "이미 매칭된 내역입니다."
		dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.shopJumunMasterIdx = m.idx, l.orderserial = m.orderno, l.shopid = m.shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
	sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_master m "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and m.orderno = '" + CStr(orderno) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and l.shopJumunMasterIdx is NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="matchsumipkum") then

	arrOrderSerial = Split(requestCheckvar(request("arrOrderSerial"),512), vbCrLf)
	tmpStr = ""

	for each orderserial in arrOrderSerial
		if (Len(orderserial) > 0) then
			if (Len(orderserial) <> 11) then
				response.write "<script>alert('잘못된 주문번호입니다.');</script>"
				response.write "잘못된 주문번호입니다." & orderserial
				dbget.close()
				response.end
			end if

			if (tmpStr = "") then
				tmpStr = " select '" + CStr(orderserial) + "' as orderserial " & vbCrLf
			else
				tmpStr = tmpStr + " union all " & vbCrLf & " select '" + CStr(orderserial) + "' as orderserial " & vbCrLf
			end if
		end if
	next

	if (tmpStr = "") then
		response.write "<script>alert('입력된 주문번호가 없습니다.');</script>"
		response.write "입력된 주문번호가 없습니다."
		dbget.close()
		response.end
	end if

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, PGuserid, orgPayDate, PGmeachulDate) " & vbCrLf
	sqlStr = sqlStr + " select l.PGgubun, l.PGkey, T.orderserial, l.sitename, l.appDivCode, l.appMethod, l.appDate, l.cancelDate, 0, 0, 0, 0, l.ipkumdate, T.orderserial, l.PGuserid, l.orgPayDate, l.PGmeachulDate " & vbCrLf
	sqlStr = sqlStr + " from " & vbCrLf
	sqlStr = sqlStr + "	db_order.dbo.tbl_onlineApp_log l " & vbCrLf
	sqlStr = sqlStr + "	join ( " & vbCrLf

	sqlStr = sqlStr + tmpStr

	sqlStr = sqlStr + "	) T " & vbCrLf
	sqlStr = sqlStr + "	on " & vbCrLf
	sqlStr = sqlStr + "		1 = 1 " & vbCrLf
	sqlStr = sqlStr + "	left join db_order.dbo.tbl_onlineApp_log l2 " & vbCrLf
	sqlStr = sqlStr + "	on " & vbCrLf
	sqlStr = sqlStr + "		1 = 1 " & vbCrLf
	sqlStr = sqlStr + "		and l.pggubun = l2.pggubun " & vbCrLf
	sqlStr = sqlStr + "		and l.pgkey = l2.pgkey " & vbCrLf
	sqlStr = sqlStr + "		and T.orderserial = l2.pgcskey " & vbCrLf
	sqlStr = sqlStr + "where " & vbCrLf
	sqlStr = sqlStr + "	1 = 1 " & vbCrLf
	sqlStr = sqlStr + "	and l.pggubun = 'bankipkum' " & vbCrLf
	sqlStr = sqlStr + "	and l.appDivCode = 'A' " & vbCrLf
	sqlStr = sqlStr + "	and l.idx = " + CStr(logidx) + " " & vbCrLf
	''sqlStr = sqlStr + "	and l.PGCSkey = '' " & vbCrLf
	sqlStr = sqlStr + "	and l2.idx is NULL " & vbCrLf
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="regReasonGubun") then

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set reasonGubun = '" + CStr(reasonGubun) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and IsNull(reasonGubun, '') not in ('030', '950') "
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="delmatchone") then

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.shopJumunMasterIdx = NULL, l.orderserial = NULL "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	''sqlStr = sqlStr + " 	and l.shopJumunMasterIdx is not NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('삭제되었습니다.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancel") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopjumun_cardApp_log c "
	sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopjumun_cardApp_log a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.cardAppNo = a.cardAppNo "
	sqlStr = sqlStr + " 		and convert(VARCHAR(10), c.appDate, 127) = convert(VARCHAR(10), a.appDate, 127) "
	''sqlStr = sqlStr + " 		and c.shopid = a.shopid "
	sqlStr = sqlStr + " 		and ((c.shopid = a.shopid) or (a.shopid is NULL and c.cardReaderID = a.cardReaderID)) "
	sqlStr = sqlStr + " 		and c.cardPrice*-1 = a.cardPrice "
	sqlStr = sqlStr + " 		and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and a.orderserial is NULL "
	sqlStr = sqlStr + " 	and c.orderserial is NULL "

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		response.write "<script>alert('에러!!\n\n매칭내역이 없습니다.');</script>"
		response.write "매칭내역이 없습니다."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_shop.dbo.tbl_shopjumun_cardApp_log "
	sqlStr = sqlStr + " set shopJumunMasterIdx = -1, orderserial = '취소매칭' "
	sqlStr = sqlStr + " where idx in (" + CStr(logidx) + ", " + CStr(matchidx) + ") "
	dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="addActLog") then

	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.pggubun = a.pggubun "
	sqlStr = sqlStr + " 	and o.pgkey = a.pgkey "
	sqlStr = sqlStr + " 	and o.pgcskey = Left(a.pgcskey, len(o.pgcskey)) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and o.idx = " + CStr(logidx) + " "

	PGCSkey = ""

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    PGCSkey = "-" + Format00(3, rsget("cnt"))
	end if
	rsget.Close

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select top 1 t.PGgubun, t.PGkey, t.PGCSkey + '" + CStr(PGCSkey) + "', t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, 0, 0, 0, 0, t.ipkumdate, t.PGuserid, t.PGmeachulDate, t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and t.appPrice <> 0 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	response.write "<script>alert('추가되었습니다.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancelOnline") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
	sqlStr = sqlStr + " 	and (convert(VARCHAR(10), IsNull(c.appDate,c.cancelDate), 127) = convert(VARCHAR(10), a.appDate, 127) or a.pggubun = 'bankipkum') "
	sqlStr = sqlStr + " 	and IsNull(c.sitename, '') = IsNull(a.sitename, '') "
	sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
	sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " and a.orderserial is NULL "
	sqlStr = sqlStr + " and c.orderserial is NULL "
	''rw sqlStr

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		response.write "<script>alert('에러!!\n\n매칭내역이 없습니다.');</script>"
		response.write "매칭내역이 없습니다."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set orderserial = '취소매칭' "
	sqlStr = sqlStr + " where idx in (" + CStr(logidx) + ", " + CStr(matchidx) + ") "
	dbget.Execute sqlStr

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set csasid = -1 "
	sqlStr = sqlStr + " where idx = " + CStr(logidx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancelOnlineDup") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
	''sqlStr = sqlStr + " 	and convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127) = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " 	and abs(datediff(d, convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127), convert(VARCHAR(10), a.appDate, 127))) <= 1 "
	sqlStr = sqlStr + " 	and c.sitename = a.sitename "
	sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
	sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " and c.csasid is NULL "
	sqlStr = sqlStr + " and c.orderserial is not NULL "
	sqlStr = sqlStr + " and c.idx > a.idx "
	''response.write sqlStr

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		'// 주문번호 있는 경우
		sqlStr = " select top 1 a.idx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
		sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
		sqlStr = sqlStr + " 	and abs(datediff(d, convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127), convert(VARCHAR(10), a.appDate, 127))) <= 10 "
		sqlStr = sqlStr + " 	and c.sitename = a.sitename "
		sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
		sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
		sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 1 = 1 "
		sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
		sqlStr = sqlStr + " and c.csasid is NULL "
		sqlStr = sqlStr + " and c.orderserial is not NULL "
		sqlStr = sqlStr + " and c.idx > a.idx "
		sqlStr = sqlStr + " and a.orderserial = c.orderserial "

		''response.write sqlStr
		matchidx = -1

		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			matchidx = rsget("idx")
		end if
		rsget.Close
	end if

	if matchidx = -1 then
		response.write "<script>alert('에러!!\n\n매칭내역이 없습니다.');</script>"
		response.write "매칭내역이 없습니다."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set csasid = -1, reasonGubun = NULL "
	sqlStr = sqlStr + " where idx = " + CStr(logidx) + " "
	dbget.Execute sqlStr

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set reasonGubun = NULL "
	sqlStr = sqlStr + " where idx = " + CStr(matchidx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('매칭되었습니다.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="getonpgdata") then

	'// ========================================================================
	'// INICIS
	if (yyyymmdd = "") then
		searchipkumdateMAX = ""
		sqlStr = " exec [db_cs].[dbo].[usp_getDayPlusWorkday] '" & Left(now(), 10) & "', 4 " & VbCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly
		if Not rsget.Eof then
			'// 근무일수 기준 D+4 일
			searchipkumdateMAX = rsget("plusworkday")
		end if
		rsget.close

		lastipkumdate = searchipkumdateMAX

		searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)))), 10)
		ipkumdate = Replace(searchipkumdate, "-", "")

		'// ========================================================================
		'// 온라인 텐텐 정산내역
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=teenxteen4&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=32"
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt04GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
		xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt14GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
		''response.write xmlURL

		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 45 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		end if

		''response.write objXML.Status

		Set objXML  = Nothing

		if (InStr(objData, "NO DATA") > 0) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('가져올 데이타가 없습니다.[1]');</script>"
			end if
			response.write "가져올 데이타가 없습니다[1]"
			dbget.close()
			response.end
		end if

		lastipkumdate = searchipkumdate
	else
		lastipkumdate = yyyymmdd

		searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)))), 10)

		ipkumdate = Replace(searchipkumdate, "-", "")

		'// ========================================================================
		'// 온라인 텐텐 정산내역
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=teenxteen4&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=32"
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt04GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
		xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt14GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"

		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 45 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		end if

		Set objXML  = Nothing

		if (InStr(objData, "NO DATA") > 0) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('가져올 데이타가 없습니다.[1]');</script>"
			end if
			response.write "가져올 데이타가 없습니다[1]"
			response.write objData
			response.end
		end if
	end if
	'' response.write objData
	'' response.end

	objData = Split(objData, "<br>")

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, "|")

		if (objLine(0) = "B") then

			PGgubun			= "inicis"

			PGuserid = objLine(4)

			if (objLine(4) = "teenxteen3") then
				sitename = "fingers"
			elseif (objLine(4) = "teenxteen4") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen5") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen6") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen8") then
				sitename = "10x10gift"
			elseif (objLine(4) = "teenxteen9") then
				sitename = "10x10mobile"
			else
				sitename = "XXX"
			end if

			if (objLine(11) = "A") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "A"
				PGCSkey		= ""

				appDate			= objLine(12)
				cancelDate		= "NULL"
			elseif (objLine(11) = "C") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "C"
				PGCSkey		= "CANCELALL"

				appDate			= objLine(12)
				cancelDate		= objLine(13)
			elseif (objLine(11) = "P") then
				'// ==============================
				'// 부분취소
				PGkey		= objLine(9)
				appDivCode	= "R"
				PGCSkey		= objLine(8)

				appDate			= "NULL"
				cancelDate		= objLine(13)
			else
				'// ==============================
				PGkey		= objLine(8)
				appDivCode = "E"
				PGCSkey		= "ERROR"
			end if

			''appMethod		= objLine(3)

			if (objLine(3) = "CC") then
				appMethod = "100"
			elseif (objLine(3) = "AC") then
				appMethod = "20"
			elseif (objLine(3) = "VA") then
				appMethod = "7"
			else
				appMethod = objLine(3)
			end if

			appPrice		= objLine(16)
			commPrice		= objLine(17)
			commVatPrice	= objLine(18)
			jungsanPrice	= objLine(20)

			ipkumdate		= objLine(5)

			'// 20130503000623
			'// (2013-05-03 00:06:23)
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if

			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end if

			'// 20130510
			'// (2013-05-10)
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

			sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
			sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			''response.write sqlStr + "<br>"
			dbget.execute sqlStr

		end if
	next

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'inicis' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('입금일자 : " + CStr(searchipkumdate) + "');</script>"
	end if

elseif (mode="getonpgdatahppre") then

	'// ========================================================================
	'// INICIS 핸드폰(기초작업)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('OK');</script>"
	end If

elseif (mode="getonpgdatahp") Then

	Call StartTimer(0)

	'// ========================================================================
	'// INICIS 핸드폰
	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateSerial(Year(Now()), Month(Now())+2, Day(Now()) - 2), 10)
	end If

	ipkumdate = Replace(yyyymmdd, "-", "")

	xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc/UrlSendExtraDc.jsp?urlid=teenteen10&passwd=cube1010??&date=" & ipkumdate & "&flgdate=P"

	objData = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

	objXML.setTimeouts 5 * 1000, 5 * 1000, 90 * 1000, 90 * 1000
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
	end if

	Set objXML  = Nothing

	if (InStr(objData, "NO DATA") > 0) then
		if  (Not IsAutoScript) then
			response.write "<script>alert('가져올 데이타가 없습니다.[1]');</script>"
		end if
		response.write "가져올 데이타가 없습니다[1]"
		response.write objData
		response.end
	end if

	''Response.Write "Elapsed time was: " & StopTimer(0)
	''dbget.Close()
	''Response.End

	objData = Split(objData, "<br>")

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	sqlStr = ""
	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, "|")

		if (objLine(0) = "B") then

			PGgubun			= "inicis"

			PGuserid = objLine(4)

			if (objLine(4) = "teenxteen3") then
				sitename = "fingers"
			elseif (objLine(4) = "teenxteen4") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen5") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen6") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen8") then
				sitename = "10x10gift"
			elseif (objLine(4) = "teenxteen9") then
				sitename = "10x10mobile"
			elseif (objLine(4) = "teenteen10") then
				if (Left(objLine(8),6) = "INIMX_") Then
					sitename = "10x10mobile"
				Else
					sitename = "10x10"
				End If
			else
				sitename = "XXX"
			end if

			if (objLine(11) = "A") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "A"
				PGCSkey		= ""

				appDate			= objLine(12)
				cancelDate		= "NULL"
			elseif (objLine(11) = "C") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "C"
				PGCSkey		= "CANCELALL"

				appDate			= objLine(12)
				cancelDate		= objLine(13)
			elseif (objLine(11) = "P") then
				'// ==============================
				'// 부분취소
				PGkey		= objLine(9)
				appDivCode	= "R"
				PGCSkey		= objLine(8)

				appDate			= "NULL"
				cancelDate		= objLine(13)
			else
				'// ==============================
				PGkey		= objLine(8)
				appDivCode = "E"
				PGCSkey		= "ERROR"
			end if

			''appMethod		= objLine(3)

			if (objLine(3) = "CC") then
				appMethod = "100"
			elseif (objLine(3) = "AC") then
				appMethod = "20"
			elseif (objLine(3) = "VA") then
				appMethod = "7"
			elseif (objLine(3) = "MO") then
				appMethod = "400"
			else
				appMethod = objLine(3)
			end if

			appPrice		= objLine(16)
			commPrice		= objLine(17)
			commVatPrice	= objLine(18)
			jungsanPrice	= objLine(20)

			ipkumdate		= objLine(5)

			'// 20130503000623
			'// (2013-05-03 00:06:23)
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if

			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end if

			'// 20130510
			'// (2013-05-10)
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)


			If (sqlStr = "") Then
				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			Else
				sqlStr = sqlStr + ", ('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			End If

			If (i <> 0) And ((i mod 500) = 0) Then
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr

				sqlStr = ""
			End If
		end if
	Next

	If (sqlStr <> "") Then
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		sqlStr = ""
	End If

	''rw "aaa" & Now()
	''dbget.close()
	''response.end

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'inicis' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('입금일자 : " + CStr(yyyymmdd) + " [" & StopTimer(0) & " sec]');</script>"
	end If

elseif (mode="getonpgdatakakaopayT") then
	'// ========================================================================
	'// 카카오PAY(거래대사)

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	If (yyyymmdd = "") Then
		'// 전날
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// 해킹대비

	''yyyymmdd = "20150819"

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gT" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// 현재 전부 모바일
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A :승인, C : 취소(전체취소 or 부분취소)
				Select Case objLine(3)
					Case "A"
						'// ==============================
						PGkey		= objLine(5)
						appDivCode	= "A"
						PGCSkey		= ""

						appDate		= objLine(2)
						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(5)
						appDivCode	= "C"
						PGCSkey		= "UNKNOWN"

						appDate		= "NULL"
						cancelDate		= objLine(2)
					Case Else
						'// ==============================
						PGkey		= objLine(5)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// 현재 카드결제만
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(8)
				commPrice		= 0
				commVatPrice	= 0
				jungsanPrice	= 0

				If appDivCode <> "A" Then
					appPrice = appPrice * -1
				End If

				ipkumdate		= ""

				'// 20130503
				'// (2013-05-03)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				''ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		if application("Svr_Info") <> "Dev" Then
			'// 테섭은 데이타 없으므로 오작동

			'// 전체취소 or 부분취소
			sqlStr = " update T "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " T.PGCSkey = (case when l.clogIdx is NULL then 'CANCELALL' else T.pgkey end) "
			sqlStr = sqlStr + " , T.appDivCode = (case when l.clogIdx is NULL then 'C' else 'R' End) "
			sqlStr = sqlStr + " , T.orderserial = (case when l.clogIdx is NULL then NULL else l.orderserial End) "
			sqlStr = sqlStr + " , T.cancelDate = (case when l.clogIdx is NULL then T.cancelDate else l.regdate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_card_cancel_log] l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		T.pgkey = l.newtid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.pggubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and T.appDivCode = 'C' "
			sqlStr = sqlStr + " 	and T.PGCSkey = 'UNKNOWN' "
			dbget.execute sqlStr

			'// 주문번호, 결제일자
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.orderserial is NULL "
			dbget.execute sqlStr

			'// 과거내역
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.orderserial is NULL "
			dbget.execute sqlStr

			'// 전체취소일자
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.cancelDate = a.finishdate "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and T.orderserial = a.orderserial "
			sqlStr = sqlStr + " 		and T.appDivCode = 'C' "
			sqlStr = sqlStr + " 		and a.divcd = 'A007' "
			sqlStr = sqlStr + " 		and a.currstate = 'B007' "
			sqlStr = sqlStr + " 		and a.deleteyn <> 'Y' "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
			dbget.execute sqlStr

		End If

		sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
		sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
		sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
		sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.idx is NULL "
		sqlStr = sqlStr + " 	and t.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		' dbget.Close()
		' Response.end

		if  (Not IsAutoScript) then
			response.write "<script>alert('거래일자 : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('거래대사 파일이 없습니다.[0]');</script>"
		end if
		response.write "거래대사 파일이 없습니다[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

elseif (mode="getonpgdatakakaopayS") then
	'// ========================================================================
	'// 카카오PAY(거래대사)

	'// C:/KMPay_jungsan/Report/cnstest22mS20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gS20150818.csv

	If (yyyymmdd = "") Then
		'// 전날
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// 해킹대비

	''yyyymmdd = "20150827"

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gS" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			''rw objLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// 현재 전부 모바일
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A : 승인, C : 취소, P: 부분취소
				Select Case objLine(2)
					Case "A"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "A"
						PGCSkey		= ""

						'// 20150303,160405
						'// 20130503000623
						'// (2013-05-03 00:06:23)
						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "C"
						PGCSkey		= "CANCELALL"

						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case "P"
						'// ==============================
						'// 부분취소
						PGkey		= objLine(17)
						appDivCode	= "R"
						PGCSkey		= objLine(8)

						appDate			= "NULL"
						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case Else
						'// ==============================
						PGkey		= objLine(8)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// 현재 카드결제만
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(11)
				If (appDivCode <> "A") Then
					appPrice = appPrice * -1
				End If

				commPrice		= objLine(13)
				commVatPrice	= Round(1.0 * commPrice * (1.0/11))

				commPrice = commPrice - commVatPrice

				If (appDivCode = "A") Then
					commPrice = commPrice * -1
					commVatPrice = commVatPrice * -1
				End If

				jungsanPrice	= appPrice + (commPrice + commVatPrice)

				ipkumdate		= objLine(14)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") Then
					If (appDate = "") Then
						appDate = "NULL"
					Else
						appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
					End If
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		sqlStr = " update l "
		sqlStr = sqlStr + " set l.commPrice = T.commPrice, l.commVatPrice = T.commVatPrice, l.jungsanPrice = T.jungsanPrice, l.ipkumdate = T.ipkumdate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and T.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 		and T.PGgubun = l.PGgubun "
		sqlStr = sqlStr + " 		and T.PGkey = l.PGkey "
		sqlStr = sqlStr + " 		and T.appDivCode = l.appDivCode "
		sqlStr = sqlStr + " 		and T.PGCSkey = l.PGCSkey "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		'// 당일 전체취소는 내역이 안온다.
		sqlStr = " update db_order.dbo.tbl_onlineApp_log "
		sqlStr = sqlStr + " set jungsanPrice = appPrice, ipkumdate = convert(varchar(10), IsNull(cancelDate,appDate), 127) "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 	and PGkey in ( "
		sqlStr = sqlStr + " 		select a.PGkey "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			db_order.dbo.tbl_onlineApp_log a "
		sqlStr = sqlStr + " 			join db_order.dbo.tbl_onlineApp_log c "
		sqlStr = sqlStr + " 			on "
		sqlStr = sqlStr + " 				1 = 1 "
		sqlStr = sqlStr + " 				and a.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 				and a.PGgubun = c.PGgubun "
		sqlStr = sqlStr + " 				and a.PGkey = c.PGkey "
		sqlStr = sqlStr + " 				and a.appDivCode = 'A' "
		sqlStr = sqlStr + " 				and c.appDivCode = 'C' "
		sqlStr = sqlStr + " 				and a.PGCSkey = '' "
		sqlStr = sqlStr + " 				and c.PGCSkey = 'CANCELALL' "
		sqlStr = sqlStr + " 				and convert(varchar(10), a.appDate, 127) = convert(varchar(10), c.cancelDate, 127) "
		sqlStr = sqlStr + " 				and a.ipkumdate = '' "
		sqlStr = sqlStr + " 				and a.ipkumdate = c.ipkumdate "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and ipkumdate = '' "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		if  (Not IsAutoScript) then
			response.write "<script>alert('정산일자 : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('정산대사 파일이 없습니다.[0]');</script>"
		end if
		response.write "정산대사 파일이 없습니다[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

	''dbget.Close
	''response.end

elseif (mode="getonpgdatakakaopay") then
	'// ========================================================================
	'// 카카오PAY

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// 해킹대비

	If (yyyymmdd = "") Then
		'// 전날
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gT" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// 현재 전부 모바일
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A : 승인, C : 취소, P: 부분취소
				Select Case objLine(2)
					Case "A"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "A"
						PGCSkey		= ""

						'// 20150303,160405
						'// 20130503000623
						'// (2013-05-03 00:06:23)
						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "C"
						PGCSkey		= "CANCELALL"

						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case "P"
						'// ==============================
						'// 부분취소
						PGkey		= objLine(17)
						appDivCode	= "R"
						PGCSkey		= objLine(8)

						appDate			= "NULL"
						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case Else
						'// ==============================
						PGkey		= objLine(8)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// 현재 카드결제만
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(11)
				commPrice		= objLine(13)
				commVatPrice	= Round(1.0 * commPrice * (1.0/11))
				jungsanPrice	= appPrice - commPrice

				ipkumdate		= objLine(14)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
		sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
		sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
		sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.idx is NULL "
		sqlStr = sqlStr + " 	and t.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		if  (Not IsAutoScript) then
			response.write "<script>alert('거래일자 : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('정산대사 파일이 없습니다.[0]');</script>"
		end if
		response.write "정산대사 파일이 없습니다[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

elseif (mode="getonpgdatauplus") then

	'// ========================================================================
	'// UPLUS

	yyyymmdd = "2016-01-08"

	if (yyyymmdd = "") then
		lastipkumdate = "2012-12-31"

		'// 매출일자
		sqlStr = " select max(PGmeachulDate) as lastipkumdate " & VbCRLF
		sqlStr = sqlStr & " from db_order.dbo.tbl_onlineApp_log " & VbCRLF
		sqlStr = sqlStr & " where PGgubun = 'uplus' " & VbCRLF
		''response.write sqlStr

		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			if Not IsNull(rsget("lastipkumdate")) then
				lastipkumdate = rsget("lastipkumdate")
			end if
		end if
		rsget.Close

		for i = 0 to 20
			'// TODO : 20일 이상 입금액이 없으면 오류
			searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)) + 1)), 10)

			if (searchipkumdate = Left(now, 10)) then
				if  (Not IsAutoScript) then
					response.write "<script>alert('가져올 데이타가 없습니다.[0]');</script>"
				end if
				response.write "가져올 데이타가 없습니다[0]" & searchipkumdate
				response.end
			end if

			ipkumdate = Replace(searchipkumdate, "-", "")

			'// ========================================================================
			'// 온라인 텐텐 정산내역
			response.write "매입일"&CStr(ipkumdate)
			xmlURL = "http://pgweb.uplus.co.kr/pg/wmp/outerpage/trxdown.jsp?mertid=tenbyten01&servicecode=ADJ&trxdate=" + CStr(ipkumdate) + "&key=2beec91670e1f2840a7ac80adde00e49"

			objData = ""

			Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

			objXML.Open "GET", xmlURL, false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()

			if objXML.Status = "200" then
				objData = BinaryToText(objXML.ResponseBody, "euc-kr")
			end if

			Set objXML  = Nothing

			if (Replace(Trim(objData), vbCrLf, "") <> "") then
				exit for
			end if

			lastipkumdate = searchipkumdate

		next

		if (i >= 20) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('가져올 데이타가 없습니다.[" + CStr(i) + "]');</script>"
			end if
			response.write "가져올 데이타가 없습니다[1]"
			response.end
		end if
	else
		'// ========================================================================
		'// 온라인 텐텐 정산내역
		response.write "매입일::"&CStr(Replace(yyyymmdd, "-", ""))
		xmlURL = "http://pgweb.uplus.co.kr/pg/wmp/outerpage/trxdown.jsp?mertid=tenbyten01&servicecode=ADJ&trxdate=" + CStr(Replace(yyyymmdd, "-", "")) + "&key=2beec91670e1f2840a7ac80adde00e49"

		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		end if

		Set objXML  = Nothing

		''response.write "aaa" & Trim(objData)

		if (Replace(Trim(objData), vbCrLf, "") = "") then
			if  (Not IsAutoScript) then
				response.write "<script>alert('가져올 데이타가 없습니다.[--]');</script>"
			end if
			response.write "가져올 데이타가 없습니다[--]"
			response.end
		end if

		searchipkumdate = yyyymmdd
	end if

	objData = Split(objData, vbCrLf)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'uplus' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

 '' response.write xmlURL
 '' response.end

	prevPGkey = ""
	prevPrevPGkey = ""
	prevAppDivCode = ""
	prevPrevAppDivCode = ""
	IsDuplicate = False
	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, ";")

		if (UBound(objLine) >= 0) then
			if (objLine(0) = "D") then

				PGgubun		= "uplus"
				PGkey		= objLine(3)
				PGuserid 	= objLine(2)

				if (PGuserid = "dacomtest") then
					sitename = "dacomtest"
				elseif (PGuserid = "tenbyten01") or (PGuserid = "tenbyten02") then
					'// PC MOBILE 구분 없음(주문내역에서 분리)
					sitename = "10x10"
				else
					sitename = "XXX"
				end if

				if (objLine(6) = "CA01") or (objLine(6) = "CS01") or (objLine(6) = "WR01") then
					'// ==============================
					appDivCode	= "A"
					PGCSkey		= ""

					appDate			= objLine(9)

					cancelDate		= "NULL"
				elseif (objLine(6) = "CA02") or (objLine(6) = "CS02") or (objLine(6) = "WR02") then
					'// ==============================
					appDivCode	= "C"
					PGCSkey		= "CANCELALL"

					appDate			= "NULL"
					cancelDate		= objLine(9)
				elseif (objLine(6) = "CA11") or (objLine(6) = "CS03") or (objLine(6) = "WR06") then
					'// ==============================
					'// 부분취소
					'// 가상계좌환불은 부분취소와 전체취소를 승인 금액으로 구분해야한다.
					appDivCode	= "R"
					PGCSkey		= objLine(9) + "-" + objLine(1)			'// 매출일자 + 일련번호

					appDate			= "NULL"
					cancelDate		= objLine(9)
				else
					'// ==============================
					appDivCode = "E"
					PGCSkey		= "ERROR"
				end if

				if (Left(objLine(6), 2) = "CS") then
					appMethod = "7"
				elseif (Left(objLine(6), 2) = "WR") then
					appMethod = "400"
				elseif (Left(objLine(6), 2) = "CA") then
					appMethod = "100"
				else
					appMethod = Left(objLine(6), 2)
				end if

				appPrice		= objLine(7)
				commVatPrice	= round(1.0 * objLine(8) * (1.0/11))
				commPrice		= objLine(8) - commVatPrice
				jungsanPrice	= objLine(7) - objLine(8)

				commPrice = commPrice * -1
				commVatPrice = commVatPrice * -1

				ipkumdate		= objLine(10)

				'// 20130510
				'// (2013-05-10)
				if (appDate <> "NULL") then
					appDate = Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(appDate, 2)
					appDate = "'" + appDate + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(cancelDate, 2)
					cancelDate = "'" + cancelDate + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				''prevPGkey, prevAppDivCode, IsDuplicate

				if (i >= 1) then
					'// 중복 승인처리(참고 : 13020397762)
					'// TODO : 내역은 주문번호 순서로 정렬되어 있다고 가정한다.
					'// 참조 : 13020397762, 13050293886, 13080752741, 16010731214, 16010731454

					IsDuplicate = False
					If (PGkey = prevPGkey) Then
						if (objLine(6) = "CS01") and (prevAppDivCode = "CS01") Then
							''중복승인
							IsDuplicate = True
						elseif (objLine(6) = "CS02") and (prevAppDivCode = "CS02") Then
							''중복취소
							IsDuplicate = True
						elseif (prevPGkey = prevPrevPGkey) Then
							''3건이상
							IsDuplicate = True
						End If
					End If

					if (prevPGkey <> "") then
						prevPrevPGkey = prevPGkey
						prevPrevAppDivCode = prevAppDivCode
					end if

					prevPGkey = PGkey
					prevAppDivCode = objLine(6)

					if (IsDuplicate = True)  then
						sqlStr = " select count(*) as cnt "
						sqlStr = sqlStr + " from "
						sqlStr = sqlStr + " db_temp.dbo.tbl_onlineApp_log_tmp "
						sqlStr = sqlStr + " where "
						sqlStr = sqlStr + " 1 = 1 "
						sqlStr = sqlStr + " and PGkey like '" + CStr(PGkey) + "%' and appDivCode = '" + appDivCode + "' "
						''response.write sqlStr

						subPgKey = ""
						rsget.Open sqlStr,dbget,1
						if Not(rsget.EOF or rsget.BOF) Then
							If rsget("cnt") > 0 Then
								subPgKey = "-" & Format00(2, rsget("cnt"))
							End If
						end if
						rsget.Close

						PGkey = PGkey + subPgKey
					end if
				end if

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				if (PGkey <> "16010731214") and (PGkey <> "16010731214-01") and (PGkey <> "16010731454") and (PGkey <> "16010731454-01") then
					dbget.execute sqlStr
				Else
					response.write sqlStr + "<br>"
				end if
			end if
		end if
	next

	sqlStr = " update db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr + " set orderserial = pgkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(pgkey) < 20 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 임시주문번호 => 주문번호
	sqlStr = " update t set t.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial, Left(o.paygatetid, (charindex('|', o.paygatetid) - 1)) as paygatetid "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 			and charindex('|', o.paygatetid) > 0 "										'// 구분자 '|'
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// 매출일 전달 또는 이번달
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.pgkey = o.paygatetid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(t.pgkey) >= 6 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 임시주문번호 => 주문번호
	sqlStr = " update t set t.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial, Left(o.paygatetid, (charindex(';', o.paygatetid) - 1)) as paygatetid "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 			and charindex(';', o.paygatetid) > 0 "										'// 구분자 ';'
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// 매출일 전달 또는 이번달
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.pgkey = o.paygatetid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(t.pgkey) >= 6 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 모바일
	sqlStr = " update t set t.sitename = '10x10mobile' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	''sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	''sqlStr = sqlStr + " 			and o.rdsite = 'mobile' "													'// 모바일
	sqlStr = sqlStr + " 			and o.beadaldiv in (4,5,7,8) "												'// 모바일(4:mobile, 5:mobile_link, 7:APP, 8:between ), 2015-07-13, skyer9
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// 매출일 전달 또는 이번달
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = o.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// PG사 매출일
	sqlStr = " update db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr + " set PGmeachulDate = convert(varchar(10), IsNull(cancelDate, appdate), 127) "
	sqlStr = sqlStr + " where pggubun = 'uplus' and PGmeachulDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 실결제일
	sqlStr = " update t set t.appdate = IsNull(o.ipkumdate, t.appdate) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = o.orderserial "
	sqlStr = sqlStr + " where t.pggubun = 'uplus' and appDivCode = 'A' and o.jumundiv not in ('6', '9') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 실취소일
	sqlStr = " update t set t.cancelDate = a.finishdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and t.appDivCode = 'C' "						'// 취소
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and r.refundresult = (t.appPrice * -1) "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 실취소일(교환주문 반품)
	sqlStr = " update t set t.cancelDate = a.finishdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = c.orgorderserial "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.chgorderserial = a.orderserial "
	sqlStr = sqlStr + " 		and t.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and r.refundresult = (t.appPrice * -1) "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''sqlStr = " delete from db_order.dbo.tbl_onlineApp_log where PGmeachulDate = '" + CStr(searchipkumdate) + "' "
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	sqlStr = " delete l "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t.PGgubun = 'uplus' "
	sqlStr = sqlStr + " 		and t.PGgubun = l.PGgubun "
	sqlStr = sqlStr + " 		and t.PGkey = l.PGkey "
	sqlStr = sqlStr + " 		and t.PGCSkey = l.PGCSkey "
	sqlStr = sqlStr + " 		and l.PGmeachulDate = '" + CStr(searchipkumdate) + "' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, t.PGmeachulDate, t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 매칭
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		l.pgkey = o.orderserial "
	sqlStr = sqlStr + " where l.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	if  (Not IsAutoScript) then
		response.write "<script>alert('거래일자 : " + CStr(searchipkumdate) + "');</script>"
	end if

elseif (mode = "matchpgdata") then

	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// 교환주문 제외
	''sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3분
	sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'' 6개월 이전 내역으로 매칭
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_log.dbo.tbl_old_order_master_2003 o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// 교환주문 제외
	sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3분
	sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "
	''response.write sqlStr + "<br>"
	'' 일단 빼고 필요할 때 수기로 실행하자(2014-09-05, skyer9)
	''dbget.execute sqlStr

	'// 원주문 승인취소
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "								'// 취소만
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "		'// 취소는 한건이므로 하루 차이나도 매칭
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'uplus', 'kakaopay') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 원주문 승인취소(OK+신용)
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "							'// 취소
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	a.orderserial = o.orderserial "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	o.orderserial = e.orderserial and e.acctdiv = '100' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "		'// 취소는 한건이므로 하루 차이나도 매칭
	sqlStr = sqlStr + " 	and e.realPayedsum = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 원주문 취소&반품
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// 취소, 부분취소
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2분
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'uplus', 'kakaopay') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 원주문 취소&반품(OK+신용)
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "							'// 취소
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	a.orderserial = o.orderserial "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	o.orderserial = e.orderserial and e.acctdiv = '100' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2분
	sqlStr = sqlStr + " 	and e.realPayedsum = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 교환주문 반품
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		l.orderserial = c.orgorderserial and c.deldate is NULL "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.chgorderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''중복매칭 확인
	'' select orderserial, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode = 'A'
	'' group by orderserial
	'' having count(*) > 1

	'' select orderserial, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode = 'C'
	'' group by orderserial
	'' having count(*) > 1

	'' select orderserial, csasid, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode <> 'A' and csasid is not NULL
	'' group by orderserial, csasid
	'' having count(*) > 1


	'// 부분취소이면서 결제당일 취소의 경우
	'// cancelDate 가 결제일 이후 날자로 지정되고 시간대만 동일하게 유지된다.
	'// 따라서 시간대만 비교해서 매칭해준다.
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	left join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate) % (24 * 60)) < 2 "			'// 동일 시간대
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "matchfingerspgdata") then

	'' sqlStr = " update l set l.orderserial = o.orderserial "
	'' sqlStr = sqlStr + " from "
	'' sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master o "
	'' sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	'' sqlStr = sqlStr + " on "
	'' sqlStr = sqlStr + " 	1 = 1 "
	'' sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	'' sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	'' sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	'' sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	'' sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// 교환주문 제외
	'' sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3분
	'' sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay')) "
	'' sqlStr = sqlStr + " where "
	'' sqlStr = sqlStr + " 	1 = 1 "
	'' sqlStr = sqlStr + " 	and l.orderserial is NULL "
	sqlStr = " exec [db_order].[dbo].[usp_TEN_PGData_Match_FingersOrder] "
	''response.write sqlStr + "<br>"
	''response.end
	dbget.execute sqlStr

	'// 원주문 취소&반품
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// 취소, 부분취소
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [ACADEMYDB].[db_academy].dbo.tbl_academy_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2분
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "matchgiftcardpgdata") then
    ''이니시스 가상계좌의 경우 입금요청 TID랑  입금완료TID랑 다른듯함  tbl_onlineApp_log 에는 입금TID가 옴
	sqlStr = " update l set l.orderserial = o.giftorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_giftcard_order o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paydateid = l.PGkey "
	sqlStr = sqlStr + " 	and o.paydateid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// 교환주문 제외
	sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) <= 5) or (l.appDivCode <> 'A'))  "		'// 5분
	sqlStr = sqlStr + " 	and ((Left(o.paydateid, 9) = 'IniTechPG') or (Left(o.paydateid, 5) = 'INIMX') or (Left(o.paydateid, 6) = 'INIswt') or (Left(o.paydateid, 6) = 'Stdpay')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and IsNull(l.orderserial, '') = '' "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 입금요청 TID 와 실제입금TID 가 서로 다르다.
	'// db_order.dbo.tbl_cyberAcctNoti_LogINI 를 이용해서 매칭시켜준다.
	sqlStr = " update l "
	sqlStr = sqlStr + " set l.orderserial = o.giftorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_giftcard_order o "
	sqlStr = sqlStr + " 	Join db_order.dbo.tbl_cyberAcctNoti_LogINI Nt "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and Nt.NO_OID=o.giftorderserial "
	sqlStr = sqlStr + " 		and Nt.isMatched='Y' "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and Nt.NO_TID = l.PGkey "
	sqlStr = sqlStr + " 		and o.paydateid is not NULL "
	sqlStr = sqlStr + " 		and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 		and o.jumundiv <> '6' "
	sqlStr = sqlStr + " 		and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) <= 5) or (l.appDivCode <> 'A')) "
	sqlStr = sqlStr + " 		and ((Left(l.PGkey , 9) = 'IniTechPG') or (Left(l.PGkey, 5) = 'INIMX') or (Left(l.PGkey, 6) = 'INIswt') or (Left(l.PGkey, 6) = 'Stdpay')) "
	sqlStr = sqlStr + " 	and appmethod='7' "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and IsNull(l.orderserial, '') = '' "
	sqlStr = sqlStr + " 		and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 원주문 취소&반품
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// 취소, 부분취소
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) <= 5 "		'// 5분
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// 전체내역 원 결제일 업데이트
	sqlStr = " update "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set orgPayDate = convert(VARCHAR(10), appDate, 127) "
	sqlStr = sqlStr + " where appDate is not NULL and orgPayDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set r.orgPayDate = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log r "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and r.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	and r.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	and r.appDivCode = 'R' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where r.appDate is NULL and a.appDate is not NULL and r.orgPayDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "makeadvprc") then

	sqlStr = " select PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[DBDATAMART].db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and paydate is not NULL "
	sqlStr = sqlStr + " 	and pgkey is not NULL "
	sqlStr = sqlStr + " 	and paydate >= '" + Left(Now, 7) + "-01' "
	sqlStr = sqlStr + " 	and paydate < '" + Left(DateAdd("m", 1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and payDivCode not in ('mil', 'dep', 'gif', '0', 'XXX') "
	sqlStr = sqlStr + " 	and not (payDivCode in ('rde') and realPayPrice = 0) "
	sqlStr = sqlStr + " 	and PGgubun <> 'KICC' "

	'// 같은 날짜 같은 금액의 환불건이 있는 경우 잘못 매칭될 수 있다.
	'// sqlStr = sqlStr + " 	and PGkey<>'15062692753'"  ''일단.제외. 상구 해결.

	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " having count(*) > 1 "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then
		response.write "중복 ERROR : " & rsget("pgkey")
		rsget.close
		dbget.close()
		response.end
	end if
	rsget.close

	sqlStr = " exec [db_summary].[dbo].[usp_Ten_appPrc_advPrc_SumMake] '" + CStr(yyyymm) + "' "
	rsget.Open sqlStr, dbget, 1

	'response.write	"<script language='javascript'>" &_
	'				"	alert('작성되었습니다.'); " &_
	'				"	history.back(); " &_
	'				"</script>"

end if

%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script language='javascript'>
alert('저장되었습니다.');
//location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
