<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 600 ''초단위
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/order/lib/xSiteOrderLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
dim IS_TEST_MODE : IS_TEST_MODE = False

Dim sqlStr, sellsite, selldate, selldateStr, mode
dim isSuccess
dim i, j, k, lp
dim orderObjArr, tmpObjArr
dim nowdate, fromdate, todate, currdate, hasMoreData, nvCount
Dim chgCode, page, gubunCode
Dim isOrderComplete, regCount
Dim arrRows, resultMessage
lp = 0
hasMoreData = "N"
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
mode		= requestCheckVar(html2db(request("mode")),32)
chgCode		= requestCheckVar(html2db(request("chgCode")),32)
gubunCode	= requestCheckVar(html2db(request("gubunCode")),32)
If chgCode = "" Then
	chgCode = "PAYED"
End If

dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False
if (selldate = "") then
	'// 오늘까지 일괄로 가져오기
	Call GetCheckStatus(sellsite, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)

	if (fromdate = todate) then
		selldateStr = fromdate
	else
		selldateStr = fromdate & " ~ " & todate
	end if
else
	fromdate = selldate
	todate = selldate
	selldateStr = fromdate
	IS_SELLDATE_FIXED = True
end if

select case sellsite
	case "ezwel"
		'// 이지웰페어
		response.write "<script>alert('ezwel : " & selldateStr & "');</script>"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "ezwel : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate, gubunCode, resultMessage)
			rw resultMessage
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "boribori1010"
		'// 보리보리
		response.write "<script>alert('boribori1010 : " & selldateStr & "');</script>"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "boribori1010 : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate, gubunCode, resultMessage)
			rw resultMessage
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "kakaostore"
	'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=kakaostore&selldate=2023-02-08&gubunCode=ShippingRequest
		'// 카카오톡스토어
		response.write "<script>alert('kakaostore : " & selldateStr & "');</script>"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if
		currdate = fromdate
		'1. 최초 테이블 비운다
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_temp.[dbo].[tbl_xSite_TMPOrder_kakaostore] WHERE orderStatus = '"&gubunCode&"' "
		dbget.Execute sqlStr
		do while (currdate <= todate)
			isOrderComplete = "N"
			page = 1
			response.write  sellsite & " : " & currdate & "<br />"
			Do Until isOrderComplete = "Y"
				Call GetOrder_kakaostore(sellsite, currdate, hasMoreData, page, gubunCode)
				If hasMoreData = "N" Then
					isOrderComplete = "Y"
				Else 
					page = page + 1
				End If
				response.flush
			Loop

			If gubunCode <> "ShippingWaiting" Then
				'2. 주문확인처리 
				sqlStr = ""
				sqlStr = sqlStr & " DELETE FROM db_temp.[dbo].[tbl_xSite_TMPOrder_kakaostoreConfirm] WHERE orderRegYN = 'Y' "
				dbget.Execute sqlStr
				Call setOrder_kakaostoreConfirm(gubunCode)

				sqlStr = ""
				sqlStr = sqlStr & " SELECT orderId "
				sqlStr = sqlStr & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_kakaostoreConfirm] "
				sqlStr = sqlStr & " WHERE orderRegYN = 'N' "
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					arrRows = rsget.getRows()
				End If
				rsget.Close
			Else
				sqlStr = ""
				sqlStr = sqlStr & " SELECT k.orderId "
				sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder_kakaostore as k "
				sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_xSite_TMPOrder t on convert(varchar(30), k.paymentId) = t.OutMallOrderSerial  "
				sqlStr = sqlStr & " 	and convert(varchar(30), k.orderId) = t.OrgDetailKey "
				sqlStr = sqlStr & " 	and t.SellSite = 'kakaostore' "
				sqlStr = sqlStr & " WHERE k.orderStatus = '"&gubunCode&"' "
				sqlStr = sqlStr & " and t.OutMallOrderSeq is null "
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					arrRows = rsget.getRows()
				End If
				rsget.Close
			End If

			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					rw "arrRows(0, i) : " & arrRows(0, i)
					Call GetOrder_kakaostoreDetail(arrRows(0, i), resultMessage)
					rw resultMessage
					If (i mod 10) = 9 Then response.flush
				Next
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "wconcept1010"
		'// w컨셉
		response.write "<script>alert('wconcept1010 : " & selldateStr & "');</script>"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		'gubunCode
		'Case 1	
		'	주문 리스트를 가져온다. getParam code : 03 (결제완료)
		'Case 2
		'	결제완료건 리스트로 상세내역을 가져와서 주문을 저장한다.
		'Case 3
		'	저장한 주문의 상태를 발주 확인 처리 한다.
		If gubunCode = "1" Then
			currdate = fromdate
			do while (currdate <= todate)
				response.write "wconcept1010 : " & currdate & "<br />"
				Call GetOrderFromExtSite(sellsite, currdate, gubunCode, resultMessage)
				rw resultMessage
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			loop
		'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=wconcept1010&selldate=2023-02-08&gubunCode=1
		ElseIf gubunCode = "2" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT outMallOrderSerial, orgDetailKey "
			sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder_Common "
			sqlStr = sqlStr & " WHERE sendLevel = '1' "
			sqlStr = sqlStr & " and sellsite = 'wconcept1010' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows()
			End If
			rsget.Close

			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					rw "arrRows(0, i) : " & arrRows(0, i) & "arrRows(1, i) : " & arrRows(1, i)
					Call GetOrder_WconceptDetail(arrRows(0, i), arrRows(1, i), resultMessage)
					rw resultMessage
					If (i mod 10) = 9 Then response.flush
				Next
			Else
				rw "No Order"
			End If
		'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=wconcept1010&gubunCode=2
		ElseIf gubunCode = "3" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT outMallOrderSerial, orgDetailKey "
			sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder_Common "
			sqlStr = sqlStr & " WHERE sendLevel = '2' "
			sqlStr = sqlStr & " and sellsite = 'wconcept1010' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows()
			End If
			rsget.Close

			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					rw "arrRows(0, i) : " & arrRows(0, i) & "arrRows(1, i) : " & arrRows(1, i)
					Call GetOrder_WconceptConfirm(arrRows(0, i), arrRows(1, i), resultMessage)
					rw resultMessage
					If (i mod 10) = 9 Then response.flush
				Next
			Else
				rw "No Order Ready"
			End If
		'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=wconcept1010&gubunCode=3
		End If
	case "benepia1010"
		'// 베네피아
		response.write "<script>alert('benepia1010 : " & selldateStr & "');</script>"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		'gubunCode
		'Case 1	
		'	주문 리스트를 가져온다. getParam code : 1:결제완료, 2:배송준비중, 3:배송중, 6:배송완료, 전체
		'Case 2
		'	결제완료건 리스트로 상세내역을 가져와서 주문을 저장한다.
		'Case 3
		'	저장한 주문의 상태를 발주 확인 처리 한다.
		If gubunCode = "1" Then
			currdate = fromdate
			do while (currdate <= todate)
				response.write "benepia1010 : " & currdate & "<br />"
				Call GetOrderFromExtSite(sellsite, currdate, gubunCode, resultMessage)
				rw resultMessage
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			loop
		'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=benepia1010&selldate=2023-02-08&gubunCode=1
		ElseIf gubunCode = "2" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT outMallOrderSerial "
			sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder_Common "
			sqlStr = sqlStr & " WHERE sendLevel = '1' "
			sqlStr = sqlStr & " and sellsite = 'benepia1010' "
			sqlStr = sqlStr & " GROUP BY outMallOrderSerial "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows()
			End If
			rsget.Close

			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					rw "arrRows(0, i) : " & arrRows(0, i)
					Call GetOrder_benepiaDetail(arrRows(0, i), resultMessage)
					rw resultMessage
					If (i mod 10) = 9 Then response.flush
				Next
			Else
				rw "No Order"
			End If
		'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=benepia1010&gubunCode=2
		ElseIf gubunCode = "3" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT outMallOrderSerial, orgDetailKey "
			sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder_Common "
			sqlStr = sqlStr & " WHERE sendLevel = '2' "
			sqlStr = sqlStr & " and sellsite = 'benepia1010' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows()
			End If
			rsget.Close

			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					rw "arrRows(0, i) : " & arrRows(0, i) & "arrRows(1, i) : " & arrRows(1, i)
					Call GetOrder_benepiaConfirm(arrRows(0, i), arrRows(1, i), resultMessage)
					rw resultMessage
					If (i mod 10) = 9 Then response.flush
				Next
			Else
				rw "No Order Ready"
			End If
		'http://localhost:11117/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=benepia1010&gubunCode=3
		End If
	case else
		response.write "잘못된 접근입니다."
		dbget.close : response.end
end select

if (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) then
	if (selldate < Left(Now(), 10)) then
		Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	elseif (selldate = Left(Now(), 10)) then
		Call SetCheckStatus(sellsite, selldate, "Y")
	end if
end if

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
