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
dim sellsite, selldate, csGubun, isSuccess
dim i, j, k
dim nowdate, fromdate, todate, currdate
dim sqlStr, resultMessage, msg
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
csGubun		= requestCheckVar(html2db(request("mode")),32)

dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False

if (selldate = "") then
	'// 오늘까지 일괄로 가져오기
	Call GetCSCheckStatus(sellsite, csGubun, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)
else
	fromdate = selldate
	todate = selldate
	IS_SELLDATE_FIXED = True
end if

select case sellsite
	Case "ezwel"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate

		do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			if (csGubun = "ordercancel") then
				Call GetCSOrder_ezwel(currdate, "CANCLEREQ", resultMessage)
				rw resultMessage
				Call GetCSOrder_ezwel(currdate, "CANCLEDONE", resultMessage)
				rw resultMessage
			elseif (csGubun = "return") then
				Call GetCSOrder_ezwel(currdate, "RETURNREQ", resultMessage)
				rw resultMessage
				Call GetCSOrder_ezwel(currdate, "RETURNDONE", resultMessage)
				rw resultMessage
			elseif (csGubun = "exchange") then
				Call GetCSOrder_ezwel(currdate, "CHANGEREQ", resultMessage)
				rw resultMessage
				Call GetCSOrder_ezwel(currdate, "CHANGEDONE", resultMessage)
				rw resultMessage
			elseif (csGubun = "all") then
				Call GetCSOrder_ezwel(currdate, "CANCLEREQ", resultMessage)
				rw resultMessage
				Call GetCSOrder_ezwel(currdate, "CANCLEDONE", resultMessage)
				rw resultMessage

				Call GetCSOrder_ezwel(currdate, "RETURNREQ", resultMessage)
				rw resultMessage
				Call GetCSOrder_ezwel(currdate, "RETURNDONE", resultMessage)
				rw resultMessage

				Call GetCSOrder_ezwel(currdate, "CHANGEREQ", resultMessage)
				rw resultMessage
				Call GetCSOrder_ezwel(currdate, "CHANGEDONE", resultMessage)
				rw resultMessage
			else
				response.write "잘못된 접근입니다."
				dbget.close : response.end
			end if

			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	Case "wconcept1010"
		If (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		End If
		currdate = fromdate

		Do While (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			If (csGubun = "all") Then
				Call GetCSOrder_wconcept(currdate, "ordercancelRequest", resultMessage)					'취소요청
				Call GetCSOrder_wconcept(currdate, "ordercancelComplete", resultMessage)				'취소완료
				rw resultMessage

				Call GetCSOrder_wconcept(currdate, "return", resultMessage)								'교환요청
				rw resultMessage

				Call GetCSOrder_wconcept(currdate, "exchange", resultMessage)							'반품요청
				rw resultMessage
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		Loop
		rw "#######################################################################################"
	Case "kakaostore"
		If (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		End If
		currdate = fromdate

		Do While (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			If (csGubun = "all") Then
				'취소건 가져오기
				Call GetCSOrder_kakaostore(currdate, "ordercancel", "ShippingCancelComplete")			'결제 취소 완료
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "ordercancel", "ShippingCancelRequestBuyer")		'구매자 배송 취소 요청
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "ordercancel", "ShippingCancelRequestSeller")		'판매자 배송 취소 요청
				rw "---------------------"

				'교환건 가져오기
				Call GetCSOrder_kakaostore(currdate, "return", "ExchangeRequest")						'교환 요청
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "return", "ExchangeShippingComplete")				'교환 완료
				rw "---------------------"

				'반품건 가져오기
				Call GetCSOrder_kakaostore(currdate, "exchange", "ReturnRequest")						'반품 요청
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "exchange", "ReturnShippingComplete")				'반품 반송 완료
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		Loop
		rw "#######################################################################################"
	Case "boribori1010"
		If (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		End If
		currdate = fromdate

		Do While (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			If (csGubun = "all") Then
				Call GetCSOrder_boribori(currdate, "ordercancel", resultMessage)						'취소
				rw resultMessage

				Call GetCSOrder_boribori(currdate, "ordersoldout", resultMessage)						'품절 취소
				rw resultMessage

				Call GetCSOrder_boribori(currdate, "return", resultMessage)								'교환
				rw resultMessage

				Call GetCSOrder_boribori(currdate, "exchange", resultMessage)							'반품
				rw resultMessage
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		Loop
		rw "#######################################################################################"
	Case "benepia1010"
		If (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		End If
		currdate = fromdate

		Do While (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			If (csGubun = "all") Then
				Call GetCSOrderCancel_benepia(currdate, resultMessage)						'취소
				rw resultMessage

				Call GetCSOrderExchange_benepia(currdate, resultMessage)					'반품
				rw resultMessage
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		Loop
		rw "#######################################################################################"
	Case "auction1010", "gmarket1010"
		If (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		End If
		currdate = fromdate

		Do while (currdate <= todate)
			response.write "<br />" & sellsite & " : " & currdate & "<br />"
			If (csGubun = "cancelorder") then
			 	Call GetCSOrderCancel_ebay(currdate, resultMessage, sellsite)				'취소
			 	rw resultMessage
			ElseIf (csGubun = "returnorder") Then
			 	Call GetCSOrderReturn_ebay(currdate, resultMessage, sellsite, 1)			'반품 / 반품요청
			 	rw resultMessage
			 	Call GetCSOrderReturn_ebay(currdate, resultMessage, sellsite, 2)			'반품 / 반품수거완료
			 	rw resultMessage
			 	Call GetCSOrderReturn_ebay(currdate, resultMessage, sellsite, 4)			'반품 / 반품환불완료
			 	rw resultMessage
			ElseIf (csGubun = "changeorder") Then
				Call GetCSOrderExchange_ebay(currdate, resultMessage, sellsite, 1)			'교환 / 교환요청/교환물품반송중
				rw resultMessage
				Call GetCSOrderExchange_ebay(currdate, resultMessage, sellsite, 2)			'교환 / 교환수거완료
				rw resultMessage
				Call GetCSOrderExchange_ebay(currdate, resultMessage, sellsite, 4)			'교환 / 교환완료
				rw resultMessage
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		Loop
		rw "#######################################################################################"
		'http://localhost:11117/admin/etc/order/xSiteCSOrder_Ins_Process.asp?sellsite=gmarket1010&selldate=2023-11-17&mode=cancelorder
		'http://localhost:11117/admin/etc/order/xSiteCSOrder_Ins_Process.asp?sellsite=gmarket1010&selldate=2023-11-17&mode=returnorder
		'http://localhost:11117/admin/etc/order/xSiteCSOrder_Ins_Process.asp?sellsite=gmarket1010&selldate=2023-11-17&mode=changeorder
	Case Else
		response.write "잘못된 접근입니다."
		dbget.close : response.end
End Select

If (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) Then
	If (selldate < Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	ElseIf (selldate = Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, selldate, "Y")
	End If

	'// 제휴몰 취소건 어드민 주문취소 : 전체취소만
	For i = 0 To 20
		msg = ""
		sqlStr = " exec [db_cs].[dbo].[usp_Ten_CheckCancelExtOrder] "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			msg = rsget("msg")
		rsget.Close

		response.write msg & "<br />"
		If (msg = "NO ORDER") Then
			Exit For
		End If
	Next
End If
%>
<% rw "OK" %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
