<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 600 ''�ʴ���
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
	'// ���ñ��� �ϰ��� ��������
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
				response.write "�߸��� �����Դϴ�."
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
				Call GetCSOrder_wconcept(currdate, "ordercancelRequest", resultMessage)					'��ҿ�û
				Call GetCSOrder_wconcept(currdate, "ordercancelComplete", resultMessage)				'��ҿϷ�
				rw resultMessage

				Call GetCSOrder_wconcept(currdate, "return", resultMessage)								'��ȯ��û
				rw resultMessage

				Call GetCSOrder_wconcept(currdate, "exchange", resultMessage)							'��ǰ��û
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
				'��Ұ� ��������
				Call GetCSOrder_kakaostore(currdate, "ordercancel", "ShippingCancelComplete")			'���� ��� �Ϸ�
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "ordercancel", "ShippingCancelRequestBuyer")		'������ ��� ��� ��û
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "ordercancel", "ShippingCancelRequestSeller")		'�Ǹ��� ��� ��� ��û
				rw "---------------------"

				'��ȯ�� ��������
				Call GetCSOrder_kakaostore(currdate, "return", "ExchangeRequest")						'��ȯ ��û
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "return", "ExchangeShippingComplete")				'��ȯ �Ϸ�
				rw "---------------------"

				'��ǰ�� ��������
				Call GetCSOrder_kakaostore(currdate, "exchange", "ReturnRequest")						'��ǰ ��û
				rw "---------------------"
				Call GetCSOrder_kakaostore(currdate, "exchange", "ReturnShippingComplete")				'��ǰ �ݼ� �Ϸ�
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
				Call GetCSOrder_boribori(currdate, "ordercancel", resultMessage)						'���
				rw resultMessage

				Call GetCSOrder_boribori(currdate, "ordersoldout", resultMessage)						'ǰ�� ���
				rw resultMessage

				Call GetCSOrder_boribori(currdate, "return", resultMessage)								'��ȯ
				rw resultMessage

				Call GetCSOrder_boribori(currdate, "exchange", resultMessage)							'��ǰ
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
				Call GetCSOrderCancel_benepia(currdate, resultMessage)						'���
				rw resultMessage

				Call GetCSOrderExchange_benepia(currdate, resultMessage)					'��ǰ
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
			 	Call GetCSOrderCancel_ebay(currdate, resultMessage, sellsite)				'���
			 	rw resultMessage
			ElseIf (csGubun = "returnorder") Then
			 	Call GetCSOrderReturn_ebay(currdate, resultMessage, sellsite, 1)			'��ǰ / ��ǰ��û
			 	rw resultMessage
			 	Call GetCSOrderReturn_ebay(currdate, resultMessage, sellsite, 2)			'��ǰ / ��ǰ���ſϷ�
			 	rw resultMessage
			 	Call GetCSOrderReturn_ebay(currdate, resultMessage, sellsite, 4)			'��ǰ / ��ǰȯ�ҿϷ�
			 	rw resultMessage
			ElseIf (csGubun = "changeorder") Then
				Call GetCSOrderExchange_ebay(currdate, resultMessage, sellsite, 1)			'��ȯ / ��ȯ��û/��ȯ��ǰ�ݼ���
				rw resultMessage
				Call GetCSOrderExchange_ebay(currdate, resultMessage, sellsite, 2)			'��ȯ / ��ȯ���ſϷ�
				rw resultMessage
				Call GetCSOrderExchange_ebay(currdate, resultMessage, sellsite, 4)			'��ȯ / ��ȯ�Ϸ�
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
		response.write "�߸��� �����Դϴ�."
		dbget.close : response.end
End Select

If (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) Then
	If (selldate < Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	ElseIf (selldate = Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, selldate, "Y")
	End If

	'// ���޸� ��Ұ� ���� �ֹ���� : ��ü��Ҹ�
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
