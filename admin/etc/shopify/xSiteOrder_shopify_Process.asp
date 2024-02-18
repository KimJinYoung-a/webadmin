<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/shopify/shopifyCls.asp"-->
<!-- #include virtual="/admin/etc/shopify/incshopifyFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sqlStr, i, sellsite, objXML, shopifyOrders, line_items, sku
Dim OutMallOrderSerial, SellDate, isSku, strObj, iRbody, iSellDate, searchDate

sellsite	= requestCheckVar(html2db(request("sellsite")),32)
'searchDate = replace(iSellDate, "-", "")
rw " Order START"
On Error Resume Next
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	If application("Svr_Info")="Dev" Then
		objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/shopify?status=open", false
	Else
		objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/shopify?status=open", false
	End If
	objXML.setRequestHeader "Content-Type", "application/json"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
	Set strObj = JSON.parse(iRbody)
		If strObj.message = "no data" Then
			rw "MayBe Today USD rate Null"
		Else
			Set shopifyOrders	= strObj.result.shopifyOrders
				For i=0 to shopifyOrders.length-1
					' isSku = "Y"
					outmallorderserial = shopifyOrders.get(i).id
					' Set line_items = shopifyOrders.get(i).line_items
					' 	For j=0 to line_items.length-1
					' 		sku = line_items.get(j).sku
					' 		If sku = "" Then
					' 			isSku = "N"
					' 		End If
					' 	Next
					' Set line_items	= nothing
					' If isSku = "Y" Then
						rw "id : " & outmallorderserial
					' End If
				Next
			Set shopifyOrders	= nothing
		End If
		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set strObj = nothing
Set objXML = nothing
response.write "<br />"
rw " Order End"

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->