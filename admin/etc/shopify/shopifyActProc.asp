<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
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
Dim itemid, action, oshopify, failCnt, chgSellYn, arrRows, isItemIdChk, mustPrice, collectionType
Dim resultMessage, strSql, SumErrStr, SumOKStr, i, strparam, mrgnRate, endItemErrMsgReplace
Dim depth1List, depth2List, depth3List
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
collectionType	= request("collectionType")
failCnt			= 0

If action = "REG" Then									'상품등록
	Call fnShopifyItemReg(itemid, action, resultMessage)
ElseIf action = "DELETE" Then							'상품삭제
	Call fnShopifyItemDel(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then						'상태변경
	Call fnShopifySellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then							'상품조회
	Call fnShopifyChkstat(itemid, action, resultMessage)
ElseIf action = "EDIT" Then								'상품수정
	Call fnShopifyItemEdit(itemid, action, resultMessage)
ElseIf action = "collectionRefresh" Then				'SmartCollection 재갱신
	If collectionType = "brand" Then
		Call fnShopifySmartCollection(resultMessage)
		If resultMessage = "success" Then
			rw "성공"
		End If
	ElseIf collectionType = "category" Then
		Set oshopify = new Cshopify
			depth1List = oshopify.getDepthGroupCodeList(1)
			depth2List = oshopify.getDepthGroupCodeList(2)
			depth3List = oshopify.getDepthGroupCodeList(3)
		Set oshopify = nothing

		rw "depth3List"
		If isArray(depth3List) Then
			For i = 0 to Ubound(depth3List, 2)
				Call fnShopifyCategoryCollectionList(depth3List(0, i), resultMessage)
				response.flush
				response.clear
			Next
		End If
		rw "---------------------"

		rw "depth2List"
		If isArray(depth2List) Then
			For i = 0 to Ubound(depth2List, 2)
				Call fnShopifyCategoryCollectionList(depth2List(0, i), resultMessage)
				response.flush
				response.clear
			Next
		End If
		rw "---------------------"

		rw "depth1List"
		If isArray(depth1List) Then
			For i = 0 to Ubound(depth1List, 2)
				Call fnShopifyCategoryCollectionList(depth1List(0, i), resultMessage)
				response.flush
				response.clear
			Next
		End If
		rw "---------------------"
	End If
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&resultMessage&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->