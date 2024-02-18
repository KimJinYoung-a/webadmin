<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/incShoplinkerFunction.asp"-->
<%
Function CheckVaildIP(ref)
	CheckVaildIP = false
	Dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.9","61.252.133.10","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","192.168.1.72")
	Dim i
	For i = 0 to UBound(VaildIP)
		If (VaildIP(i)=ref) Then
			CheckVaildIP = true
			Exit Function
		End If
	Next
End Function

Dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
If (Not CheckVaildIP(ref)) Then
    dbget.Close()
    response.end
End If

Dim act     : act = requestCheckVar(request("act"),32)
Dim param1  : param1 = requestCheckVar(request("param1"),32)
Dim sqlStr, i, paramData, retVal
Dim retCnt : retCnt = 0
Dim cnt
Dim itemidArr, iitemid, ret
Dim oshoplinker, ierrStr

select Case act
	Case "ShopLinkerSoldOutItem" ''품절처리 상품.
		SET oshoplinker = new CShoplinker
			oshoplinker.FPageSize       = 20
			oshoplinker.FCurrPage       = 1
			oshoplinker.FRectShoplinkerNotReg  = "D"		''등록완료(외부몰 연결)
			oshoplinker.FRectShoplinkerYes10x10No = "on"
			oshoplinker.FRectOrdType = "B"
			oshoplinker.FRectFailCntOverExcept="5"			'' 5회 이상 실패내역 제낌.
			oshoplinker.getShoplinkerRegedItemList

			If (oshoplinker.FResultCount < 1) Then
				response.Write "S_NONE"
				dbget.Close() 
				Set oshoplinker= Nothing
				response.end
			End If

			For i = 0 to oshoplinker.FResultCount - 1
				itemidArr = itemidArr & oshoplinker.FItemList(i).FItemID &","
			Next
		Set oshoplinker= Nothing

		If (Right(itemidArr,1)=",") Then itemidArr = Left(itemidArr , Len(itemidArr) - 1)
		itemidArr = split(itemidArr,",")

		For i=0 to UBound(itemidArr)
			If (itemidArr(i)<>"") Then
				iitemid = Trim(itemidArr(i))
				ierrStr=""
				ret = editSellStatusShoplinkerOneItem(iitemid, ierrStr, "N")
				If (Not ret) Then
					rw ierrStr
				End If
			End If
		Next
		''response.Write "itemidArr="&itemidArr
		response.Write "<br>"&retVal
	Case ELSE
	response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->