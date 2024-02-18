<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<!-- #include virtual="/admin/etc/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oCoupang, failCnt, chgSellYn, arrRows, isItemIdChk, maeipdiv, mustPrice
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, i, tCoupangGoodno, errVendorItemId
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0
Select Case action
	Case "REGDELIVERY"
		isItemIdChk = "N"
		itemid			= requestCheckVar(request("itemid"),32)
	Case Else				isItemIdChk = "Y"
End Select

If isItemIdChk = "Y" Then
	If itemid="" or itemid="0" Then
		response.write "<script>alert('상품번호가 없습니다.')</script>"
		response.end
	ElseIf Not(isNumeric(itemid)) Then
		response.write "<script>alert('잘못된 상품번호입니다.')</script>"
		response.end
	Else
		'정수형태로 변환
		itemid = CLng(getNumeric(itemid))
	End If
End If

'######################################################## Coupang API ########################################################
If action = "REGDELIVERY" Then								'출고지등록
	maeipdiv = fnBrandmaeipdiv(itemid)
	Call fnCoupangDeliveryReg(itemid, maeipdiv, iErrStr)
ElseIf action = "REG" Then									'상품등록
	SET oCoupang = new CCoupang
		oCoupang.FRectItemID	= itemid
		oCoupang.getCoupangNotRegOneItem
	    If (oCoupang.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Coupang_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oCoupang.FOneItem.checkTenItemOptionValid Then
				Call fnCoupangItemReg(itemid, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oCoupang = nothing
ElseIf action = "CHKSTAT" Then								'상품 조회
	tCoupangGoodno = getCoupangGoodno(itemid)
	Call fnCoupangStatChk(itemid, tCoupangGoodno, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EditSellYn" Then							'상품 상태 변경
	arrRows = getCoupangVendorItemidList(itemid)
	If isArray(arrRows) Then
		For i = 0 To UBound(arrRows,2)
			Call fnCoupangSellyn(itemid, chgSellYn, arrRows(0, i), errVendorItemId)
			If errVendorItemId <> "" Then
				SumErrStr = SumErrStr & errVendorItemId & ","
			End If
		Next
		iErrStr = ArrErrStrInfo(action, chgSellYn, itemid, SumErrStr)
	Else
		iErrStr = "ERR||"&itemid&"||[상태변경] 조회부터 실행하세요. by kjy"
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "DELETE" Then								'상품 삭제
	tCoupangGoodno = getCoupangGoodno(itemid)
	Call fnCoupangDelete(itemid, tCoupangGoodno, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "PRICE" Then								'가격 수정
	arrRows = getCoupangVendorItemidList(itemid)
	If isArray(arrRows) Then
		For i = 0 To UBound(arrRows,2)
			Call fnCoupangPrice(itemid, arrRows(0, i), arrRows(1, i), arrRows(2, i), errVendorItemId)
			If errVendorItemId <> "" Then
				SumErrStr = SumErrStr & errVendorItemId & ","
			End If
			mustPrice = arrRows(1, i)
		Next
		iErrStr = ArrErrStrInfo(action, mustPrice, itemid, SumErrStr)
	Else
		iErrStr = "ERR||"&itemid&"||[가격수정] 조회부터 실행하세요. by kjy"
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "QTY" Then									'재고 수정
	arrRows = getCoupangVendorItemidList(itemid)
	If isArray(arrRows) Then
		For i = 0 To UBound(arrRows,2)
			Call fnCoupangQuantity(itemid, arrRows(0, i), arrRows(3, i), arrRows(4, i), errVendorItemId)
			If errVendorItemId <> "" Then
				SumErrStr = SumErrStr & errVendorItemId & ","
			End If
		Next
		iErrStr = ArrErrStrInfo(action, "", itemid, SumErrStr)
	Else
		iErrStr = "ERR||"&itemid&"||[재고수정] 조회부터 실행하세요. by kjy"
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("coupang", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDIT" Then									'상품 수정
	SET oCoupang = new CCoupang
		oCoupang.FRectItemID	= itemid
		oCoupang.getCoupangEditOneItem
		If oCoupang.FResultCount = 0 Then
	    	failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||수정가능한 상품이 아닙니다."
		Else
			arrRows = getCoupangVendorItemidList(itemid)
			'######################################## 1-1. 품절 처리 ########################################
			If (oCoupang.FOneItem.FmaySoldOut = "Y") OR (oCoupang.FOneItem.IsMayLimitSoldout = "Y") OR (oCoupang.FOneItem.IsAllOptionChange = "Y") Then
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						Call fnCoupangSellyn(itemid, "N", arrRows(0, i), errVendorItemId)
						If errVendorItemId <> "" Then
							SumErrStr = SumErrStr & errVendorItemId & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("EditSellYn", "N", itemid, SumErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			Else
			'######################################## 1-2. 판매 처리 ########################################
				If (oCoupang.FOneItem.FCoupangSellYn = "N" AND oCoupang.FOneItem.IsSoldOut = False) Then
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							Call fnCoupangSellyn(itemid, "Y", arrRows(0, i), errVendorItemId)
							If errVendorItemId <> "" Then
								SumErrStr = SumErrStr & errVendorItemId & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EditSellYn", "Y", itemid, SumErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
			'######################################## 2. 재고 수정 ########################################
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						Call fnCoupangQuantity(itemid, arrRows(0, i), arrRows(3, i), arrRows(4, i), errVendorItemId)
						If errVendorItemId <> "" Then
							SumErrStr = SumErrStr & errVendorItemId & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("QTY", "", itemid, SumErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			'######################################## 3. 수정 처리 ########################################
				Call fnCoupangItemEdit(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If
	SET oCoupang = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("coupang", itemid, SumErrStr)
		Call SugiQueLogInsert("coupang", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regItem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("coupang", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&iErrStr&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 500);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->