<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<!-- #include virtual="/admin/etc/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, failCnt, oCoupang, getMustprice, chgSellYn, vOptCnt, i
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, chgImageNm, arrRows, errVendorItemId, tCoupangGoodno
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
If itemid="" or itemid="0" Then
	response.write "<script>alert('상품번호가 없습니다.')</script>"
	response.end
ElseIf Not(isNumeric(itemid)) Then
	response.write "<script>alert('잘못된 상품번호입니다.')</script>"
	response.end
Else
	'정수형태로 변환
	itemid=CLng(getNumeric(itemid))
End If
'######################################################## coupang API ########################################################
If mallid = "coupang" Then
	If action = "SOLDOUT" Then				'상태변경
		arrRows = getCoupangVendorItemidList(itemid)
		If isArray(arrRows) Then
			For i = 0 To UBound(arrRows,2)
				Call fnCoupangSellyn(itemid, "N", arrRows(0, i), errVendorItemId)
				If errVendorItemId <> "" Then
					SumErrStr = SumErrStr & errVendorItemId & ","
				End If
			Next
			iErrStr = ArrErrStrInfo("EditSellYn", "N", itemid, SumErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||[상태변경] 조회부터 실행하세요. by kjy"
		End If
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		'http://webadmin.10x10.co.kr/admin/etc/coupang/CoupangProc.asp?itemid=1891798&mallid=coupang&action=SOLDOUT
	ElseIf action = "PRICE" Then			'가격수정
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
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		'http://webadmin.10x10.co.kr/admin/etc/coupang/CoupangProc.asp?itemid=1891798&mallid=coupang&action=PRICE
	ElseIf action = "CHKSTAT" Then			'상품조회
		tCoupangGoodno = getCoupangGoodno(itemid)
		Call fnCoupangStatChk(itemid, tCoupangGoodno, iErrStr)
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("coupang", itemid, iErrStr)
		End If
		'http://webadmin.10x10.co.kr/admin/etc/coupang/CoupangProc.asp?itemid=1891798&mallid=coupang&action=CHKSTAT
	ElseIf action = "EDIT" Then				'상품수정
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

		'OK던 ERR이던 editQuecnt에 + 1을 시킴..
		'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regItem SET " & VBCRLF
		strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
		strSql = strSql & " ,coupangLastUpdate = getdate()  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql	

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("coupang", itemid, SumErrStr)
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://webadmin.10x10.co.kr/admin/etc/coupang/CoupangProc.asp?itemid=1891798&mallid=coupang&action=EDIT
	End If
End If
'###################################################### Sabangnet API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->