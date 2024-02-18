<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/incShoplinkerFunction.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim subcmd : subcmd = requestCheckVar(request("subcmd"),10)
Dim oshoplinker, i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim SuccCNT, FailCNT, ret, alertMsg
Dim retFlag
Dim iMessage
retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")

If (cmdparam = "RegSelect") Then				''선택상품 실제 등록
	SuccCNT = 0
	FailCNT = 0
	arrItemid = split(arrItemid, ",")
	Dim q
	For q = 0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(q))
		ret = regShoplinkerOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next

	alertMsg = ""&SuccCNT&"건 등록 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam = "EDTSelect") Then			''선택상품 실제 수정
	SuccCNT = 0
	FailCNT = 0
	arrItemid = split(arrItemid, ",")
	Dim t
	For t = 0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(t))
		ret = edtShoplinkerOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next

	alertMsg = ""&SuccCNT&"건 수정 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam = "RegPoom") Then				''선택상품 품질인증정보 등록
	SuccCNT = 0
	FailCNT = 0
	arrItemid = split(arrItemid, ",")
	Dim w
	Dim qqqqqqqq
	For w = 0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(w))
		ret = regShoplinkerPoomOK(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next
	alertMsg = ""&SuccCNT&"건 등록 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam = "EditSellYn") Then			''선택상품 판매상태 변경
	Dim l
	arrItemid = split(arrItemid, ",")
	For l=0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(l))
		ret = editSellStatusShoplinkerOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam = "EditOutMall") Then			''선택상품 외부몰 상품 수정
	Dim b, c, mallprdidlist
	arrItemid = split(arrItemid, ",")
	For b=0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(b))
		mallprdidlist = getMallPrdid(iitemid)
		IF isArray(mallprdidlist) THEN
			For c = 0 To UBound(mallprdidlist,2)
				ret = editOutmallShoplinkerOneItem(iitemid, ierrStr, mallprdidlist(0,c)&"^*^*"&mallprdidlist(1,c))
				If (Not ret) Then
					rw ierrStr
				End If
			Next
		End If

'		ret = editOutmallShoplinkerOneItem(iitemid, ierrStr)
'		If (Not ret) Then
'			rw ierrStr
'		End If
	Next
ElseIf (cmdparam = "SearchITEM") Then			''선택상품코드 조회
	SuccCNT = 0
	FailCNT = 0
	Dim g
	arrItemid = split(arrItemid, ",")
	For g=0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(g))
		ret = ShoplinkerSearchItem(iitemid, ierrStr)
		If (Not ret) Then
			FailCNT = FailCNT + 1
			rw ierrStr
		Else
			SuccCNT = SuccCNT + 1
		End If
	Next

	alertMsg = ""&SuccCNT&"건 조회완료 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
End If

If (alertMsg <> "") Then
	IF (IsAutoScript) Then
		rw alertMsg
	Else
		response.write "<script>alert('"&alertMsg&"');</script>"
	End If
End if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->