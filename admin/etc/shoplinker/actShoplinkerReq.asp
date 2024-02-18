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

If (cmdparam = "RegSelect") Then				''���û�ǰ ���� ���
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

	alertMsg = ""&SuccCNT&"�� ��� "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"�� ���� "
	End If
ElseIf (cmdparam = "EDTSelect") Then			''���û�ǰ ���� ����
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

	alertMsg = ""&SuccCNT&"�� ���� "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"�� ���� "
	End If
ElseIf (cmdparam = "RegPoom") Then				''���û�ǰ ǰ���������� ���
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
	alertMsg = ""&SuccCNT&"�� ��� "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"�� ���� "
	End If
ElseIf (cmdparam = "EditSellYn") Then			''���û�ǰ �ǸŻ��� ����
	Dim l
	arrItemid = split(arrItemid, ",")
	For l=0 To UBound(arrItemid)
		iitemid = Trim(arrItemid(l))
		ret = editSellStatusShoplinkerOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam = "EditOutMall") Then			''���û�ǰ �ܺθ� ��ǰ ����
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
ElseIf (cmdparam = "SearchITEM") Then			''���û�ǰ�ڵ� ��ȸ
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

	alertMsg = ""&SuccCNT&"�� ��ȸ�Ϸ� "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"�� ���� "
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