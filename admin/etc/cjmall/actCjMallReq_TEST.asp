<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall/cjmallitemcls.asp"-->
<!-- #include virtual="/admin/etc/cjmall/incCJmallFunction_TEST.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),16)
Dim sday : sday = request("sday")
Dim cksel : cksel = request("cksel")
Dim subcmd : subcmd = request("subcmd")
Dim iitemid, ret, sqlStr, AssignedRow, i
Dim alertMsg, ierrStr
Dim SuccCNT, FailCNT
dim todate, stdt, maxloop

If (cmdparam="RegSelectWait") Then   ''���û�ǰ �������.
	Dim k, j, failid
	k = 0
	j = 0
	cksel = Trim(cksel)
	For i=0 to Ubound(Split(cksel,","))
		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid, isnull(c.infodiv,'') as Coninfodiv, m.infodiv as Mapinfodiv, m.cdmKey " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i  " & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid " & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_cjMall_prdDiv_mapping as m on i.cate_large = m.tencatelarge and i.cate_mid = m.tencatemid and i.cate_small = m.tencatesmall and c.infodiv = m.infodiv " & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as S on i.cate_large = S.tencatelarge and i.cate_mid = S.tencatemid and i.cate_small = S.tencatesmall " & VBCRLF
		sqlStr = sqlStr & " WHERE i.itemid = "&Split(cksel,",")(i) & VBCRLF
		sqlStr = sqlStr & " and S.mallid = 'cjmall' "
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			If rsget("Coninfodiv") <> rsget("Mapinfodiv") Then			'tbl_item_contents�� infodiv�� �������� �־
				k = k + 1
				failid = rsget("itemid") & "," & failid
			ElseIf rsget("Coninfodiv") = rsget("Mapinfodiv") Then		'tbl_cjmall_regItem�� infodiv�� cdmŰ ����
				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_cjmall_regItem " & VBCRLF
				sqlStr = sqlStr & " (itemid, regdate, reguserid, cjmallStatCD, infodiv, cdmKey) values " & VBCRLF
				sqlStr = sqlStr & " ("&Split(cksel,",")(i)&", getdate(), '"&session("SSBctID")&"', 0, '"&rsget("Mapinfodiv")&"', '"&rsget("cdmKey")&"') " & VBCRLF
				dbget.Execute sqlStr
				j = j + 1
			End If
		End If
		rsget.Close
	Next
	response.write "<script>alert('"&j&" �� ������ϵ�.\n"&k&" �� ��Ͻ��е�("&failid&") ');parent.location.reload();</script>"
ElseIf (cmdparam="DelSelectWait") Then   ''���û�ǰ �������.
	cksel = Trim(cksel)
	if Right(cksel,1)="," then cksel=Left(cksel,Len(cksel)-1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_cjmall_regItem " & VBCRLF
	sqlStr = sqlStr & " WHERE cjmallStatCD in (0, -1)" & VBCRLF
	sqlStr = sqlStr & " and itemid in ("&cksel&")"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� ���� ������.');parent.location.reload();</script>"
ElseIf (cmdparam="RegSelect") Then	''���û�ǰ �ǵ��.
	SuccCNT = 0
	FailCNT = 0
	cksel = split(cksel, ",")
	Dim q
	For q = 0 To UBound(cksel)
		iitemid = Trim(cksel(q))
		ret = regCjMallOneItem(iitemid, ierrStr)
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
ElseIf (cmdparam="EditSelect") Then	''���û�ǰ ���� �� ����.
	Dim s
	cksel = split(cksel, ",")
	For s=0 To UBound(cksel)
		iitemid=Trim(cksel(s))
		ret = editCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="EditPriceSelect") Then	''���û�ǰ ���� �� ����.
	Dim p
	cksel = split(cksel, ",")
	For p=0 To UBound(cksel)
		iitemid=Trim(cksel(p))
		ret = editPriceCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam = "EdSaleDTSel") Then ''���û�ǰ ��ǰ ����.
	Dim e
	Dim tenOptCnt, cjOptCnt
	cksel = split(cksel, ",")
	For e = 0 To UBound(cksel)
		iitemid = Trim(cksel(e))
		ret = editDTCjmallOneItem(iitemid, ierrStr)
	Next
ElseIf (cmdparam="LIST") Then  		''���ε� ��ǰ���� �� �Ⱓ���� �˻�		http://testscm.10x10.co.kr/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST
	listCjMallItem()
ElseIf (cmdparam="DayLIST") Then	''���ε� ��ǰ���� �����Ⱓ���� �˻�		http://testscm.10x10.co.kr/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday=0
	daylistCjMallItem(sday)
ElseIf (cmdparam="EditSellYn") Then ''���û�ǰ �ǸŻ��� ����
	Dim l
	cksel = split(cksel,",")
	For l=0 To UBound(cksel)
		iitemid=Trim(cksel(l))
		ret = editSellStatusCjmallOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="EditQty") Then ''���û�ǰ ���� ����
	Dim y
	cksel = split(cksel,",")
	For y=0 To UBound(cksel)
		iitemid=Trim(cksel(y))
		ret = editqtyCjmallOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="cjmallOrdreg") Then ''�ֹ���� ��ȸ

    todate = LEFT(CStr(now()),10)
    maxloop = 10
    stdt = getLastOrderInputDT()
    sday = stdt
    for i=0 to maxloop
        rw sday & "�ֹ��� ��Ͻ��� ======================================"
    	call getCjOrderList("ORDLIST", sday)
    	rw sday & "�ֹ���Ұ� ��Ͻ��� ======================================"
    	call getCjOrderList("ORDCANCELLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallCsreg1") Then ''CS��� ��ȸ(��ǰ)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15�� �̻� CS�� ������ �� ���� CS������ �������� ���Ѵ�.

	'// ========================================================================
    stdt = getLastCSInputDT("return")
	''rw stdt
    sday = stdt
    for i=0 to maxloop
        rw sday & " CS��[ȸ������] ��ȸ ��Ͻ��� ======================================"
    	call getCjCsList("CSLIST", sday)
		Call UpdateLastCSInputDT("return", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
		rw ""
    next

ElseIf (cmdparam="cjmallCsreg2") Then ''CS��� ��ȸ(�ֹ����)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15�� �̻� CS�� ������ �� ���� CS������ �������� ���Ѵ�.

	'// ========================================================================
    stdt = getLastCSInputDT("ordercancel")
    sday = stdt
    for i=0 to maxloop
		rw sday & " CS��[�ֹ�����:���] ��ȸ ��Ͻ��� ======================================"
    	call getCjCsListInOrder("CSORDCANCELLIST", sday)
		Call UpdateLastCSInputDT("ordercancel", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next

ElseIf (cmdparam="cjmallCsreg3") Then ''CS��� ��ȸ(CS��� : ��ȯ��� ��)
    todate = LEFT(CStr(now()),10)
    maxloop = 15					'// TODO : 15�� �̻� CS�� ������ �� ���� CS������ �������� ���Ѵ�.

	'// ========================================================================
    stdt = getLastCSInputDT("order")
    sday = stdt
    for i=0 to maxloop
        rw sday & " CS��[�ֹ�����:���,������] ��ȸ ��Ͻ��� ======================================"
    	call getCjCsListInOrder("CSORDLIST", sday)
		Call UpdateLastCSInputDT("order", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next

ElseIf (cmdparam="cjmallCancelreg") Then ''�ֹ���Ҹ�� ��ȸ
    sday = LEFT(CStr(now()),10)
	call getCjOrderCancelList(sday)

	rw sday
ElseIf (cmdparam="cjmallCommonCode") Then ''�����ڵ� ��ȸ
	Dim ccd
	ccd = request("CommCD")
	call getcjCommonCodeList(ccd)
Else
	rw "������ ["&cmdparam&"]"
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