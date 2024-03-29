<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 600 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/outMall/cjmall/cjmallitemcls_CT.asp"-->
<!-- #include virtual="/admin/outMall/cjmall/incCJmallFunction_CT.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),30)
Dim sday : sday = request("sday")
Dim cksel : cksel = request("cksel")
Dim subcmd : subcmd = request("subcmd")
Dim retFlag : retFlag = request("retFlag")
Dim iitemid, ret, sqlStr, AssignedRow, i
Dim alertMsg, ierrStr
Dim SuccCNT, FailCNT
dim todate, stdt, maxloop
dim ArrRows
Dim s

If (cmdparam="RegSelectWait") Then   ''선택상품 예정등록.
	Dim k, j, failid
	k = 0
	j = 0
	cksel = Trim(cksel)
	For i=0 to Ubound(Split(cksel,","))
		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid, isnull(c.infodiv,'') as Coninfodiv, m.infodiv as Mapinfodiv, m.cdmKey " & VBCRLF
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i  " & VBCRLF
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item_contents as c on i.itemid = c.itemid " & VBCRLF
		sqlStr = sqlStr & " JOIN db_outMall.dbo.tbl_cjMall_prdDiv_mapping as m on i.cate_large = m.tencatelarge and i.cate_mid = m.tencatemid and i.cate_small = m.tencatesmall and c.infodiv = m.infodiv " & VBCRLF
		sqlStr = sqlStr & " JOIN db_outMall.dbo.tbl_OutMall_CateMap_Summary as S on i.cate_large = S.tencatelarge and i.cate_mid = S.tencatemid and i.cate_small = S.tencatesmall " & VBCRLF
		sqlStr = sqlStr & " WHERE i.itemid = "&Split(cksel,",")(i) & VBCRLF
		sqlStr = sqlStr & " and S.mallid = 'cjmall' "
		rsCTget.Open sqlStr,dbCTget,1
		If not rsCTget.EOF Then
			If rsCTget("Coninfodiv") <> rsCTget("Mapinfodiv") Then			'tbl_item_contents에 infodiv가 없을수도 있어서
				k = k + 1
				failid = rsCTget("itemid") & "," & failid
			ElseIf rsCTget("Coninfodiv") = rsCTget("Mapinfodiv") Then		'tbl_cjmall_regItem에 infodiv와 cdm키 저장
				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_outMall.dbo.tbl_cjmall_regItem " & VBCRLF
				sqlStr = sqlStr & " (itemid, regdate, reguserid, cjmallStatCD, infodiv, cdmKey) values " & VBCRLF
				sqlStr = sqlStr & " ("&Split(cksel,",")(i)&", getdate(), '"&session("SSBctID")&"', 0, '"&rsCTget("Mapinfodiv")&"', '"&rsCTget("cdmKey")&"') " & VBCRLF
				dbCTget.Execute sqlStr
				j = j + 1
			End If
		End If
		rsCTget.Close
	Next
	response.write "<script>alert('"&j&" 건 예정등록됨.\n"&k&" 건 등록실패됨("&failid&") ');parent.location.reload();</script>"
ElseIf (cmdparam="DelSelectWait") Then   ''선택상품 예정등록.
	cksel = Trim(cksel)
	if Right(cksel,1)="," then cksel=Left(cksel,Len(cksel)-1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_outMall.dbo.tbl_cjmall_regItem " & VBCRLF
	sqlStr = sqlStr & " WHERE cjmallStatCD in (0, -1)" & VBCRLF
	sqlStr = sqlStr & " and itemid in ("&cksel&")"
	dbCTget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 예정 삭제됨.');parent.location.reload();</script>"
ElseIf (cmdparam="RegSelect") Then	''선택상품 실등록.
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

	alertMsg = ""&SuccCNT&"건 등록 "
	If (FailCNT > 0) Then
		alertMsg = alertMsg & ""&FailCNT&"건 실패 "
	End If
ElseIf (cmdparam="EditSelect") Then	''선택상품 정보 실 수정.

	cksel = split(cksel, ",")
	For s=0 To UBound(cksel)
		iitemid=Trim(cksel(s))
		ret = editCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="EditSelect2") Then	''선택상품 정보 실 수정 + 단품수정.

	cksel = split(cksel, ",")
	For s=0 To UBound(cksel)
		iitemid=Trim(cksel(s))
		ierrStr = ""
		ret = editCjmallOneItem(iitemid, ierrStr)                   ''상품정보
		If (Not ret) Then
			rw ierrStr
		End If
		ierrStr = ""
		ret = editqtyCjmallOneItem(iitemid, ierrStr, "")            ''수량수정
		If (Not ret) Then
			rw ierrStr
		End If
		ierrStr = ""
		ret = editDTCjmallOneItem(iitemid, ierrStr)                 ''단품판매여부
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="EdDateSel") Then	''선택상품 정보 실 수정.
	Dim v
	cksel = split(cksel,",")
	For y=0 To UBound(cksel)
		iitemid=Trim(cksel(v))
		ret = editDateCjmallOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="EditPriceSelect") Then	''선택상품 가격 실 수정.
	Dim p, qw
	cksel = split(cksel, ",")
'	For qw=0 To UBound(cksel)
'		iitemid=Trim(cksel(qw))
'		ret = editSellPriceCjmallOneItem(iitemid, ierrStr)
'		If (Not ret) Then
'			rw ierrStr
'		End If
'	Next
	For p=0 To UBound(cksel)
		iitemid=Trim(cksel(p))
		ret = editPriceCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next

ElseIf (cmdparam="EditPriceSelect2") Then	''선택상품 판매가 수정.(관리자만)
	Dim ww
	cksel = split(cksel, ",")
	For ww=0 To UBound(cksel)
		iitemid=Trim(cksel(ww))
		ret = editSellPriceCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next

	For ww=0 To UBound(cksel)
		iitemid=Trim(cksel(ww))
		ret = editPriceCjmallOneItem(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam = "EdSaleDTSel") Then ''선택상품 단품 수정.
	Dim e
	Dim tenOptCnt, cjOptCnt
	cksel = split(cksel, ",")
	For e = 0 To UBound(cksel)
		iitemid = Trim(cksel(e))
		ret = editDTCjmallOneItem(iitemid, ierrStr)
	Next

	if (retFlag<>"") then
        Response.Write "<script language=javascript>parent."&retFlag&";</script>"
        response.end
    end if
ElseIf (cmdparam = "EdPriceQtyDT") Then ''선택상품 가격/수량/단품판매/ 수정.
	cksel = split(cksel, ",")
	For e = 0 To UBound(cksel)
		iitemid = Trim(cksel(e))

		ierrStr = ""
		ret = editPriceCjmallOneItem(iitemid, ierrStr)              ''가격
		If (Not ret) Then
			rw ierrStr
		End If

		ierrStr = ""
		ret = editqtyCjmallOneItem(iitemid, ierrStr, "")            ''수량수정
		If (Not ret) Then
			rw ierrStr
		End If

		ierrStr = ""
		ret = editDTCjmallOneItem(iitemid, ierrStr)                 ''단품판매
		If (Not ret) Then
			rw ierrStr
		End If
	Next

	if (retFlag<>"") then
        Response.Write "<script language=javascript>parent."&retFlag&";</script>"
        response.end
    end if
ElseIf (cmdparam="LIST") Then  		''승인된 상품인지 총 기간동안 검색		http://testscm.10x10.co.kr/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST
	listCjMallItem(sday)
ElseIf (cmdparam="DayLIST") Then	''승인된 상품인지 일정기간동안 검색		http://testscm.10x10.co.kr/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday=0
	daylistCjMallItem(sday)
ELSEIF (cmdparam="confirmItem") Then    '' 선택상품 승인확인
    cksel = split(cksel,",")
	For i=0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		ret = oneCjMallItemConfirm(iitemid, ierrStr)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ELSEIF (cmdparam="confirmItemAuto") Then    '' 승인,판매상태 확인 Batch
    cksel = ""
    if (subcmd="1") then
        sqlStr = "select top 15 r.itemid "
        sqlStr = sqlStr & "	from db_outMall.dbo.tbl_cjmall_regitem r"
        sqlStr = sqlStr & "	Join db_AppWish.dbo.tbl_item i"
	    sqlStr = sqlStr & "	on r.itemid=i.itemid"
        sqlStr = sqlStr & "	where r.cjMallStatcd=3" ''-1: 등록실패 , 0: 등록예정, 1: 전송시도 , 3:승인대기
        '''sqlStr = sqlStr & "	and (i.sellyn<>'Y' or i.sellcash>isNULL(r.cjmallprice,0))"
        sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.cjmallsellyn='X' THEN '0' ELSE r.cjmallsellyn END), r.cjmallLastUpdate , r.itemid desc"

'        sqlStr = "select top 15 r.itemid "
'        sqlStr = sqlStr & "	from db_item.dbo.tbl_cjmall_regitem r"
'        sqlStr = sqlStr & "	Join db_item.dbo.tbl_item i"
'	    sqlStr = sqlStr & "	on r.itemid=i.itemid"
'        sqlStr = sqlStr & "	where r.cjMallStatcd>3" ''-1: 등록실패 , 0: 등록예정, 1: 전송시도 , 3:승인대기
'        sqlStr = sqlStr & "	and i.optionCnt>0"
'        '''sqlStr = sqlStr & "	and (i.sellyn<>'Y' or i.sellcash>isNULL(r.cjmallprice,0))"
'        sqlStr = sqlStr & "	order by r.regedOptCnt, r.lastStatCheckDate, (CASE WHEN r.cjmallsellyn='X' THEN '0' ELSE r.cjmallsellyn END), r.cjmallLastUpdate , r.itemid desc"


    else
        sqlStr = "select top 15 r.itemid "
        sqlStr = sqlStr & "	from db_outMall.dbo.tbl_cjmall_regitem r"
        sqlStr = sqlStr & "	where cjMallStatcd>0" ''-1: 등록실패 , 0: 등록예정, 1: 전송시도
       '' sqlStr = sqlStr & "	and lastPriceCheckDate>'2013-11-03' and lastPriceCheckDate<'2013-11-21 16:00:00' and lastStatCheckDate<'2013-11-22 12:00:00'" ''임시
        ''sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.cjmallsellyn='X' THEN '0' ELSE r.cjmallsellyn END), r.cjmallLastUpdate , r.itemid desc"
        sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.cjmallsellyn='X' THEN '0' ELSE r.cjmallsellyn END), r.cjmallLastUpdate , r.itemid desc"
    end if

    rsCTget.Open sqlStr,dbCTget,1
    if not rsCTget.Eof then
        ArrRows = rsCTget.getRows()
    end if
    rsCTget.close

    if isArray(ArrRows) then
        For i =0 To UBound(ArrRows,2)
            cksel = cksel + CStr(ArrRows(0,i)) + ","
        Next
    else
        rw "S_NONE"
        dbCTget.Close() : response.end
    end if

    cksel = split(cksel,",")
	For i=0 To UBound(cksel)
		iitemid=Trim(cksel(i))
		if (iitemid<>"") then
    		ret = oneCjMallItemConfirm(iitemid, ierrStr)
    		If (Not ret) Then
    			rw ierrStr
    		End If
        end if
	Next
ElseIf (cmdparam="EditSellYn") Then ''선택상품 판매상태 수정
	Dim l
	cksel = split(cksel,",")
	For l=0 To UBound(cksel)
		iitemid=Trim(cksel(l))
		ret = editSellStatusCjmallOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="EditQty") Then ''선택상품 수량 수정
	Dim y
	cksel = split(cksel,",")
	For y=0 To UBound(cksel)
		iitemid=Trim(cksel(y))
		ret = editqtyCjmallOneItem(iitemid, ierrStr, subcmd)
		If (Not ret) Then
			rw ierrStr
		End If
	Next
ElseIf (cmdparam="cjmallOrdreg") Then ''주문목록 조회

    todate = LEFT(CStr(now()),10)
    maxloop = 4 ''추석때 발주 안함 추석이후 영업일에 주문 통보됨// 연휴시 이값을 길게 할것. (일반적으로 3,4일 적당)
    stdt = getLastOrderInputDT()
    sday = stdt
    for i=0 to maxloop
        rw sday & "주문건 등록시작 ======================================"
    	call getCjOrderList("ORDLIST", sday)

    	'' 취소쪽은 CS건에서함. 주석처리 2013/08/05
    	''rw sday & "주문취소건 등록시작 ======================================"
    	''call getCjOrderList("ORDCANCELLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallOrdUp") Then ''주문목록 실판매가 업데이트

    todate = LEFT(CStr(now()),10)
    maxloop = 1
    stdt = getLastOrderInputDTUp()
    if (request("stdt")<>"") then stdt=request("stdt")
    rw stdt
    if stdt>"2014-08-28" then response.end

    sday = stdt
    for i=0 to maxloop-1
        rw sday & "주문건 등록시작 ======================================"
    	call getCjOrderList("ORDLISTUP", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
    rw "<form name=frmR method=post action=''><input type='hidden' name='cmdparam' value='cjmallOrdUp'><input type='hidden' name='stdt' value='"&sday&"'><input type='button' name='reloadBtn' value='reload' onClick='document.frmR.submit();'></form>"

    if (sday<"2014-08-27") then
    response.write "<script>"
    response.write "setTimeout(function(){document.frmR.submit();},2000);"
    response.write "</script>"
    end if

ElseIf (cmdparam="cjmallCsreg") Then ''CS목록 조회
    todate = LEFT(CStr(now()),10)
    maxloop = 10
    stdt = LEFT(CStr(DATEADD("d",-1,now())),10)
    sday = stdt
    for i=0 to maxloop
        rw sday & "CS건 조회 등록시작 ======================================"
    	call getCjCsList("CSLIST", sday)

    	sday = left(CStr(dateadd("d",1,sday)),10)

    	if (CDate(sday)>CDate(todate)) then Exit For
    	rw ""
    next
ElseIf (cmdparam="cjmallCancelreg") Then ''주문취소목록 조회
    sday = LEFT(CStr(now()),10)
	call getCjOrderCancelList(sday)

	rw sday
ElseIf (cmdparam="cjmallCommonCode") Then ''공통코드 조회
	Dim ccd
	ccd = request("CommCD")
	call getcjCommonCodeList(ccd)
Else
	rw "미지정 ["&cmdparam&"]"
End If

If (alertMsg <> "") Then
	IF (IsAutoScript) Then
		rw alertMsg
	Else
		response.write "<script>alert('"&alertMsg&"');</script>"
	End If
End if
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->