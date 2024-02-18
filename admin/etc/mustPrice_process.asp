<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Function Math_RoundOff( su1, decimalPlaces)
    Dim sTemp, i, antilog, fraction '진수, 소수
    antilog = 1

    If decimalPlaces > 0 Then ' 10의 0승 이상 처리
        antilog = 10 ^ decimalPlaces
        sTemp = Fix( su1 / antilog + 0.5 ) * antilog
    Else ' 1 자리와 그 이하 처리
        sTemp = Round( su1 + 0.000001 , -(decimalPlaces))
    End if
    Math_RoundOff = sTemp
End Function


Dim mallgubun, itemid, mustPrice, mustBuyPrice, startDate, startDateTime, endDate, endDateTime, mode, i, idx, mustMargin, orgpricestartDate, orgpricestartDateTime, orgpriceendDate, orgpriceendDateTime
Dim sqlStr, cnt, AssignedRow, calcuBuyPrice, mallid
Dim arrDelItemid : arrDelItemid = request("cksel")
mallgubun   = request("mallgubun")
itemid      = request("itemid")
mustPrice   = request("mustPrice")
mustBuyPrice = request("mustBuyPrice")
startDate   = request("startDate")
startDateTime = request("startDateTime")
endDate     = request("endDate")
endDateTime = request("endDateTime")
orgpricestartDate   = request("orgpricestartDate")
orgpricestartDateTime = request("orgpricestartDateTime")
orgpriceendDate     = request("orgpriceendDate")
orgpriceendDateTime = request("orgpriceendDateTime")
mode        = request("mode")
idx         = request("idx")
mustMargin  = request("mustMargin")
startDate = startDate & " " & startDateTime
endDate = endDate & " " & endDateTime
mallid      = request("mallid")
If orgpricestartDate <> "" Then
    orgpricestartDate = orgpricestartDate & " " & orgpricestartDateTime
    orgpriceendDate = orgpriceendDate & " " & orgpriceendDateTime
End If
'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
arrItemid = Split(itemid, ",")

If mode = "I" Then
    If mustMargin = "" Then
        sqlStr = ""
        sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item "
        sqlStr = sqlStr & " WHERE itemid in ("& itemid &") "
        sqlStr = sqlStr & " and mwdiv <> 'M' "
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            cnt = rsget("cnt")
        rsget.Close

        If cnt > 0 Then
            Response.Write "<script language=javascript>alert('매입 아닌 상품에 특가마진이 입력되지 않았습니다.');self.close();</script>"
            dbget.close()	:	response.End
        End If
    End If

    If mallGubun = "nvstorefarm" AND (Trim(orgpricestartDate) = "" OR Trim(orgpriceendDate) = "") Then
        Response.Write "<script language=javascript>alert('정상가 판매기간을 확인해주세요');history.back();</script>"
        dbget.close()	:	response.End
    End If

    For i = 0 To Ubound(arrItemid)
        calcuBuyPrice = 0
        If mustMargin <> "" Then
            calcuBuyPrice = Math_RoundOff(mustPrice - (mustPrice * (mustMargin / 100)), 0)
        End If
        sqlStr = ""
        sqlStr = sqlStr & " IF EXISTS(SELECT TOP 1 itemid from db_etcmall.dbo.tbl_outmall_mustPriceItem WHERE itemid = '"& arrItemid(i) &"' and mallgubun = '"& mallGubun &"' ) "
        sqlStr = sqlStr & " BEGIN "
        sqlStr = sqlStr & "     UPDATE db_etcmall.dbo.tbl_outmall_mustPriceItem SET "
        sqlStr = sqlStr & "     mustPrice = '"& mustPrice &"' " & VBCRLF
        sqlStr = sqlStr & "     ,mustBuyPrice = '"& calcuBuyPrice &"' " & VBCRLF
        sqlStr = sqlStr & "     ,mustMargin = '"& mustMargin &"' " & VBCRLF
        sqlStr = sqlStr & "     ,startDate = '"& startDate &"' " & VBCRLF
        sqlStr = sqlStr & "     ,endDate = '"& endDate &"' " & VBCRLF
        If mallGubun = "nvstorefarm" Then
            sqlStr = sqlStr & "     ,orgpricestartDate = '"& orgpricestartDate &"' " & VBCRLF
            sqlStr = sqlStr & "     ,orgpriceendDate = '"& orgpriceendDate &"' " & VBCRLF
        End If
        sqlStr = sqlStr & "     ,lastUpdate = getdate() " & VBCRLF
        sqlStr = sqlStr & "     ,lastUpdateUserId = '"& session("ssBctID") &"' " & VBCRLF
        sqlStr = sqlStr & "     WHERE itemid = '"& arrItemid(i) &"' and mallgubun = '"& mallGubun &"' " & VBCRLF
        sqlStr = sqlStr & " END ELSE "
        sqlStr = sqlStr & " BEGIN "
        If mallGubun = "nvstorefarm" Then
            sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_outmall_mustPriceItem (mallgubun, itemid, mustPrice, mustBuyPrice, mustMargin, startDate, endDate, orgpricestartDate, orgpriceendDate, regDate, regUserId) VALUES " & VBCRLF
            sqlStr = sqlStr & " ('"& mallGubun &"', '"& arrItemid(i) &"', '"& mustPrice &"', '"& calcuBuyPrice &"', '"& mustMargin &"', '"& startDate &"', '"& endDate &"', '"& orgpricestartDate &"', '"& orgpriceendDate &"', getdate(), '"&  session("ssBctID") &"' ) "
        Else
            sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_outmall_mustPriceItem (mallgubun, itemid, mustPrice, mustBuyPrice, mustMargin, startDate, endDate, regDate, regUserId) VALUES " & VBCRLF
            sqlStr = sqlStr & " ('"& mallGubun &"', '"& arrItemid(i) &"', '"& mustPrice &"', '"& calcuBuyPrice &"', '"& mustMargin &"', '"& startDate &"', '"& endDate &"', getdate(), '"&  session("ssBctID") &"' ) "
        End If
        sqlStr = sqlStr & " END "
        dbget.Execute(sqlStr)
    Next
    Response.Write "<script language=javascript>alert('저장 하였습니다.');opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End
ElseIf mode = "U" Then
    If mallid = "nvstorefarm" AND (Trim(orgpricestartDate) = "" OR Trim(orgpriceendDate) = "") Then
        Response.Write "<script language=javascript>alert('정상가 판매기간을 확인해주세요');history.back();</script>"
        dbget.close()	:	response.End
    End If

    calcuBuyPrice = 0
    If mustMargin <> "" Then
        calcuBuyPrice = Math_RoundOff(mustPrice - (mustPrice * (mustMargin / 100)), 0)
    End If
    sqlStr = ""
    sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_outmall_mustPriceItem SET " & VBCRLF
    sqlStr = sqlStr & " mustPrice = '"& mustPrice &"' " & VBCRLF
    sqlStr = sqlStr & " ,mustBuyPrice = '"& calcuBuyPrice &"' " & VBCRLF
    sqlStr = sqlStr & " ,mustMargin = '"& mustMargin &"' " & VBCRLF
    sqlStr = sqlStr & " ,startDate = '"& startDate &"' " & VBCRLF
    sqlStr = sqlStr & " ,endDate = '"& endDate &"' " & VBCRLF
    If mallid = "nvstorefarm" Then
        sqlStr = sqlStr & " ,orgpricestartDate = '"& orgpricestartDate &"' " & VBCRLF
        sqlStr = sqlStr & " ,orgpriceendDate = '"& orgpriceendDate &"' " & VBCRLF
    End If
    sqlStr = sqlStr & " ,lastUpdate = getdate() " & VBCRLF
    sqlStr = sqlStr & " ,lastUpdateUserId = '"& session("ssBctID") &"' " & VBCRLF
    sqlStr = sqlStr & " WHERE idx = '"& idx &"' " & VBCRLF
    dbget.Execute(sqlStr)
    Response.Write "<script language=javascript>alert('수정 하였습니다.');opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End
ElseIf mode = "D" Then
	If Right(arrDelItemid,1) = "," Then arrDelItemid = Left(arrDelItemid, Len(arrDelItemid) - 1)
	sqlStr = ""
	sqlStr = sqlStr & " DELETE R " & VbCrlf
    sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem R " & VbCrlf
    sqlStr = sqlStr & " WHERE R.itemid in (" & arrDelItemid & ")" & VbCrlf
    sqlStr = sqlStr & " and mallgubun = '"& mallgubun &"' " & VbCrlf
	dbget.Execute sqlStr,AssignedRow
    Response.Write "<script language=javascript>alert('"& AssignedRow &"건 삭제 하였습니다.');parent.location.reload();</script>"
    dbget.close()	:	response.End
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->