<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매입상품원가관리
' History : 2022.01.17 이상구 생성
'           2022.08.19 한용민 수정(세금계산서 내용 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<%

dim mode, indt, updt, deldt, didx, arrDidx, arrReportNo, arrReportPrice, result, i, oCPurchasedProduct
dim idx, codeList, reportIdx, reportNo, reportPrice, totalNo, totalPrice, reguserid, regusername
dim ordercode, sqlStr, REPORT_EXIST, anbunSuplyPrice, anbunVatPrice, anbunBuyPrice, detailidx
dim ppMasterIdx, ppGubun, groupCode, anbunType, suplyPrice, vatPrice, yyyymm, refer, PriceEditCount
dim taxregdate, taxlinkidx, billsiteCode, neotaxno, eseroEvalSeq, AssignedRow, suplyPriceSum, vatPriceSum, buyPriceSum
dim title, productidx, existsReportCount
    mode = requestCheckVar(request("mode"), 32)
    idx = requestCheckVar(request("idx"), 32)
    ordercode = requestCheckVar(request("ordercode"), 32)
    reguserid = session("ssBctId")
    regusername = html2db(session("ssBctCname"))
    reportIdx = requestCheckVar(getNumeric(trim(request("reportIdx"))), 10)
    ppMasterIdx = requestCheckVar(getNumeric(trim(request("ppMasterIdx"))), 10)
    ppGubun = requestCheckVar(request("ppGubun"), 32)
    groupCode = requestCheckVar(request("groupCode"), 32)
    anbunType = requestCheckVar(request("anbunType"), 32)
    suplyPrice = requestCheckVar(request("suplyPrice"), 32)
    vatPrice = requestCheckVar(request("vatPrice"), 32)
    yyyymm = requestCheckVar(request("yyyymm"), 32)
    anbunSuplyPrice = request("anbunSuplyPrice")   ' requestcheckvar 쓰지 말것..짤림
    anbunVatPrice = request("anbunVatPrice")   ' requestcheckvar 쓰지 말것..짤림
    anbunBuyPrice = request("anbunBuyPrice")   ' requestcheckvar 쓰지 말것..짤림
    detailidx = request("detailidx")   ' requestcheckvar 쓰지 말것..짤림
    taxregdate = requestCheckVar(trim(request("taxregdate")),10)
    taxlinkidx = requestCheckVar(getNumeric(trim(request("taxlinkidx"))),10)
    billsiteCode = requestCheckVar(trim(request("billsiteCode")),2)
    neotaxno = requestCheckVar(trim(request("neotaxno")),30)
    eseroEvalSeq = requestCheckVar(trim(request("eseroEvalSeq")),24)
    title = requestCheckVar(trim(request("title")),128)
    productidx = requestCheckVar(getNumeric(trim(request("productidx"))), 10)

REPORT_EXIST = False
if (reportIdx = "") then
    reportIdx = "0"
end if
existsReportCount=0
if (reportIdx <> 0) then
    '// 품의서 있음
    REPORT_EXIST = True
end if

refer = request.ServerVariables("HTTP_REFERER")

if (mode = "insmaster") then
    sqlStr = " insert into [db_storage].[dbo].[tbl_pp_product_master]( " & vbcrlf
    sqlStr = sqlStr & " 	codeList, reportIdx, reportNo, reportPrice, orderNo, orderPrice, ipgoNo, ipgoPrice, reguserid, regusername, title" & vbcrlf
    sqlStr = sqlStr & " ) " & vbcrlf
    sqlStr = sqlStr & " values('', 0, 0, 0, 0, 0, 0, 0, '" & reguserid & "', '" & regusername & "', convert(nvarchar(128),N'"& html2db(title) &"'))"
    dbget.Execute sqlStr

    sqlStr ="select SCOPE_IDENTITY() "

	rsget.open sqlStr,dbget,2
	IF not rsget.Eof Then
		idx = rsget(0)
	End IF
	rsget.close

    if (ordercode <> "") then
        if not(CheckOrderCodeBusinessNumberExists(ordercode,idx)) then
            response.write "<script type='text/javascript'>"
            response.write "    alert('주문서에 연결된 브랜드가 텐바이텐 사업자가 아닙니다.');"
            response.write "</script>"
            response.write "주문서에 연결된 브랜드가 텐바이텐 사업자가 아닙니다."
            dbget.close()	:	response.End
        end if

        if CheckOrderCodeExists(idx, ordercode) = True then
            response.write "<script language='javascript'>"
            response.write "alert('다른 품의자료에 연결된 주문서입니다.');"
            response.write "</script>"
            response.write "에러 : 다른 품의자료에 연결된 주문서입니다."
            dbget.close()	:	response.End
        end if

        Call AddOrderCode(idx, ordercode)
    end if
    Call UpdateOrderCodeList(idx)
    Call UpdateMasterInfo(idx)

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "location.replace('/admin/newstorage/PurchasedProductList.asp?menupos=" & request("menupos") & "');"
    response.write "</script>"

elseif (mode = "modimaster") then
    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_master] " & vbcrlf
    sqlStr = sqlStr & " set reguserid = '" & reguserid & "', regusername = '" & regusername & "', updt = getdate()," & vbcrlf
    sqlStr = sqlStr & " title=convert(nvarchar(128),N'"& html2db(title) &"') where" & vbcrlf
    sqlStr = sqlStr & " idx = " & idx
    dbget.Execute sqlStr

    set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = idx
    oCPurchasedProduct.GetPurchasedProductMaster

    if oCPurchasedProduct.FOneItem.FreportIdx <> 0 then
        REPORT_EXIST = True
    end if

    'if Not REPORT_EXIST then
        arrDidx = request("didx")   ' requestcheckvar 쓰지 말것..짤림
        arrReportNo = request("reportNo")   ' requestcheckvar 쓰지 말것..짤림
        arrReportPrice = request("reportPrice")   ' requestcheckvar 쓰지 말것..짤림

        '' response.write arrDidx & "<br />"
        '' response.write arrReportNo & "<br />"
        '' response.write arrReportPrice & "<br />"

        arrDidx = Split(arrDidx, ",")
        arrReportNo = Split(arrReportNo, ",")
        arrReportPrice = Split(arrReportPrice, ",")

        '' response.write UBound(arrDidx) & "<br />"

        for i = 0 to UBound(arrDidx)
            if Trim(arrDidx(i)) <> "" then
                sqlStr = " update [db_storage].[dbo].[tbl_pp_product_item_detail] "
                sqlStr = sqlStr & " set reportNo = " & arrReportNo(i) & ", reportPrice = " & arrReportPrice(i) & " "
                sqlStr = sqlStr & " where idx = " & arrDidx(i) & " "
                ''response.write sqlStr & "<br />"
                dbget.Execute sqlStr
            end if
        next
    'end if

    if (ordercode <> "") then
        if not(CheckOrderCodeBusinessNumberExists(ordercode,idx)) then
            response.write "<script type='text/javascript'>"
            response.write "    alert('주문서에 연결된 브랜드가 텐바이텐 사업자가 아닙니다.');"
            response.write "</script>"
            response.write "주문서에 연결된 브랜드가 텐바이텐 사업자가 아닙니다."
            dbget.close()	:	response.End
        end if

        if CheckOrderCodeExists(idx, ordercode) = True then
            response.write "<script language='javascript'>"
            response.write "alert('다른 품의자료에 연결된 주문서입니다.');"
            response.write "</script>"
            response.write "에러 : 다른 품의자료에 연결된 주문서입니다."
            dbget.close()	:	response.End
        end if

        Call AddOrderCode(idx, ordercode)
    end if
    Call UpdateOrderCodeList(idx)
    Call UpdateMasterInfo(idx)

    '' dbget.close()	:	response.End

elseif (mode = "taxregchange") then
	sqlStr = "update db_storage.dbo.tbl_pp_product_sheet_master"
	sqlStr = sqlStr & " set taxregdate='" & taxregdate & "'"
	sqlStr = sqlStr & " ,taxinputdate=getdate()"
	sqlStr = sqlStr & " ,finishflag=(CASE WHEN finishflag='1' THEN '3' ELSE finishflag END)"

	IF (neotaxno<>"") or (taxlinkidx="") then
	    sqlStr = sqlStr & " ,neotaxno='"&neotaxno&"'"
	    sqlStr = sqlStr & " ,billsiteCode='"&billsiteCode&"'"+ VbCrlf
    end if

    sqlStr = sqlStr & " ,eseroEvalSeq='"&eseroEvalSeq&"' where"+ VbCrlf
	sqlStr = sqlStr & " idx=" + CStr(idx)

    'response.write sqlStr & "<Br>"
    dbget.execute sqlStr

	if (taxlinkidx="") then
	    sqlStr = " exec db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne] '"&eseroEvalSeq&"',4,"&idx&""

        'response.write sqlStr & "<br>"
	    dbget.Execute sqlStr,AssignedRow
	    ''if (AssignedRow<1) then AssignedRow=0
	    ''response.write "<script>alert('Tax 매핑 : "&AssignedRow&" 건');</script>"
	end if

elseif mode="delTaxInfo" then
	sqlstr = " update db_storage.dbo.tbl_pp_product_sheet_master" + VbCrlf
    sqlstr = sqlstr + " set taxlinkidx=NULL"
    sqlstr = sqlstr + " ,neotaxno=NULL"
    sqlstr = sqlstr + " ,eseroevalseq=NULL"
    sqlstr = sqlstr + " ,taxregdate=NULL"
    sqlstr = sqlstr + " ,taxinputdate=NULL"
    sqlstr = sqlstr + " ,billsitecode=NULL"
    sqlstr = sqlstr + " where idx=" + CStr(idx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag in ('0','1')"  + VbCrlf

    'response.write sqlStr & "<Br>"
    dbget.execute sqlStr

elseif (mode = "rmordr") then
    '// 주문서 제외
    if (ordercode <> "") then
        Call DelOrderCode(idx, ordercode)
        Call UpdateOrderCodeList(idx)
        Call UpdateMasterInfo(idx)
    end if

elseif (mode = "rmsheetordr") then
    '// 주문서 제외
    if (ordercode <> "") then
        Call DelOrderCodeFromSheet(idx, ordercode)

        set oCPurchasedProduct = new CPurchasedProduct
        oCPurchasedProduct.FRectIdx = CHKIIF(idx="", "-1", idx)
        oCPurchasedProduct.GetPurchasedProductSheetMaster

        oCPurchasedProduct.FRectIdx = oCPurchasedProduct.FOneItem.FppMasterIdx
        oCPurchasedProduct.GetPurchasedProductMaster
        ppMasterIdx = oCPurchasedProduct.FOneItem.Fidx

        if oCPurchasedProduct.FOneItem.FreportIdx <> 0 then
            REPORT_EXIST = True
        end if

        sqlStr = " select count(m.idx) as cnt "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_storage].[dbo].[tbl_pp_product_sheet_master] m "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & "     and m.ppMasterIdx = " & ppMasterIdx
        sqlStr = sqlStr & "     and (m.codeList like '%" & ordercode & ",%' or m.codeList like '%" & ordercode & "') "		'// 쉼표있거나 마지막 내역
        sqlStr = sqlStr & "     and m.deldt is NULL "

	    rsget.Open sqlStr, dbget, 1
	    if Not rsget.Eof then
		    if (rsget("cnt") <> 0) then
                ordercode = ""
            end if
	    end if
	    rsget.Close

        if ordercode <> "" then
            Call DelOrderCode(ppMasterIdx, ordercode)
            Call UpdateOrderCodeList(ppMasterIdx)
        end if

        Call UpdateSheetDetail(idx)
    end if

    response.write "<script language='javascript'>"
    response.write "opener.location.reload(); "
    response.write "</script>"

elseif (mode = "delmaster") then

    if REPORT_EXIST then
	    response.write "<script language='javascript'>"
	    response.write "alert('에러 : 품의번호가 있습니다.');"
	    response.write "</script>"
        response.write "에러 : 품의번호가 있습니다."
        dbget.close()	:	response.End
    end if

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_master] "
    sqlStr = sqlStr & " set reguserid = '" & reguserid & "', regusername = '" & regusername & "', deldt = getdate() "
    sqlStr = sqlStr & " where idx = " & idx
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('삭제 되었습니다.');"
    response.write "location.replace('/admin/newstorage/PurchasedProductList.asp?menupos=" & request("menupos") & "');"
    response.write "</script>"

elseif (mode = "inssheetmaster") then


    sqlStr = " insert into [db_storage].[dbo].[tbl_pp_product_sheet_master]( "
    sqlStr = sqlStr & " 	ppMasterIdx, yyyymm, codeList, ppGubun, groupCode, anbunType, suplyPrice, vatPrice, buyPrice "
    sqlStr = sqlStr & " ) "
    sqlStr = sqlStr & " values('" & ppMasterIdx & "', '" & yyyymm & "', '', '" & ppGubun & "', '" & groupCode & "', '" & anbunType & "', '" & suplyPrice & "', '" & vatPrice & "', '" & (suplyPrice*1 + vatPrice*1) & "') "

    dbget.Execute(sqlStr)

    sqlStr ="select SCOPE_IDENTITY() "

	rsget.open sqlStr,dbget,2
	IF not rsget.Eof Then
		idx = rsget(0)
	End IF
	rsget.close

    set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = ppMasterIdx
    oCPurchasedProduct.GetPurchasedProductMaster

    if oCPurchasedProduct.FOneItem.FreportIdx <> 0 then
        REPORT_EXIST = True
    end if

    if (ordercode <> "") then
        '// 주문서 기준 업데이트
        Call AddOrderCodeToSheet(idx, ordercode)
        Call AddOrderCode(ppMasterIdx, ordercode)

        Call UpdateOrderCodeList(ppMasterIdx)
        Call UpdateSheetDetail(idx)
    elseif (yyyymm <> "") then
        '// 입고월 기준 업데이트
        Call UpdateSheetDetailByMonth(idx, ppMasterIdx, yyyymm)
    end if

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "opener.location.reload(); "
    response.write "location.href = 'PurchasedProductSheetModify.asp?idx=" & idx & "';"
    response.write "</script>"
    dbget.close()	:	response.End

elseif (mode = "modisheetmaster") then

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_sheet_master] "
    sqlStr = sqlStr & " set updt = getdate(), yyyymm = '" & yyyymm & "', ppGubun = '" & ppGubun & "', groupCode = '" & groupCode & "', anbunType = '" & anbunType & "' "
    sqlStr = sqlStr & " , suplyPrice = '" & suplyPrice & "', vatPrice = '" & vatPrice & "', buyPrice = '" & (suplyPrice*1 + vatPrice*1) & "' "
    sqlStr = sqlStr & " where idx = " & idx
    ''response.write sqlStr
    dbget.Execute(sqlStr)

    set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProduct.GetPurchasedProductSheetMaster

    oCPurchasedProduct.FRectIdx = oCPurchasedProduct.FOneItem.FppMasterIdx
    oCPurchasedProduct.GetPurchasedProductMaster
    ppMasterIdx = oCPurchasedProduct.FOneItem.Fidx

    if (anbunType = "G203") then
        ''anbunSuplyPrice, anbunVatPrice, anbunBuyPrice, detailidx

        anbunSuplyPrice = Split(anbunSuplyPrice, ",")
        anbunVatPrice = Split(anbunVatPrice, ",")
        anbunBuyPrice = Split(anbunBuyPrice, ",")
        detailidx = Split(detailidx, ",")

        for i = 0 to UBound(anbunSuplyPrice)
            anbunSuplyPrice(i) = Trim(anbunSuplyPrice(i))
            anbunVatPrice(i) = Trim(anbunVatPrice(i))
            anbunBuyPrice(i) = Trim(anbunBuyPrice(i))
            detailidx(i) = Trim(detailidx(i))
            if anbunSuplyPrice(i) <> "" and anbunVatPrice(i) <> "" and anbunBuyPrice(i) <> "" and detailidx(i) <> "" then
                if anbunBuyPrice(i)=0 then
                    anbunSuplyPrice(i)=0
                    anbunVatPrice(i)=0
                end if
                sqlStr = " update [db_storage].[dbo].[tbl_pp_product_sheet_detail] "
                sqlStr = sqlStr & " set suplyPriceSum = " & anbunSuplyPrice(i) & ", vatPriceSum = " & anbunVatPrice(i) & ", buyPriceSum = " & anbunBuyPrice(i)
                sqlStr = sqlStr & " where idx = " & detailidx(i)
                ''response.write sqlStr
                dbget.Execute sqlStr

            end if
        next

        suplyPriceSum=0
        vatPriceSum=0
        buyPriceSum=0
        sqlStr = " select isnull(sum(sd.suplyPriceSum),0) as suplyPriceSum, isnull(sum(sd.vatPriceSum),0) as vatPriceSum, isnull(sum(sd.buyPriceSum),0) as buyPriceSum"
        sqlStr = sqlStr & " from db_storage.dbo.tbl_pp_product_sheet_detail sd with (nolock)"
        sqlStr = sqlStr & " join db_storage.dbo.tbl_pp_product_sheet_master sm with (nolock)"
        sqlStr = sqlStr & " 	on sd.masterIdx=sm.idx"
        sqlStr = sqlStr & " 	and sd.ordercode=sm.yyyymm"
        sqlStr = sqlStr & " 	and sm.deldt is NULL"
        sqlStr = sqlStr & " where sd.deldt is NULL and sd.masteridx = " & idx

        'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if Not rsget.Eof then
            suplyPriceSum = rsget("suplyPriceSum")
            vatPriceSum = rsget("vatPriceSum")
            buyPriceSum = rsget("buyPriceSum")
        end if
        rsget.Close

        'response.write "suplyPriceSum:" & suplyPriceSum & "<Br>"
        'response.write "suplyPrice:" & suplyPrice & "<Br>"
        'response.write "vatPriceSum:" & vatPriceSum & "<Br>"
        'response.write "vatPrice:" & vatPrice & "<Br>"
        'response.write "buyPriceSum:" & buyPriceSum & "<Br>"
        'response.write "suplyPrice*1 + vatPrice:" & suplyPrice*1 + vatPrice*1 & "<Br>"

        PriceEditCount="0"
        ' 디테일의 합계값과 마스터의 합계값이 안맞을 경우 합산해서 마스터 엎어침    ' 2022.08.23 한용민
        if (suplyPriceSum<>0 and suplyPriceSum<>suplyPrice) or (vatPriceSum<>0 and vatPriceSum<>vatPrice) or (buyPriceSum<>0 and buyPriceSum<>suplyPrice*1 + vatPrice*1) then
            sqlStr = "update db_storage.dbo.tbl_pp_product_sheet_master set" & vbcrlf

            if (suplyPriceSum<>0 and suplyPriceSum<>suplyPrice) then
                sqlStr = sqlStr & " suplyPrice="& suplyPriceSum &"" & vbcrlf
                PriceEditCount=PriceEditCount+1
            end if
            if (vatPriceSum<>0 and vatPriceSum<>vatPrice) then
                if PriceEditCount>0 then sqlStr = sqlStr & " , "
                sqlStr = sqlStr & " vatPrice="& vatPriceSum &"" & vbcrlf
                PriceEditCount=PriceEditCount+1
            end if
            if (buyPriceSum<>0 and buyPriceSum<>suplyPrice*1 + vatPrice*1) then
                if PriceEditCount>0 then sqlStr = sqlStr & " , "
                sqlStr = sqlStr & " buyPrice="& buyPriceSum &"" & vbcrlf
            end if

            sqlStr = sqlStr & " where idx="& idx &""

            'response.write sqlStr & "<Br>"
            dbget.execute sqlStr
        end if
    end if

    if oCPurchasedProduct.FOneItem.FreportIdx <> 0 then
        REPORT_EXIST = True
    end if

    if (ordercode <> "") then
        '// 주문서 기준 업데이트
        Call AddOrderCodeToSheet(idx, ordercode)

        Call AddOrderCode(ppMasterIdx, ordercode)

        Call UpdateOrderCodeList(ppMasterIdx)
        Call UpdateSheetDetail(idx)
    elseif (yyyymm <> "") then
        '// 입고월 기준 업데이트
        Call UpdateSheetDetailByMonth(idx, ppMasterIdx, yyyymm)
    end if

    response.write "<script language='javascript'>"
    response.write "opener.location.reload(); "
    response.write "location.href = 'PurchasedProductSheetModify.asp?idx=" & idx & "';"
    response.write "</script>"
    dbget.close()	:	response.End

elseif (mode = "delsheetmaster") then

    set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProduct.GetPurchasedProductSheetMaster

    oCPurchasedProduct.FRectIdx = oCPurchasedProduct.FOneItem.FppMasterIdx
    oCPurchasedProduct.GetPurchasedProductMaster
    ppMasterIdx = oCPurchasedProduct.FOneItem.Fidx

    sqlStr = " update [db_storage].[dbo].[tbl_pp_product_sheet_master] "
    sqlStr = sqlStr & " set deldt = getdate() "
    sqlStr = sqlStr & " where idx = " & idx
    dbget.Execute(sqlStr)

    Call UpdateSheetDetailByMonth(idx, ppMasterIdx, yyyymm)

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "opener.location.reload(); opener.focus(); window.close(); "
    response.write "</script>"
    dbget.close()	:	response.End

elseif (mode = "doapplycogs") then

    sqlStr = " exec [db_storage].[dbo].[usp_Ten_PP_cogs_Update] " & idx & ", 'Y' "
    dbget.Execute(sqlStr)

    Call UpdateOrderCodeList(idx)
    Call UpdateMasterInfo(idx)

elseif (mode = "doapplyipgocogs") then

    sqlStr = " exec [db_storage].[dbo].[usp_Ten_PP_cogs_ipgo_Update] " & idx & ", 'Y' "
    dbget.Execute(sqlStr)

    Call UpdateOrderCodeList(idx)
    Call UpdateMasterInfo(idx)

elseif (mode = "doapplyipgotoorder") then

    sqlStr = " exec [db_storage].[dbo].[usp_Ten_PP_ipgo_to_order_Update] " & idx & ", 'Y' "
    dbget.Execute(sqlStr)

    Call UpdateOrderCodeList(idx)
    Call UpdateMasterInfo(idx)

elseif (mode = "ReportIdxEdit") then
    if productidx="" or isnull(productidx) then
        response.write "<script type='text/javascript'>"
        response.write "alert('원가idx가 지정되지 않았습니다.');"
        response.write "</script>"
        dbget.close()	:	response.End
    end if

    existsReportCount=0
    sqlStr = " select count(reportIdx) as existsReportCount"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_eAppReport with (nolock)"
    sqlStr = sqlStr & " where isusing=1 and reportIdx = "& reportIdx &""

    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        existsReportCount = rsget("existsReportCount")
    end if
    rsget.Close

    if existsReportCount < 1 then
        response.write "<script type='text/javascript'>"
        response.write "alert('존재하는 품의번호가 아닙니다.');"
        response.write "</script>"
        response.write "존재하는 품의번호가 아닙니다."
        dbget.close()	:	response.End
    end if

    ' 이미 등록 되어 있는 품의번호 지운다.
	sqlStr = "update db_partner.dbo.tbl_eAppReport"
	sqlStr = sqlStr & " set scmlinkNo=0 where"
	sqlStr = sqlStr & " edmsIdx in (102,103,104) and scmlinkNo = "& productidx &""

    'response.write sqlStr & "<Br>"
    dbget.execute sqlStr

	sqlStr = "update db_partner.dbo.tbl_eAppReport"
	sqlStr = sqlStr & " set scmlinkNo="& productidx &""
	sqlStr = sqlStr & " ,edmsIdx=(CASE WHEN edmsIdx in (102,103,104) THEN edmsIdx ELSE 102 END) where"
	sqlStr = sqlStr & " reportIdx = "& reportIdx &""

    'response.write sqlStr & "<Br>"
    dbget.execute sqlStr

elseif (mode = "ReportIdxDel") then
    if productidx="" or isnull(productidx) then
        response.write "<script type='text/javascript'>"
        response.write "alert('원가idx가 지정되지 않았습니다.');"
        response.write "</script>"
        dbget.close()	:	response.End
    end if

    existsReportCount=0
    sqlStr = " select count(reportIdx) as existsReportCount"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_eAppReport with (nolock)"
    sqlStr = sqlStr & " where isusing=1 and reportIdx = "& reportIdx &""

    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        existsReportCount = rsget("existsReportCount")
    end if
    rsget.Close

    if existsReportCount < 1 then
        response.write "<script type='text/javascript'>"
        response.write "alert('존재하는 품의번호가 아닙니다.');"
        response.write "</script>"
        response.write "존재하는 품의번호가 아닙니다."
        dbget.close()	:	response.End
    end if

	sqlStr = "update db_partner.dbo.tbl_eAppReport"
	sqlStr = sqlStr & " set scmlinkNo=0"
	sqlStr = sqlStr & " ,edmsIdx=(CASE WHEN edmsIdx in (102,103,104) THEN edmsIdx ELSE 102 END) where"
	sqlStr = sqlStr & " reportIdx = "& reportIdx &""

    'response.write sqlStr & "<Br>"
    dbget.execute sqlStr

else
	response.write "<script language='javascript'>"
	response.write "alert('잘못된 접근입니다.');"
	response.write "</script>"
    response.write "잘못된 접근입니다."
    dbget.close()	:	response.End
end if

response.write "<script language='javascript'>"
response.write "alert('저장 되었습니다.');"
response.write "location.replace('"&refer&"');"
response.write "</script>"

Function inArray(element, arr)
    dim i
    inArray = False
    For i = 0 To Ubound(arr)
        If Trim(arr(i)) = Trim(element) Then
            inArray = True
            Exit Function
        End If
    Next
End Function

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
