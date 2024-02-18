<%@ language=vbscript %>
<% option explicit %>
<%
''Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, yyyymm, yyyymmdd, silent
dim itemgubun, itemid, itemoption, shopid

mode = request("mode")
yyyymm = request("yyyymm")
silent = request("silent")

shopid = requestCheckvar(request("shopid"),32)
itemgubun = requestCheckvar(request("itemgubun"),32)
itemid = requestCheckvar(request("itemid"),32)
itemoption = requestCheckvar(request("itemoption"),32)

dim sqlStr, resultrows

yyyymmdd = yyyymm + "-01"
if (DateDiff("m", yyyymmdd, Now()) > 1) then
	''response.write "지난달까지만 적용가능합니다."
	''dbget.close()	:	response.End
end if

if mode="makeStockBeginStock" then
    '// 기초재고
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_BeginStock] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""기초재고 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockIpgo" then
    '// 입고
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_Ipgo] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""입고 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockMove" then
    '// 이동
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_Move] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""이동 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockSell" then
    '// 판매
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_Sell] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""판매 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockSellOnGift" then
    '// 사은품 판매
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_SellOnGift] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""사은품 판매 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockSellUpcheWitak" then
    '// 매장위탁판매
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_SellUpcheWitak] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""매장위탁판매 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockShopLoss" then
    '// 삽로스 + 샵개별입고

    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_ShopIpgo] '" & yyyymm & "' "
    dbget.execute sqlStr

    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_ShopLoss] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""샵개별입고+삽로스 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockCsChulgo" then
    '// CS출고
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_CsChulgo] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""CS출고 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeWitakSell2Maeip" then
    '// 판매(출고)분(업배) 매입정산
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_WitakSell2Maeip] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""판매(출고)분(업배) 매입정산 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockEndStock" then
    '// 기말재고
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_EndStock] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('적용 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""기말재고 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="excitem" then
    '// 재고자산 제외상품
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_Exc_Item] '" & yyyymm & "', '" & shopid & "', '" & itemgubun & "', " & itemid & ", '" & itemoption & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('삭제 되었습니다.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""삭제 OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

else
    '// 잘못된 접근
	response.write "mode=" + mode
	dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
