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
	''response.write "�����ޱ����� ���밡���մϴ�."
	''dbget.close()	:	response.End
end if

if mode="makeStockBeginStock" then
    '// �������
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_BeginStock] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""������� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockIpgo" then
    '// �԰�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_Ipgo] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""�԰� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockMove" then
    '// �̵�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_Move] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""�̵� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockSell" then
    '// �Ǹ�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_Sell] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""�Ǹ� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockSellOnGift" then
    '// ����ǰ �Ǹ�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_SellOnGift] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""����ǰ �Ǹ� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockSellUpcheWitak" then
    '// ������Ź�Ǹ�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_SellUpcheWitak] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""������Ź�Ǹ� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockShopLoss" then
    '// ��ν� + �������԰�

    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_ShopIpgo] '" & yyyymm & "' "
    dbget.execute sqlStr

    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_ShopLoss] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""�������԰�+��ν� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockCsChulgo" then
    '// CS���
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_CsChulgo] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""CS��� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeWitakSell2Maeip" then
    '// �Ǹ�(���)��(����) ��������
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_WitakSell2Maeip] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""�Ǹ�(���)��(����) �������� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="makeStockEndStock" then
    '// �⸻���
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_New_Make_EndStock] '" & yyyymm & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""�⸻��� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

elseif mode="excitem" then
    '// ����ڻ� ���ܻ�ǰ
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_monthly_Stock_MaeipLedger_Exc_Item] '" & yyyymm & "', '" & shopid & "', '" & itemgubun & "', " & itemid & ", '" & itemoption & "' "
    dbget.execute sqlStr

    if (Not IsAutoScript) and silent = "" then
	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>opener.location.reload();window.close();</script>"
    elseif silent <> "" then
        Response.charset ="euc-kr"
        response.write "{""code"": ""000"",	""message"": ""���� OK""}"
    else
        response.write "OK"
    end if
	dbget.close()	:	response.End

else
    '// �߸��� ����
	response.write "mode=" + mode
	dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
