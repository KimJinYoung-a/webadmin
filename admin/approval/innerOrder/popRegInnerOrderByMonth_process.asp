<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%


dim mode, target
dim yyyy1, mm1, yyyymm

mode = requestCheckvar(Request("mode"),32)
target = requestCheckvar(Request("target"),32)

yyyy1 = requestCheckvar(Request("yyyy1"),32)
mm1 = requestCheckvar(Request("mm1"),32)
yyyymm = yyyy1 + "-" + mm1

dim sqlStr

if (mode = "reginsertshopchulgo") then

	sqlStr = " exec db_partner.dbo.usp_Ten_Shop_InnerOrder_insertShopChulgo '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

elseif (mode="reginsertupcheshopmaeip") then

	sqlStr = " exec db_partner.dbo.usp_Ten_Shop_InnerOrder_insertUpcheShopMaeip '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

elseif (mode="reginsertupcheshopwitak") then

	sqlStr = " exec db_partner.dbo.usp_Ten_Shop_InnerOrder_insertUpcheShopWitak '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

elseif (mode="reginsertshopwitaksell") then

	sqlStr = " exec db_partner.dbo.usp_Ten_Shop_InnerOrder_insertShopWitakSell '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

elseif (mode="reginsertparttoonline") then

	sqlStr = " exec db_partner.dbo.usp_Ten_Part_InnerOrder_insertPartToOnline '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

elseif (mode="reginsertparttooffline") then

	sqlStr = " exec db_partner.dbo.usp_Ten_Part_InnerOrder_insertPartToOffline '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	'response.write sqlStr
	dbget.Execute sqlStr

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

elseif (mode="reginsertall") then

	if (target = "01") then
		'// 01. �¶����Ǹ�(���̶��)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_OnSell_Maechul '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "02") then
		'// 02. �¶��θ���(���̶��)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_OnSell_Maeip '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "03") then
		'// 03. ������(ON��ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_ON '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "04") then
		'// 04. ��Ÿ����(ON��ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_ON_ETC '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "05") then
		'// 05. ������(OFF��ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_OFF '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "06") then
		'// 06. ��Ÿ����(OFF��ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_OFF_ETC '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "07") then
		'// 07. ������(��Ź��ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_Witak '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "08") then
		'// 08. �������
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_SHOP '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "09") then
		'// 09. ��ü��Ź
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_JungSan_UW '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "10") then
		'// 10. ��Ÿ����
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_JungSan_ETC '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "11") then
		'// 11. ������(��һ�ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_ITS '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "12") then
		'// 12. ��Ÿ����(��һ�ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopChulgo_ITS_ETC '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "13") then
		'// 13. �����Ǹ�(��һ�ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopSell_ITS '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	elseif (target = "14") then
		'// 14. ��Ÿ�Ǹ�(��һ�ǰ)
		sqlStr = " exec db_partner.dbo.usp_Ten_InnerOrder_ShopSell_ITS_ETC '" + CStr(yyyymm) + "', '" & session("ssbctid") & "'"
	response.write sqlStr
		dbget.Execute sqlStr
	else
		response.write "<script>alert('�߸��� �����Դϴ�.');</script>"
		response.end
	end if

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.href='popRegInnerOrderByMonth.asp?yyyy1=" + CStr(yyyy1)+ "&mm1=" + CStr(mm1) + "'</script>"
    'dbget.close()	:	response.End

else
	'// ����
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
