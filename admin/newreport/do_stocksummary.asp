<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 60*15		' �������� ������ۼ� 15��
%>
<%
'###########################################################
' Description : ����ڻ�
' History : �̻� ����
'			2023.05.03 �ѿ�� ����(�˻������߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
dim mode, yyyymm
dim itemid, itemoption
mode = request("mode")
yyyymm = request("yyyymm")

itemid = request("itemid")
itemoption = request("itemoption")

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn, i
dim mwgubun, buseo, itemgubun, stplace, purchasetype, showsuply, dtype, makerid, shopid, etcjungsantype, showDiff
dim brandUseYN
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	minusinc   = requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   = requestCheckvar(request("purchasetype"),10)
	stplace     = requestCheckvar(request("stplace"),10)
	showsuply   = requestCheckvar(request("showsuply"),10)
	dtype       = requestCheckvar(request("dtype"),10)
	makerid     = requestCheckvar(request("makerid"),32)
	shopid     = requestCheckvar(request("shopid"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)
	showDiff      = requestCheckvar(request("showDiff"),10)
	brandUseYN      = requestCheckvar(request("brandUseYN"),10)

if (makerid<>"") then dtype=""
if (sysorreal="") then sysorreal="sys"  ''real
if (research="") and (bPriceGbn = "") then
    bPriceGbn="V"
end if
if (stplace="") then
    stplace="L"
	showDiff = "Y"
end if
if (research="") then
	if (itemgubun = "") then
		'itemgubun = "AA"
	end if
	if (buseo = "") then
		buseo = "3X"
	end if
end if

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy
dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy
dim totSellno, totSellBuy   , subSellno, subSellBuy
dim totOffChulno, totOffChulBuy  , subOffChulno, subOffChulBuy
dim totEtcChulno, totEtcChulBuy  , subEtcChulno, subEtcChulBuy
dim totCsChulno, totCsChulBuy    , subCsChulno, subCsChulBuy
dim iURL, iURLEtc, nBusiName, diffStock, diffStockPrc, diffStockW
DIM isGroupByBrand : isGroupByBrand = (dtype="mk")
Dim isItemList : isItemList = (makerid<>"")

dim totErrBadItemno, totErrBadItemBuy, subErrBadItemno, subErrBadItemBuy
dim totMoveItemno, totMoveItemBuy, subMoveItemno, subMoveItemBuy
dim totErrRealCheckno, totErrRealCheckBuy, subErrRealCheckno, subErrRealCheckBuy
dim totRealStockno, totRealStockBuy, subRealStockno, subRealStockBuy
dim totErrRealCheckBuyPlus, totErrRealCheckBuyMinus


dim sqlStr, resultrows
dim diffMonth

' "[���]����ڻ�>>����ڻ�-����" ����ڻ����ۼ� ��ư
if mode="monthlystock" then
    ''tbl_monthly_logisstock_summary :: DailyLogisStockMaker_��_2��55 / DailyLogisStockMaker_ThisDate_��_7��55 �����ٿ� ���Ե�.
    '' ���� ����Ŀ ������ ���� �� ���ۼ� .// �� ������ ������ ���ۼ�(������Ʈ)

    ''diffMonth = dateDiff("m",yyyymm+"-01",now())
 'rw "������"
 'response.end

	'// ������� ����
    sqlStr = " db_summary.dbo.sp_Ten_monthly_Acc_LogisStockMake '"&yyyymm&"'"
    dbget.execute sqlStr,resultrows

    response.write "<br>������� ����"

    '// ���Ա��� �Է�
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr,resultrows

    response.write "<br>���Ա��� �Է�"

    '// ����÷��� ����
    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','3pl'" ' 3pl
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','C'"  ' �������
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','W'"  ' �����Ź
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','M'"  ' �¶��θ���.
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','F'"  ' ��������.
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','E'"  ' ��Ÿ���.
    dbget.execute sqlStr,resultrows

    response.write "<br>����÷��� ����"

	'// �԰��� ���Ӹ� ����
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'L' "
    dbget.execute sqlStr,resultrows

    response.write "<br>�԰��� ���Ӹ� ����"

	'// ��ո��԰� ���
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'L' "
    dbget.execute sqlStr,resultrows

    response.write "<br>��ո��԰� ���"

''	''�� ����Ÿ ����
''	sqlStr = "delete from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary"+ VbCrlf
''	sqlStr = sqlStr + " where yyyymm='" + yyyymm + "'"+ VbCrlf
''	dbget.execute sqlStr
''
''
''	''��ǰ�� ���°��. �⺻�� �Է�
''	sqlStr = "insert into [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary"+ VbCrlf
''	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption"+ VbCrlf
''	sqlStr = sqlStr + " ,ipgono,reipgono,totipgono,offchulgono,offrechulgono,etcchulgono"+ VbCrlf
''	sqlStr = sqlStr + " ,etcrechulgono,totchulgono,sellno,resellno,totsellno"+ VbCrlf
''	sqlStr = sqlStr + " ,errcsno,errbaditemno,errrealcheckno,erretcno,toterrno"+ VbCrlf
''	sqlStr = sqlStr + " ,offsellno,totsysstock,availsysstock,realstock,lossno)"+ VbCrlf
''
''	sqlStr = sqlStr + " 	select '" + yyyymm + "',itemgubun,itemid,itemoption,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,sum(etcchulgono) as etcchulgono,"+ VbCrlf
''	sqlStr = sqlStr + " 	sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,sum(sellno) as sellno,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(resellno) as resellno,sum(totsellno) as totsellno,"+ VbCrlf
''	sqlStr = sqlStr + " 	sum(errcsno) as errcsno,sum(errbaditemno) as reipgono,sum(errrealcheckno) as reipgono,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(erretcno) as erretcno,sum(toterrno) as toterrno,"+ VbCrlf
''	sqlStr = sqlStr + " 	sum(offsellno) as offsellno,sum(totsysstock) as totsysstock,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(availsysstock) as availsysstock,sum(realstock) as realstock,sum(lossno) as lossno"+ VbCrlf
''	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary"+ VbCrlf
''	sqlStr = sqlStr + " 	where yyyymm<='" + yyyymm + "'"+ VbCrlf
''	sqlStr = sqlStr + " 	group by itemgubun,itemid,itemoption "+ VbCrlf
''
''	 dbget.execute sqlStr, resultrows

    '' �ۼ��� ���Ա��� �Է� // �ۼ��� ���԰�// �μ�����// �ۼ��� makerid // ��ո��԰� ����
    ''if (resultrows>0) then
    ''    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_mwFlagUpdate '"&yyyymm&"'"
    ''    dbget.execute sqlStr
    ''end if
	response.write "<script type='text/javascript'>"
    response.write "    alert('�ۼ� �Ǿ����ϴ�.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

' "[���]����ڻ�>>����ڻ�-����" �Ϻ���������ۼ�STEP1 ��ư
elseif mode="dailystock1" then
    '-- �Ϻ� ����� ���Ӹ� ���ó�¥ �¶��� �Ǹ�����
    sqlStr = "exec db_summary.[dbo].[sp_Ten_recentOnlineSell_Update_All]"
    dbget.execute sqlStr,resultrows

    response.write "<br>�Ϻ� ����� ���Ӹ� ���ó�¥ �¶��� �Ǹ�����"

    '-- �Ϻ� ����� ���Ӹ� ���ó�¥ ���ⵥ����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_daily_logisstock_summary_currentdate_make_ipchuldata]"
    dbget.execute sqlStr,resultrows

    response.write "<br>�Ϻ� ����� ���Ӹ� ���ó�¥ ���ⵥ����"

    '-- ���� ������� ���Ӹ� ���ó�¥����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_monthly_logisstock_summary_currentdate_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���� ������� ���Ӹ� ���ó�¥����"

    '-- ���繰����� ���Ӹ� ���ó�¥����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� ���ó�¥����"

    '-- ���繰����� ���Ӹ� ���ó�¥���� �¶��� �Ǹ�����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_on_ipkum_chulgo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� ���ó�¥���� �¶��� �Ǹ�����"

    '-- 3PL �¶��� 7�� �Ǹż���(usp_TEN_current_logisstock_summary_currentdate_make_on_ipkum_chulgo �� ���Խ�Ŵ)
    ''sqlStr = "exec [db_summary].[dbo].[usp_TPL_7dayOnlineSell_Update]"
    ''dbget.execute sqlStr,resultrows

    response.write "<br>3PL �¶��� 7�� �Ǹż���"

    '-- ���繰����� ���Ӹ� ���ó�¥���� ��������7���Ǹż��� ������Ʈ
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offchulgo7days]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� ���ó�¥���� ��������7���Ǹż��� ������Ʈ"

    '-- ���繰����� ���Ӹ� ���ó�¥���� �������� �����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offipchul]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� ���ó�¥���� �������� �����"

    '-- ���繰����� ���Ӹ� ���ó�¥���� ��Ÿ���� ������Ʈ
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_etcinfo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� ���ó�¥���� ��Ÿ���� ������Ʈ"

    '-- ���繰����� ���Ӹ� ���ó�¥���� ��Ÿ���� ������Ʈ2
    sqlStr = "exec [db_summary].[dbo].[usp_Ten_Refresh_MakeItem_RequireNO] NULL, NULL, NULL"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� ���ó�¥���� ��Ÿ���� ������Ʈ2"

	response.write "<script type='text/javascript'>"
    response.write "    alert('�ۼ� �Ǿ����ϴ�.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

' "[���]����ڻ�>>����ڻ�-����" �Ϻ���������ۼ�STEP2 ��ư
elseif mode="dailystock2" then
    '-- �Ϻ� ����� ���Ӹ� �̹��� �Ǹŵ�����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_daily_logisstock_summary_currentmonth_make_online_selldata]"
    dbget.execute sqlStr,resultrows

    response.write "<br>�Ϻ� ����� ���Ӹ� �̹��� �Ǹŵ�����"

    '-- 3pl stock STEP1
    sqlStr = "exec [db_summary].[dbo].[usp_TPL_recentOnlineSell_Update_All]"
    dbget.execute sqlStr,resultrows

    response.write "<br>3pl stock STEP1"

    '-- ���繰����� ���Ӹ� �̹��� ���� ��ġ����Ŀ ���ⱸ��
    sqlStr = "exec db_summary.[dbo].sp_Ten_recentOnlineSell_Update_With_6MonthAgo_loop_item_standing"
   	dbget.CommandTimeout = 60*5   ' 5��
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� �̹��� ���� ��ġ����Ŀ ���ⱸ��"

    '-- �Ϻ� ����� ���Ӹ� �̹��� ���ⵥ����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_daily_logisstock_summary_currentmonth_make_ipchuldata]"
    dbget.execute sqlStr,resultrows

    response.write "<br>�Ϻ� ����� ���Ӹ� �̹��� ���ⵥ����"

    '-- ���� ������� ���Ӹ� �̹��� ����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_monthly_logisstock_summary_currentdate_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� ����"

    '-- ���繰����� ���Ӹ� �̹��� ����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentmonth_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� �̹��� ����"

    '-- ���繰����� ���Ӹ� �̹��� ���� �¶��� �Ǹ�����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_on_ipkum_chulgo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� �̹��� ���� �¶��� �Ǹ�����"

    '-- 3pl stock STEP2
    sqlStr = "exec [db_summary].[dbo].[usp_TPL_recentOnlineJupsu_Update_All]"
    dbget.execute sqlStr,resultrows

    response.write "<br>3pl stock STEP2"

    '-- 3pl stock STEP3
    sqlStr = "exec [db_summary].[dbo].[usp_TPL_7dayOnlineSell_Update]"
    dbget.execute sqlStr,resultrows

    response.write "<br>3pl stock STEP3"

    '-- ���繰����� ���Ӹ� �̹��� ���� ��������7���Ǹż��� ������Ʈ
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offchulgo7days]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� �̹��� ���� ��������7���Ǹż��� ������Ʈ"

    '-- ���繰����� ���Ӹ� �̹��� ���� �������� �����
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offipchul]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� �̹��� ���� �������� �����"

    '-- ���繰����� ���Ӹ� �̹��� ���� ��Ÿ���� ������Ʈ
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_etcinfo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���繰����� ���Ӹ� �̹��� ���� ��Ÿ���� ������Ʈ"

    '-- �ΰŽ� ������� �ۼ�
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_recentOnlineSell_Update_ACA] 'thefingerscollabo'"
    dbget.execute sqlStr,resultrows

    response.write "<br>�ΰŽ� ������� �ۼ�"

    '-- ���� �Ǹ� ��ǰ ��� ��� ���̺� �ۼ�
    sqlStr = "exec db_summary.[dbo].[usp_Ten_ShopItem_Front_RealStockMake]"
    dbget.execute sqlStr,resultrows

    response.write "<br>���� �Ǹ� ��ǰ ��� ��� ���̺� �ۼ�"

	response.write "<script type='text/javascript'>"
    response.write "    alert('�ۼ� �Ǿ����ϴ�.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

elseif mode="monthlystockexl" then

    sqlStr = "exec [db_datamart].[dbo].[sp_Ten_monthlystock_Asset_Make] '" & yyyy1 & "-" & mm1 & "', 'L','" & shopid & "','"&buseo&"','"&itemgubun&"','"&mwgubun&"','"&vatyn&"','"&purchasetype&"','"&CHKIIF(showsuply="on",1,0)&"','"&CHKIIF(dtype="mk",1,0)&"','"&etcjungsantype&"','" & brandUseYN & "',''"
    db3_dbget.CommandTimeout = 60*5   ' 5��
    db3_dbget.execute sqlStr

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="monthlystockoneitem" then
	'// �������, ���Ӹ� ���ۼ�(����, ��Ź��ǰ)

    sqlStr = " db_summary.[dbo].[sp_Ten_monthly_Acc_LogisStockMake_OneItem] '" + CStr(yyyymm) + "', '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
    dbget.execute sqlStr,resultrows

    sqlStr = " db_summary.[dbo].[sp_Ten_monthlyLogisstock_mwFlagUpdate_OneItem] '" + CStr(yyyymm) + "', '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
    dbget.execute sqlStr,resultrows

    sqlStr = " db_summary.[dbo].[sp_Ten_monthly_Stockledger_Make_OneItem] '" + CStr(yyyymm) + "','L', '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('�ۼ� �Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="monthlystocksum" then
	'// ���Ӹ����� ���ۼ�(����+����)

    sqlStr = " db_summary.dbo.sp_Ten_monthly_Stockledger_Make '"&yyyymm&"', 'L' "
    dbget.execute sqlStr,resultrows

    sqlStr = " db_summary.dbo.sp_Ten_monthly_Stockledger_Make '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('�ۼ� �Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly" then

	''������(3�ܰ�� ����)
	response.end

    sqlStr = "exec db_summary.dbo.sp_Ten_monthly_Acc_ShopstockMake '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

	'// �԰��� ���Ӹ�
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

	'// ��ո��԰�(����)
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyShopstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr

    ''�������м�>>�������μ��ͼ��Ӹ� �� �⸻����ۼ�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"' "
    dbget.Execute sqlStr

    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopdailystock1" then
    ' �Ϻ� ������� ���Ӹ� �̹��� ����,�Ǹ� ������
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_RecentUpdateALL"
   	dbget.CommandTimeout = 60*5   ' 5��
    dbget.execute sqlStr, resultrows

    response.write "<br>�Ϻ� ������� ���Ӹ� �̹��� ����,�Ǹ� ������"

    ' ���� ������� ���Ӹ� �̹��� �Ǹ� ������
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_7daysSellUpdate"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� �Ǹ� ������"

    ' ���� ������� ���Ӹ� �̹��� ����� ���ֹ� ������Ʈ
    sqlStr = "exec db_summary.[dbo].[sp_Ten_Shop_Stock_PreOrderUpdate_ALL]"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� ����� ���ֹ� ������Ʈ"

    ' ���� ������� ���Ӹ� �̹��� ���� �����. �̵��߼���.
    sqlStr = "exec db_summary.dbo.[usp_Ten_ShopChulgo_Update]"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� ���� �����. �̵��߼���."

    ' ���� ������� ���Ӹ� �̹��� ���� �����. ��ǰ�߼���.
    sqlStr = "exec db_summary.dbo.[usp_Ten_ShopReturn_Update]"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� ���� �����. ��ǰ�߼���."

	response.write "<script type='text/javascript'>"
    response.write "    alert('�ۼ� �Ǿ����ϴ�.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthlystock" then
    ' �귣�� ����. �귣�� ���걸�� �ۼ�
    sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_ShopDesigner_Make] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>�귣�� ����. �귣�� ���걸�� �ۼ�"

    ' ���� ���������� ���Ӹ� �̹��� ���Լ���
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_SetMWDiv] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ���������� ���Ӹ� �̹��� ���Լ���"

    ' ���� ������� ���Ӹ� �̹��� �����
    sqlStr = "exec db_summary.dbo.sp_Ten_monthly_Acc_ShopstockMake '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� �����"

    ' ���� ������� ���Ӹ� �̹��� ����� ���Ա���,��ո��԰�
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyShopstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>���� ������� ���Ӹ� �̹��� ����� ���Ա���,��ո��԰�"

    ' �԰��� ���Ӹ� ����
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'S'"
    dbget.execute sqlStr, resultrows

    response.write "<br>�԰��� ���Ӹ� ����"

    ' ��ո��԰� ���
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'S'"
    dbget.execute sqlStr, resultrows

    response.write "<br>��ո��԰� ���"

    ' �⸻����ۼ�.
    sqlStr = "exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>�⸻����ۼ�"

	response.write "<script type='text/javascript'>"
    response.write "    alert('�ۼ� �Ǿ����ϴ�.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly10" then


    'sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Acc_ShopstockMake_20150506_1] '"&yyyymm&"'"
    'dbget.execute sqlStr

    'dbget.close()	:	response.End

    sqlStr = "exec db_summary.dbo.sp_Ten_monthly_Acc_ShopstockMake '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows
'rw sqlStr
    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly101" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_monthlyTable] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly102" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '1'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly103" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '2'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly104" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '3'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly105" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '4'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly11" then

    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyShopstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr

    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly20" then

	'// �԰��� ���Ӹ�
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly21" then

	'// ��ո��԰�(����)
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly30" then

    ''�������м�>>�������μ��ͼ��Ӹ� �� �⸻����ۼ�
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"' "
    dbget.Execute sqlStr

    response.write "<script>alert('�ۼ� �Ǿ����ϴ�. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

' "[���]����ڻ�>>������ ���������ۼ� ��ư
elseif mode="stockovervalue" then
	' ������ ������ �԰��� ����. ����
    sqlStr = "exec db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Logis '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>������ ������ �԰��� ����. ����"

	' ������ ������ �԰��� ����. ����
    sqlStr = "exec db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Shop '"&yyyymm&"'"
    dbget.execute sqlStr

    response.write "<br>������ ������ �԰��� ����. ����"

    response.write "<script type='text/javascript'>"
    response.write "    alert('�ۼ� �Ǿ����ϴ�.[" + CStr(yyyymm) + "]');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

' "[���]����ڻ�>>����ڻ�(����)" �ۼ�����(�������) , �����ġ(��ü) ���� ��ư
elseif mode="meaipsummake" then
    '�������/����
    if (request("atype")="S") then

''        if (request("ptype")="A") then
''            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '"&yyyymm&"','L'"
''            dbget.execute sqlStr
''
''            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '"&yyyymm&"','S'"
''            dbget.execute sqlStr
''        else
''            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '"&yyyymm&"','"&request("ptype")&"'"
''            dbget.execute sqlStr
''        end if

        '// ������� ���Ҵ���. ���Ի�ǰ �԰�/����� ���Ӹ�
		sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_Make] '"&yyyymm&"','"&CHKIIF(request("ptype")="A","",request("ptype"))&"'"
		''response.write sqlStr : dbget.close : response.end
        dbget.execute sqlStr

        response.write "<br>������� ���Ҵ���. ���Ի�ǰ �԰�/����� ���Ӹ�"

		'// ����ڻ� �ۼ��� ��Ÿ��� ��ո��԰� ������Ʈ
		sqlStr = " exec [db_summary].[dbo].[sp_Ten_monthly_EtcChulgoList_Apply_avgBuyPrice] '" & yyyymm & "' "
		dbget.execute sqlStr

        response.write "<br>����ڻ� �ۼ��� ��Ÿ��� ��ո��԰� ������Ʈ"

        response.write "<script type='text/javascript'>"
        response.write "    alert('�ۼ� �Ǿ����ϴ�.[" + CStr(yyyymm) + "].');"
    	response.write "    opener.location.reload();"
        'response.write "    self.close();"
        response.write "</script>"
    	dbget.close()	:	response.End

    elseif (request("atype")="J") then
        sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_JungsanSum_Make] '"&yyyymm&"','"&CHKIIF(request("ptype")="A","",request("ptype"))&"'"
	'rw sqlStr
        dbget.execute sqlStr

        response.write "<script type='text/javascript'>"
        response.write "    alert('�ۼ� �Ǿ����ϴ�.[" + CStr(yyyymm) + "]..');"
    	response.write "    opener.location.reload();"
        'response.write "    self.close();"
        response.write "</script>"
    	dbget.close()	:	response.End
    else
        response.write "ERR:"&request("atype")
    end if
elseif mode="meaipsumcopy" then
	sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_COPY] '" + CStr(yyyymm) + "'"
	dbget.execute sqlStr

	response.write "<script>alert('���� �Ǿ����ϴ�.[" + CStr(yyyymm) + "]');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
elseif mode="meaipsumdel" then
	sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_DEL] '" + CStr(yyyymm) + "'"
	dbget.execute sqlStr

	response.write "<script>alert('���� �Ǿ����ϴ�.[" + CStr(yyyymm) + "]');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
else
	response.write "mode=" + mode
	dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
