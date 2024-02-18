<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������fix
' History : �̻� ����
'			2023.08.04 �ѿ�� ����(��ǰ���к��� �׷��������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls_2.asp"-->
<%
Dim isViewUser, yyyy1,mm1,isusing,sysorreal, research, newitem, minusinc, bPriceGbn, vatyn, mwgubun, buseo
dim itemgubun, mygubun, purchasetype, stplace, shopid, swSppPrc, etcjungsantype, nowdate, ojaego, i
dim subTotBuySum1, subTotBuySum2, subTotBuySum3, subTotBuySum4, subTotBuySum5, subTotBuySum6, subTotBuySum7
dim subTotBuySum8, subTotBuySum11, subTotBuySum12, subTotBuySum13, subTotBuySum14, subTotBuySum, subTotOverValueSum
dim sub_totStockNo, totBuySum1, totBuySum2, totBuySum3, totBuySum4, totBuySum5, totBuySum6, totBuySum7, totBuySum8
dim totBuySum11, totBuySum12, totBuySum13, totBuySum14, totBuySum, totOverValueSum, tot_totStockNo
dim totno, totbuy, subTotno, subTotbuy, totPreno, totPrebuy , subPreno, subPrebuy'', totavgBuy, offtotavgBuy
dim totIpno,totIpBuy, subIpno, subIpBuy, totLossno, totLossBuy, subLossno, subLossBuy, iURL
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	mygubun     = requestCheckvar(request("mygubun"),10)
	minusinc   	= requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   	= requestCheckvar(request("purchasetype"),10)
	vatyn       	= requestCheckvar(request("vatyn"),10)
	stplace       	= requestCheckvar(request("stplace"),10)
	shopid       	= requestCheckvar(request("shopid"),32)
	swSppPrc	= requestCheckvar(request("swSppPrc"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)

isViewUser = FALSE ''(session("ssAdminPsn")="17")
if (sysorreal="") then sysorreal="sys"  ''real
if (isViewUser="") then sysorreal="sys"
if (isViewUser="") then bPriceGbn="P"
if (isViewUser="") then isusing=""
if (research="") and (etcjungsantype="") then etcjungsantype="41" ''����+�Ǹź�

if (research="") then
    bPriceGbn="V"
	buseo = "3X"
	mwgubun = "M"
	swSppPrc = "Y"
end if

if (stplace="") then
    stplace="L"
end if

if (mygubun = "") then
	mygubun = "M"
end if

if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

set ojaego = new CMonthlyMaeipLedge
	ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
	ojaego.FRectGubun = "sys"						'//sysorreal
	ojaego.FRectMwDiv    = mwgubun
	ojaego.FRectItemGubun = itemgubun
	ojaego.FRectTargetGbn = buseo
	ojaego.FRectVatYn    = vatyn
	ojaego.FRectShopID    = shopid
	ojaego.FRectShopSuplyPrice    = swSppPrc
	ojaego.FRectetcjungsantype = etcjungsantype
	ojaego.FRectPriceGubun = bPriceGbn

	if (stplace = "L") then
		ojaego.GetJeagoOverValueSum
	else
		ojaego.FRectLastIpgoGBN = stplace
		ojaego.GetJeagoOverValueSum_Shop
	end if

%>
<script type='text/javascript'>

function reActMonthSummary() {

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('do_stocksummary.asp?mode=stockovervalue&yyyymm=' + yyyymm,'reActMonthSummary','width=100,height=100');
	popwin.focus();
}

/*
function pop_exceldown() {
	<%
	'// ������ �԰���(����)
	'// db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Logis
	'// ������ �԰���(����)
	'// db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Shop
	'// ������ �԰���(���Ա��к�)
	'// db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_Make]
	%>
	var popwin = window.open("/admin/newreport/monthlystock_overValue_csv.asp?exYYYY=<%'= yyyy1 %>&exMM=<%'= mm1 %>&stplace=<%'= stplace %>&sysorreal=<%'= sysorreal %>&bPriceGbn=<%'= bPriceGbn %>&mygubun=<%'= mygubun %>","pop_exceldown","width=600,height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
 */

function pop_exceldown(){
	alert('�ٿ�ε����Դϴ�. ��ٷ��ּ���.');
	document.frmexcel.target = "xLink";
	document.frmexcel.exYYYY.value = '<%= yyyy1 %>';
	document.frmexcel.exMM.value = '<%= mm1 %>';
	document.frmexcel.stplace.value = '<%= stplace %>';
	document.frmexcel.sysorreal.value = '<%= sysorreal %>';
	document.frmexcel.bPriceGbn.value = '<%= bPriceGbn %>';
	document.frmexcel.mygubun.value = '<%= mygubun %>';
	<% 'document.frmexcel.action = "/admin/newreport/monthlystock_overValue_csv.asp" %>
	document.frmexcel.action = "/admin/newreport/monthlystock_overValue_excel.asp"
	document.frmexcel.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
		&nbsp;&nbsp;|&nbsp;&nbsp;
		��������
		<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >��ü
		<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >����
		<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >�鼼
		&nbsp;&nbsp;
		<input type="checkbox" name="swSppPrc" value="Y" <%= CHKIIF(swSppPrc="Y","checked","") %> >���ް��� ǥ��
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if (Not isViewUser) then %>
		<font color="#CC3333">�����:</font>
		<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >�ý������
		<!--
		<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >�ǻ����
		-->
		&nbsp;&nbsp;
		<% end if %>

		<font color="#CC3333">���Ա���:</font>
		<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
		<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >����(���̶��(��) �����Ź����)
		<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >��Ź
		<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >������
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	<font color="#CC3333">���̳ʽ�����:</font>
	<input type="radio" name="minusinc" value="" <%= CHKIIF(minusinc="","checked","") %> >���̳ʽ���� ����(��ü)
	<!--
	<input type="radio" name="minusinc" value="N" <%= CHKIIF(minusinc="N","checked","") %> >���̳ʽ���� ����
	-->
	&nbsp;&nbsp;
	<% if (Not isViewUser) then %>
	<font color="#CC3333">���԰�����:</font>
	<input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >�ۼ��ø��԰�
	<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >��ո��԰�
	&nbsp;
	<font color="#CC3333">�����Ⱓ:</font>
	<input type="radio" name="mygubun" value="M" <%= CHKIIF(mygubun="M","checked","") %>  >����
	<input type="radio" name="mygubun" value="Y" <%= CHKIIF(mygubun="Y","checked","") %>  >������
	<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<font color="#CC3333">�����ġ:</font>
		<select name="stplace">
			<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >����</option>
			<option value="M" <%= CHKIIF(stplace="M","selected" ,"") %> >����(���Ա��к�)</option>
			<option value="">---------</option>
			<option value="T" <%= CHKIIF(stplace="T","selected" ,"") %> >����(�����԰���)</option>
			<option value="S" <%= CHKIIF(stplace="S","selected" ,"") %> >����(�����԰���)</option>
		</select>
		&nbsp;
		<font color="#CC3333">�μ�����:</font>
		<% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
		&nbsp;
		<font color="#CC3333">��ǰ����:</font>
		<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		&nbsp;
		<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
			&nbsp;
			����(������� �˻���) : <% Call drawSelectBoxAccShop(yyyy1 + "-" + mm1, "", "shopid", shopid) %>
			&nbsp;
			&nbsp;
			������:
			<select class="select" name="etcjungsantype"  >
			<option value="">-����-</option>
			<option value="1" <%=CHKIIF(etcjungsantype="1","selected","")%> >�Ǹź�����</option>
			<option value="2" <%=CHKIIF(etcjungsantype="2","selected","")%> >��������</option>
			<option value="3" <%=CHKIIF(etcjungsantype="3","selected","")%> >����������</option>
			<option value="4" <%=CHKIIF(etcjungsantype="4","selected","")%> >����������</option>
			<option value="41" <%=CHKIIF(etcjungsantype="41","selected","")%> >������+�Ǹź�����</option>
			</select>
		<% end if %>
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		* �����԰���� �������� <font color="red">������</font>�� �����մϴ�.<br>
		* �������� 1���� �Ѵ� ��ǰ�� ��� <font color="red">���������(����򰡼ս�)</font>�� �����մϴ�.<br>
		* �������� 1��-2�� ������ ��� ���԰� ��� 50% �� �������� �����մϴ�.<br>
		* �������� 2���� �Ѵ� ��� ���԰� ��� 100% �� �������� �����մϴ�.<br>
		* ����(���Ա��к�) = �������Ի�ǰ�� �����԰���, �� �̿� ��ǰ�� �����԰���.<br><br>
		* <font color="red">�������� ������ �ʴ� ���</font><br>
		&nbsp; - 1. [���]����ڻ�>>����ڻ�(����) -- ���ۼ� (���� / ����)<br>
		&nbsp; - 2. [�濵]����ڻ�>>����ڻ�(����) FIX -- ���� �� ����<br>
	</td>
	<td align="right" valign="bottom">
		<% If stplace = "L" OR stplace = "T" OR stplace = "M" Then %>
			<input type="button" value="������ �ٿ�ε�" onclick="pop_exceldown();" class="button_s">
		<% End If %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="4">��ǰ����</td>
	<td rowspan="2" width="60">������</td>

	<% if (mygubun = "Y") then %>
		<td rowspan="2" width="100"><%= yyyy1 %></td>
		<td rowspan="2" width="100"><%= (yyyy1 - 1) %></td>
		<td rowspan="2" width="100"><%= (yyyy1 - 2) %></td>
		<td rowspan="2" width="100">~ <%= (yyyy1 - 3) %></td>
	<% else %>
		<td rowspan="2" width="100">1����~3����</td>
		<td rowspan="2" width="100">4����~6����</td>
		<td rowspan="2" width="100">7����~12����</td>
		<td rowspan="2" width="100">13����~18����</td>
		<td rowspan="2" width="100">19����~24����</td>
		<td rowspan="2" width="100">2���ʰ�</td>
	<% end if %>

	<td rowspan="2" width="100">NULL</td>
	<td rowspan="2" width="100">�Ѱ�</td>
	<td rowspan="2" width="100">���������</td>
	<td rowspan="2" width="100">����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td >�μ�</td>
	<td >����</td>
	<td >�ڵ屸��</td>
	<td >���Ա���</td>
</tr>
<% for i=0 to ojaego.FResultCount-1 %>
<%
if (ojaego.FItemList(i).Fitemgubun <> "75") and (ojaego.FItemList(i).Fitemgubun <> "80") and (ojaego.FItemList(i).Fitemgubun <> "85") then
	totBuySum1 = totBuySum1 + ojaego.FItemList(i).FTotBuySum1
	totBuySum2 = totBuySum2 + ojaego.FItemList(i).FTotBuySum2
	totBuySum3 = totBuySum3 + ojaego.FItemList(i).FTotBuySum3
	totBuySum4 = totBuySum4 + ojaego.FItemList(i).FTotBuySum4
	totBuySum5 = totBuySum5 + ojaego.FItemList(i).FTotBuySum5
	totBuySum6 = totBuySum6 + ojaego.FItemList(i).FTotBuySum6
	totBuySum7 = totBuySum7 + ojaego.FItemList(i).FTotBuySum7
	totBuySum8 = totBuySum8 + ojaego.FItemList(i).FTotBuySum8
	totBuySum11 = totBuySum11 + ojaego.FItemList(i).FTotBuySum11
	totBuySum12 = totBuySum12 + ojaego.FItemList(i).FTotBuySum12
	totBuySum13 = totBuySum13 + ojaego.FItemList(i).FTotBuySum13
	totBuySum14 = totBuySum14 + ojaego.FItemList(i).FTotBuySum14
	totBuySum = totBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	tot_totStockNo = tot_totStockNo + ojaego.FItemList(i).FtotStockNo

	subTotBuySum1 = subTotBuySum1 + ojaego.FItemList(i).FTotBuySum1
	subTotBuySum2 = subTotBuySum2 + ojaego.FItemList(i).FTotBuySum2
	subTotBuySum3 = subTotBuySum3 + ojaego.FItemList(i).FTotBuySum3
	subTotBuySum4 = subTotBuySum4 + ojaego.FItemList(i).FTotBuySum4
	subTotBuySum5 = subTotBuySum5 + ojaego.FItemList(i).FTotBuySum5
	subTotBuySum6 = subTotBuySum6 + ojaego.FItemList(i).FTotBuySum6
	subTotBuySum7 = subTotBuySum7 + ojaego.FItemList(i).FTotBuySum7
	subTotBuySum8 = subTotBuySum8 + ojaego.FItemList(i).FTotBuySum8
	subTotBuySum11 = subTotBuySum11 + ojaego.FItemList(i).FTotBuySum11
	subTotBuySum12 = subTotBuySum12 + ojaego.FItemList(i).FTotBuySum12
	subTotBuySum13 = subTotBuySum13 + ojaego.FItemList(i).FTotBuySum13
	subTotBuySum14 = subTotBuySum14 + ojaego.FItemList(i).FTotBuySum14
	subTotBuySum = subTotBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	sub_totStockNo = sub_totStockNo + ojaego.FItemList(i).FtotStockNo

	iURL = "monthlystock_overValue_detail_2.asp?menupos="& menupos &"&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
	iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&shopid&"&swSppPrc="&swSppPrc
	iURL = iURL + "&sysorreal=" & sysorreal + "&mygubun=" & mygubun & "&etcjungsantype="&etcjungsantype
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getBusiName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getITemGubunName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
		<td align="right"><a href="<%= iURL %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FtotStockNo,0) %></a></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><a href="<%= iURL + "&monthGubun=11" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum11,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=12" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum12,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=13" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum13,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=14" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum14,0) %></a></td>
		<% else %>
		<td align="right"><a href="<%= iURL + "&monthGubun=1" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum1,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=2" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum2,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=3" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum3,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=7" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum7,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=8" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum8,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=5" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum5,0) %></a></td>
		<% end if %>

		<td align="right"><a href="<%= iURL + "&monthGubun=6" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum6,0) %></a></td>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPriceYear,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPriceYear),0) %></td>
		<% else %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPrice,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPrice),0) %></td>
		<% end if %>
	</tr>
<% end if %>
<% next %>
<tr align="center" bgcolor="#EEFFEE">
	<td></td>
	<td>��ǰ�Ұ�</td>
	<td></td>
	<td></td>
	<td align="right"><%= FormatNumber(sub_totStockNo,0) %></td>

	<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(subTotBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum14,0) %></td>
	<% else %>
		<td align="right" ><%= FormatNumber(subTotBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum5,0) %></td>
	<% end if %>

	<td align="right" ><%= FormatNumber(subTotBuySum6,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum,0) %></b></td>
	<td align="right" ><%= FormatNumber(subTotOverValueSum,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum - subTotOverValueSum,0) %></b></td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="15"></td>
</tr>
<%
subTotBuySum1 = 0
subTotBuySum2 = 0
subTotBuySum3 = 0
subTotBuySum4 = 0
subTotBuySum5 = 0
subTotBuySum6 = 0
subTotBuySum7 = 0
subTotBuySum8 = 0
subTotBuySum11 = 0
subTotBuySum12 = 0
subTotBuySum13 = 0
subTotBuySum14 = 0
subTotBuySum = 0
subTotOverValueSum = 0
sub_totStockNo = 0
%>
<% for i=0 to ojaego.FResultCount-1 %>
<%
if (ojaego.FItemList(i).Fitemgubun = "75") or (ojaego.FItemList(i).Fitemgubun = "80") or (ojaego.FItemList(i).Fitemgubun = "85") then
	totBuySum1 = totBuySum1 + ojaego.FItemList(i).FTotBuySum1
	totBuySum2 = totBuySum2 + ojaego.FItemList(i).FTotBuySum2
	totBuySum3 = totBuySum3 + ojaego.FItemList(i).FTotBuySum3
	totBuySum4 = totBuySum4 + ojaego.FItemList(i).FTotBuySum4
	totBuySum5 = totBuySum5 + ojaego.FItemList(i).FTotBuySum5
	totBuySum6 = totBuySum6 + ojaego.FItemList(i).FTotBuySum6
	totBuySum7 = totBuySum7 + ojaego.FItemList(i).FTotBuySum7
	totBuySum8 = totBuySum8 + ojaego.FItemList(i).FTotBuySum8
	totBuySum11 = totBuySum11 + ojaego.FItemList(i).FTotBuySum11
	totBuySum12 = totBuySum12 + ojaego.FItemList(i).FTotBuySum12
	totBuySum13 = totBuySum13 + ojaego.FItemList(i).FTotBuySum13
	totBuySum14 = totBuySum14 + ojaego.FItemList(i).FTotBuySum14
	totBuySum = totBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	tot_totStockNo = tot_totStockNo + ojaego.FItemList(i).FtotStockNo

	subTotBuySum1 = subTotBuySum1 + ojaego.FItemList(i).FTotBuySum1
	subTotBuySum2 = subTotBuySum2 + ojaego.FItemList(i).FTotBuySum2
	subTotBuySum3 = subTotBuySum3 + ojaego.FItemList(i).FTotBuySum3
	subTotBuySum4 = subTotBuySum4 + ojaego.FItemList(i).FTotBuySum4
	subTotBuySum5 = subTotBuySum5 + ojaego.FItemList(i).FTotBuySum5
	subTotBuySum6 = subTotBuySum6 + ojaego.FItemList(i).FTotBuySum6
	subTotBuySum7 = subTotBuySum7 + ojaego.FItemList(i).FTotBuySum7
	subTotBuySum8 = subTotBuySum8 + ojaego.FItemList(i).FTotBuySum8
	subTotBuySum11 = subTotBuySum11 + ojaego.FItemList(i).FTotBuySum11
	subTotBuySum12 = subTotBuySum12 + ojaego.FItemList(i).FTotBuySum12
	subTotBuySum13 = subTotBuySum13 + ojaego.FItemList(i).FTotBuySum13
	subTotBuySum14 = subTotBuySum14 + ojaego.FItemList(i).FTotBuySum14
	subTotBuySum = subTotBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	sub_totStockNo = sub_totStockNo + ojaego.FItemList(i).FtotStockNo

	iURL = "monthlystock_overValue_detail_2.asp?menupos="& menupos &"&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
	iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&shopid&"&swSppPrc="&swSppPrc
	iURL = iURL + "&sysorreal=" & sysorreal + "&mygubun=" & mygubun & "&etcjungsantype="&etcjungsantype
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getBusiName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getITemGubunName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
		<td align="right"><a href="<%= iURL %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FtotStockNo,0) %></a></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><a href="<%= iURL + "&monthGubun=11" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum11,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=12" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum12,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=13" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum13,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=14" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum14,0) %></a></td>
		<% else %>
		<td align="right"><a href="<%= iURL + "&monthGubun=1" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum1,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=2" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum2,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=3" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum3,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=7" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum7,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=8" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum8,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=5" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum5,0) %></a></td>
		<% end if %>

		<td align="right"><a href="<%= iURL + "&monthGubun=6" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum6,0) %></a></td>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPriceYear,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPriceYear),0) %></td>
		<% else %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPrice,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPrice),0) %></td>
		<% end if %>
	</tr>
<% end if %>
<% next %>
<tr align="center" bgcolor="#EEFFEE">
	<td></td>
	<td>����ǰ�Ұ�</td>
	<td></td>
	<td></td>
	<td align="right"><%= FormatNumber(sub_totStockNo,0) %></td>

	<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(subTotBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum14,0) %></td>
	<% else %>
		<td align="right" ><%= FormatNumber(subTotBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum5,0) %></td>
	<% end if %>

	<td align="right" ><%= FormatNumber(subTotBuySum6,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum,0) %></b></td>
	<td align="right" ><%= FormatNumber(subTotOverValueSum,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum - subTotOverValueSum,0) %></b></td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="15"></td>
</tr>
<%
subTotBuySum1 = 0
subTotBuySum2 = 0
subTotBuySum3 = 0
subTotBuySum4 = 0
subTotBuySum5 = 0
subTotBuySum6 = 0
subTotBuySum7 = 0
subTotBuySum8 = 0
subTotBuySum11 = 0
subTotBuySum12 = 0
subTotBuySum13 = 0
subTotBuySum14 = 0
subTotBuySum = 0
subTotOverValueSum = 0
sub_totStockNo = 0
%>
<tr align="center" bgcolor="#FFFFFF">
	<td></td>
	<td>�Ѱ�</td>
	<td></td>
	<td></td>
	<td align="right"><%= FormatNumber(tot_totStockNo,0) %></td>

	<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(totBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum14,0) %></td>
	<% else %>
		<td align="right" ><%= FormatNumber(totBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum5,0) %></td>
	<% end if %>

	<td align="right" ><%= FormatNumber(totBuySum6,0) %></td>
	<td align="right" ><b><%= FormatNumber(totBuySum,0) %></b></td>
	<td align="right" ><%= FormatNumber(totOverValueSum,0) %></td>
	<td align="right" ><b><%= FormatNumber(totBuySum - totOverValueSum,0) %></b></td>
</tr>
</table>

<form name="frmexcel" method="post" style="margin:0px;">
<input type="hidden" name="exYYYY">
<input type="hidden" name="exMM">
<input type="hidden" name="stplace">
<input type="hidden" name="sysorreal">
<input type="hidden" name="bPriceGbn">
<input type="hidden" name="mygubun">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% else %>
	<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
