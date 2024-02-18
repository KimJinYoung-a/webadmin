<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcostpermeachulcls.asp"-->
<%

dim research
dim yyyy1,mm1,designer,rectorder, shopid, shopdiv, minusstockexclude
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
shopid  = requestCheckVar(request("shopid"),32)
shopdiv = requestCheckVar(request("shopdiv"),32)
minusstockexclude = requestCheckVar(request("minusstockexclude"),32)
research = requestCheckVar(request("research"),32)

if (minusstockexclude = "") then
	minusstockexclude = "N"
end if

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim oanal
set oanal = new CAnalysis
oanal.FRectYYYYMM = yyyy1 + "-" + mm1
oanal.FRectShopID=shopid

if Left(shopid,11)="streetshop0" then
	oanal.getOffMonthGainSum
elseif Left(shopid,11)="streetshop8" then
	oanal.getFrnMonthGainSum
end if

if (research = "") then
	shopid = "streetshop011"
end if



Dim oOffCost
set oOffCost = new COffShopCostPerMeachul
oOffCost.FRectYYYYMM = yyyy1 + "-" + mm1
oOffCost.FRectShopID=shopid
oOffCost.FRectShopDiv=shopdiv
oOffCost.FRectMinusStockExclude = minusstockexclude
oOffCost.GetOffShopCostSumByShopNew

Dim oOffCostJs
set oOffCostJs = new COffShopCostPerMeachul
oOffCostJs.FRectYYYYMM = yyyy1 + "-" + mm1
oOffCostJs.FRectShopID=shopid
oOffCostJs.FRectShopDiv=shopdiv
oOffCostJs.FRectMinusStockExclude = minusstockexclude
oOffCostJs.GetOffShopCostSumByJungsanNew

dim i
Dim ttlSell, ttlBuy, ttlChulSum, ttlSuplySum, innerMOrd, innerFOrd, iShopGainSum, iShopCost
Dim StockPricePrevMonth, StockPriceThisMonth
%>
<script language='javascript'>
function MakeMonthlyBrandSellSum(yyyymm){
	var popwin=window.open('dooffgainsum.asp?menupos=<%= menupos %>&mode=MakeMonthlyBrandSellSum&yyyymm='+yyyymm,'dooffgainsum','width=100,height=100');
	popwin.focus();
}

function MakeMonthlyBrandStockSum(yyyymm){
	var popwin=window.open('dooffgainsum.asp?menupos=<%= menupos %>&mode=MakeMonthlyBrandStockSum&yyyymm='+yyyymm,'MakeMonthlyBrandStockSum','width=100,height=100');
	popwin.focus();
}

function popOffGainDetail(yyyy,mm,shopid,commcd){
    var popwin=window.open('offgainsumDetail.asp?menupos=<%= menupos %>&yyyy1='+yyyy+'&mm1='+mm+'&shopid='+shopid+'&commcd='+commcd,'offgainsumDetail','width=900,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		��������:<% DrawYMBox yyyy1,mm1 %> &nbsp;&nbsp;
		���� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;

		���屸�� : <% drawSelectBoxShopDiv "shopdiv",shopdiv,true %>
		&nbsp;
		���̳ʽ���� :
		<select name="minusstockexclude">
			<option value="N" <% if (minusstockexclude = "N") then %>selected<% end if %> >����</option>
			<option value="Y" <% if (minusstockexclude = "Y") then %>selected<% end if %> >����</option>
		</select>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>

<% if (C_ADMIN_AUTH) then %>
	������ :
	<input type=button class=button value="�������ڷ� (��)�ۼ�" onclick="MakeMonthlyBrandSellSum('<%= yyyy1 %>-<%= mm1 %>');">
	&nbsp;
	<input type=button class=button value="�⸻��� �ۼ�" onclick="MakeMonthlyBrandStockSum('<%= yyyy1 %>-<%= mm1 %>');">
<% end if %>
<p>

<% if (oOffCost.FResultCount<1) then %>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr>
	<td bgcolor=#FFFFFF colspan=8>�������ڷ��ۼ� �����</td>
</tr>
</table>
<% else %>

<p>

<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
    <td width="50">���屸��</td>
    <td >����</td>
	<td width="110">�Ѹ���<br>(S)</td>
	<td width="10"></td>
	<td width="110">�������<br>(B)</td>
	<td width="110">��ü�����<br>(J)</td>
	<td width="110">��������<br>ON���Ի�ǰ<br>(���԰�)<br>(M)</td>
	<td width="110">��������<br>OF���Ի�ǰ<br>(���԰�)<br>(F)</td>
	<td width="110">��������հ�<br>(A)=J+M+F</td>
	<td width="110">�⸻���<br>(E)</td>
	<td width="110">����<br>(G)=B+A-E</td>
	<td width="100">������<br>=(S-A)/S</td>
	<td width="110">��������<br>(���԰�)<br>(C)</td>
	<td width="60">��Ÿ����</td>
</tr>
<% for i=0 to oOffCost.FResultCount-1 %>
<%
ttlSell     = ttlSell + oOffCost.FItemList(i).FttlSell
ttlBuy      = ttlBuy + oOffCost.FItemList(i).FttlBuy
ttlChulSum  = ttlChulSum + oOffCost.FItemList(i).FttlChulSum
ttlSuplySum = ttlSuplySum + oOffCost.FItemList(i).FttlSuplySum
innerMOrd   = innerMOrd + oOffCost.FItemList(i).FinnerMOrd
innerFOrd   = innerFOrd + oOffCost.FItemList(i).FinnerFOrd
iShopGainSum = iShopGainSum + oOffCost.FItemList(i).getShopGainSum
iShopCost = iShopCost + oOffCost.FItemList(i).getCostPrice
StockPricePrevMonth     = StockPricePrevMonth + oOffCost.FItemList(i).FstockPricePrevMonth
StockPriceThisMonth     = StockPriceThisMonth + oOffCost.FItemList(i).FstockPriceThisMonth
%>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center"><%= oOffCost.FItemList(i).FDivName %></td>
    <td align="center"><a href="?menupos=<%= menupos %>&shopid=<%= oOffCost.FItemList(i).FShopID %>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>"><%= oOffCost.FItemList(i).FShopName %></a></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCost.FItemList(i).FttlSell,0) %></td>
    <td></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCost.FItemList(i).FstockPricePrevMonth,0) %></td>
    <td>
        <% if (oOffCost.FItemList(i).FShopid="") then %>
        <a href="javascript:popOffGainDetail('<%=yyyy1%>','<%=mm1%>','','B021');"><%= FormatNumber(oOffCost.FItemList(i).FttlBuy,0) %></a>
        <% else %>
        <%= FormatNumber(oOffCost.FItemList(i).FttlBuy,0) %>
        <% end if %>
    </td>
    <td><%= FormatNumber(oOffCost.FItemList(i).FinnerMOrd,0) %></td>
    <td><%= FormatNumber(oOffCost.FItemList(i).FinnerFOrd,0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber((oOffCost.FItemList(i).FttlBuy + oOffCost.FItemList(i).FinnerMOrd + oOffCost.FItemList(i).FinnerFOrd),0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCost.FItemList(i).FstockPriceThisMonth,0) %></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCost.FItemList(i).getCostPrice,0) %></td>
    <td><%= oOffCost.FItemList(i).getShopGainPro %></td>
    <td><%= FormatNumber(oOffCost.FItemList(i).FttlChulSum,0) %></td>
	<td><%= FormatNumber(oOffCost.FItemList(i).FetcBuy,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center">�հ�</td>
    <td></td>
    <td><%= FormatNumber(ttlSell,0) %></td>
    <td></td>
    <td><%= FormatNumber(StockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(ttlBuy,0) %></td>
    <td><%= FormatNumber(innerMOrd,0) %></td>
    <td><%= FormatNumber(innerFOrd,0) %></td>
    <td><%= FormatNumber((ttlBuy + innerMOrd + innerFOrd),0) %></td>
    <td><%= FormatNumber(StockPriceThisMonth,0) %></td>
    <td><%= FormatNumber(iShopCost,0) %></td>
    <td>
    <% if ttlSell<>0 then %>
    <%= CLNG(iShopGainSum/ttlSell*100*100)/100 %>
    <% end if %>
    </td>
    <td><%= FormatNumber(ttlChulSum,0) %></td>
	<td></td>
</tr>
</table>
<p>
<%
ttlSell     = 0
ttlBuy      = 0
ttlChulSum  = 0
ttlSuplySum = 0
innerMOrd   = 0
innerFOrd   = 0
iShopGainSum =0
iShopCost	= 0
StockPricePrevMonth = 0
StockPriceThisMonth = 0
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
    <td >����</td>
	<td width="110">�Ѹ���<br>(S)</td>
	<td width="10"></td>
	<td width="110">�������<br>(B)</td>
	<td width="110">��ü�����<br>(J)</td>
	<td width="110">��������<br>ON���Ի�ǰ<br>(���԰�)<br>(M)</td>
	<td width="110">��������<br>OF���Ի�ǰ<br>(���԰�)<br>(F)</td>
	<td width="110">��������հ�<br>(A)=J+M+F</td>
	<td width="110">�⸻���<br>(E)</td>
	<td width="110">����<br>(G)=B+A-E</td>
	<td width="100">������<br>=(S-A)/S</td>
	<td width="110">��������<br>(���԰�)<br>(C)</td>
	<td width="60">��Ÿ����</td>
</tr>
<% for i=0 to oOffCostJs.FResultCount-1 %>
<%
ttlSell     = ttlSell + oOffCostJs.FItemList(i).FttlSell
ttlBuy      = ttlBuy + oOffCostJs.FItemList(i).FttlBuy
ttlChulSum  = ttlChulSum + oOffCostJs.FItemList(i).FttlChulSum
ttlSuplySum = ttlSuplySum + oOffCostJs.FItemList(i).FttlSuplySum
innerMOrd   = innerMOrd + oOffCostJs.FItemList(i).FinnerMOrd
innerFOrd   = innerFOrd + oOffCostJs.FItemList(i).FinnerFOrd
iShopGainSum = iShopGainSum + oOffCostJs.FItemList(i).getShopGainSum
iShopCost = iShopCost + oOffCostJs.FItemList(i).getCostPrice
StockPricePrevMonth     = StockPricePrevMonth + oOffCostJs.FItemList(i).FstockPricePrevMonth
StockPriceThisMonth     = StockPriceThisMonth + oOffCostJs.FItemList(i).FstockPriceThisMonth
%>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center"><a href="javascript:popOffGainDetail('<%=yyyy1%>','<%=mm1%>','<%=shopid%>','<%= oOffCostJs.FItemList(i).FComm_cd%>');"><%= oOffCostJs.FItemList(i).FComm_name %></a></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCostJs.FItemList(i).FttlSell,0) %></td>
    <td></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCostJs.FItemList(i).FstockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(oOffCostJs.FItemList(i).FttlBuy,0) %></td>
    <td><%= FormatNumber(oOffCostJs.FItemList(i).FinnerMOrd,0) %></td>
    <td><%= FormatNumber(oOffCostJs.FItemList(i).FinnerFOrd,0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber((oOffCostJs.FItemList(i).FttlBuy + oOffCostJs.FItemList(i).FinnerMOrd + oOffCostJs.FItemList(i).FinnerFOrd),0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCostJs.FItemList(i).FstockPriceThisMonth,0) %></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCostJs.FItemList(i).getCostPrice,0) %></td>
    <td><%= oOffCostJs.FItemList(i).getShopGainPro %></td>
    <td>
        <% if (oOffCostJs.FItemList(i).isCheckChulSumDiff) then %>
        <%= FormatNumber(oOffCostJs.FItemList(i).FttlChulSum,0) %>
        <br><font color=red>(<%= FormatNumber(oOffCostJs.FItemList(i).getChulSumDiffValue,0) %>)</font>
        <% else %>
        <%= FormatNumber(oOffCostJs.FItemList(i).FttlChulSum,0) %>
        <% end if %>
    </td>
	<td><%= FormatNumber(oOffCostJs.FItemList(i).FetcBuy,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center">�հ�</td>
    <td><%= FormatNumber(ttlSell,0) %></td>
    <td></td>
    <td><%= FormatNumber(StockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(ttlBuy,0) %></td>
    <td><%= FormatNumber(innerMOrd,0) %></td>
    <td><%= FormatNumber(innerFOrd,0) %></td>
    <td><%= FormatNumber((ttlBuy + innerMOrd + innerFOrd),0) %></td>
    <td><%= FormatNumber(StockPriceThisMonth,0) %></td>
    <td><%= FormatNumber(iShopCost,0) %></td>
    <td>
    <% if ttlSell<>0 then %>
    <%= CLNG(iShopGainSum/ttlSell*100*100)/100%>
    <% end if %>
    </td>
    <td><%= FormatNumber(ttlChulSum,0) %></td>
	<td></td>
</tr>
</table>
<br>* ��౸���� ����� �귣�尡 ������ ��� <font color="red">�������ݾ�</font>�� �޶����ϴ�.
<br>* <font color="red">������ �������</font>�� �����մϴ�.

<% end if %>


<br>
<br>
<hr>
<b><font color=red>(������)</font></b><br>
<% if (shopid<>"") and (oanal.FResultCount>0) then %>
<span class=a>* <%= shopid %></span>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr>
	<td bgcolor=#FFFFFF colspan=8>�⺻����</td>
</tr>
<tr bgcolor="#DDDDFF" align=center>
	<td>�Ѹ���</td>
	<td>���ϸ������</td>
	<td>������������</td>
	<td>����������</td>
	<td>���� �Ǹ���</td>
	<td>��ü �����</td>
	<td>���� �����</td>
	<td>�������</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align=right><%= FormatNumber(oanal.GetTotalMeachul,0) %></td>
	<td align=right><%= FormatNumber(oanal.FOneItem.FTotSpendMile,0) %></td>
	<td align=center><%= oanal.GetIpjumSusu %> %</td>
	<td align=right><%= FormatNumber(oanal.getIpjumSusuSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.GetTotalMeachul - oanal.getIpjumSusuSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.GetTotalRealSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.GetTotalShopSuplySum,0) %></td>
	<td align=right><%= FormatNumber(oanal.GetTotalMinusCharge,0) %></td>
</tr>
</table>
<br>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr>
	<td bgcolor=#FFFFFF colspan=9>�󼼳���</td>
</tr>
<tr bgcolor="#DDDDFF" align=center>
	<td>����</td>
	<td>����	</td>
	<td>����(��ü �����)</td>
	<td >����(���� �����)</td>
	<td ><font color="#888888">�԰�</font></td>
	<td ><font color="#888888">��ǰ</font></td>
	<td>����</td>
	<td>���� ������</td>
	<td>���� ����</td>
</tr>
<% for i=0 to oanal.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" align=center>
	<td align=center>
    	<% if oanal.FItemList(i).FChargeDivName<>"" then %>
    	    <%= oanal.FItemList(i).FChargeDivName %>
    	<% else %>
    	    <%= oanal.FItemList(i).getJungSanChargeDivName %>(<%= oanal.FItemList(i).FChargeDiv %>)
    	<% end if %>
	</td>
	<td align=right><font color="<%= ChkIIF(oanal.FItemList(i).FChargeDivName="" and oanal.FItemList(i).Ftotsum<>0,"#CC3333","#000000") %>"><%= FormatNumber(oanal.FItemList(i).Ftotsum,0) %></font></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Frealjungsansum,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Fshopsuplysum,0) %></td>
	<td align=right><font color="#888888"><%= FormatNumber(oanal.FItemList(i).Fchul_shopsuplysum,0) %></font></td>
	<td align=right><font color="#888888"><%= FormatNumber(oanal.FItemList(i).Fre_shopsuplysum,0) %></font></td>
	<td align=center>
	<% if oanal.FItemList(i).Ftotsum<>0 then %>
	<%= CLng(oanal.FItemList(i).Fminuscharge/oanal.FItemList(i).Ftotsum*100*100)/100 %>%
	<% end if %>
	</td>
	<td align=right></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Fminuscharge,0) %></td>
</tr>
<% next %>
<tr bgcolor="#EEEEEE">
	<td align=center>��</td>
	<td align=right>
	<% if (oanal.GetTotalMeachul=0) and (shopid="streetshop011") then %>
	<input type=button value="�������ڷ��ۼ�" onclick="MakeMonthlyBrandSellSum('<%= yyyy1 %>-<%= mm1 %>');">
	<% else %>
	<%= FormatNumber(oanal.GetTotalMeachul,0) %>
	<% end if %>
	</td>
	<td align=right><%= FormatNumber(oanal.GetTotalRealSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.GetTotalShopSuplySum,0) %></td>
	<td align=right><font color="#888888"><%= FormatNumber(oanal.GetTotalShop_ChulSuplySum,0) %></font></td>
	<td align=right><font color="#888888"><%= FormatNumber(oanal.GetTotalShop_ReSuplySum,0) %></font></td>
	<td align=center>
	<% if oanal.GetTotalMeachul<>0 then %>
	<%= CLng(oanal.GetTotalMinusCharge/oanal.GetTotalMeachul*100*100)/100 %>%
	<% end if %>
	</td>
	<td align=right><%= FormatNumber(oanal.getIpjumSusuSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.GetTotalMinusCharge-oanal.getIpjumSusuSum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align=center></td>
	<td align=right><font color="#AAAAAA">(<%= FormatNumber(oanal.FOneItem.FTotSum,0) %>)</font></td>
	<td colspan="7"><font color="#AAAAAA">(�Ǹ��� pos ����) = �Ѹ����հ� - ���ϸ������ ���� : <%= FormatNumber(oanal.FOneItem.FTotSum - (oanal.GetTotalMeachul - oanal.FOneItem.FTotSpendMile),0) %></font> </td>
</tr>
<%
dim sqlStr, bufStr
sqlStr = "select top 100 j.makerid from "
sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master j"
sqlStr = sqlStr + " left join  [db_shop].[dbo].tbl_shop_designer d"
sqlStr = sqlStr + " on j.makerid=d.makerid"
sqlStr = sqlStr + " and d.shopid='" + shopid + "'"
sqlStr = sqlStr + " where yyyymm='" + yyyy1 + "-" + mm1 + "'"
sqlStr = sqlStr + " and d.makerid is NULL"

'rsget.open sqlStr,dbget,1
'
'if Not rsget.Eof then
'   do until rsget.Eof
'        bufStr = bufStr & rsget("makerid") &","
'        rsget.movenext
'   loop
'end if
'
'rsget.close

%>
<tr>
	<td bgcolor=#FFFFFF colspan=9>
	    * ��Ź������� �Ǹű��� ������̸�, ����������� �԰���� �Ѹ��Ծ��Դϴ�.
	<br>* ���걸���� ���°�� () ���걸�� ���� Ȯ�� ���
	<br>
	    <font color="blue"><%= bufStr %></font>
	</td>
</tr>
</table>
<% end if %>
<%
set oanal = Nothing
set oOffCost = Nothing
set oOffCostJs = Nothing
%>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
