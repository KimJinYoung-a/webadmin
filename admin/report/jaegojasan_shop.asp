<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/jaego_offline_cls.asp"-->
<%

dim shopid, sortorder
shopid = request("shopid")
sortorder = request("sortorder")

if (sortorder = "") then
        sortorder = "chargediv"
end if


dim ojaegoshop
set ojaegoshop = new CJaegoOffline

ojaegoshop.FRectShopid = shopid
ojaegoshop.FRectSortOrder = sortorder

ojaegoshop.GetOfflineJeagoSumByShopByMaker


dim i
dim TotalTenMaeipSellPriceSum, TotalTenMaeipBuyPriceSum, TotalTenWitakSellPriceSum, TotalTenWitakBuyPriceSum
dim TotalUpcheWitakSellPriceSum, TotalUpcheWitakBuyPriceSum, TotalUpcheMaeipSellPriceSum, TotalUpcheMaeipBuyPriceSum

TotalTenMaeipSellPriceSum = 0
TotalTenMaeipBuyPriceSum = 0
TotalTenWitakSellPriceSum = 0
TotalTenWitakBuyPriceSum = 0
TotalUpcheWitakSellPriceSum = 0
TotalUpcheWitakBuyPriceSum = 0
TotalUpcheMaeipSellPriceSum = 0
TotalUpcheMaeipBuyPriceSum = 0

%>
<script language='javascript'>
function popOfflineShopBrandIpChul(shopid,makerid){
	var popwin = window.open("/admin/stock/brandstock_off.asp?shopid=" + shopid + "&makerid=" + makerid,"ipchuldetail","width=1000,height=620,scrollbars=yes, resizable=yes");
	popwin.focus();
}
function popOfflineShopBrandStockJasan(shopid,makerid){
	var popwin = window.open("/admin/report/pop_jaegojasan_shop_brand.asp?shopid=" + shopid + "&makerid=" + makerid,"jasandetail","width=1000,height=620,scrollbars=yes, resizable=yes");
	popwin.focus();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>SHOP ����ڻ�</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>OFFLINE �� ���� ���� ����� ���� ����ڻ� �����Դϴ�.(���� ���� 1�� ����)
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	�� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
	        	���ļ��� : <input type="radio" name="sortorder" value="makerid" <% if (sortorder = "makerid") then %>checked<% end if %>> �귣��&nbsp;&nbsp; <input type="radio" name="sortorder" value="chargediv" <% if (sortorder = "chargediv") then %>checked<% end if %>> �ŷ�����
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="130" rowspan="2">-</td>
    	<td width="30" rowspan="2">���<br>����</td>
    	<td width="150" colspan="2">�ٸ�</td>
    	<td width="150" colspan="2">����</td>
    	<td width="150" colspan="2">����</td>
    	<td width="150" colspan="2">����</td>
    	<td rowspan="2">���</td>
    </tr>
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="75">�Һ��ڰ�</td>
    	<td width="75">���ް�</td>
    	<td width="75">�Һ��ڰ�</td>
    	<td width="75">���ް�</td>
    	<td width="75">�Һ��ڰ�</td>
    	<td width="75">���ް�</td>
    	<td width="75">�Һ��ڰ�</td>
    	<td width="75">���ް�</td>
    </tr>
<%

TotalTenMaeipSellPriceSum = 0
TotalTenMaeipBuyPriceSum = 0
TotalTenWitakSellPriceSum = 0
TotalTenWitakBuyPriceSum = 0
TotalUpcheWitakSellPriceSum = 0
TotalUpcheWitakBuyPriceSum = 0
TotalUpcheMaeipSellPriceSum = 0
TotalUpcheMaeipBuyPriceSum = 0

%>
<% for i=0 to ojaegoshop.FResultCount-1 %>
        <%
        TotalTenMaeipSellPriceSum = TotalTenMaeipSellPriceSum + ojaegoshop.FItemList(i).FTenMaeipSellPriceSum
        TotalTenMaeipBuyPriceSum = TotalTenMaeipBuyPriceSum + ojaegoshop.FItemList(i).FTenMaeipBuyPriceSum
        TotalTenWitakSellPriceSum = TotalTenWitakSellPriceSum + ojaegoshop.FItemList(i).FTenWitakSellPriceSum
        TotalTenWitakBuyPriceSum = TotalTenWitakBuyPriceSum + ojaegoshop.FItemList(i).FTenWitakBuyPriceSum
        TotalUpcheWitakSellPriceSum = TotalUpcheWitakSellPriceSum + ojaegoshop.FItemList(i).FUpcheWitakSellPriceSum
        TotalUpcheWitakBuyPriceSum = TotalUpcheWitakBuyPriceSum + ojaegoshop.FItemList(i).FUpcheWitakBuyPriceSum
        TotalUpcheMaeipSellPriceSum = TotalUpcheMaeipSellPriceSum + ojaegoshop.FItemList(i).FUpcheMaeipSellPriceSum
        TotalUpcheMaeipBuyPriceSum = TotalUpcheMaeipBuyPriceSum + ojaegoshop.FItemList(i).FUpcheMaeipBuyPriceSum
        %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td align="left"><a href="javascript:popOfflineShopBrandIpChul('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= ojaegoshop.FItemList(i).FMakerid %></a></td>
    	<td align="center"><font color="<%= ojaegoshop.FItemList(i).getChargeDivColor %>"><%= ojaegoshop.FItemList(i).getChargeDivName %></font></td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FTenMaeipSellPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FTenMaeipSellPriceSum,0) %></a>
        <% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FTenMaeipBuyPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FTenMaeipBuyPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FTenWitakSellPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FTenWitakSellPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FTenWitakBuyPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FTenWitakBuyPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FUpcheWitakSellPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheWitakSellPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FUpcheWitakBuyPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheWitakBuyPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FUpcheMaeipSellPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheMaeipSellPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td align="right">
        <% if (ojaegoshop.FItemList(i).FUpcheMaeipBuyPriceSum <> 0) then %>
    	  <a href="javascript:popOfflineShopBrandStockJasan('<%= shopid %>','<%= ojaegoshop.FItemList(i).FMakerid %>')"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheMaeipBuyPriceSum,0) %></a>
    	<% end if %>
    	</td>
    	<td></td>
    </tr>
<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td align="right"></td>
    	<td align="right"><%= FormatNumber(TotalTenMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalTenMaeipBuyPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalTenWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalTenWitakBuyPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalUpcheWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalUpcheWitakBuyPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalUpcheMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalUpcheMaeipBuyPriceSum,0) %></td>
    	<td></td>
    </tr>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<%

set ojaegoshop = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->