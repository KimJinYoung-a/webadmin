<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/jaego_offline_cls.asp"-->
<%

dim shopid, makerid, hasstock
shopid = request("shopid")
makerid = request("makerid")
hasstock = request("hasstock")

if (hasstock = "") then
        hasstock = "N"
end if


dim ojaegoshop
set ojaegoshop = new CJaegoOffline

ojaegoshop.FRectShopid = shopid
ojaegoshop.FRectMakerid = makerid
ojaegoshop.FRectDisplayHasOnly = hasstock

ojaegoshop.GetOfflineJeagoSumByShopByMakerByItem


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

dim divname

if (ojaegoshop.FResultCount > 0) then
        divname = ojaegoshop.FItemList(0).getChargeDivName
end if

%>
<script language='javascript'>
function popOfflineShopBrandItemDetail(shopid,itemgubun, itemid, itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock_shop.asp?shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"itemipchuldetail","width=1000,height=620,scrollbars=yes, resizable=yes");
	popwin.focus();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr valign="bottom">
		<td width="10" height="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" height="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top">
		<td height="20" background="/images/tbl_blue_round_04.gif"></td>
		<td height="20" background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>OFFLINE�귣�庰�����Ȳ</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			�������� ����/�귣�庰 ��ǰ��� �����Դϴ�..
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td height="10"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td height="10" background="/images/tbl_blue_round_08.gif"></td>
		<td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
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
		��ü : <% drawSelectBoxDesignerwithName "makerid",makerid  %> &nbsp;&nbsp;
		���������� : <input type=checkbox name=hasstock value="Y" <% if (hasstock = "Y") then %>checked<% end if %>> &nbsp;&nbsp;
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>


<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="#BABABA" class="a">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="50">�̹���</td>
    	<td width="80">���ڵ�</td>
    	<td width="200">��ǰ��<br>(�ɼǸ�)</td>
	<td width="50">�ǸŰ�</td>
    	<td width="30">�ŷ�<br>����</td>
    	<td width="35">�԰�<br>����</td>
    	<td width="35">��ǰ<br>����</td>
    	<td width="35">�԰�<br>��ü</td>
    	<td width="35">��ǰ<br>��ü</td>
    	<td width="35">��<br>�Ǹ�</td>
    	<td width="35" bgcolor="F4F4F4">�ý���<br>���</td>
    	<td width="35">����</td>
    	<td width="35">�ҷ�</td>
    	<td width="35" bgcolor="F4F4F4">��ȿ<br>���</td>
    	<td width="35">����</td>
    	<td width="35" bgcolor="F4F4F4">����<br>���</td>
    	<td>���</td>
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
	<tr align="center" <% if (ojaegoshop.FItemList(i).Fsysstockno >= 0) then %>bgcolor="#FFFFFF"<% else %>bgcolor="F4F4F4"<% end if %>>
    	<td><img src="<%= ojaegoshop.FItemList(i).Fimgsmall %>" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
    	<td><%= ojaegoshop.FItemList(i).GetBarCode %></td>
    	<td align="left">
    	  <a href="javascript:popOfflineShopBrandItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).FItemgubun %>','<%= ojaegoshop.FItemList(i).FItemId %>','<%= ojaegoshop.FItemList(i).FItemoption %>')"><%= ojaegoshop.FItemList(i).FItemName %></a>
        <% if (ojaegoshop.FItemList(i).FItemOptionName <> "") then %>
    	  <br>(<%= ojaegoshop.FItemList(i).FItemOptionName %>)
        <% end if %>
    	</td>
	<td><%= formatNumber(ojaegoshop.FItemList(i).Fshopitemprice,0) %></td>
    	<td><font color="<%= ojaegoshop.FItemList(i).getChargeDivColor %>"><%= ojaegoshop.FItemList(i).getChargeDivName %></font></td>
    	<td><a href="javascript:popOfflineShopBrandItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).FItemgubun %>','<%= ojaegoshop.FItemList(i).FItemId %>','<%= ojaegoshop.FItemList(i).FItemoption %>')"><%= ojaegoshop.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:popOfflineShopBrandItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).FItemgubun %>','<%= ojaegoshop.FItemList(i).FItemId %>','<%= ojaegoshop.FItemList(i).FItemoption %>')"><%= ojaegoshop.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:popOfflineShopBrandItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).FItemgubun %>','<%= ojaegoshop.FItemList(i).FItemId %>','<%= ojaegoshop.FItemList(i).FItemoption %>')"><%= ojaegoshop.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:popOfflineShopBrandItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).FItemgubun %>','<%= ojaegoshop.FItemList(i).FItemId %>','<%= ojaegoshop.FItemList(i).FItemoption %>')"><%= ojaegoshop.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:popOfflineShopBrandItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).FItemgubun %>','<%= ojaegoshop.FItemList(i).FItemId %>','<%= ojaegoshop.FItemList(i).FItemoption %>')"><%= (ojaegoshop.FItemList(i).Fsellno + ojaegoshop.FItemList(i).Fresellno) %></a></td>
    	<td bgcolor="F4F4F4">
        <% if (ojaegoshop.FItemList(i).Fsysstockno < 0) then %>
    	  <font color="red"><%= ojaegoshop.FItemList(i).Fsysstockno %></font>
        <% else %>
          <%= ojaegoshop.FItemList(i).Fsysstockno %>
        <% end if %>
    	</td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    	<td bgcolor="F4F4F4">
	</td>
    	<td></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
		<td colspan="12">total</td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
	</tr>

</table>



<%
set ojaegoshop = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->