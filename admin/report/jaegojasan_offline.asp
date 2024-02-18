<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/jaego_offline_cls.asp"-->
<%

dim shopid
shopid = request("shopid")


dim ojaegooff
set ojaegooff = new CJaegoOffline

ojaegooff.GetOfflineJeagoSumByShop


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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>OFFLINE 재고자산</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>OFFLINE 에 대한 현재고 기준 재고자산 정보입니다.(전일 새벽 1시 기준)
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

<!-- 표 상단바 시작-->
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
	        	&nbsp;&nbsp;
	        </td>
	        <td valign="top" align="right">
	        	&nbsp;
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="160" rowspan="2">-</td>
    	<td width="150" colspan="2">텐매</td>
    	<td width="150" colspan="2">텐위</td>
    	<td width="150" colspan="2">업위</td>
    	<td width="150" colspan="2">업매</td>
    	<td rowspan="2">비고</td>
    </tr>
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="75">소비자가</td>
    	<td width="75">공급가</td>
    	<td width="75">소비자가</td>
    	<td width="75">공급가</td>
    	<td width="75">소비자가</td>
    	<td width="75">공급가</td>
    	<td width="75">소비자가</td>
    	<td width="75">공급가</td>
    </tr>
<% for i=0 to ojaegooff.FResultCount-1 %>
        <%
        TotalTenMaeipSellPriceSum = TotalTenMaeipSellPriceSum + ojaegooff.FItemList(i).FTenMaeipSellPriceSum
        TotalTenMaeipBuyPriceSum = TotalTenMaeipBuyPriceSum + ojaegooff.FItemList(i).FTenMaeipBuyPriceSum
        TotalTenWitakSellPriceSum = TotalTenWitakSellPriceSum + ojaegooff.FItemList(i).FTenWitakSellPriceSum
        TotalTenWitakBuyPriceSum = TotalTenWitakBuyPriceSum + ojaegooff.FItemList(i).FTenWitakBuyPriceSum
        TotalUpcheWitakSellPriceSum = TotalUpcheWitakSellPriceSum + ojaegooff.FItemList(i).FUpcheWitakSellPriceSum
        TotalUpcheWitakBuyPriceSum = TotalUpcheWitakBuyPriceSum + ojaegooff.FItemList(i).FUpcheWitakBuyPriceSum
        TotalUpcheMaeipSellPriceSum = TotalUpcheMaeipSellPriceSum + ojaegooff.FItemList(i).FUpcheMaeipSellPriceSum
        TotalUpcheMaeipBuyPriceSum = TotalUpcheMaeipBuyPriceSum + ojaegooff.FItemList(i).FUpcheMaeipBuyPriceSum
        %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= ojaegooff.FItemList(i).FShopid %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FTenMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FTenMaeipBuyPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FTenWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FTenWitakBuyPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FUpcheWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FUpcheWitakBuyPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FUpcheMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegooff.FItemList(i).FUpcheMaeipBuyPriceSum,0) %></td>
    	<td align="left"><%= ojaegooff.FItemList(i).FShopName %></td>
    </tr>
<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>총계</td>
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

<!-- 표 하단바 시작-->
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
<!-- 표 하단바 끝-->
<%

set ojaegooff = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->